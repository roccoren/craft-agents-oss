/**
 * TeamsAdapter — in-process Bot Framework adapter.
 *
 * Teams is inbound-HTTP + outbound-REST, so (unlike Discord/WhatsApp) there is
 * no subprocess. The adapter owns a local HTTP listener; a public tunnel (BYO
 * URL in Phase 1) points at it. Each inbound turn stores a ConversationReference
 * so replies can be sent proactively via `continueConversationAsync`.
 */
import { createServer, type IncomingMessage as NodeReq, type Server } from 'node:http'
import { createRequire } from 'node:module'
import type {
  PlatformAdapter,
  PlatformConfig,
  AdapterCapabilities,
  IncomingMessage,
  ButtonPress,
  SentMessage,
  InlineButton,
  SendOptions,
  MessagingLogger,
} from '../../types'
import { formatForTeams, buildButtonCard } from './format'
import { BotResponse, toBotRequest } from './http-shim'
import { activityToIncoming, activityToButtonPress, type TeamsActivity } from './translate'
import { parseTeamsCredentials, testTeamsCredentials, type TeamsCredentials } from './credentials'

export { parseTeamsCredentials, testTeamsCredentials }
export type { TeamsCredentials }

const NOOP_LOGGER: MessagingLogger = {
  info: () => {}, warn: () => {}, error: () => {}, child: () => NOOP_LOGGER,
}

/** Minimal TurnContext surface the adapter uses. */
export interface TeamsTurnContextLike {
  activity: TeamsActivity & { channelId?: string; serviceUrl?: string }
  sendActivity(activity: unknown): Promise<{ id?: string } | undefined>
  updateActivity(activity: unknown): Promise<unknown>
}

/** Minimal CloudAdapter surface — injectable so tests avoid real botbuilder. */
export interface TeamsCloudAdapterLike {
  process(
    req: { body: unknown; headers: Record<string, string>; method?: string },
    res: BotResponse,
    logic: (context: TeamsTurnContextLike) => Promise<void>,
  ): Promise<void>
  continueConversationAsync(
    appId: string,
    reference: unknown,
    logic: (context: TeamsTurnContextLike) => Promise<void>,
  ): Promise<void>
}

export interface TeamsConfig extends PlatformConfig {
  appId: string
  appPassword: string
  tenantId?: string
  /** Local listener port. 0 = ephemeral (tests). Default 3978. */
  localPort?: number
  /** Injectable CloudAdapter factory (tests). Defaults to the real botbuilder one. */
  cloudAdapterFactory?: (creds: TeamsCredentials) => TeamsCloudAdapterLike
  logger?: MessagingLogger
}

export type TeamsEvent =
  | { type: 'connected'; identity: string }
  | { type: 'error'; message: string }
  | { type: 'unavailable'; reason: string; message: string }

type EventHandler = (event: TeamsEvent) => void

const MESSAGING_PATH = '/api/messages'

/** Build a proactive ConversationReference from an inbound activity. */
function referenceFromActivity(
  activity: TeamsActivity & { channelId?: string; serviceUrl?: string },
): unknown {
  return {
    channelId: activity.channelId,
    serviceUrl: activity.serviceUrl,
    conversation: activity.conversation,
    bot: activity.recipient,
    user: activity.from,
  }
}

export class TeamsAdapter implements PlatformAdapter {
  readonly platform = 'teams' as const
  readonly capabilities: AdapterCapabilities = {
    messageEditing: true,
    inlineButtons: true,
    maxButtons: 6,
    maxMessageLength: 20000,
    markdown: 'teams',
    webhookSupport: false,
  }

  private cloud: TeamsCloudAdapterLike | null = null
  private server: Server | null = null
  private appId = ''
  private log: MessagingLogger = NOOP_LOGGER
  private started = false
  private messageHandler: ((msg: IncomingMessage) => Promise<void>) | null = null
  private buttonHandler: ((press: ButtonPress) => Promise<void>) | null = null
  private eventHandlers = new Set<EventHandler>()
  private conversationRefs = new Map<string, unknown>()

  async initialize(config: PlatformConfig): Promise<void> {
    const cfg = config as TeamsConfig
    if (!cfg.appId) throw new Error('Teams: appId is required')
    if (!cfg.appPassword) throw new Error('Teams: appPassword is required')
    if (this.server) throw new Error('Teams adapter already initialized')

    this.log = (cfg.logger ?? NOOP_LOGGER).child({ component: 'teams-adapter', platform: 'teams' })
    this.appId = cfg.appId

    const factory = cfg.cloudAdapterFactory ?? defaultCloudAdapterFactory
    this.cloud = factory({ appId: cfg.appId, appPassword: cfg.appPassword, tenantId: cfg.tenantId })

    await this.startListener(cfg.localPort ?? 3978)
    this.started = true
    this.fireEvent({ type: 'connected', identity: cfg.appId })
  }

  private startListener(port: number): Promise<void> {
    return new Promise((resolve, reject) => {
      const server = createServer((req, res) => {
        if (req.method !== 'POST' || !req.url?.startsWith(MESSAGING_PATH)) {
          res.statusCode = 404
          res.end()
          return
        }
        this.handleHttp(req, res).catch((err) => {
          this.log.error('Teams inbound handler failed', { event: 'teams_inbound_error', error: err })
          if (!res.writableEnded) { res.statusCode = 500; res.end() }
        })
      })
      server.on('error', reject)
      server.listen(port, '127.0.0.1', () => {
        this.server = server
        resolve()
      })
    })
  }

  private async handleHttp(req: NodeReq, res: import('node:http').ServerResponse): Promise<void> {
    const chunks: Buffer[] = []
    for await (const c of req) chunks.push(c as Buffer)
    let body: unknown = {}
    try { body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}') } catch { /* empty */ }

    const botReq = toBotRequest(req, body)
    const botRes = new BotResponse()
    await this.cloud!.process(botReq, botRes, (context) => this.onTurn(context))
    const outcome = await botRes.finished
    res.statusCode = outcome.status
    if (outcome.body !== undefined) {
      res.setHeader('Content-Type', 'application/json')
      res.end(JSON.stringify(outcome.body))
    } else {
      res.end()
    }
  }

  private async onTurn(context: TeamsTurnContextLike): Promise<void> {
    await this.handleActivity(context.activity)
  }

  /**
   * Core inbound dispatch: store the proactive conversation reference and route
   * the activity to the message/button handlers. Exposed (not private) so it can
   * be driven deterministically in tests without a real socket round-trip; the
   * HTTP path calls it via `onTurn`.
   */
  async handleActivity(
    activity: TeamsActivity & { channelId?: string; serviceUrl?: string },
  ): Promise<void> {
    const channelId = activity.conversation?.id
    if (channelId) this.conversationRefs.set(channelId, referenceFromActivity(activity))

    const press = activityToButtonPress(activity)
    if (press && this.buttonHandler) { await this.buttonHandler(press); return }

    const msg = activityToIncoming(activity)
    if (msg && this.messageHandler) { await this.messageHandler(msg) }
  }

  async destroy(): Promise<void> {
    this.started = false
    this.conversationRefs.clear()
    const server = this.server
    this.server = null
    this.cloud = null
    if (!server) return
    await new Promise<void>((resolve) => server.close(() => resolve()))
  }

  isConnected(): boolean {
    return this.started && this.server !== null
  }

  onMessage(handler: (msg: IncomingMessage) => Promise<void>): void { this.messageHandler = handler }
  onButtonPress(handler: (press: ButtonPress) => Promise<void>): void { this.buttonHandler = handler }

  onEvent(handler: EventHandler): () => void {
    this.eventHandlers.add(handler)
    return () => this.eventHandlers.delete(handler)
  }

  getLocalPort(): number | undefined {
    const addr = this.server?.address()
    return addr && typeof addr === 'object' ? addr.port : undefined
  }

  private async sendViaRef(
    channelId: string,
    logic: (ctx: TeamsTurnContextLike) => Promise<void>,
  ): Promise<void> {
    const ref = this.conversationRefs.get(channelId)
    if (!ref) throw new Error(`Teams: no conversation reference for channel ${channelId}`)
    if (!this.cloud) throw new Error('Teams adapter is not initialized')
    await this.cloud.continueConversationAsync(this.appId, ref, logic)
  }

  async sendText(channelId: string, text: string, _opts?: SendOptions): Promise<SentMessage> {
    let messageId = ''
    await this.sendViaRef(channelId, async (ctx) => {
      const r = await ctx.sendActivity({ type: 'message', text: formatForTeams(text) })
      messageId = r?.id ?? ''
    })
    return { platform: 'teams', channelId, messageId }
  }

  async editMessage(channelId: string, messageId: string, text: string, _opts?: SendOptions): Promise<void> {
    await this.sendViaRef(channelId, async (ctx) => {
      await ctx.updateActivity({ id: messageId, type: 'message', text: formatForTeams(text) })
    })
  }

  async sendButtons(
    channelId: string, text: string, buttons: InlineButton[], _opts?: SendOptions,
  ): Promise<SentMessage> {
    let messageId = ''
    const card = buildButtonCard(text, buttons.map((b) => ({ id: b.id, label: b.label, data: b.data })))
    await this.sendViaRef(channelId, async (ctx) => {
      const r = await ctx.sendActivity({ type: 'message', attachments: [card] })
      messageId = r?.id ?? ''
    })
    return { platform: 'teams', channelId, messageId }
  }

  async clearButtons(channelId: string, messageId: string, _opts?: SendOptions): Promise<void> {
    try {
      await this.sendViaRef(channelId, async (ctx) => {
        await ctx.updateActivity({ id: messageId, type: 'message', text: '' })
      })
    } catch { /* best-effort */ }
  }

  async sendTyping(channelId: string, _opts?: SendOptions): Promise<void> {
    try {
      await this.sendViaRef(channelId, async (ctx) => { await ctx.sendActivity({ type: 'typing' }) })
    } catch { /* cosmetic */ }
  }

  async sendFile(
    channelId: string, _file: Buffer, filename: string, caption?: string, _opts?: SendOptions,
  ): Promise<SentMessage> {
    // v1 limitation: Teams file upload needs the file-consent/Graph flow. Post
    // the caption (and filename) as text so the turn is not silently dropped.
    this.log.warn('Teams file upload not supported in v1; sending caption text', {
      event: 'teams_file_unsupported', filename,
    })
    const text = caption ? `${caption} (attachment: ${filename})` : `Attachment: ${filename}`
    return this.sendText(channelId, text)
  }

  private fireEvent(event: TeamsEvent): void {
    for (const h of this.eventHandlers) { try { h(event) } catch { /* isolate */ } }
  }
}

/** Real botbuilder-backed factory. Loaded lazily (ESM-safe) so tests never load it. */
function defaultCloudAdapterFactory(creds: TeamsCredentials): TeamsCloudAdapterLike {
  const require = createRequire(import.meta.url)
  const { CloudAdapter, ConfigurationBotFrameworkAuthentication } = require('botbuilder')
  const auth = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: creds.appId,
    MicrosoftAppPassword: creds.appPassword,
    MicrosoftAppType: creds.tenantId ? 'SingleTenant' : 'MultiTenant',
    MicrosoftAppTenantId: creds.tenantId ?? '',
  })
  return new CloudAdapter(auth) as unknown as TeamsCloudAdapterLike
}
