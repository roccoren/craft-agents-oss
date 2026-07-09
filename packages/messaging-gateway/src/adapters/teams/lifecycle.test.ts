import { describe, it, expect, afterEach } from 'bun:test'
import { Readable } from 'node:stream'
import { TeamsAdapter, type TeamsConfig, type TeamsCloudAdapterLike } from './index'
import type { IncomingMessage, ButtonPress } from '../../types'

const cleanups: Array<() => Promise<void>> = []
afterEach(async () => {
  for (const c of cleanups.splice(0)) { try { await c() } catch { /* best-effort */ } }
})

/** A fake CloudAdapter: `process` invokes the turn logic with a context built
 *  from the posted body; `continueConversationAsync` records sent activities. */
function fakeFactory(record: { sent: unknown[] }): () => TeamsCloudAdapterLike {
  return () => ({
    async process(req, res, logic) {
      const context = {
        activity: req.body as never,
        async sendActivity(a: unknown) { record.sent.push(a); return { id: 'sent-1' } },
        async updateActivity(a: unknown) { record.sent.push(a); return undefined },
      }
      await logic(context)
      res.status(200); res.send(); res.end()
    },
    async continueConversationAsync(_appId, _ref, logic) {
      const context = {
        activity: {} as never,
        async sendActivity(a: unknown) { record.sent.push(a); return { id: 'proactive-1' } },
        async updateActivity(a: unknown) { record.sent.push(a); return undefined },
      }
      await logic(context)
    },
  })
}

async function makeAdapter(record: { sent: unknown[] }): Promise<TeamsAdapter> {
  const adapter = new TeamsAdapter()
  const cfg: TeamsConfig = {
    appId: 'app-1', appPassword: 'pw', localPort: 0,
    cloudAdapterFactory: fakeFactory(record),
  }
  await adapter.initialize(cfg)
  cleanups.push(() => adapter.destroy())
  return adapter
}

describe('TeamsAdapter lifecycle', () => {
  it('reports connected after initialize and disconnected after destroy', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    expect(adapter.isConnected()).toBe(true)
    expect(typeof adapter.getLocalPort()).toBe('number')
    await adapter.destroy()
    expect(adapter.isConnected()).toBe(false)
  })
})

describe('TeamsAdapter inbound dispatch (handleActivity)', () => {
  it('translates a personal message into an IncomingMessage and stores a reference', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    const seen: IncomingMessage[] = []
    adapter.onMessage(async (m) => { seen.push(m) })
    await adapter.handleActivity({
      type: 'message', id: 'a1', text: 'hi', from: { id: 'u1', name: 'Al' },
      recipient: { id: 'bot-1' }, conversation: { id: 'c1', conversationType: 'personal' },
      channelId: 'msteams', serviceUrl: 'https://smba.example',
    })
    expect(seen.length).toBe(1)
    expect(seen[0]!.channelId).toBe('c1')
    expect(seen[0]!.isDM).toBe(true)
    // Reference was stored, so a proactive reply now succeeds.
    const sent = await adapter.sendText('c1', 'reply')
    expect(sent.messageId).toBe('proactive-1')
  })

  it('dispatches an Adaptive Card submit as a ButtonPress', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    const seen: ButtonPress[] = []
    adapter.onButtonPress(async (p) => { seen.push(p) })
    await adapter.handleActivity({
      type: 'message', id: 'a2', text: '', value: { buttonId: 'yes', data: 'x' },
      from: { id: 'u1' }, recipient: { id: 'bot-1' },
      conversation: { id: 'c1', conversationType: 'personal' },
    })
    expect(seen.length).toBe(1)
    expect(seen[0]!.buttonId).toBe('yes')
  })
})

describe('TeamsAdapter proactive send', () => {
  it('sendText fails before any conversation reference is known', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    await expect(adapter.sendText('unknown', 'hi')).rejects.toThrow(/no conversation/i)
  })

  it('sendText posts the formatted text via continueConversationAsync', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    await adapter.handleActivity({
      type: 'message', id: 'a1', text: 'hi', from: { id: 'u1' }, recipient: { id: 'bot-1' },
      conversation: { id: 'c1', conversationType: 'personal' },
    })
    const sent = await adapter.sendText('c1', 'reply')
    expect(sent.platform).toBe('teams')
    expect(sent.messageId).toBe('proactive-1')
    expect(record.sent.some((a) => (a as { text?: string }).text === 'reply')).toBe(true)
  })
})

describe('TeamsAdapter HTTP path', () => {
  /** Build a fake Node request (Readable body) + response capturing the outcome. */
  function fakeHttp(body: unknown) {
    const req = Readable.from([Buffer.from(JSON.stringify(body))]) as unknown as import('node:http').IncomingMessage
    ;(req as { headers: Record<string, string> }).headers = { 'content-type': 'application/json' }
    ;(req as { method: string }).method = 'POST'
    ;(req as { url: string }).url = '/api/messages'
    let status = 0
    let ended = false
    const res = {
      statusCode: 0,
      setHeader() {},
      end() { status = res.statusCode; ended = true },
      get writableEnded() { return ended },
    } as unknown as import('node:http').ServerResponse
    return { req, res, getStatus: () => status }
  }

  it('processes a POSTed activity through the shim and replies 200', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    const seen: IncomingMessage[] = []
    adapter.onMessage(async (m) => { seen.push(m) })
    const { req, res, getStatus } = fakeHttp({
      type: 'message', id: 'a1', text: 'hi', from: { id: 'u1' },
      recipient: { id: 'bot-1' }, conversation: { id: 'c1', conversationType: 'personal' },
    })
    await (adapter as unknown as { handleHttp(r: unknown, s: unknown): Promise<void> }).handleHttp(req, res)
    expect(getStatus()).toBe(200)
    expect(seen.length).toBe(1)
  })
})
