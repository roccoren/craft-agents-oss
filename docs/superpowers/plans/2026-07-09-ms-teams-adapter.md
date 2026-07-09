# Microsoft Teams Messaging Channel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Microsoft Teams as a first-class messaging channel (bidirectional bot, session-bound) alongside Telegram, WhatsApp, Lark, and Discord.

**Architecture:** Bot Framework (Azure Bot Service). Unlike existing channels (outbound sockets), Teams is inbound-HTTP + outbound-REST, so the adapter runs **in-process** using `botbuilder`'s `CloudAdapter` — no subprocess worker. The adapter owns a small local HTTP listener that a public tunnel (BYO URL in Phase 1; on-demand Azure Dev Tunnels in Phase 2) points at. Replies are **proactive**: each inbound turn stores a `ConversationReference` keyed by conversation id, and later sends run via `continueConversationAsync`.

**Tech Stack:** TypeScript, Bun (test runner: `bun:test`), `botbuilder` (^4.23.0), Electron (renderer: React + i18next), monorepo (`workspace:*` packages).

## Global Constraints

- Platform id string is `'teams'` everywhere (`PlatformType` union member).
- Credentials stored in the `messaging_bearer` credential row, `name = 'teams'`, JSON `{ appId, appPassword, tenantId? }`.
- Tenancy "either": `tenantId` present ⇒ single-tenant (`MicrosoftAppType='SingleTenant'` + `MicrosoftAppTenantId`); absent ⇒ multi-tenant.
- `approvalChannel` default for Teams is `'chat'` (supports Adaptive Card buttons), like Telegram/Discord.
- Keep the product name "Microsoft Teams" / "Teams" in English across all locales.
- i18n parity is enforced: every key added to `en.json` MUST be added to `de.json`, `es.json`, `hu.json`, `ja.json`, `pl.json`, `zh-Hans.json`.
- Adding `'teams'` to `PlatformType` will surface TypeScript errors at every exhaustive site (`Record<PlatformType, …>` maps in `registry.ts`); those are the required edit points — fix each.
- Per-package check: `bun run typecheck` (or `bun run tsc --noEmit`). Root gate: `bun run validate:ci`.
- Run tests with `bun test <path>`.
- No real network / no real Azure Bot / no real tunnel in tests — mock everything.

---

## File Structure

**New files (Phase 1):**
- `packages/messaging-gateway/src/adapters/teams/index.ts` — `TeamsAdapter implements PlatformAdapter`, credentials parse/test helpers, `TeamsEvent` bus.
- `packages/messaging-gateway/src/adapters/teams/format.ts` — Markdown → Teams text + Adaptive Card builder.
- `packages/messaging-gateway/src/adapters/teams/http-shim.ts` — adapt Node `http` req/res to the `botbuilder` `CloudAdapter.process` shape.
- `packages/messaging-gateway/src/adapters/teams/tunnel/index.ts` — `TunnelProvider` interface + `ByoTunnelProvider`.
- Tests: `.../teams/format.test.ts`, `.../teams/lifecycle.test.ts`, `.../teams/tunnel/byo.test.ts`.

**New files (Phase 2):**
- `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.ts` — `DevTunnelBinaryProvisioner` (lazy download + cache).
- `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.ts` — `DevTunnelProvider` (spawn/parse/restart).
- Tests: `.../tunnel/devtunnel-binary.test.ts`, `.../tunnel/devtunnel.test.ts`.

**New files (Phase 1 UI):**
- `apps/electron/src/renderer/components/messaging/TeamsConnectDialog.tsx`
- `apps/electron/src/renderer/assets/messaging-icons/teams.svg`

**Modified (wiring):** `types.ts`, `registry.ts`, `index.ts`, `package.json`, `bootstrap.ts` (messaging-gateway); `router.ts` (Phase 3); `channels.ts`, `routing.ts` (shared); `messaging-registry-interface.ts`, `handlers/rpc/messaging.ts` (server-core); `channel-map.ts`, `main/index.ts`, preload types (electron); `MessagingPlatformIcon.tsx`, `MessagingSettingsPage.tsx`, `MessagingSessionMenuItem.tsx`, `PairingCodeDialog.tsx`, `atoms/messaging.ts`, `playground/mock-utils.ts` (renderer); 7 locale JSON files; `doc-links.ts`, `word-lists.ts`.

---

## PHASE 1 — Core adapter (personal chat) + BYO URL + full wiring + UI

---

### Task 1: Add `botbuilder` dependency + type-system wiring

**Files:**
- Modify: `packages/messaging-gateway/package.json`
- Modify: `packages/messaging-gateway/src/types.ts`

**Interfaces:**
- Produces: `PlatformType` now includes `'teams'`; `MessagingConfig.platforms.teams`; `BindingConfig.teamsChannelTrigger`; markdown union member `'teams'`.

- [ ] **Step 1: Add the dependency**

Edit `packages/messaging-gateway/package.json` dependencies (add alphabetically near the top of `dependencies`):

```json
    "botbuilder": "^4.23.0",
```

Run: `cd packages/messaging-gateway && bun install`
Expected: `botbuilder` and transitive `botframework-connector` resolve.

- [ ] **Step 2: Extend `PlatformType` and capabilities markdown union**

In `packages/messaging-gateway/src/types.ts`:

```ts
export type PlatformType = 'telegram' | 'whatsapp' | 'lark' | 'discord' | 'teams'
```

And extend the markdown union in `AdapterCapabilities`:

```ts
  markdown: 'v2' | 'whatsapp' | 'lark-post' | 'discord' | 'teams'
```

- [ ] **Step 3: Generalise the `isDM` / `mentionedBot` doc comments**

In `IncomingMessage`, update the two comments that say "Discord only" to read "Discord and Teams only" (fields stay the same; no code change beyond the comment). Search for `Discord only:` in `types.ts` and adjust the wording for `isDM` and `mentionedBot`.

- [ ] **Step 4: Add the Teams platform config**

In `MessagingConfig.platforms`, add after the `discord?` member:

```ts
    teams?: {
      enabled: boolean
      /** How the public messaging endpoint is provided. */
      tunnelMode: 'byo' | 'devtunnel'
      /** Bring-your-own stable HTTPS base URL (no trailing slash). Used when tunnelMode==='byo'. */
      byoUrl?: string
      /** Persistent Dev Tunnel id (Phase 2). Used when tunnelMode==='devtunnel'. */
      devtunnelId?: string
      /** Computed `<publicBase>/api/messages` shown in the UI to paste into Azure. */
      messagingEndpoint?: string
    }
```

- [ ] **Step 5: Add the per-binding channel trigger**

In `BindingConfig`, add after `discordGuildTrigger`:

```ts
  /**
   * Teams-only: in a bound channel or group chat, decides which messages
   * route to the session.
   *  - `'mention'` (default) — only messages that @mention the bot route.
   *  - `'all'` — every message routes.
   * Ignored for 1:1 personal chats (always route) and non-Teams platforms.
   */
  teamsChannelTrigger: 'mention' | 'all'
```

Add to `DEFAULT_BINDING_CONFIG`:

```ts
  teamsChannelTrigger: 'mention',
```

In `normalizeBindingConfig`, the spread of `base` + `config` already carries `teamsChannelTrigger`; no extra migration logic needed (default `'mention'` is safe). Confirm the returned object includes it via the `...base, ...config` spread.

- [ ] **Step 6: Typecheck**

Run: `cd packages/messaging-gateway && bun run typecheck`
Expected: FAIL — `registry.ts` now errors because `Record<PlatformType, …>` maps are missing `teams`. That is expected and fixed in Task 6. (If you want a clean gate here, comment nothing — proceed; Task 6 restores green.)

- [ ] **Step 7: Commit**

```bash
git add packages/messaging-gateway/package.json packages/messaging-gateway/src/types.ts bun.lock
git commit -m "feat(messaging): add teams to platform types + binding trigger config"
```

---

### Task 2: Teams Markdown/Adaptive-Card formatter

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/format.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/format.test.ts`

**Interfaces:**
- Produces: `formatForTeams(text: string): string`; `buildButtonCard(text: string, buttons: {id:string;label:string;data?:string}[]): unknown` (returns an Adaptive Card attachment object).

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/format.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { formatForTeams, buildButtonCard } from './format'

describe('formatForTeams', () => {
  it('passes plain text through unchanged', () => {
    expect(formatForTeams('hello **world**')).toBe('hello **world**')
  })

  it('does not rewrite headings inside fenced code', () => {
    const src = '```\n# not a heading\n```'
    expect(formatForTeams(src)).toBe(src)
  })

  it('downgrades ATX headings to bold', () => {
    expect(formatForTeams('# Title')).toBe('**Title**')
  })
})

describe('buildButtonCard', () => {
  it('builds an Adaptive Card attachment with one Action.Submit per button', () => {
    const card = buildButtonCard('Approve?', [
      { id: 'yes', label: 'Yes', data: 'y' },
      { id: 'no', label: 'No' },
    ]) as {
      contentType: string
      content: { actions: Array<{ type: string; title: string; data: { buttonId: string; data?: string } }> }
    }
    expect(card.contentType).toBe('application/vnd.microsoft.card.adaptive')
    expect(card.content.actions).toHaveLength(2)
    expect(card.content.actions[0]!.type).toBe('Action.Submit')
    expect(card.content.actions[0]!.title).toBe('Yes')
    expect(card.content.actions[0]!.data).toEqual({ buttonId: 'yes', data: 'y' })
    expect(card.content.actions[1]!.data).toEqual({ buttonId: 'no', data: undefined })
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/format.test.ts`
Expected: FAIL — module `./format` not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/format.ts`:

```ts
/**
 * Markdown → Teams formatting.
 *
 * Teams renders a Markdown subset in `message` activity text (bold, italic,
 * inline code, fenced code, links, lists). Like Discord, agent output over-uses
 * ATX headings for short labels; we downgrade them to bold. Everything else is
 * passed through unchanged.
 */
export function formatForTeams(text: string): string {
  const lines = text.split('\n')
  let inFence = false
  const out: string[] = []
  for (const line of lines) {
    if (/^\s*```/.test(line)) {
      inFence = !inFence
      out.push(line)
      continue
    }
    if (inFence) {
      out.push(line)
      continue
    }
    const heading = /^\s{0,3}(#{1,6})\s+(.*)$/.exec(line)
    if (heading) {
      const content = heading[2]!.trim()
      out.push(content.length > 0 ? `**${content}**` : '')
      continue
    }
    out.push(line)
  }
  return out.join('\n')
}

/**
 * Build an Adaptive Card attachment carrying one `Action.Submit` per button.
 * On click Teams posts a `message` activity whose `value` is the action's
 * `data` object — so we stash `{ buttonId, data }` there for the adapter to map
 * back to a `ButtonPress`.
 */
export function buildButtonCard(
  text: string,
  buttons: Array<{ id: string; label: string; data?: string }>,
): unknown {
  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [{ type: 'TextBlock', text: formatForTeams(text), wrap: true }],
      actions: buttons.map((b) => ({
        type: 'Action.Submit',
        title: b.label,
        data: { buttonId: b.id, data: b.data },
      })),
    },
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/format.test.ts`
Expected: PASS (5 assertions across 4 tests).

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/format.ts packages/messaging-gateway/src/adapters/teams/format.test.ts
git commit -m "feat(teams): add Markdown->Teams formatter + Adaptive Card builder"
```

---

### Task 3: HTTP request/response shim for `CloudAdapter.process`

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/http-shim.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/http-shim.test.ts`

**Interfaces:**
- Produces: `toBotRequest(req: IncomingMessage, body: unknown): { body: unknown; headers: Record<string,string>; method?: string }` and `class BotResponse` implementing the minimal `{ status(code); send(body); end(); header(k,v) }` surface `CloudAdapter.process` calls, plus `finished: Promise<{ status: number; body?: unknown }>`.

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/http-shim.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { BotResponse, toBotRequest } from './http-shim'
import type { IncomingMessage } from 'node:http'

describe('toBotRequest', () => {
  it('lowercases headers and attaches the parsed body', () => {
    const fake = { headers: { Authorization: 'Bearer x' }, method: 'POST' } as unknown as IncomingMessage
    const r = toBotRequest(fake, { type: 'message' })
    expect(r.headers.authorization).toBe('Bearer x')
    expect(r.method).toBe('POST')
    expect(r.body).toEqual({ type: 'message' })
  })
})

describe('BotResponse', () => {
  it('captures status + body and resolves finished on end()', async () => {
    const res = new BotResponse()
    res.status(201)
    res.send({ id: 'abc' })
    res.end()
    await expect(res.finished).resolves.toEqual({ status: 201, body: { id: 'abc' } })
  })

  it('defaults status to 200 when only end() is called', async () => {
    const res = new BotResponse()
    res.end()
    await expect(res.finished).resolves.toEqual({ status: 200, body: undefined })
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/http-shim.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/http-shim.ts`:

```ts
/**
 * Minimal adapters between Node's `http` server objects and the request/response
 * shape `botbuilder`'s `CloudAdapter.process(req, res, logic)` expects.
 *
 * `CloudAdapter.process` reads `req.body`, `req.headers`, `req.method`, and
 * writes via `res.status(code)`, `res.send(body)`, `res.end()`. We don't hand it
 * a raw Node response (it lacks `.status()/.send()`), so we capture the outcome
 * in-memory and flush it to the real Node response ourselves.
 */
import type { IncomingMessage } from 'node:http'

export interface BotRequest {
  body: unknown
  headers: Record<string, string>
  method?: string
}

export function toBotRequest(req: IncomingMessage, body: unknown): BotRequest {
  const headers: Record<string, string> = {}
  for (const [k, v] of Object.entries(req.headers)) {
    if (typeof v === 'string') headers[k.toLowerCase()] = v
    else if (Array.isArray(v)) headers[k.toLowerCase()] = v.join(', ')
  }
  return { body, headers, method: req.method }
}

export class BotResponse {
  private statusCode = 200
  private payload: unknown = undefined
  private resolveFinished!: (r: { status: number; body?: unknown }) => void
  readonly finished: Promise<{ status: number; body?: unknown }>

  constructor() {
    this.finished = new Promise((resolve) => {
      this.resolveFinished = resolve
    })
  }

  status(code: number): this {
    this.statusCode = code
    return this
  }

  header(_key: string, _value: string): this {
    return this
  }

  send(body?: unknown): this {
    if (body !== undefined) this.payload = body
    return this
  }

  end(): void {
    this.resolveFinished({ status: this.statusCode, body: this.payload })
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/http-shim.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/http-shim.ts packages/messaging-gateway/src/adapters/teams/http-shim.test.ts
git commit -m "feat(teams): add Node http <-> botbuilder request/response shim"
```

---

### Task 4: Credentials parse + token-based test helper

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/credentials.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/credentials.test.ts`

**Interfaces:**
- Produces:
  - `interface TeamsCredentials { appId: string; appPassword: string; tenantId?: string }`
  - `parseTeamsCredentials(raw: string): TeamsCredentials`
  - `testTeamsCredentials(creds: TeamsCredentials, fetchImpl?: typeof fetch): Promise<{ success: boolean; botName?: string; error?: string }>`

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/credentials.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { parseTeamsCredentials, testTeamsCredentials } from './credentials'

describe('parseTeamsCredentials', () => {
  it('parses appId + appPassword (+ optional tenantId)', () => {
    const c = parseTeamsCredentials(JSON.stringify({ appId: 'a', appPassword: 'p', tenantId: 't' }))
    expect(c).toEqual({ appId: 'a', appPassword: 'p', tenantId: 't' })
  })

  it('rejects missing appId', () => {
    expect(() => parseTeamsCredentials(JSON.stringify({ appPassword: 'p' }))).toThrow(/appId/)
  })

  it('rejects non-JSON', () => {
    expect(() => parseTeamsCredentials('nope')).toThrow(/valid JSON/)
  })
})

describe('testTeamsCredentials', () => {
  const okFetch = (async () =>
    new Response(JSON.stringify({ access_token: 'tok' }), { status: 200 })) as unknown as typeof fetch

  it('uses the botframework.com authority when multi-tenant', async () => {
    let calledUrl = ''
    const spy = (async (url: string) => {
      calledUrl = url
      return new Response(JSON.stringify({ access_token: 'tok' }), { status: 200 })
    }) as unknown as typeof fetch
    const r = await testTeamsCredentials({ appId: 'a', appPassword: 'p' }, spy)
    expect(r.success).toBe(true)
    expect(calledUrl).toContain('/botframework.com/')
  })

  it('uses the tenant authority when single-tenant', async () => {
    let calledUrl = ''
    const spy = (async (url: string) => {
      calledUrl = url
      return new Response(JSON.stringify({ access_token: 'tok' }), { status: 200 })
    }) as unknown as typeof fetch
    await testTeamsCredentials({ appId: 'a', appPassword: 'p', tenantId: 'tid-123' }, spy)
    expect(calledUrl).toContain('/tid-123/')
  })

  it('reports failure on a 401', async () => {
    const bad = (async () => new Response('bad', { status: 401 })) as unknown as typeof fetch
    const r = await testTeamsCredentials({ appId: 'a', appPassword: 'p' }, bad)
    expect(r.success).toBe(false)
  })

  it('succeeds with a valid token', async () => {
    const r = await testTeamsCredentials({ appId: 'a', appPassword: 'p' }, okFetch)
    expect(r.success).toBe(true)
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/credentials.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/credentials.ts`:

```ts
/**
 * Teams (Azure Bot) credentials: stored as JSON in the `messaging_bearer`
 * credential row, `name = 'teams'`.
 */
export interface TeamsCredentials {
  appId: string
  appPassword: string
  /** Present ⇒ single-tenant app; absent ⇒ multi-tenant. */
  tenantId?: string
}

export function parseTeamsCredentials(raw: string): TeamsCredentials {
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    throw new Error('Teams credentials are not valid JSON')
  }
  const obj = parsed as { appId?: unknown; appPassword?: unknown; tenantId?: unknown }
  if (typeof obj !== 'object' || obj === null) {
    throw new Error('Teams credentials must be an object')
  }
  if (typeof obj.appId !== 'string' || obj.appId.length === 0) {
    throw new Error('Teams credentials must include a non-empty "appId"')
  }
  if (typeof obj.appPassword !== 'string' || obj.appPassword.length === 0) {
    throw new Error('Teams credentials must include a non-empty "appPassword"')
  }
  const creds: TeamsCredentials = { appId: obj.appId, appPassword: obj.appPassword }
  if (typeof obj.tenantId === 'string' && obj.tenantId.length > 0) creds.tenantId = obj.tenantId
  return creds
}

/**
 * Validate credentials by acquiring a Bot Connector app token via the OAuth2
 * client-credentials flow. Single-tenant uses the tenant authority; multi-tenant
 * uses the shared `botframework.com` authority. No inbound traffic required.
 */
export async function testTeamsCredentials(
  creds: TeamsCredentials,
  fetchImpl: typeof fetch = fetch,
): Promise<{ success: boolean; botName?: string; error?: string }> {
  if (!creds.appId || !creds.appPassword) {
    return { success: false, error: 'App ID and password are required' }
  }
  const authority = creds.tenantId ?? 'botframework.com'
  const url = `https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`
  const form = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: creds.appId,
    client_secret: creds.appPassword,
    scope: 'https://api.botframework.com/.default',
  })
  try {
    const res = await fetchImpl(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form.toString(),
    })
    if (res.status === 401 || res.status === 400) {
      return { success: false, error: 'Invalid App ID or password' }
    }
    if (!res.ok) {
      return { success: false, error: `Azure AD error (${res.status})` }
    }
    const body = (await res.json()) as { access_token?: string }
    if (!body.access_token) {
      return { success: false, error: 'No access token returned' }
    }
    return { success: true, botName: creds.appId }
  } catch (err) {
    return { success: false, error: err instanceof Error ? err.message : 'Network error' }
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/credentials.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/credentials.ts packages/messaging-gateway/src/adapters/teams/credentials.test.ts
git commit -m "feat(teams): add credential parsing + token-based validation"
```

---

### Task 5: Tunnel provider interface + `ByoTunnelProvider`

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/tunnel/index.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/tunnel/byo.test.ts`

**Interfaces:**
- Produces:
  - `interface TunnelProvider { start(localPort: number): Promise<{ publicUrl: string }>; stop(): Promise<void>; onUrlChange(cb: (u: string) => void): void; isRunning(): boolean }`
  - `class ByoTunnelProvider implements TunnelProvider` (constructed with `{ baseUrl: string }`)
  - `normalizeBaseUrl(url: string): string` (strips trailing slash, requires https)

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/byo.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { ByoTunnelProvider, normalizeBaseUrl } from './index'

describe('normalizeBaseUrl', () => {
  it('strips a trailing slash', () => {
    expect(normalizeBaseUrl('https://x.example.com/')).toBe('https://x.example.com')
  })
  it('rejects non-https', () => {
    expect(() => normalizeBaseUrl('http://x.example.com')).toThrow(/https/)
  })
})

describe('ByoTunnelProvider', () => {
  it('returns the configured base url and reports running', async () => {
    const p = new ByoTunnelProvider({ baseUrl: 'https://bot.example.com/' })
    const { publicUrl } = await p.start(3978)
    expect(publicUrl).toBe('https://bot.example.com')
    expect(p.isRunning()).toBe(true)
    await p.stop()
    expect(p.isRunning()).toBe(false)
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/byo.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/index.ts`:

```ts
/**
 * TunnelProvider — supplies the public HTTPS base URL that Azure/Teams uses to
 * reach the adapter's local HTTP listener.
 *
 *  - `ByoTunnelProvider` (Phase 1): user-supplied stable URL, no process.
 *  - `DevTunnelProvider` (Phase 2): spawns a persistent Azure Dev Tunnel.
 */
export interface TunnelProvider {
  /** Start the tunnel pointing at `localPort`; resolves the public base URL. */
  start(localPort: number): Promise<{ publicUrl: string }>
  stop(): Promise<void>
  onUrlChange(cb: (publicUrl: string) => void): void
  isRunning(): boolean
}

export function normalizeBaseUrl(url: string): string {
  const trimmed = url.trim().replace(/\/+$/, '')
  if (!/^https:\/\//i.test(trimmed)) {
    throw new Error('Messaging endpoint base URL must start with https://')
  }
  return trimmed
}

export class ByoTunnelProvider implements TunnelProvider {
  private readonly baseUrl: string
  private running = false

  constructor(opts: { baseUrl: string }) {
    this.baseUrl = normalizeBaseUrl(opts.baseUrl)
  }

  async start(_localPort: number): Promise<{ publicUrl: string }> {
    this.running = true
    return { publicUrl: this.baseUrl }
  }

  async stop(): Promise<void> {
    this.running = false
  }

  onUrlChange(_cb: (publicUrl: string) => void): void {
    // BYO URL never changes at runtime.
  }

  isRunning(): boolean {
    return this.running
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/byo.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/tunnel/index.ts packages/messaging-gateway/src/adapters/teams/tunnel/byo.test.ts
git commit -m "feat(teams): add TunnelProvider interface + ByoTunnelProvider"
```

---

### Task 6: Pure activity translation (Activity → IncomingMessage / ButtonPress)

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/translate.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/translate.test.ts`

**Interfaces:**
- Produces:
  - `stripMentions(text: string): string`
  - `activityToIncoming(activity: TeamsActivity): IncomingMessage | null` (null when bot-authored or not a text message)
  - `activityToButtonPress(activity: TeamsActivity): ButtonPress | null`
  - `type TeamsActivity` — a structural subset of a botbuilder `Activity` used by these functions (`type`, `id`, `text`, `value`, `timestamp`, `from`, `recipient`, `conversation`, `entities`).

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/translate.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { stripMentions, activityToIncoming, activityToButtonPress, type TeamsActivity } from './translate'

const base: TeamsActivity = {
  type: 'message',
  id: 'act-1',
  text: 'hello',
  timestamp: '2026-07-09T00:00:00.000Z',
  from: { id: 'u1', name: 'Alice' },
  recipient: { id: 'bot-1', name: 'Bot' },
  conversation: { id: 'conv-1', conversationType: 'personal' },
  entities: [],
}

describe('stripMentions', () => {
  it('removes <at> tags and trims', () => {
    expect(stripMentions('<at>Bot</at> do it')).toBe('do it')
  })
})

describe('activityToIncoming', () => {
  it('maps a personal message with isDM=true', () => {
    const msg = activityToIncoming(base)!
    expect(msg.platform).toBe('teams')
    expect(msg.channelId).toBe('conv-1')
    expect(msg.messageId).toBe('act-1')
    expect(msg.senderId).toBe('u1')
    expect(msg.senderName).toBe('Alice')
    expect(msg.isDM).toBe(true)
    expect(msg.mentionedBot).toBe(false)
    expect(msg.text).toBe('hello')
  })

  it('detects a bot mention in a channel', () => {
    const act: TeamsActivity = {
      ...base,
      conversation: { id: 'chan-1', conversationType: 'channel' },
      text: '<at>Bot</at> ping',
      entities: [{ type: 'mention', mentioned: { id: 'bot-1' } }],
    }
    const msg = activityToIncoming(act)!
    expect(msg.isDM).toBe(false)
    expect(msg.mentionedBot).toBe(true)
    expect(msg.text).toBe('ping')
  })

  it('drops bot-authored messages', () => {
    const act: TeamsActivity = { ...base, from: { id: 'bot-1', name: 'Bot', role: 'bot' } }
    expect(activityToIncoming(act)).toBeNull()
  })

  it('returns null for an Adaptive Card submit (button) activity', () => {
    const act: TeamsActivity = { ...base, text: '', value: { buttonId: 'yes' } }
    expect(activityToIncoming(act)).toBeNull()
  })
})

describe('activityToButtonPress', () => {
  it('maps an Action.Submit value to a ButtonPress', () => {
    const act: TeamsActivity = { ...base, text: '', value: { buttonId: 'yes', data: 'x' } }
    const p = activityToButtonPress(act)!
    expect(p.platform).toBe('teams')
    expect(p.channelId).toBe('conv-1')
    expect(p.buttonId).toBe('yes')
    expect(p.data).toBe('x')
  })

  it('returns null when there is no buttonId', () => {
    expect(activityToButtonPress(base)).toBeNull()
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/translate.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/translate.ts`:

```ts
import type { IncomingMessage, ButtonPress } from '../../types'

/** Structural subset of a botbuilder Activity used by the translators. */
export interface TeamsActivity {
  type?: string
  id?: string
  text?: string
  value?: unknown
  timestamp?: string
  from?: { id?: string; name?: string; role?: string }
  recipient?: { id?: string; name?: string }
  conversation?: { id?: string; conversationType?: string }
  entities?: Array<{ type?: string; mentioned?: { id?: string } }>
}

/** Remove Teams `<at>…</at>` mention tags and collapse surrounding whitespace. */
export function stripMentions(text: string): string {
  return text.replace(/<at>.*?<\/at>/gi, '').replace(/\s+/g, ' ').trim()
}

function submitValue(activity: TeamsActivity): { buttonId?: string; data?: string } | null {
  const v = activity.value
  if (typeof v === 'object' && v !== null && typeof (v as { buttonId?: unknown }).buttonId === 'string') {
    return v as { buttonId: string; data?: string }
  }
  return null
}

export function activityToIncoming(activity: TeamsActivity): IncomingMessage | null {
  if (activity.type !== 'message') return null
  // Adaptive Card submits arrive as `message` activities with a value payload
  // and no text — those are button presses, not chat messages.
  if (submitValue(activity)) return null
  if (activity.from?.role === 'bot') return null

  const botId = activity.recipient?.id
  const mentionedBot = Boolean(
    activity.entities?.some((e) => e.type === 'mention' && e.mentioned?.id === botId),
  )
  const text = stripMentions(activity.text ?? '')

  return {
    platform: 'teams',
    channelId: activity.conversation?.id ?? '',
    messageId: activity.id ?? '',
    senderId: activity.from?.id ?? '',
    senderName: activity.from?.name,
    senderIsBot: activity.from?.role === 'bot',
    isDM: activity.conversation?.conversationType === 'personal',
    mentionedBot,
    text,
    timestamp: activity.timestamp ? Date.parse(activity.timestamp) : Date.now(),
    raw: activity,
  }
}

export function activityToButtonPress(activity: TeamsActivity): ButtonPress | null {
  const v = submitValue(activity)
  if (!v || !v.buttonId) return null
  return {
    platform: 'teams',
    channelId: activity.conversation?.id ?? '',
    messageId: activity.id ?? '',
    senderId: activity.from?.id ?? '',
    senderName: activity.from?.name,
    buttonId: v.buttonId,
    data: v.data,
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/translate.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/translate.ts packages/messaging-gateway/src/adapters/teams/translate.test.ts
git commit -m "feat(teams): add pure Activity->IncomingMessage/ButtonPress translation"
```

---

### Task 7: `TeamsAdapter` (in-process CloudAdapter + HTTP listener + proactive sends)

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/index.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/lifecycle.test.ts`

**Interfaces:**
- Consumes: `formatForTeams`, `buildButtonCard` (Task 2); `BotResponse`, `toBotRequest` (Task 3); `activityToIncoming`, `activityToButtonPress`, `TeamsActivity` (Task 6); `parseTeamsCredentials`, `TeamsCredentials`, `testTeamsCredentials` (Task 4); `PlatformAdapter`, `IncomingMessage`, `ButtonPress`, `SentMessage`, `InlineButton`, `AdapterCapabilities` (types).
- Produces:
  - `class TeamsAdapter implements PlatformAdapter`
  - `interface TeamsConfig extends PlatformConfig { appId; appPassword; tenantId?; localPort?; cloudAdapterFactory?; logger? }`
  - `type TeamsEvent = { type:'connected'; identity:string } | { type:'error'; message:string } | { type:'unavailable'; reason:string; message:string }`
  - `interface TeamsCloudAdapterLike { process(req, res, logic): Promise<void>; continueConversationAsync(appId, ref, logic): Promise<void> }`
  - re-exports `parseTeamsCredentials`, `testTeamsCredentials`, `TeamsCredentials`
  - method `onEvent(handler): () => void`, `getLocalPort(): number | undefined`

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/lifecycle.test.ts`:

```ts
import { describe, it, expect, afterEach } from 'bun:test'
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
        activity: req.body,
        async sendActivity(a: unknown) { record.sent.push(a); return { id: 'sent-1' } },
        async updateActivity(a: unknown) { record.sent.push(a); return undefined },
      }
      await logic(context)
      res.status(200); res.send(); res.end()
    },
    async continueConversationAsync(_appId, _ref, logic) {
      const context = {
        activity: {},
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

function post(port: number, body: unknown): Promise<Response> {
  return fetch(`http://127.0.0.1:${port}/api/messages`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
  })
}

describe('TeamsAdapter inbound', () => {
  it('translates a posted personal message into an IncomingMessage', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    const seen: IncomingMessage[] = []
    adapter.onMessage(async (m) => { seen.push(m) })
    const port = adapter.getLocalPort()!
    const res = await post(port, {
      type: 'message', id: 'a1', text: 'hi', from: { id: 'u1', name: 'Al' },
      recipient: { id: 'bot-1' }, conversation: { id: 'c1', conversationType: 'personal' },
      channelId: 'msteams', serviceUrl: 'https://smba.example',
    })
    expect(res.status).toBe(200)
    await new Promise((r) => setTimeout(r, 20))
    expect(seen.length).toBe(1)
    expect(seen[0]!.channelId).toBe('c1')
    expect(seen[0]!.isDM).toBe(true)
  })

  it('dispatches an Adaptive Card submit as a ButtonPress', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    const seen: ButtonPress[] = []
    adapter.onButtonPress(async (p) => { seen.push(p) })
    const port = adapter.getLocalPort()!
    await post(port, {
      type: 'message', id: 'a2', text: '', value: { buttonId: 'yes', data: 'x' },
      from: { id: 'u1' }, recipient: { id: 'bot-1' },
      conversation: { id: 'c1', conversationType: 'personal' },
    })
    await new Promise((r) => setTimeout(r, 20))
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

  it('sendText works after an inbound message stores the reference', async () => {
    const record = { sent: [] as unknown[] }
    const adapter = await makeAdapter(record)
    adapter.onMessage(async () => {})
    const port = adapter.getLocalPort()!
    await post(port, {
      type: 'message', id: 'a1', text: 'hi', from: { id: 'u1' }, recipient: { id: 'bot-1' },
      conversation: { id: 'c1', conversationType: 'personal' },
      channelId: 'msteams', serviceUrl: 'https://smba.example',
    })
    await new Promise((r) => setTimeout(r, 20))
    const sent = await adapter.sendText('c1', 'reply')
    expect(sent.platform).toBe('teams')
    expect(sent.messageId).toBe('proactive-1')
    expect(record.sent.some((a) => (a as { text?: string }).text === 'reply')).toBe(true)
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/lifecycle.test.ts`
Expected: FAIL — module `./index` not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/index.ts`:

```ts
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
    const activity = context.activity
    const channelId = activity.conversation?.id
    if (channelId) this.conversationRefs.set(channelId, referenceFromActivity(activity))

    const press = activityToButtonPress(activity)
    if (press && this.buttonHandler) { void this.buttonHandler(press); return }

    const msg = activityToIncoming(activity)
    if (msg && this.messageHandler) { void this.messageHandler(msg) }
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/lifecycle.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 5: Typecheck the package (types only; registry still needs Task 8)**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/`
Expected: PASS (all teams adapter tests).

- [ ] **Step 6: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/index.ts packages/messaging-gateway/src/adapters/teams/lifecycle.test.ts
git commit -m "feat(teams): add in-process TeamsAdapter (CloudAdapter + HTTP + proactive sends)"
```

---

### Task 8: Registry wiring (`registry.ts`) + `index.ts` export + `bootstrap.ts`

**Files:**
- Modify: `packages/messaging-gateway/src/registry.ts`
- Modify: `packages/messaging-gateway/src/index.ts`
- Modify: `packages/messaging-gateway/src/bootstrap.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/registry-teams.test.ts` (new; smoke test of test/save credential path with a fake registry construction is heavy, so we unit-test the tunnel-endpoint computation helper instead — see Step 8)

**Interfaces:**
- Consumes: `TeamsAdapter`, `parseTeamsCredentials`, `testTeamsCredentials`, `TeamsCredentials`, `TeamsEvent` (Task 7); `ByoTunnelProvider`, `TunnelProvider`, `normalizeBaseUrl` (Task 5).
- Produces on the registry class:
  - `testTeamsCredentials(creds: TeamsCredentials): Promise<{ success: boolean; botName?: string; error?: string }>`
  - `saveTeamsCredentials(workspaceId: string, creds: TeamsSaveInput): Promise<{ messagingEndpoint: string }>` where `TeamsSaveInput = TeamsCredentials & { tunnelMode: 'byo' | 'devtunnel'; byoUrl?: string }`
  - `MessagingGatewayRegistryOptions.teams?: { localPort?: number }`

> NOTE: adding `'teams'` to `PlatformType` (Task 1) makes the two `runtime: Record<PlatformType, …>` literals fail typecheck. Steps 4 and 7 fix them; that is the compiler telling you where the exhaustive sites are.

- [ ] **Step 1: Imports + helper**

At the top of `registry.ts`, next to the Discord import (line ~33), add:

```ts
import { TeamsAdapter, parseTeamsCredentials, testTeamsCredentials as testTeamsToken, type TeamsCredentials, type TeamsEvent } from './adapters/teams/index'
import { ByoTunnelProvider, normalizeBaseUrl, type TunnelProvider } from './adapters/teams/tunnel/index'
```

Add near the bottom helper functions (next to `createRuntime`, line ~1820):

```ts
export interface TeamsSaveInput extends TeamsCredentials {
  tunnelMode: 'byo' | 'devtunnel'
  byoUrl?: string
}

function teamsMessagingEndpoint(publicBaseUrl: string): string {
  return `${normalizeBaseUrl(publicBaseUrl)}/api/messages`
}
```

- [ ] **Step 2: Registry options**

In `MessagingGatewayRegistryOptions` (next to the `discord?` block, line ~81), add:

```ts
  /** Optional Teams config — enables the Teams (Bot Framework) adapter. */
  teams?: {
    /** Local HTTP listener port the tunnel points at. Default 3978. */
    localPort?: number
  }
```

- [ ] **Step 3: WorkspaceState fields**

In `interface WorkspaceState` (next to `discord: DiscordAdapter | null`, line ~99), add:

```ts
  teams: TeamsAdapter | null
  teamsOffEvent?: () => void
  teamsTunnel?: TunnelProvider | null
```

- [ ] **Step 4: Runtime records (two sites)**

In `initializeWorkspace`'s runtime clone (line ~257-260) add after the `discord:` line:

```ts
      teams: cloneRuntime(state.runtime.teams),
```

In `bootstrapWorkspace`'s runtime literal (line ~1073-1078) add after the `discord:` line, and initialize `teams: null` in the state object next to `discord: null`:

```ts
        teams: createRuntime('teams', isPlatformConfigured(cfg, 'teams')),
```
```ts
      teams: null,
```

- [ ] **Step 5: `initializeWorkspace` connect trigger**

After the Discord `if (isPlatformConfigured(config, 'discord')) { … }` block (line ~180-193), add:

```ts
    if (isPlatformConfigured(config, 'teams')) {
      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: false, state: 'connecting', lastError: undefined,
      })
      void this.tryConnectTeams(workspaceId, state).catch((err) => {
        this.log.error('failed to connect Teams on init', { event: 'teams_connect_failed', workspaceId, error: err })
      })
    }
```

- [ ] **Step 6: `updateConfig` teardown**

Next to the Discord teardown (lines ~280-289 and ~331-338), add analogous cleanup:

After `await state.gateway.unregisterAdapter('discord').catch(() => {})` add:
```ts
    await state.gateway.unregisterAdapter('teams').catch(() => {})
```
After the discord destroy block add:
```ts
    state.teamsOffEvent?.()
    state.teamsOffEvent = undefined
    if (state.teams) { await state.teams.destroy().catch(() => {}); state.teams = null }
    if (state.teamsTunnel) { await state.teamsTunnel.stop().catch(() => {}); state.teamsTunnel = null }
```

In the `for (const platform of ['telegram', 'whatsapp', 'lark', 'discord'] as const)` loop (line ~321), add `'teams'`:
```ts
    for (const platform of ['telegram', 'whatsapp', 'lark', 'discord', 'teams'] as const) {
```
And add a teams branch mirroring the `if (!configured && platform === 'discord')` block:
```ts
      if (!configured && platform === 'teams') {
        state.teamsOffEvent?.(); state.teamsOffEvent = undefined
        if (state.teams) { await state.teams.destroy().catch(() => {}); state.teams = null }
        if (state.teamsTunnel) { await state.teamsTunnel.stop().catch(() => {}); state.teamsTunnel = null }
      }
```

- [ ] **Step 7: `testTeamsCredentials` + `saveTeamsCredentials`**

Next to `saveDiscordCredentials` (line ~725), add:

```ts
  async testTeamsCredentials(
    creds: TeamsCredentials,
  ): Promise<{ success: boolean; botName?: string; error?: string }> {
    return testTeamsToken(creds)
  }

  async saveTeamsCredentials(
    workspaceId: string,
    input: TeamsSaveInput,
  ): Promise<{ messagingEndpoint: string }> {
    const test = await testTeamsToken(input)
    if (!test.success) throw new Error(test.error ?? 'Invalid Teams credentials')
    if (input.tunnelMode === 'byo') {
      if (!input.byoUrl) throw new Error('A public HTTPS URL is required for bring-your-own tunnel mode')
      normalizeBaseUrl(input.byoUrl) // throws on non-https
    }

    await this.opts.credentialManager.set(
      { type: 'messaging_bearer', workspaceId, name: 'teams' },
      { value: JSON.stringify({ appId: input.appId, appPassword: input.appPassword, tenantId: input.tenantId }) },
    )

    const state = this.workspaces.get(workspaceId) ?? this.bootstrapWorkspace(workspaceId)
    state.configStore.update({
      enabled: true,
      platforms: { teams: { enabled: true, tunnelMode: input.tunnelMode, byoUrl: input.byoUrl } },
    })

    this.setPlatformRuntime(workspaceId, state, 'teams', {
      configured: true, connected: false, state: 'connecting', lastError: undefined,
    })

    await this.tryConnectTeams(workspaceId, state)
    await state.gateway.start()

    const endpoint = state.configStore.get().platforms.teams?.messagingEndpoint
    return { messagingEndpoint: endpoint ?? '' }
  }
```

- [ ] **Step 8: `tryConnectTeams`**

Next to `tryConnectDiscord` (line ~760), add:

```ts
  private async tryConnectTeams(workspaceId: string, state: WorkspaceState): Promise<void> {
    const cfg = state.configStore.get().platforms.teams
    const cred = await this.opts.credentialManager
      .get({ type: 'messaging_bearer', workspaceId, name: 'teams' })
      .catch(() => null)
    if (!cfg || !cred?.value) {
      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: false, state: 'error', lastError: 'Teams credentials are missing.',
      })
      return
    }

    let creds: TeamsCredentials
    try { creds = parseTeamsCredentials(cred.value) } catch (err) {
      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: false, state: 'error',
        lastError: err instanceof Error ? err.message : 'Teams credentials are malformed',
      })
      return
    }

    // Tear down any prior adapter/tunnel (reconnect path).
    await state.gateway.unregisterAdapter('teams').catch(() => {})
    state.teamsOffEvent?.(); state.teamsOffEvent = undefined
    if (state.teams) { await state.teams.destroy().catch(() => {}); state.teams = null }
    if (state.teamsTunnel) { await state.teamsTunnel.stop().catch(() => {}); state.teamsTunnel = null }

    try {
      // Phase 1: BYO tunnel only. (Phase 2 adds the devtunnel branch here.)
      if (cfg.tunnelMode !== 'byo' || !cfg.byoUrl) {
        throw new Error('Teams requires a bring-your-own HTTPS URL in this build')
      }
      const tunnel: TunnelProvider = new ByoTunnelProvider({ baseUrl: cfg.byoUrl })
      const localPort = this.opts.teams?.localPort ?? 3978
      const { publicUrl } = await tunnel.start(localPort)
      state.teamsTunnel = tunnel

      const adapter = new TeamsAdapter()
      state.teams = adapter
      state.teamsOffEvent = adapter.onEvent((ev) => this.onTeamsEvent(workspaceId, ev))
      await adapter.initialize({
        appId: creds.appId, appPassword: creds.appPassword, tenantId: creds.tenantId,
        localPort,
        logger: this.log.child({ component: 'teams-adapter', workspaceId, platform: 'teams' }),
      })
      state.gateway.registerAdapter(adapter)

      state.configStore.update({
        platforms: { teams: { ...cfg, messagingEndpoint: teamsMessagingEndpoint(publicUrl) } },
      })

      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: true, state: 'connected',
        identity: creds.appId, lastError: undefined,
      })
    } catch (err) {
      this.log.error('failed to connect Teams', { event: 'teams_connect_failed', workspaceId, error: err })
      state.teamsOffEvent?.(); state.teamsOffEvent = undefined
      state.teams = null
      if (state.teamsTunnel) { await state.teamsTunnel.stop().catch(() => {}); state.teamsTunnel = null }
      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: false, state: 'error',
        lastError: err instanceof Error ? err.message : String(err),
      })
      throw err
    }
  }

  private onTeamsEvent(workspaceId: string, event: TeamsEvent): void {
    const state = this.workspaces.get(workspaceId)
    if (!state) return
    if (event.type === 'error' || event.type === 'unavailable') {
      this.setPlatformRuntime(workspaceId, state, 'teams', {
        configured: true, connected: false, state: 'error',
        lastError: event.type === 'error' ? event.message : event.message,
      })
    }
  }
```

- [ ] **Step 9: Export from `index.ts`**

In `packages/messaging-gateway/src/index.ts`, next to the Discord export block, add:

```ts
export {
  TeamsAdapter,
  parseTeamsCredentials,
  testTeamsCredentials,
  type TeamsConfig,
  type TeamsCredentials,
  type TeamsEvent,
} from './adapters/teams/index'
```

- [ ] **Step 10: `bootstrap.ts` options passthrough**

In `packages/messaging-gateway/src/bootstrap.ts`, next to the `discord?` option (line ~49) add a `teams?: { localPort?: number }` field to the bootstrap options interface, and in the registry-options construction (next to the `...(opts.discord ? … )` spread at line ~87) add:

```ts
    ...(opts.teams ? { teams: { localPort: opts.teams.localPort } } : {}),
```

- [ ] **Step 11: Typecheck**

Run: `cd packages/messaging-gateway && bun run typecheck`
Expected: PASS (all `Record<PlatformType, …>` sites now include `teams`).

- [ ] **Step 12: Run the whole gateway test suite**

Run: `cd packages/messaging-gateway && bun test`
Expected: PASS (existing suites unaffected; teams adapter suites pass).

- [ ] **Step 13: Commit**

```bash
git add packages/messaging-gateway/src/registry.ts packages/messaging-gateway/src/index.ts packages/messaging-gateway/src/bootstrap.ts
git commit -m "feat(messaging): wire TeamsAdapter into registry + bootstrap (BYO tunnel)"
```

---

### Task 9: Shared protocol channels + routing classification

**Files:**
- Modify: `packages/shared/src/protocol/channels.ts`
- Modify: `packages/shared/src/protocol/routing.ts`
- Test: `packages/shared/src/protocol/__tests__/routing.test.ts` (existing — must stay green)

**Interfaces:**
- Produces: `RPC_CHANNELS.messaging.TEST_TEAMS = 'messaging:testTeams'`, `RPC_CHANNELS.messaging.SAVE_TEAMS = 'messaging:saveTeams'`.

- [ ] **Step 1: Add the channels**

In `channels.ts`, after `SAVE_DISCORD` (line ~434):

```ts
    TEST_TEAMS: 'messaging:testTeams',
    SAVE_TEAMS: 'messaging:saveTeams',
```

- [ ] **Step 2: Classify in routing.ts**

In `routing.ts`, in the same LOCAL_ONLY messaging block after `SAVE_DISCORD` (line ~456):

```ts
  RPC_CHANNELS.messaging.TEST_TEAMS,
  RPC_CHANNELS.messaging.SAVE_TEAMS,
```

(Both go in the SAME set as `TEST_DISCORD`/`SAVE_DISCORD` — locate that set and append there.)

- [ ] **Step 3: Run the exhaustiveness test**

Run: `cd packages/shared && bun test src/protocol/__tests__/routing.test.ts`
Expected: PASS — "every channel is classified exactly once".

- [ ] **Step 4: Commit**

```bash
git add packages/shared/src/protocol/channels.ts packages/shared/src/protocol/routing.ts
git commit -m "feat(protocol): add TEST_TEAMS/SAVE_TEAMS messaging channels"
```

---

### Task 10: server-core registry interface + RPC handlers

**Files:**
- Modify: `packages/server-core/src/handlers/messaging-registry-interface.ts`
- Modify: `packages/server-core/src/handlers/rpc/messaging.ts`

**Interfaces:**
- Consumes: `RPC_CHANNELS.messaging.TEST_TEAMS/SAVE_TEAMS` (Task 9); `IMessagingGatewayRegistry` gains teams methods matching `registry.ts` (Task 8).
- Produces: interface methods `testTeamsCredentials`, `saveTeamsCredentials`.

- [ ] **Step 1: Extend the interface**

In `messaging-registry-interface.ts`, after the Discord methods (line ~205):

```ts
  /** Validate Teams (Azure Bot) credentials by acquiring a Bot Connector token. */
  testTeamsCredentials(creds: {
    appId: string
    appPassword: string
    tenantId?: string
  }): Promise<{ success: boolean; botName?: string; error?: string }>

  /** Save Teams credentials + tunnel config and (re)initialize the adapter. */
  saveTeamsCredentials(workspaceId: string, creds: {
    appId: string
    appPassword: string
    tenantId?: string
    tunnelMode: 'byo' | 'devtunnel'
    byoUrl?: string
  }): Promise<{ messagingEndpoint: string }>
```

- [ ] **Step 2: Register the RPC handlers**

In `handlers/rpc/messaging.ts`, after the `SAVE_DISCORD` handler (line ~70):

```ts
  server.handle(RPC_CHANNELS.messaging.TEST_TEAMS, async (
    _ctx,
    creds: { appId: string; appPassword: string; tenantId?: string },
  ) => {
    return registry.testTeamsCredentials(creds)
  })

  server.handle(RPC_CHANNELS.messaging.SAVE_TEAMS, async (
    ctx,
    creds: { appId: string; appPassword: string; tenantId?: string; tunnelMode: 'byo' | 'devtunnel'; byoUrl?: string },
  ) => {
    if (!ctx.workspaceId) throw new Error('Missing workspaceId')
    const result = await registry.saveTeamsCredentials(ctx.workspaceId, creds)
    return { success: true, ...result }
  })
```

- [ ] **Step 3: Typecheck**

Run: `cd packages/server-core && bun run typecheck`
Expected: PASS (the `MessagingGatewayRegistry` in messaging-gateway already implements the new methods from Task 8).

- [ ] **Step 4: Commit**

```bash
git add packages/server-core/src/handlers/messaging-registry-interface.ts packages/server-core/src/handlers/rpc/messaging.ts
git commit -m "feat(server-core): add TEST_TEAMS/SAVE_TEAMS RPC handlers"
```

---

### Task 11: Electron transport channel-map + main registry options

**Files:**
- Modify: `apps/electron/src/transport/channel-map.ts`
- Modify: `apps/electron/src/main/index.ts`
- Modify: the preload/electronAPI type surface (search for where `saveDiscordCredentials` is typed on `window.electronAPI`; likely `apps/electron/src/shared/types.ts` or a preload `.d.ts`)

**Interfaces:**
- Produces: `window.electronAPI.testTeamsCredentials(creds)` and `saveTeamsCredentials(creds)` bound to the RPC channels; registry options carry `teams: { localPort }`.

- [ ] **Step 1: channel-map entries**

In `channel-map.ts`, after `saveDiscordCredentials` (line ~420):

```ts
  testTeamsCredentials: invoke(RPC_CHANNELS.messaging.TEST_TEAMS),
  saveTeamsCredentials: invoke(RPC_CHANNELS.messaging.SAVE_TEAMS),
```

- [ ] **Step 2: main/index.ts registry options**

In `apps/electron/src/main/index.ts`, in the `createMessagingBootstrap({ … })` call, after the `discord: { … }` block (line ~690-694), add:

```ts
  // Teams runs in-process (no worker). It hosts a local HTTP listener that the
  // tunnel points at; the default port is 3978.
  teams: {
    localPort: 3978,
  },
```

- [ ] **Step 3: Preload type surface**

Find the interface that declares `saveDiscordCredentials` on the electron API (grep `saveDiscordCredentials` under `apps/electron/src`). Add matching declarations:

```ts
  testTeamsCredentials(creds: { appId: string; appPassword: string; tenantId?: string }): Promise<{ success: boolean; botName?: string; error?: string }>
  saveTeamsCredentials(creds: { appId: string; appPassword: string; tenantId?: string; tunnelMode: 'byo' | 'devtunnel'; byoUrl?: string }): Promise<{ success: boolean; messagingEndpoint: string }>
```

- [ ] **Step 4: Typecheck the electron app**

Run: `cd apps/electron && bun run typecheck`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add apps/electron/src/transport/channel-map.ts apps/electron/src/main/index.ts apps/electron/src/shared/types.ts
git commit -m "feat(electron): wire Teams credential RPC + in-process adapter options"
```

---

### Task 12: Renderer icon + platform-union updates

**Files:**
- Create: `apps/electron/src/renderer/assets/messaging-icons/teams.svg`
- Modify: `apps/electron/src/renderer/components/messaging/MessagingPlatformIcon.tsx`
- Modify: `apps/electron/src/renderer/components/messaging/MessagingSessionMenuItem.tsx`
- Modify: `apps/electron/src/renderer/components/messaging/PairingCodeDialog.tsx`
- Modify: `apps/electron/src/renderer/atoms/messaging.ts`
- Modify: `apps/electron/src/renderer/playground/mock-utils.ts`

**Interfaces:**
- Produces: `'teams'` added to every renderer `MessagingPlatform` / platform union; Teams icon registered.

- [ ] **Step 1: Add the icon**

Create `apps/electron/src/renderer/assets/messaging-icons/teams.svg` (simple mark, Teams purple):

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">
  <rect width="24" height="24" rx="5" fill="#6264A7"/>
  <text x="12" y="16" font-family="Segoe UI, Arial, sans-serif" font-size="12" font-weight="700" fill="#ffffff" text-anchor="middle">T</text>
</svg>
```

- [ ] **Step 2: Register in MessagingPlatformIcon.tsx**

Add the import (after `discordIcon`):
```ts
import teamsIcon from '@/assets/messaging-icons/teams.svg'
```
Extend the union and both records:
```ts
type MessagingPlatform = 'telegram' | 'whatsapp' | 'lark' | 'discord' | 'teams'
```
```ts
  discord: discordIcon,
  teams: teamsIcon,
```
```ts
  discord: { bg: '#5865F2', initial: 'D' },
  teams: { bg: '#6264A7', initial: 'T' },
```

- [ ] **Step 3: Update the other unions**

In each of these files, change the union `'telegram' | 'whatsapp' | 'lark' | 'discord'` to add `| 'teams'`:
- `MessagingSessionMenuItem.tsx` (the exported `MessagingPlatform` type, line ~30)
- `PairingCodeDialog.tsx` (the `platform` prop, line ~22)
- `atoms/messaging.ts` (the `pairing` dialog `platform` field, line ~63)
- `playground/mock-utils.ts` (`AllowListPlatform`, line ~37)

- [ ] **Step 4: Typecheck**

Run: `cd apps/electron && bun run typecheck`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add apps/electron/src/renderer/assets/messaging-icons/teams.svg apps/electron/src/renderer/components/messaging/MessagingPlatformIcon.tsx apps/electron/src/renderer/components/messaging/MessagingSessionMenuItem.tsx apps/electron/src/renderer/components/messaging/PairingCodeDialog.tsx apps/electron/src/renderer/atoms/messaging.ts apps/electron/src/renderer/playground/mock-utils.ts
git commit -m "feat(electron-ui): add Teams icon + platform-union entries"
```

---

### Task 13: TeamsConnectDialog + Settings card

**Files:**
- Create: `apps/electron/src/renderer/components/messaging/TeamsConnectDialog.tsx`
- Modify: `apps/electron/src/renderer/pages/settings/MessagingSettingsPage.tsx`

**Interfaces:**
- Consumes: `window.electronAPI.testTeamsCredentials`, `saveTeamsCredentials` (Task 11); i18n keys (Task 14).
- Produces: `<TeamsConnectDialog open onOpenChange reconfigure onSaved? />`.

- [ ] **Step 1: Create the dialog**

Create `apps/electron/src/renderer/components/messaging/TeamsConnectDialog.tsx` (adapted from `DiscordConnectDialog`, with App ID / App Password / optional Tenant ID / tunnel-mode toggle + BYO URL, and a post-save endpoint display):

```tsx
/**
 * TeamsConnectDialog — Azure Bot (Bot Framework) connect flow for Microsoft
 * Teams. Collects App ID + App Password (+ optional Tenant ID) and a public
 * messaging endpoint (bring-your-own HTTPS URL). After save, shows the
 * `<publicUrl>/api/messages` value to paste into the Azure Bot resource.
 */
import * as React from 'react'
import { Check, X, Copy } from 'lucide-react'
import { useTranslation } from 'react-i18next'
import { toast } from 'sonner'
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription, DialogFooter,
} from '@/components/ui/dialog'
import { Button } from '@/components/ui/button'
import { Spinner } from '@craft-agent/ui'
import { SettingsSecretInput } from '@/components/settings'

interface TeamsConnectDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  reconfigure?: boolean
  onSaved?: () => void
}

type TestResult =
  | { state: 'idle' } | { state: 'testing' } | { state: 'success' } | { state: 'error'; error: string }

export function TeamsConnectDialog({ open, onOpenChange, reconfigure = false, onSaved }: TeamsConnectDialogProps) {
  const { t } = useTranslation()
  const [appId, setAppId] = React.useState('')
  const [appPassword, setAppPassword] = React.useState('')
  const [tenantId, setTenantId] = React.useState('')
  const [byoUrl, setByoUrl] = React.useState('')
  const [saving, setSaving] = React.useState(false)
  const [test, setTest] = React.useState<TestResult>({ state: 'idle' })
  const [endpoint, setEndpoint] = React.useState<string | null>(null)

  React.useEffect(() => {
    if (!open) {
      setAppId(''); setAppPassword(''); setTenantId(''); setByoUrl('')
      setTest({ state: 'idle' }); setSaving(false); setEndpoint(null)
    }
  }, [open])

  const ready = appId.trim().length > 0 && appPassword.trim().length > 0
  const canSave = ready && byoUrl.trim().length > 0 && test.state === 'success'

  const handleTest = async () => {
    if (!ready) return
    setTest({ state: 'testing' })
    try {
      const result = await window.electronAPI.testTeamsCredentials({
        appId: appId.trim(), appPassword: appPassword.trim(),
        tenantId: tenantId.trim() || undefined,
      })
      setTest(result.success ? { state: 'success' } : { state: 'error', error: result.error ?? t('common.error') })
    } catch (err) {
      setTest({ state: 'error', error: err instanceof Error ? err.message : t('common.error') })
    }
  }

  const handleSave = async () => {
    if (!canSave) return
    setSaving(true)
    try {
      const res = await window.electronAPI.saveTeamsCredentials({
        appId: appId.trim(), appPassword: appPassword.trim(),
        tenantId: tenantId.trim() || undefined,
        tunnelMode: 'byo', byoUrl: byoUrl.trim(),
      })
      setEndpoint(res.messagingEndpoint)
      toast.success(t('settings.messaging.teams.saved'))
      onSaved?.()
    } catch (err) {
      toast.error(err instanceof Error ? err.message : t('settings.messaging.teams.saveFailed'))
    } finally {
      setSaving(false)
    }
  }

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-[520px]">
        <DialogHeader>
          <DialogTitle>
            {reconfigure ? t('settings.messaging.teams.reconfigureTitle') : t('settings.messaging.teams.connectTitle')}
          </DialogTitle>
          <DialogDescription className="whitespace-pre-line">
            {t('settings.messaging.teams.instructions')}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-3 py-2">
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.appIdLabel')}</div>
            <SettingsSecretInput value={appId} onChange={setAppId} placeholder={t('settings.messaging.teams.appIdPlaceholder')} disabled={saving} />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.appPasswordLabel')}</div>
            <SettingsSecretInput value={appPassword} onChange={setAppPassword} placeholder={t('settings.messaging.teams.appPasswordPlaceholder')} disabled={saving} />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.tenantIdLabel')}</div>
            <SettingsSecretInput value={tenantId} onChange={setTenantId} placeholder={t('settings.messaging.teams.tenantIdPlaceholder')} disabled={saving} />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.byoUrlLabel')}</div>
            <SettingsSecretInput value={byoUrl} onChange={setByoUrl} placeholder={t('settings.messaging.teams.byoUrlPlaceholder')} disabled={saving} />
          </div>

          <div className="flex items-center gap-2">
            <Button variant="outline" size="sm" onClick={handleTest} disabled={!ready || test.state === 'testing' || saving}>
              {test.state === 'testing' && <Spinner className="mr-1 text-[14px]" />}
              {t('settings.messaging.teams.testConnection')}
            </Button>
            {test.state === 'success' && (
              <span className="inline-flex items-center gap-1 text-xs text-emerald-600 dark:text-emerald-400">
                <Check className="h-3.5 w-3.5" />{t('settings.messaging.teams.testOk')}
              </span>
            )}
            {test.state === 'error' && (
              <span className="inline-flex items-center gap-1 text-xs text-destructive">
                <X className="h-3.5 w-3.5" />{test.error}
              </span>
            )}
          </div>

          {endpoint && (
            <div className="rounded-md border border-border bg-muted/40 p-2 text-xs">
              <div className="mb-1 font-medium">{t('settings.messaging.teams.endpointTitle')}</div>
              <div className="flex items-center gap-2">
                <code className="truncate">{endpoint}</code>
                <Button variant="ghost" size="icon" className="h-6 w-6" onClick={() => { void navigator.clipboard.writeText(endpoint); toast.success(t('common.copied')) }}>
                  <Copy className="h-3.5 w-3.5" />
                </Button>
              </div>
              <div className="mt-1 text-muted-foreground">{t('settings.messaging.teams.endpointHint')}</div>
            </div>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={() => onOpenChange(false)} disabled={saving}>
            {endpoint ? t('common.done') : t('common.cancel')}
          </Button>
          {!endpoint && (
            <Button variant="outline" size="sm" onClick={handleSave} disabled={!canSave || saving}>
              {saving && <Spinner className="mr-1 text-[14px]" />}
              {t('settings.messaging.teams.save')}
            </Button>
          )}
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}
```

> If `common.copied` / `common.done` keys don't exist, reuse existing common keys or add them alongside the Teams keys in Task 14.

- [ ] **Step 2: Add the Settings card + dialog mount**

In `MessagingSettingsPage.tsx`:

Import (after the Discord import, line ~54):
```ts
import { TeamsConnectDialog } from '@/components/messaging/TeamsConnectDialog'
```
Add the platform card (after the Discord `SettingsCard`, line ~127-129):
```tsx
  <SettingsCard>
    <PlatformRow platform="teams" workspaceId={activeWorkspace.id} />
  </SettingsCard>
```
Extend the `Platform` union + label map (line ~142-149):
```ts
type Platform = 'telegram' | 'whatsapp' | 'lark' | 'discord' | 'teams'
```
```ts
  discord: 'settings.messaging.discord.title',
  teams: 'settings.messaging.teams.title',
```
Mount the dialog (after the `platform === 'discord'` block, line ~481-483):
```tsx
      {platform === 'teams' && (
        <TeamsConnectDialog open={connectOpen} onOpenChange={setConnectOpen} reconfigure={reconfigure} />
      )}
```

- [ ] **Step 3: Typecheck**

Run: `cd apps/electron && bun run typecheck`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add apps/electron/src/renderer/components/messaging/TeamsConnectDialog.tsx apps/electron/src/renderer/pages/settings/MessagingSettingsPage.tsx
git commit -m "feat(electron-ui): add Teams connect dialog + settings card"
```

---

### Task 14: i18n keys (all 7 locales)

**Files:**
- Modify: `packages/shared/src/i18n/locales/en.json` (+ `de.json`, `es.json`, `hu.json`, `ja.json`, `pl.json`, `zh-Hans.json`)

> Scope note: the Discord channel added **no** `doc-links.ts` / `word-lists.ts` entries (verified — no `discord` matches there), so this plan follows that precedent and does **not** touch those files.

**Interfaces:**
- Produces: `settings.messaging.teams.*` keys used by `TeamsConnectDialog` (Task 13) and the settings PlatformRow.

- [ ] **Step 1: Add the English keys**

Add to `en.json` (flat, dotted keys; they must remain sorted — placed between the `settings.messaging.lark.*`/`discord.*` group and `settings.messaging.telegram.*`):

```json
"settings.messaging.teams.apiType": "Bot Framework",
"settings.messaging.teams.appIdLabel": "App ID (Client ID)",
"settings.messaging.teams.appIdPlaceholder": "Microsoft App ID from Azure",
"settings.messaging.teams.appPasswordLabel": "App Password (Client Secret)",
"settings.messaging.teams.appPasswordPlaceholder": "Client secret value",
"settings.messaging.teams.byoUrlLabel": "Public HTTPS URL",
"settings.messaging.teams.byoUrlPlaceholder": "https://your-tunnel.example.com",
"settings.messaging.teams.connectTitle": "Connect Microsoft Teams",
"settings.messaging.teams.connected": "Connected to Teams",
"settings.messaging.teams.disconnected": "Teams disconnected",
"settings.messaging.teams.endpointHint": "Paste this into your Azure Bot resource's Messaging endpoint, then save.",
"settings.messaging.teams.endpointTitle": "Messaging endpoint",
"settings.messaging.teams.instructions": "1. In the Azure portal, create an Azure Bot resource and a Microsoft App (App ID + client secret)\n2. Copy the App ID, client secret value, and (for single-tenant apps) the Tenant ID\n3. Provide a public HTTPS URL that forwards to this app (a reverse proxy or a persistent tunnel)\n4. Test, then Save — copy the shown Messaging endpoint into the Azure Bot's configuration\n5. Add the Teams channel in Azure, install the bot, and DM it or @mention it in a channel",
"settings.messaging.teams.notConnected": "Not connected",
"settings.messaging.teams.reconfigureTitle": "Replace Teams credentials",
"settings.messaging.teams.save": "Save",
"settings.messaging.teams.saveFailed": "Could not save Teams credentials",
"settings.messaging.teams.saved": "Teams bot saved",
"settings.messaging.teams.tenantIdLabel": "Tenant ID (optional)",
"settings.messaging.teams.tenantIdPlaceholder": "Directory (tenant) ID for single-tenant apps",
"settings.messaging.teams.testConnection": "Test connection",
"settings.messaging.teams.testOk": "Credentials accepted",
"settings.messaging.teams.title": "Microsoft Teams"
```

If `common.copied` / `common.done` are absent, add them too (check first: grep `"common.done"` in en.json).

- [ ] **Step 2: Replicate in the other 6 locales**

For each of `de.json`, `es.json`, `hu.json`, `ja.json`, `pl.json`, `zh-Hans.json`: add the SAME keys, translated in the style of that locale's existing `settings.messaging.discord.*` entries. Keep "Microsoft Teams" / "Teams" / "Bot Framework" / "App ID" / "Tenant ID" in English. Keys must be present in ALL locales (parity) and sorted.

- [ ] **Step 3: Run i18n validation**

Run: `bun run validate:ci` (or the specific i18n check, e.g. `bun run i18n:check` if it exists — grep `package.json` scripts).
Expected: PASS — parity + sorted checks green.

- [ ] **Step 4: Commit**

```bash
git add packages/shared/src/i18n/locales
git commit -m "feat(i18n): add Microsoft Teams messaging strings across locales"
```

---

### Task 15: Phase 1 integration verification

**Files:** none (verification only).

- [ ] **Step 1: Typecheck all touched packages**

Run: `bun run typecheck` (root) or per package: `cd packages/messaging-gateway && bun run typecheck && cd ../shared && bun run typecheck && cd ../server-core && bun run typecheck && cd ../../apps/electron && bun run typecheck`
Expected: PASS everywhere.

- [ ] **Step 2: Run the full messaging-gateway + shared test suites**

Run: `cd packages/messaging-gateway && bun test && cd ../shared && bun test`
Expected: PASS.

- [ ] **Step 3: Root CI gate**

Run: `bun run validate:ci`
Expected: PASS.

- [ ] **Step 4: Commit (if any incidental fixes were needed)**

```bash
git commit -am "chore(teams): phase 1 verification fixes" || echo "nothing to commit"
```

**Phase 1 complete: Teams works end-to-end with a bring-your-own HTTPS URL (personal chat, text/edit/typing/Adaptive-Card buttons).**

---

## PHASE 2 — On-demand Azure Dev Tunnels (persistent, stable URL)

---

### Task 16: `DevTunnelBinaryProvisioner` (lazy download + cache)

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.test.ts`

**Interfaces:**
- Produces:
  - `resolveDevtunnelRid(platform: NodeJS.Platform, arch: string): string` (throws on unsupported)
  - `class DevTunnelBinaryProvisioner` constructed with `{ cacheDir: string; fetchImpl?: typeof fetch; extractZip?: (zip: Buffer, destDir: string) => Promise<string> }` exposing `ensure(): Promise<string>`

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.test.ts`:

```ts
import { describe, it, expect, afterEach } from 'bun:test'
import { mkdtempSync, rmSync, existsSync, writeFileSync, mkdirSync, chmodSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { resolveDevtunnelRid, DevTunnelBinaryProvisioner } from './devtunnel-binary'

const cleanups: Array<() => void> = []
afterEach(() => { for (const c of cleanups.splice(0)) { try { c() } catch { /* */ } } })
function tmp(): string { const d = mkdtempSync(join(tmpdir(), 'dt-')); cleanups.push(() => rmSync(d, { recursive: true, force: true })); return d }

describe('resolveDevtunnelRid', () => {
  it('maps linux x64', () => { expect(resolveDevtunnelRid('linux', 'x64')).toBe('linux-x64') })
  it('maps darwin arm64 to the zip rid', () => { expect(resolveDevtunnelRid('darwin', 'arm64')).toBe('osx-arm64-zip') })
  it('maps win32 x64', () => { expect(resolveDevtunnelRid('win32', 'x64')).toBe('win-x64') })
  it('throws on unsupported', () => { expect(() => resolveDevtunnelRid('linux', 'ppc64')).toThrow(/unsupported/i) })
})

describe('DevTunnelBinaryProvisioner', () => {
  it('returns a cached binary without downloading', async () => {
    const cacheDir = tmp()
    const rid = resolveDevtunnelRid(process.platform, process.arch)
    const dir = join(cacheDir, rid)
    mkdirSync(dir, { recursive: true })
    const bin = join(dir, process.platform === 'win32' ? 'devtunnel.exe' : 'devtunnel')
    writeFileSync(bin, 'x'); chmodSync(bin, 0o755)
    let fetched = false
    const p = new DevTunnelBinaryProvisioner({ cacheDir, fetchImpl: (async () => { fetched = true; return new Response('') }) as unknown as typeof fetch })
    expect(await p.ensure()).toBe(bin)
    expect(fetched).toBe(false)
  })

  it('downloads a raw (non-zip) binary when cache is empty', async () => {
    const cacheDir = tmp()
    const fetchImpl = (async () => new Response(Buffer.from([0x7f, 0x45, 0x4c, 0x46]))) as unknown as typeof fetch
    const p = new DevTunnelBinaryProvisioner({ cacheDir, fetchImpl })
    const bin = await p.ensure()
    expect(existsSync(bin)).toBe(true)
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/devtunnel-binary.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.ts`:

```ts
/**
 * DevTunnelBinaryProvisioner — fetches the Microsoft `devtunnel` CLI on demand
 * and caches it under `<cacheDir>/<rid>/`. Nothing is bundled in the app; the
 * download happens the first time a user connects Teams in devtunnel mode.
 */
import { existsSync, mkdirSync, writeFileSync, chmodSync } from 'node:fs'
import { join } from 'node:path'
import { Buffer } from 'node:buffer'

const DOWNLOAD_BASE = 'https://aka.ms/TunnelsCliDownload'

export function resolveDevtunnelRid(platform: NodeJS.Platform, arch: string): string {
  if (platform === 'linux' && arch === 'x64') return 'linux-x64'
  if (platform === 'darwin' && arch === 'arm64') return 'osx-arm64-zip'
  if (platform === 'darwin' && arch === 'x64') return 'osx-x64-zip'
  if (platform === 'win32' && arch === 'x64') return 'win-x64'
  throw new Error(`Unsupported platform/arch for devtunnel: ${platform}/${arch}`)
}

function isZip(buf: Buffer): boolean {
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b // 'PK'
}

export interface DevTunnelBinaryProvisionerOptions {
  cacheDir: string
  fetchImpl?: typeof fetch
  /** Extract a zip buffer, returning the path to the `devtunnel` executable inside destDir. */
  extractZip?: (zip: Buffer, destDir: string) => Promise<string>
}

export class DevTunnelBinaryProvisioner {
  private readonly cacheDir: string
  private readonly fetchImpl: typeof fetch
  private readonly extractZip?: (zip: Buffer, destDir: string) => Promise<string>

  constructor(opts: DevTunnelBinaryProvisionerOptions) {
    this.cacheDir = opts.cacheDir
    this.fetchImpl = opts.fetchImpl ?? fetch
    this.extractZip = opts.extractZip
  }

  async ensure(): Promise<string> {
    const rid = resolveDevtunnelRid(process.platform, process.arch)
    const destDir = join(this.cacheDir, rid)
    const exe = process.platform === 'win32' ? 'devtunnel.exe' : 'devtunnel'
    const binPath = join(destDir, exe)
    if (existsSync(binPath)) return binPath

    mkdirSync(destDir, { recursive: true })
    const res = await this.fetchImpl(`${DOWNLOAD_BASE}/${rid}`)
    if (!res.ok) throw new Error(`Failed to download devtunnel (${res.status})`)
    const buf = Buffer.from(await res.arrayBuffer())

    if (isZip(buf)) {
      if (!this.extractZip) throw new Error('devtunnel archive requires an extractZip implementation')
      const extracted = await this.extractZip(buf, destDir)
      chmodSync(extracted, 0o755)
      return extracted
    }
    writeFileSync(binPath, buf)
    if (process.platform !== 'win32') chmodSync(binPath, 0o755)
    return binPath
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/devtunnel-binary.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.ts packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel-binary.test.ts
git commit -m "feat(teams): add on-demand devtunnel binary provisioner"
```

---

### Task 17: `DevTunnelProvider` (spawn persistent tunnel, parse URL, restart)

**Files:**
- Create: `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.ts`
- Test: `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.test.ts`

**Interfaces:**
- Consumes: `TunnelProvider` (Task 5), `DevTunnelBinaryProvisioner` (Task 16).
- Produces:
  - `parseTunnelUrl(line: string): string | null` (extracts the `https://…devtunnels.ms` hosting URL)
  - `class DevTunnelProvider implements TunnelProvider` constructed with `{ binPath: string; tunnelId?: string; spawnImpl?: typeof import('node:child_process').spawn; onTunnelId?: (id: string) => void }`

> IMPLEMENTATION NOTE: the exact `devtunnel host` stdout format varies by CLI version. `parseTunnelUrl` targets the printed `https://<id>-<port>.<cluster>.devtunnels.ms` connect URL. The implementer MUST run `devtunnel host --allow-anonymous` once locally and confirm the regex matches the installed version's output before finalizing.

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.test.ts`:

```ts
import { describe, it, expect } from 'bun:test'
import { EventEmitter } from 'node:events'
import { parseTunnelUrl, DevTunnelProvider } from './devtunnel'

describe('parseTunnelUrl', () => {
  it('extracts a devtunnels.ms hosting URL', () => {
    const line = 'Connect via browser: https://abc123-3978.usw2.devtunnels.ms'
    expect(parseTunnelUrl(line)).toBe('https://abc123-3978.usw2.devtunnels.ms')
  })
  it('returns null for unrelated lines', () => {
    expect(parseTunnelUrl('Logged in as user@example.com')).toBeNull()
  })
})

describe('DevTunnelProvider', () => {
  it('resolves the public URL parsed from spawn stdout', async () => {
    const fakeChild = Object.assign(new EventEmitter(), {
      stdout: new EventEmitter(), stderr: new EventEmitter(), kill: () => {},
    })
    const spawnImpl = (() => fakeChild) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', tunnelId: 't1', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => {
      fakeChild.stdout.emit('data', Buffer.from('Connect via browser: https://t1-3978.usw2.devtunnels.ms\n'))
    }, 10)
    const { publicUrl } = await started
    expect(publicUrl).toBe('https://t1-3978.usw2.devtunnels.ms')
    expect(p.isRunning()).toBe(true)
    await p.stop()
  })
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/devtunnel.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Write minimal implementation**

Create `packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.ts`:

```ts
/**
 * DevTunnelProvider — hosts a persistent Azure Dev Tunnel so the public URL is
 * stable across restarts. Requires a prior one-time `devtunnel user login`.
 */
import { spawn as nodeSpawn, type ChildProcess } from 'node:child_process'
import type { TunnelProvider } from './index'

const URL_RE = /(https:\/\/[a-z0-9-]+\.[a-z0-9-]+\.devtunnels\.ms)/i

export function parseTunnelUrl(line: string): string | null {
  const m = URL_RE.exec(line)
  return m ? m[1]! : null
}

export interface DevTunnelProviderOptions {
  binPath: string
  /** Persistent tunnel id. When omitted, `devtunnel host` creates an anonymous session URL. */
  tunnelId?: string
  spawnImpl?: typeof nodeSpawn
  onUrl?: (url: string) => void
}

export class DevTunnelProvider implements TunnelProvider {
  private proc: ChildProcess | null = null
  private publicUrl: string | null = null
  private urlHandlers = new Set<(u: string) => void>()
  private readonly opts: DevTunnelProviderOptions
  private stopping = false

  constructor(opts: DevTunnelProviderOptions) {
    this.opts = opts
    if (opts.onUrl) this.urlHandlers.add(opts.onUrl)
  }

  start(localPort: number): Promise<{ publicUrl: string }> {
    const spawnImpl = this.opts.spawnImpl ?? nodeSpawn
    const args = ['host', '--allow-anonymous', '-p', String(localPort)]
    if (this.opts.tunnelId) args.push(this.opts.tunnelId)
    const proc = spawnImpl(this.opts.binPath, args, { stdio: ['ignore', 'pipe', 'pipe'] })
    this.proc = proc

    return new Promise((resolve, reject) => {
      const onData = (chunk: Buffer) => {
        for (const line of chunk.toString('utf8').split('\n')) {
          const url = parseTunnelUrl(line)
          if (url) {
            this.publicUrl = url
            for (const h of this.urlHandlers) h(url)
            resolve({ publicUrl: url })
          }
        }
      }
      proc.stdout?.on('data', onData)
      proc.on('exit', (code) => {
        this.proc = null
        if (!this.stopping && !this.publicUrl) reject(new Error(`devtunnel exited (${code}). Run "devtunnel user login" first.`))
      })
      proc.on('error', reject)
    })
  }

  async stop(): Promise<void> {
    this.stopping = true
    const proc = this.proc
    this.proc = null
    this.publicUrl = null
    if (proc) proc.kill()
  }

  onUrlChange(cb: (publicUrl: string) => void): void { this.urlHandlers.add(cb) }
  isRunning(): boolean { return this.proc !== null }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/adapters/teams/tunnel/devtunnel.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.ts packages/messaging-gateway/src/adapters/teams/tunnel/devtunnel.test.ts
git commit -m "feat(teams): add DevTunnelProvider (persistent tunnel host + URL parse)"
```

---

### Task 18: Wire devtunnel mode into registry + electron main + dialog

**Files:**
- Modify: `packages/messaging-gateway/src/registry.ts` (`tryConnectTeams`, options), `bootstrap.ts`
- Modify: `apps/electron/src/main/index.ts`
- Modify: `apps/electron/src/renderer/components/messaging/TeamsConnectDialog.tsx`
- Modify: i18n locales (tunnel-mode strings)

**Interfaces:**
- Consumes: `DevTunnelProvider` (Task 17), `DevTunnelBinaryProvisioner` (Task 16).
- Produces: `MessagingGatewayRegistryOptions.teams.devtunnelCacheDir`; devtunnel branch in `tryConnectTeams`; dialog tunnel-mode toggle.

- [ ] **Step 1: Registry options + devtunnel branch**

In `MessagingGatewayRegistryOptions.teams` (Task 8, Step 2), add:
```ts
    /** Cache dir for the on-demand devtunnel binary (Phase 2). */
    devtunnelCacheDir?: string
```
In `tryConnectTeams` (Task 8, Step 8), replace the "BYO tunnel only" guard with a mode switch:
```ts
      let tunnel: TunnelProvider
      if (cfg.tunnelMode === 'devtunnel') {
        if (!this.opts.teams?.devtunnelCacheDir) throw new Error('devtunnel cache dir is not configured on this host')
        const provisioner = new DevTunnelBinaryProvisioner({ cacheDir: this.opts.teams.devtunnelCacheDir })
        const binPath = await provisioner.ensure()
        tunnel = new DevTunnelProvider({ binPath, tunnelId: cfg.devtunnelId })
      } else {
        if (!cfg.byoUrl) throw new Error('A public HTTPS URL is required for bring-your-own tunnel mode')
        tunnel = new ByoTunnelProvider({ baseUrl: cfg.byoUrl })
      }
```
Add the imports at the top of `registry.ts`:
```ts
import { DevTunnelProvider } from './adapters/teams/tunnel/devtunnel'
import { DevTunnelBinaryProvisioner } from './adapters/teams/tunnel/devtunnel-binary'
```
In `saveTeamsCredentials` (Task 8, Step 7), relax the BYO-only validation so `tunnelMode === 'devtunnel'` is accepted (skip the `byoUrl` requirement when mode is devtunnel).

- [ ] **Step 2: bootstrap.ts + electron main pass the cache dir**

In `bootstrap.ts` teams options: add `devtunnelCacheDir?: string` and pass it through.
In `apps/electron/src/main/index.ts` teams options (Task 11, Step 2):
```ts
  teams: {
    localPort: 3978,
    devtunnelCacheDir: join(homedir(), '.craft-agent', 'devtunnel'),
  },
```

- [ ] **Step 3: Dialog tunnel-mode toggle**

In `TeamsConnectDialog.tsx`, add a `tunnelMode` state (`'byo' | 'devtunnel'`, default `'devtunnel'`). Render a two-option toggle. When `'byo'`, show the BYO URL field (existing); when `'devtunnel'`, hide it and show `t('settings.messaging.teams.devtunnelHint')` (explains the one-time `devtunnel user login`). Pass `tunnelMode` + `byoUrl` (only when byo) into `saveTeamsCredentials`. `canSave` requires `byoUrl` only in byo mode.

- [ ] **Step 4: i18n strings for tunnel mode**

Add to all locales: `settings.messaging.teams.tunnelModeLabel`, `.tunnelModeDevtunnel` ("Automatic (Dev Tunnel)"), `.tunnelModeByo` ("Bring your own URL"), `.devtunnelHint` ("Requires a one-time `devtunnel user login` in a terminal. The tunnel keeps a stable URL across restarts.").

- [ ] **Step 5: Verify**

Run: `cd packages/messaging-gateway && bun test && bun run typecheck && cd ../../apps/electron && bun run typecheck`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(teams): add on-demand devtunnel mode across registry, main, and dialog"
```

**Phase 2 complete: Teams connects from the desktop app with a stable Dev Tunnel URL and no bundled binary.**

---

## PHASE 3 — Channel + group surfaces (mention trigger)

The `TeamsAdapter` already emits `isDM` (`conversationType === 'personal'`) and
`mentionedBot`, and `BindingConfig.teamsChannelTrigger` already exists (Task 1).
This phase adds the router gate so bound channels/group chats honour the trigger.

---

### Task 19: Router gate for Teams channel/group trigger

**Files:**
- Modify: `packages/messaging-gateway/src/router.ts`
- Test: `packages/messaging-gateway/src/__tests__/router.test.ts` (add cases) or a new `router-teams.test.ts`

**Interfaces:**
- Consumes: `IncomingMessage.isDM`/`mentionedBot`, `BindingConfig.teamsChannelTrigger`.

- [ ] **Step 1: Write the failing test**

Create `packages/messaging-gateway/src/__tests__/router-teams.test.ts` modeled on the existing `router.test.ts` harness (copy its adapter/binding-store/session-manager mocks from `router.test.ts`; keep them identical). Add:

```ts
// (reuse the mock setup from router.test.ts: makeRouter(), a fake bound session)
it('ignores an un-mentioned Teams channel message when trigger=mention', async () => {
  const { router, adapter, routedSessions } = makeRouter({ teamsChannelTrigger: 'mention' })
  await router.route(adapter, teamsMsg({ isDM: false, mentionedBot: false, text: 'hello' }))
  expect(routedSessions).toHaveLength(0)
})

it('routes an un-mentioned Teams channel message when trigger=all', async () => {
  const { router, adapter, routedSessions } = makeRouter({ teamsChannelTrigger: 'all' })
  await router.route(adapter, teamsMsg({ isDM: false, mentionedBot: false, text: 'hello' }))
  expect(routedSessions).toHaveLength(1)
})

it('always routes a Teams personal (DM) message', async () => {
  const { router, adapter, routedSessions } = makeRouter({ teamsChannelTrigger: 'mention' })
  await router.route(adapter, teamsMsg({ isDM: true, mentionedBot: false, text: 'hello' }))
  expect(routedSessions).toHaveLength(1)
})
```

(Provide `teamsMsg(overrides)` returning an `IncomingMessage` with `platform: 'teams'`, and `makeRouter({ teamsChannelTrigger })` that binds a session whose `config.teamsChannelTrigger` is set. Follow the exact shapes used in `router.test.ts`.)

- [ ] **Step 2: Run test to verify it fails**

Run: `cd packages/messaging-gateway && bun test src/__tests__/router-teams.test.ts`
Expected: FAIL — un-mentioned message is currently routed (no gate yet).

- [ ] **Step 3: Add the gate**

In `router.ts`, right after the existing Discord guild-trigger gate (line ~84, inside `if (binding) { … }`), add:

```ts
    // Teams channel/group-trigger gate: mirror the Discord gate. Personal
    // chats (isDM) always route; channel/group messages route only when the
    // bot is @mentioned unless the binding opts into 'all'.
    if (
      msg.platform === 'teams' &&
      msg.isDM === false &&
      msg.mentionedBot !== true &&
      binding.config.teamsChannelTrigger !== 'all'
    ) {
      this.log.info('ignoring un-mentioned teams channel message', {
        event: 'teams_trigger_skipped',
        channelId: msg.channelId,
        sessionId: binding.sessionId,
        bindingId: binding.id,
      })
      return
    }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cd packages/messaging-gateway && bun test src/__tests__/router-teams.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add packages/messaging-gateway/src/router.ts packages/messaging-gateway/src/__tests__/router-teams.test.ts
git commit -m "feat(teams): gate channel/group messages behind teamsChannelTrigger"
```

---

### Task 20: Final verification + PR

**Files:** none (verification + delivery).

- [ ] **Step 1: Full typecheck**

Run: `bun run typecheck`
Expected: PASS across all packages/apps.

- [ ] **Step 2: Full test suites for touched packages**

Run: `cd packages/messaging-gateway && bun test && cd ../shared && bun test && cd ../server-core && bun test`
Expected: PASS.

- [ ] **Step 3: Root CI gate**

Run: `bun run validate:ci`
Expected: PASS (typecheck, lint, i18n parity/sorted, coverage).

- [ ] **Step 4: Manual smoke checklist (documented, not automated)**

- Create an Azure Bot + Microsoft App; note App ID, secret, (tenant).
- In Settings → Messaging → Microsoft Teams: enter creds, choose a tunnel mode, Test → success, Save → endpoint shown.
- Paste the endpoint into Azure Bot config; add the Teams channel; install the bot.
- DM the bot: `/new`, then send a message → session responds.
- In a channel: @mention the bot → routes; plain message → ignored (trigger=mention).
- Tap an Adaptive Card approval button → approval registered.

- [ ] **Step 5: Open the PR**

```bash
git push -u origin <branch>
gh pr create --title "feat: add Microsoft Teams as a connected messaging channel" --body "Implements docs/superpowers/specs/2026-07-09-ms-teams-adapter-design.md. Phases 1-3: in-process Bot Framework adapter (personal + channel + group), BYO URL + on-demand Dev Tunnel, Adaptive Card approvals."
```

---

## Notes for the implementer

- **botbuilder under Bun:** `botbuilder` is pure JS; if a Bun-specific import issue arises, the `defaultCloudAdapterFactory` uses `require('botbuilder')` lazily so tests (which inject a fake) never load it. Verify a real connect under the Electron/Node runtime, not just `bun test`.
- **`CloudAdapter.process` req/res shape:** the `http-shim` targets the `{ body, headers, method }` + `{ status/send/end }` surface. If the installed `botbuilder` version expects a different response contract, adjust `BotResponse` accordingly (keep the `finished` promise).
- **devtunnel stdout parsing (Task 17):** confirm the URL regex against the installed CLI version before shipping Phase 2.
- **Security:** the adapter's HTTP listener binds `127.0.0.1` only; the tunnel is the sole public surface, and `CloudAdapter.process` validates the Bot Connector JWT on every inbound activity.
