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
