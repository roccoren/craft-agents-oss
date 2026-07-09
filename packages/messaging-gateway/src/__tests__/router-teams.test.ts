/**
 * Router tests for the Teams channel/group-trigger gate.
 *
 * Personal (DM) messages always route; channel/group messages route only when
 * the bot is @mentioned unless the binding opts into `teamsChannelTrigger: 'all'`.
 */
import { afterEach, beforeEach, describe, expect, it, mock } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { Router } from '../router'
import { BindingStore } from '../binding-store'
import type { Commands } from '../commands'
import type { IncomingMessage, PlatformAdapter } from '../types'

let storeDir: string

beforeEach(() => {
  storeDir = mkdtempSync(join(tmpdir(), 'router-teams-store-'))
})
afterEach(() => {
  rmSync(storeDir, { recursive: true, force: true })
})

function teamsMsg(overrides: Partial<IncomingMessage> = {}): IncomingMessage {
  return {
    platform: 'teams',
    channelId: 'chan-1',
    messageId: '1',
    senderId: 'user-1',
    text: 'hello',
    timestamp: Date.now(),
    raw: {},
    ...overrides,
  }
}

function makeFakeAdapter(): PlatformAdapter {
  const noop = async () => {
    throw new Error('unused')
  }
  return {
    platform: 'teams',
    capabilities: {
      messageEditing: true,
      inlineButtons: true,
      maxButtons: 6,
      maxMessageLength: 20000,
      markdown: 'teams',
      webhookSupport: false,
    },
    initialize: noop,
    destroy: noop,
    isConnected: () => true,
    onMessage: () => {},
    onButtonPress: () => {},
    sendText: mock(async () => ({ platform: 'teams', channelId: 'chan-1', messageId: 'm' })),
    editMessage: noop,
    sendButtons: noop,
    sendTyping: async () => {},
    sendFile: noop,
  } as unknown as PlatformAdapter
}

function makeRouter(trigger: 'mention' | 'all') {
  const store = new BindingStore(storeDir)
  store.bind('ws1', 'sess-A', 'teams', 'chan-1', undefined, { teamsChannelTrigger: trigger })
  const sessionManager = { sendMessage: mock(async () => {}) }
  const commands = { handle: mock(async () => {}) }
  const router = new Router(
    sessionManager as unknown as never,
    store,
    commands as unknown as Commands,
  )
  return { router, sessionManager }
}

describe('Router — Teams trigger gate', () => {
  it('ignores an un-mentioned channel message when trigger=mention', async () => {
    const { router, sessionManager } = makeRouter('mention')
    await router.route(makeFakeAdapter(), teamsMsg({ isDM: false, mentionedBot: false }))
    expect(sessionManager.sendMessage).toHaveBeenCalledTimes(0)
  })

  it('routes an @mentioned channel message when trigger=mention', async () => {
    const { router, sessionManager } = makeRouter('mention')
    await router.route(makeFakeAdapter(), teamsMsg({ isDM: false, mentionedBot: true }))
    expect(sessionManager.sendMessage).toHaveBeenCalledTimes(1)
  })

  it('routes an un-mentioned channel message when trigger=all', async () => {
    const { router, sessionManager } = makeRouter('all')
    await router.route(makeFakeAdapter(), teamsMsg({ isDM: false, mentionedBot: false }))
    expect(sessionManager.sendMessage).toHaveBeenCalledTimes(1)
  })

  it('always routes a personal (DM) message regardless of mention', async () => {
    const { router, sessionManager } = makeRouter('mention')
    await router.route(makeFakeAdapter(), teamsMsg({ isDM: true, mentionedBot: false }))
    expect(sessionManager.sendMessage).toHaveBeenCalledTimes(1)
  })
})
