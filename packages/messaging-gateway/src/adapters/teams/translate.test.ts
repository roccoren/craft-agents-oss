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
