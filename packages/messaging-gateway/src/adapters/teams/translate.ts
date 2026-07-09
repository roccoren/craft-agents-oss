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
