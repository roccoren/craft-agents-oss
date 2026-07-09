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
