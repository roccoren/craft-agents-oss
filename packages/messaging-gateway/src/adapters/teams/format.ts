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
