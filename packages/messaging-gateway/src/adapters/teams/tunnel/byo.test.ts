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
