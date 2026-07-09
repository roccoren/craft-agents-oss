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
