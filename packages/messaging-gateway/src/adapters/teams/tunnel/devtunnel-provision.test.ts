import { describe, it, expect } from 'bun:test'
import { EventEmitter } from 'node:events'
import { parseTunnelId, DevTunnelResourceProvisioner } from './devtunnel-provision'

describe('parseTunnelId', () => {
  it('extracts the id from the real devtunnel CLI shape `{ tunnel: { tunnelId } }`', () => {
    // Verbatim payload observed from a real `devtunnel create --allow-anonymous
    // --json` run (Windows, devtunnel CLI, 2026-07). The field is `tunnelId`,
    // not `id` as Microsoft's own docs samples suggest -- this is the actual
    // shape that matters.
    const real = {
      tunnel: {
        tunnelId: 'sneaky-plane-vvjxm25.jpe1',
        hostConnections: 0,
        clientConnections: 0,
        labels: [],
        tunnelExpiration: '30 days',
        description: '',
        currentUploadRate: '0 MB/s (limit: 20 MB/s)',
        currentDownloadRate: '0 MB/s (limit: 20 MB/s)',
        accessControl: [{ type: 'Anonymous', subjects: [], scopes: ['connect'] }],
      },
    }
    expect(parseTunnelId(JSON.stringify(real))).toBe('sneaky-plane-vvjxm25.jpe1')
  })

  it('extracts the id from a `{ tunnel: { id } }` JSON shape', () => {
    expect(parseTunnelId(JSON.stringify({ tunnel: { id: 'abc123' } }))).toBe('abc123')
  })

  it('extracts the id from a flat `{ tunnelId }` JSON shape', () => {
    expect(parseTunnelId(JSON.stringify({ tunnelId: 'flat-tunnel-id' }))).toBe('flat-tunnel-id')
  })

  it('extracts the id from a flat `{ id }` JSON shape', () => {
    expect(parseTunnelId(JSON.stringify({ id: 'xyz789' }))).toBe('xyz789')
  })

  it('returns null for non-JSON output', () => {
    expect(parseTunnelId('not json')).toBeNull()
  })

  it('returns null when no id field is present', () => {
    expect(parseTunnelId(JSON.stringify({ tunnel: { name: 'foo' } }))).toBeNull()
  })
})

describe('DevTunnelResourceProvisioner', () => {
  function fakeChild(): EventEmitter & { stdout: EventEmitter; stderr: EventEmitter } {
    return Object.assign(new EventEmitter(), { stdout: new EventEmitter(), stderr: new EventEmitter() })
  }

  it('resolves the tunnel id parsed from `devtunnel create --json` stdout', async () => {
    const child = fakeChild()
    const spawnImpl = (() => child) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelResourceProvisioner({ binPath: '/fake/devtunnel', spawnImpl })
    const created = p.create()
    child.stdout.emit('data', Buffer.from(JSON.stringify({ tunnel: { id: 'persist-1' } })))
    child.emit('exit', 0)
    await expect(created).resolves.toBe('persist-1')
  })

  it('rejects when the process exits non-zero', async () => {
    const child = fakeChild()
    const spawnImpl = (() => child) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelResourceProvisioner({ binPath: '/fake/devtunnel', spawnImpl })
    const created = p.create()
    child.stderr.emit('data', Buffer.from('not logged in'))
    child.emit('exit', 1)
    await expect(created).rejects.toThrow(/devtunnel create failed/)
  })

  it('rejects when stdout has no parseable tunnel id', async () => {
    const child = fakeChild()
    const spawnImpl = (() => child) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelResourceProvisioner({ binPath: '/fake/devtunnel', spawnImpl })
    const created = p.create()
    child.stdout.emit('data', Buffer.from('{}'))
    child.emit('exit', 0)
    await expect(created).rejects.toThrow(/no tunnel id/)
  })
})
