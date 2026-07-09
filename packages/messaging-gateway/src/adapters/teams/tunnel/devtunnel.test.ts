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
    expect(p.isRunning()).toBe(false)
  })

  it('rejects when the process exits before a URL is seen', async () => {
    const fakeChild = Object.assign(new EventEmitter(), {
      stdout: new EventEmitter(), stderr: new EventEmitter(), kill: () => {},
    })
    const spawnImpl = (() => fakeChild) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => { fakeChild.emit('exit', 1) }, 10)
    await expect(started).rejects.toThrow(/devtunnel exited/)
  })

  it('includes the real stderr text in the rejection instead of a hardcoded guess', async () => {
    const fakeChild = Object.assign(new EventEmitter(), {
      stdout: new EventEmitter(), stderr: new EventEmitter(), kill: () => {},
    })
    const spawnImpl = (() => fakeChild) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', tunnelId: 't1', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => {
      fakeChild.stderr.emit('data', Buffer.from('Error: port 3978 is already forwarded on this tunnel\n'))
      fakeChild.emit('exit', 1)
    }, 10)
    await expect(started).rejects.toThrow(/port 3978 is already forwarded/)
  })
})
