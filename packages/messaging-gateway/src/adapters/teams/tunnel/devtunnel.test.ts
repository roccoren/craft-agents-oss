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

function fakeChild(): EventEmitter & { stdout: EventEmitter; stderr: EventEmitter; kill: () => void } {
  return Object.assign(new EventEmitter(), { stdout: new EventEmitter(), stderr: new EventEmitter(), kill: () => {} })
}

describe('DevTunnelProvider — no tunnelId (ad-hoc/temporary session)', () => {
  it('hosts directly with -p, no port-create step', async () => {
    const child = fakeChild()
    const calls: string[][] = []
    const spawnImpl = ((_bin: string, args: string[]) => { calls.push(args); return child }) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => {
      child.stdout.emit('data', Buffer.from('Connect via browser: https://t1-3978.usw2.devtunnels.ms\n'))
    }, 10)
    const { publicUrl } = await started
    expect(publicUrl).toBe('https://t1-3978.usw2.devtunnels.ms')
    expect(calls).toEqual([['host', '--allow-anonymous', '-p', '3978']])
    expect(p.isRunning()).toBe(true)
    await p.stop()
    expect(p.isRunning()).toBe(false)
  })

  it('rejects when the process exits before a URL is seen', async () => {
    const child = fakeChild()
    const spawnImpl = (() => child) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => { child.emit('exit', 1) }, 10)
    await expect(started).rejects.toThrow(/devtunnel exited/)
  })

  it('includes the real stderr text in the rejection instead of a hardcoded guess', async () => {
    const child = fakeChild()
    const spawnImpl = (() => child) as unknown as typeof import('node:child_process').spawn
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', spawnImpl })
    const started = p.start(3978)
    setTimeout(() => {
      child.stderr.emit('data', Buffer.from('Error: port 3978 is already forwarded on this tunnel\n'))
      child.emit('exit', 1)
    }, 10)
    await expect(started).rejects.toThrow(/port 3978 is already forwarded/)
  })
})

describe('DevTunnelProvider — persistent tunnelId', () => {
  /**
   * Real CLI behaviour (confirmed against a live devtunnel run): hosting a
   * previously-created (persistent) tunnel with `-p <port>` combined in the
   * same `host` invocation fails with "Batch update of ports is not
   * supported. Add, update, or delete ports individually instead." The port
   * must be registered via a separate `devtunnel port create` call first;
   * `host` is then invoked with no `-p` at all.
   */
  function twoPhaseSpawn(): {
    spawnImpl: typeof import('node:child_process').spawn
    portCreateChild: ReturnType<typeof fakeChild>
    hostChild: ReturnType<typeof fakeChild>
    calls: string[][]
  } {
    const portCreateChild = fakeChild()
    const hostChild = fakeChild()
    const calls: string[][] = []
    let call = 0
    const spawnImpl = ((_bin: string, args: string[]) => {
      calls.push(args)
      call += 1
      return call === 1 ? portCreateChild : hostChild
    }) as unknown as typeof import('node:child_process').spawn
    return { spawnImpl, portCreateChild, hostChild, calls }
  }

  it('registers the port individually before hosting, then hosts without -p', async () => {
    const { spawnImpl, portCreateChild, hostChild, calls } = twoPhaseSpawn()
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', tunnelId: 't1', spawnImpl })
    const started = p.start(3978)

    await new Promise((r) => setTimeout(r, 5))
    portCreateChild.emit('exit', 0)

    await new Promise((r) => setTimeout(r, 5))
    hostChild.stdout.emit('data', Buffer.from('Connect via browser: https://t1-3978.usw2.devtunnels.ms\n'))

    const { publicUrl } = await started
    expect(publicUrl).toBe('https://t1-3978.usw2.devtunnels.ms')
    expect(calls[0]).toEqual(['port', 'create', 't1', '-p', '3978', '--allow-anonymous'])
    expect(calls[1]).toEqual(['host', 't1'])
  })

  it('tolerates "port already exists" from a prior partial run and still hosts', async () => {
    const { spawnImpl, portCreateChild, hostChild, calls } = twoPhaseSpawn()
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', tunnelId: 't1', spawnImpl })
    const started = p.start(3978)

    await new Promise((r) => setTimeout(r, 5))
    portCreateChild.stderr.emit('data', Buffer.from('Error: a port with this number already exists\n'))
    portCreateChild.emit('exit', 1)

    await new Promise((r) => setTimeout(r, 5))
    hostChild.stdout.emit('data', Buffer.from('Connect via browser: https://t1-3978.usw2.devtunnels.ms\n'))

    const { publicUrl } = await started
    expect(publicUrl).toBe('https://t1-3978.usw2.devtunnels.ms')
    expect(calls[1]).toEqual(['host', 't1'])
  })

  it('rejects without ever hosting when port registration genuinely fails', async () => {
    const { spawnImpl, portCreateChild, calls } = twoPhaseSpawn()
    const p = new DevTunnelProvider({ binPath: '/fake/devtunnel', tunnelId: 't1', spawnImpl })
    const started = p.start(3978)

    await new Promise((r) => setTimeout(r, 5))
    portCreateChild.stderr.emit('data', Buffer.from('Error: not logged in\n'))
    portCreateChild.emit('exit', 1)

    await expect(started).rejects.toThrow(/port create failed.*not logged in/s)
    expect(calls.length).toBe(1) // host was never invoked
  })
})
