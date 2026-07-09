import { describe, it, expect, afterEach } from 'bun:test'
import { mkdtempSync, rmSync, existsSync, writeFileSync, mkdirSync, chmodSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { resolveDevtunnelRid, DevTunnelBinaryProvisioner } from './devtunnel-binary'

const cleanups: Array<() => void> = []
afterEach(() => { for (const c of cleanups.splice(0)) { try { c() } catch { /* */ } } })
function tmp(): string { const d = mkdtempSync(join(tmpdir(), 'dt-')); cleanups.push(() => rmSync(d, { recursive: true, force: true })); return d }

describe('resolveDevtunnelRid', () => {
  it('maps linux x64', () => { expect(resolveDevtunnelRid('linux', 'x64')).toBe('linux-x64') })
  it('maps darwin arm64 to the zip rid', () => { expect(resolveDevtunnelRid('darwin', 'arm64')).toBe('osx-arm64-zip') })
  it('maps win32 x64', () => { expect(resolveDevtunnelRid('win32', 'x64')).toBe('win-x64') })
  it('throws on unsupported', () => { expect(() => resolveDevtunnelRid('linux', 'ppc64')).toThrow(/unsupported/i) })
})

describe('DevTunnelBinaryProvisioner', () => {
  it('returns a cached binary without downloading', async () => {
    const cacheDir = tmp()
    const rid = resolveDevtunnelRid(process.platform, process.arch)
    const dir = join(cacheDir, rid)
    mkdirSync(dir, { recursive: true })
    const bin = join(dir, process.platform === 'win32' ? 'devtunnel.exe' : 'devtunnel')
    writeFileSync(bin, 'x'); chmodSync(bin, 0o755)
    let fetched = false
    const p = new DevTunnelBinaryProvisioner({ cacheDir, fetchImpl: (async () => { fetched = true; return new Response('') }) as unknown as typeof fetch })
    expect(await p.ensure()).toBe(bin)
    expect(fetched).toBe(false)
  })

  it('downloads a raw (non-zip) binary when cache is empty', async () => {
    const cacheDir = tmp()
    const fetchImpl = (async () => new Response(Buffer.from([0x7f, 0x45, 0x4c, 0x46]))) as unknown as typeof fetch
    const p = new DevTunnelBinaryProvisioner({ cacheDir, fetchImpl })
    const bin = await p.ensure()
    expect(existsSync(bin)).toBe(true)
  })

  it('rejects a zip archive without an extractZip implementation', async () => {
    const cacheDir = tmp()
    const fetchImpl = (async () => new Response(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) as unknown as typeof fetch
    const p = new DevTunnelBinaryProvisioner({ cacheDir, fetchImpl })
    await expect(p.ensure()).rejects.toThrow(/extractZip/)
  })
})
