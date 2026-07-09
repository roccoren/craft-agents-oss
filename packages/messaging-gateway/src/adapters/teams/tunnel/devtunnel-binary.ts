/**
 * DevTunnelBinaryProvisioner — fetches the Microsoft `devtunnel` CLI on demand
 * and caches it under `<cacheDir>/<rid>/`. Nothing is bundled in the app; the
 * download happens the first time a user connects Teams in devtunnel mode.
 */
import { existsSync, mkdirSync, writeFileSync, chmodSync } from 'node:fs'
import { join } from 'node:path'
import { Buffer } from 'node:buffer'

const DOWNLOAD_BASE = 'https://aka.ms/TunnelsCliDownload'

export function resolveDevtunnelRid(platform: NodeJS.Platform, arch: string): string {
  if (platform === 'linux' && arch === 'x64') return 'linux-x64'
  if (platform === 'darwin' && arch === 'arm64') return 'osx-arm64-zip'
  if (platform === 'darwin' && arch === 'x64') return 'osx-x64-zip'
  if (platform === 'win32' && arch === 'x64') return 'win-x64'
  throw new Error(`Unsupported platform/arch for devtunnel: ${platform}/${arch}`)
}

function isZip(buf: Buffer): boolean {
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b // 'PK'
}

export interface DevTunnelBinaryProvisionerOptions {
  cacheDir: string
  fetchImpl?: typeof fetch
  /** Extract a zip buffer, returning the path to the `devtunnel` executable inside destDir. */
  extractZip?: (zip: Buffer, destDir: string) => Promise<string>
}

export class DevTunnelBinaryProvisioner {
  private readonly cacheDir: string
  private readonly fetchImpl: typeof fetch
  private readonly extractZip?: (zip: Buffer, destDir: string) => Promise<string>

  constructor(opts: DevTunnelBinaryProvisionerOptions) {
    this.cacheDir = opts.cacheDir
    this.fetchImpl = opts.fetchImpl ?? fetch
    this.extractZip = opts.extractZip
  }

  async ensure(): Promise<string> {
    const rid = resolveDevtunnelRid(process.platform, process.arch)
    const destDir = join(this.cacheDir, rid)
    const exe = process.platform === 'win32' ? 'devtunnel.exe' : 'devtunnel'
    const binPath = join(destDir, exe)
    if (existsSync(binPath)) return binPath

    mkdirSync(destDir, { recursive: true })
    const res = await this.fetchImpl(`${DOWNLOAD_BASE}/${rid}`)
    if (!res.ok) throw new Error(`Failed to download devtunnel (${res.status})`)
    const buf = Buffer.from(await res.arrayBuffer())

    if (isZip(buf)) {
      if (!this.extractZip) throw new Error('devtunnel archive requires an extractZip implementation')
      const extracted = await this.extractZip(buf, destDir)
      chmodSync(extracted, 0o755)
      return extracted
    }
    writeFileSync(binPath, buf)
    if (process.platform !== 'win32') chmodSync(binPath, 0o755)
    return binPath
  }
}
