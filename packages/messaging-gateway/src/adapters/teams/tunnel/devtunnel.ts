/**
 * DevTunnelProvider — hosts a persistent Azure Dev Tunnel so the public URL is
 * stable across restarts. Requires a prior one-time `devtunnel user login`.
 *
 * NOTE: the exact `devtunnel host` stdout format varies by CLI version.
 * `parseTunnelUrl` targets the printed `https://<id>.<cluster>.devtunnels.ms`
 * connect URL; verify against the installed CLI before shipping.
 */
import { spawn as nodeSpawn, type ChildProcess } from 'node:child_process'
import type { TunnelProvider } from './index'

const URL_RE = /(https:\/\/[a-z0-9-]+\.[a-z0-9-]+\.devtunnels\.ms)/i

export function parseTunnelUrl(line: string): string | null {
  const m = URL_RE.exec(line)
  return m ? m[1]! : null
}

export interface DevTunnelProviderOptions {
  binPath: string
  /** Persistent tunnel id. When omitted, `devtunnel host` creates an anonymous session URL. */
  tunnelId?: string
  spawnImpl?: typeof nodeSpawn
  onUrl?: (url: string) => void
}

export class DevTunnelProvider implements TunnelProvider {
  private proc: ChildProcess | null = null
  private publicUrl: string | null = null
  private urlHandlers = new Set<(u: string) => void>()
  private readonly opts: DevTunnelProviderOptions
  private stopping = false

  constructor(opts: DevTunnelProviderOptions) {
    this.opts = opts
    if (opts.onUrl) this.urlHandlers.add(opts.onUrl)
  }

  start(localPort: number): Promise<{ publicUrl: string }> {
    const spawnImpl = this.opts.spawnImpl ?? nodeSpawn
    const args = ['host', '--allow-anonymous', '-p', String(localPort)]
    if (this.opts.tunnelId) args.push(this.opts.tunnelId)
    const proc = spawnImpl(this.opts.binPath, args, { stdio: ['ignore', 'pipe', 'pipe'] })
    this.proc = proc

    return new Promise((resolve, reject) => {
      let stderr = ''
      const onData = (chunk: Buffer) => {
        for (const line of chunk.toString('utf8').split('\n')) {
          const url = parseTunnelUrl(line)
          if (url) {
            this.publicUrl = url
            for (const h of this.urlHandlers) h(url)
            resolve({ publicUrl: url })
          }
        }
      }
      proc.stdout?.on('data', onData)
      proc.stderr?.on('data', (chunk: Buffer) => { stderr += chunk.toString('utf8') })
      proc.on('exit', (code) => {
        this.proc = null
        if (!this.stopping && !this.publicUrl) {
          // Surface the CLI's own stderr instead of guessing why it failed —
          // a hardcoded "run devtunnel user login" message was actively
          // misleading when the real cause was unrelated (e.g. an argument
          // or port-forwarding error), and the user had already logged in.
          const detail = stderr.trim()
          reject(new Error(
            detail
              ? `devtunnel exited (${code}): ${detail}`
              : `devtunnel exited (${code}) with no output. If this persists, try "devtunnel user login".`,
          ))
        }
      })
      proc.on('error', reject)
    })
  }

  async stop(): Promise<void> {
    this.stopping = true
    const proc = this.proc
    this.proc = null
    this.publicUrl = null
    if (proc) proc.kill()
  }

  onUrlChange(cb: (publicUrl: string) => void): void { this.urlHandlers.add(cb) }
  isRunning(): boolean { return this.proc !== null }
}
