/**
 * DevTunnelResourceProvisioner — ensures a *persistent* Azure Dev Tunnel
 * resource exists, returning its stable tunnel id.
 *
 * Root-cause context: `devtunnel host` invoked with NO tunnel id creates a
 * brand-new *temporary* tunnel — with a fresh random public URL — on every
 * single invocation. Nothing in this codebase previously called
 * `devtunnel create`, so every app restart silently generated a new URL,
 * leaving the Azure Bot's "Messaging endpoint" pointing at a stale, dead
 * tunnel (symptom: bot goes completely silent, including `/new`, until the
 * user manually re-pastes the new endpoint into Azure).
 *
 * The fix: create the tunnel resource once via `devtunnel create --json`,
 * persist the returned id (see registry.ts `tryConnectTeams`), and always
 * host with that same id thereafter — `devtunnel host <id>` reuses the
 * existing resource and its stable `https://<id>-<port>.<region>.devtunnels.ms`
 * URL shape across restarts.
 */
import { spawn as nodeSpawn } from 'node:child_process'

/** Parse the tunnel id out of `devtunnel create --json` stdout. Tolerates
 * both the real CLI shape (`{ tunnel: { tunnelId } }`, confirmed against a
 * live Windows devtunnel run) and the `id` variants Microsoft's docs samples
 * show, since exact CLI output has varied across versions/platforms. */
export function parseTunnelId(jsonOutput: string): string | null {
  try {
    const parsed = JSON.parse(jsonOutput) as {
      tunnel?: { tunnelId?: string; id?: string }
      tunnelId?: string
      id?: string
    }
    return parsed.tunnel?.tunnelId ?? parsed.tunnel?.id ?? parsed.tunnelId ?? parsed.id ?? null
  } catch {
    return null
  }
}

export interface DevTunnelResourceProvisionerOptions {
  binPath: string
  spawnImpl?: typeof nodeSpawn
}

export class DevTunnelResourceProvisioner {
  private readonly binPath: string
  private readonly spawnImpl: typeof nodeSpawn

  constructor(opts: DevTunnelResourceProvisionerOptions) {
    this.binPath = opts.binPath
    this.spawnImpl = opts.spawnImpl ?? nodeSpawn
  }

  /**
   * Run `devtunnel create --allow-anonymous --json`, returning the new
   * tunnel's stable id. Requires a prior one-time `devtunnel user login`
   * (surfaced to the caller as a rejected promise if missing).
   */
  create(): Promise<string> {
    return new Promise((resolve, reject) => {
      const proc = this.spawnImpl(this.binPath, ['create', '--allow-anonymous', '--json'], {
        stdio: ['ignore', 'pipe', 'pipe'],
      })
      let stdout = ''
      let stderr = ''
      proc.stdout?.on('data', (chunk: Buffer) => { stdout += chunk.toString('utf8') })
      proc.stderr?.on('data', (chunk: Buffer) => { stderr += chunk.toString('utf8') })
      proc.on('error', reject)
      proc.on('exit', (code) => {
        if (code !== 0) {
          reject(new Error(`devtunnel create failed (exit ${code}): ${stderr.trim() || 'no output'}`))
          return
        }
        const id = parseTunnelId(stdout)
        if (!id) {
          reject(new Error(`devtunnel create succeeded but no tunnel id found in output: ${stdout.trim()}`))
          return
        }
        resolve(id)
      })
    })
  }
}
