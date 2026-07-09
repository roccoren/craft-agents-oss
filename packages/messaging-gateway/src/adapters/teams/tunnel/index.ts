/**
 * TunnelProvider — supplies the public HTTPS base URL that Azure/Teams uses to
 * reach the adapter's local HTTP listener.
 *
 *  - `ByoTunnelProvider` (Phase 1): user-supplied stable URL, no process.
 *  - `DevTunnelProvider` (Phase 2): spawns a persistent Azure Dev Tunnel.
 */
export interface TunnelProvider {
  /** Start the tunnel pointing at `localPort`; resolves the public base URL. */
  start(localPort: number): Promise<{ publicUrl: string }>
  stop(): Promise<void>
  onUrlChange(cb: (publicUrl: string) => void): void
  isRunning(): boolean
}

export function normalizeBaseUrl(url: string): string {
  const trimmed = url.trim().replace(/\/+$/, '')
  if (!/^https:\/\//i.test(trimmed)) {
    throw new Error('Messaging endpoint base URL must start with https://')
  }
  return trimmed
}

export class ByoTunnelProvider implements TunnelProvider {
  private readonly baseUrl: string
  private running = false

  constructor(opts: { baseUrl: string }) {
    this.baseUrl = normalizeBaseUrl(opts.baseUrl)
  }

  async start(_localPort: number): Promise<{ publicUrl: string }> {
    this.running = true
    return { publicUrl: this.baseUrl }
  }

  async stop(): Promise<void> {
    this.running = false
  }

  onUrlChange(_cb: (publicUrl: string) => void): void {
    // BYO URL never changes at runtime.
  }

  isRunning(): boolean {
    return this.running
  }
}
