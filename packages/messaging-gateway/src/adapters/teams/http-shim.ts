/**
 * Minimal adapters between Node's `http` server objects and the request/response
 * shape `botbuilder`'s `CloudAdapter.process(req, res, logic)` expects.
 *
 * `CloudAdapter.process` reads `req.body`, `req.headers`, `req.method`, and
 * writes via `res.status(code)`, `res.send(body)`, `res.end()`. We don't hand it
 * a raw Node response (it lacks `.status()/.send()`), so we capture the outcome
 * in-memory and flush it to the real Node response ourselves.
 */
import type { IncomingMessage } from 'node:http'

export interface BotRequest {
  body: unknown
  headers: Record<string, string>
  method?: string
}

export function toBotRequest(req: IncomingMessage, body: unknown): BotRequest {
  const headers: Record<string, string> = {}
  for (const [k, v] of Object.entries(req.headers)) {
    if (typeof v === 'string') headers[k.toLowerCase()] = v
    else if (Array.isArray(v)) headers[k.toLowerCase()] = v.join(', ')
  }
  return { body, headers, method: req.method }
}

export class BotResponse {
  private statusCode = 200
  private payload: unknown = undefined
  private resolveFinished!: (r: { status: number; body?: unknown }) => void
  readonly finished: Promise<{ status: number; body?: unknown }>

  constructor() {
    this.finished = new Promise((resolve) => {
      this.resolveFinished = resolve
    })
  }

  status(code: number): this {
    this.statusCode = code
    return this
  }

  header(_key: string, _value: string): this {
    return this
  }

  send(body?: unknown): this {
    if (body !== undefined) this.payload = body
    return this
  }

  end(): void {
    this.resolveFinished({ status: this.statusCode, body: this.payload })
  }
}
