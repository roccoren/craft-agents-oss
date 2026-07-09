import { describe, it, expect } from 'bun:test'
import { BotResponse, toBotRequest } from './http-shim'
import type { IncomingMessage } from 'node:http'

describe('toBotRequest', () => {
  it('lowercases headers and attaches the parsed body', () => {
    const fake = { headers: { Authorization: 'Bearer x' }, method: 'POST' } as unknown as IncomingMessage
    const r = toBotRequest(fake, { type: 'message' })
    expect(r.headers.authorization).toBe('Bearer x')
    expect(r.method).toBe('POST')
    expect(r.body).toEqual({ type: 'message' })
  })
})

describe('BotResponse', () => {
  it('captures status + body and resolves finished on end()', async () => {
    const res = new BotResponse()
    res.status(201)
    res.send({ id: 'abc' })
    res.end()
    await expect(res.finished).resolves.toEqual({ status: 201, body: { id: 'abc' } })
  })

  it('defaults status to 200 when only end() is called', async () => {
    const res = new BotResponse()
    res.end()
    await expect(res.finished).resolves.toEqual({ status: 200, body: undefined })
  })
})
