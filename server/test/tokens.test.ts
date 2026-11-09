import { describe, expect, it } from 'bun:test'
import { deriveKeys } from '../src/keys'
import { issueToken, MAX_TOKEN_CHARS, newSid, verifyToken } from '../src/tokens'

const keys = await deriveKeys(new Uint8Array(32).fill(3))
const other = await deriveKeys(new Uint8Array(32).fill(4))
const docRef = 'a'.repeat(24), sid = newSid(), now = 1_000_000

async function issued() { return issueToken(keys.kToken, { docRef, sid, ttlSec: 300, now }) }
function mutatePayload(token: string, patch: Record<string, unknown>): string {
  const [h, p, s] = token.split('.')
  const claims = JSON.parse(atob(p!.replace(/-/g, '+').replace(/_/g, '/')))
  const np = btoa(JSON.stringify({ ...claims, ...patch })).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
  return `${h}.${np}.${s}`
}

describe('token contract (§6.2)', () => {
  it('round-trips with exactly the specified claims', async () => {
    const t = await issued()
    expect(t.token.length).toBeLessThanOrEqual(MAX_TOKEN_CHARS)
    const r = await verifyToken(keys.kToken, t.token, { now: now + 1, docRef, sid })
    expect(r.ok).toBe(true)
    if (r.ok) expect(Object.keys(r.claims).sort()).toEqual(['docRef', 'exp', 'iat', 'jti', 'scope', 'sid', 'v'])
    expect(sid).toMatch(/^[0-9a-f]{12}$/)
  })
  it('each check fails with its own reason', async () => {
    const { token } = await issued()
    const sigFlip = token.slice(0, -2) + (token.at(-2) === 'A' ? 'B' : 'A') + token.at(-1)
    expect(await verifyToken(keys.kToken, sigFlip, { now, docRef, sid })).toEqual({ ok: false, reason: 'signature' })
    expect(await verifyToken(other.kToken, token, { now, docRef, sid })).toEqual({ ok: false, reason: 'signature' })
    expect(await verifyToken(keys.kToken, mutatePayload(token, { exp: now + 999 }), { now, docRef, sid })).toEqual({ ok: false, reason: 'signature' })
    expect(await verifyToken(keys.kToken, token, { now: now + 300, docRef, sid })).toEqual({ ok: false, reason: 'exp' })
    expect(await verifyToken(keys.kToken, token, { now, docRef: 'b'.repeat(24), sid })).toEqual({ ok: false, reason: 'docRef' })
    expect(await verifyToken(keys.kToken, token, { now, docRef, sid: 'ffffffffffff' })).toEqual({ ok: false, reason: 'sid' })
    expect(await verifyToken(keys.kToken, 'a.b.c', { now, docRef, sid })).toEqual({ ok: false, reason: 'malformed' })
    expect(await verifyToken(keys.kToken, 42, { now, docRef, sid })).toEqual({ ok: false, reason: 'malformed' })
    expect(await verifyToken(keys.kToken, 'x'.repeat(513), { now, docRef, sid })).toEqual({ ok: false, reason: 'malformed' })
  })
  it('a validly signed token with another scope is refused', async () => {
    // Sign a scope-x token with the real key to reach the scope check.
    const enc = new TextEncoder()
    const b = (o: object) => btoa(JSON.stringify(o)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
    const body = `${b({ alg: 'HS256', typ: 'JWT' })}.${b({ v: 1, docRef, sid, exp: now + 10, iat: now, jti: 'j', scope: 'x' })}`
    const sig = new Uint8Array(await crypto.subtle.sign('HMAC', keys.kToken, enc.encode(body)))
    const s = btoa(String.fromCharCode(...sig)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
    expect(await verifyToken(keys.kToken, `${body}.${s}`, { now, docRef, sid })).toEqual({ ok: false, reason: 'scope' })
  })
})
