/** The §6.2 token contract: HS256 under kToken, claims exactly
 * { v, docRef, sid, exp, iat, jti, scope: 'stream' }. Time is injected. */
import { hex } from './keys'

export interface Claims { v: 1; docRef: string; sid: string; exp: number; iat: number; jti: string; scope: 'stream' }
export type VerifyFailure = 'signature' | 'exp' | 'docRef' | 'sid' | 'scope' | 'malformed'
export type VerifyResult = { ok: true; claims: Claims } | { ok: false; reason: VerifyFailure }

export const MAX_TOKEN_CHARS = 512
const enc = new TextEncoder(), dec = new TextDecoder('utf-8', { fatal: true })
const HEADER = b64url(enc.encode(JSON.stringify({ alg: 'HS256', typ: 'JWT' })))

function b64url(bytes: Uint8Array): string {
  let s = ''
  for (const b of bytes) s += String.fromCharCode(b)
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}
function unb64url(s: string): Uint8Array {
  if (!/^[A-Za-z0-9_-]*$/.test(s)) throw new Error('b64url')
  return Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (s.length % 4)) % 4)), (c) => c.charCodeAt(0))
}

export const randomHex = (bytes: number) => hex(crypto.getRandomValues(new Uint8Array(bytes)))
/** 48-bit session id (§6.2): never the viewer's identity. */
export const newSid = () => randomHex(6)

export async function issueToken(kToken: CryptoKey, o: { docRef: string; sid: string; ttlSec: number; now: number }): Promise<{ token: string; jti: string; exp: number }> {
  const jti = randomHex(12)
  const claims: Claims = { v: 1, docRef: o.docRef, sid: o.sid, exp: o.now + o.ttlSec, iat: o.now, jti, scope: 'stream' }
  const body = `${HEADER}.${b64url(enc.encode(JSON.stringify(claims)))}`
  const sig = new Uint8Array(await crypto.subtle.sign('HMAC', kToken, enc.encode(body)))
  const token = `${body}.${b64url(sig)}`
  if (token.length > MAX_TOKEN_CHARS) throw new Error('token exceeds 512 chars')
  return { token, jti, exp: claims.exp }
}

function isClaims(c: unknown): c is Claims {
  if (!c || typeof c !== 'object') return false
  const o = c as Record<string, unknown>
  return o.v === 1 && typeof o.docRef === 'string' && typeof o.sid === 'string' && typeof o.jti === 'string' &&
    Number.isInteger(o.exp) && Number.isInteger(o.iat) && typeof o.scope === 'string'
}

/** Signature first (nothing in an unsigned token is trusted), then exp,
 * docRef, the sid bound to the socket, and scope. */
export async function verifyToken(kToken: CryptoKey, token: unknown, ctx: { now: number; docRef: string; sid: string }): Promise<VerifyResult> {
  if (typeof token !== 'string' || token.length > MAX_TOKEN_CHARS) return { ok: false, reason: 'malformed' }
  const parts = token.split('.')
  if (parts.length !== 3 || parts[0] !== HEADER) return { ok: false, reason: 'malformed' }
  let sig: Uint8Array, claims: unknown
  try { sig = unb64url(parts[2]!) } catch { return { ok: false, reason: 'malformed' } }
  if (!(await crypto.subtle.verify('HMAC', kToken, sig as BufferSource, enc.encode(`${parts[0]}.${parts[1]}`)))) return { ok: false, reason: 'signature' }
  try { claims = JSON.parse(dec.decode(unb64url(parts[1]!))) } catch { return { ok: false, reason: 'malformed' } }
  if (!isClaims(claims)) return { ok: false, reason: 'malformed' }
  if (ctx.now >= claims.exp) return { ok: false, reason: 'exp' }
  if (claims.docRef !== ctx.docRef) return { ok: false, reason: 'docRef' }
  if (claims.sid !== ctx.sid) return { ok: false, reason: 'sid' }
  if (claims.scope !== 'stream') return { ok: false, reason: 'scope' }
  return { ok: true, claims }
}
