/**
 * Piece 5 (spec §4.3; plan D1, D13, D14; Task 9): the media text messages and their limits, shared by the server and
 * the (lazy) client media path. Nothing eager imports this module.
 *
 * Every parser takes a decoded JSON value and throws a `MediaWireError` on any key outside its set, a wrong type, or a
 * value out of range; it never coerces. Client → server: `media-show`, `media-cursor`, `media-request`, `media-end`.
 * Server → client: `media-show-ok`, `media-start` (then `chunks` MEDIA_CHUNK messages, kind 9, stored, under the same
 * reqId), `media-refused`.
 */
import type { MediaRefusal } from './media-types'
import { MEDIA_ID, SHOW_PLAN_MEDIA_LIMITS } from './show-plan'

export const MEDIA_WIRE_LIMITS: Readonly<{ chunkBytes: 4194304; sendChunk: 1048512; audioBytes: 33554432; textBytes: 65536 }> = Object.freeze({
  /** MEDIA_CHUNK: ≤ 4 MiB encoded and decoded per message (the admission bound). */
  chunkBytes: 4194304,
  /** D13: the server sends 1 MiB − 64 B per chunk, one fragment each. */
  sendChunk: 1048512,
  /** A whole audio file, declared in `media-start` and enforced while receiving. */
  audioBytes: 33554432,
  /** Every media text message. */
  textBytes: 65536,
} as const)

export type MediaStart = { reqId: number; type: 'media-start'; mediaId: string; contentType: 'audio/mp4'; durationMs: number; bytes: number; chunks: number }
export type MediaRefused = { reqId: number; type: 'media-refused'; mediaId?: string; reason: MediaRefusal | 'rate-limited' | 'policy-refused' | 'invalid-media' }
export type MediaShowOk = { reqId: number; type: 'media-show-ok'; queued: number }

export class MediaWireError extends Error { constructor(what: string) { super(`media wire: ${what}`); this.name = 'MediaWireError' } }

export const MEDIA_REFUSED_REASONS: ReadonlySet<MediaRefused['reason']> = new Set<MediaRefused['reason']>([
  'media-too-long', 'media-too-large', 'media-timeout', 'media-busy', 'media-unreadable', 'media-linked', 'media-unavailable',
  'rate-limited', 'policy-refused', 'invalid-media',
])

const fail = (what: string): never => { throw new MediaWireError(what) }
/** The value as a record whose keys are all in `required` ∪ `optional`, with every `required` key present. */
function record(m: unknown, type: string, required: readonly string[], optional: readonly string[] = []): Record<string, unknown> {
  if (!m || typeof m !== 'object' || Array.isArray(m)) return fail('not an object')
  const o = m as Record<string, unknown>
  if (o.type !== type) fail(`type is not ${type}`)
  for (const k of Object.keys(o)) if (k !== 'type' && !required.includes(k) && !optional.includes(k)) fail(`unknown key ${k}`)
  for (const k of required) if (!(k in o)) fail(`missing ${k}`)
  return o
}
const u32 = (v: unknown, what: string): number => (Number.isInteger(v) && (v as number) >= 0 && (v as number) <= 0xffffffff ? v as number : fail(what))
const reqIdOf = (v: unknown): number => { const n = u32(v, 'reqId'); return n === 0 ? fail('reqId') : n }
const mediaIdOf = (v: unknown): string => (typeof v === 'string' && MEDIA_ID.test(v) ? v : fail('mediaId'))
const tokenOf = (v: unknown): string => (typeof v === 'string' && v.length > 0 && v.length <= 4096 ? v : fail('token'))

/** Client: throws on any key outside the set, a contentType other than audio/mp4, bytes over `audioBytes`, or
 * chunks ≠ ceil(bytes / sendChunk). */
export function parseMediaStart(m: unknown): MediaStart {
  const o = record(m, 'media-start', ['reqId', 'mediaId', 'contentType', 'durationMs', 'bytes', 'chunks'])
  if (o.contentType !== 'audio/mp4') fail('contentType')
  const durationMs = u32(o.durationMs, 'durationMs')
  if (durationMs > SHOW_PLAN_MEDIA_LIMITS.durationMs) fail('durationMs')
  const bytes = u32(o.bytes, 'bytes')
  if (bytes < 1 || bytes > MEDIA_WIRE_LIMITS.audioBytes) fail('bytes')
  const chunks = u32(o.chunks, 'chunks')
  if (chunks !== Math.ceil(bytes / MEDIA_WIRE_LIMITS.sendChunk)) fail('chunks')
  return { reqId: reqIdOf(o.reqId), type: 'media-start', mediaId: mediaIdOf(o.mediaId), contentType: 'audio/mp4', durationMs, bytes, chunks }
}

/** Client. */
export function parseMediaShowOk(m: unknown): MediaShowOk {
  const o = record(m, 'media-show-ok', ['reqId', 'queued'])
  return { reqId: reqIdOf(o.reqId), type: 'media-show-ok', queued: u32(o.queued, 'queued') }
}

/** Client. */
export function parseMediaRefused(m: unknown): MediaRefused {
  const o = record(m, 'media-refused', ['reqId', 'reason'], ['mediaId'])
  if (!MEDIA_REFUSED_REASONS.has(o.reason as MediaRefused['reason'])) fail('reason')
  return { reqId: reqIdOf(o.reqId), type: 'media-refused', ...(o.mediaId !== undefined ? { mediaId: mediaIdOf(o.mediaId) } : {}), reason: o.reason as MediaRefused['reason'] }
}

/** Server: `media-show {reqId, token, unit}` (D1). `unit` is range-checked by the server against its document. */
export function parseMediaShow(m: unknown): { reqId: number; token: string; unit: number } {
  const o = record(m, 'media-show', ['reqId', 'token', 'unit'])
  return { reqId: reqIdOf(o.reqId), token: tokenOf(o.token), unit: u32(o.unit, 'unit') }
}

/** Server: `media-cursor {unit}` (D1; no reply, no token). */
export function parseMediaCursor(m: unknown): { unit: number } {
  const o = record(m, 'media-cursor', ['unit'])
  return { unit: u32(o.unit, 'unit') }
}

/** Server: `media-request {reqId, token, mediaId}`. `mediaId` is any short string: the server answers anything that
 * is not one of the session's own ids (another session's, a cache key, a placeholder) `invalid-media`. */
export function parseMediaRequest(m: unknown): { reqId: number; token: string; mediaId: string } {
  const o = record(m, 'media-request', ['reqId', 'token', 'mediaId'])
  const mediaId = typeof o.mediaId === 'string' && o.mediaId.length >= 1 && o.mediaId.length <= 256 ? o.mediaId : fail('mediaId')
  return { reqId: reqIdOf(o.reqId), token: tokenOf(o.token), mediaId }
}

/** Server: `media-end {}` (D1; no reply). */
export function parseMediaEnd(m: unknown): Record<string, never> {
  record(m, 'media-end', [])
  return {}
}
