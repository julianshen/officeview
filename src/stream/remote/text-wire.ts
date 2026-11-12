/**
 * Stage 3a text ops (spec §4.1): the client's typed view of the wire.
 *
 * Rects are flat `[x, y, w, h, …]` in page px at scale 1, relative to their
 * unit. Every received answer is checked and rebuilt from its checked fields
 * before a consumer sees it; a malformed one fails its request.
 */

export interface RemotePoint { unit: number; x: number; y: number }
/** A text range: the whole document, or the span between two unit points. */
export type RemoteRange = { all: true } | { anchor: RemotePoint; focus: RemotePoint }

export type TextRequest =
  | { type: 'search'; query: string; caseSensitive?: boolean }
  | { type: 'select'; range: RemoteRange }
  | { type: 'copy'; range: RemoteRange }
  | { type: 'a11y'; unit: number; paragraph?: number }

export interface UnitRects { unit: number; rects: number[] }
export type TextMessage =
  | { type: 'search-result'; matches: UnitRects[] }
  | { type: 'search-done'; total: number; truncated: boolean }
  | { type: 'select-result'; units: UnitRects[]; paragraphs: number; last: boolean }
  | { type: 'copy-result'; seq: number; text: string; last: boolean }
  | { type: 'a11y-result'; seq: number; paragraphs: Array<{ rect: number[]; text?: string }>; last: boolean }
  /** Client-side, never on the wire: the request was re-issued (after a resume);
   * everything it delivered before this is void and its answer starts over. */
  | { type: 'restart' }

/** Every text-service message is at most this many bytes of UTF-8 (the server's chunk size). */
export const TEXT_MESSAGE_BYTES = 48 * 1024
/** Per-request totals a text answer may never exceed (the server's R6 caps, or
 * a fixed bound where the server has none). */
export const TEXT_LIMITS = { copyBytes: 4 * 1024 * 1024, searchHits: 5000, selectRects: 20_000, a11yBytes: 4 * 1024 * 1024, a11yParagraphs: 20_000 } as const
/** The wire types of text-service answers (the 48 KiB check applies to these only). */
export const TEXT_ANSWER_TYPES = new Set(['search-result', 'search-done', 'select-result', 'copy-result', 'a11y-result'])

/** What one request has received so far (reset when it is re-issued). */
export interface TextTally { hits: number; rects: number; paragraphs: number; bytes: number }
export const emptyTally = (): TextTally => ({ hits: 0, rects: 0, paragraphs: 0, bytes: 0 })

/** UTF-8 byte length of `s` (lone surrogates count as U+FFFD, 3 bytes). */
export function utf8Bytes(s: string): number {
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c < 0x80) n += 1
    else if (c < 0x800) n += 2
    else if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length && (s.charCodeAt(i + 1) & 0xfc00) === 0xdc00) { n += 4; i++ }
    else n += 3
  }
  return n
}

/**
 * Adds `m` to the request's tally; returns why the answer is over a limit (or,
 * for `search-done`, why its total disagrees with the hits delivered), else
 * undefined. Called before the consumer sees `m`.
 */
export function admit(t: TextTally, m: TextMessage): string | undefined {
  switch (m.type) {
    case 'search-result':
      t.hits += m.matches.length
      return t.hits > TEXT_LIMITS.searchHits ? `search answer over ${TEXT_LIMITS.searchHits} hits` : undefined
    case 'search-done':
      return m.total !== t.hits ? `search total ${m.total} differs from the ${t.hits} hits delivered` : undefined
    case 'select-result':
      for (const u of m.units) t.rects += u.rects.length / 4
      return t.rects > TEXT_LIMITS.selectRects ? `select answer over ${TEXT_LIMITS.selectRects} rects` : undefined
    case 'copy-result':
      t.bytes += utf8Bytes(m.text)
      return t.bytes > TEXT_LIMITS.copyBytes ? 'copy answer over 4 MiB of UTF-8' : undefined
    case 'a11y-result':
      t.paragraphs += m.paragraphs.length
      for (const p of m.paragraphs) if (p.text !== undefined) t.bytes += utf8Bytes(p.text)
      if (t.paragraphs > TEXT_LIMITS.a11yParagraphs) return `a11y answer over ${TEXT_LIMITS.a11yParagraphs} paragraphs`
      return t.bytes > TEXT_LIMITS.a11yBytes ? 'a11y answer over 4 MiB of UTF-8' : undefined
    default: return undefined
  }
}

/** The refusal codes a text request rejects with; any other is `policy-refused`. */
export const TEXT_REFUSALS = new Set(['rate-limited', 'policy-refused', 'session-rotated', 'text-unavailable', 'copy-too-large', 'copy-quota-exhausted'] as const)
export type TextRefusalCode = typeof TEXT_REFUSALS extends Set<infer C> ? C : never

const isCount = (v: unknown): v is number => Number.isInteger(v) && (v as number) >= 0
/** Flat `[x, y, w, h, …]`: finite numbers, whole quads, w and h never negative (x/y may be). */
const isRects = (v: unknown): v is number[] => {
  if (!Array.isArray(v) || v.length % 4 !== 0) return false
  for (let i = 0; i < v.length; i++) if (!Number.isFinite(v[i]) || (i % 4 >= 2 && v[i] < 0)) return false
  return true
}
const unitRects = (v: unknown, unitCount: number): UnitRects[] | undefined => {
  const ok = Array.isArray(v) && v.every((u) => !!u && typeof u === 'object' && isCount(u.unit) && u.unit < unitCount && isRects(u.rects))
  return ok ? (v as UnitRects[]).map((u) => ({ unit: u.unit, rects: u.rects })) : undefined
}
const isParagraph = (v: unknown): v is { rect: number[]; text?: string } => {
  if (!v || typeof v !== 'object') return false
  const p = v as { rect?: unknown; text?: unknown }
  return isRects(p.rect) && (p.rect as number[]).length === 4 && (p.text === undefined || typeof p.text === 'string')
}

/** The checked answer `m` for a request of type `op`, or undefined when it is
 * malformed, names a unit outside `0..unitCount-1`, has a negative rect
 * dimension, belongs to another op, or (copy/a11y) is not the `seq` expected next. */
export function textMessage(op: TextRequest['type'], m: Record<string, unknown>, seq: number, unitCount: number): TextMessage | undefined {
  switch (m.type) {
    case 'search-result': {
      const matches = op === 'search' ? unitRects(m.matches, unitCount) : undefined
      return matches && { type: 'search-result', matches }
    }
    case 'search-done':
      return op === 'search' && isCount(m.total) && typeof m.truncated === 'boolean' ? { type: 'search-done', total: m.total, truncated: m.truncated } : undefined
    case 'select-result': {
      const units = op === 'select' ? unitRects(m.units, unitCount) : undefined
      return units && isCount(m.paragraphs) && (m.last === undefined || m.last === true) ? { type: 'select-result', units, paragraphs: m.paragraphs, last: m.last === true } : undefined
    }
    case 'copy-result':
      return op === 'copy' && m.seq === seq && typeof m.text === 'string' && typeof m.last === 'boolean' ? { type: 'copy-result', seq, text: m.text, last: m.last } : undefined
    case 'a11y-result': {
      const ok = op === 'a11y' && m.seq === seq && typeof m.last === 'boolean' && Array.isArray(m.paragraphs) && m.paragraphs.every(isParagraph)
      return ok ? { type: 'a11y-result', seq, paragraphs: (m.paragraphs as Array<{ rect: number[]; text?: string }>).map((p) => ({ rect: p.rect, ...(p.text !== undefined ? { text: p.text } : {}) })), last: m.last as boolean } : undefined
    }
    default: return undefined
  }
}

/** The answer's final message: the request settles on it. */
export const isTerminal = (m: TextMessage): boolean => m.type === 'search-done' || ('last' in m && m.last)

/** The wire body of a request, built field by field (nothing else rides along). */
export function requestBody(r: TextRequest): Record<string, unknown> {
  switch (r.type) {
    case 'search': return { query: r.query, ...(r.caseSensitive !== undefined ? { caseSensitive: r.caseSensitive } : {}) }
    case 'a11y': return { unit: r.unit, ...(r.paragraph !== undefined ? { paragraph: r.paragraph } : {}) }
    default: return { range: 'all' in r.range ? { all: true } : { anchor: point(r.range.anchor), focus: point(r.range.focus) } }
  }
}
const point = (p: RemotePoint): RemotePoint => ({ unit: p.unit, x: p.x, y: p.y })
