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

/** The refusal codes a text request rejects with; any other is `policy-refused`. */
export const TEXT_REFUSALS = new Set(['rate-limited', 'policy-refused', 'session-rotated', 'text-unavailable', 'copy-too-large', 'copy-quota-exhausted'] as const)
export type TextRefusalCode = typeof TEXT_REFUSALS extends Set<infer C> ? C : never

const isCount = (v: unknown): v is number => Number.isInteger(v) && (v as number) >= 0
const isRects = (v: unknown): v is number[] => Array.isArray(v) && v.length % 4 === 0 && v.every((n) => Number.isFinite(n))
const isUnitRects = (v: unknown): v is UnitRects => !!v && typeof v === 'object' && isCount((v as UnitRects).unit) && isRects((v as UnitRects).rects)
const unitRects = (v: unknown): UnitRects[] | undefined => Array.isArray(v) && v.every(isUnitRects) ? v.map((u) => ({ unit: u.unit, rects: u.rects })) : undefined
const isParagraph = (v: unknown): v is { rect: number[]; text?: string } => {
  if (!v || typeof v !== 'object') return false
  const p = v as { rect?: unknown; text?: unknown }
  return isRects(p.rect) && (p.rect as number[]).length === 4 && (p.text === undefined || typeof p.text === 'string')
}

/** The checked answer `m` for a request of type `op`, or undefined when it is
 * malformed, belongs to another op, or (copy/a11y) is not the `seq` expected next. */
export function textMessage(op: TextRequest['type'], m: Record<string, unknown>, seq: number): TextMessage | undefined {
  switch (m.type) {
    case 'search-result': {
      const matches = op === 'search' ? unitRects(m.matches) : undefined
      return matches && { type: 'search-result', matches }
    }
    case 'search-done':
      return op === 'search' && isCount(m.total) && typeof m.truncated === 'boolean' ? { type: 'search-done', total: m.total, truncated: m.truncated } : undefined
    case 'select-result': {
      const units = op === 'select' ? unitRects(m.units) : undefined
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
