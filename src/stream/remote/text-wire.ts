/**
 * Stage 3a text ops (spec §4.1): the client's typed view of the wire.
 *
 * Rects are flat `[x, y, w, h, …]` in page px at scale 1, relative to their
 * unit. Every received answer is checked and rebuilt from its checked fields
 * before a consumer sees it; a malformed one fails its request.
 */
import type { OperationRefusal } from '../../xlsx/view/contracts'
import {
  A11Y_MAX_CELLS, decodeA11yCells, decodeDetection, decodeFindHits, decodeFitResult, decodeSheetView, decodeValuesPage, encodeWireSelection, FIND_MAX_HITS, SHEET_MAX_COLS, SHEET_MAX_ROWS,
  type SheetFindHit, type SheetFitRequest, type SheetOrder, type SheetResizeRequest, type SheetViewChange, type ValuesPage, type WireDetection, type WireFilterSpec,
  type WireRange, type WireRule, type WireSelection, type WireSortSpec, type WireTyped,
} from '../sheet-wire'

export interface RemotePoint { unit: number; x: number; y: number }
/** A text range: the whole document, or the span between two unit points. */
export type RemoteRange = { all: true } | { anchor: RemotePoint; focus: RemotePoint }

export type TextRequest =
  | { type: 'search'; query: string; caseSensitive?: boolean }
  | { type: 'select'; range: RemoteRange }
  | { type: 'copy'; range: RemoteRange }
  | { type: 'a11y'; unit: number; paragraph?: number }
  // Stage 4b (`sheet-view@1`, S4 §6.7; P6, P7, P9–P11): the sheet ops of a tiles-mode session.
  | { type: 'sheet-find'; query: string; matchCase?: boolean }
  | { type: 'sheet-copy'; sheetId: string; viewGen: number; selection: WireSelection }
  | { type: 'sheet-a11y'; sheetId: string; viewGen: number; rows: Array<[number, number]>; cols: [number, number] }
  | ({ type: 'sheet-fit' } & SheetFitRequest)
  | ({ type: 'sheet-resize' } & SheetResizeRequest)
  // Stage 4c (S4 §6.6; P4–P8): the sort ops of a `sheetSort` session.
  | ({ type: 'sheet-sort'; sheetId: string; viewGen: number } & WireSortSpec)
  | ({ type: 'sheet-filter'; sheetId: string; viewGen: number } & WireFilterSpec)
  | ({ type: 'filter-values' } & ValuesRequest)
  | { type: 'range-detect'; sheetId: string; viewGen: number; kind: 'sort' | 'filter'; selection: WireSelection; override?: { header: 0 | 1 } }
  | { type: 'sheet-order'; sheetId: string }

/** A value list's first page, or a later page by its cursor (P6). */
export type ValuesRequest =
  | { sheetId: string; viewGen: number; range: WireRange; col: number; search?: { text: string; wildcards: boolean } }
  | { sheetId: string; viewGen: number; cursor: string }

export interface UnitRects { unit: number; rects: number[] }
export type TextMessage =
  | { type: 'search-result'; matches: UnitRects[] }
  | { type: 'search-done'; total: number; truncated: boolean }
  | { type: 'select-result'; units: UnitRects[]; paragraphs: number; last: boolean }
  | { type: 'copy-result'; seq: number; text: string; last: boolean }
  | { type: 'a11y-result'; seq: number; paragraphs: Array<{ rect: number[]; text?: string }>; last: boolean }
  /** Stage 4b answers (`sheet-copy` is answered by `copy-result`, the Stage 3 shape). */
  | { type: 'sheet-find-result'; hits: SheetFindHit[] }
  | { type: 'sheet-find-done'; total: number; truncated: boolean }
  | { type: 'sheet-a11y-result'; seq: number; cells: Array<{ row: number; col: number; text: string }>; last: boolean }
  | { type: 'sheet-fit-result'; size: number }
  | { type: 'sheet-view'; view: SheetViewChange }
  /** Stage 4c answers: a whole `sheet-order` answer (the session assembles its parts, P7), a value-list page, a detection. */
  | { type: 'sheet-order'; order: SheetOrder }
  | { type: 'filter-values-result'; page: ValuesPage }
  | { type: 'range-detect-result'; detection: WireDetection }
  /** Client-side, never on the wire: the request was re-issued (after a resume);
   * everything it delivered before this is void and its answer starts over. */
  | { type: 'restart' }

/** Every text-service message is at most this many bytes of UTF-8 (the server's chunk size). */
export const TEXT_MESSAGE_BYTES = 48 * 1024
/** Per-request totals a text answer may never exceed (the server's R6 caps, or
 * a fixed bound where the server has none). */
export const TEXT_LIMITS = { copyBytes: 4 * 1024 * 1024, searchHits: 5000, selectRects: 20_000, a11yBytes: 4 * 1024 * 1024, a11yParagraphs: 20_000 } as const
/** Backstop: no one text request is answered in more than this many messages. */
export const TEXT_MAX_MESSAGES = 4096
/** The wire types of text-service answers (the 48 KiB check applies to these only). */
export const TEXT_ANSWER_TYPES = new Set(['search-result', 'search-done', 'select-result', 'copy-result', 'a11y-result', 'sheet-find-result', 'sheet-find-done', 'sheet-a11y-result', 'sheet-fit-result', 'sheet-view', 'filter-values-result', 'range-detect-result'])
/** The Stage 4b request types: sent only on a `sheet-view@1` session (`manifest.sheetView === 1`). */
export const SHEET_TEXT_TYPES = new Set<TextRequest['type']>(['sheet-find', 'sheet-copy', 'sheet-a11y', 'sheet-fit', 'sheet-resize'])
/** The Stage 4c request types: sent only on a `sheetSort` session (`manifest.sheetSort === 1`, P1). */
export const SHEET_SORT_TYPES = new Set<TextRequest['type']>(['sheet-sort', 'sheet-filter', 'filter-values', 'range-detect', 'sheet-order'])
/** The Stage 4c requests answered by `sheet-order` parts (P7), each at most `LAYOUT_PART_BYTES` (not the 48 KiB of a text answer). */
export const SHEET_ORDER_TYPES = new Set<TextRequest['type']>(['sheet-sort', 'sheet-filter', 'sheet-order'])

/** What one request has received so far (reset when it is re-issued). */
export interface TextTally {
  messages: number; hits: number; rects: number; paragraphs: number; bytes: number
  /** Stage 4b a11y: the cells delivered, and their `row * SHEET_MAX_COLS + col` keys (a repeat fails the answer). */
  cells: number; seen?: Set<number>
}
export const emptyTally = (): TextTally => ({ messages: 0, hits: 0, rects: 0, paragraphs: 0, bytes: 0, cells: 0 })

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
  if (++t.messages > TEXT_MAX_MESSAGES) return `answer over ${TEXT_MAX_MESSAGES} messages`
  // A non-final chunk always carries payload (the server never sends an empty
  // one); only a single final answer may be empty.
  const empty = m.type === 'search-result' ? !m.matches.length
    : m.type === 'select-result' ? !m.last && m.units.every((u) => !u.rects.length)
    : m.type === 'copy-result' ? !m.last && m.text === ''
    : m.type === 'a11y-result' ? !m.last && !m.paragraphs.length
    : m.type === 'sheet-find-result' ? !m.hits.length
    : m.type === 'sheet-a11y-result' ? !m.last && !m.cells.length
    : false
  if (empty) return `empty non-final ${m.type} chunk`
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
    case 'sheet-find-result':
      t.hits += m.hits.length
      return t.hits > FIND_MAX_HITS ? `sheet-find answer over ${FIND_MAX_HITS} hits` : undefined
    case 'sheet-find-done':
      return m.total !== t.hits ? `sheet-find total ${m.total} differs from the ${t.hits} hits delivered` : undefined
    case 'sheet-a11y-result':
      t.cells += m.cells.length
      t.seen ??= new Set()
      for (const c of m.cells) {
        const k = c.row * SHEET_MAX_COLS + c.col
        if (t.seen.has(k)) return `sheet-a11y answer repeats cell ${c.row},${c.col}`
        t.seen.add(k)
        t.bytes += utf8Bytes(c.text)
      }
      if (t.cells > A11Y_MAX_CELLS) return `sheet-a11y answer over ${A11Y_MAX_CELLS} cells`
      return t.bytes > TEXT_LIMITS.a11yBytes ? 'sheet-a11y answer over 4 MiB of UTF-8' : undefined
    default: return undefined
  }
}

/** The refusal codes a text request rejects with; any other is `policy-refused`. */
export const TEXT_REFUSALS = new Set(['rate-limited', 'policy-refused', 'session-rotated', 'text-unavailable', 'copy-too-large', 'copy-quota-exhausted',
  // Stage 4b sheet ops (P12; `unsupported-op` from a server without `sheet-view@1`).
  'stale-view', 'hidden-target', 'coverage-unknown', 'view-state-limit', 'unsupported-op'] as const)
export type TextRefusalCode = typeof TEXT_REFUSALS extends Set<infer C> ? C : never
/** The refusal codes a Stage 4c sort op rejects with (S4 §6.6; P4, P8, P13); any other is `policy-refused`. */
export const SORT_REFUSALS = new Set<SortRefusalCode>([...TEXT_REFUSALS, 'stale-filter-values', 'filter-condition-refused', 'view-refused', 'cancelled'])
export type SortRefusalCode = TextRefusalCode | 'stale-filter-values' | 'filter-condition-refused' | 'view-refused' | 'cancelled'

/**
 * The local strings a `view-refused` refusal is rebuilt with (P13: the server drops every `message`). `pivot` is the
 * local engine's own text (`PIVOT_SHEET_CHANGE_MESSAGE`, src/xlsx/view/operationGuard.ts, pinned by a test: no runtime
 * import from `src/xlsx/**` here); the others are fixed texts for reasons whose local message names the local cause.
 * `too-large` (Task 14a) is the server's own reason: the sheet's sort data could not be prepared inside its deadline. It
 * is rebuilt as the local `resource-limit` refusal with this text.
 */
export const VIEW_REFUSAL_MESSAGES = Object.freeze({
  pivot: "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.",
  'ambiguous-selection': 'Select a cell inside the data to sort or filter.',
  'invalid-range': 'Invalid operation range.',
  'invalid-rule': 'Invalid sort or filter rule.',
  'resource-limit': 'View exceeds the session memory budget.',
  unsupported: 'This operation is not supported.',
  'source-error': 'The source could not be read; retry.',
  'too-large': 'This sheet is too large to sort or filter in this viewing mode.',
})
const RECT_REASONS = new Set(['merge', 'fixed-array', 'dynamic-spill', 'outside-column-data', 'pivot'])
const MESSAGE_REASONS = new Set(['ambiguous-selection', 'invalid-range', 'invalid-rule', 'resource-limit', 'unsupported'])
const VIEW_REFUSAL_KEYS = new Set(['reqId', 'type', 'code', 'reason', 'range', 'count', 'at'])
const inGrid = (v: unknown, max: number): v is number => Number.isInteger(v) && (v as number) >= 0 && (v as number) < max

/**
 * The local refusal of a `refused {code: 'view-refused', reason, range?, count?, at?}` (P13), or undefined when it is
 * off that contract: an unknown reason (`incomplete` included: it cannot occur on a resident key view), a missing or
 * extra field, a rect that is not `{r0, r1, c0, c1}` inside the grid with `r0 ≤ r1`, `c0 ≤ c1`. `drawingId` is dropped
 * on the wire, so a drawing refusal names none.
 */
export function viewRefusal(m: Record<string, unknown>): OperationRefusal | undefined {
  if (!Object.keys(m).every((k) => VIEW_REFUSAL_KEYS.has(k))) return undefined
  const reason = m.reason, extra = (keys: string[]) => ['range', 'count', 'at'].some((k) => !keys.includes(k) && k in m)
  if (typeof reason !== 'string') return undefined
  if (RECT_REASONS.has(reason)) {
    const r = m.range as Record<string, unknown> | undefined
    if (extra(['range']) || !r || typeof r !== 'object' || Object.keys(r).length !== 4) return undefined
    if (!inGrid(r.r0, SHEET_MAX_ROWS) || !inGrid(r.r1, SHEET_MAX_ROWS) || !inGrid(r.c0, SHEET_MAX_COLS) || !inGrid(r.c1, SHEET_MAX_COLS) || r.r1 < r.r0 || r.c1 < r.c0) return undefined
    const range = { r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 }
    return reason === 'pivot' ? { reason, range, message: VIEW_REFUSAL_MESSAGES.pivot } : { reason: reason as 'merge' | 'fixed-array' | 'dynamic-spill' | 'outside-column-data', range }
  }
  if (reason === 'formula-value-unavailable') {
    const at = m.at
    if (extra(['count', 'at']) || !Number.isInteger(m.count) || (m.count as number) < 1 || !Array.isArray(at) || at.length !== 2 || !inGrid(at[0], SHEET_MAX_ROWS) || !inGrid(at[1], SHEET_MAX_COLS)) return undefined
    return { reason, count: m.count as number, firstAddress: { row: at[0], col: at[1] } }
  }
  if (extra([])) return undefined
  if (reason === 'drawing-impact') return { reason, drawingId: '' }
  if (reason === 'source-error') return { reason, code: 'page-error', retryable: false, message: VIEW_REFUSAL_MESSAGES['source-error'] }
  if (reason === 'too-large') return { reason: 'resource-limit', message: VIEW_REFUSAL_MESSAGES['too-large'] }
  if (MESSAGE_REASONS.has(reason)) {
    const r = reason as keyof typeof VIEW_REFUSAL_MESSAGES & ('ambiguous-selection' | 'invalid-range' | 'invalid-rule' | 'resource-limit' | 'unsupported')
    return { reason: r, message: VIEW_REFUSAL_MESSAGES[r] }
  }
  return undefined
}

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

/** What `textMessage` checks an answer against: the manifest's unit and (tiles mode) sheet counts. */
export interface TextBounds { unitCount: number; sheetCount: number }

/** The checked answer `m` for `request`, or undefined when it is malformed,
 * names a unit outside `0..unitCount-1`, has a negative rect dimension,
 * belongs to another op, or (copy/a11y) is not the `seq` expected next.
 * Stage 4b answers are also held to their request (Review: hostile answers):
 * a hit names a listed sheet, an a11y cell lies once inside the window asked,
 * a `sheet-view` names the sheet, the change asked and no older viewGen. */
export function textMessage(request: TextRequest, m: Record<string, unknown>, seq: number, bounds: TextBounds): TextMessage | undefined {
  const op = request.type
  if (op.startsWith('sheet-') || SHEET_SORT_TYPES.has(op)) {
    try { return sheetMessage(request, m, seq, bounds.sheetCount) } catch { return undefined }
  }
  const unitCount = bounds.unitCount
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

/** A Stage 4b answer, checked by the `sheet-wire` decoders (which throw `InvalidOperand`) and against `r`. */
function sheetMessage(r: TextRequest, m: Record<string, unknown>, seq: number, sheetCount: number): TextMessage | undefined {
  switch (m.type) {
    case 'sheet-find-result': {
      if (r.type !== 'sheet-find') return undefined
      const hits = decodeFindHits(m)
      return hits.every((h) => h.sheetIndex < sheetCount) ? { type: 'sheet-find-result', hits } : undefined
    }
    case 'sheet-find-done':
      return r.type === 'sheet-find' && isCount(m.total) && typeof m.truncated === 'boolean' ? { type: 'sheet-find-done', total: m.total, truncated: m.truncated } : undefined
    case 'copy-result':
      return r.type === 'sheet-copy' && m.seq === seq && typeof m.text === 'string' && typeof m.last === 'boolean' ? { type: 'copy-result', seq, text: m.text, last: m.last } : undefined
    case 'sheet-a11y-result': {
      if (r.type !== 'sheet-a11y' || m.seq !== seq || typeof m.last !== 'boolean') return undefined
      const cells = decodeA11yCells(m)
      const inRows = (row: number) => r.rows.some(([a, b]) => a <= row && row <= b)
      if (!cells.every((c) => inRows(c.row) && r.cols[0] <= c.col && c.col <= r.cols[1])) return undefined
      return { type: 'sheet-a11y-result', seq, cells, last: m.last }
    }
    case 'sheet-fit-result':
      return r.type === 'sheet-fit' ? { type: 'sheet-fit-result', size: decodeFitResult(m) } : undefined
    case 'sheet-view': {
      if (r.type !== 'sheet-resize') return undefined
      const view = decodeSheetView(m)
      if (view.sheetId !== r.sheetId || view.viewGen < r.viewGen) return undefined
      const c = view.change
      const same = 'reset' in r ? 'reset' in c : !('reset' in c) && c.axis === r.axis && c.index === r.index
      return same ? { type: 'sheet-view', view } : undefined
    }
    case 'filter-values-result':
      return r.type === 'filter-values' ? { type: 'filter-values-result', page: decodeValuesPage(m) } : undefined
    case 'range-detect-result':
      return r.type === 'range-detect' ? { type: 'range-detect-result', detection: decodeDetection(m) } : undefined
    default: return undefined
  }
}

/** The answer's final message: the request settles on it. */
export const isTerminal = (m: TextMessage): boolean =>
  m.type === 'search-done' || m.type === 'sheet-find-done' || m.type === 'sheet-fit-result' || m.type === 'sheet-view'
  || m.type === 'sheet-order' || m.type === 'filter-values-result' || m.type === 'range-detect-result' || ('last' in m && m.last)

/** The wire body of a request, built field by field (nothing else rides along). */
export function requestBody(r: TextRequest): Record<string, unknown> {
  switch (r.type) {
    case 'search': return { query: r.query, ...(r.caseSensitive !== undefined ? { caseSensitive: r.caseSensitive } : {}) }
    case 'a11y': return { unit: r.unit, ...(r.paragraph !== undefined ? { paragraph: r.paragraph } : {}) }
    case 'sheet-find': return { query: r.query, ...(r.matchCase !== undefined ? { matchCase: r.matchCase } : {}) }
    case 'sheet-copy': return { sheetId: r.sheetId, viewGen: r.viewGen, selection: encodeWireSelection(r.selection) }
    case 'sheet-a11y': return { sheetId: r.sheetId, viewGen: r.viewGen, rows: r.rows.map(([a, b]) => [a, b]), cols: [r.cols[0], r.cols[1]] }
    case 'sheet-fit': return { sheetId: r.sheetId, viewGen: r.viewGen, axis: r.axis, index: r.index, pane: box(r.pane), ...(r.regions?.length ? { regions: r.regions.map(box) } : {}) }
    case 'sheet-resize': return 'reset' in r ? { sheetId: r.sheetId, viewGen: r.viewGen, reset: true } : { sheetId: r.sheetId, viewGen: r.viewGen, axis: r.axis, index: r.index, size: r.size }
    case 'sheet-sort': return { sheetId: r.sheetId, viewGen: r.viewGen, range: rangeBody(r.range), keys: r.keys.map((k) => ({ col: k.col, order: k.order })), caseSensitive: r.caseSensitive }
    case 'sheet-filter': return { sheetId: r.sheetId, viewGen: r.viewGen, range: rangeBody(r.range), columns: r.columns.map(([col, rule]) => [col, ruleBody(rule)]) }
    case 'filter-values':
      if ('cursor' in r) return { sheetId: r.sheetId, viewGen: r.viewGen, cursor: r.cursor }
      return { sheetId: r.sheetId, viewGen: r.viewGen, range: rangeBody(r.range), col: r.col, ...(r.search ? { search: { text: r.search.text, wildcards: r.search.wildcards } } : {}) }
    case 'range-detect': return { sheetId: r.sheetId, viewGen: r.viewGen, kind: r.kind, selection: encodeWireSelection(r.selection), ...(r.override ? { override: { header: r.override.header } } : {}) }
    case 'sheet-order': return { sheetId: r.sheetId }
    default: return { range: 'all' in r.range ? { all: true } : { anchor: point(r.range.anchor), focus: point(r.range.focus) } }
  }
}
const point = (p: RemotePoint): RemotePoint => ({ unit: p.unit, x: p.x, y: p.y })
const box = (b: { x: number; y: number; width: number; height: number }) => ({ x: b.x, y: b.y, width: b.width, height: b.height })
const rangeBody = (r: WireRange): WireRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, header: r.header, totals: r.totals })
const day = (d: { year: number; month: number; day: number }) => ({ year: d.year, month: d.month, day: d.day })
function typedBody(t: WireTyped): WireTyped {
  switch (t.kind) {
    case 'text': return { kind: 'text', op: t.op, first: t.first, wildcards: t.wildcards }
    case 'number': return { kind: 'number', op: t.op, first: t.first, ...(t.second !== undefined ? { second: t.second } : {}) }
    case 'date': return { kind: 'date', op: t.op, first: day(t.first), ...(t.second !== undefined ? { second: day(t.second) } : {}) }
    case 'datePreset': return { kind: 'datePreset', preset: t.preset }
  }
}
function ruleBody(rule: WireRule): WireRule {
  switch (rule.kind) {
    case 'committed': return { kind: 'committed' }
    case 'values': return { kind: 'values', list: rule.list, mode: rule.mode, ids: rule.ids.map(([a, b]): [number, number] => [a, b]), blanks: rule.blanks }
    case 'compound': return { kind: 'compound', connective: rule.connective, first: typedBody(rule.first), second: typedBody(rule.second) }
    default: return typedBody(rule)
  }
}
