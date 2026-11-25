/**
 * Stage 4a: the `sheet-tiles@1` wire constants (S4 §6, R12) and the `sheet-layout` codec (S4 §6.4, R14), shared by
 * the stream server and client. React-free, and it imports nothing from `src/xlsx/**` at runtime (types only, F5), so
 * the stream runtime's module set and the eager gates do not move: `TILE_ZOOM_STEPS`, `SHEET_MAX_ROWS` and
 * `SHEET_MAX_COLS` are literals, and tests pin them to the engine's constants.
 */
import type { AxisSizeInterval } from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import { InvalidOperand } from './ops'
import { decodeRasterBody, encodeRasterBody } from './raster'
import { decodeUnit, encodeUnit } from './unit'

/** The HELLO capability that selects tiles mode for an XLSX session. */
export const SHEET_TILES_FEATURE = 'sheet-tiles@1'
/** A tile is 512 × 512 device px. */
export const TILE_PX = 512
/** At most 64 tiles per batch. */
export const TILE_BATCH_MAX = 64
export const TILE_DPRS = [1, 2] as const
export const TILE_PALETTES = ['light', 'dark'] as const
/** Equal to `SHEET_ZOOM_STEPS` (src/xlsx/view/sheetZoom.ts); a tile's zoom travels as an index into it. */
export const TILE_ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4] as const
export type TilePalette = typeof TILE_PALETTES[number]
/** HELLO `features`: at most 16 entries, each at most 32 characters (R12). */
export const HELLO_FEATURES_MAX = 16, HELLO_FEATURE_CHARS = 32
/** A tiles manifest lists at most 256 sheets, each name at most 31 characters (Excel's own bound). A workbook past
 * either bound is served in units mode. */
export const MAX_SHEETS = 256, SHEET_NAME_CHARS = 31

// ---- XLSX_TILE body (S4 §6.5, R13) ----

/** One `XLSX_TILE` (kind 5) body: an `encodeUnit` body with `unitId = sheetIndex` and discriminator 3, whose body is
 * `u32 viewGen, u8 zoomStep, u8 dpr, u16 0, u32 col, u32 row`, then the `UNIT_RASTER` raster body
 * (`u32 w, u32 h, u32 pngLen, PNG`) with `w = h = 512`. */
export interface TileBody { sheetIndex: number; viewGen: number; zoomStep: number; dpr: 1 | 2; col: number; row: number; png: Uint8Array; w: number; h: number }
const TILE_HEAD_BYTES = 16
const u32 = (n: unknown): n is number => Number.isInteger(n) && (n as number) >= 0 && (n as number) <= 0xffffffff
function tileFieldsProblem(t: { sheetIndex: unknown; viewGen: unknown; zoomStep: unknown; dpr: unknown; col: unknown; row: unknown; w: unknown; h: unknown }): string | undefined {
  if (!u32(t.sheetIndex) || !u32(t.viewGen) || !u32(t.col) || !u32(t.row)) return 'header field'
  if (!Number.isInteger(t.zoomStep) || (t.zoomStep as number) < 0 || (t.zoomStep as number) >= TILE_ZOOM_STEPS.length) return 'zoomStep'
  if (t.dpr !== 1 && t.dpr !== 2) return 'dpr'
  if (t.w !== TILE_PX || t.h !== TILE_PX) return 'tile size'
  return undefined
}

/** Throws `InvalidOperand` for a field outside the wire contract. The PNG is not inspected (the server encodes it). */
export function encodeTileBody(t: TileBody): Uint8Array {
  const problem = tileFieldsProblem(t)
  if (problem) throw new InvalidOperand(`XLSX_TILE: ${problem}`)
  const raster = encodeRasterBody(t.w, t.h, t.png)
  const body = new Uint8Array(TILE_HEAD_BYTES + raster.length), v = new DataView(body.buffer)
  v.setUint32(0, t.viewGen, true); v.setUint8(4, t.zoomStep); v.setUint8(5, t.dpr); v.setUint16(6, 0, true); v.setUint32(8, t.col, true); v.setUint32(12, t.row, true)
  body.set(raster, TILE_HEAD_BYTES)
  return encodeUnit({ unitId: t.sheetIndex, discriminator: 3, body })
}

/** The inverse of `encodeTileBody`. Throws `InvalidOperand` for anything else: a short or malformed unit header, a
 * discriminator other than 3, a reserved field not 0, a `zoomStep` past `TILE_ZOOM_STEPS`, a `dpr` outside {1, 2},
 * a raster body that is not a PNG whose IHDR is 512 × 512. Reads the IHDR only (nothing is decoded). */
export function decodeTileBody(bytes: Uint8Array): TileBody {
  try {
    const unit = decodeUnit(bytes)
    if (unit.discriminator !== 3) throw new InvalidOperand('XLSX_TILE: discriminator')
    if (unit.body.length < TILE_HEAD_BYTES) throw new InvalidOperand('XLSX_TILE: truncated')
    const v = new DataView(unit.body.buffer, unit.body.byteOffset, unit.body.byteLength)
    if (v.getUint16(6, true) !== 0) throw new InvalidOperand('XLSX_TILE: reserved')
    const head = { sheetIndex: unit.unitId, viewGen: v.getUint32(0, true), zoomStep: v.getUint8(4), dpr: v.getUint8(5), col: v.getUint32(8, true), row: v.getUint32(12, true) }
    const r = decodeRasterBody(unit.body.subarray(TILE_HEAD_BYTES), { maxRgbaBytes: TILE_PX * TILE_PX * 4 })
    const t = { ...head, png: r.png, w: r.w, h: r.h }
    const problem = tileFieldsProblem(t)
    if (problem) throw new InvalidOperand(`XLSX_TILE: ${problem}`)
    return t as TileBody
  } catch (e) {
    if (e instanceof InvalidOperand) throw e
    throw new InvalidOperand(`XLSX_TILE: ${(e as Error).message}`)
  }
}

// ---- sheet-layout (S4 §6.4, R14) ----


/** The sheet's row and column bounds. Literals (F5): a test pins them to the engine's `MAX_ROWS` / `MAX_COLS`. */
export const SHEET_MAX_ROWS = 1_048_576, SHEET_MAX_COLS = 16_384
/** R14: per-list entry caps (row and column intervals and collapsed runs each count against `intervals`), and at
 * most 64 messages per answer. */
export const LAYOUT_LIMITS = Object.freeze({ intervals: 65_536, merges: 65_536, filters: 1_024, parts: 64 })
/** Every `sheet-layout` message is at most 64 KiB of UTF-8, the reqId included. */
export const LAYOUT_PART_BYTES = 65_536
/** Room kept in each part for the `"reqId":4294967295,` ws.ts adds after encoding (19 bytes), with slack. */
const REQ_ID_RESERVE = 64

export interface StreamSheetLayout {
  sheetId: string; viewGen: number
  defaults: { rowPx: number; colPx: number }
  rows: AxisSizeInterval[]; cols: AxisSizeInterval[]
  collapsedRows: Array<{ start: number; end: number }>; collapsedCols: Array<{ start: number; end: number }>
  frozen?: { rows: number; cols: number; topRow?: number; leftCol?: number }
  merges: SourceRect[]; used?: SourceRect
  filters: Array<{ kind: 'autoFilter' | 'table'; range: SourceRect; header: 0 | 1; totals: 0 | 1 }>
  sort: null; filterSummary: Array<{ col: number; kind: string }>
  authorHidden: boolean; counts: { visibleRows: number; totalRows: number }
  /** 4b (P11): the session's resized rows and columns, `[index, px]` ascending by index. Optional so a 4a-shaped layout
   * still types; the encoder omits an empty list (an identity view's bytes equal 4a's) and the decoder leaves an empty one absent (read `?? []`). */
  sessionRows?: Array<[number, number]>; sessionCols?: Array<[number, number]>
}

/** A layout over an R14 cap (any list, or more than 64 messages): the answer is refused `layout-too-large`. */
export class LayoutTooLarge extends Error { constructor(m = 'sheet layout over a cap') { super(m); this.name = 'LayoutTooLarge' } }
/** A received layout that breaks the wire contract (the sheet shows its load-failure notice). */
export class LayoutInvalid extends Error { constructor(m: string) { super(`sheet-layout: ${m}`); this.name = 'LayoutInvalid' } }

/** The lists that chunk, in part order, and their caps. */
const LAYOUT_LISTS = ['rows', 'cols', 'collapsedRows', 'collapsedCols', 'merges', 'filters', 'sessionRows', 'sessionCols'] as const
type LayoutList = typeof LAYOUT_LISTS[number]
const listCap = (k: LayoutList): number => k === 'merges' ? LAYOUT_LIMITS.merges : k === 'filters' ? LAYOUT_LIMITS.filters : LAYOUT_LIMITS.intervals
const listOf = (l: StreamSheetLayout, k: LayoutList): readonly unknown[] => l[k] ?? []
/** Longest `filterSummary[].kind` (4a sends none; 4c sends rule kinds). */
const FILTER_KIND_CHARS = 32
const utf8 = new TextEncoder()
const bytesOf = (v: unknown) => utf8.encode(JSON.stringify(v)).length

const interval = (x: AxisSizeInterval) => ({ start: x.start, end: x.end, sizeCss: x.sizeCss })
const run = (x: { start: number; end: number }) => ({ start: x.start, end: x.end })
const rectOf = (r: SourceRect) => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 })
const filterOf = (f: StreamSheetLayout['filters'][number]) => ({ kind: f.kind, range: rectOf(f.range), header: f.header, totals: f.totals })
const sized = (x: [number, number]) => [x[0], x[1]]
const copyItem: Record<LayoutList, (x: never) => object> = { rows: interval, cols: interval, collapsedRows: run, collapsedCols: run, merges: rectOf, filters: filterOf, sessionRows: sized, sessionCols: sized }

/**
 * The parts of one answer, without `reqId` (the sender adds it first): `{type: 'sheet-layout', seq, last, head?, …}`.
 * `head` (seq 0 only) holds every scalar field; the lists fill parts greedily in `LAYOUT_LISTS` order, each part at
 * most 64 KiB with its reqId. Throws `LayoutTooLarge` over a list cap or past 64 parts. Only the listed fields are
 * read, so nothing else on `l` can travel.
 */
export function encodeSheetLayout(l: StreamSheetLayout): object[] {
  for (const k of LAYOUT_LISTS) if (listOf(l, k).length > listCap(k)) throw new LayoutTooLarge(`${k}: ${listOf(l, k).length} > ${listCap(k)}`)
  if (l.filterSummary.length > LAYOUT_LIMITS.filters) throw new LayoutTooLarge('filterSummary')
  const head = {
    sheetId: l.sheetId, viewGen: l.viewGen, defaults: { rowPx: l.defaults.rowPx, colPx: l.defaults.colPx },
    ...(l.frozen ? { frozen: { rows: l.frozen.rows, cols: l.frozen.cols, ...(l.frozen.topRow !== undefined ? { topRow: l.frozen.topRow } : {}), ...(l.frozen.leftCol !== undefined ? { leftCol: l.frozen.leftCol } : {}) } } : {}),
    ...(l.used ? { used: rectOf(l.used) } : {}),
    sort: null, filterSummary: l.filterSummary.map((s) => ({ col: s.col, kind: s.kind })),
    authorHidden: l.authorHidden, counts: { visibleRows: l.counts.visibleRows, totalRows: l.counts.totalRows },
  }
  const budget = LAYOUT_PART_BYTES - REQ_ID_RESERVE
  const parts: Array<Record<string, unknown>> = []
  const open = (): [Record<string, unknown>, number] => {
    const p: Record<string, unknown> = { type: 'sheet-layout', seq: parts.length, last: false, ...(parts.length === 0 ? { head } : {}) }
    return [p, bytesOf(p)]
  }
  let [cur, size] = open()
  for (const k of LAYOUT_LISTS) {
    // The session sizes travel ascending by index whatever order the caller built them in (the decoder requires it).
    const items = k === 'sessionRows' || k === 'sessionCols' ? [...listOf(l, k) as Array<[number, number]>].sort((a, b) => a[0] - b[0]) : listOf(l, k)
    for (const x of items) {
      const item = copyItem[k](x as never), n = bytesOf(item)
      // A comma before the item, or `,"k":[]` around the list's first item in this part.
      const cost = (p: Record<string, unknown>) => n + (p[k] ? 1 : k.length + 6)
      if (size + cost(cur) > budget) {
        parts.push(cur)
        if (parts.length >= LAYOUT_LIMITS.parts) throw new LayoutTooLarge(`more than ${LAYOUT_LIMITS.parts} parts`)
        ;[cur, size] = open()
      }
      size += cost(cur)
      ;((cur[k] ??= []) as object[]).push(item)
    }
  }
  cur.last = true
  parts.push(cur)
  return parts
}

type Obj = Record<string, unknown>
const isObj = (v: unknown): v is Obj => typeof v === 'object' && v !== null && !Array.isArray(v)
const fail = (m: string): never => { throw new LayoutInvalid(m) }
/** `v` is a plain object with only `allowed` keys and every `required` one. */
function shape(v: unknown, allowed: readonly string[], required: readonly string[], what: string): Obj {
  if (!isObj(v)) return fail(`${what} is not an object`)
  for (const k of Object.keys(v)) if (!allowed.includes(k)) fail(`${what}: unknown key ${k}`)
  for (const k of required) if (!(k in v)) fail(`${what}: missing ${k}`)
  return v
}
const int = (v: unknown, lo: number, hi: number, what: string): number => {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < lo || v > hi) fail(`${what} is not an integer in [${lo}, ${hi}]`)
  return v as number
}
const size = (v: unknown, what: string): number => {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < 0) fail(`${what} is not a finite size ≥ 0`)
  return v as number
}
const rowIndex = (v: unknown, what: string) => int(v, 0, SHEET_MAX_ROWS - 1, what)
const colIndex = (v: unknown, what: string) => int(v, 0, SHEET_MAX_COLS - 1, what)
function readRect(v: unknown, what: string): SourceRect {
  const o = shape(v, ['r0', 'r1', 'c0', 'c1'], ['r0', 'r1', 'c0', 'c1'], what)
  const r = { r0: rowIndex(o.r0, `${what}.r0`), r1: rowIndex(o.r1, `${what}.r1`), c0: colIndex(o.c0, `${what}.c0`), c1: colIndex(o.c1, `${what}.c1`) }
  if (r.r0 > r.r1 || r.c0 > r.c1) fail(`${what} is inverted`)
  return r
}
function readRun(v: unknown, index: (v: unknown, what: string) => number, sized: boolean, what: string) {
  const keys = sized ? ['start', 'end', 'sizeCss'] : ['start', 'end']
  const o = shape(v, keys, keys, what)
  const start = index(o.start, `${what}.start`), end = index(o.end, `${what}.end`)
  if (start > end) fail(`${what} is inverted`)
  return sized ? { start, end, sizeCss: size(o.sizeCss, `${what}.sizeCss`) } : { start, end }
}
/** A session size entry: `[index, px]`. */
function readSized(v: unknown, index: (v: unknown, what: string) => number, what: string): [number, number] {
  if (!Array.isArray(v) || v.length !== 2) return fail(`${what} is not an [index, px] pair`)
  return [index(v[0], `${what}[0]`), size(v[1], `${what}[1]`)]
}
const readItem: Record<LayoutList, (v: unknown, what: string) => never> = {
  rows: (v, w) => readRun(v, rowIndex, true, w) as never,
  cols: (v, w) => readRun(v, colIndex, true, w) as never,
  collapsedRows: (v, w) => readRun(v, rowIndex, false, w) as never,
  collapsedCols: (v, w) => readRun(v, colIndex, false, w) as never,
  merges: (v, w) => readRect(v, w) as never,
  sessionRows: (v, w) => readSized(v, rowIndex, w) as never,
  sessionCols: (v, w) => readSized(v, colIndex, w) as never,
  filters: (v, w) => {
    const o = shape(v, ['kind', 'range', 'header', 'totals'], ['kind', 'range', 'header', 'totals'], w)
    if (o.kind !== 'autoFilter' && o.kind !== 'table') fail(`${w}.kind`)
    return { kind: o.kind, range: readRect(o.range, `${w}.range`), header: int(o.header, 0, 1, `${w}.header`), totals: int(o.totals, 0, 1, `${w}.totals`) } as never
  },
}
const HEAD_KEYS = ['sheetId', 'viewGen', 'defaults', 'frozen', 'used', 'sort', 'filterSummary', 'authorHidden', 'counts'] as const
function readHead(v: unknown): Omit<StreamSheetLayout, LayoutList> {
  const h = shape(v, HEAD_KEYS, ['sheetId', 'viewGen', 'defaults', 'sort', 'filterSummary', 'authorHidden', 'counts'], 'head')
  if (typeof h.sheetId !== 'string' || !/^(0|[1-9][0-9]{0,8})$/.test(h.sheetId)) fail('head.sheetId')
  const d = shape(h.defaults, ['rowPx', 'colPx'], ['rowPx', 'colPx'], 'head.defaults')
  if (h.sort !== null) fail('head.sort')
  if (typeof h.authorHidden !== 'boolean') fail('head.authorHidden')
  if (!Array.isArray(h.filterSummary) || h.filterSummary.length > LAYOUT_LIMITS.filters) fail('head.filterSummary')
  const filterSummary = (h.filterSummary as unknown[]).map((s, i) => {
    const o = shape(s, ['col', 'kind'], ['col', 'kind'], `head.filterSummary[${i}]`)
    if (typeof o.kind !== 'string' || o.kind.length === 0 || o.kind.length > FILTER_KIND_CHARS) fail(`head.filterSummary[${i}].kind`)
    return { col: colIndex(o.col, `head.filterSummary[${i}].col`), kind: o.kind as string }
  })
  const c = shape(h.counts, ['visibleRows', 'totalRows'], ['visibleRows', 'totalRows'], 'head.counts')
  const counts = { visibleRows: int(c.visibleRows, 0, SHEET_MAX_ROWS, 'head.counts.visibleRows'), totalRows: int(c.totalRows, 0, SHEET_MAX_ROWS, 'head.counts.totalRows') }
  if (counts.visibleRows > counts.totalRows) fail('head.counts')
  let frozen: StreamSheetLayout['frozen']
  if (h.frozen !== undefined) {
    const f = shape(h.frozen, ['rows', 'cols', 'topRow', 'leftCol'], ['rows', 'cols'], 'head.frozen')
    frozen = { rows: int(f.rows, 0, SHEET_MAX_ROWS - 1, 'head.frozen.rows'), cols: int(f.cols, 0, SHEET_MAX_COLS - 1, 'head.frozen.cols') }
    if (f.topRow !== undefined) frozen.topRow = rowIndex(f.topRow, 'head.frozen.topRow')
    if (f.leftCol !== undefined) frozen.leftCol = colIndex(f.leftCol, 'head.frozen.leftCol')
  }
  return {
    sheetId: h.sheetId as string, viewGen: int(h.viewGen, 0, 0xffffffff, 'head.viewGen'),
    defaults: { rowPx: size(d.rowPx, 'head.defaults.rowPx'), colPx: size(d.colPx, 'head.defaults.colPx') },
    ...(frozen ? { frozen } : {}), ...(h.used !== undefined ? { used: readRect(h.used, 'head.used') } : {}),
    sort: null, filterSummary, authorHidden: h.authorHidden as boolean, counts,
  }
}
const PART_KEYS = ['reqId', 'type', 'seq', 'last', 'head', ...LAYOUT_LISTS] as const

/**
 * One whole answer, its parts in arrival order. Throws `LayoutInvalid` for: no parts or more than 64; a part that is
 * not `sheet-layout`, or whose reqId differs from the first's; a seq gap; `head` absent from seq 0 or present on a
 * later part; `last` anywhere but on the final part (or not there); an unknown key at any depth; a list over its cap
 * (counted across parts); a non-integer index or count; a row index outside [0, SHEET_MAX_ROWS) or a column index
 * outside [0, SHEET_MAX_COLS); an inverted run or rect; a negative or non-finite size.
 */
export function decodeSheetLayout(parts: unknown[]): StreamSheetLayout {
  if (!Array.isArray(parts) || parts.length === 0 || parts.length > LAYOUT_LIMITS.parts) fail('part count')
  const lists = Object.fromEntries(LAYOUT_LISTS.map((k) => [k, [] as unknown[]])) as Record<LayoutList, unknown[]>
  let head: Omit<StreamSheetLayout, LayoutList> | undefined
  let reqId: unknown
  parts.forEach((raw, i) => {
    const p = shape(raw, PART_KEYS, ['type', 'seq', 'last'], `part ${i}`)
    if (p.type !== 'sheet-layout') fail(`part ${i}: type`)
    if (i === 0) reqId = p.reqId
    else if (p.reqId !== reqId) fail(`part ${i}: reqId`)
    if (p.seq !== i) fail(`part ${i}: seq ${String(p.seq)} (gap)`)
    if (typeof p.last !== 'boolean' || p.last !== (i === parts.length - 1)) fail(`part ${i}: last`)
    if (i === 0) head = readHead(p.head)
    else if ('head' in p) fail(`part ${i}: head after seq 0`)
    for (const k of LAYOUT_LISTS) {
      if (!(k in p)) continue
      const list = p[k]
      if (!Array.isArray(list)) fail(`part ${i}: ${k} is not a list`)
      if (lists[k].length + (list as unknown[]).length > listCap(k)) fail(`${k} over its cap`)
      for (const x of list as unknown[]) lists[k].push(readItem[k](x, `${k}[${lists[k].length}]`))
    }
  })
  for (const k of ['sessionRows', 'sessionCols'] as const) {
    const l = lists[k] as Array<[number, number]>
    for (let i = 1; i < l.length; i++) if (l[i]![0] <= l[i - 1]![0]) fail(`${k} is not strictly ascending`)
  }
  const { sessionRows, sessionCols, ...required } = lists
  const out = { ...head!, ...(required as unknown as Pick<StreamSheetLayout, Exclude<LayoutList, 'sessionRows' | 'sessionCols'>>) } as StreamSheetLayout
  // A layout with no session sizes leaves both keys absent (4a-shaped); consumers read `layout.sessionRows ?? []`.
  if (sessionRows!.length > 0) out.sessionRows = sessionRows as Array<[number, number]>
  if (sessionCols!.length > 0) out.sessionCols = sessionCols as Array<[number, number]>
  return out
}

// ---- Stage 4b: the `sheet-view@1` ops (S4 §6.7; plan P6, P7, P9–P11) ----

/** The HELLO capability that adds selection, find, copy, a11y text, fit and resize to a tiles-mode sheet (P1). */
export const SHEET_VIEW_FEATURE = 'sheet-view@1'
/** A selection holds at most 256 ranges (P7). */
export const SELECTION_LIMITS = Object.freeze({ ranges: 256 })
/** One `sheet-a11y` request covers at most 2,000 cells (S4 §6.7). */
export const A11Y_MAX_CELLS = 2_000
/** One `sheet-find` answer lists at most 5,000 hits; past that `truncated` is true (S3, P6). */
export const FIND_MAX_HITS = 5_000
/** A find query is at most 256 UTF-16 units (S3). */
export const QUERY_MAX_UNITS = 256
/** At most 4 `regions` on a `sheet-fit` request (P9). */
const FIT_MAX_REGIONS = 4

export interface SheetFindHit { sheetIndex: number; row: number; col: number; hidden?: 'filter' }
/** `total` is the number of hits sent; `truncated` says more existed (P6, F15). There is no cursor. */
export interface SheetFindAnswer { hits: SheetFindHit[]; total: number; truncated: boolean }
export interface Box { x: number; y: number; width: number; height: number }
export interface SheetFitRequest { sheetId: string; viewGen: number; axis: 'row' | 'col'; index: number; pane: Box; regions?: Box[] }
/** The resize arguments without the view generation; a union, so it is not `Omit`-able (F6). */
export type SheetResizeArgs = { sheetId: string } & ({ axis: 'row' | 'col'; index: number; size: number } | { reset: true })
export type SheetResizeRequest = SheetResizeArgs & { viewGen: number }
export interface SheetViewChange { sheetId: string; viewGen: number; change: { axis: 'row' | 'col'; index: number; size: number } | { reset: true } }

/** The selection as it travels (P7): interval-only, every range the product of its row and column intervals. */
export interface WireSelection {
  active: [number, number] | null
  anchor: [number, number] | null
  ranges: Array<{ rows: Array<[number, number]>; cols: Array<[number, number]> }>
}

const SHEET_ID = /^(0|[1-9][0-9]{0,8})$/
const isInt = (v: unknown, lo: number, hi: number): v is number => typeof v === 'number' && Number.isInteger(v) && v >= lo && v <= hi
const isFiniteGe0 = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v) && v >= 0
const hasOnly = (o: Obj, keys: readonly string[]) => Object.keys(o).every((k) => keys.includes(k))

/** The encoded form of `s`: only the three wire fields, so nothing else on the object travels. */
export function encodeWireSelection(s: WireSelection): object {
  return {
    active: s.active ? [s.active[0], s.active[1]] : null,
    anchor: s.anchor ? [s.anchor[0], s.anchor[1]] : null,
    ranges: s.ranges.map((r) => ({ rows: r.rows.map(([a, b]) => [a, b]), cols: r.cols.map(([a, b]) => [a, b]) })),
  }
}

function cellPair(v: unknown): [number, number] | null | undefined {
  if (v === null) return null
  if (!Array.isArray(v) || v.length !== 2 || !isInt(v[0], 0, SHEET_MAX_ROWS - 1) || !isInt(v[1], 0, SHEET_MAX_COLS - 1)) return undefined
  return [v[0], v[1]]
}
/** At least one inclusive `[lo, hi]` interval within `[0, max)`, ascending and disjoint; undefined for anything else. */
function intervalList(v: unknown, max: number): Array<[number, number]> | undefined {
  if (!Array.isArray(v) || v.length === 0) return undefined
  const out: Array<[number, number]> = []
  for (const p of v as unknown[]) {
    if (!Array.isArray(p) || p.length !== 2 || !isInt(p[0], 0, max - 1) || !isInt(p[1], 0, max - 1) || p[1] < p[0]) return undefined
    if (out.length > 0 && p[0] <= out[out.length - 1]![1]) return undefined
    out.push([p[0], p[1]])
  }
  return out
}

/** Strict: exactly `{active, anchor, ranges}`, ≤ 256 ranges, each exactly `{rows, cols}` of ascending disjoint in-range
 * intervals. Anything else (a `kind` key, a non-integer, an out-of-range index) is undefined (⇒ `policy-refused`). */
export function decodeWireSelection(v: unknown): WireSelection | undefined {
  if (!isObj(v) || !hasOnly(v, ['active', 'anchor', 'ranges']) || !('active' in v) || !('anchor' in v) || !('ranges' in v)) return undefined
  const active = cellPair(v.active), anchor = cellPair(v.anchor)
  if (active === undefined || anchor === undefined) return undefined
  if (!Array.isArray(v.ranges) || v.ranges.length > SELECTION_LIMITS.ranges) return undefined
  const ranges: WireSelection['ranges'] = []
  for (const r of v.ranges as unknown[]) {
    if (!isObj(r) || !hasOnly(r, ['rows', 'cols']) || !('rows' in r) || !('cols' in r)) return undefined
    const rows = intervalList(r.rows, SHEET_MAX_ROWS), cols = intervalList(r.cols, SHEET_MAX_COLS)
    if (!rows || !cols) return undefined
    ranges.push({ rows, cols })
  }
  return { active, anchor, ranges }
}

// ---- requests (server side): the typed args, or undefined (⇒ policy-refused) ----

const sheetIndexOf = (id: unknown, sheetCount: number): number | undefined =>
  typeof id === 'string' && SHEET_ID.test(id) && Number(id) < sheetCount ? Number(id) : undefined

/** `sheet-find {query, matchCase?, scope?}`. `wholeCell` may be absent or false (local find has none); `true` is refused. */
export function findArgs(m: Obj, sheetCount: number): { query: string; matchCase: boolean; scope: string } | undefined {
  const { query, matchCase, scope, wholeCell } = m
  if (typeof query !== 'string' || query.length === 0 || query.length > QUERY_MAX_UNITS) return undefined
  if (matchCase !== undefined && typeof matchCase !== 'boolean') return undefined
  if (wholeCell !== undefined && wholeCell !== false) return undefined
  if (scope !== undefined && scope !== 'workbook' && sheetIndexOf(scope, sheetCount) === undefined) return undefined
  return { query, matchCase: matchCase === true, scope: scope === undefined ? 'workbook' : (scope as string) }
}

/** `sheet-copy {sheetId, viewGen, selection}`. */
export function copyArgs(m: Obj, sheetCount: number): { sheetId: string; sheetIndex: number; viewGen: number; selection: WireSelection } | undefined {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  if (sheetIndex === undefined || !u32(m.viewGen)) return undefined
  const selection = decodeWireSelection(m.selection)
  if (!selection) return undefined
  return { sheetId: m.sheetId as string, sheetIndex, viewGen: m.viewGen, selection }
}

/** `sheet-a11y {sheetId, viewGen, rows: [[r0, r1], …], cols: [c0, c1]}`: ≤ 2,000 cells, columns in `[0, SHEET_MAX_COLS)`. */
export function a11yArgs(m: Obj, sheetCount: number): { sheetId: string; sheetIndex: number; viewGen: number; rows: Array<[number, number]>; cols: [number, number] } | undefined {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  if (sheetIndex === undefined || !u32(m.viewGen)) return undefined
  const rows = intervalList(m.rows, SHEET_MAX_ROWS)
  const c = m.cols
  if (!rows || !Array.isArray(c) || c.length !== 2 || !isInt(c[0], 0, SHEET_MAX_COLS - 1) || !isInt(c[1], 0, SHEET_MAX_COLS - 1) || c[1] < c[0]) return undefined
  let cells = 0
  for (const [a, b] of rows) cells += (b - a + 1) * (c[1] - c[0] + 1)
  if (cells > A11Y_MAX_CELLS) return undefined
  return { sheetId: m.sheetId as string, sheetIndex, viewGen: m.viewGen, rows, cols: [c[0], c[1]] }
}

function readBox(v: unknown): Box | undefined {
  if (!isObj(v) || !hasOnly(v, ['x', 'y', 'width', 'height'])) return undefined
  if (!isFiniteGe0(v.x) || !isFiniteGe0(v.y) || !isFiniteGe0(v.width) || !isFiniteGe0(v.height)) return undefined
  return { x: v.x, y: v.y, width: v.width, height: v.height }
}
const axisIndex = (axis: unknown, index: unknown): number | undefined =>
  axis === 'row' ? (isInt(index, 0, SHEET_MAX_ROWS - 1) ? index : undefined) : axis === 'col' ? (isInt(index, 0, SHEET_MAX_COLS - 1) ? index : undefined) : undefined

/** `sheet-fit {sheetId, viewGen, axis, index, pane, regions?}` (P9): boxes finite and non-negative, ≤ 4 regions. The
 * server clamps the boxes to the sheet extent. */
export function fitArgs(m: Obj, sheetCount: number): (SheetFitRequest & { sheetIndex: number }) | undefined {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  if (sheetIndex === undefined || !u32(m.viewGen)) return undefined
  const index = axisIndex(m.axis, m.index)
  const pane = readBox(m.pane)
  if (index === undefined || !pane) return undefined
  let regions: Box[] | undefined
  if (m.regions !== undefined) {
    if (!Array.isArray(m.regions) || m.regions.length > FIT_MAX_REGIONS) return undefined
    regions = []
    for (const r of m.regions as unknown[]) { const b = readBox(r); if (!b) return undefined; regions.push(b) }
  }
  return { sheetId: m.sheetId as string, sheetIndex, viewGen: m.viewGen, axis: m.axis as 'row' | 'col', index, pane, ...(regions ? { regions } : {}) }
}

/** `sheet-resize {sheetId, viewGen, axis, index, size}` or `{sheetId, viewGen, reset: true}` (P11); the two forms never
 * mix. `size` is any finite number ≥ 0; the server clamps it (`clampSize`). */
export function resizeArgs(m: Obj, sheetCount: number): (SheetResizeRequest & { sheetIndex: number }) | undefined {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  if (sheetIndex === undefined || !u32(m.viewGen)) return undefined
  const head = { sheetId: m.sheetId as string, sheetIndex, viewGen: m.viewGen }
  if (m.reset !== undefined) {
    if (m.reset !== true || m.axis !== undefined || m.index !== undefined || m.size !== undefined) return undefined
    return { ...head, reset: true }
  }
  const index = axisIndex(m.axis, m.index)
  if (index === undefined || !isFiniteGe0(m.size)) return undefined
  return { ...head, axis: m.axis as 'row' | 'col', index, size: m.size }
}

// ---- answers (client side): throw `InvalidOperand` for anything off the wire contract ----

const bad = (m: string): never => { throw new InvalidOperand(m) }
/** `m` is the message, or its list itself. */
const listOfMsg = (m: unknown, key: string, what: string, max: number): unknown[] => {
  const list = Array.isArray(m) ? m : isObj(m) ? m[key] : undefined
  if (!Array.isArray(list)) return bad(`${what}: ${key} is not a list`)
  if (list.length > max) return bad(`${what}: more than ${max} ${key}`)
  return list
}

/** `sheet-find-result` hits: `[sheetIndex, row, col]` or `[sheetIndex, row, col, 1]` (a hit hidden by the filter). */
export function decodeFindHits(m: unknown): SheetFindHit[] {
  return listOfMsg(m, 'hits', 'sheet-find-result', FIND_MAX_HITS).map((h, i): SheetFindHit => {
    if (!Array.isArray(h) || (h.length !== 3 && h.length !== 4)) return bad(`sheet-find-result: hits[${i}] shape`)
    if (!isInt(h[0], 0, MAX_SHEETS - 1) || !isInt(h[1], 0, SHEET_MAX_ROWS - 1) || !isInt(h[2], 0, SHEET_MAX_COLS - 1)) return bad(`sheet-find-result: hits[${i}] position`)
    if (h.length === 4 && h[3] !== 1) return bad(`sheet-find-result: hits[${i}] flag`)
    return { sheetIndex: h[0], row: h[1], col: h[2], ...(h.length === 4 ? { hidden: 'filter' as const } : {}) }
  })
}

/** `sheet-a11y-result` cells: `[row, col, text]`. */
export function decodeA11yCells(m: unknown): Array<{ row: number; col: number; text: string }> {
  return listOfMsg(m, 'cells', 'sheet-a11y-result', A11Y_MAX_CELLS).map((c, i) => {
    if (!Array.isArray(c) || c.length !== 3) return bad(`sheet-a11y-result: cells[${i}] shape`)
    if (!isInt(c[0], 0, SHEET_MAX_ROWS - 1) || !isInt(c[1], 0, SHEET_MAX_COLS - 1) || typeof c[2] !== 'string') return bad(`sheet-a11y-result: cells[${i}] fields`)
    return { row: c[0], col: c[1], text: c[2] }
  })
}

/** `sheet-fit-result {size}`: a finite size ≥ 0. */
export function decodeFitResult(m: unknown): number {
  if (!isObj(m) || !isFiniteGe0(m.size)) return bad('sheet-fit-result: size')
  return m.size
}

/** `sheet-view {sheetId, viewGen, change}`; `reqId` and `type` are the envelope. */
export function decodeSheetView(m: unknown): SheetViewChange {
  if (!isObj(m) || !hasOnly(m, ['reqId', 'type', 'sheetId', 'viewGen', 'change'])) return bad('sheet-view: shape')
  if (typeof m.sheetId !== 'string' || !SHEET_ID.test(m.sheetId) || !u32(m.viewGen)) return bad('sheet-view: sheetId/viewGen')
  const c = m.change
  if (!isObj(c)) return bad('sheet-view: change')
  if ('reset' in c) {
    if (c.reset !== true || Object.keys(c).length !== 1) return bad('sheet-view: change.reset')
    return { sheetId: m.sheetId, viewGen: m.viewGen, change: { reset: true } }
  }
  if (!hasOnly(c, ['axis', 'index', 'size']) || axisIndex(c.axis, c.index) === undefined || !isFiniteGe0(c.size)) return bad('sheet-view: change')
  return { sheetId: m.sheetId, viewGen: m.viewGen, change: { axis: c.axis as 'row' | 'col', index: c.index as number, size: c.size } }
}
