/**
 * Stage 4a: the `sheet-tiles@1` wire constants (S4 §6, R12) and the `sheet-layout` codec (S4 §6.4, R14), shared by
 * the stream server and client. React-free, and it imports nothing from `src/xlsx/**` at runtime (types only, F5), so
 * the stream runtime's module set and the eager gates do not move: `TILE_ZOOM_STEPS`, `SHEET_MAX_ROWS` and
 * `SHEET_MAX_COLS` are literals, and tests pin them to the engine's constants.
 */
import type { AxisSizeInterval, CalendarDay, DateOp, DatePreset, NumberOp, TextOp } from '../xlsx/view/contracts'
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
  /** 4c (F1): the session's committed sort, or null (an identity view, whose bytes equal 4a's and 4b's). */
  sort: { range: WireRange; keys: WireSortSpec['keys'] } | null; filterSummary: Array<{ col: number; kind: string }>
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
    sort: l.sort === null ? null : (sortHead(l.sort) ?? fail('sort')), filterSummary: l.filterSummary.map((s) => ({ col: s.col, kind: s.kind })),
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
  const sort = h.sort === null ? null : (sortHead(h.sort) ?? fail('head.sort'))
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
    sort, filterSummary, authorHidden: h.authorHidden as boolean, counts,
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

// ---- Stage 4c: sort, filter, value lists, range detection and the view delta (S4 §6; plan P4–P7) ----

/** At most 64 sort keys and 64 filter columns (P4); Excel's own sort-level limit. */
export const SORT_FILTER_MAX_RULES = 64
/** A text operand or a value-list search is at most 256 UTF-16 units (S4 §6.6). */
export const FILTER_TEXT_UNITS = 256
/** `entries`: distinct entries per list; `page`: entries per answer; `intervals`: `ids` intervals per values rule (P6, P15). */
export const VALUE_LIST_LIMITS = Object.freeze({ entries: 10_000, page: 500, intervals: 4_096 })
/** A `sheet-order` answer: at most 160 parts, one band, at most one sheet of rows (P7). */
export const ORDER_LIMITS = Object.freeze({ parts: 160, bands: 1, rows: 1_048_576 })

/** A sort or filter range: inclusive source rows and columns, with the header and totals row counts. */
export interface WireRange { r0: number; r1: number; c0: number; c1: number; header: 0 | 1; totals: 0 | 1 }
export interface WireSortSpec { range: WireRange; keys: Array<{ col: number; order: 'asc' | 'desc' }>; caseSensitive: boolean }
export type WireTyped =
  | { kind: 'text'; op: TextOp; first: string; wildcards: boolean }
  | { kind: 'number'; op: NumberOp; first: number; second?: number }
  | { kind: 'date'; op: DateOp; first: CalendarDay; second?: CalendarDay }
  | { kind: 'datePreset'; preset: DatePreset }
export type WireRule =
  | WireTyped
  | { kind: 'compound'; connective: 'and' | 'or'; first: WireTyped; second: WireTyped }
  | { kind: 'values'; list: string; mode: 'include' | 'exclude'; ids: Array<[number, number]>; blanks: boolean }
  | { kind: 'committed' }
/** `columns` are ascending by column, each at most once. */
export interface WireFilterSpec { range: WireRange; columns: Array<[number, WireRule]> }
export interface SheetOrderHead {
  sheetId: string; viewGen: number; counts: { visible: number; total: number }
  sort: { range: WireRange; keys: WireSortSpec['keys'] } | null; filterSummary: Array<{ col: number; kind: string }>
}
/** `sourceAtOffset[i]` is the source row shown at display row `r0 + i`; `mask.words` is the display-slot bitset over `[r0, r1]`. */
export interface SheetOrder {
  head: SheetOrderHead
  bands: Array<{ r0: number; r1: number; sourceAtOffset: Uint32Array }>
  mask: { r0: number; r1: number; words: Uint32Array } | null
}
export interface ValuesPage {
  list: string
  items: Array<{ id: number; text: string; selected: boolean; error: boolean }>
  blanks: { present: boolean; selected: boolean }
  total: number; truncated: boolean; cursor?: string
  classification?: { kind: 'number' | 'date' | 'text' | 'empty'; hasNumeric: boolean }
}
export interface WireDetection { range: WireRange; source: 'table' | 'region' | 'used' | 'explicit'; header: 'metadata' | 'heuristic' | 'override'; overrideIgnored: boolean }

/** `o` has every `required` key, no key outside `required` and `optional`. */
function keysOk(o: Obj, required: readonly string[], optional: readonly string[] = []): boolean {
  return required.every((k) => k in o) && Object.keys(o).every((k) => required.includes(k) || optional.includes(k))
}
const UNITS_MAX_CURSOR = 256, LIST_ID_UNITS = 64

const RANGE_KEYS = ['r0', 'r1', 'c0', 'c1', 'header', 'totals'] as const
/** A range inside the grid with a body of at least one row, or undefined. */
function wireRange(v: unknown): WireRange | undefined {
  if (!isObj(v) || !keysOk(v, RANGE_KEYS)) return undefined
  const { r0, r1, c0, c1, header, totals } = v
  if (!isInt(r0, 0, SHEET_MAX_ROWS - 1) || !isInt(r1, 0, SHEET_MAX_ROWS - 1) || !isInt(c0, 0, SHEET_MAX_COLS - 1) || !isInt(c1, 0, SHEET_MAX_COLS - 1)) return undefined
  if (r1 < r0 || c1 < c0 || !isInt(header, 0, 1) || !isInt(totals, 0, 1) || r1 - r0 + 1 - header - totals < 1) return undefined
  return { r0, r1, c0, c1, header: header as 0 | 1, totals: totals as 0 | 1 }
}

/** 1 to 64 keys, each exactly `{col, order}` with `col` inside `range`'s columns. */
function sortKeys(v: unknown, range: WireRange): WireSortSpec['keys'] | undefined {
  if (!Array.isArray(v) || v.length === 0 || v.length > SORT_FILTER_MAX_RULES) return undefined
  const keys: WireSortSpec['keys'] = []
  for (const k of v as unknown[]) {
    if (!isObj(k) || !keysOk(k, ['col', 'order']) || !isInt(k.col, range.c0, range.c1) || (k.order !== 'asc' && k.order !== 'desc')) return undefined
    keys.push({ col: k.col, order: k.order })
  }
  return keys
}

/** The `{range, keys}` a layout or an order head carries. */
function sortHead(v: unknown): { range: WireRange; keys: WireSortSpec['keys'] } | undefined {
  if (!isObj(v) || !keysOk(v, ['range', 'keys'])) return undefined
  const range = wireRange(v.range), keys = range && sortKeys(v.keys, range)
  return range && keys ? { range, keys } : undefined
}

const TEXT_OPS: readonly TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
const NUMBER_OPS: readonly NumberOp[] = ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'between']
const DATE_OPS: readonly DateOp[] = ['eq', 'neq', 'before', 'after', 'between']
const DATE_PRESETS: readonly DatePreset[] = ['today', 'yesterday', 'tomorrow', 'thisWeek', 'thisMonth', 'thisYear']
const DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

function calendarDay(v: unknown): CalendarDay | undefined {
  if (!isObj(v) || !keysOk(v, ['year', 'month', 'day']) || !isInt(v.year, 1, 9999) || !isInt(v.month, 1, 12)) return undefined
  const leap = v.year % 4 === 0 && (v.year % 100 !== 0 || v.year % 400 === 0)
  const days = v.month === 2 && leap ? 29 : DAYS_IN_MONTH[v.month - 1]!
  return isInt(v.day, 1, days) ? { year: v.year, month: v.month, day: v.day } : undefined
}

/** Only the first and, for `between`, the second operand; an operand on any other op is refused. */
function typedRule(v: unknown): WireTyped | undefined {
  if (!isObj(v)) return undefined
  switch (v.kind) {
    case 'text':
      if (!keysOk(v, ['kind', 'op', 'first', 'wildcards']) || !TEXT_OPS.includes(v.op as TextOp)) return undefined
      if (typeof v.first !== 'string' || v.first.length > FILTER_TEXT_UNITS || typeof v.wildcards !== 'boolean') return undefined
      return { kind: 'text', op: v.op as TextOp, first: v.first, wildcards: v.wildcards }
    case 'number': {
      if (!keysOk(v, ['kind', 'op', 'first'], ['second']) || !NUMBER_OPS.includes(v.op as NumberOp)) return undefined
      const between = v.op === 'between'
      if (!Number.isFinite(v.first) || (between ? !Number.isFinite(v.second) : 'second' in v)) return undefined
      return { kind: 'number', op: v.op as NumberOp, first: v.first as number, ...(between ? { second: v.second as number } : {}) }
    }
    case 'date': {
      if (!keysOk(v, ['kind', 'op', 'first'], ['second']) || !DATE_OPS.includes(v.op as DateOp)) return undefined
      const first = calendarDay(v.first), between = v.op === 'between', second = between ? calendarDay(v.second) : undefined
      if (!first || (between ? !second : 'second' in v)) return undefined
      return { kind: 'date', op: v.op as DateOp, first, ...(second ? { second } : {}) }
    }
    case 'datePreset':
      return keysOk(v, ['kind', 'preset']) && DATE_PRESETS.includes(v.preset as DatePreset) ? { kind: 'datePreset', preset: v.preset as DatePreset } : undefined
    default:
      return undefined
  }
}

/** Ascending, disjoint inclusive `[i0, i1]` intervals, at most 4,096. Whether an id exists is the server's. */
function idIntervals(v: unknown): Array<[number, number]> | undefined {
  if (!Array.isArray(v) || v.length > VALUE_LIST_LIMITS.intervals) return undefined
  const out: Array<[number, number]> = []
  for (const p of v as unknown[]) {
    if (!Array.isArray(p) || p.length !== 2 || !isInt(p[0], 0, 0xffffffff) || !isInt(p[1], 0, 0xffffffff) || p[1] < p[0]) return undefined
    if (out.length > 0 && p[0] <= out[out.length - 1]![1]) return undefined
    out.push([p[0], p[1]])
  }
  return out
}

const listId = (v: unknown): v is string => typeof v === 'string' && v.length > 0 && v.length <= LIST_ID_UNITS

function wireRule(v: unknown): WireRule | undefined {
  if (!isObj(v)) return undefined
  if (v.kind === 'committed') return keysOk(v, ['kind']) ? { kind: 'committed' } : undefined
  if (v.kind === 'values') {
    if (!keysOk(v, ['kind', 'list', 'mode', 'ids', 'blanks']) || !listId(v.list) || (v.mode !== 'include' && v.mode !== 'exclude') || typeof v.blanks !== 'boolean') return undefined
    const ids = idIntervals(v.ids)
    return ids ? { kind: 'values', list: v.list, mode: v.mode, ids, blanks: v.blanks } : undefined
  }
  if (v.kind === 'compound') {
    if (!keysOk(v, ['kind', 'connective', 'first', 'second']) || (v.connective !== 'and' && v.connective !== 'or')) return undefined
    const first = typedRule(v.first), second = typedRule(v.second)
    return first && second ? { kind: 'compound', connective: v.connective, first, second } : undefined
  }
  return typedRule(v)
}

/** At most 64 `[col, rule]` pairs, ascending and distinct, each column inside `range`'s columns. */
function filterColumns(v: unknown, range: WireRange): WireFilterSpec['columns'] | undefined {
  if (!Array.isArray(v) || v.length > SORT_FILTER_MAX_RULES) return undefined
  const out: WireFilterSpec['columns'] = []
  for (const p of v as unknown[]) {
    if (!Array.isArray(p) || p.length !== 2 || !isInt(p[0], range.c0, range.c1) || (out.length > 0 && p[0] <= out[out.length - 1]![0])) return undefined
    const rule = wireRule(p[1])
    if (!rule) return undefined
    out.push([p[0], rule])
  }
  return out
}

type SheetArgs = { sheetId: string; sheetIndex: number }
const sheetArgs = (m: Obj, sheetCount: number): (SheetArgs & { viewGen: number }) | undefined => {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  return sheetIndex === undefined || !u32(m.viewGen) ? undefined : { sheetId: m.sheetId as string, sheetIndex, viewGen: m.viewGen }
}

/** `sheet-sort {sheetId, viewGen, range, keys, caseSensitive?}` (P4's shape rules; targets and coverage are the server's). */
export function sortArgs(m: Obj, sheetCount: number): (SheetArgs & { viewGen: number } & WireSortSpec) | undefined {
  const head = sheetArgs(m, sheetCount), range = wireRange(m.range)
  if (!head || !range) return undefined
  const keys = sortKeys(m.keys, range)
  if (!keys || (m.caseSensitive !== undefined && typeof m.caseSensitive !== 'boolean')) return undefined
  return { ...head, range, keys, caseSensitive: m.caseSensitive === true }
}

/** `sheet-filter {sheetId, viewGen, range, columns: [[col, rule], …]}`. An empty `columns` clears the filter. */
export function filterArgs(m: Obj, sheetCount: number): (SheetArgs & { viewGen: number } & WireFilterSpec) | undefined {
  const head = sheetArgs(m, sheetCount), range = wireRange(m.range)
  if (!head || !range) return undefined
  const columns = filterColumns(m.columns, range)
  return columns ? { ...head, range, columns } : undefined
}

export type ValuesArgs = SheetArgs & { viewGen: number } & ({ cursor: string } | { range: WireRange; col: number; search?: { text: string; wildcards: boolean } })
/** `filter-values {sheetId, viewGen, range, col, search?}` for a first page, or `{sheetId, viewGen, cursor}` for a later one; never both. */
export function valuesArgs(m: Obj, sheetCount: number): ValuesArgs | undefined {
  const head = sheetArgs(m, sheetCount)
  if (!head) return undefined
  if (m.cursor !== undefined) {
    if (typeof m.cursor !== 'string' || m.cursor.length === 0 || m.cursor.length > UNITS_MAX_CURSOR) return undefined
    if (m.range !== undefined || m.col !== undefined || m.search !== undefined) return undefined
    return { ...head, cursor: m.cursor }
  }
  const range = wireRange(m.range)
  if (!range || !isInt(m.col, range.c0, range.c1)) return undefined
  if (m.search === undefined) return { ...head, range, col: m.col }
  const s = m.search
  if (!isObj(s) || !keysOk(s, ['text', 'wildcards']) || typeof s.text !== 'string' || s.text.length > FILTER_TEXT_UNITS || typeof s.wildcards !== 'boolean') return undefined
  return { ...head, range, col: m.col, search: { text: s.text, wildcards: s.wildcards } }
}

/** `range-detect {sheetId, viewGen, kind, selection, override?: {header}}`. */
export function detectArgs(m: Obj, sheetCount: number): (SheetArgs & { viewGen: number; kind: 'sort' | 'filter'; selection: WireSelection; override?: { header: 0 | 1 } }) | undefined {
  const head = sheetArgs(m, sheetCount)
  if (!head || (m.kind !== 'sort' && m.kind !== 'filter')) return undefined
  const selection = decodeWireSelection(m.selection)
  if (!selection) return undefined
  if (m.override === undefined) return { ...head, kind: m.kind, selection }
  const o = m.override
  if (!isObj(o) || !keysOk(o, ['header']) || !isInt(o.header, 0, 1)) return undefined
  return { ...head, kind: m.kind, selection, override: { header: o.header as 0 | 1 } }
}

/** `sheet-order {sheetId}`. */
export function orderArgs(m: Obj, sheetCount: number): SheetArgs | undefined {
  const sheetIndex = sheetIndexOf(m.sheetId, sheetCount)
  return sheetIndex === undefined ? undefined : { sheetId: m.sheetId as string, sheetIndex }
}

// ---- sheet-order parts (P7) ----

function orderHead(v: unknown): SheetOrderHead | undefined {
  if (!isObj(v) || !keysOk(v, ['sheetId', 'viewGen', 'counts', 'sort', 'filterSummary'])) return undefined
  if (typeof v.sheetId !== 'string' || !SHEET_ID.test(v.sheetId) || !u32(v.viewGen)) return undefined
  const c = v.counts
  if (!isObj(c) || !keysOk(c, ['visible', 'total']) || !isInt(c.visible, 0, SHEET_MAX_ROWS) || !isInt(c.total, 0, SHEET_MAX_ROWS) || c.visible > c.total) return undefined
  const sort = v.sort === null ? null : sortHead(v.sort)
  if (sort === undefined || !Array.isArray(v.filterSummary) || v.filterSummary.length > LAYOUT_LIMITS.filters) return undefined
  const filterSummary: SheetOrderHead['filterSummary'] = []
  for (const s of v.filterSummary as unknown[]) {
    if (!isObj(s) || !keysOk(s, ['col', 'kind']) || !isInt(s.col, 0, SHEET_MAX_COLS - 1) || typeof s.kind !== 'string' || s.kind.length === 0 || s.kind.length > FILTER_KIND_CHARS) return undefined
    filterSummary.push({ col: s.col, kind: s.kind })
  }
  return { sheetId: v.sheetId, viewGen: v.viewGen, counts: { visible: c.visible, total: c.total }, sort, filterSummary }
}

/** `entries` is a permutation of `[r0, r1]`. */
function isPermutation(entries: Uint32Array, r0: number, r1: number): boolean {
  const n = r1 - r0 + 1
  if (entries.length !== n) return false
  const seen = new Uint8Array(n)
  for (const e of entries) {
    if (e < r0 || e > r1 || seen[e - r0]) return false
    seen[e - r0] = 1
  }
  return true
}
const rowSpan = (r0: unknown, r1: unknown): number | undefined =>
  isInt(r0, 0, SHEET_MAX_ROWS - 1) && isInt(r1, 0, SHEET_MAX_ROWS - 1) && r1 >= r0 && r1 - r0 + 1 <= ORDER_LIMITS.rows ? r1 - r0 + 1 : undefined
const maskWords = (rows: number) => Math.ceil(rows / 32)

function wordsToBase64(words: Uint32Array, from: number, count: number): string {
  const bytes = new Uint8Array(count * 4), v = new DataView(bytes.buffer)
  for (let i = 0; i < count; i++) v.setUint32(i * 4, words[from + i]!, true)
  let bin = ''
  for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode(...bytes.subarray(i, i + 8192))
  return btoa(bin)
}
const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/
function base64ToWords(data: unknown, what: string): Uint32Array {
  if (typeof data !== 'string' || data.length === 0 || data.length % 4 !== 0 || !BASE64.test(data)) return bad(`sheet-order: ${what} data`)
  const bin = atob(data)
  if (bin.length === 0 || bin.length % 4 !== 0) return bad(`sheet-order: ${what} data length`)
  const out = new Uint32Array(bin.length / 4), v = new DataView(new ArrayBuffer(bin.length))
  for (let i = 0; i < bin.length; i++) v.setUint8(i, bin.charCodeAt(i))
  for (let i = 0; i < out.length; i++) out[i] = v.getUint32(i * 4, true)
  return out
}

/** Words per part for a chunk key: the largest count whose base64 fits the part budget beside its envelope. */
function wordsPerPart(key: 'band' | 'mask'): number {
  const probe = { type: 'sheet-order', seq: 9_999, last: false, [key]: { r0: SHEET_MAX_ROWS, r1: SHEET_MAX_ROWS, from: SHEET_MAX_ROWS, data: '' } }
  const room = Math.floor((LAYOUT_PART_BYTES - REQ_ID_RESERVE - bytesOf(probe)) / 4)
  return Math.floor((room * 3) / 4)
}

/**
 * The parts of one `sheet-order` answer, without `reqId`: `{type, seq, last, head?, band?, mask?}`. `head` travels alone
 * in seq 0; the band, then the mask, follow as `{r0, r1, from, data}` chunks (base64 of little-endian `Uint32`s from
 * entry `from`), each part at most 64 KiB with its reqId. Throws `InvalidOperand` for an order that breaks a cap or
 * the permutation.
 */
export function encodeSheetOrder(o: SheetOrder): object[] {
  const head = orderHead(o.head)
  if (!head) return bad('sheet-order: head')
  if (o.bands.length > ORDER_LIMITS.bands) return bad('sheet-order: more than one band')
  const bodies: object[] = [{ head }]
  for (const b of o.bands) {
    if (rowSpan(b.r0, b.r1) === undefined || !isPermutation(b.sourceAtOffset, b.r0, b.r1)) return bad('sheet-order: band is not a permutation of its rows')
    const per = wordsPerPart('band')
    for (let from = 0; from < b.sourceAtOffset.length; from += per) bodies.push({ band: { r0: b.r0, r1: b.r1, from, data: wordsToBase64(b.sourceAtOffset, from, Math.min(per, b.sourceAtOffset.length - from)) } })
  }
  if (o.mask) {
    const m = o.mask, rows = rowSpan(m.r0, m.r1)
    if (rows === undefined || m.words.length !== maskWords(rows)) return bad('sheet-order: mask size')
    const per = wordsPerPart('mask')
    for (let from = 0; from < m.words.length; from += per) bodies.push({ mask: { r0: m.r0, r1: m.r1, from, data: wordsToBase64(m.words, from, Math.min(per, m.words.length - from)) } })
  }
  if (bodies.length > ORDER_LIMITS.parts) return bad(`sheet-order: more than ${ORDER_LIMITS.parts} parts`)
  return bodies.map((b, i) => ({ type: 'sheet-order', seq: i, last: i === bodies.length - 1, ...b }))
}

const ORDER_PART_KEYS = ['reqId', 'type', 'seq', 'last', 'head', 'band', 'mask'] as const
interface Chunked { r0: number; r1: number; words: Uint32Array; filled: number }

/** Reads one `{r0, r1, from, data}` chunk into `acc` (created by the first), requiring it to continue exactly where the last ended. */
function takeChunk(acc: Chunked | undefined, v: unknown, what: 'band' | 'mask'): Chunked {
  if (!isObj(v) || !keysOk(v, ['r0', 'r1', 'from', 'data'])) return bad(`sheet-order: ${what}`)
  const rows = rowSpan(v.r0, v.r1)
  if (rows === undefined) return bad(`sheet-order: ${what} rows`)
  const chunk = base64ToWords(v.data, what)
  const total = what === 'band' ? rows : maskWords(rows)
  const cur = acc ?? { r0: v.r0 as number, r1: v.r1 as number, words: new Uint32Array(total), filled: 0 }
  if (cur.r0 !== v.r0 || cur.r1 !== v.r1) return bad(`sheet-order: ${what} range changed`)
  if (v.from !== cur.filled) return bad(`sheet-order: ${what} ${(v.from as number) < cur.filled ? 'overlap' : 'gap'}`)
  if (cur.filled + chunk.length > total) return bad(`sheet-order: ${what} past its rows`)
  cur.words.set(chunk, cur.filled)
  cur.filled += chunk.length
  return cur
}

/**
 * One whole answer, its parts in arrival order. Strict: throws `InvalidOperand` for no parts or more than 160; a part
 * that is not `sheet-order`, whose reqId differs from the first's, whose `seq` is not its position, or whose `last` is
 * not exactly on the final part; a `head` anywhere but seq 0, or missing there; a band or mask chunk that leaves a gap
 * or an overlap, changes range, or runs past its rows; an unfinished band or mask; a band that is not a permutation of
 * its rows; an unknown key at any depth.
 */
export function decodeSheetOrder(parts: unknown[]): SheetOrder {
  if (!Array.isArray(parts) || parts.length === 0 || parts.length > ORDER_LIMITS.parts) return bad('sheet-order: part count')
  let head: SheetOrderHead | undefined, reqId: unknown, band: Chunked | undefined, mask: Chunked | undefined
  parts.forEach((raw, i) => {
    if (!isObj(raw) || !hasOnly(raw, ORDER_PART_KEYS) || raw.type !== 'sheet-order') return bad(`sheet-order: part ${i}`)
    if (i === 0) reqId = raw.reqId
    else if (raw.reqId !== reqId) return bad(`sheet-order: part ${i} reqId`)
    if (raw.seq !== i) return bad(`sheet-order: part ${i} seq ${String(raw.seq)}`)
    if (raw.last !== (i === parts.length - 1)) return bad(`sheet-order: part ${i} last`)
    if (i === 0) {
      if (!('head' in raw)) return bad('sheet-order: no head')
      head = orderHead(raw.head)
      if (!head) return bad('sheet-order: head')
    } else if ('head' in raw) return bad(`sheet-order: head on part ${i}`)
    if ('band' in raw) band = takeChunk(band, raw.band, 'band')
    if ('mask' in raw) mask = takeChunk(mask, raw.mask, 'mask')
  })
  if (band && (band.filled !== band.words.length || !isPermutation(band.words, band.r0, band.r1))) return bad('sheet-order: band is not a permutation of its rows')
  if (mask && mask.filled !== mask.words.length) return bad('sheet-order: mask is incomplete')
  return {
    head: head!,
    bands: band ? [{ r0: band.r0, r1: band.r1, sourceAtOffset: band.words }] : [],
    mask: mask ? { r0: mask.r0, r1: mask.r1, words: mask.words } : null,
  }
}

// ---- value lists and range detection (P6) ----

const VALUES_PAGE_KEYS = ['reqId', 'type', 'list', 'items', 'blanks', 'total', 'truncated', 'cursor', 'classification'] as const
const CLASS_KINDS = ['number', 'date', 'text', 'empty'] as const

/** `filter-values-result {list, items: [[id, text, flags], …], blanks, total, truncated, cursor?, classification?}`; `flags` is 1 = selected, 2 = every merged value is an error. */
export function decodeValuesPage(m: unknown): ValuesPage {
  if (!isObj(m) || !hasOnly(m, VALUES_PAGE_KEYS)) return bad('filter-values-result: shape')
  if (!listId(m.list) || !isInt(m.total, 0, VALUE_LIST_LIMITS.entries) || typeof m.truncated !== 'boolean') return bad('filter-values-result: fields')
  const total = m.total, raw = m.items
  if (!Array.isArray(raw) || raw.length > VALUE_LIST_LIMITS.page) return bad('filter-values-result: items')
  const items = raw.map((it, i) => {
    if (!Array.isArray(it) || it.length !== 3) return bad(`filter-values-result: items[${i}] shape`)
    if (!isInt(it[0], 0, total - 1) || typeof it[1] !== 'string' || it[1].length > FILTER_TEXT_UNITS || !isInt(it[2], 0, 3)) return bad(`filter-values-result: items[${i}] fields`)
    return { id: it[0], text: it[1], selected: (it[2] & 1) === 1, error: (it[2] & 2) === 2 }
  })
  const b = m.blanks
  if (!isObj(b) || !keysOk(b, ['present', 'selected']) || typeof b.present !== 'boolean' || typeof b.selected !== 'boolean') return bad('filter-values-result: blanks')
  const out: ValuesPage = { list: m.list, items, blanks: { present: b.present, selected: b.selected }, total, truncated: m.truncated }
  if (m.cursor !== undefined) {
    if (typeof m.cursor !== 'string' || m.cursor.length === 0 || m.cursor.length > UNITS_MAX_CURSOR) return bad('filter-values-result: cursor')
    out.cursor = m.cursor
  }
  if (m.classification !== undefined) {
    const c = m.classification
    if (!isObj(c) || !keysOk(c, ['kind', 'hasNumeric']) || !CLASS_KINDS.includes(c.kind as typeof CLASS_KINDS[number]) || typeof c.hasNumeric !== 'boolean') return bad('filter-values-result: classification')
    out.classification = { kind: c.kind as typeof CLASS_KINDS[number], hasNumeric: c.hasNumeric }
  }
  return out
}

const DETECT_SOURCES = ['table', 'region', 'used', 'explicit'] as const
const DETECT_HEADERS = ['metadata', 'heuristic', 'override'] as const
/** `range-detect-result {range, source, header, overrideIgnored}`. */
export function decodeDetection(m: unknown): WireDetection {
  if (!isObj(m) || !hasOnly(m, ['reqId', 'type', 'range', 'source', 'header', 'overrideIgnored']) || !('range' in m)) return bad('range-detect-result: shape')
  const range = wireRange(m.range)
  if (!range || !DETECT_SOURCES.includes(m.source as typeof DETECT_SOURCES[number]) || !DETECT_HEADERS.includes(m.header as typeof DETECT_HEADERS[number]) || typeof m.overrideIgnored !== 'boolean') return bad('range-detect-result: fields')
  return { range, source: m.source as WireDetection['source'], header: m.header as WireDetection['header'], overrideIgnored: m.overrideIgnored }
}
