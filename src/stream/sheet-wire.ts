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
}

/** A layout over an R14 cap (any list, or more than 64 messages): the answer is refused `layout-too-large`. */
export class LayoutTooLarge extends Error { constructor(m = 'sheet layout over a cap') { super(m); this.name = 'LayoutTooLarge' } }
/** A received layout that breaks the wire contract (the sheet shows its load-failure notice). */
export class LayoutInvalid extends Error { constructor(m: string) { super(`sheet-layout: ${m}`); this.name = 'LayoutInvalid' } }

/** The lists that chunk, in part order, and their caps. */
const LAYOUT_LISTS = ['rows', 'cols', 'collapsedRows', 'collapsedCols', 'merges', 'filters'] as const
type LayoutList = typeof LAYOUT_LISTS[number]
const listCap = (k: LayoutList): number => k === 'merges' ? LAYOUT_LIMITS.merges : k === 'filters' ? LAYOUT_LIMITS.filters : LAYOUT_LIMITS.intervals
/** Longest `filterSummary[].kind` (4a sends none; 4c sends rule kinds). */
const FILTER_KIND_CHARS = 32
const utf8 = new TextEncoder()
const bytesOf = (v: unknown) => utf8.encode(JSON.stringify(v)).length

const interval = (x: AxisSizeInterval) => ({ start: x.start, end: x.end, sizeCss: x.sizeCss })
const run = (x: { start: number; end: number }) => ({ start: x.start, end: x.end })
const rectOf = (r: SourceRect) => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 })
const filterOf = (f: StreamSheetLayout['filters'][number]) => ({ kind: f.kind, range: rectOf(f.range), header: f.header, totals: f.totals })
const copyItem: Record<LayoutList, (x: never) => object> = { rows: interval, cols: interval, collapsedRows: run, collapsedCols: run, merges: rectOf, filters: filterOf }

/**
 * The parts of one answer, without `reqId` (the sender adds it first): `{type: 'sheet-layout', seq, last, head?, …}`.
 * `head` (seq 0 only) holds every scalar field; the lists fill parts greedily in `LAYOUT_LISTS` order, each part at
 * most 64 KiB with its reqId. Throws `LayoutTooLarge` over a list cap or past 64 parts. Only the listed fields are
 * read, so nothing else on `l` can travel.
 */
export function encodeSheetLayout(l: StreamSheetLayout): object[] {
  for (const k of LAYOUT_LISTS) if (l[k].length > listCap(k)) throw new LayoutTooLarge(`${k}: ${l[k].length} > ${listCap(k)}`)
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
    for (const x of l[k]) {
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
const readItem: Record<LayoutList, (v: unknown, what: string) => never> = {
  rows: (v, w) => readRun(v, rowIndex, true, w) as never,
  cols: (v, w) => readRun(v, colIndex, true, w) as never,
  collapsedRows: (v, w) => readRun(v, rowIndex, false, w) as never,
  collapsedCols: (v, w) => readRun(v, colIndex, false, w) as never,
  merges: (v, w) => readRect(v, w) as never,
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
  return { ...head!, ...(lists as unknown as Pick<StreamSheetLayout, LayoutList>) }
}
