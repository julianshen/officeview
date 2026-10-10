/** S0.2 deterministic fixture family (spec §12.5, plan S0.2).
 * Every builder derives from (seed, row, col): a counter-based hash mixes the
 * seed into values (F1/F2/F3/F4/F5/F7; F8 parity vectors stay fixed by design
 * and F6 stays arithmetic per spec). No Math.random, no rows×cols allocation.
 * Sparse fixtures index only populated coordinates; dense fixtures answer from
 * arithmetic over row/column intervals clamped to each fixture's used extent,
 * so no summary ever enumerates the 2^34 domain. */
import { createHash } from 'node:crypto'
import type {
  AuthoritativeCell, CellRead, ColumnEntry, CompactSourceRowSet, Completeness,
  OverviewGroup, Page, ReadOptions, SingleCellPayload, SourceErrorCode, SourceStamp, SourceStore,
  SparseRowPage, StructuralResult, ValuePolicy,
} from '../../src/xlsx/view/contracts'
import type { SourceCell, SourceRect, SourceRow } from '../../src/xlsx/view/coordinates'
import {
  MAX_COLS, MAX_ROWS, cellId, isValidCoordinate, isValidSourceRect,
} from '../../src/xlsx/view/coordinates'

export type FixtureId = 'F1' | 'F2' | 'F3' | 'F4' | 'F5' | 'F6' | 'F7' | 'F8'

/** FNV-1a over UTF-16 code units; 32-bit integer ops are engine-deterministic. */
export function hashSeed(text: string): number {
  let hash = 0x811c9dc5
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i)
    hash = Math.imul(hash, 0x01000193)
  }
  return hash >>> 0
}

/** Counter-based avalanche mixer: the per-address value stream for a seed.
 * Uses only 32-bit integer ops (no cell-ID packing); distinct seeds diverge. */
export function cellHash(seed: number, row: number, col: number): number {
  let h = (seed >>> 0) ^ 0x9e3779b9
  h = Math.imul(h ^ row, 0x85ebca6b)
  h = Math.imul(h ^ col, 0xc2b2ae35)
  h ^= h >>> 13
  h = Math.imul(h, 0x5bd1e995)
  h ^= h >>> 15
  return h >>> 0
}

/** Mulberry32: tiny, fast, fully specified by 32-bit integer semantics. */
export function createPrng(seed: number): () => number {
  let state = seed >>> 0
  return () => {
    state = (state + 0x6d2b79f5) >>> 0
    let t = Math.imul(state ^ (state >>> 15), 1 | state)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

export const FIXTURE_SEEDS: Record<FixtureId, number> = {
  F1: hashSeed('xlsx-viewer:F1'), F2: hashSeed('xlsx-viewer:F2'),
  F3: hashSeed('xlsx-viewer:F3'), F4: hashSeed('xlsx-viewer:F4'),
  F5: hashSeed('xlsx-viewer:F5'), F6: hashSeed('xlsx-viewer:F6'),
  F7: hashSeed('xlsx-viewer:F7'), F8: hashSeed('xlsx-viewer:F8'),
}

/** Style indexes shared by the builders and the streaming package writer. */
export const STYLE_DEFAULT = 0
export const STYLE_HEADER = 1
export const STYLE_NUMBER = 2
export const STYLE_DATE = 3

/** ST_Xstring encoding: lone surrogates, carriage returns and C0 controls
 * become _xHHHH_ (XML line-ending normalization would otherwise lose \r);
 * literal _xHHHH_ text is guarded so single-pass decoding stays exact. */
export function encodeXlsxString(value: string): string {
  const guarded = value.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_')
  let out = ''
  for (let i = 0; i < guarded.length; i++) {
    const unit = guarded.charCodeAt(i)
    const lead = unit >= 0xd800 && unit <= 0xdbff
    const trail = unit >= 0xdc00 && unit <= 0xdfff
    if (lead) {
      const next = guarded.charCodeAt(i + 1)
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += guarded[i] + guarded[i + 1]
        i++
        continue
      }
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    if (trail) {
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    if (unit < 0x20 && unit !== 0x09 && unit !== 0x0a) {
      out += `_x${unit.toString(16).toUpperCase().padStart(4, '0')}_`
      continue
    }
    out += guarded[i]
  }
  return escapeXml(out)
}

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export interface FixtureCell {
  value: string | number | boolean | null
  formula?: string
  hasCachedValue: boolean
  valueIsError: boolean
  styleIndex: number
  sharedFormula?: { si: number; ref?: string }
  arrayRef?: string
  dynamicArray?: boolean
  formulaType?: string
  ca?: boolean
}

export interface FixtureParams {
  seed?: number
  /** F1 sampled NNZ including corners (default 100000). */
  nnz?: number
  /** Row count override for F2/F3/F4/F7 reduced-scale tests. */
  rowCount?: number
  /** Column count override for F3 reduced-scale tests. */
  colCount?: number
  /** F3 dense (default) or sparse-wide lattice variant. */
  variant?: 'dense' | 'sparse-wide'
}

export interface FixtureStructure {
  merges: SourceRect[]
  tables: Array<{ range: SourceRect; headerRows: 0 | 1; totalsRows: 0 | 1 }>
  spills: Array<{ anchor: SourceCell; kind: 'fixed' | 'dynamic'; saved?: SourceRect; evaluated?: SourceRect }>
  drawings: Array<
    | { kind: 'oneCell'; id: string; from: SourceCell; bounds: SourceRect }
    | { kind: 'twoCell'; id: string; from: SourceCell; to: SourceCell; bounds: SourceRect }
    | { kind: 'absolute'; id: string }
  >
  /** Sparse style bands for the S1 metadata importer (no structural-query field carries them). */
  styleIntervals: Array<{ range: SourceRect; styleIndex: number }>
}

export interface DomainSummary {
  totalRows: number
  populatedRows: number
  populatedCells: number
  uniformStyle: number | null
}

export interface SyntheticStoreOptions {
  valuePolicy?: ValuePolicy
  pending?: SourceRect[]
  errors?: Array<{ range: SourceRect; code: 'page-error' | 'storage-error'; retryable: boolean }>
  /** Probe hook for boundedness tests: replaces the backend cell reader. */
  read?: Partial<CellReader>
}

/** Cheap per-address predicates (no value materialization) for S1/S5 consumers. */
export interface CellReader {
  cellAt(row: number, col: number): FixtureCell | null
  isPopulated(row: number, col: number): boolean
  /** Valid only where isPopulated is true; must agree with cellAt's styleIndex. */
  styleOf(row: number, col: number): number
}

export interface CellRef {
  row: number
  col: number
  cell: FixtureCell
}

export interface FixtureHandle extends CellReader {
  readonly id: FixtureId
  readonly seed: number
  readonly rowCount: number
  readonly colCount: number
  readonly used: SourceRect
  nnz(): number
  structure(): FixtureStructure
  /** Exact enumeration for small enumerable fixtures (F1 sampled set, F5 data, F8 cases). */
  enumerateHead(limit: number): CellRef[]
  summarizeDomain(rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>): DomainSummary
  scanCells(range: SourceRect, offset: number, limit: number, read?: CellReader, exclude?: SourceRect[]): CellRef[]
  createStore(options?: SyntheticStoreOptions): SourceStore
  canonicalSamples(): string
}

export function sha256Hex(text: string): string {
  return createHash('sha256').update(text, 'utf8').digest('hex')
}

export function readOptionsFor(store: SourceStore): ReadOptions {
  return { signal: new AbortController().signal, stamp: { ...store.stamp } }
}

/** Head-scale query hashes (documented scope) shared by the script and tests. */
export async function headQueryHashes(handle: FixtureHandle): Promise<{ readRowsHead: string; overviewHead: string }> {
  const store = handle.createStore()
  const options = readOptionsFor(store)
  const headCols = Math.min(3, handle.colCount - 1)
  const rowsPage = await store.readRows({ r0: 0, r1: 4, c0: 0, c1: headCols }, undefined, options)
  const overviewPage = await store.queryOverview({
    groups: [{ groupId: 'head', rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 999 }] }, cols: [{ c0: 0, c1: headCols }] }],
    maxGroups: 2048,
  }, options)
  return {
    readRowsHead: sha256Hex(JSON.stringify(rowsPage)),
    overviewHead: sha256Hex(JSON.stringify(overviewPage)),
  }
}

/** Page window for readRows; callers bound ranges, the adapter bounds pages. */
export const READ_ROWS_PAGE_SIZE = 200
/** Admission guard: a single readRows call resolving more cells refuses loudly. */
export const READ_ROWS_MAX_CELLS = 262144
/** Hard ceiling for overview output (spec §4.1 desktop cap; no profile is carried). */
export const MAX_OVERVIEW_OUTPUT_GROUPS = 2048

function rectContains(rect: SourceRect, row: number, col: number): boolean {
  return row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1
}

function rectsIntersect(a: SourceRect, b: SourceRect): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}

function intersectRect(a: SourceRect, b: SourceRect): SourceRect | null {
  const r0 = Math.max(a.r0, b.r0)
  const r1 = Math.min(a.r1, b.r1)
  const c0 = Math.max(a.c0, b.c0)
  const c1 = Math.min(a.c1, b.c1)
  return r0 <= r1 && c0 <= c1 ? { r0, r1, c0, c1 } : null
}

/** Count of integers r in [lo,hi] with r % mod === rem (mod > 0). */
function countModRange(lo: number, hi: number, mod: number, rem: number): number {
  if (lo > hi) return 0
  const first = lo + (((rem - lo) % mod) + mod) % mod
  return first > hi ? 0 : Math.floor((hi - first) / mod) + 1
}

function countRowSetRows(set: CompactSourceRowSet): number {
  switch (set.kind) {
    case 'intervals': return set.intervals.reduce((sum, iv) => sum + (iv.r1 - iv.r0 + 1), 0)
    case 'list': return set.rows.length
    case 'bitset': return set.count
    case 'ref': return 0
  }
}

function rowSetContains(set: CompactSourceRowSet, row: number): boolean {
  switch (set.kind) {
    case 'intervals': return set.intervals.some((iv) => row >= iv.r0 && row <= iv.r1)
    case 'list': {
      let lo = 0
      let hi = set.rows.length - 1
      while (lo <= hi) {
        const mid = (lo + hi) >> 1
        if (set.rows[mid] === row) return true
        if (set.rows[mid] < row) lo = mid + 1
        else hi = mid - 1
      }
      return false
    }
    case 'bitset': {
      if (row < set.r0 || row > set.r1) return false
      const i = row - set.r0
      return (set.words[i >> 5] & (1 << (i & 31))) !== 0
    }
    case 'ref': return false
  }
}

function* rowSetMembers(set: CompactSourceRowSet): Generator<number> {
  switch (set.kind) {
    case 'intervals':
      for (const iv of set.intervals) for (let row = iv.r0; row <= iv.r1; row++) yield row
      return
    case 'list':
      yield* set.rows
      return
    case 'bitset': {
      const length = set.r1 - set.r0 + 1
      for (let i = 0; i < length; i++) {
        if ((set.words[i >> 5] & (1 << (i & 31))) !== 0) yield set.r0 + i
      }
      return
    }
    case 'ref': throw new Error('row-set refs resolve only against their issuing store')
  }
}

/** Row sets and rect row-ranges overlap without enumerating either side. */
function rowSetOverlapsRange(set: CompactSourceRowSet, r0: number, r1: number): boolean {
  switch (set.kind) {
    case 'intervals': return set.intervals.some((iv) => iv.r0 <= r1 && r0 <= iv.r1)
    case 'list': {
      let lo = 0
      let hi = set.rows.length - 1
      while (lo <= hi) {
        const mid = (lo + hi) >> 1
        if (set.rows[mid] < r0) lo = mid + 1
        else if (set.rows[mid] > r1) hi = mid - 1
        else return true
      }
      return false
    }
    case 'bitset': {
      const start = Math.max(r0, set.r0)
      const end = Math.min(r1, set.r1)
      for (let row = start; row <= end; row++) {
        const i = row - set.r0
        if ((set.words[i >> 5] & (1 << (i & 31))) !== 0) return true
      }
      return false
    }
    case 'ref': return false
  }
}

function toAuthoritative(row: number, col: number, cell: FixtureCell, valuePolicy: ValuePolicy): AuthoritativeCell {
  return {
    id: cellId(row, col), row, col,
    value: cell.value, hasCachedValue: cell.hasCachedValue, valueIsError: cell.valueIsError,
    formula: cell.formula, styleIndex: cell.styleIndex,
    sharedFormula: cell.sharedFormula, arrayRef: cell.arrayRef, dynamicArray: cell.dynamicArray,
    formulaType: cell.formulaType, ca: cell.ca,
    valueProvenance: valuePolicy,
    ...(valuePolicy === 'cached' && cell.formula !== undefined && !cell.hasCachedValue ? { unavailable: true as const } : {}),
  }
}

function toSingleCell(row: number, col: number, cell: FixtureCell, valuePolicy: ValuePolicy): SingleCellPayload {
  const { unavailable: _dropped, ...payload } = toAuthoritative(row, col, cell, valuePolicy)
  return payload
}

function isUnavailableEligible(cell: FixtureCell): boolean {
  return cell.formula !== undefined && !cell.hasCachedValue
}

const F1_CORNERS: Array<{ row: number; col: number; cell: FixtureCell }> = [
  { row: 0, col: 0, cell: { value: 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  { row: 0, col: MAX_COLS - 1, cell: { value: 'XFD1', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  { row: MAX_ROWS - 1, col: 0, cell: { value: 'A1048576', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  { row: MAX_ROWS - 1, col: MAX_COLS - 1, cell: { value: 2 ** 34 - 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER } },
]

function f1ValueFor(id: number): FixtureCell {
  switch (id % 6) {
    case 0: case 1: return { value: id % 10000, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 2: case 3: return { value: `c${id % 997}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
    default: return { value: id % 2 === 0, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  }
}

/** Uniform (row, col) sampling: rows and columns drawn separately so every
 * column is reachable (a single 2^34 draw can only hit multiples of 4). */
function f1SampledIds(seed: number, nnz: number): Set<number> {
  const prng = createPrng(seed)
  const ids = new Set<number>()
  for (const corner of F1_CORNERS) ids.add(cellId(corner.row, corner.col))
  while (ids.size < nnz) {
    const row = Math.floor(prng() * MAX_ROWS)
    const col = Math.floor(prng() * MAX_COLS)
    ids.add(row * MAX_COLS + col)
  }
  return ids
}

const F2_HEADERS = ['id', 'date', 'flag', 'name', 'amount', 'ratio', 'code', 'note']
const F2_POOL = ['alpha', 'beta', 'gamma', 'delta']

function f2TotalsFormula(rowCount: number): string {
  return `SUM(B2:B${rowCount - 1})`
}

function f2BodyCell(seed: number, row: number, col: number): FixtureCell | null {
  switch (col) {
    case 0: return { value: row, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 1: return { value: 45000 + ((row * 7 + (seed % 365)) % 365), hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DATE }
    case 2: return { value: cellHash(seed, row, 2) % 2 === 0, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
    case 3: return { value: F2_POOL[cellHash(seed, row, 3) % F2_POOL.length], hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
    case 4: return row % 7 === 3
      ? null
      : { value: row * 1.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 5: return row % 101 === 0
      ? { value: '#DIV/0!', hasCachedValue: true, valueIsError: true, styleIndex: STYLE_DEFAULT }
      : { value: row + 0.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 6: return { value: row * 2 + 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    default: return { value: `note-${cellHash(seed, row, 7) % 16}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  }
}

function f2BlankCount(rowCount: number): number {
  const first = 1 + (((3 - 1) % 7) + 7) % 7
  const last = rowCount - 2
  return first > last ? 0 : Math.floor((last - first) / 7) + 1
}

function f2BodyStyle(seed: number, row: number, col: number): number {
  void seed
  switch (col) {
    case 0: case 4: case 6: return STYLE_NUMBER
    case 1: return STYLE_DATE
    case 5: return row % 101 === 0 ? STYLE_DEFAULT : STYLE_NUMBER
    default: return STYLE_DEFAULT
  }
}

function f3DenseCell(seed: number, row: number, col: number): FixtureCell {
  const hash = cellHash(seed, row, col)
  if ((row + col) % 8 === 0) return { value: `w${hash % 256}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  return { value: (hash % 100000) / 10, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
}

function f3SparsePopulated(row: number, col: number): boolean {
  const id = row * MAX_COLS + col
  return id % 64 === 7 || col === MAX_COLS - 1
}

/** 1M unique raw UTF-16 strings: twins share a base and differ only by the
 * named variant (case, accent, trailing zero), so value-list dedupe and
 * display-collision handling are genuinely exercised. */
export const F4_STRING_COUNT = 1_000_000
export const F4_LONG_BLOB_INDEX = 999983
export type F4Family = 'natural' | 'accent' | 'case' | 'lone' | 'collide' | 'long' | 'astral' | 'plain' | 'blob'
const F4_FAMILIES: F4Family[] = ['natural', 'accent', 'case', 'lone', 'collide', 'long', 'astral', 'plain']
export function f4Family(index: number): F4Family {
  if (index === F4_LONG_BLOB_INDEX) return 'blob'
  return F4_FAMILIES[index % 8]
}
function f4TwinSub(seed: number, index: number): 0 | 1 {
  return ((Math.floor(index / 8) + (seed & 1)) % 2) as 0 | 1
}
function f4Base(index: number): number {
  return Math.floor(index / 16)
}
export function f4StringFor(seed: number, index: number): string {
  if (index === F4_LONG_BLOB_INDEX) return `blob-${index}-` + 'x'.repeat(100_000)
  const sub = f4TwinSub(seed, index)
  const base = f4Base(index)
  switch (f4Family(index)) {
    case 'natural': return `item-${index}`
    case 'accent': return sub === 0 ? `cafe-${base}` : `café-${base}`
    case 'case': return sub === 0 ? `alpha-${base}` : `Alpha-${base}`
    case 'lone': return `${sub === 0 ? '\uD800' : '\uDE00'}-lone-${index}`
    case 'collide': return sub === 0 ? `price-${base}-1.50` : `price-${base}-1.5`
    case 'long': return `long-${index}-` + 'y'.repeat(200 + (index % 500))
    case 'astral': return `😀-${index}`
    case 'plain': return `row-${index}`
    case 'blob': return `blob-${index}-` + 'x'.repeat(100_000)
  }
}
export function f4StringAt(index: number): string {
  return f4StringFor(FIXTURE_SEEDS.F4, index)
}

export interface F8Case {
  note: string
  address: SourceCell
  cell: FixtureCell
  /** Emit through the shared-string table to exercise the t=s path. */
  shared?: boolean
}

export function f8Cases(): F8Case[] {
  return [
    { note: 'plain-number', address: { row: 0, col: 0 }, cell: { value: 1.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER } },
    { note: 'plain-boolean', address: { row: 0, col: 1 }, cell: { value: true, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
    { note: 'cached-error', address: { row: 0, col: 2 }, cell: { value: '#DIV/0!', hasCachedValue: true, valueIsError: true, styleIndex: STYLE_DEFAULT, formula: '1/0' } },
    { note: 'error-looking-text', address: { row: 0, col: 3 }, cell: { value: '#DIV/0!', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
    { note: 'cached-formula', address: { row: 1, col: 0 }, cell: { value: 3, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'SUM(A1:A2)' } },
    { note: 'cache-less-formula', address: { row: 1, col: 1 }, cell: { value: null, hasCachedValue: false, valueIsError: false, styleIndex: STYLE_DEFAULT, formula: 'WEBSERVICE("https://example.invalid")' } },
    { note: 'unsupported-formula-cache', address: { row: 1, col: 2 }, cell: { value: 42, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'WEBSERVICE("https://example.invalid")' } },
    { note: 'date-style', address: { row: 2, col: 0 }, cell: { value: 44927.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DATE } },
    { note: 'lone-surrogate', address: { row: 2, col: 1 }, cell: { value: '\uDE00-lone', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
    { note: 'shared-string', address: { row: 2, col: 2 }, cell: { value: 'shared-hello', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }, shared: true },
    { note: 'spill-child', address: { row: 3, col: 0 }, cell: { value: 7, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER } },
    { note: 'carriage-return', address: { row: 3, col: 1 }, cell: { value: 'a\rb', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
    { note: 'table-col-name', address: { row: 4, col: 0 }, cell: { value: 't-name', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER } },
    { note: 'table-col-value', address: { row: 4, col: 1 }, cell: { value: 't-value', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER } },
    { note: 'table-col-total', address: { row: 4, col: 2 }, cell: { value: 't-total', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER } },
    { note: 'sparse-edge-cell', address: { row: MAX_ROWS - 1, col: MAX_COLS - 1 }, cell: { value: 'edge', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  ]
}

/** Committed native-compat oracle tests this parity family leans on. F8 cases
 * stay fixed vectors by design (a seed would void their parity meaning). */
export function f8CorpusRefs(): string[] {
  return [
    'tests/xlsx-native-strings.test.ts',
    'tests/xlsx-native-display.test.ts',
    'tests/xlsx-formula-calc-policy.test.ts',
  ]
}

export interface F7StressSpec {
  sortCols: number[]
  hiddenIntervals: Array<{ r0: number; r1: number }>
  tinyRowHeights: number[]
  zoomLevels: number[]
  deviceScales: number[]
  slowJobs: Array<{ kind: 'sort' | 'filter'; rows: number }>
  sheetCount: number
  rowCount: number
  seed: number
}

export function f7Spec(): F7StressSpec {
  return {
    sortCols: [0, 1, 2, 3, 4, 5, 6, 7],
    hiddenIntervals: [
      { r0: 5, r1: 5 },
      { r0: 10, r1: 12 },
      { r0: 100, r1: 199 },
    ],
    tinyRowHeights: [0.5, 1, 2],
    zoomLevels: [0.25, 1, 2, 4],
    deviceScales: [1, 2],
    slowJobs: [
      { kind: 'sort', rows: 100_000 },
      { kind: 'filter', rows: 100_000 },
    ],
    sheetCount: 12,
    rowCount: 100_000,
    seed: FIXTURE_SEEDS.F7,
  }
}

export interface F7SheetHandle {
  readonly sheetId: string
  readonly seed: number
  readonly rowCount: number
  readonly colCount: number
  nnz(): number
  cellAt(row: number, col: number): FixtureCell | null
  canonicalSamples(): string
}

/** Max-dimension sparse sheets: full axes, ~1000 seeded cells each. */
export function f7Sheets(count: number): F7SheetHandle[] {
  const sheets: F7SheetHandle[] = []
  for (let index = 0; index < count; index++) {
    const seed = hashSeed(`xlsx-viewer:F7sheet:${index}`)
    const sampled = f1SampledIds(seed, 1000)
    sheets.push({
      sheetId: `F7-sheet-${index}`,
      seed,
      rowCount: MAX_ROWS,
      colCount: MAX_COLS,
      nnz: () => sampled.size,
      cellAt: (row, col) => {
        if (!isValidCoordinate(row, col)) return null
        const lookup = row * MAX_COLS + col
        return sampled.has(lookup) ? { value: lookup % 1000, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER } : null
      },
      canonicalSamples: () => sha256Hex(JSON.stringify({ sheetId: `F7-sheet-${index}`, seed, nnz: sampled.size })),
    })
  }
  return sheets
}

const F7_POOL_SIZES = [3, 5, 8, 4, 6, 10, 12]

function f5Structure(): FixtureStructure {
  return {
    merges: [
      { r0: 0, r1: 1, c0: 0, c1: 1 },
      { r0: 99, r1: 199999, c0: 1, c1: 100 },
    ],
    tables: [
      { range: { r0: 4, r1: 20, c0: 0, c1: 3 }, headerRows: 1, totalsRows: 1 },
      { range: { r0: 24, r1: 30, c0: 0, c1: 2 }, headerRows: 0, totalsRows: 0 },
    ],
    spills: [
      { anchor: { row: 32, col: 0 }, kind: 'fixed', saved: { r0: 32, r1: 34, c0: 0, c1: 1 } },
      { anchor: { row: 40, col: 3 }, kind: 'dynamic', saved: { r0: 40, r1: 50, c0: 3, c1: 3 }, evaluated: { r0: 40, r1: 50, c0: 3, c1: 3 } },
    ],
    drawings: [
      { kind: 'oneCell', id: 'drawing-1', from: { row: 70, col: 2 }, bounds: { r0: 70, r1: 74, c0: 2, c1: 4 } },
      { kind: 'twoCell', id: 'drawing-2', from: { row: 75, col: 2 }, to: { row: 80, col: 5 }, bounds: { r0: 75, r1: 80, c0: 2, c1: 5 } },
      { kind: 'absolute', id: 'drawing-3' },
    ],
    styleIntervals: [
      { range: { r0: 4, r1: 4, c0: 0, c1: 3 }, styleIndex: STYLE_HEADER },
      { range: { r0: 0, r1: 99, c0: 5, c1: 5 }, styleIndex: STYLE_NUMBER },
    ],
  }
}

function f5CellAt(seed: number, row: number, col: number): FixtureCell | null {
  if ((row === 44 || row === 45 || row === 46) && col === 3) return { value: 7, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
  if (row === 32 && col === 0) return { value: 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'SUM(A33:A35)', arrayRef: 'A33:B35', formulaType: 'array' }
  if (row === 40 && col === 3) return { value: 2, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'SEQUENCE(11)', dynamicArray: true }
  if (row === 4 && col <= 3) return { value: `head-${col}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
  if (row < 10 && col < 10 && (row + col) % 3 === 0) return { value: row * 10 + col + (seed % 5), hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  return null
}

function f5Nnz(seed: number): number {
  let count = 0
  for (let row = 0; row < 100; row++) for (let col = 0; col < 10; col++) {
    if (f5CellAt(seed, row, col) !== null) count++
  }
  return count + 4
}

function f5Enumerate(seed: number, visit: (row: number, col: number, cell: FixtureCell) => void): void {
  for (let row = 0; row < 100; row++) for (let col = 0; col < 10; col++) {
    const cell = f5CellAt(seed, row, col)
    if (cell !== null) visit(row, col, cell)
  }
  for (const [row, col] of [[44, 3], [46, 3], [32, 0], [40, 3]] as Array<[number, number]>) {
    const cell = f5CellAt(seed, row, col)
    if (cell !== null) visit(row, col, cell)
  }
}

const F6_OVERRIDES: Array<{ row: number; col: number; cell: FixtureCell }> = [
  { row: 0, col: 0, cell: { value: 'origin', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  { row: MAX_ROWS - 1, col: MAX_COLS - 1, cell: { value: 'edge', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
]

function f6CellAt(row: number, col: number): FixtureCell {
  for (const override of F6_OVERRIDES) {
    if (override.row === row && override.col === col) return override.cell
  }
  return { value: (row * MAX_COLS + col) % 1000, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
}

function canonicalCellList(cells: CellRef[]): string {
  return JSON.stringify(cells.map(({ row, col, cell }) => [row, col, cell.value, cell.formula ?? null, cell.styleIndex, cell.hasCachedValue, cell.valueIsError]))
}

function lowerBoundId(sorted: number[], target: number): number {
  let lo = 0
  let hi = sorted.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (sorted[mid] < target) lo = mid + 1
    else hi = mid
  }
  return lo
}

export function createFixture(id: FixtureId, params: FixtureParams = {}): FixtureHandle {
  const seed = params.seed ?? FIXTURE_SEEDS[id]
  switch (id) {
    case 'F1': {
      const target = params.nnz ?? 100_000
      const sampled = f1SampledIds(seed, target)
      const sorted = [...sampled].sort((a, b) => a - b)
      const cornerCells = new Map(F1_CORNERS.map((c) => [cellId(c.row, c.col), c.cell]))
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col)) return null
        const lookup = cellId(row, col)
        const corner = cornerCells.get(lookup)
        if (corner !== undefined) return corner
        return sampled.has(lookup) ? f1ValueFor(lookup) : null
      }
      const isPopulated = (row: number, col: number): boolean => {
        if (!isValidCoordinate(row, col)) return false
        const lookup = cellId(row, col)
        return cornerCells.has(lookup) || sampled.has(lookup)
      }
      const styleOf = (row: number, col: number): number => {
        const corner = cornerCells.get(cellId(row, col))
        if (corner !== undefined) return corner.styleIndex
        return f1ValueFor(cellId(row, col)).styleIndex
      }
      const enumerate = (visit: (row: number, col: number, cell: FixtureCell) => void): void => {
        for (const lookup of sorted) {
          const row = Math.floor(lookup / MAX_COLS)
          const corner = cornerCells.get(lookup)
          visit(row, lookup - row * MAX_COLS, corner ?? f1ValueFor(lookup))
        }
      }
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used,
        nnz: () => sampled.size,
        cellAt, isPopulated, styleOf,
        enumeratePopulated: enumerate,
        enumerateHead: (limit) => {
          const out: CellRef[] = []
          for (const lookup of sorted) {
            if (out.length >= limit) break
            const row = Math.floor(lookup / MAX_COLS)
            const col = lookup - row * MAX_COLS
            const corner = cornerCells.get(lookup)
            out.push({ row, col, cell: corner ?? f1ValueFor(lookup) })
          }
          return out
        },
        scanFrom: (range, offset, limit, exclude = []) => {
          const out: CellRef[] = []
          let skipped = 0
          const start = lowerBoundId(sorted, range.r0 * MAX_COLS)
          for (let i = start; i < sorted.length && out.length < limit; i++) {
            const lookup = sorted[i]
            const row = Math.floor(lookup / MAX_COLS)
            if (row > range.r1) break
            const col = lookup - row * MAX_COLS
            if (col < range.c0 || col > range.c1) continue
            if (exclude.some((rect) => rectContains(rect, row, col))) continue
            if (skipped < offset) {
              skipped++
              continue
            }
            const corner = cornerCells.get(lookup)
            out.push({ row, col, cell: corner ?? f1ValueFor(lookup) })
          }
          return out
        },
      })
    }
    case 'F2': {
      const rowCount = params.rowCount ?? MAX_ROWS
      let totals: { date: number; amount: number } | null = null
      const totalsOf = () => {
        if (totals === null) {
          let date = 0
          let amount = 0
          for (let row = 1; row <= rowCount - 2; row++) {
            date += 45000 + ((row * 7 + (seed % 365)) % 365)
            if (row % 7 !== 3) amount += row * 1.5
          }
          totals = { date, amount }
        }
        return totals
      }
      const inGrid = (row: number, col: number): boolean => isValidCoordinate(row, col) && col <= 7 && row < rowCount
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!inGrid(row, col)) return null
        if (row === 0) {
          return { value: F2_HEADERS[col], hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
        }
        if (row === rowCount - 1 && rowCount > 1) {
          if (col === 0) return { value: 'total', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
          if (col === 1) return { value: totalsOf().date, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: f2TotalsFormula(rowCount) }
          if (col === 3) return { value: 'end', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
          if (col === 4) return { value: totalsOf().amount, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
          return null
        }
        return f2BodyCell(seed, row, col)
      }
      const isPopulated = (row: number, col: number): boolean => {
        if (!inGrid(row, col)) return false
        if (row === 0) return true
        if (row === rowCount - 1 && rowCount > 1) return col === 0 || col === 1 || col === 3 || col === 4
        return !(col === 4 && row % 7 === 3)
      }
      const styleOf = (row: number, col: number): number => {
        if (row === 0) return STYLE_HEADER
        if (row === rowCount - 1 && rowCount > 1) {
          if (col === 0) return STYLE_HEADER
          if (col === 1 || col === 4) return STYLE_NUMBER
          return STYLE_DEFAULT
        }
        return f2BodyStyle(seed, row, col)
      }
      const used: SourceRect = { r0: 0, r1: Math.max(rowCount - 1, 0), c0: 0, c1: 7 }
      return makeHandle({
        id, seed, rowCount, colCount: 8, used,
        nnz: () => {
          if (rowCount <= 0) return 0
          if (rowCount === 1) return 8
          return 8 + (rowCount - 2) * 8 - f2BlankCount(rowCount) + 4
        },
        cellAt, isPopulated, styleOf,
      })
    }
    case 'F3': {
      const rowCount = params.rowCount ?? 256
      const colCount = params.colCount ?? MAX_COLS
      const variant = params.variant ?? 'dense'
      const inGrid = (row: number, col: number): boolean => isValidCoordinate(row, col) && row < rowCount && col < colCount
      const populated = (row: number, col: number): boolean => {
        if (!inGrid(row, col)) return false
        if (variant === 'sparse-wide') return f3SparsePopulated(row, col)
        return true
      }
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!populated(row, col)) return null
        return f3DenseCell(seed, row, col)
      }
      const used: SourceRect = { r0: 0, r1: Math.max(rowCount - 1, 0), c0: 0, c1: Math.max(colCount - 1, 0) }
      return makeHandle({
        id, seed, rowCount, colCount, used, variant,
        nnz: () => {
          if (variant === 'dense') return rowCount * colCount
          let count = 0
          for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) {
            if (f3SparsePopulated(row, col)) count++
          }
          return count
        },
        cellAt,
        isPopulated: populated,
        styleOf: (row, col) => ((row + col) % 8 === 0 ? STYLE_DEFAULT : STYLE_NUMBER),
      })
    }
    case 'F4': {
      const rowCount = params.rowCount ?? F4_STRING_COUNT
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col) || col !== 0 || row >= rowCount) return null
        return { value: f4StringFor(seed, row), hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
      }
      const used: SourceRect = { r0: 0, r1: Math.max(rowCount - 1, 0), c0: 0, c1: 0 }
      return makeHandle({
        id, seed, rowCount, colCount: 1, used, nnz: () => rowCount, cellAt,
        isPopulated: (row, col) => isValidCoordinate(row, col) && col === 0 && row < rowCount,
        styleOf: () => STYLE_DEFAULT,
      })
    }
    case 'F5': {
      const used: SourceRect = { r0: 0, r1: 199999, c0: 0, c1: 100 }
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col)) return null
        return f5CellAt(seed, row, col)
      }
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used, nnz: () => f5Nnz(seed), cellAt,
        isPopulated: (row, col) => cellAt(row, col) !== null,
        styleOf: (row, col) => cellAt(row, col)?.styleIndex ?? STYLE_DEFAULT,
        enumeratePopulated: (visit) => {
          f5Enumerate(seed, visit)
        },
        enumerateHead: (limit) => {
          const out: CellRef[] = []
          f5Enumerate(seed, (row, col, cell) => {
            if (out.length < limit) out.push({ row, col, cell })
          })
          return out
        },
        scanFrom: (range, offset, limit, exclude = []) => {
          const out: CellRef[] = []
          let skipped = 0
          f5Enumerate(seed, (row, col, cell) => {
            if (out.length >= limit) return
            if (row < range.r0 || row > range.r1 || col < range.c0 || col > range.c1) return
            if (exclude.some((rect) => rectContains(rect, row, col))) return
            if (skipped < offset) {
              skipped++
              return
            }
            out.push({ row, col, cell })
          })
          return out
        },
      })
    }
    case 'F6': {
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used, nnz: () => 2 ** 34,
        cellAt: (row, col) => {
          if (!isValidCoordinate(row, col)) return null
          return f6CellAt(row, col)
        },
        isPopulated: (row, col) => isValidCoordinate(row, col),
        styleOf: (row, col) => f6CellAt(row, col).styleIndex,
      })
    }
    case 'F7': {
      const spec = f7Spec()
      const rowCount = params.rowCount ?? spec.rowCount
      const inGrid = (row: number, col: number): boolean => {
        if (!isValidCoordinate(row, col) || col > 7) return false
        return row < rowCount || row === MAX_ROWS - 1
      }
      const valueAt = (row: number, col: number): number | string => {
        if (row === MAX_ROWS - 1) return row * 8 + col
        if (col === 7) return `s${(row + (seed % 64)) % 64}`
        return (cellHash(seed, row, col) % F7_POOL_SIZES[col]) + col * 100
      }
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!inGrid(row, col)) return null
        const value = valueAt(row, col)
        return { value, hasCachedValue: true, valueIsError: false, styleIndex: col === 7 ? STYLE_DEFAULT : STYLE_NUMBER }
      }
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: 7 }
      return makeHandle({
        id, seed, rowCount, colCount: 8, used,
        nnz: () => rowCount * 8 + 8,
        cellAt,
        isPopulated: inGrid,
        styleOf: (_row, col) => (col === 7 ? STYLE_DEFAULT : STYLE_NUMBER),
      })
    }
    case 'F8': {
      const cases = f8Cases()
      const byAddress = new Map(cases.map((c) => [`${c.address.row}:${c.address.col}`, c.cell]))
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
      const ordered = [...cases].sort((a, b) => a.address.row - b.address.row || a.address.col - b.address.col)
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used,
        nnz: () => cases.length,
        sampleAddrs: cases.map((c) => ({ row: c.address.row, col: c.address.col })),
        cellAt: (row, col) => byAddress.get(`${row}:${col}`) ?? null,
        isPopulated: (row, col) => byAddress.has(`${row}:${col}`),
        styleOf: (row, col) => byAddress.get(`${row}:${col}`)?.styleIndex ?? STYLE_DEFAULT,
        enumeratePopulated: (visit) => {
          for (const c of ordered) visit(c.address.row, c.address.col, c.cell)
        },
        enumerateHead: (limit) => ordered.slice(0, limit).map((c) => ({ row: c.address.row, col: c.address.col, cell: c.cell })),
        scanFrom: (range, offset, limit, exclude = []) => {
          const out: CellRef[] = []
          let skipped = 0
          for (const c of ordered) {
            if (out.length >= limit) break
            if (!rectContains(range, c.address.row, c.address.col)) continue
            if (exclude.some((rect) => rectContains(rect, c.address.row, c.address.col))) continue
            if (skipped < offset) {
              skipped++
              continue
            }
            out.push({ row: c.address.row, col: c.address.col, cell: c.cell })
          }
          return out
        },
      })
    }
  }
}

function makeHandle(options: {
  id: FixtureId; seed: number; rowCount: number; colCount: number; used: SourceRect
  nnz: () => number; cellAt: (row: number, col: number) => FixtureCell | null
  isPopulated: (row: number, col: number) => boolean; styleOf: (row: number, col: number) => number
  variant?: 'dense' | 'sparse-wide'
  sampleAddrs?: Array<{ row: number; col: number }>
  enumeratePopulated?: (visit: (row: number, col: number, cell: FixtureCell) => void) => void
  scanFrom?: (range: SourceRect, offset: number, limit: number, exclude?: SourceRect[]) => CellRef[]
  enumerateHead?: (limit: number) => CellRef[]
}): FixtureHandle {
  const { id, seed, rowCount, colCount, used } = options
  const selfReader: CellReader = { cellAt: options.cellAt, isPopulated: options.isPopulated, styleOf: options.styleOf }
  const structure = (): FixtureStructure => {
    if (id !== 'F5') {
      return { merges: [], tables: [], spills: [], drawings: [], styleIntervals: [] }
    }
    return f5Structure()
  }
  const summarizeDomain = (
    rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>,
  ): DomainSummary => {
    if (rows.kind === 'ref') return { totalRows: 0, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
    const totalRows = countRowSetRows(rows)
    if (id === 'F6') return summarizeF6(rows, cols, totalRows)
    if (options.enumeratePopulated !== undefined) {
      let populatedRows = 0
      let populatedCells = 0
      let style: number | null | undefined
      const seen = new Set<number>()
      options.enumeratePopulated((row, col, cell) => {
        if (!rowSetContains(rows, row)) return
        if (!cols.some((span) => col >= span.c0 && col <= span.c1)) return
        populatedCells++
        if (style === undefined) style = cell.styleIndex
        else if (style !== cell.styleIndex) style = null
        if (!seen.has(row)) {
          seen.add(row)
          populatedRows++
        }
      })
      return { totalRows, populatedRows, populatedCells, uniformStyle: style === undefined ? STYLE_DEFAULT : style }
    }
    const ctx: SummaryContext = { seed, rowCount, colCount }
    switch (id) {
      case 'F2': return summarizeF2(ctx, rows, cols, totalRows)
      case 'F3': return summarizeF3(ctx, options.variant ?? 'dense', rows, cols, totalRows)
      case 'F4': return summarizeF4(ctx, rows, cols, totalRows)
      case 'F7': return summarizeF7(ctx, rows, cols, totalRows)
      default: return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
    }
  }
  const scanCells = (
    range: SourceRect, offset: number, limit: number, read: CellReader = selfReader, exclude: SourceRect[] = [],
  ): CellRef[] => {
    if (options.scanFrom !== undefined) return options.scanFrom(range, offset, limit, exclude)
    const clamped = intersectRect(range, used)
    if (clamped === null) return []
    const out: CellRef[] = []
    let skipped = 0
    for (let row = clamped.r0; row <= clamped.r1 && out.length < limit; row++) {
      for (let col = clamped.c0; col <= clamped.c1 && out.length < limit; col++) {
        if (exclude.some((rect) => rectContains(rect, row, col))) continue
        const cell = read.cellAt(row, col)
        if (cell === null) continue
        if (skipped < offset) {
          skipped++
          continue
        }
        out.push({ row, col, cell })
      }
    }
    return out
  }
  const sampleAddrs = (): Array<{ row: number; col: number }> => {
    switch (id) {
      case 'F1': return [{ row: 0, col: 0 }, { row: 0, col: MAX_COLS - 1 }, { row: MAX_ROWS - 1, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }]
      case 'F2': {
        const last = Math.max(rowCount - 1, 0)
        return [
          { row: 0, col: 0 }, { row: 0, col: 3 }, { row: 1, col: 1 }, { row: 1, col: 4 },
          { row: 3, col: 4 }, { row: 101, col: 5 }, { row: 10, col: 2 }, { row: 10, col: 6 },
          { row: last, col: 0 }, { row: last, col: 1 }, { row: last, col: 4 },
        ]
      }
      case 'F3': return [{ row: 0, col: 0 }, { row: 0, col: colCount - 1 }, { row: rowCount - 1, col: 0 }, { row: rowCount - 1, col: colCount - 1 }, { row: 0, col: 7 }]
      case 'F4': return [0, 1, 2, 3, 4, 5, 6, 7, F4_LONG_BLOB_INDEX, F4_STRING_COUNT - 1].map((row) => ({ row, col: 0 }))
      case 'F5': return [{ row: 0, col: 0 }, { row: 3, col: 0 }, { row: 4, col: 1 }, { row: 32, col: 0 }, { row: 40, col: 3 }, { row: 45, col: 3 }]
      case 'F6': return [{ row: 0, col: 0 }, { row: 0, col: 1 }, { row: 100, col: 5 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }]
      case 'F7': return [{ row: 0, col: 0 }, { row: 0, col: 7 }, { row: 999, col: 3 }, { row: MAX_ROWS - 1, col: 0 }]
      case 'F8': return (options.sampleAddrs ?? []).map(({ row, col }) => ({ row, col }))
    }
  }
  const canonicalSamples = (): string => {
    const head = sampleAddrs()
      .map(({ row, col }) => ({ row, col, cell: options.cellAt(row, col) }))
      .filter((entry): entry is { row: number; col: number; cell: FixtureCell } => entry.cell !== null)
    const struct = id === 'F5' ? f5Structure() : null
    return JSON.stringify({ id, seed, rowCount, colCount, nnz: memoizedNnz(), head: canonicalCellList(head), struct })
  }
  const enumerateHead = (limit: number): CellRef[] => {
    if (options.enumerateHead !== undefined) return options.enumerateHead(limit)
    throw new Error(`${id} is not enumerable at full scale; bound a range with scanCells`)
  }
  let cachedNnz: number | null = null
  const memoizedNnz = (): number => {
    if (cachedNnz === null) cachedNnz = options.nnz()
    return cachedNnz
  }
  return {
    id, seed, rowCount, colCount, used, nnz: memoizedNnz,
    cellAt: options.cellAt, isPopulated: options.isPopulated, styleOf: options.styleOf,
    structure, enumerateHead, summarizeDomain, scanCells, canonicalSamples,
    createStore: (storeOptions) => createSyntheticStore(
      { id, used, cellAt: options.cellAt, isPopulated: options.isPopulated, styleOf: options.styleOf, structure, summarizeDomain, scanCells }, storeOptions,
    ),
  }
}

interface SummaryContext {
  seed: number
  rowCount: number
  colCount: number
}

function clampColSpans(cols: ReadonlyArray<{ c0: number; c1: number }>, lo: number, hi: number): Array<{ c0: number; c1: number }> {
  const out: Array<{ c0: number; c1: number }> = []
  for (const span of cols) {
    const c0 = Math.max(span.c0, lo)
    const c1 = Math.min(span.c1, hi)
    if (c0 <= c1) out.push({ c0, c1 })
  }
  return out
}

function spanWidth(spans: ReadonlyArray<{ c0: number; c1: number }>): number {
  return spans.reduce((sum, span) => sum + (span.c1 - span.c0 + 1), 0)
}

function clampIntervals(intervals: ReadonlyArray<{ r0: number; r1: number }>, lo: number, hi: number): Array<{ r0: number; r1: number }> {
  const out: Array<{ r0: number; r1: number }> = []
  for (const iv of intervals) {
    const r0 = Math.max(iv.r0, lo)
    const r1 = Math.min(iv.r1, hi)
    if (r0 <= r1) out.push({ r0, r1 })
  }
  return out
}

/** Caller-bounded members of a list/bitset row set inside [lo, hi], ascending. */
function filterMembers(rows: CompactSourceRowSet, lo: number, hi: number): number[] {
  const out: number[] = []
  for (const row of rowSetMembers(rows)) {
    if (row >= lo && row <= hi) out.push(row)
  }
  return out
}

function existsModIntervals(intervals: ReadonlyArray<{ r0: number; r1: number }>, mod: number, rem: number): boolean {
  return intervals.some((iv) => countModRange(iv.r0, iv.r1, mod, rem) > 0)
}

function existsModMembers(members: ReadonlyArray<number>, mod: number, rem: number): boolean {
  return members.some((row) => ((row % mod) + mod) % mod === rem)
}

function countModMembers(members: ReadonlyArray<number>, mod: number, rem: number): number {
  let count = 0
  for (const row of members) {
    if (((row % mod) + mod) % mod === rem) count++
  }
  return count
}

function summarizeF6(rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>, totalRows: number): DomainSummary {
  const width = spanWidth(clampColSpans(cols, 0, MAX_COLS - 1))
  const populatedCells = totalRows * width
  if (populatedCells === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  const hits = F6_OVERRIDES.filter((o) => rowSetContains(rows, o.row) && cols.some((span) => o.col >= span.c0 && o.col <= span.c1))
  const styles = new Set<number>()
  if (populatedCells > hits.length) styles.add(STYLE_NUMBER)
  for (const hit of hits) styles.add(hit.cell.styleIndex)
  return {
    totalRows, populatedRows: totalRows, populatedCells,
    uniformStyle: styles.size === 1 ? [...styles][0] : null,
  }
}

function summarizeF2(ctx: SummaryContext, rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>, totalRows: number): DomainSummary {
  const spans = clampColSpans(cols, 0, 7)
  const width = spanWidth(spans)
  const n = ctx.rowCount
  if (width === 0 || n === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  const inBody = (row: number): boolean => row >= 1 && row <= n - 2
  let header = false
  let totals = false
  let bodyIntervals: Array<{ r0: number; r1: number }> = []
  let bodyMembers: number[] = []
  let useMembers = false
  if (rows.kind === 'intervals') {
    header = rows.intervals.some((iv) => iv.r0 <= 0 && 0 <= iv.r1)
    totals = n >= 2 && rows.intervals.some((iv) => iv.r0 <= n - 1 && n - 1 <= iv.r1)
    bodyIntervals = clampIntervals(rows.intervals, 1, n - 2)
  } else {
    useMembers = true
    const members = filterMembers(rows, 0, n - 1)
    header = members.includes(0)
    totals = n >= 2 && members.includes(n - 1)
    bodyMembers = members.filter(inBody)
  }
  const bodyRows = useMembers
    ? bodyMembers.length
    : bodyIntervals.reduce((sum, iv) => sum + (iv.r1 - iv.r0 + 1), 0)
  const blanksInBody = useMembers
    ? countModMembers(bodyMembers, 7, 3)
    : bodyIntervals.reduce((sum, iv) => sum + countModRange(iv.r0, iv.r1, 7, 3), 0)
  let bodyCells = 0
  for (const span of spans) {
    for (let col = span.c0; col <= span.c1; col++) {
      bodyCells += col === 4 ? bodyRows - blanksInBody : bodyRows
    }
  }
  const totalsWidth = spans.reduce((sum, span) => {
    const c0 = Math.max(span.c0, 0)
    const c1 = Math.min(span.c1, 7)
    let count = 0
    for (let col = c0; col <= c1; col++) {
      if (col === 0 || col === 1 || col === 3 || col === 4) count++
    }
    return sum + count
  }, 0)
  const populatedCells = (header ? width : 0) + bodyCells + (totals ? totalsWidth : 0)
  const populatedRows = (header ? 1 : 0) + bodyRows + (totals ? 1 : 0)
  if (populatedCells === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  const styles = new Set<number>()
  if (header) styles.add(STYLE_HEADER)
  if (totals) {
    for (const span of spans) {
      for (let col = Math.max(span.c0, 0); col <= Math.min(span.c1, 7); col++) {
        if (col === 0) styles.add(STYLE_HEADER)
        else if (col === 1 || col === 4) styles.add(STYLE_NUMBER)
        else if (col === 3) styles.add(STYLE_DEFAULT)
      }
    }
  }
  if (bodyRows > 0) {
    const bodyStyle = (col: number): number[] => {
      if (col === 0 || col === 4 || col === 6) return [STYLE_NUMBER]
      if (col === 1) return [STYLE_DATE]
      if (col === 5) {
        const err = useMembers ? existsModMembers(bodyMembers, 101, 0) : existsModIntervals(bodyIntervals, 101, 0)
        const plain = useMembers ? bodyMembers.length - countModMembers(bodyMembers, 101, 0)
          : bodyRows - bodyIntervals.reduce((sum, iv) => sum + countModRange(iv.r0, iv.r1, 101, 0), 0)
        return [...(err ? [STYLE_DEFAULT] : []), ...(plain > 0 ? [STYLE_NUMBER] : [])]
      }
      return [STYLE_DEFAULT]
    }
    for (const span of spans) {
      for (let col = Math.max(span.c0, 0); col <= Math.min(span.c1, 7); col++) {
        for (const style of bodyStyle(col)) styles.add(style)
      }
    }
  }
  return { totalRows, populatedRows, populatedCells, uniformStyle: styles.size === 1 ? [...styles][0] : null }
}

function summarizeF3(ctx: SummaryContext, variant: 'dense' | 'sparse-wide', rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>, totalRows: number): DomainSummary {
  const spans = clampColSpans(cols, 0, ctx.colCount - 1)
  const width = spanWidth(spans)
  const rowCount = ((): number => {
    if (rows.kind === 'intervals') return clampIntervals(rows.intervals, 0, ctx.rowCount - 1).reduce((sum, iv) => sum + (iv.r1 - iv.r0 + 1), 0)
    return filterMembers(rows, 0, ctx.rowCount - 1).length
  })()
  if (width === 0 || rowCount === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  if (variant === 'dense') {
    const populatedCells = rowCount * width
    return { totalRows, populatedRows: rowCount, populatedCells, uniformStyle: denseResidueStyle(rows, spans, populatedCells, ctx) }
  }
  let latticeCols = 0
  for (const span of spans) latticeCols += countModRange(span.c0, span.c1, 64, 7)
  const lastCol = spans.some((span) => MAX_COLS - 1 >= span.c0 && MAX_COLS - 1 <= span.c1) && MAX_COLS - 1 < ctx.colCount ? 1 : 0
  const perRow = latticeCols + lastCol
  const populatedCells = rowCount * perRow
  if (populatedCells === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  return { totalRows, populatedRows: rowCount, populatedCells, uniformStyle: sparseResidueStyle(rows, ctx.rowCount, populatedCells) }
}

/** Exact style uniformity for residue-ruled dense grids (style 0 iff (r+c)%8===0). */
function denseResidueStyle(
  rows: CompactSourceRowSet, spans: ReadonlyArray<{ c0: number; c1: number }>,
  populatedCells: number, ctx: SummaryContext,
): number | null {
  if (populatedCells === 1) {
    const single = singlePopulatedAddr(rows, spans, ctx)
    return single === null ? null : (single.row + single.col) % 8 === 0 ? STYLE_DEFAULT : STYLE_NUMBER
  }
  const width = spanWidth(spans)
  if (width >= 8) return null
  let zero = false
  let two = false
  const visitRow = (row: number): boolean => {
    for (const span of spans) {
      const first = (row + span.c0) % 8
      if (first === 0 || first + (span.c1 - span.c0) >= 8) zero = true
      else two = true
      if (span.c1 > span.c0) two = true
      if (zero && two) return true
    }
    return false
  }
  if (rows.kind === 'intervals') {
    for (const iv of clampIntervals(rows.intervals, 0, ctx.rowCount - 1)) {
      for (let row = iv.r0; row <= iv.r1; row++) {
        if (visitRow(row)) return null
      }
    }
  } else {
    for (const row of filterMembers(rows, 0, ctx.rowCount - 1)) {
      if (visitRow(row)) return null
    }
  }
  if (!zero && !two) return STYLE_DEFAULT
  return zero && !two ? STYLE_DEFAULT : !zero && two ? STYLE_NUMBER : null
}

/** Sparse lattice columns are all ≡7 (mod 64), hence ≡7 (mod 8): style varies by row only. */
function sparseResidueStyle(rows: CompactSourceRowSet, fixtureRowCount: number, populatedCells: number): number | null {
  if (populatedCells === 1) {
    const only = rows.kind === 'intervals'
      ? clampIntervals(rows.intervals, 0, fixtureRowCount - 1)[0]?.r0
      : filterMembers(rows, 0, fixtureRowCount - 1)[0]
    if (only === undefined) return null
    return (only + 7) % 8 === 0 ? STYLE_DEFAULT : STYLE_NUMBER
  }
  const residues = new Set<number>()
  const visit = (row: number): boolean => {
    residues.add(((row + 7) % 8 + 8) % 8)
    return residues.size > 1
  }
  if (rows.kind === 'intervals') {
    const clamped = clampIntervals(rows.intervals, 0, fixtureRowCount - 1)
    if (clamped.some((iv) => iv.r1 - iv.r0 >= 7)) return null
    for (const iv of clamped) {
      for (let row = iv.r0; row <= iv.r1; row++) {
        if (visit(row)) return null
      }
    }
  } else {
    for (const row of filterMembers(rows, 0, fixtureRowCount - 1)) {
      if (visit(row)) return null
    }
  }
  if (residues.size === 0) return STYLE_DEFAULT
  return residues.has(0) && residues.size === 1 ? STYLE_DEFAULT : residues.size === 1 ? STYLE_NUMBER : null
}

function singlePopulatedAddr(
  rows: CompactSourceRowSet, spans: ReadonlyArray<{ c0: number; c1: number }>, ctx: SummaryContext,
): { row: number; col: number } | null {
  const col = spans.length === 1 && spans[0].c0 === spans[0].c1 ? spans[0].c0 : null
  if (col === null || col < 0 || col >= ctx.colCount) return null
  if (rows.kind === 'intervals') {
    const clamped = clampIntervals(rows.intervals, 0, ctx.rowCount - 1)
    if (clamped.length === 1 && clamped[0].r0 === clamped[0].r1) return { row: clamped[0].r0, col }
    return null
  }
  const members = filterMembers(rows, 0, ctx.rowCount - 1)
  if (members.length === 1) return { row: members[0], col }
  return null
}

function summarizeF4(ctx: SummaryContext, rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>, totalRows: number): DomainSummary {
  const hasCol = cols.some((span) => span.c0 <= 0 && 0 <= span.c1)
  if (!hasCol) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  const rowCount = rows.kind === 'intervals'
    ? clampIntervals(rows.intervals, 0, ctx.rowCount - 1).reduce((sum, iv) => sum + (iv.r1 - iv.r0 + 1), 0)
    : filterMembers(rows, 0, ctx.rowCount - 1).length
  if (rowCount === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  return { totalRows, populatedRows: rowCount, populatedCells: rowCount, uniformStyle: STYLE_DEFAULT }
}

function summarizeF7(ctx: SummaryContext, rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>, totalRows: number): DomainSummary {
  const spans = clampColSpans(cols, 0, 7)
  const width = spanWidth(spans)
  const inUsed = (row: number): boolean => (row >= 0 && row < ctx.rowCount) || row === MAX_ROWS - 1
  const rowCount = rows.kind === 'intervals'
    ? clampIntervals(rows.intervals, 0, MAX_ROWS - 1).filter((iv) => iv.r0 < ctx.rowCount || iv.r1 >= MAX_ROWS - 1)
      .reduce((sum, iv) => {
        let count = 0
        if (iv.r0 < ctx.rowCount) count += Math.min(iv.r1, ctx.rowCount - 1) - iv.r0 + 1
        if (iv.r1 >= MAX_ROWS - 1 && iv.r0 <= MAX_ROWS - 1) count += 1
        return sum + count
      }, 0)
    : filterMembers(rows, 0, MAX_ROWS - 1).filter(inUsed).length
  const populatedCells = rowCount * width
  if (populatedCells === 0) return { totalRows, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
  const hasStrings = spans.some((span) => span.c0 <= 7 && 7 <= span.c1)
  const hasNumbers = spans.some((span) => span.c0 <= 6)
  const uniform = hasStrings && hasNumbers ? null : hasStrings ? STYLE_DEFAULT : hasNumbers ? STYLE_NUMBER : STYLE_DEFAULT
  return { totalRows, populatedRows: width > 0 ? rowCount : 0, populatedCells, uniformStyle: uniform }
}

function validateDomain(rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>): { code: 'invalid-coordinate' | 'invalid-cursor'; message: string } | null {
  for (const span of cols) {
    if (!Number.isSafeInteger(span.c0) || !Number.isSafeInteger(span.c1) || span.c0 > span.c1 || span.c0 < 0 || span.c1 >= MAX_COLS) {
      return { code: 'invalid-coordinate', message: 'overview column span is outside the worksheet domain' }
    }
  }
  switch (rows.kind) {
    case 'intervals':
      for (const iv of rows.intervals) {
        if (!Number.isSafeInteger(iv.r0) || !Number.isSafeInteger(iv.r1) || iv.r0 > iv.r1 || iv.r0 < 0 || iv.r1 >= MAX_ROWS) {
          return { code: 'invalid-coordinate', message: 'overview row interval is outside the worksheet domain' }
        }
      }
      return null
    case 'list': {
      if (!Number.isSafeInteger(rows.r0) || !Number.isSafeInteger(rows.r1) || rows.r0 > rows.r1 || rows.r0 < 0 || rows.r1 >= MAX_ROWS) {
        return { code: 'invalid-coordinate', message: 'overview row-list band is outside the worksheet domain' }
      }
      let previous = -1
      for (const row of rows.rows) {
        if (!Number.isSafeInteger(row) || row < rows.r0 || row > rows.r1 || row <= previous) {
          return { code: 'invalid-coordinate', message: 'overview row lists hold ascending in-band source rows' }
        }
        previous = row
      }
      return null
    }
    case 'bitset': {
      if (!Number.isSafeInteger(rows.r0) || !Number.isSafeInteger(rows.r1) || rows.r0 > rows.r1 || rows.r0 < 0 || rows.r1 >= MAX_ROWS) {
        return { code: 'invalid-coordinate', message: 'overview bitset band is outside the worksheet domain' }
      }
      if (rows.words.length !== Math.ceil((rows.r1 - rows.r0 + 1) / 32)) {
        return { code: 'invalid-coordinate', message: 'overview bitset words must cover exactly ceil((r1-r0+1)/32) words' }
      }
      return null
    }
    case 'ref': return { code: 'invalid-cursor', message: 'row-set refs resolve only against their issuing store' }
  }
}

function cursorTag(queryKey: string): string {
  return hashSeed(queryKey).toString(36)
}

function makeCursor(prefix: string, offset: number, queryKey: string): string {
  return `${prefix}:${offset}:${cursorTag(queryKey)}`
}

function parseCursor(cursor: string | undefined, prefix: string, queryKey: string): number | null {
  if (cursor === undefined) return 0
  const match = new RegExp(`^${prefix}:(\\d+):([0-9a-z]+)$`).exec(cursor)
  if (match === null) return null
  if (match[2] !== cursorTag(queryKey)) return null
  const offset = Number(match[1])
  return Number.isSafeInteger(offset) ? offset : null
}

interface StoreBackend {
  id: FixtureId
  used: SourceRect
  cellAt: (row: number, col: number) => FixtureCell | null
  isPopulated: (row: number, col: number) => boolean
  styleOf: (row: number, col: number) => number
  structure: () => FixtureStructure
  summarizeDomain: (rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>) => DomainSummary
  scanCells: (range: SourceRect, offset: number, limit: number, read?: CellReader, exclude?: SourceRect[]) => CellRef[]
}

function stalePage<T>(stamp: SourceStamp, code: 'stale-source'): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error: { code, retryable: false, message: 'source stamp changed or store released' } }
}

function badCursorPage<T>(stamp: SourceStamp): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-cursor', retryable: false, message: 'cursor belongs to another query or stamp' } }
}

function badLimitPage<T>(stamp: SourceStamp): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-cursor', retryable: false, message: 'column limit must be a positive integer' } }
}

/** Nominal EMU geometry for fixture drawings (documented placeholder scale:
 * 64px columns, 20px rows at 9525 EMU/px). S3 resolves real axis geometry. */
const EMU_PER_PX = 9525
function nominalExtent(bounds: SourceRect): { x: number; y: number; width: number; height: number } {
  return { x: 0, y: 0, width: (bounds.c1 - bounds.c0 + 1) * 64 * EMU_PER_PX, height: (bounds.r1 - bounds.r0 + 1) * 20 * EMU_PER_PX }
}

function createSyntheticStore(backend: StoreBackend, storeOptions: SyntheticStoreOptions = {}): SourceStore {
  const valuePolicy = storeOptions.valuePolicy ?? 'engine'
  const pending = storeOptions.pending ?? []
  const errors = storeOptions.errors ?? []
  const stamp: SourceStamp = { docId: `viewer-${backend.id}`, sheetId: `${backend.id}-sheet-1`, valueGen: 0 }
  let released = false
  const gone = (candidate: SourceStamp): boolean => released
    || candidate.docId !== stamp.docId || candidate.sheetId !== stamp.sheetId || candidate.valueGen !== stamp.valueGen
  const read: CellReader = {
    cellAt: storeOptions.read?.cellAt ?? backend.cellAt,
    isPopulated: storeOptions.read?.isPopulated ?? backend.isPopulated,
    styleOf: storeOptions.read?.styleOf ?? backend.styleOf,
  }

  const pendingAt = (row: number, col: number): number => pending.findIndex((rect) => rectContains(rect, row, col))
  const errorAt = (row: number, col: number) => errors.find((entry) => rectContains(entry.range, row, col))

  const toRead = (row: number, col: number): CellRead => {
    if (!isValidCoordinate(row, col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const failure = errorAt(row, col)
    if (failure !== undefined) return { kind: 'error', code: failure.code, retryable: failure.retryable }
    const page = pendingAt(row, col)
    if (page >= 0) return { kind: 'unknown', pageId: `pending-${page}` }
    if (!read.isPopulated(row, col)) return { kind: 'blank' }
    const cell = read.cellAt(row, col)
    if (cell === null) return { kind: 'blank' }
    if (valuePolicy === 'cached' && isUnavailableEligible(cell)) {
      return { kind: 'unavailable', cell: toSingleCell(row, col, cell, valuePolicy), reason: 'formula-value-unavailable' }
    }
    return { kind: 'cell', cell: toSingleCell(row, col, cell, valuePolicy) }
  }

  const entryFor = (row: number, col: number): ColumnEntry => {
    const resolved = toRead(row, col)
    switch (resolved.kind) {
      case 'cell': return { kind: 'cell', row, cell: resolved.cell }
      case 'blank': return { kind: 'blank', r0: row, r1: row }
      case 'unknown': return { kind: 'unknown', r0: row, r1: row, pageId: resolved.pageId }
      case 'error': return { kind: 'error', r0: row, r1: row, code: resolved.code, retryable: resolved.retryable }
      case 'unavailable': return { kind: 'unavailable', row, cell: resolved.cell, reason: resolved.reason }
    }
  }

  const rangeOverlaps = (rects: SourceRect[], r0: number, r1: number, c0: number, c1: number): boolean => rects.some(
    (rect) => rect.r0 <= r1 && r0 <= rect.r1 && rect.c0 <= c1 && c0 <= rect.c1,
  )
  const firstOverlapError = (r0: number, r1: number, c0: number, c1: number) => errors.find(
    (entry) => entry.range.r0 <= r1 && r0 <= entry.range.r1 && entry.range.c0 <= c1 && c0 <= entry.range.c1,
  )

  return {
    stamp,
    valuePolicy,
    extent: { rows: 1048576, cols: 16384, used: { ...backend.used } },
    readCell: async (at, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return { kind: 'error', code: 'stale-source', retryable: false }
      return toRead(at.row, at.col)
    },
    readRows: async (range, cursor, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return stalePage(stamp, 'stale-source')
      if (!isValidSourceRect(range)) {
        return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-source-rect', retryable: false, message: 'range is outside the worksheet domain' } }
      }
      const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1)
      if (area > READ_ROWS_MAX_CELLS) {
        return { items: [], completeness: 'incomplete', stamp, error: { code: 'resource-limit', retryable: false, message: `range resolves ${area} cells, bound is ${READ_ROWS_MAX_CELLS}` } }
      }
      const queryKey = `rr:${range.r0},${range.r1},${range.c0},${range.c1}`
      const start = parseCursor(cursor, 'row', queryKey)
      if (start === null) return badCursorPage(stamp)
      const first = Math.max(range.r0, start)
      if (first > range.r1) return { items: [], completeness: 'complete', stamp }
      const end = Math.min(range.r1, first + READ_ROWS_PAGE_SIZE - 1)
      const items: SparseRowPage[] = []
      let completeness: Completeness = 'complete'
      let failure: { code: SourceErrorCode; retryable: boolean } | null = null
      for (let row: SourceRow = first; row <= end; row++) {
        const rowError = errors.find((entry) => entry.range.r0 <= row && row <= entry.range.r1 && entry.range.c0 <= range.c1 && range.c0 <= entry.range.c1)
        if (rowError !== undefined) {
          completeness = 'incomplete'
          failure = rowError
          continue
        }
        const rowPending = pending.some((rect) => rect.r0 <= row && row <= rect.r1 && rect.c0 <= range.c1 && range.c0 <= rect.c1)
        if (rowPending) {
          if (completeness === 'complete') completeness = 'loading'
          continue
        }
        const cells: AuthoritativeCell[] = []
        for (let col = range.c0; col <= range.c1; col++) {
          if (!read.isPopulated(row, col)) continue
          const cell = read.cellAt(row, col)
          if (cell === null) continue
          cells.push(toAuthoritative(row, col, cell, valuePolicy))
        }
        items.push({ row, cells })
      }
      return {
        items, completeness, stamp,
        error: failure === null ? undefined : { code: failure.code, retryable: failure.retryable, message: 'row range overlaps a failed region' },
        nextCursor: end < range.r1 ? makeCursor('row', end + 1, queryKey) : undefined,
      }
    },
    queryColumn: async (request, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return stalePage(stamp, 'stale-source')
      if (!isValidCoordinate(0, request.col)) {
        return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-coordinate', retryable: false, message: 'column is outside the worksheet domain' } }
      }
      if (request.rows.kind === 'ref') return badCursorPage(stamp)
      if (!Number.isSafeInteger(request.limit) || request.limit <= 0) return badLimitPage(stamp)
      const queryKey = `qc:${request.col}|${request.limit}|${JSON.stringify(request.rows)}`
      const offset = parseCursor(request.cursor, 'cell', queryKey)
      if (offset === null) return badCursorPage(stamp)
      const items: ColumnEntry[] = []
      let skipped = 0
      let consumed = 0
      let finished = true
      let blankStart: number | null = null
      let previousRow: number | null = null
      let loading = false
      let failure: { code: SourceErrorCode; retryable: boolean } | null = null
      for (const row of rowSetMembers(request.rows)) {
        if (skipped < offset) {
          skipped++
          continue
        }
        if (items.length >= request.limit) {
          finished = false
          break
        }
        consumed++
        const entry = entryFor(row, request.col)
        if (entry.kind === 'blank') {
          if (blankStart !== null && previousRow !== null && row === previousRow + 1) {
            items[items.length - 1] = { kind: 'blank', r0: blankStart, r1: row }
          } else {
            blankStart = row
            items.push({ kind: 'blank', r0: row, r1: row })
          }
        } else {
          blankStart = null
          items.push(entry)
          if (entry.kind === 'unknown') loading = true
          if (entry.kind === 'error' && failure === null) failure = entry
        }
        previousRow = row
      }
      return {
        items, stamp,
        completeness: failure !== null ? 'incomplete' : loading ? 'loading' : 'complete',
        error: failure === null ? undefined : { code: failure.code, retryable: failure.retryable, message: 'column domain overlaps a failed region' },
        nextCursor: finished ? undefined : makeCursor('cell', offset + consumed, queryKey),
      }
    },
    scanPopulated: async function* (request, options) {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) {
        yield stalePage(stamp, 'stale-source')
        return
      }
      if (!isValidSourceRect(request.range)) {
        yield { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-source-rect', retryable: false, message: 'range is outside the worksheet domain' } }
        return
      }
      if (!Number.isSafeInteger(request.limit) || request.limit <= 0) {
        yield badLimitPage(stamp)
        return
      }
      const queryKey = `sc:${request.range.r0},${request.range.r1},${request.range.c0},${request.range.c1}|${request.limit}`
      const offset = parseCursor(request.cursor, 'cell', queryKey)
      if (offset === null) {
        yield badCursorPage(stamp)
        return
      }
      const excluded = [...errors.map((entry) => entry.range), ...pending]
      const found = backend.scanCells(request.range, offset, request.limit, read, excluded)
      const failure = firstOverlapError(request.range.r0, request.range.r1, request.range.c0, request.range.c1)
      const waiting = rangeOverlaps(pending, request.range.r0, request.range.r1, request.range.c0, request.range.c1)
      yield {
        items: found.map(({ row, col, cell }) => toAuthoritative(row, col, cell, valuePolicy)),
        completeness: failure !== undefined ? 'incomplete' : waiting ? 'loading' : 'complete',
        stamp,
        error: failure === undefined ? undefined : { code: failure.code, retryable: failure.retryable, message: 'scan range overlaps a failed region' },
        nextCursor: found.length === request.limit ? makeCursor('cell', offset + request.limit, queryKey) : undefined,
      }
    },
    queryStructural: async (rect, options) => {
      options.signal.throwIfAborted()
      const structure = backend.structure()
      const indexedMerges = structure.merges.map((range, index) => ({ id: `merge-${index}`, range }))
      const indexedTables = structure.tables.map((table, index) => ({ id: `table-${index}`, ...table }))
      const empty: StructuralResult = { stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [] }
      if (gone(options.stamp)) return { ...empty, completeness: 'incomplete', error: { code: 'stale-source', retryable: false, message: 'source stamp changed or store released' } }
      if (!isValidSourceRect(rect)) {
        return { ...empty, completeness: 'incomplete', error: { code: 'invalid-source-rect', retryable: false, message: 'rect is outside the worksheet domain' } }
      }
      const failure = firstOverlapError(rect.r0, rect.r1, rect.c0, rect.c1)
      const missing = pending.filter((region) => rectsIntersect(region, rect))
      return {
        stamp,
        completeness: failure !== undefined ? 'incomplete' : missing.length > 0 ? 'loading' : 'complete',
        merges: indexedMerges.filter((merge) => rectsIntersect(merge.range, rect))
          .map((merge) => ({ id: merge.id, range: merge.range, anchor: { row: merge.range.r0, col: merge.range.c0 } })),
        tables: indexedTables.filter((table) => rectsIntersect(table.range, rect))
          .map((table) => ({ id: table.id, range: table.range, headerRows: table.headerRows, totalsRows: table.totalsRows })),
        spills: structure.spills.filter((spill) => {
          const bounds: SourceRect[] = [spill.saved, spill.evaluated].filter((extent): extent is SourceRect => extent !== undefined)
          return bounds.some((extent) => rectsIntersect(extent, rect))
            || (spill.anchor.row >= rect.r0 && spill.anchor.row <= rect.r1 && spill.anchor.col >= rect.c0 && spill.anchor.col <= rect.c1)
        }),
        // Absolute anchors are viewport-fixed with no source bounds; rect
        // queries report source-indexed records only (placement is an S5 concern).
        drawings: structure.drawings.filter((drawing): drawing is Extract<FixtureStructure['drawings'][number], { kind: 'oneCell' | 'twoCell' }> =>
          drawing.kind !== 'absolute' && rectsIntersect(drawing.bounds, rect)).map((drawing) => {
          const from = { ...drawing.from, offsetXEmu: 0, offsetYEmu: 0 }
          if (drawing.kind === 'oneCell') {
            return { id: drawing.id, kind: 'oneCell', from, extent: nominalExtent(drawing.bounds), bounds: drawing.bounds }
          }
          const to = { ...drawing.to, offsetXEmu: 0, offsetYEmu: 0 }
          return { id: drawing.id, kind: 'twoCell', from, to, extent: nominalExtent(drawing.bounds), bounds: drawing.bounds }
        }),
        missing,
        error: failure === undefined ? undefined : { code: failure.code, retryable: failure.retryable, message: 'structural rect overlaps a failed region' },
      }
    },
    queryOverview: async (request, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return stalePage(stamp, 'stale-source')
      if (!Number.isSafeInteger(request.maxGroups) || request.maxGroups < 1 || request.maxGroups > MAX_OVERVIEW_OUTPUT_GROUPS) {
        return { items: [], completeness: 'incomplete', stamp, error: { code: 'resource-limit', retryable: false, message: `maxGroups must be within 1..${MAX_OVERVIEW_OUTPUT_GROUPS}` } }
      }
      for (const domain of request.groups) {
        const invalid = validateDomain(domain.rows, domain.cols)
        if (invalid !== null) {
          return { items: [], completeness: 'incomplete', stamp, error: { code: invalid.code, retryable: false, message: invalid.message } }
        }
      }
      const queryKey = `ov:${request.groups.map((group) => group.groupId).join('|')}|${request.maxGroups}`
      const start = parseCursor(request.cursor, 'group', queryKey)
      if (start === null) return badCursorPage(stamp)
      const items: OverviewGroup[] = []
      const end = Math.min(request.groups.length, start + request.maxGroups)
      for (let index = start; index < end; index++) {
        const domain = request.groups[index]
        const failure = domain.cols.map((span) => firstOverlapError(
          domain.rows.kind === 'intervals' && domain.rows.intervals.length > 0 ? domain.rows.intervals[0].r0 : 0,
          domain.rows.kind === 'intervals' && domain.rows.intervals.length > 0 ? domain.rows.intervals[domain.rows.intervals.length - 1].r1 : MAX_ROWS - 1,
          span.c0, span.c1,
        )).find((entry) => entry !== undefined)
        void failure
        const errorOverlap = errors.some((entry) => rowSetOverlapsRange(domain.rows, entry.range.r0, entry.range.r1)
          && domain.cols.some((span) => span.c0 <= entry.range.c1 && entry.range.c0 <= span.c1))
        if (errorOverlap) {
          const first = errors.find((entry) => rowSetOverlapsRange(domain.rows, entry.range.r0, entry.range.r1)
            && domain.cols.some((span) => span.c0 <= entry.range.c1 && entry.range.c0 <= span.c1))!
          return { items: [], completeness: 'incomplete', stamp, error: { code: first.code, retryable: first.retryable, message: 'overview domain overlaps a failed region' } }
        }
        const overlapsPending = pending.some((rect) => rowSetOverlapsRange(domain.rows, rect.r0, rect.r1)
          && domain.cols.some((span) => span.c0 <= rect.c1 && rect.c0 <= span.c1))
        if (overlapsPending) {
          items.push({ groupId: domain.groupId, content: 'unknown', pageId: `pending-overview-${index}` })
          continue
        }
        const summary = backend.summarizeDomain(domain.rows, domain.cols)
        const width = spanWidth(domain.cols)
        const totalCells = summary.totalRows * width
        if (summary.populatedCells === 0) {
          items.push({
            groupId: domain.groupId, populatedCount: 0, countKind: 'exact',
            style: { kind: 'uniform', styleIndex: STYLE_DEFAULT }, content: 'blank',
          })
          continue
        }
        const mixed = summary.populatedCells < totalCells
        items.push({
          groupId: domain.groupId, populatedCount: summary.populatedCells, countKind: 'exact',
          style: summary.uniformStyle === null ? { kind: 'mixed' } : { kind: 'uniform', styleIndex: summary.uniformStyle },
          content: mixed ? 'mixed' : 'populated',
        })
      }
      return { items, completeness: 'complete', stamp, nextCursor: end < request.groups.length ? makeCursor('group', end, queryKey) : undefined }
    },
    ensureComplete: async (rect, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) {
        return { range: rect, completeness: 'incomplete', stamp, missing: [], error: { code: 'stale-source', retryable: false, message: 'source stamp changed or store released' } }
      }
      const failure = firstOverlapError(rect.r0, rect.r1, rect.c0, rect.c1)
      const missing = pending.filter((region) => rectsIntersect(region, rect))
      if (failure !== undefined) {
        return { range: rect, completeness: 'incomplete', stamp, missing, error: { code: failure.code, retryable: failure.retryable, message: 'range overlaps a failed region' } }
      }
      return { range: rect, completeness: missing.length > 0 ? 'loading' : 'complete', stamp, missing }
    },
    release: async () => {
      released = true
    },
  }
}
