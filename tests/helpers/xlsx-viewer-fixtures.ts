/** S0.2 deterministic fixture family (spec §12.5, plan S0.2).
 * Every builder is a pure function of (seed, row, col): no Math.random, no
 * rows×cols allocation. Sparse fixtures index only populated coordinates;
 * dense fixtures compute values arithmetically per address. */
import { createHash } from 'node:crypto'
import type {
  AuthoritativeCell, CellRead, ColumnEntry, CompactSourceRowSet, Completeness,
  OverviewGroup, Page, ReadOptions, SingleCellPayload, SourceStamp, SourceStore,
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
  /** F1 sampled NNZ (default 100000). */
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
}

export interface CellRef {
  row: number
  col: number
  cell: FixtureCell
}

export interface FixtureHandle {
  readonly id: FixtureId
  readonly seed: number
  readonly rowCount: number
  readonly colCount: number
  readonly used: SourceRect
  nnz(): number
  cellAt(row: number, col: number): FixtureCell | null
  structure(): FixtureStructure
  /** Exact enumeration for small enumerable fixtures (F1 sampled set, F8 cases).
   * Dense and full-axis fixtures are not enumerable; use scanCells instead. */
  enumerateHead(limit: number): CellRef[]
  summarizeDomain(rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>): DomainSummary
  scanCells(range: SourceRect, offset: number, limit: number): CellRef[]
  createStore(options?: SyntheticStoreOptions): SourceStore
  canonicalSamples(): string
}

export function sha256Hex(text: string): string {
  return createHash('sha256').update(text, 'utf8').digest('hex')
}

export function readOptionsFor(store: SourceStore): ReadOptions {
  return { signal: new AbortController().signal, stamp: { ...store.stamp } }
}

/** Page window for readRows; callers bound ranges, the adapter bounds pages. */
export const READ_ROWS_PAGE_SIZE = 200
/** Admission guard: a single readRows call resolving more cells refuses loudly. */
export const READ_ROWS_MAX_CELLS = 262144

function rectContains(rect: SourceRect, row: number, col: number): boolean {
  return row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1
}

function rectsIntersect(a: SourceRect, b: SourceRect): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
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

function f1SampledIds(seed: number, nnz: number): Set<number> {
  const prng = createPrng(seed)
  const ids = new Set<number>()
  for (const corner of F1_CORNERS) ids.add(cellId(corner.row, corner.col))
  while (ids.size < nnz + F1_CORNERS.length) {
    ids.add(Math.floor(prng() * 2 ** 34))
  }
  return ids
}

const F2_HEADERS = ['id', 'date', 'flag', 'name', 'amount', 'ratio', 'code', 'note']
const F2_POOL = ['alpha', 'beta', 'gamma', 'delta']

function f2TotalsFormula(rowCount: number): string {
  return `SUM(B2:B${rowCount - 1})`
}

function f2BodyCell(row: number, col: number): FixtureCell | null {
  switch (col) {
    case 0: return { value: row, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 1: return { value: 45000 + ((row * 7) % 365), hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DATE }
    case 2: return { value: row % 2 === 0, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
    case 3: return { value: F2_POOL[(row * 5 + 1) % F2_POOL.length], hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
    case 4: return row % 7 === 3
      ? null
      : { value: row * 1.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 5: return row % 101 === 0
      ? { value: '#DIV/0!', hasCachedValue: true, valueIsError: true, styleIndex: STYLE_DEFAULT }
      : { value: row + 0.5, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    case 6: return { value: row * 2 + 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
    default: return { value: `note-${row % 16}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  }
}

function f2BlankCount(rowCount: number): number {
  const first = 1 + (((3 - 1) % 7) + 7) % 7
  const last = rowCount - 2
  return first > last ? 0 : Math.floor((last - first) / 7) + 1
}

function f3DenseCell(row: number, col: number): FixtureCell {
  if ((row + col) % 8 === 0) return { value: `w${col % 256}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  return { value: ((row * MAX_COLS + col) % 100000) / 10, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
}

function f3SparsePopulated(row: number, col: number): boolean {
  const id = row * MAX_COLS + col
  return id % 64 === 7 || col === MAX_COLS - 1
}

/** 1M unique raw UTF-16 strings across accents, case, suffixes, surrogates. */
export const F4_STRING_COUNT = 1_000_000
export const F4_LONG_BLOB_INDEX = 999983
export function f4StringAt(index: number): string {
  if (index === F4_LONG_BLOB_INDEX) return `blob-${index}-` + 'x'.repeat(100_000)
  switch (index % 8) {
    case 0: return `item-${index}`
    case 1: return index % 2 === 0 ? `café-${index}` : `naïve-${index}`
    case 2: return index % 2 === 0 ? `Alpha-${index}` : `alpha-${index}`
    case 3: return `${index % 2 === 0 ? '\uD800' : '\uDE00'}-lone-${index}`
    case 4: return index % 2 === 0 ? `collide-${index}-1.5` : `collide-${index}-1.50`
    case 5: return `long-${index}-` + 'y'.repeat(200 + (index % 500))
    case 6: return `😀-${index}`
    default: return `row-${index}`
  }
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
    { note: 'sparse-edge-cell', address: { row: MAX_ROWS - 1, col: MAX_COLS - 1 }, cell: { value: 'edge', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT } },
  ]
}

export interface F7StressSpec {
  sortCols: number[]
  hiddenIntervals: SourceRect[]
  sizeOverrides: Array<{ start: number; end: number; sizeCss: number }>
  sheetCount: number
  rowCount: number
  seed: number
}

export function f7Spec(): F7StressSpec {
  return {
    sortCols: [0, 1, 2, 3, 4, 5, 6, 7],
    hiddenIntervals: [
      { r0: 5, r1: 5, c0: 0, c1: 7 },
      { r0: 10, r1: 12, c0: 0, c1: 7 },
      { r0: 100, r1: 199, c0: 0, c1: 7 },
    ],
    sizeOverrides: [{ start: 0, end: 0, sizeCss: 12 }],
    sheetCount: 12,
    rowCount: 100_000,
    seed: FIXTURE_SEEDS.F7,
  }
}

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
  }
}

function f5CellAt(row: number, col: number): FixtureCell | null {
  if (row === 45 && col === 3) return { value: 7, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
  if (row === 32 && col === 0) return { value: 1, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'SUM(A33:A35)', arrayRef: 'A33:B35', formulaType: 'array' }
  if (row === 40 && col === 3) return { value: 2, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: 'SEQUENCE(11)' }
  if (row === 4 && col <= 3) return { value: `head-${col}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
  if (row < 10 && col < 10 && (row + col) % 3 === 0) return { value: row * 10 + col, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
  return null
}

function f5Nnz(): number {
  let count = 0
  for (let row = 0; row < 100; row++) for (let col = 0; col < 10; col++) {
    if (f5CellAt(row, col) !== null) count++
  }
  return count
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

function firstPopulated(
  cellAt: (row: number, col: number) => FixtureCell | null,
  rowCount: number, colCount: number, limit: number,
): CellRef[] {
  const out: CellRef[] = []
  const stride = Math.max(1, Math.floor((rowCount * colCount) / (limit * 64)))
  for (let row = 0; row < rowCount && out.length < limit; row += 1) {
    for (let col = 0; col < colCount && out.length < limit; col += stride) {
      const cell = cellAt(row, col)
      if (cell !== null) out.push({ row, col, cell })
    }
  }
  return out
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
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used,
        nnz: () => sampled.size,
        cellAt,
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
        enumeratePopulated: (visit) => {
          for (const lookup of sorted) {
            const row = Math.floor(lookup / MAX_COLS)
            visit(row, lookup - row * MAX_COLS, f1ValueFor(lookup))
          }
        },
        scanFrom: (range, offset, limit) => {
          const out: CellRef[] = []
          let skipped = 0
          const start = lowerBoundId(sorted, range.r0 * MAX_COLS)
          for (let i = start; i < sorted.length && out.length < limit; i++) {
            const lookup = sorted[i]
            const row = Math.floor(lookup / MAX_COLS)
            if (row > range.r1) break
            const col = lookup - row * MAX_COLS
            if (col < range.c0 || col > range.c1) continue
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
            date += 45000 + ((row * 7) % 365)
            if (row % 7 !== 3) amount += row * 1.5
          }
          totals = { date, amount }
        }
        return totals
      }
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col) || col > 7 || row >= rowCount) return null
        if (row === 0) {
          return { value: F2_HEADERS[col], hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
        }
        if (row === rowCount - 1) {
          if (col === 0) return { value: 'total', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_HEADER }
          if (col === 1) return { value: totalsOf().date, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER, formula: f2TotalsFormula(rowCount) }
          if (col === 3) return { value: 'end', hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
          if (col === 4) return { value: totalsOf().amount, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
          return null
        }
        return f2BodyCell(row, col)
      }
      const used: SourceRect = { r0: 0, r1: rowCount - 1, c0: 0, c1: 7 }
      return makeHandle({
        id, seed, rowCount, colCount: 8, used,
        nnz: () => 8 + (rowCount - 2) * 8 - f2BlankCount(rowCount) + 4,
        cellAt,
      })
    }
    case 'F3': {
      const rowCount = params.rowCount ?? 256
      const colCount = params.colCount ?? MAX_COLS
      const variant = params.variant ?? 'dense'
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col) || row >= rowCount || col >= colCount) return null
        if (variant === 'sparse-wide') return f3SparsePopulated(row, col) ? f3DenseCell(row, col) : null
        return f3DenseCell(row, col)
      }
      const used: SourceRect = { r0: 0, r1: rowCount - 1, c0: 0, c1: colCount - 1 }
      return makeHandle({
        id, seed, rowCount, colCount, used,
        nnz: () => {
          if (variant === 'dense') return rowCount * colCount
          let count = 0
          for (let row = 0; row < rowCount; row++) for (let col = 0; col < colCount; col++) {
            if (f3SparsePopulated(row, col)) count++
          }
          return count
        },
        cellAt,
      })
    }
    case 'F4': {
      const rowCount = params.rowCount ?? F4_STRING_COUNT
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col) || col !== 0 || row >= rowCount) return null
        return { value: f4StringAt(row), hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
      }
      const used: SourceRect = { r0: 0, r1: rowCount - 1, c0: 0, c1: 0 }
      return makeHandle({ id, seed, rowCount, colCount: 1, used, nnz: () => rowCount, cellAt })
    }
    case 'F5': {
      const used: SourceRect = { r0: 0, r1: 199999, c0: 0, c1: 100 }
      return makeHandle({ id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used, nnz: f5Nnz, cellAt: f5CellAt })
    }
    case 'F6': {
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
      return makeHandle({ id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used, nnz: () => 2 ** 34, cellAt: f6CellAt })
    }
    case 'F7': {
      const spec = f7Spec()
      const rowCount = params.rowCount ?? spec.rowCount
      const cellAt = (row: number, col: number): FixtureCell | null => {
        if (!isValidCoordinate(row, col) || col > 7 || row >= rowCount) return null
        if (col === 7) return { value: `s${row % 64}`, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_DEFAULT }
        return { value: row * 8 + col, hasCachedValue: true, valueIsError: false, styleIndex: STYLE_NUMBER }
      }
      const used: SourceRect = { r0: 0, r1: rowCount - 1, c0: 0, c1: 7 }
      return makeHandle({ id, seed, rowCount, colCount: 8, used, nnz: () => rowCount * 8, cellAt })
    }
    case 'F8': {
      const cases = f8Cases()
      const byAddress = new Map(cases.map((c) => [`${c.address.row}:${c.address.col}`, c.cell]))
      const used: SourceRect = { r0: 0, r1: MAX_ROWS - 1, c0: 0, c1: MAX_COLS - 1 }
      const ordered = [...cases].sort((a, b) => a.address.row - b.address.row || a.address.col - b.address.col)
      return makeHandle({
        id, seed, rowCount: MAX_ROWS, colCount: MAX_COLS, used,
        nnz: () => cases.length,
        cellAt: (row, col) => byAddress.get(`${row}:${col}`) ?? null,
        enumerateHead: (limit) => ordered.slice(0, limit).map((c) => ({ row: c.address.row, col: c.address.col, cell: c.cell })),
        enumeratePopulated: (visit) => {
          for (const c of ordered) visit(c.address.row, c.address.col, c.cell)
        },
        scanFrom: (range, offset, limit) => {
          const out: CellRef[] = []
          let skipped = 0
          for (const c of ordered) {
            if (out.length >= limit) break
            if (!rectContains(range, c.address.row, c.address.col)) continue
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
  enumeratePopulated?: (visit: (row: number, col: number, cell: FixtureCell) => void) => void
  scanFrom?: (range: SourceRect, offset: number, limit: number) => CellRef[]
  enumerateHead?: (limit: number) => CellRef[]
}): FixtureHandle {
  const { id, seed, rowCount, colCount, used } = options
  const structure = (): FixtureStructure => {
    if (id !== 'F5') return { merges: [], tables: [], spills: [], drawings: [] }
    return f5Structure()
  }
  const summarizeDomain = (rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>): DomainSummary => {
    if (rows.kind === 'ref') return { totalRows: 0, populatedRows: 0, populatedCells: 0, uniformStyle: STYLE_DEFAULT }
    const totalRows = countRowSetRows(rows)
    if (id === 'F6') {
      const width = cols.reduce((sum, span) => sum + (span.c1 - span.c0 + 1), 0)
      const populatedCells = totalRows * width
      return { totalRows, populatedRows: totalRows, populatedCells, uniformStyle: populatedCells <= 1 ? STYLE_NUMBER : null }
    }
    let populatedRows = 0
    let populatedCells = 0
    let style: number | null | undefined
    const noteStyle = (styleIndex: number) => {
      if (style === undefined) style = styleIndex
      else if (style !== styleIndex) style = null
    }
    if (options.enumeratePopulated !== undefined) {
      const seen = new Set<number>()
      options.enumeratePopulated((row, col, cell) => {
        if (!rowSetContains(rows, row)) return
        if (!cols.some((span) => col >= span.c0 && col <= span.c1)) return
        populatedCells++
        noteStyle(cell.styleIndex)
        if (!seen.has(row)) {
          seen.add(row)
          populatedRows++
        }
      })
    } else {
      for (const row of rowSetMembers(rows)) {
        let rowPopulated = false
        for (const span of cols) {
          for (let col = span.c0; col <= span.c1; col++) {
            const cell = options.cellAt(row, col)
            if (cell === null) continue
            populatedCells++
            rowPopulated = true
            noteStyle(cell.styleIndex)
          }
        }
        if (rowPopulated) populatedRows++
      }
    }
    return { totalRows, populatedRows, populatedCells, uniformStyle: style === undefined ? STYLE_DEFAULT : style }
  }
  const scanCells = (range: SourceRect, offset: number, limit: number): CellRef[] => {
    if (options.scanFrom !== undefined) return options.scanFrom(range, offset, limit)
    const out: CellRef[] = []
    let skipped = 0
    for (let row = range.r0; row <= range.r1 && out.length < limit; row++) {
      for (let col = range.c0; col <= range.c1 && out.length < limit; col++) {
        const cell = options.cellAt(row, col)
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
  const canonicalSamples = (): string => {
    const head = firstPopulated(options.cellAt, Math.min(rowCount, 4096), Math.min(colCount, 64), 24)
    return JSON.stringify({ id, seed, rowCount, colCount, nnz: memoizedNnz(), head: canonicalCellList(head) })
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
    id, seed, rowCount, colCount, used, nnz: memoizedNnz, cellAt: options.cellAt, structure, enumerateHead,
    summarizeDomain, scanCells, canonicalSamples,
    createStore: (storeOptions) => createSyntheticStore(
      { id, used, cellAt: options.cellAt, structure, summarizeDomain, scanCells }, storeOptions,
    ),
  }
}

interface StoreBackend {
  id: FixtureId
  used: SourceRect
  cellAt: (row: number, col: number) => FixtureCell | null
  structure: () => FixtureStructure
  summarizeDomain: (rows: CompactSourceRowSet, cols: ReadonlyArray<{ c0: number; c1: number }>) => DomainSummary
  scanCells: (range: SourceRect, offset: number, limit: number) => CellRef[]
}

function stalePage<T>(stamp: SourceStamp, code: 'stale-source'): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error: { code, retryable: false, message: 'source stamp changed or store released' } }
}

function badCursorPage<T>(stamp: SourceStamp): Page<T> {
  return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-cursor', retryable: false, message: 'cursor belongs to another query or stamp' } }
}

function parseOffset(cursor: string | undefined, prefix: string): number | null {
  if (cursor === undefined) return 0
  const match = new RegExp(`^${prefix}:(\\d+)$`).exec(cursor)
  if (match === null) return null
  const offset = Number(match[1])
  return Number.isSafeInteger(offset) ? offset : null
}

function createSyntheticStore(backend: StoreBackend, storeOptions: SyntheticStoreOptions = {}): SourceStore {
  const valuePolicy = storeOptions.valuePolicy ?? 'engine'
  const pending = storeOptions.pending ?? []
  const errors = storeOptions.errors ?? []
  const stamp: SourceStamp = { docId: `viewer-${backend.id}`, sheetId: `${backend.id}-sheet-1`, valueGen: 0 }
  let released = false
  const gone = (candidate: SourceStamp): boolean => released
    || candidate.docId !== stamp.docId || candidate.sheetId !== stamp.sheetId || candidate.valueGen !== stamp.valueGen

  const pendingAt = (row: number, col: number): number => pending.findIndex((rect) => rectContains(rect, row, col))
  const errorAt = (row: number, col: number) => errors.find((entry) => rectContains(entry.range, row, col))

  const toRead = (row: number, col: number): CellRead => {
    if (!isValidCoordinate(row, col)) return { kind: 'error', code: 'invalid-coordinate', retryable: false }
    const failure = errorAt(row, col)
    if (failure !== undefined) return { kind: 'error', code: failure.code, retryable: failure.retryable }
    const page = pendingAt(row, col)
    if (page >= 0) return { kind: 'unknown', pageId: `pending-${page}` }
    const cell = backend.cellAt(row, col)
    if (cell === null) return { kind: 'blank' }
    if (valuePolicy === 'cached' && isUnavailableEligible(cell)) {
      return { kind: 'unavailable', cell: toSingleCell(row, col, cell, valuePolicy), reason: 'formula-value-unavailable' }
    }
    return { kind: 'cell', cell: toSingleCell(row, col, cell, valuePolicy) }
  }

  const entryFor = (row: number, col: number): ColumnEntry => {
    const read = toRead(row, col)
    switch (read.kind) {
      case 'cell': return { kind: 'cell', row, cell: read.cell }
      case 'blank': return { kind: 'blank', r0: row, r1: row }
      case 'unknown': return { kind: 'unknown', r0: row, r1: row, pageId: read.pageId }
      case 'error': return { kind: 'error', r0: row, r1: row, code: read.code, retryable: read.retryable }
      case 'unavailable': return { kind: 'unavailable', row, cell: read.cell, reason: read.reason }
    }
  }

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
      const start = parseOffset(cursor, 'row')
      if (start === null) return badCursorPage(stamp)
      const first = Math.max(range.r0, start)
      if (first > range.r1) return { items: [], completeness: 'complete', stamp }
      const end = Math.min(range.r1, first + READ_ROWS_PAGE_SIZE - 1)
      const items: SparseRowPage[] = []
      let completeness: Completeness = 'complete'
      for (let row: SourceRow = first; row <= end; row++) {
        const rowPending = pending.some((rect) => rect.r0 <= row && row <= rect.r1 && rect.c0 <= range.c1 && range.c0 <= rect.c1)
        if (rowPending) {
          completeness = 'loading'
          continue
        }
        const cells: AuthoritativeCell[] = []
        for (let col = range.c0; col <= range.c1; col++) {
          const cell = backend.cellAt(row, col)
          if (cell === null || errorAt(row, col) !== undefined) continue
          cells.push(toAuthoritative(row, col, cell, valuePolicy))
        }
        items.push({ row, cells })
      }
      return { items, completeness, stamp, nextCursor: end < range.r1 ? `row:${end + 1}` : undefined }
    },
    queryColumn: async (request, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return stalePage(stamp, 'stale-source')
      if (!isValidCoordinate(0, request.col)) {
        return { items: [], completeness: 'incomplete', stamp, error: { code: 'invalid-coordinate', retryable: false, message: 'column is outside the worksheet domain' } }
      }
      if (request.rows.kind === 'ref') return badCursorPage(stamp)
      if (!Number.isSafeInteger(request.limit) || request.limit < 0) return badCursorPage(stamp)
      const offset = parseOffset(request.cursor, 'cell')
      if (offset === null) return badCursorPage(stamp)
      const items: ColumnEntry[] = []
      let skipped = 0
      let consumed = 0
      let finished = true
      let blankStart: number | null = null
      let blankEnd = 0
      let loading = false
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
          if (blankStart === null) {
            blankStart = row
            items.push({ kind: 'blank', r0: row, r1: row })
          } else {
            blankEnd = row
            items[items.length - 1] = { kind: 'blank', r0: blankStart, r1: blankEnd }
          }
          continue
        }
        items.push(entry)
        if (entry.kind === 'unknown') loading = true
      }
      return {
        items, stamp,
        completeness: loading ? 'loading' : 'complete',
        nextCursor: finished ? undefined : `cell:${offset + consumed}`,
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
      if (!Number.isSafeInteger(request.limit) || request.limit < 0) {
        yield badCursorPage(stamp)
        return
      }
      const offset = parseOffset(request.cursor, 'cell')
      if (offset === null) {
        yield badCursorPage(stamp)
        return
      }
      const found = backend.scanCells(request.range, offset, request.limit)
      yield {
        items: found.map(({ row, col, cell }) => toAuthoritative(row, col, cell, valuePolicy)),
        completeness: 'complete',
        stamp,
        nextCursor: found.length === request.limit ? `cell:${offset + request.limit}` : undefined,
      }
    },
    queryStructural: async (rect, options) => {
      options.signal.throwIfAborted()
      const empty: StructuralResult = { stamp, completeness: 'complete', merges: [], tables: [], spills: [], drawings: [], missing: [] }
      if (gone(options.stamp)) return { ...empty, completeness: 'incomplete', error: { code: 'stale-source', retryable: false, message: 'source stamp changed or store released' } }
      if (!isValidSourceRect(rect)) {
        return { ...empty, completeness: 'incomplete', error: { code: 'invalid-source-rect', retryable: false, message: 'rect is outside the worksheet domain' } }
      }
      const structure = backend.structure()
      // Absolute anchors are viewport-fixed with no source bounds; rect queries
      // report only source-indexed records (absolute placement belongs to S5).
      const indexed = structure.drawings.filter((drawing) => drawing.kind === 'absolute'
        ? false
        : rectsIntersect(drawing.bounds, rect))
      return {
        stamp,
        completeness: 'complete',
        merges: structure.merges.filter((range) => rectsIntersect(range, rect))
          .map((range, index) => ({ id: `merge-${index}`, range, anchor: { row: range.r0, col: range.c0 } })),
        tables: structure.tables.filter((table) => rectsIntersect(table.range, rect))
          .map((table, index) => ({ id: `table-${index}`, range: table.range, headerRows: table.headerRows, totalsRows: table.totalsRows })),
        spills: structure.spills.filter((spill) => {
          const bounds: SourceRect[] = [spill.saved, spill.evaluated].filter((rect): rect is SourceRect => rect !== undefined)
          return bounds.some((extent) => rectsIntersect(extent, rect))
            || (spill.anchor.row >= rect.r0 && spill.anchor.row <= rect.r1 && spill.anchor.col >= rect.c0 && spill.anchor.col <= rect.c1)
        }),
        drawings: indexed.map((drawing) => {
          if (drawing.kind === 'absolute') return { id: drawing.id, kind: 'absolute', extent: { x: 0, y: 0, width: 0, height: 0 } } as const
          const from = { ...drawing.from, offsetXEmu: 0, offsetYEmu: 0 }
          if (drawing.kind === 'oneCell') {
            return { id: drawing.id, kind: 'oneCell', from, extent: { x: 0, y: 0, width: 0, height: 0 }, bounds: drawing.bounds }
          }
          const to = { ...drawing.to, offsetXEmu: 0, offsetYEmu: 0 }
          return { id: drawing.id, kind: 'twoCell', from, to, extent: { x: 0, y: 0, width: 0, height: 0 }, bounds: drawing.bounds }
        }),
        missing: [],
      }
    },
    queryOverview: async (request, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) return stalePage(stamp, 'stale-source')
      const start = parseOffset(request.cursor, 'group')
      if (start === null) return badCursorPage(stamp)
      const bound = Number.isSafeInteger(request.maxGroups) && request.maxGroups >= 0 ? request.maxGroups : 0
      const items: OverviewGroup[] = []
      const end = Math.min(request.groups.length, start + bound)
      for (let index = start; index < end; index++) {
        const domain = request.groups[index]
        if (domain.rows.kind === 'ref') return badCursorPage(stamp)
        const overlapsPending = pending.some((rect) => rowSetOverlapsRange(domain.rows, rect.r0, rect.r1)
          && domain.cols.some((span) => span.c0 <= rect.c1 && rect.c0 <= span.c1))
        if (overlapsPending) {
          items.push({ groupId: domain.groupId, content: 'unknown', pageId: `pending-overview-${index}` })
          continue
        }
        const summary = backend.summarizeDomain(domain.rows, domain.cols)
        if (summary.totalRows === 0 || summary.populatedRows === 0) {
          items.push({
            groupId: domain.groupId, populatedCount: 0, countKind: 'exact',
            style: { kind: 'uniform', styleIndex: STYLE_DEFAULT }, content: 'blank',
          })
          continue
        }
        const mixed = summary.populatedRows < summary.totalRows
        items.push({
          groupId: domain.groupId, populatedCount: summary.populatedCells, countKind: 'exact',
          style: summary.uniformStyle === null ? { kind: 'mixed' } : { kind: 'uniform', styleIndex: summary.uniformStyle },
          content: mixed ? 'mixed' : 'populated',
        })
      }
      return { items, completeness: 'complete', stamp, nextCursor: end < request.groups.length ? `group:${end}` : undefined }
    },
    ensureComplete: async (rect, options) => {
      options.signal.throwIfAborted()
      if (gone(options.stamp)) {
        return { range: rect, completeness: 'incomplete', stamp, missing: [], error: { code: 'stale-source', retryable: false, message: 'source stamp changed or store released' } }
      }
      const missing = pending.filter((region) => rectsIntersect(region, rect))
      return { range: rect, completeness: missing.length > 0 ? 'loading' : 'complete', stamp, missing }
    },
    release: async () => {
      released = true
    },
  }
}
