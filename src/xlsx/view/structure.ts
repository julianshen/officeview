/** Structural range index and populated-chunk indexes (spec sections 3, 4.5, 6.1, 6.3).
 * Pure, source-independent: S1 feeds these from the store/page APIs; sort/filter
 * guards consume them. Rectangles stay true two-dimensional bounds; a multi-row
 * partial-column rectangle is never flattened into one cell-ID interval (spec invariant 4). */
import type {
  AuthoritativeCell,
  Completeness,
  DrawingAnchor,
  MergeExtent,
  PopulatedQuery,
  ReadOptions,
  SourceColInterval,
  SourceRowInterval,
  SourceStamp,
  SourceStore,
  SpillExtent,
  StructuralResult,
  TableExtent,
} from './contracts'
import {
  MAX_COLS,
  MAX_ROWS,
  validateCoordinate,
  validateSourceRect,
  type SourceCell,
  type SourceRect,
} from './coordinates'

export const STRUCTURAL_PAGE_DEFAULT_LIMIT = 1024
export const RECTANGLE_PAGE_DEFAULT_LIMIT = 4096

function rectKey(rect: SourceRect): string {
  return `${rect.r0},${rect.r1},${rect.c0},${rect.c1}`
}
function columnsOverlap(a0: number, a1: number, b0: number, b1: number): boolean {
  return a0 <= b1 && b0 <= a1
}
function checkLimit(limit: number | undefined, what: string): number {
  const value = limit ?? STRUCTURAL_PAGE_DEFAULT_LIMIT
  if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid ${what} page limit`)
  return value
}

/** True two-dimensional rectangle index over caller-owned records.
 * Segment tree over r0-sorted entries with max-r1 augmentation: O(n log n) build,
 * O(records) memory, O(log n + rowHits) queries with an O(1) column check per hit.
 * A full-domain merge is one entry; narrow-column giant rectangles keep exact
 * column bounds, so no phantom coverage appears. Immutable after build. */
export interface RectangleQueryOptions { limit?: number; cursor?: string }
export interface RectanglePage<T> { items: readonly T[]; nextCursor?: string }
interface SortedEntry<T> { r0: number; r1: number; c0: number; c1: number; value: T }
export class RectangleIndex<T> {
  private readonly sorted: readonly SortedEntry<T>[]
  private readonly maxR1: Float64Array
  readonly size: number
  private constructor(sorted: readonly SortedEntry<T>[], maxR1: Float64Array) {
    this.sorted = sorted
    this.maxR1 = maxR1
    this.size = sorted.length
  }
  static build<T>(entries: readonly { rect: SourceRect; value: T }[]): RectangleIndex<T> {
    const sorted = entries.map((entry, order) => {
      validateSourceRect(entry.rect)
      return { r0: entry.rect.r0, r1: entry.rect.r1, c0: entry.rect.c0, c1: entry.rect.c1, value: entry.value, order }
    }).sort((a, b) => a.r0 - b.r0 || a.r1 - b.r1 || a.c0 - b.c0 || a.c1 - b.c1 || a.order - b.order)
      .map(({ r0, r1, c0, c1, value }) => ({ r0, r1, c0, c1, value }))
    const tree = new Float64Array(Math.max(1, 4 * sorted.length)).fill(-1)
    const build = (node: number, left: number, right: number): number => {
      if (left + 1 === right) return (tree[node] = sorted[left].r1)
      const mid = (left + right) >> 1
      return (tree[node] = Math.max(build(node * 2, left, mid), build(node * 2 + 1, mid, right)))
    }
    if (sorted.length > 0) build(1, 0, sorted.length)
    return new RectangleIndex(Object.freeze(sorted), tree)
  }
  /** All hits in deterministic r0 order; callers slice for bounded pages. */
  collect(rect: SourceRect): readonly T[] {
    validateSourceRect(rect)
    const out: T[] = []
    const total = this.sorted.length
    if (total === 0) return Object.freeze(out)
    let lo = 0, hi = total
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.sorted[mid].r0 <= rect.r1) lo = mid + 1
      else hi = mid
    }
    const visit = (node: number, left: number, right: number): void => {
      if (left >= hi || right <= 0 || this.maxR1[node] < rect.r0) return
      if (left + 1 === right) {
        const entry = this.sorted[left]
        if (entry.r1 >= rect.r0 && columnsOverlap(entry.c0, entry.c1, rect.c0, rect.c1)) out.push(entry.value)
        return
      }
      const mid = (left + right) >> 1
      visit(node * 2, left, mid)
      visit(node * 2 + 1, mid, right)
    }
    visit(1, 0, total)
    return Object.freeze(out)
  }
  query(rect: SourceRect, options?: RectangleQueryOptions): RectanglePage<T> {
    const limit = options?.limit === undefined ? RECTANGLE_PAGE_DEFAULT_LIMIT
      : (!Number.isSafeInteger(options.limit) || options.limit <= 0
        ? (() => { throw new RangeError('Invalid rectangle page limit') })() : options.limit)
    const offset = decodeOffset(options?.cursor, rect, 'rectangle')
    const hits = this.collect(rect)
    const items = Object.freeze(hits.slice(offset, offset + limit))
    return offset + limit < hits.length
      ? { items, nextCursor: `rectindex:${rectKey(rect)}:${offset + limit}` }
      : { items }
  }
}
function decodeOffset(cursor: string | undefined, rect: SourceRect, what: string): number {
  if (cursor === undefined) return 0
  const parts = cursor.split(':')
  const offset = Number(parts[parts.length - 1])
  const prefix = what === 'rectangle' ? 'rectindex' : 's3struct'
  if (parts.length !== 3 || parts[0] !== prefix || parts[1] !== rectKey(rect)
    || !Number.isSafeInteger(offset) || offset < 0 || parts[2] !== String(offset)) {
    throw new RangeError(`Invalid ${what} cursor`)
  }
  return offset
}

/** Paged structural answer. The frozen store wire type (StructuralResult) carries no
 * cursor, so bounded paging lives here; S1 adapters page through this index. */
export interface StructuralPage {
  completeness: Completeness
  missing: readonly SourceRect[]
  merges: readonly MergeExtent[]
  tables: readonly TableExtent[]
  spills: readonly SpillExtent[]
  drawings: readonly CellAnchoredDrawing[]
  nextCursor?: string
}
export interface StructuralQueryOptions { limit?: number; cursor?: string }
export type CellAnchoredDrawing = Extract<DrawingAnchor, { kind: 'oneCell' | 'twoCell' }>
export type AbsoluteDrawing = Extract<DrawingAnchor, { kind: 'absolute' }>
/** Immutable index set over one StructuralResult. Spills are indexed per extent
 * (saved, evaluated) plus their anchor cell, so child-only intersections match even
 * when the anchor lies outside the query; absolute anchors are retained separately
 * and never enter the spatial index. */
export class StructuralIndex {
  readonly stamp: SourceStamp
  readonly completeness: Completeness
  readonly missing: readonly SourceRect[]
  readonly absoluteDrawings: readonly AbsoluteDrawing[]
  readonly recordCount: number
  private readonly merges: RectangleIndex<MergeExtent>
  private readonly tables: RectangleIndex<TableExtent>
  private readonly spillEntries: RectangleIndex<SpillExtent>
  private readonly drawings: RectangleIndex<CellAnchoredDrawing>
  constructor(result: StructuralResult) {
    this.stamp = { ...result.stamp }
    this.completeness = result.completeness
    this.missing = Object.freeze(result.missing.map(rect => ({ ...rect })))
    const absolute: AbsoluteDrawing[] = []
    const drawingEntries: { rect: SourceRect; value: CellAnchoredDrawing }[] = []
    for (const drawing of result.drawings) {
      if (drawing.kind === 'absolute') absolute.push(drawing)
      else drawingEntries.push({ rect: { ...drawing.bounds }, value: drawing })
    }
    const spillEntries: { rect: SourceRect; value: SpillExtent }[] = []
    for (const spill of result.spills) {
      if (spill.saved !== undefined) spillEntries.push({ rect: { ...spill.saved }, value: spill })
      if (spill.evaluated !== undefined
        && (spill.saved === undefined || rectKey(spill.evaluated) !== rectKey(spill.saved))) {
        spillEntries.push({ rect: { ...spill.evaluated }, value: spill })
      }
      spillEntries.push({
        rect: { r0: spill.anchor.row, r1: spill.anchor.row, c0: spill.anchor.col, c1: spill.anchor.col },
        value: spill,
      })
    }
    this.merges = RectangleIndex.build(result.merges.map(merge => ({ rect: merge.range, value: merge })))
    this.tables = RectangleIndex.build(result.tables.map(table => ({ rect: table.range, value: table })))
    this.spillEntries = RectangleIndex.build(spillEntries)
    this.drawings = RectangleIndex.build(drawingEntries)
    this.absoluteDrawings = Object.freeze(absolute)
    this.recordCount = this.merges.size + this.tables.size + this.spillEntries.size + this.drawings.size
  }
  query(rect: SourceRect, options?: StructuralQueryOptions): StructuralPage {
    validateSourceRect(rect)
    const limit = checkLimit(options?.limit, 'structural')
    const offset = decodeOffset(options?.cursor, rect, 'structural')
    const spills: SpillExtent[] = []
    const seen = new Set<SpillExtent>()
    for (const spill of this.spillEntries.collect(rect)) {
      if (!seen.has(spill)) {
        seen.add(spill)
        spills.push(spill)
      }
    }
    const ordered: {
      merges: readonly MergeExtent[]; tables: readonly TableExtent[]
      spills: readonly SpillExtent[]; drawings: readonly CellAnchoredDrawing[]
    } = {
      merges: this.merges.collect(rect),
      tables: this.tables.collect(rect),
      spills: Object.freeze(spills),
      drawings: this.drawings.collect(rect),
    }
    const flattened: (MergeExtent | TableExtent | SpillExtent | CellAnchoredDrawing)[] = [
      ...ordered.merges, ...ordered.tables, ...ordered.spills, ...ordered.drawings,
    ]
    const kinds = [
      ...ordered.merges.map(() => 0), ...ordered.tables.map(() => 1),
      ...ordered.spills.map(() => 2), ...ordered.drawings.map(() => 3),
    ]
    const pageMerges: MergeExtent[] = []
    const pageTables: TableExtent[] = []
    const pageSpills: SpillExtent[] = []
    const pageDrawings: CellAnchoredDrawing[] = []
    for (let index = offset; index < Math.min(flattened.length, offset + limit); index++) {
      const item = flattened[index]
      if (kinds[index] === 0) pageMerges.push(item as MergeExtent)
      else if (kinds[index] === 1) pageTables.push(item as TableExtent)
      else if (kinds[index] === 2) pageSpills.push(item as SpillExtent)
      else pageDrawings.push(item as CellAnchoredDrawing)
    }
    const page: StructuralPage = {
      completeness: this.completeness,
      missing: this.missing,
      merges: Object.freeze(pageMerges),
      tables: Object.freeze(pageTables),
      spills: Object.freeze(pageSpills),
      drawings: Object.freeze(pageDrawings),
    }
    return offset + limit < flattened.length
      ? { ...page, nextCursor: `s3struct:${rectKey(rect)}:${offset + limit}` }
      : page
  }
}

/** Sparse populated-cell feed. Only truly populated cells are items: style-only
 * blanks are never data, and unknown/incomplete pages report through completeness
 * plus missing domains instead of a fake safe state. */
export interface PopulatedCell { row: number; col: number }
export interface PopulatedPage {
  items: readonly PopulatedCell[]
  completeness: Completeness
  missing: readonly SourceRect[]
  nextCursor?: string
}
export interface PopulatedCellSource {
  iterate(range: SourceRect, cursor?: string): Promise<PopulatedPage>
}
export type OutsideColumnAnswer =
  | { kind: 'clear' }
  | { kind: 'occupied'; at: SourceCell }
  | { kind: 'unknown'; missing: readonly SourceRect[] }
export interface ChunkBuildResult<T> { index: T; complete: boolean; missing: readonly SourceRect[] }
async function consumeBand(source: PopulatedCellSource, band: SourceRowInterval, pageLimit: number): Promise<{
  byRow: Map<number, number[]>; complete: boolean; missing: readonly SourceRect[]
}> {
  validateCoordinate(band.r0, 0)
  validateCoordinate(band.r1, 0)
  if (band.r0 > band.r1) throw new RangeError('Reversed chunk band')
  if (!Number.isSafeInteger(pageLimit) || pageLimit <= 0) throw new RangeError('Invalid chunk page limit')
  const range: SourceRect = { r0: band.r0, r1: band.r1, c0: 0, c1: MAX_COLS - 1 }
  const byRow = new Map<number, number[]>()
  const missing: SourceRect[] = []
  let complete = true
  let cursor: string | undefined
  const seen = new Set<string>()
  for (;;) {
    const page = await source.iterate(range, cursor)
    for (const cell of page.items) {
      validateCoordinate(cell.row, cell.col)
      if (cell.row < range.r0 || cell.row > range.r1) throw new RangeError('Populated cell outside the requested band')
      let cols = byRow.get(cell.row)
      if (cols === undefined) byRow.set(cell.row, (cols = []))
      cols.push(cell.col)
    }
    if (page.completeness !== 'complete') {
      complete = false
      for (const rect of page.missing) missing.push(rect)
    }
    if (page.nextCursor === undefined) break
    if (seen.has(page.nextCursor)) throw new RangeError('Populated cursor made no progress')
    seen.add(page.nextCursor)
    cursor = page.nextCursor
  }
  return { byRow, complete, missing: Object.freeze(missing) }
}
function toColumnIntervals(cols: number[]): SourceColInterval[] {
  const sorted = [...cols].sort((a, b) => a - b)
  const out: SourceColInterval[] = []
  for (const col of sorted) {
    const previous = out[out.length - 1]
    if (previous && previous.c1 + 1 >= col) previous.c1 = Math.max(previous.c1, col)
    else out.push({ c0: col, c1: col })
  }
  return out
}
function toRowIntervals(rows: number[]): SourceRowInterval[] {
  const sorted = [...rows].sort((a, b) => a - b)
  const out: SourceRowInterval[] = []
  for (const row of sorted) {
    const previous = out[out.length - 1]
    if (previous && previous.r1 + 1 >= row) previous.r1 = Math.max(previous.r1, row)
    else out.push({ r0: row, r1: row })
  }
  return out
}
/** Per-row populated-column intervals for one row band. Answers the
 * outside-columns query without scanning all 16,384 columns per candidate row:
 * only populated cells are ever visited. */
export class RowChunkIndex {
  private readonly rows: ReadonlyMap<number, readonly SourceColInterval[]>
  private readonly orderedRows: readonly number[]
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  private constructor(rows: ReadonlyMap<number, readonly SourceColInterval[]>, complete: boolean, missing: readonly SourceRect[]) {
    this.rows = rows
    this.orderedRows = Object.freeze([...rows.keys()].sort((a, b) => a - b))
    this.complete = complete
    this.missing = missing
  }
  static async build(source: PopulatedCellSource, band: SourceRowInterval, pageLimit = 4096): Promise<ChunkBuildResult<RowChunkIndex>> {
    const consumed = await consumeBand(source, band, pageLimit)
    const rows = new Map<number, readonly SourceColInterval[]>()
    for (const [row, cols] of consumed.byRow) rows.set(row, Object.freeze(toColumnIntervals(cols)))
    return { index: new RowChunkIndex(rows, consumed.complete, consumed.missing), complete: consumed.complete, missing: consumed.missing }
  }
  columnsOf(row: number): readonly SourceColInterval[] {
    validateCoordinate(row, 0)
    return this.rows.get(row) ?? []
  }
  populatedRows(): readonly SourceRowInterval[] {
    return Object.freeze(toRowIntervals([...this.orderedRows]))
  }
  outsideColumnData(c0: number, c1: number): OutsideColumnAnswer {
    validateCoordinate(0, c0)
    validateCoordinate(0, c1)
    if (c0 > c1) throw new RangeError('Reversed outside-column span')
    if (!this.complete) return { kind: 'unknown', missing: this.missing }
    for (const row of this.orderedRows) {
      for (const interval of this.rows.get(row)!) {
        if (interval.c0 < c0) return { kind: 'occupied', at: { row, col: interval.c0 } }
        if (interval.c1 > c1) return { kind: 'occupied', at: { row, col: Math.max(interval.c0, c1 + 1) } }
      }
    }
    return { kind: 'clear' }
  }
}
/** Transposed per-column populated-row intervals for the same band. */
export class ColumnChunkIndex {
  private readonly cols: ReadonlyMap<number, readonly SourceRowInterval[]>
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  private constructor(cols: ReadonlyMap<number, readonly SourceRowInterval[]>, complete: boolean, missing: readonly SourceRect[]) {
    this.cols = cols
    this.complete = complete
    this.missing = missing
  }
  static async build(source: PopulatedCellSource, band: SourceRowInterval, pageLimit = 4096): Promise<ChunkBuildResult<ColumnChunkIndex>> {
    const consumed = await consumeBand(source, band, pageLimit)
    const byCol = new Map<number, number[]>()
    for (const [row, cols] of consumed.byRow) {
      for (const col of cols) {
        let rows = byCol.get(col)
        if (rows === undefined) byCol.set(col, (rows = []))
        rows.push(row)
      }
    }
    const cols = new Map<number, readonly SourceRowInterval[]>()
    for (const [col, rows] of byCol) cols.set(col, Object.freeze(toRowIntervals(rows)))
    return { index: new ColumnChunkIndex(cols, consumed.complete, consumed.missing), complete: consumed.complete, missing: consumed.missing }
  }
  rowsOf(col: number): readonly SourceRowInterval[] {
    validateCoordinate(0, col)
    return this.cols.get(col) ?? []
  }
}
function intervalContainsCol(intervals: readonly SourceColInterval[], col: number): boolean {
  let lo = 0, hi = intervals.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (intervals[mid].c1 < col) lo = mid + 1
    else hi = mid
  }
  return lo < intervals.length && intervals[lo].c0 <= col
}
function intervalContainsRow(intervals: readonly SourceRowInterval[], row: number): boolean {
  let lo = 0, hi = intervals.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (intervals[mid].r1 < row) lo = mid + 1
    else hi = mid
  }
  return lo < intervals.length && intervals[lo].r0 <= row
}
/** Bounding box of the 4-connected populated component containing the seed, found
 * through sparse adjacency only. Diagonally touching cells are separate regions.
 * An unpopulated seed has no region. Excel-rectangle expansion beyond the
 * connected component (including blank interior) belongs to S7 range resolution. */
export function currentRegion(seed: SourceCell, rows: RowChunkIndex, cols: ColumnChunkIndex): SourceRect | null {
  validateCoordinate(seed.row, seed.col)
  if (!intervalContainsCol(rows.columnsOf(seed.row), seed.col)
    || !intervalContainsRow(cols.rowsOf(seed.col), seed.row)) return null
  const visited = new Set<number>()
  const box: SourceRect = { r0: seed.row, r1: seed.row, c0: seed.col, c1: seed.col }
  const queue: { row: number; c0: number; c1: number }[] = [{ row: seed.row, c0: seed.col, c1: seed.col }]
  const keyOf = (row: number, c0: number): number => row * MAX_COLS + c0
  for (const seedInterval of rows.columnsOf(seed.row)) {
    if (seedInterval.c0 <= seed.col && seed.col <= seedInterval.c1) {
      queue[0] = { row: seed.row, c0: seedInterval.c0, c1: seedInterval.c1 }
      visited.add(keyOf(seed.row, seedInterval.c0))
      box.c0 = seedInterval.c0
      box.c1 = seedInterval.c1
      break
    }
  }
  while (queue.length > 0) {
    const current = queue.pop()!
    box.r0 = Math.min(box.r0, current.row)
    box.r1 = Math.max(box.r1, current.row)
    box.c0 = Math.min(box.c0, current.c0)
    box.c1 = Math.max(box.c1, current.c1)
    for (const next of [current.row - 1, current.row + 1]) {
      if (next < 0 || next >= MAX_ROWS) continue
      for (const interval of rows.columnsOf(next)) {
        if (interval.c1 < current.c0 || interval.c0 > current.c1) continue
        const key = keyOf(next, interval.c0)
        if (visited.has(key)) continue
        visited.add(key)
        queue.push({ row: next, c0: interval.c0, c1: interval.c1 })
      }
    }
  }
  return box
}
/** Store-backed populated source so S1 can feed the chunk indexes from
 * scanPopulated. Store cursors pass through opaquely; the store itself rejects
 * foreign cursors with an invalid-cursor error page, which surfaces as a throw.
 * Loading pages report the requested range as missing (store pages carry no
 * missing-rect list); operational failures throw instead of faking a page. */
export function fromStoreScan(store: SourceStore, options: ReadOptions, pageLimit = 4096): PopulatedCellSource {
  if (!Number.isSafeInteger(pageLimit) || pageLimit <= 0) throw new RangeError('Invalid store scan page limit')
  return {
    async iterate(range: SourceRect, cursor?: string): Promise<PopulatedPage> {
      validateSourceRect(range)
      if (cursor !== undefined && (typeof cursor !== 'string' || cursor.length === 0)) {
        throw new RangeError('Invalid populated cursor')
      }
      const request: PopulatedQuery = cursor === undefined
        ? { range, limit: pageLimit }
        : { range, limit: pageLimit, cursor }
      for await (const page of store.scanPopulated(request, options)) {
        if (page.error !== undefined) throw new RangeError(`Populated scan failed: ${page.error.code}`)
        const items = Object.freeze(page.items.map((cell: AuthoritativeCell) => ({ row: cell.row, col: cell.col })))
        if (page.completeness === 'complete') return { items, completeness: page.completeness, missing: [], nextCursor: page.nextCursor }
        return { items, completeness: page.completeness, missing: Object.freeze([{ ...range }]), nextCursor: page.nextCursor }
      }
      throw new RangeError('Populated scan ended without a page')
    },
  }
}
