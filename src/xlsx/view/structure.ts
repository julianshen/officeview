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
 * O(records) memory. Row filtering is O(log n + rowHits); the O(1) column check
 * runs per row hit, so column-narrow queries over tall records visit rowHits
 * leaves per hit found (worst case O(n) node visits for full-height records;
 * full-width band queries have rowHits == hits). Cursors resume at entry indexes,
 * so pages never re-collect prior hits: each page costs O(log n) redescent plus
 * its own examinations. A full-domain merge is one entry; narrow-column giant
 * rectangles keep exact column bounds, so no phantom coverage appears.
 * Immutable after build. */
export interface RectangleQueryOptions { limit?: number; cursor?: string; countVisits?: (visits: number) => void }
export interface RectanglePage<T> { items: readonly T[]; nextCursor?: string }
export interface WalkOptions<T> { pred?: (value: T) => boolean; countVisits?: (visits: number) => void }
export interface WalkResult<T> { items: readonly T[]; resume: number; exhausted: boolean }
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
    return this.walkFrom(rect, 0, Number.MAX_SAFE_INTEGER).items
  }
  /** First hit in index order with early exit; no full drain. */
  first(rect: SourceRect): T | undefined {
    return this.walkFrom(rect, 0, 1).items[0]
  }
  /** First hit satisfying pred with early exit; later hits are never visited. */
  firstWhere(rect: SourceRect, pred: (value: T) => boolean): T | undefined {
    return this.walkFrom(rect, 0, 1, { pred }).items[0]
  }
  /** Low-level resumable primitive: examines candidate leaves from startIndex,
   * emitting those that pass the column check and pred, stopping after maxEmit
   * emissions. resume advances past every examined leaf, so callers make strict
   * progress; exhausted is true only when the walk ran to completion. */
  walkFrom(rect: SourceRect, startIndex: number, maxEmit: number, options?: WalkOptions<T>): WalkResult<T> {
    validateSourceRect(rect)
    if (!Number.isSafeInteger(startIndex) || startIndex < 0) throw new RangeError('Invalid rectangle walk origin')
    if (!Number.isSafeInteger(maxEmit) || maxEmit < 1) throw new RangeError('Invalid rectangle walk budget')
    return this.walk(rect, startIndex, maxEmit, options?.pred, options?.countVisits)
  }
  private walk(
    rect: SourceRect, startIndex: number, maxEmit: number,
    pred: ((value: T) => boolean) | undefined, report: ((visits: number) => void) | undefined,
  ): WalkResult<T> {
    const out: T[] = []
    let resume = startIndex
    let visits = 0
    const total = this.sorted.length
    let lo = 0, hi = total
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.sorted[mid].r0 <= rect.r1) lo = mid + 1
      else hi = mid
    }
    const visit = (node: number, left: number, right: number): boolean => {
      visits++
      if (right <= startIndex || left >= hi || this.maxR1[node] < rect.r0) return false
      if (left + 1 === right) {
        const entry = this.sorted[left]
        resume = left + 1
        if (entry.r1 >= rect.r0 && columnsOverlap(entry.c0, entry.c1, rect.c0, rect.c1)
          && (pred === undefined || pred(entry.value))) {
          out.push(entry.value)
          if (out.length >= maxEmit) return true
        }
        return false
      }
      const mid = (left + right) >> 1
      return visit(node * 2, left, mid) || visit(node * 2 + 1, mid, right)
    }
    let stopped = false
    if (total > 0) stopped = visit(1, 0, total)
    if (report !== undefined) report(visits)
    return { items: Object.freeze(out), resume, exhausted: !stopped }
  }
  private decodeCursor(cursor: string | undefined, rect: SourceRect): number {
    if (cursor === undefined) return 0
    const parts = cursor.split(':')
    if (parts.length !== 3 || parts[0] !== 'rectindex2' || parts[1] !== rectKey(rect)) {
      throw new RangeError('Invalid rectangle cursor')
    }
    const offset = Number(parts[2])
    if (!Number.isSafeInteger(offset) || offset < 0 || offset > this.size || parts[2] !== String(offset)) {
      throw new RangeError('Invalid rectangle cursor')
    }
    return offset
  }
  query(rect: SourceRect, options?: RectangleQueryOptions): RectanglePage<T> {    validateSourceRect(rect)
    const limit = options?.limit === undefined ? RECTANGLE_PAGE_DEFAULT_LIMIT
      : (!Number.isSafeInteger(options.limit) || options.limit <= 0
        ? (() => { throw new RangeError('Invalid rectangle page limit') })() : options.limit)
    const startIndex = this.decodeCursor(options?.cursor, rect)
    const found = this.walk(rect, startIndex, limit, undefined, options?.countVisits)
    const items = Object.freeze(found.items)
    return found.exhausted
      ? { items }
      : { items, nextCursor: `rectindex2:${rectKey(rect)}:${found.resume}` }
  }
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
export interface StructuralQueryOptions { limit?: number; cursor?: string; countVisits?: (visits: number) => void }
export type CellAnchoredDrawing = Extract<DrawingAnchor, { kind: 'oneCell' | 'twoCell' }>
export type AbsoluteDrawing = Extract<DrawingAnchor, { kind: 'absolute' }>
export type StructuralHit =
  | { kind: 'merge'; record: MergeExtent }
  | { kind: 'table'; record: TableExtent }
  | { kind: 'spill'; record: SpillExtent }
  | { kind: 'drawing'; record: CellAnchoredDrawing }
/** One entry per spill over the union of saved, evaluated and anchor extents.
 * The union may cover gaps between extents, so queries verify exact matches;
 * single entries keep cursors exactamente resumable with no cross-page dedup. */
function spillUnion(spill: SpillExtent): SourceRect {
  let r0 = spill.anchor.row, r1 = spill.anchor.row, c0 = spill.anchor.col, c1 = spill.anchor.col
  for (const extent of [spill.saved, spill.evaluated]) {
    if (extent === undefined) continue
    r0 = Math.min(r0, extent.r0); r1 = Math.max(r1, extent.r1)
    c0 = Math.min(c0, extent.c0); c1 = Math.max(c1, extent.c1)
  }
  return { r0, r1, c0, c1 }
}
/** Exact spill match mirroring the store predicate: an extent intersection or
 * the anchor cell inside the query (child-only counts either way). */
function spillMatches(spill: SpillExtent, rect: SourceRect): boolean {
  for (const extent of [spill.saved, spill.evaluated]) {
    if (extent !== undefined && extent.r0 <= rect.r1 && rect.r0 <= extent.r1
      && extent.c0 <= rect.c1 && rect.c0 <= extent.c1) return true
  }
  return spill.anchor.row >= rect.r0 && spill.anchor.row <= rect.r1
    && spill.anchor.col >= rect.c0 && spill.anchor.col <= rect.c1
}
/** Deep clone-then-freeze for plain structural records. The index owns its
 * copies: later caller mutation cannot affect queries, and returned records
 * reject mutation attempts instead of aliasing caller state. */
function frozenClone<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(frozenClone)) as unknown as T
  if (typeof value === 'object' && value !== null) {
    const out: Record<string, unknown> = {}
    for (const [key, entry] of Object.entries(value)) out[key] = frozenClone(entry)
    return Object.freeze(out) as T
  }
  return value
}
/** Immutable index set over one StructuralResult. Each spill is one entry over
 * the union of saved, evaluated and anchor extents, verified exactly on query,
 * so child-only intersections match even when the anchor lies outside the query
 * and cursors stay exactly resumable with no cross-page dedup. Absolute anchors
 * are retained separately and never enter the spatial index. All records are
 * cloned and deep-frozen: the index owns its copies. */
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
    this.stamp = frozenClone(result.stamp)
    // A set error poisons the result even when completeness claims otherwise:
    // errored structural data must wait, never authorize.
    this.completeness = result.error !== undefined ? 'incomplete' : result.completeness
    this.missing = Object.freeze(result.missing.map(rect => frozenClone(rect)))
    const absolute: AbsoluteDrawing[] = []
    const drawingEntries: { rect: SourceRect; value: CellAnchoredDrawing }[] = []
    for (const drawing of result.drawings) {
      const record = frozenClone(drawing)
      if (record.kind === 'absolute') absolute.push(record)
      else drawingEntries.push({ rect: record.bounds, value: record })
    }
    const spillEntries: { rect: SourceRect; value: SpillExtent }[] = []
    for (const spill of result.spills) {
      const record = frozenClone(spill)
      spillEntries.push({ rect: spillUnion(record), value: record })
    }
    const mergeEntries = result.merges.map(merge => {
      const record = frozenClone(merge)
      return { rect: record.range, value: record }
    })
    const tableEntries = result.tables.map(table => {
      const record = frozenClone(table)
      return { rect: record.range, value: record }
    })
    this.merges = RectangleIndex.build(mergeEntries)
    this.tables = RectangleIndex.build(tableEntries)
    this.spillEntries = RectangleIndex.build(spillEntries)
    this.drawings = RectangleIndex.build(drawingEntries)
    this.absoluteDrawings = Object.freeze(absolute)
    this.recordCount = this.merges.size + this.tables.size + this.spillEntries.size + this.drawings.size
  }
  /** First hit across kinds in merge/table/spill/drawing order with early exit.
   * Later kinds are never visited once a hit is found. */
  firstWhere(rect: SourceRect, pred: (hit: StructuralHit) => boolean): StructuralHit | undefined {
    validateSourceRect(rect)
    const merge = this.merges.firstWhere(rect, record => pred({ kind: 'merge', record }))
    if (merge !== undefined) return { kind: 'merge', record: merge }
    const table = this.tables.firstWhere(rect, record => pred({ kind: 'table', record }))
    if (table !== undefined) return { kind: 'table', record: table }
    const spill = this.spillEntries.firstWhere(rect, record => spillMatches(record, rect) && pred({ kind: 'spill', record }))
    if (spill !== undefined) return { kind: 'spill', record: spill }
    const drawing = this.drawings.firstWhere(rect, record => pred({ kind: 'drawing', record }))
    if (drawing !== undefined) return { kind: 'drawing', record: drawing }
    return undefined
  }
  /** All cell-anchored drawings intersecting rect in one walk; the guard Y-checks
   * every drawing, so collection (not paging) is the work here. */
  drawingsIn(rect: SourceRect): readonly CellAnchoredDrawing[] {
    return this.drawings.collect(rect)
  }
  query(rect: SourceRect, options?: StructuralQueryOptions): StructuralPage {
    validateSourceRect(rect)
    const limit = checkLimit(options?.limit, 'structural')
    const state = decodeStructuralCursor(options?.cursor, rect)
    let remaining = limit
    let visits = 0
    const probe = options?.countVisits === undefined ? undefined : (n: number) => { visits += n }
    const take = <T>(index: RectangleIndex<T>, start: number, pred?: (value: T) => boolean) => {
      if (remaining === 0) {
        const empty: T[] = []
        return { items: Object.freeze(empty), resume: start, exhausted: start >= index.size }
      }
      const found = index.walkFrom(rect, start, remaining, { pred, countVisits: probe })
      remaining -= found.items.length
      return found
    }
    const merges = take(this.merges, state.m)
    const tables = take(this.tables, state.t)
    const spills = take(this.spillEntries, state.s, spill => spillMatches(spill, rect))
    const drawings = take(this.drawings, state.d)
    const page: StructuralPage = {
      completeness: this.completeness,
      missing: this.missing,
      merges: merges.items,
      tables: tables.items,
      spills: spills.items,
      drawings: drawings.items,
    }
    if (probe !== undefined && options?.countVisits !== undefined) options.countVisits(visits)
    const exhausted = merges.exhausted && tables.exhausted && spills.exhausted && drawings.exhausted
    return exhausted ? page : {
      ...page,
      nextCursor: `s3struct2:${rectKey(rect)}:${merges.resume}:${tables.resume}:${spills.resume}:${drawings.resume}`,
    }
  }
}
function decodeStructuralCursor(cursor: string | undefined, rect: SourceRect): { m: number; t: number; s: number; d: number } {
  const fresh = { m: 0, t: 0, s: 0, d: 0 }
  if (cursor === undefined) return fresh
  const parts = cursor.split(':')
  if (parts.length !== 6 || parts[0] !== 's3struct2' || parts[1] !== rectKey(rect)) {
    throw new RangeError('Invalid structural cursor')
  }
  const numbers = parts.slice(2).map(Number)
  if (!numbers.every((n, i) => Number.isSafeInteger(n) && n >= 0 && parts[i + 2] === String(n))) {
    throw new RangeError('Invalid structural cursor')
  }
  return { m: numbers[0], t: numbers[1], s: numbers[2], d: numbers[3] }
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
  /** Pages are source-sized; `limit` caps items per page when the source honors
   * it (the store does). Cursors are scoped to one range and one limit: a cursor
   * resumed with a different range or limit is rejected with a RangeError. */
  iterate(range: SourceRect, cursor?: string, limit?: number): Promise<PopulatedPage>
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
    const page = await source.iterate(range, cursor, pageLimit)
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
export interface RowIntervalView {
  readonly band: SourceRowInterval
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  columnsOf(row: number): readonly SourceColInterval[]
}
export interface ColumnIntervalView {
  readonly band: SourceRowInterval
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  rowsOf(col: number): readonly SourceRowInterval[]
}
export class RowChunkIndex implements RowIntervalView {
  private readonly rows: ReadonlyMap<number, readonly SourceColInterval[]>
  private readonly orderedRows: readonly number[]
  readonly band: SourceRowInterval
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  private constructor(rows: ReadonlyMap<number, readonly SourceColInterval[]>, band: SourceRowInterval, complete: boolean, missing: readonly SourceRect[]) {
    this.rows = rows
    this.orderedRows = Object.freeze([...rows.keys()].sort((a, b) => a - b))
    this.band = Object.freeze({ ...band })
    this.complete = complete
    this.missing = missing
  }
  static async build(source: PopulatedCellSource, band: SourceRowInterval, pageLimit = 4096): Promise<ChunkBuildResult<RowChunkIndex>> {
    const consumed = await consumeBand(source, band, pageLimit)
    const rows = new Map<number, readonly SourceColInterval[]>()
    for (const [row, cols] of consumed.byRow) rows.set(row, Object.freeze(toColumnIntervals(cols)))
    return { index: new RowChunkIndex(rows, band, consumed.complete, consumed.missing), complete: consumed.complete, missing: consumed.missing }
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
/** Streams the two outside-column rects `[0,c0-1]` and `[c1+1,16383]` for one
 * populated cell, stopping at the first one. Key/range columns are never read,
 * no per-cell index is retained (O(1) working state), and `pageLimit` is passed
 * to every source request. A found cell is decisive even on a loading page;
 * without one, loading pages report `unknown` instead of a fake `clear`. */
export async function findOutsideColumnData(
  source: PopulatedCellSource, band: SourceRowInterval, c0: number, c1: number, pageLimit = 4096,
): Promise<OutsideColumnAnswer> {
  validateCoordinate(band.r0, 0)
  validateCoordinate(band.r1, 0)
  if (band.r0 > band.r1) throw new RangeError('Reversed chunk band')
  validateCoordinate(0, c0)
  validateCoordinate(0, c1)
  if (c0 > c1) throw new RangeError('Reversed outside-column span')
  if (!Number.isSafeInteger(pageLimit) || pageLimit <= 0) throw new RangeError('Invalid chunk page limit')
  const rects: SourceRect[] = []
  if (c0 > 0) rects.push({ r0: band.r0, r1: band.r1, c0: 0, c1: c0 - 1 })
  if (c1 < MAX_COLS - 1) rects.push({ r0: band.r0, r1: band.r1, c0: c1 + 1, c1: MAX_COLS - 1 })
  const missing: SourceRect[] = []
  let complete = true
  for (const rect of rects) {
    let cursor: string | undefined
    const seen = new Set<string>()
    for (;;) {
      const page = await source.iterate(rect, cursor, pageLimit)
      for (const cell of page.items) {
        validateCoordinate(cell.row, cell.col)
        if (cell.row < rect.r0 || cell.row > rect.r1 || cell.col < rect.c0 || cell.col > rect.c1) {
          throw new RangeError('Populated cell outside the requested range')
        }
        return { kind: 'occupied', at: { row: cell.row, col: cell.col } }
      }
      if (page.completeness !== 'complete') {
        complete = false
        for (const miss of page.missing) missing.push(miss)
      }
      if (page.nextCursor === undefined) break
      if (seen.has(page.nextCursor)) throw new RangeError('Populated cursor made no progress')
      seen.add(page.nextCursor)
      cursor = page.nextCursor
    }
  }
  if (!complete) return { kind: 'unknown', missing: Object.freeze(missing) }
  return { kind: 'clear' }
}
/** Transposed per-column populated-row intervals for the same band. */
export class ColumnChunkIndex implements ColumnIntervalView {
  private readonly cols: ReadonlyMap<number, readonly SourceRowInterval[]>
  readonly band: SourceRowInterval
  readonly complete: boolean
  readonly missing: readonly SourceRect[]
  private constructor(cols: ReadonlyMap<number, readonly SourceRowInterval[]>, band: SourceRowInterval, complete: boolean, missing: readonly SourceRect[]) {
    this.cols = cols
    this.band = Object.freeze({ ...band })
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
    return { index: new ColumnChunkIndex(cols, band, consumed.complete, consumed.missing), complete: consumed.complete, missing: consumed.missing }
  }
  rowsOf(col: number): readonly SourceRowInterval[] {
    validateCoordinate(0, col)
    return this.cols.get(col) ?? []
  }
}
/** Sorted sub-array of column intervals overlapping [lo,hi] by binary search:
 * O(log n + overlaps), never a full scan. Returned slices share no state. */
export function overlappingRange(intervals: readonly SourceColInterval[], lo: number, hi: number): readonly SourceColInterval[] {
  let start = 0, end = intervals.length
  while (start < end) {
    const mid = (start + end) >> 1
    if (intervals[mid].c1 < lo) start = mid + 1
    else end = mid
  }
  let stop = start
  while (stop < intervals.length && intervals[stop].c0 <= hi) stop++
  return intervals.slice(start, stop)
}
/** Row-interval variant of overlappingRange. */
export function overlappingRows(intervals: readonly SourceRowInterval[], lo: number, hi: number): readonly SourceRowInterval[] {
  let start = 0, end = intervals.length
  while (start < end) {
    const mid = (start + end) >> 1
    if (intervals[mid].r1 < lo) start = mid + 1
    else end = mid
  }
  let stop = start
  while (stop < intervals.length && intervals[stop].r0 <= hi) stop++
  return intervals.slice(start, stop)
}
/** Excel current-region semantics: starting from the seed's populated cell,
 * repeatedly grow the bounding box while ANY populated cell exists in the
 * one-cell ring around it, to a fixpoint. The ring is checked sparsely: edge
 * rows through the row index, side columns through the column index, each an
 * O(log n) overlap probe (never a dense scan, never a per-stripe pair scan).
 * A diagonally touching cell joins through the ring (A1+B2 yields A1:B2); blank
 * interior stays inside the box. An unpopulated seed has no region (null).
 * Incomplete indexes, a column index narrower than the row band, or growth that
 * would need rows outside the indexed band yield `unknown`: the caller must
 * widen the band and retry rather than silently truncating (spec section 6.1:
 * unknown pages cannot be boundaries). */
export function currentRegion(
  seed: SourceCell, rows: RowIntervalView, cols: ColumnIntervalView,
): SourceRect | null | { kind: 'unknown'; missing: readonly SourceRect[] } {
  validateCoordinate(seed.row, seed.col)
  if (!rows.complete || !cols.complete) {
    return { kind: 'unknown', missing: Object.freeze([...rows.missing, ...cols.missing]) }
  }
  if (cols.band.r0 > rows.band.r0 || cols.band.r1 < rows.band.r1) {
    return {
      kind: 'unknown',
      missing: Object.freeze([{ r0: rows.band.r0, r1: rows.band.r1, c0: 0, c1: MAX_COLS - 1 }]),
    }
  }
  const inRows = overlappingRange(rows.columnsOf(seed.row), seed.col, seed.col).length > 0
  const inCols = overlappingRows(cols.rowsOf(seed.col), seed.row, seed.row).length > 0
  if (!inRows || !inCols) {
    if (seed.row < rows.band.r0 || seed.row > rows.band.r1) {
      return { kind: 'unknown', missing: Object.freeze([{ r0: seed.row, r1: seed.row, c0: seed.col, c1: seed.col }]) }
    }
    return null
  }
  let box: SourceRect = { r0: seed.row, r1: seed.row, c0: seed.col, c1: seed.col }
  for (;;) {
    let r0 = box.r0, r1 = box.r1, c0 = box.c0, c1 = box.c1
    for (const row of [box.r0 - 1, box.r1 + 1]) {
      if (row < 0 || row >= MAX_ROWS) continue
      if (row < rows.band.r0 || row > rows.band.r1) {
        return { kind: 'unknown', missing: Object.freeze([{ ...box }]) }
      }
      const found = overlappingRange(rows.columnsOf(row), box.c0 - 1, box.c1 + 1)
      if (found.length > 0) {
        r0 = Math.min(r0, row)
        r1 = Math.max(r1, row)
        for (const interval of found) {
          c0 = Math.min(c0, interval.c0)
          c1 = Math.max(c1, interval.c1)
        }
      }
    }
    for (const col of [box.c0 - 1, box.c1 + 1]) {
      if (col < 0 || col >= MAX_COLS) continue
      const found = overlappingRows(cols.rowsOf(col), box.r0, box.r1)
      if (found.length > 0) {
        c0 = Math.min(c0, col)
        c1 = Math.max(c1, col)
        for (const interval of found) {
          r0 = Math.min(r0, interval.r0)
          r1 = Math.max(r1, interval.r1)
        }
      }
    }
    if (r0 === box.r0 && r1 === box.r1 && c0 === box.c0 && c1 === box.c1) return box
    box = { r0, r1, c0, c1 }
  }
}
/** Store-backed populated source so S1 can feed the chunk indexes from
 * scanPopulated. Store cursors pass through opaquely; the store itself rejects
 * foreign cursors with an invalid-cursor error page, which surfaces as a throw.
 * Loading pages report the requested range as missing (store pages carry no
 * missing-rect list); operational failures throw instead of faking a page. */
export function fromStoreScan(store: SourceStore, options: ReadOptions, pageLimit = 4096): PopulatedCellSource {
  if (!Number.isSafeInteger(pageLimit) || pageLimit <= 0) throw new RangeError('Invalid store scan page limit')
  return {
    async iterate(range: SourceRect, cursor?: string, limit?: number): Promise<PopulatedPage> {
      validateSourceRect(range)
      if (cursor !== undefined && (typeof cursor !== 'string' || cursor.length === 0)) {
        throw new RangeError('Invalid populated cursor')
      }
      const effective = limit ?? pageLimit
      if (!Number.isSafeInteger(effective) || effective <= 0) throw new RangeError('Invalid store scan page limit')
      const request: PopulatedQuery = cursor === undefined
        ? { range, limit: effective }
        : { range, limit: effective, cursor }
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
