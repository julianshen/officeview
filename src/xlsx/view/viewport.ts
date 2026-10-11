/**
 * Full-coordinate renderer seam (S2.3, spec §§5.2, 5.3, 5.4, 11.1, 11.5).
 *
 * The interactive path never reads the legacy capped `rowY`/`colX` prefix
 * arrays: geometry is the S2.1 full-capacity axis/projection model, a viewport
 * plan locates its display window first, admits detail from bounded metadata
 * (window counts, projected device size, per-row column indexes, merge and
 * drawing indexes) and only then resolves source cells. Every source coordinate
 * is painted at its own logical offset (or its sorted display slot) or not at
 * all; no missing geometry falls back to the origin. Typed `AxisMetricsError`
 * failures are mapped to diagnosed fallbacks here so they never reach a caller
 * as an unhandled exception.
 */
import type { XlsxCell, XlsxDrawing, XlsxMergeRange, XlsxSheet } from '../types'
import type { DrawingContentShape } from '../../drawing/content'
import type {
  AxisMetrics, AxisRun, AxisSizeInterval, AxisWindow, CacheIdentity, CompactSourceRowSet, CompleteViewportContent, Completeness, IntervalSet, LogicalCellBox,
  LogicalViewport, OverviewDomain, OverviewGroup, RowOrder, RowPermutationBand, RowVisibility, SourceFailure, SourceStore, SparseRowPage,
  ViewportCell, ViewportOverviewGroup, ViewportResult, WorkerDiagnostic,
} from './contracts'
import { MAX_COLS, MAX_ROWS, isValidCoordinate, type SourceCell } from './coordinates'
import { AxisMetricsError, createAxisMetrics, type AxisMetricsErrorCode, type SparseAxisMetrics } from './axes'
import { createSheetMetrics, type ProjectedSheetMetrics } from './projection'
import { RectangleIndex } from './structure'
import { cellSourceRow, drawingExtent } from './sheetCells'
import { createIntervalSet } from './visibility'

// Static-renderer helpers live in a leaf module so the eager static path never
// reaches the viewport planner; re-exported for existing callers.
export { cellSourceRow, drawingExtent }

/** Excel default geometry in CSS pixels; identical to the static renderer so an
 * ordinary small sheet lays out the same in both paths. */
export const SHEET_DEFAULT_ROW_CSS = 20
export const SHEET_DEFAULT_COL_CSS = 64
const CHAR_CSS = 7
const COL_PADDING_CSS = 5
/** §11.1: individual detail needs projected width and height of at least four device pixels. */
export const MIN_DETAIL_DEVICE_PX = 4
/** Float tolerance for that rule: 0.4 CSS px x 10 computes 3.9999999999999996. */
const DETAIL_EPSILON = 1e-9
function belowDetail(sizeCss: number, scale: number): boolean {
  return sizeCss * scale < MIN_DETAIL_DEVICE_PX - DETAIL_EPSILON
}
/** How far (logical CSS px) left/right of the pane unclipped cell text is fetched
 * as a layout dependency. The static renderer never clips ordinary text, so a crop
 * of it shows overflow from neighbouring cells; the viewer reproduces that within
 * this bounded reach (the cells count against the record budget). */
export const TEXT_OVERFLOW_REACH_CSS = 2048

/** Stable diagnostic code per typed axis failure (spec §5.3, plan S2.3). */
export const AXIS_ERROR_DIAGNOSTICS: Readonly<Record<AxisMetricsErrorCode, string>> = Object.freeze({
  CONTRADICTORY_OVERRIDES: 'axis-contradictory-overrides',
  PRECISION_EXCEEDED: 'axis-precision-exceeded',
  TOTAL_NOT_FINITE: 'axis-total-not-finite',
  INVALID_METADATA: 'axis-invalid-metadata',
  INVALID_QUERY: 'axis-invalid-query',
})

/** Proposed §11.1 limits on desktop (D) and mobile (M): per frame, individual
 * records (cells plus drawings), source page requests and gridline instances; per
 * 512-device-pixel tile, records and page requests. `overviewGroups` is the §4.1
 * bounded `queryOverview` output. Beyond any of them the affected region is coarse
 * before cell content is read. */
export const VIEWPORT_DETAIL_LIMITS = Object.freeze({
  D: Object.freeze({ records: 20_000, gridlines: 4096, pages: 32, overviewGroups: 2048, tile: Object.freeze({ records: 5000, pages: 8 }) }),
  M: Object.freeze({ records: 10_000, gridlines: 2048, pages: 16, overviewGroups: 1024, tile: Object.freeze({ records: 2500, pages: 4 }) }),
})
interface DetailLimits { records: number; gridlines: number; pages: number; overviewGroups: number; tile: { records: number; pages: number } }
/** Tile edge for per-tile admission (spec §11.1 proposal; `tileCache.TILE_DEVICE_PX`). */
const TILE_DEVICE_PX = 512
export const COARSE_NOTICE = 'Zoom in or inspect a cell for detail'
export type ViewportProfile = keyof typeof VIEWPORT_DETAIL_LIMITS

/** Collapse an index only when every record for it is hidden: duplicate row
 * records are invalid OOXML and a stray hidden duplicate must not hide data. */
function hiddenRowIndexes(sheet: XlsxSheet): { hidden: number[]; conflicting: number } {
  const all = new Map<number, boolean>()
  for (const row of sheet.rows) {
    if (isValidCoordinate(row.index, 0)) all.set(row.index, (all.get(row.index) ?? true) && row.hidden === true)
  }
  const hidden = [...all].filter(([, value]) => value).map(([index]) => index)
  const conflicting = new Set(sheet.rows.filter(row => row.hidden === true && all.get(row.index) === false).map(row => row.index)).size
  return { hidden, conflicting }
}

/** The authored-hidden snapshot (`SheetView.baseHidden`, spec §5.1/§5.2). Feed it to
 * `visibilityFromHidden`/`projectVisibility` to build the canonical display mask. */
export function sheetBaseHidden(sheet: XlsxSheet): IntervalSet {
  return createIntervalSet(hiddenRowIndexes(sheet).hidden.map(index => ({ r0: index, r1: index })))
}

/** Nodes one top-level drawing paints (group children and diagram shapes included):
 * its record weight against the per-frame budget. */
function drawingWeight(node: XlsxDrawing | DrawingContentShape): number {
  let weight = 1
  for (const child of (node.children ?? []) as Array<XlsxDrawing | DrawingContentShape>) weight += drawingWeight(child)
  if ('content' in node && node.content?.kind === 'diagram') for (const shape of node.content.shapes) weight += drawingWeight(shape as DrawingContentShape)
  return weight
}

interface DrawingBox { left: number; top: number; right: number; bottom: number; index: number; weight: number }
/** Float64 box index over drawing extents: top-sorted entries in a segment tree whose
 * nodes carry max-bottom and the x-span of their subtree, so typical frames visit
 * O(log n + hits) nodes. Every query is visit- and weight-budgeted: an adversarial
 * layout that defeats pruning ends the query (coarse frame) instead of a full scan. */
class BoxIndex {
  private readonly maxBottom: Float64Array
  private readonly minLeft: Float64Array
  private readonly maxRight: Float64Array
  private constructor(private readonly entries: readonly DrawingBox[]) {
    const size = Math.max(1, 4 * entries.length)
    this.maxBottom = new Float64Array(size).fill(-Infinity)
    this.minLeft = new Float64Array(size).fill(Infinity)
    this.maxRight = new Float64Array(size).fill(-Infinity)
    const build = (node: number, lo: number, hi: number): void => {
      if (lo + 1 === hi) {
        const entry = entries[lo]
        this.maxBottom[node] = entry.bottom; this.minLeft[node] = entry.left; this.maxRight[node] = entry.right
        return
      }
      const mid = Math.floor((lo + hi) / 2)
      build(node * 2, lo, mid)
      build(node * 2 + 1, mid, hi)
      this.maxBottom[node] = Math.max(this.maxBottom[node * 2], this.maxBottom[node * 2 + 1])
      this.minLeft[node] = Math.min(this.minLeft[node * 2], this.minLeft[node * 2 + 1])
      this.maxRight[node] = Math.max(this.maxRight[node * 2], this.maxRight[node * 2 + 1])
    }
    if (entries.length) build(1, 0, entries.length)
  }
  static build(boxes: readonly DrawingBox[]): BoxIndex {
    return new BoxIndex([...boxes].sort((a, b) => a.top - b.top || a.index - b.index))
  }
  /** Original indexes of boxes touching the query, ascending, or null once the hits'
   * weight exceeds `budget` or the walk exceeds its visit allowance. */
  query(left: number, top: number, right: number, bottom: number, budget: number): number[] | null {
    const hits: number[] = []
    let weight = 0, visits = 0
    const maxVisits = 2 * Math.max(0, budget) + 64
    let hi = 0, upper = this.entries.length
    while (hi < upper) {
      const mid = Math.floor((hi + upper) / 2)
      if (this.entries[mid].top <= bottom) hi = mid + 1
      else upper = mid
    }
    let exhausted = false
    const visit = (node: number, lo: number, end: number): void => {
      if (exhausted || lo >= hi || this.maxBottom[node] < top || this.minLeft[node] > right || this.maxRight[node] < left) return
      if (++visits > maxVisits) { exhausted = true; return }
      if (lo + 1 === end) {
        const entry = this.entries[lo]
        if (entry.right >= left && entry.left <= right) {
          hits.push(entry.index)
          weight += entry.weight
          if (weight > budget) exhausted = true
        }
        return
      }
      const mid = Math.floor((lo + end) / 2)
      visit(node * 2, lo, mid)
      visit(node * 2 + 1, mid, end)
    }
    if (this.entries.length) visit(1, 0, this.entries.length)
    return exhausted ? null : hits.sort((a, b) => a - b)
  }
}

export interface SheetGeometryOptions {
  order?: RowOrder
  /** Canonical display mask (spec §5.2). When supplied it alone decides which rows
   * collapse, so a filter may reveal a snapshot-hidden body row; authored row
   * heights are kept. Without it the authored-hidden snapshot collapses. */
  visibility?: RowVisibility
}
/** Built once per geometry generation (sheet + order + visibility), never per paint. */
export interface SheetGeometry {
  readonly kind: 'ready'
  readonly metrics: ProjectedSheetMetrics
  /** Fallbacks taken while building (dropped sizes, conflicting/invalid metadata). */
  readonly diagnostics: readonly WorkerDiagnostic[]
  /** Populated cells of one source row, ascending by column; undefined for blank rows. */
  cellsOf(sourceRow: number): readonly XlsxCell[] | undefined
  readonly merges: RectangleIndex<XlsxMergeRange>
  /** Top-level drawings touching a logical CSS box, in authored paint order, or null
   * when their node weight or the index walk exceeds `budget` records. */
  drawingsIn(left: number, top: number, right: number, bottom: number, budget: number): readonly XlsxDrawing[] | null
}
export type SheetGeometryResult =
  | SheetGeometry
  | { readonly kind: 'diagnostic'; readonly diagnostics: readonly WorkerDiagnostic[] }

/** Logical CSS box of one painted record; merge anchors span their whole range. */
export interface CellPlacement { cell: XlsxCell; source: SourceCell; x: number; y: number; width: number; height: number; merge?: XlsxMergeRange }
export interface MergePlacement { range: XlsxMergeRange; x: number; y: number; width: number; height: number }
export interface ViewportDetail {
  kind: 'detail'
  /** Row-major (display slot, column) paint order. */
  cells: readonly CellPlacement[]
  merges: readonly MergePlacement[]
  /** Drawings intersecting the pane, in authored order. */
  drawings: readonly XlsxDrawing[]
  /** Distinct logical CSS positions of the visible grid boundaries. */
  rowLines: readonly number[]
  colLines: readonly number[]
  /** Full logical sheet extent (the worksheet continues past the used range). */
  extent: { width: number; height: number }
  diagnostics: readonly WorkerDiagnostic[]
  /** Sub-detail regions of a mixed pane, summarized instead of enumerated. */
  coarse: readonly CoarseGroup[]
  /** The one exact active-cell record, admitted within the budget in any mode. */
  active?: CellPlacement
  work: ViewportWork
}

/** Instrumented per-frame work (plan S5.1): every count is bounded by the caps. */
export interface ViewportWork {
  /** Source page requests (store reads) issued for the frame. */
  pages: number
  /** Cell index entries visited during admission and enumeration. */
  visitedRecords: number
  /** Detail records admitted (cells plus drawing nodes). */
  records: number
  layoutsBuilt: number
  layoutsReused: number
  /** Distinct boundary instances, and painted gridline segments after merge masking. */
  gridlines: number
  gridlineSpans: number
  /** Axis runs visited while planning (never per coordinate of a uniform run). */
  runsVisited: number
  overviewGroups: number
  overviewPages: number
  /** Exact semantic/inspector records (the active cell). */
  semanticNodes: number
  mergeQueries: number
}
export function emptyWork(): ViewportWork {
  return { pages: 0, visitedRecords: 0, records: 0, layoutsBuilt: 0, layoutsReused: 0, gridlines: 0, gridlineSpans: 0, runsVisited: 0, overviewGroups: 0, overviewPages: 0, semanticNodes: 0, mergeQueries: 0 }
}
/** One bounded display pixel group: inclusive display slots/columns whose boxes
 * intersect the group's logical box (a row straddling a bucket edge is summarized
 * in both adjacent groups), never a source bounding rectangle. */
export interface CoarseGroup { groupId: string; bounds: LogicalCellBox; slots: { start: number; end: number }; cols: { c0: number; c1: number } }
/** Whole-frame coarse: every visible region is summarized; `groups` are the
 * bounded display pixel groups (no cell content was read). */
export interface ViewportCoarse {
  kind: 'coarse'
  diagnostics: readonly WorkerDiagnostic[]
  groups: readonly CoarseGroup[]
  extent: { width: number; height: number }
  active?: CellPlacement
  work: ViewportWork
}
export type ViewportPlan =
  | ViewportDetail
  | ViewportCoarse
  | { kind: 'pending'; diagnostics: readonly WorkerDiagnostic[] }
  | { kind: 'diagnostic'; diagnostics: readonly WorkerDiagnostic[] }

export function axisDiagnostic(error: AxisMetricsError, axis: 'rows' | 'cols' | 'projection' | 'viewport'): WorkerDiagnostic {
  return { code: AXIS_ERROR_DIAGNOSTICS[error.code], message: `${axis}: ${error.message}` }
}

function validSize(value: number | undefined): value is number {
  return value !== undefined && Number.isFinite(value) && value >= 0
}
/** Static rounding for ordinary sizes (layout parity), but a positive authored size
 * never rounds away: it keeps its exact value and the detail rule decides. */
function cssSize(exact: number): number {
  const rounded = Math.round(exact)
  return rounded > 0 || exact === 0 ? rounded : exact
}

/** Try the authored axis, then without authored sizes, then plain defaults. Each
 * rejected attempt leaves its typed diagnostic; hidden intervals survive the
 * first fallback because they never contradict (all zero). */
function buildAxis(axis: 'rows' | 'cols', logicalCount: number, defaultSizeCss: number, overrides: AxisSizeInterval[], collapsed: Array<{ start: number; end: number }>, diagnostics: WorkerDiagnostic[]): SparseAxisMetrics | null {
  const attempts = [{ overrides, collapsed }, { overrides: [], collapsed }, { overrides: [], collapsed: [] }]
  for (const attempt of attempts) {
    try {
      return createAxisMetrics({ logicalCount, defaultSizeCss, ...attempt })
    } catch (error) {
      if (!(error instanceof AxisMetricsError)) throw error
      diagnostics.push(axisDiagnostic(error, axis))
    }
  }
  return null
}

/** Build the full-capacity geometry for a small object sheet. Authored-hidden rows
 * and columns collapse unconditionally (spec §5.3), with or without drawings,
 * unless a canonical visibility mask is supplied for rows. */
export function buildSheetGeometry(sheet: XlsxSheet, options: SheetGeometryOptions = {}): SheetGeometryResult {
  const diagnostics: WorkerDiagnostic[] = []
  const cellsByRow = new Map<number, XlsxCell[]>()
  const rowOverrides: AxisSizeInterval[] = []
  let skippedRows = 0, skippedCells = 0, inconsistent = 0
  for (const row of sheet.rows) {
    const rowValid = isValidCoordinate(row.index, 0)
    if (!rowValid) skippedRows++
    else if (row.customHeight && row.heightPt !== undefined) {
      if (validSize(row.heightPt)) rowOverrides.push({ start: row.index, end: row.index, sizeCss: cssSize(row.heightPt * (96 / 72)) })
      else diagnostics.push({ code: 'invalid-dimension', message: `row ${row.index + 1}: ignored height ${row.heightPt}` })
    }
    for (const item of row.cells) {
      const source = cellSourceRow(row, item)
      if (!isValidCoordinate(source, item.col)) { skippedCells++; continue }
      if (source !== row.index) inconsistent++
      let cells = cellsByRow.get(source)
      if (!cells) cellsByRow.set(source, cells = [])
      cells.push(item)
    }
  }
  // Stable: equal columns keep authored order, exactly as the static loop paints them.
  for (const cells of cellsByRow.values()) cells.sort((a, b) => a.col - b.col)
  const { hidden, conflicting } = hiddenRowIndexes(sheet)
  const rowCollapsed = options.visibility ? [] : hidden.map(index => ({ start: index, end: index }))
  if (conflicting) diagnostics.push({ code: 'conflicting-row-metadata', message: `${conflicting} row index(es) carry both hidden and visible records; kept visible` })
  if (inconsistent) diagnostics.push({ code: 'inconsistent-row-coordinate', message: `${inconsistent} cell(s) carry a reference outside their row record; placed by the cell reference` })
  if (skippedRows || skippedCells) diagnostics.push({ code: 'coordinate-outside-domain', message: `skipped ${skippedRows} row record(s) and ${skippedCells} cell(s) outside the worksheet domain` })

  const colOverrides: AxisSizeInterval[] = []
  const colCollapsed: Array<{ start: number; end: number }> = []
  for (const spec of sheet.cols) {
    if (!isValidCoordinate(0, spec.min) || !isValidCoordinate(0, spec.max) || spec.min > spec.max) {
      diagnostics.push({ code: 'coordinate-outside-domain', message: `column spec ${spec.min + 1}:${spec.max + 1} ignored` })
      continue
    }
    if (spec.hidden) colCollapsed.push({ start: spec.min, end: spec.max })
    else if (spec.widthChars !== undefined) {
      if (validSize(spec.widthChars)) colOverrides.push({ start: spec.min, end: spec.max, sizeCss: cssSize(spec.widthChars * CHAR_CSS + COL_PADDING_CSS) })
      else diagnostics.push({ code: 'invalid-dimension', message: `column spec ${spec.min + 1}:${spec.max + 1}: ignored width ${spec.widthChars}` })
    }
  }

  const rows = buildAxis('rows', MAX_ROWS, SHEET_DEFAULT_ROW_CSS, rowOverrides, rowCollapsed, diagnostics)
  const cols = buildAxis('cols', MAX_COLS, SHEET_DEFAULT_COL_CSS, colOverrides, colCollapsed, diagnostics)
  if (!rows || !cols) return { kind: 'diagnostic', diagnostics }
  let metrics: ProjectedSheetMetrics
  try {
    metrics = createSheetMetrics({ rows, cols, order: options.order, visibility: options.visibility })
  } catch (error) {
    // A projection that cannot be built must not silently show another order.
    if (!(error instanceof AxisMetricsError)) throw error
    diagnostics.push(axisDiagnostic(error, 'projection'))
    return { kind: 'diagnostic', diagnostics }
  }
  const mergeEntries: Array<{ rect: { r0: number; r1: number; c0: number; c1: number }; value: XlsxMergeRange }> = []
  for (const range of sheet.mergeRanges ?? []) {
    if (!isValidCoordinate(range.minRow, range.minCol) || !isValidCoordinate(range.maxRow, range.maxCol)
      || range.minRow > range.maxRow || range.minCol > range.maxCol) {
      diagnostics.push({ code: 'coordinate-outside-domain', message: 'merge range outside the worksheet domain ignored' })
      continue
    }
    mergeEntries.push({ rect: { r0: range.minRow, r1: range.maxRow, c0: range.minCol, c1: range.maxCol }, value: range })
  }
  const drawings = sheet.drawings ?? []
  const boxes = drawings.map((node, index) => ({ ...drawingExtent([node]), index, weight: drawingWeight(node) }))
    .filter(box => box.left <= box.right && box.top <= box.bottom)
  const drawingIndex = BoxIndex.build(boxes)
  return Object.freeze({
    kind: 'ready' as const,
    metrics,
    diagnostics: Object.freeze(diagnostics),
    cellsOf: (sourceRow: number) => cellsByRow.get(sourceRow),
    merges: RectangleIndex.build(mergeEntries),
    drawingsIn(left: number, top: number, right: number, bottom: number, budget: number) {
      const hits = drawingIndex.query(left, top, right, bottom, budget)
      return hits && hits.map(index => drawings[index])
    },
  })
}

interface LocatedMerge { minRow: number; maxRow: number; range: XlsxMergeRange }
/** Point location over one frame's covering merges. A segment tree over the frame's
 * column span stores each merge at O(log C) canonical nodes; each node's list is
 * sorted by first row with a prefix maximum of last rows, so a lookup walks one
 * root-to-leaf path with a binary search per node: O(log C log M) per cell instead
 * of a scan of every merge. Field reads happen once per merge at build. */
class MergeLocator {
  private readonly nodes = new Map<number, { entries: LocatedMerge[]; reach: Float64Array }>()
  constructor(ranges: readonly XlsxMergeRange[], private readonly c0: number, private readonly c1: number) {
    const lists = new Map<number, LocatedMerge[]>()
    const insert = (node: number, lo: number, hi: number, a: number, b: number, entry: LocatedMerge): void => {
      if (b < lo || a > hi) return
      if (a <= lo && hi <= b) {
        let list = lists.get(node)
        if (!list) lists.set(node, list = [])
        list.push(entry)
        return
      }
      const mid = Math.floor((lo + hi) / 2)
      insert(node * 2, lo, mid, a, b, entry)
      insert(node * 2 + 1, mid + 1, hi, a, b, entry)
    }
    for (const range of ranges) {
      const { minRow, maxRow, minCol, maxCol } = range
      if (c1 >= c0) insert(1, c0, c1, Math.max(minCol, c0), Math.min(maxCol, c1), { minRow, maxRow, range })
    }
    for (const [node, entries] of lists) {
      entries.sort((a, b) => a.minRow - b.minRow)
      const reach = new Float64Array(entries.length)
      entries.forEach((entry, index) => { reach[index] = Math.max(index ? reach[index - 1] : -1, entry.maxRow) })
      this.nodes.set(node, { entries, reach })
    }
  }
  at(row: number, col: number): XlsxMergeRange | undefined {
    if (col < this.c0 || col > this.c1) return undefined
    let node = 1, lo = this.c0, hi = this.c1
    for (;;) {
      const list = this.nodes.get(node)
      if (list) {
        let low = 0, high = list.entries.length
        while (low < high) {
          const mid = Math.floor((low + high) / 2)
          if (list.entries[mid].minRow <= row) low = mid + 1
          else high = mid
        }
        // Valid merges are disjoint, so this normally checks one entry; overlapping
        // (invalid) input stays correct by walking back while the prefix reach allows.
        for (let index = low - 1; index >= 0 && list.reach[index] >= row; index--) {
          if (list.entries[index].maxRow >= row) return list.entries[index].range
        }
      }
      if (lo === hi) return undefined
      const mid = Math.floor((lo + hi) / 2)
      if (col <= mid) { node = node * 2; hi = mid } else { node = node * 2 + 1; lo = mid + 1 }
    }
  }
}

function firstCellAtOrAfter(cells: readonly XlsxCell[], col: number): number {
  let low = 0, high = cells.length
  while (low < high) {
    const mid = Math.floor((low + high) / 2)
    if (cells[mid].col < col) low = mid + 1
    else high = mid
  }
  return low
}

export interface PlanOptions {
  /** One exact detail record kept within budget in every mode (spec §11.1). */
  activeCell?: SourceCell
  /** Merge-query cache reused across frames of one geometry generation. */
  merges?: MergeQueryCache
  /** Remaining §11.1 PER-FRAME budget shared by every tile of one presented frame
   * (plan S5.2). Each detail tile is admitted against what is left and charges its
   * records and gridline instances; once spent, later tiles of the frame go coarse
   * before any cell content is read. */
  frameBudget?: FrameBudget
}
export interface FrameBudget { records: number; gridlines: number }
export function frameBudgetFor(profile: ViewportProfile = 'D'): FrameBudget {
  const limits = VIEWPORT_DETAIL_LIMITS[profile]
  return { records: limits.records, gridlines: limits.gridlines }
}

/** Plan one frame. Admission happens before any cell is read: window counts
 * (O(log n)), projected device size per run, distinct boundary instances,
 * per-row column-index counts, merge/drawing index queries and per-tile counts.
 * Sub-detail runs become bounded coarse pixel groups while the rest of the pane
 * keeps detail; a density/cap crossing makes the whole frame coarse. Source rows
 * outside the window are never scanned. */
export function planViewport(geometry: SheetGeometryResult | undefined, viewport: LogicalViewport, profile: ViewportProfile = 'D', options: PlanOptions = {}): ViewportPlan {
  if (!geometry) return { kind: 'pending', diagnostics: [{ code: 'geometry-pending', message: 'sheet geometry has not been published' }] }
  if (geometry.kind !== 'ready') return { kind: 'diagnostic', diagnostics: geometry.diagnostics }
  try {
    const base = VIEWPORT_DETAIL_LIMITS[profile], budget = options.frameBudget
    // A budgeted plan is one tile of a frame: every record it builds, overflow
    // dependencies included, also obeys the per-tile cap.
    const limits = budget ? { ...base, records: Math.max(0, Math.min(base.records, base.tile.records, budget.records)), gridlines: Math.max(0, Math.min(base.gridlines, budget.gridlines)) } : base
    const plan = detailPlan(geometry, viewport, limits, options)
    if (budget && plan.kind === 'detail') {
      budget.records -= plan.work.records
      budget.gridlines -= plan.work.gridlines
    }
    return plan
  } catch (error) {
    if (!(error instanceof AxisMetricsError)) throw error
    return { kind: 'diagnostic', diagnostics: [...geometry.diagnostics, axisDiagnostic(error, 'viewport')] }
  }
}

/** Projected detail scale (spec §5.5): validated SheetZoom x DPR. */
function detailScale(viewport: LogicalViewport): number {
  const { width, height } = viewport.box
  if (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0) {
    throw new AxisMetricsError('INVALID_QUERY', 'Viewport size must be finite and nonnegative')
  }
  const scale = viewport.zoom * viewport.deviceScale
  if (!Number.isFinite(viewport.zoom) || viewport.zoom <= 0 || !Number.isFinite(viewport.deviceScale) || viewport.deviceScale <= 0 || !Number.isFinite(scale)) {
    throw new AxisMetricsError('INVALID_QUERY', 'Viewport zoom and device scale must be finite and positive')
  }
  return scale
}

interface CssSpan { start: number; end: number }
interface AxisPlan {
  /** Detail indexes, ascending. */
  detail: number[]
  /** Boundary positions of the detail runs, coalesced by device pixel. */
  lines: number[]
  /** Contiguous logical spans of sub-detail runs, and of detail runs. */
  coarse: CssSpan[]
  detailSpans: CssSpan[]
}
/** Classify one axis window run by run. Uniform runs are classified once; only
 * detail runs enumerate their indexes (bounded by the gridline budget). Returns
 * null when the window needs more runs or boundaries than the budget allows. */
function planAxis(window: AxisWindow, start: (index: number) => number, scale: number, gridlines: number, work: ViewportWork): AxisPlan | null {
  const plan: AxisPlan = { detail: [], lines: [], coarse: [], detailSpans: [] }
  const runBudget = 2 * gridlines + 64
  let runs = 0
  const extend = (spans: CssSpan[], from: number, to: number) => {
    const last = spans[spans.length - 1]
    if (last && last.end === from) last.end = to
    else spans.push({ start: from, end: to })
  }
  const push = (value: number) => {
    const last = plan.lines[plan.lines.length - 1]
    // Coalesce coincident boundaries (zero-extent indexes between runs share one)
    // and boundaries that land on the same device pixel.
    if (last === undefined || (last !== value && Math.round(last * scale) !== Math.round(value * scale))) plan.lines.push(value)
  }
  for (const run of window.runs) {
    work.runsVisited++
    if (++runs > runBudget) return null
    const from = run.startCss, to = run.startCss + (run.end - run.start + 1) * run.sizeCss
    if (belowDetail(run.sizeCss, scale)) { extend(plan.coarse, from, to); continue }
    extend(plan.detailSpans, from, to)
    for (let index = run.start; index <= run.end; index++) {
      push(start(index))
      plan.detail.push(index)
      if (plan.lines.length > gridlines) return null
    }
    push(start(run.end + 1))
  }
  return plan
}

/** Bucket logical spans into device-pixel groups of at least four device pixels,
 * doubling the bucket until the whole frame fits the overview output bound. */
function pixelGroups(regions: ReadonlyArray<{ rows: CssSpan; cols: CssSpan }>, scale: number, maxGroups: number): Array<{ rows: CssSpan; cols: CssSpan }> {
  let size = MIN_DETAIL_DEVICE_PX / scale
  const count = (s: number) => regions.reduce((sum, region) => sum + Math.ceil((region.rows.end - region.rows.start) / s - 1e-9) * Math.ceil((region.cols.end - region.cols.start) / s - 1e-9), 0)
  while (count(size) > maxGroups) size *= 2
  const cut = (span: CssSpan): CssSpan[] => {
    const out: CssSpan[] = []
    for (let at = span.start; at < span.end - 1e-9; at += size) out.push({ start: at, end: Math.min(span.end, at + size) })
    return out
  }
  const groups: Array<{ rows: CssSpan; cols: CssSpan }> = []
  for (const region of regions) for (const rows of cut(region.rows)) for (const cols of cut(region.cols)) groups.push({ rows, cols })
  return groups
}

/** Inclusive indexes whose boxes intersect [start,end) on an axis, or null. */
function indexesIn(axis: AxisMetrics, span: CssSpan): { start: number; end: number } | null {
  const first = axis.locate(span.start)
  if (first === null) return null
  const located = axis.locate(span.end)
  let last = located === null ? axis.logicalCount - 1 : located
  if (located !== null && axis.start(located) >= span.end) last = located - 1
  return last >= first ? { start: first, end: last } : null
}

function coarseGroups(metrics: ProjectedSheetMetrics, regions: ReadonlyArray<{ rows: CssSpan; cols: CssSpan }>, scale: number, maxGroups: number): CoarseGroup[] {
  const groups: CoarseGroup[] = []
  for (const group of pixelGroups(regions.filter(region => region.rows.end > region.rows.start && region.cols.end > region.cols.start), scale, maxGroups)) {
    const slots = indexesIn(metrics.rows, group.rows), cols = indexesIn(metrics.cols, group.cols)
    if (!slots || !cols) continue
    groups.push({
      groupId: `g${groups.length}`,
      bounds: { x: group.cols.start, y: group.rows.start, width: group.cols.end - group.cols.start, height: group.rows.end - group.rows.start },
      slots, cols: { c0: cols.start, c1: cols.end },
    })
  }
  return groups
}

/** Window-level planning shared by the object and store paths. */
type WindowPlan =
  | { kind: 'detail'; scale: number; rows: AxisPlan; cols: AxisPlan; coarse: CoarseGroup[]; extent: { width: number; height: number }; work: ViewportWork }
  | { kind: 'coarse'; scale: number; message: string; groups: CoarseGroup[]; extent: { width: number; height: number }; work: ViewportWork }
function planWindow(metrics: ProjectedSheetMetrics, viewport: LogicalViewport, limits: DetailLimits): WindowPlan {
  const scale = detailScale(viewport)
  const { x, y, width, height } = viewport.box
  const work = emptyWork()
  const extent = { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
  const paneRows = { start: Math.max(0, y), end: Math.min(extent.height, y + height) }
  const paneCols = { start: Math.max(0, x), end: Math.min(extent.width, x + width) }
  const wholeCoarse = (message: string): WindowPlan => ({ kind: 'coarse', scale, message, extent, work, groups: coarseGroups(metrics, [{ rows: paneRows, cols: paneCols }], scale, limits.overviewGroups) })
  const rowWindow = metrics.rows.visibleBetween(y, y + height)
  const colWindow = metrics.cols.visibleBetween(x, x + width)
  // Necessary bound from O(log n) counts: each visible index owns one boundary, plus one closing each axis.
  if (rowWindow.count + colWindow.count + 2 > limits.gridlines) {
    return wholeCoarse(`${rowWindow.count} rows x ${colWindow.count} columns exceed ${limits.gridlines} gridline instances`)
  }
  const rows = planAxis(rowWindow, index => metrics.rows.start(index), scale, limits.gridlines, work)
  const cols = rows && planAxis(colWindow, index => metrics.cols.start(index), scale, limits.gridlines, work)
  if (!rows || !cols || rows.lines.length + cols.lines.length > limits.gridlines) return wholeCoarse(`boundary instances exceed ${limits.gridlines}`)
  if (!rows.detail.length && rows.coarse.length || !cols.detail.length && cols.coarse.length) {
    return wholeCoarse(`visible rows/columns project below ${MIN_DETAIL_DEVICE_PX} device pixels at zoom ${viewport.zoom} x DPR ${viewport.deviceScale}`)
  }
  work.gridlines = rows.lines.length + cols.lines.length
  // Mixed pane: sub-detail row spans across the pane, sub-detail column spans
  // beside each detail row block. Detail cells are only detail rows x detail columns.
  const regions: Array<{ rows: CssSpan; cols: CssSpan }> = []
  for (const span of rows.coarse) regions.push({ rows: clip(span, paneRows), cols: paneCols })
  for (const block of rows.detailSpans) for (const span of cols.coarse) regions.push({ rows: clip(block, paneRows), cols: clip(span, paneCols) })
  return { kind: 'detail', scale, rows, cols, extent, work, coarse: regions.length ? coarseGroups(metrics, regions, scale, limits.overviewGroups) : [] }
}
function clip(span: CssSpan, to: CssSpan): CssSpan {
  return { start: Math.max(span.start, to.start), end: Math.min(span.end, to.end) }
}

/** Per-tile record admission (spec §11.1): count each record once per 512
 * device-pixel tile it touches inside the pane. Tiles align to the logical sheet. */
function tileOverflow(boxes: Iterable<{ x: number; y: number; width: number; height: number }>, viewport: LogicalViewport, scale: number, cap: number): boolean {
  const { x, y, width, height } = viewport.box
  const tile = TILE_DEVICE_PX / scale
  const counts = new Map<string, number>()
  for (const box of boxes) {
    const left = Math.max(box.x, x), right = Math.min(box.x + box.width, x + width)
    const top = Math.max(box.y, y), bottom = Math.min(box.y + box.height, y + height)
    if (!(right > left && bottom > top)) continue
    for (let ty = Math.floor(top / tile); ty * tile < bottom; ty++) {
      for (let tx = Math.floor(left / tile); tx * tile < right; tx++) {
        const key = `${tx},${ty}`
        const next = (counts.get(key) ?? 0) + 1
        if (next > cap) return true
        counts.set(key, next)
      }
    }
  }
  return false
}

/** Merge index queries reused across frames of one geometry generation (plan
 * S5.1). Keys are the query rectangle and limit; the geometry object is the
 * generation (a new sheet/order/visibility builds a new geometry). */
export class MergeQueryCache {
  private readonly byGeometry = new WeakMap<SheetGeometry, Map<string, ReturnType<SheetGeometry['merges']['query']>>>()
  hits = 0
  misses = 0
  constructor(readonly maxEntries = 64) {}
  query(geometry: SheetGeometry, rect: { r0: number; r1: number; c0: number; c1: number }, limit: number, work: ViewportWork): ReturnType<SheetGeometry['merges']['query']> {
    let entries = this.byGeometry.get(geometry)
    if (!entries) this.byGeometry.set(geometry, entries = new Map())
    const key = `${rect.r0},${rect.r1},${rect.c0},${rect.c1},${limit}`
    const hit = entries.get(key)
    if (hit) { this.hits++; return hit }
    this.misses++
    work.mergeQueries++
    const page = geometry.merges.query(rect, { limit })
    entries.set(key, page)
    if (entries.size > this.maxEntries) entries.delete(entries.keys().next().value!)
    return page
  }
}

/** The exact active cell (spec §11.1): its own record regardless of overview mode,
 * merge anchors spanning their range. Undefined when blank or collapsed. */
function activePlacement(geometry: SheetGeometry, at: SourceCell | undefined, work: ViewportWork): CellPlacement | undefined {
  if (!at || !isValidCoordinate(at.row, at.col)) return undefined
  const { metrics } = geometry
  const range = geometry.merges.firstWhere({ r0: at.row, r1: at.row, c0: at.col, c1: at.col }, () => true)
  const anchor = range ? { row: range.minRow, col: range.minCol } : at
  const slot = metrics.order.sourceToDisplay(anchor.row)
  if (!(metrics.rows.size(slot) > 0) || !(metrics.cols.size(anchor.col) > 0)) return undefined
  const cells = geometry.cellsOf(anchor.row)
  const index = cells ? firstCellAtOrAfter(cells, anchor.col) : -1
  if (!cells || index >= cells.length || cells[index].col !== anchor.col) return undefined
  work.semanticNodes = 1
  const x0 = metrics.cols.start(anchor.col), y0 = metrics.rows.start(slot)
  const width = range ? metrics.cols.start(range.maxCol + 1) - x0 : metrics.cols.size(anchor.col)
  const height = range ? metrics.rows.start(range.maxRow + 1) - y0 : metrics.rows.size(slot)
  return { cell: cells[index], source: anchor, x: x0, y: y0, width, height, ...(range ? { merge: range } : {}) }
}

function detailPlan(geometry: SheetGeometry, viewport: LogicalViewport, limits: DetailLimits, options: PlanOptions): ViewportPlan {
  const { metrics } = geometry
  const { x, width } = viewport.box
  const window = planWindow(metrics, viewport, limits)
  const work = window.work
  const active = activePlacement(geometry, options.activeCell, work)
  const coarse = (message: string, groups = window.kind === 'coarse' ? window.groups : coarseGroups(metrics, [{ rows: { start: Math.max(0, viewport.box.y), end: Math.min(window.extent.height, viewport.box.y + viewport.box.height) }, cols: { start: Math.max(0, x), end: Math.min(window.extent.width, x + width) } }], window.scale, limits.overviewGroups)): ViewportPlan => {
    work.overviewGroups = groups.length
    return { kind: 'coarse', diagnostics: [...geometry.diagnostics, { code: 'detail-limit', message }], groups, extent: window.extent, work, ...(active ? { active } : {}) }
  }
  if (window.kind === 'coarse') return coarse(window.message)
  const { scale, extent } = window
  const rowLines = window.rows.lines, colLines = window.cols.lines
  const diagnostics: WorkerDiagnostic[] = [...geometry.diagnostics]
  if (window.coarse.length) diagnostics.push({ code: 'detail-coarse-region', message: `${window.coarse.length} sub-detail pixel group(s) summarized at zoom ${viewport.zoom} x DPR ${viewport.deviceScale}` })
  work.overviewGroups = window.coarse.length
  const slots = window.rows.detail
  const detailCols = window.cols.detail
  const cells: Array<CellPlacement & { slot: number }> = []
  const merges: MergePlacement[] = []
  const done = (drawings: readonly XlsxDrawing[]): ViewportDetail => {
    cells.sort((a, b) => a.slot - b.slot || a.source.col - b.source.col)
    work.records = cells.length + drawings.length
    return { kind: 'detail', cells, merges, drawings, rowLines, colLines, extent, diagnostics, coarse: window.coarse, work, ...(active ? { active } : {}) }
  }
  if (!slots.length || !detailCols.length) return done([])
  const visibleC0 = detailCols[0], visibleC1 = detailCols[detailCols.length - 1]
  // Overflow dependencies: cells whose unclipped text may reach into the pane.
  // They obey the same detail admission: walking outward from the pane, columns
  // join only while every run projects to four device pixels, and a side whose
  // window alone exceeds the boundary budget contributes no dependencies.
  const dependencies = (from: number, to: number): AxisRun[] => {
    if (!(to > from)) return []
    const window = metrics.cols.visibleBetween(from, to)
    if (window.count > limits.gridlines) return []
    return [...window.runs]
  }
  let c0 = visibleC0, c1 = visibleC1
  const leftRuns = dependencies(Math.max(0, x - TEXT_OVERFLOW_REACH_CSS), Math.max(0, x))
  for (let index = leftRuns.length - 1; index >= 0 && !belowDetail(leftRuns[index].sizeCss, scale); index--) c0 = Math.min(c0, leftRuns[index].start)
  for (const run of dependencies(Math.min(extent.width, x + width), Math.min(extent.width, x + width + TEXT_OVERFLOW_REACH_CSS))) {
    if (belowDetail(run.sizeCss, scale)) break
    c1 = Math.max(c1, run.end)
  }
  // Sub-detail columns inside the envelope are summarized, never enumerated.
  const colIsDetail = (col: number) => col < visibleC0 || col > visibleC1 || !belowDetail(metrics.cols.size(col), scale)

  // Merges never span a permutation band (the S3/S7 sort guard refuses that), so
  // outside bands display slots equal source rows and the slot envelope is the
  // source query. A merge that does touch a band is painted cell by cell.
  const rect = { r0: slots[0], r1: slots[slots.length - 1], c0, c1 }
  const page = options.merges ? options.merges.query(geometry, rect, limits.records, work) : (work.mergeQueries++, geometry.merges.query(rect, { limit: limits.records }))
  if (page.nextCursor !== undefined) return coarse(`more than ${limits.records} merged ranges intersect the frame`)
  // Admission from indexes only: per-row column counts by binary search, no cell content.
  let records = page.items.length
  work.visitedRecords += page.items.length
  for (const slot of slots) {
    const rowCells = geometry.cellsOf(metrics.order.displayToSource(slot))
    if (rowCells) records += firstCellAtOrAfter(rowCells, c1 + 1) - firstCellAtOrAfter(rowCells, c0)
    if (records > limits.records) return coarse(`more than ${limits.records} cell records in the frame`)
  }
  const drawings = geometry.drawingsIn(x, viewport.box.y, x + width, viewport.box.y + viewport.box.height, limits.records - records)
  if (!drawings) return coarse(`more than ${limits.records} cell and drawing records in the frame`)

  const covering: XlsxMergeRange[] = []
  const bands = metrics.order.bands
  for (const range of page.items) {
    if (bands.some(band => band.r0 <= range.maxRow && band.r1 >= range.minRow)) {
      diagnostics.push({ code: 'merge-in-reordered-band', message: 'merged range inside a reordered band is painted per cell' })
      continue
    }
    const left = metrics.cols.start(range.minCol), top = metrics.rows.start(range.minRow)
    const placement = { range, x: left, y: top, width: metrics.cols.start(range.maxCol + 1) - left, height: metrics.rows.start(range.maxRow + 1) - top }
    covering.push(range)
    merges.push(placement)
    if (placement.width <= 0 || placement.height <= 0) continue
    const anchor = geometry.cellsOf(range.minRow)
    const at = anchor ? firstCellAtOrAfter(anchor, range.minCol) : -1
    // An anchor whose own row/column lies outside the window still owns the merged paint.
    if (anchor && at < anchor.length && anchor[at].col === range.minCol) {
      cells.push({ cell: anchor[at], source: { row: range.minRow, col: range.minCol }, slot: range.minRow, ...placement, merge: range })
    }
  }
  const locator = new MergeLocator(covering, c0, c1)
  for (const slot of slots) {
    const source = metrics.order.displayToSource(slot)
    const rowCells = geometry.cellsOf(source)
    if (!rowCells) continue
    const top = metrics.rows.start(slot), rowHeight = metrics.rows.size(slot)
    const end = firstCellAtOrAfter(rowCells, c1 + 1)
    for (let index = firstCellAtOrAfter(rowCells, c0); index < end; index++) {
      const item = rowCells[index]
      const col = item.col
      work.visitedRecords++
      const colWidth = metrics.cols.size(col)
      // Collapsed rows/columns produce no fill, text or hit target (spec §5.3).
      if (colWidth <= 0 || !colIsDetail(col) || locator.at(source, col)) continue
      cells.push({ cell: item, source: { row: source, col }, slot, x: metrics.cols.start(col), y: top, width: colWidth, height: rowHeight })
    }
  }
  if (tileOverflow(cells, viewport, scale, limits.tile.records)) {
    cells.length = 0
    merges.length = 0
    return coarse(`more than ${limits.tile.records} records in one ${TILE_DEVICE_PX}-device-pixel tile`)
  }
  return done(drawings)
}

// ---------------------------------------------------------------------------
// Exact source membership of display groups and the store-backed viewport
// ---------------------------------------------------------------------------

/** Positive-size runs visited per membership request before an explicit
 * resource outcome (never an approximated bounding box). */
const MAX_SET_RUNS = 65_536
const MAX_SET_INTERVALS = 4096

/** Exact source membership of the visible display slots [start,end] (spec §4.1):
 * identity stretches as source intervals, permuted bands as a sorted source list.
 * Hidden and zero-size slots are excluded. Returns null when the slots fragment
 * into more runs than the bounded budget (explicit resource outcome). */
export function sourceSetsForSlots(metrics: ProjectedSheetMetrics, start: number, end: number): CompactSourceRowSet[] | null {
  if (!(end >= start)) return []
  const bands = [...metrics.order.bands].sort((a, b) => a.r0 - b.r0)
  const intervals: Array<{ r0: number; r1: number }> = []
  const listed: number[] = []
  const pushInterval = (r0: number, r1: number) => {
    const last = intervals[intervals.length - 1]
    if (last && last.r1 + 1 === r0) last.r1 = r1
    else intervals.push({ r0, r1 })
  }
  const pushSlots = (s0: number, s1: number) => {
    let at = s0
    for (const band of bands) {
      if (band.r1 < at || band.r0 > s1) continue
      if (band.r0 > at) pushInterval(at, band.r0 - 1)
      const b0 = Math.max(at, band.r0), b1 = Math.min(s1, band.r1)
      for (let slot = b0; slot <= b1; slot++) listed.push(band.sourceAtOffset[slot - band.r0])
      at = b1 + 1
      if (at > s1) return
    }
    if (at <= s1) pushInterval(at, s1)
  }
  const from = metrics.rows.start(start)
  const to = metrics.rows.start(end) + metrics.rows.size(end)
  let runs = 0
  for (const run of metrics.rows.visibleBetween(from, Math.max(to, from)).runs) {
    if (++runs > MAX_SET_RUNS) return null
    const s0 = Math.max(start, run.start), s1 = Math.min(end, run.end)
    if (s1 >= s0) pushSlots(s0, s1)
  }
  const sets: CompactSourceRowSet[] = []
  if (intervals.length > MAX_SET_INTERVALS) sets.push(bitsetOf(intervals))
  else if (intervals.length) sets.push({ kind: 'intervals', intervals })
  if (listed.length) {
    const rows = Uint32Array.from(listed).sort()
    sets.push({ kind: 'list', r0: rows[0], r1: rows[rows.length - 1], rows })
  }
  return sets
}
function bitsetOf(intervals: ReadonlyArray<{ r0: number; r1: number }>): CompactSourceRowSet {
  const r0 = intervals[0].r0, r1 = intervals[intervals.length - 1].r1
  const words = new Uint32Array(Math.ceil((r1 - r0 + 1) / 32))
  let count = 0
  for (const interval of intervals) {
    for (let row = interval.r0; row <= interval.r1; row++) {
      const offset = row - r0
      words[Math.floor(offset / 32)] = (words[Math.floor(offset / 32)] | (1 << (offset % 32))) >>> 0
      count++
    }
  }
  return { kind: 'bitset', space: 'source', r0, r1, words, count }
}

export interface StoreViewportOptions {
  profile?: ViewportProfile
  signal: AbortSignal
  /** Publication guard (spec §10.1): a late result whose owner is no longer
   * current is reported `stale` and never published. */
  isCurrent?: (owner: CacheIdentity) => boolean
  /** `tile` applies the per-tile record/page caps instead of the per-frame ones. */
  scope?: 'frame' | 'tile'
  activeCell?: SourceCell
}
export interface StoreViewport {
  kind: 'viewport'
  owner: CacheIdentity
  completeness: Completeness
  error?: SourceFailure
  /** Detail cells (and explicit unknown rows) at their logical boxes; merge
   * anchors carry their merged box, covered cells are omitted. */
  cells: readonly ViewportCell[]
  merges: readonly StoreMerge[]
  /** Coarse summaries assigned to their display pixel boxes. */
  groups: readonly ViewportOverviewGroup[]
  coarse: readonly CoarseGroup[]
  rowLines: readonly number[]
  colLines: readonly number[]
  active?: ViewportCell
  work: ViewportWork
  diagnostics: readonly WorkerDiagnostic[]
}
export type StoreViewportOutcome =
  | StoreViewport
  | { kind: 'stale' | 'cancelled'; owner: CacheIdentity; work: ViewportWork; diagnostics: readonly WorkerDiagnostic[] }
  | { kind: 'diagnostic'; diagnostics: readonly WorkerDiagnostic[] }

class Cancelled extends Error {}

/** Store-backed frame (spec §§4.1, 11.1): admission from geometry and structural
 * metadata, then
 * bounded `readRows` pages for detail and analytic `queryOverview` for coarse
 * groups. A frame that needs more pages than its budget declares the rest
 * loading; unknown summaries are loading, never blank; failures stay explicit.
 * TODO(S5.2): the object path's text-overflow dependency columns and drawing
 * admission are not yet mirrored here; this path is not wired to the worker yet,
 * and must gain both before the worker `viewport` command publishes it. */
export async function queryViewport(store: SourceStore, metrics: ProjectedSheetMetrics, viewport: LogicalViewport, owner: CacheIdentity, options: StoreViewportOptions): Promise<StoreViewportOutcome> {
  const profile = VIEWPORT_DETAIL_LIMITS[options.profile ?? 'D']
  const limits = options.scope === 'tile' ? { ...profile, records: profile.tile.records, pages: profile.tile.pages } : profile
  const diagnostics: WorkerDiagnostic[] = []
  let window: WindowPlan
  try {
    window = planWindow(metrics, viewport, limits)
  } catch (error) {
    if (!(error instanceof AxisMetricsError)) throw error
    return { kind: 'diagnostic', diagnostics: [axisDiagnostic(error, 'viewport')] }
  }
  const work = window.work
  if (options.signal.aborted) return { kind: 'cancelled', owner, work, diagnostics }
  const readOptions = { signal: options.signal, stamp: store.stamp }
  try {
    let coarse = window.kind === 'coarse' ? window.groups : window.coarse
    let detail = window.kind === 'detail' ? window : null
    if (window.kind === 'coarse') diagnostics.push({ code: 'detail-limit', message: window.message })
    const state = { completeness: 'complete' as Completeness, error: undefined as SourceFailure | undefined }
    const merge = (completeness: Completeness, error?: SourceFailure) => {
      if (completeness === 'incomplete') { state.completeness = 'incomplete'; state.error ??= error }
      else if (completeness === 'loading' && state.completeness === 'complete') state.completeness = 'loading'
    }
    let cells: ViewportCell[] = []
    const mergePlacements: StoreMerge[] = []
    // An empty axis (pane past the last row/column) has nothing to read.
    if (detail && detail.rows.detail.length && detail.cols.detail.length) {
      // Admission before any cell read, from geometry plus structural metadata:
      // a source may populate every cell, but a merged range is one record (its
      // anchor across the merged box), exactly as the object path counts records.
      const structural = await structuralMerges(store, metrics, detail.rows.detail, detail.cols.detail, readOptions, merge, diagnostics)
      if (options.signal.aborted) throw new Cancelled()
      const boxes = candidateBoxes(metrics, detail.rows.detail, detail.cols.detail, structural, limits.records)
      const tiles = boxes !== null && tileOverflow(boxes, viewport, window.scale, profile.tile.records)
      if (boxes === null || tiles) {
        diagnostics.push({ code: 'detail-limit', message: tiles ? `more than ${profile.tile.records} records in one ${TILE_DEVICE_PX}-device-pixel tile` : `more than ${limits.records} candidate records in the frame` })
        coarse = coarseGroups(metrics, [{ rows: { start: Math.max(0, viewport.box.y), end: Math.min(window.extent.height, viewport.box.y + viewport.box.height) }, cols: { start: Math.max(0, viewport.box.x), end: Math.min(window.extent.width, viewport.box.x + viewport.box.width) } }], window.scale, limits.overviewGroups)
        detail = null
      } else {
        if (coarse.length) diagnostics.push({ code: 'detail-coarse-region', message: `${coarse.length} sub-detail pixel group(s) summarized` })
        cells = await readDetail(store, metrics, detail.rows.detail, detail.cols.detail, limits.pages, readOptions, work, merge)
        cells = await applyMerges(store, cells, structural, viewport, limits.pages, readOptions, work, merge, mergePlacements)
      }
    }
    const groups = coarse.length ? await overview(store, metrics, coarse, limits.overviewGroups, limits.pages, readOptions, work, merge) : []
    let active: ViewportCell | undefined
    if (options.activeCell && isValidCoordinate(options.activeCell.row, options.activeCell.col)) {
      const at = options.activeCell
      const read = await store.readCell(at, readOptions)
      const box = metrics.sourceCellToLogical(at)
      // Blank and collapsed cells have no exact record; loading/error reads stay
      // explicit placeholder records so the inspector can say so.
      if (box && read.kind !== 'blank') { active = { ...read, source: at, bounds: box } as ViewportCell; work.semanticNodes = 1 }
    }
    if (options.signal.aborted) return { kind: 'cancelled', owner, work, diagnostics }
    if (options.isCurrent && !options.isCurrent(owner)) return { kind: 'stale', owner, work, diagnostics }
    work.records = cells.filter(item => item.kind === 'cell' || item.kind === 'unavailable').length
    work.overviewGroups = groups.length
    return {
      kind: 'viewport', owner, completeness: state.completeness, ...(state.error ? { error: state.error } : {}),
      cells, merges: mergePlacements, groups, coarse, rowLines: detail ? detail.rows.lines : [], colLines: detail ? detail.cols.lines : [],
      ...(active ? { active } : {}), work, diagnostics,
    }
  } catch (error) {
    if (error instanceof Cancelled || options.signal.aborted) return { kind: 'cancelled', owner, work, diagnostics }
    if (error instanceof AxisMetricsError) return { kind: 'diagnostic', diagnostics: [...diagnostics, axisDiagnostic(error, 'viewport')] }
    throw error
  }
}

/** Candidate record boxes of a store detail grid: every cell outside a merge
 * (a source may populate any of them) plus one merged box per visible merge.
 * Null once the count exceeds `cap`, before enumerating further. */
function candidateBoxes(metrics: ProjectedSheetMetrics, slots: readonly number[], cols: readonly number[], merges: readonly StoreMerge[], cap: number): LogicalCellBox[] | null {
  const boxes: LogicalCellBox[] = merges.map(item => item.bounds)
  if (boxes.length > cap) return null
  const locator = merges.length ? new MergeLocator(merges.map(item => item.range), 0, MAX_COLS - 1) : null
  for (const slot of slots) {
    const row = metrics.order.displayToSource(slot)
    const y = metrics.rows.start(slot), height = metrics.rows.size(slot)
    for (const col of cols) {
      if (locator?.at(row, col)) continue
      boxes.push({ x: metrics.cols.start(col), y, width: metrics.cols.size(col), height })
      if (boxes.length > cap) return null
    }
  }
  return boxes
}

type Merge = (completeness: Completeness, error?: SourceFailure) => void
export interface StoreMerge { range: XlsxMergeRange; bounds: LogicalCellBox }

/** Structural intersection (spec §4.5) over the detail source envelope: one
 * bounded index query; ranges stay rectangles, never per-covered-cell entries. */
async function structuralMerges(store: SourceStore, metrics: ProjectedSheetMetrics, slots: readonly number[], cols: readonly number[], options: ReadOptionsLike, merge: Merge, diagnostics: WorkerDiagnostic[]): Promise<StoreMerge[]> {
  let r0 = MAX_ROWS, r1 = -1
  for (const slot of slots) { const row = metrics.order.displayToSource(slot); r0 = Math.min(r0, row); r1 = Math.max(r1, row) }
  const result = await store.queryStructural({ r0, r1, c0: cols[0], c1: cols[cols.length - 1] }, options)
  merge(result.completeness, result.error)
  const out: StoreMerge[] = []
  for (const extent of result.merges) {
    const range = { minRow: extent.range.r0, maxRow: extent.range.r1, minCol: extent.range.c0, maxCol: extent.range.c1 }
    if (metrics.order.bands.some(band => band.r0 <= range.maxRow && band.r1 >= range.minRow)) {
      diagnostics.push({ code: 'merge-in-reordered-band', message: 'merged range inside a reordered band is painted per cell' })
      continue
    }
    const x = metrics.cols.start(range.minCol), y = metrics.rows.start(range.minRow)
    out.push({ range, bounds: { x, y, width: metrics.cols.start(range.maxCol + 1) - x, height: metrics.rows.start(range.maxRow + 1) - y } })
  }
  return out
}

/** Covered cells disappear; each visible merge's anchor paints across the merged
 * box, read individually (within the page budget) when its own row is offscreen. */
async function applyMerges(store: SourceStore, cells: ViewportCell[], merges: readonly StoreMerge[], viewport: LogicalViewport, pageBudget: number, options: ReadOptionsLike, work: ViewportWork, merge: Merge, placements: StoreMerge[]): Promise<ViewportCell[]> {
  if (!merges.length) return cells
  const locator = new MergeLocator(merges.map(item => item.range), 0, MAX_COLS - 1)
  const byRange = new Map(merges.map(item => [item.range, item]))
  const anchors = new Set<XlsxMergeRange>()
  const out: ViewportCell[] = []
  for (const item of cells) {
    const range = item.kind === 'unknown' ? undefined : locator.at(item.source.row, item.source.col)
    if (!range) { out.push(item); continue }
    if (range.minRow !== item.source.row || range.minCol !== item.source.col) continue
    anchors.add(range)
    out.push({ ...item, bounds: byRange.get(range)!.bounds })
  }
  const { x, y, width, height } = viewport.box
  for (const item of merges) {
    const { bounds, range } = item
    if (!(bounds.width > 0 && bounds.height > 0) || bounds.x >= x + width || bounds.x + bounds.width <= x || bounds.y >= y + height || bounds.y + bounds.height <= y) continue
    placements.push(item)
    if (anchors.has(range)) continue
    if (work.pages >= pageBudget) { merge('loading'); out.push({ kind: 'unknown', pageId: 'page-budget', source: { row: range.minRow, col: range.minCol }, bounds }); continue }
    work.pages++
    const read = await store.readCell({ row: range.minRow, col: range.minCol }, options)
    if (options.signal.aborted) throw new Cancelled()
    if (read.kind === 'blank') continue
    if (read.kind === 'unknown') merge('loading')
    if (read.kind === 'error') merge('incomplete', { code: read.code, retryable: read.retryable, message: 'merge anchor read failed' })
    out.push({ ...read, source: { row: range.minRow, col: range.minCol }, bounds } as ViewportCell)
  }
  return out
}
interface ReadOptionsLike { signal: AbortSignal; stamp: SourceStore['stamp'] }

/** Detail rows as ascending exact source runs (sorted bands map consecutive
 * slots to disjoint sources); one bounded `readRows` page per request, at most
 * `pageBudget` per frame. Rows past the budget are explicit unknown records. */
async function readDetail(store: SourceStore, metrics: ProjectedSheetMetrics, slots: readonly number[], cols: readonly number[], pageBudget: number, options: ReadOptionsLike, work: ViewportWork, merge: Merge): Promise<ViewportCell[]> {
  const c0 = cols[0], c1 = cols[cols.length - 1]
  const detailCol = new Set(cols)
  const sources = slots.map(slot => metrics.order.displayToSource(slot)).sort((a, b) => a - b)
  const runs: Array<{ r0: number; r1: number }> = []
  for (const row of sources) {
    const last = runs[runs.length - 1]
    if (last && last.r1 + 1 === row) last.r1 = row
    else runs.push({ r0: row, r1: row })
  }
  const cells: ViewportCell[] = []
  const seen = new Set<number>()
  const unknownRow = (row: number, pageId: string) => {
    const slot = metrics.order.sourceToDisplay(row)
    const x = metrics.cols.start(c0)
    cells.push({ kind: 'unknown', pageId, source: { row, col: c0 }, bounds: { x, y: metrics.rows.start(slot), width: metrics.cols.start(c1) + metrics.cols.size(c1) - x, height: metrics.rows.size(slot) } })
  }
  // Each cell carries its own source row: a page item is not assumed to hold a
  // single row (the SourceStore contract does not pin that), so rows are marked
  // seen and placed per cell.
  const place = (page: SparseRowPage, range: { r0: number; r1: number }) => {
    seen.add(page.row)
    for (const cell of page.cells) {
      work.visitedRecords++
      const row = cell.row
      if (!(row >= range.r0 && row <= range.r1) || !detailCol.has(cell.col)) continue
      seen.add(row)
      const slot = metrics.order.sourceToDisplay(row)
      const bounds = { x: metrics.cols.start(cell.col), y: metrics.rows.start(slot), width: metrics.cols.size(cell.col), height: metrics.rows.size(slot) }
      const { unavailable, ...payload } = cell
      cells.push(unavailable
        ? { kind: 'unavailable', cell: payload, reason: 'formula-value-unavailable', source: { row, col: cell.col }, bounds }
        : { kind: 'cell', cell: payload, source: { row, col: cell.col }, bounds })
    }
  }
  let budgetHit = false
  for (const run of runs) {
    // A complete run's omitted rows are blank (sparse pages); a run with a
    // loading/failed page, or one cut by the page budget, reports its unseen
    // rows as explicit unknown records instead.
    let complete = !budgetHit
    let cursor: string | undefined
    while (!budgetHit) {
      if (work.pages >= pageBudget) { budgetHit = true; complete = false; break }
      work.pages++
      const page = await store.readRows({ r0: run.r0, r1: run.r1, c0, c1 }, cursor, options)
      if (options.signal.aborted) throw new Cancelled()
      merge(page.completeness, page.error)
      if (page.completeness !== 'complete') complete = false
      for (const item of page.items) place(item, run)
      cursor = page.nextCursor
      if (cursor === undefined) break
    }
    if (!complete) for (let row = run.r0; row <= run.r1; row++) if (!seen.has(row)) unknownRow(row, budgetHit ? 'page-budget' : `rows:${run.r0}-${run.r1}`)
  }
  if (budgetHit) merge('loading')
  cells.sort((a, b) => a.bounds.y - b.bounds.y || a.bounds.x - b.bounds.x)
  return cells
}

/** Coarse groups → exact source domains → bounded `queryOverview` pages →
 * aggregates assigned to display boxes (spec §4.1). A group splits into at most
 * two domains (identity intervals and a permuted-band list); its aggregate is the
 * exact combination, unknown if any part is unknown, never a sampled value. */
async function overview(store: SourceStore, metrics: ProjectedSheetMetrics, groups: readonly CoarseGroup[], maxGroups: number, pageBudget: number, options: ReadOptionsLike, work: ViewportWork, merge: Merge): Promise<ViewportOverviewGroup[]> {
  const domains: OverviewDomain[] = []
  const parts = new Map<string, string[]>()
  for (const group of groups) {
    const sets = sourceSetsForSlots(metrics, group.slots.start, group.slots.end)
    if (!sets) {
      merge('incomplete', { code: 'resource-limit', retryable: false, message: `display group ${group.groupId} fragments beyond ${MAX_SET_RUNS} runs` })
      continue
    }
    const ids: string[] = []
    sets.forEach((rows, index) => {
      const groupId = `${group.groupId}.${index}`
      ids.push(groupId)
      domains.push({ groupId, rows, cols: [group.cols] })
    })
    parts.set(group.groupId, ids)
  }
  const results = new Map<string, OverviewGroup>()
  let cursor: string | undefined
  do {
    // Output is bounded per page; a store that keeps paging past the frame's
    // page budget ends in an explicit resource outcome, never a loop.
    if (work.overviewPages >= pageBudget) {
      merge('incomplete', { code: 'resource-limit', retryable: true, message: `overview needs more than ${pageBudget} pages` })
      break
    }
    work.overviewPages++
    const page = await store.queryOverview({ groups: domains, maxGroups, cursor }, options)
    if (options.signal.aborted) throw new Cancelled()
    merge(page.completeness, page.error)
    for (const item of page.items) results.set(item.groupId, item)
    cursor = page.nextCursor
    if (page.error && !page.items.length && cursor === undefined) break
  } while (cursor !== undefined)
  const out: ViewportOverviewGroup[] = []
  for (const group of groups) {
    const ids = parts.get(group.groupId)
    if (!ids) continue
    const items = ids.map(id => results.get(id))
    if (items.some(item => item === undefined)) continue // failed domain: reported via page.error, never shown as blank
    const known = items as OverviewGroup[]
    const unknown = known.find((item): item is Extract<OverviewGroup, { content: 'unknown' }> => item.content === 'unknown')
    if (unknown) {
      // Unknown coverage is loading even when the page itself reports complete.
      merge('loading')
      out.push({ groupId: group.groupId, content: 'unknown', pageId: unknown.pageId, bounds: group.bounds })
      continue
    }
    const aggregates = known as Array<Exclude<OverviewGroup, { content: 'unknown' }>>
    if (!aggregates.length) { out.push({ groupId: group.groupId, populatedCount: 0, countKind: 'exact', style: { kind: 'uniform', styleIndex: 0 }, content: 'blank', bounds: group.bounds }); continue }
    const first = aggregates[0].style
    const uniform = first.kind === 'uniform' && aggregates.every(item => item.style.kind === 'uniform' && item.style.styleIndex === first.styleIndex)
    const contents = new Set(aggregates.map(item => item.content))
    out.push({
      groupId: group.groupId,
      populatedCount: aggregates.reduce((sum, item) => sum + item.populatedCount, 0),
      countKind: aggregates.every(item => item.countKind === 'exact') ? 'exact' : 'estimated',
      style: uniform ? first : { kind: 'mixed' },
      content: contents.size === 1 ? aggregates[0].content : 'mixed',
      bounds: group.bounds,
    })
  }
  return out
}

/** Re-exported for typed consumers of permuted bands. */
export type { RowPermutationBand }

/** Worker reply adapter (spec §§4.1, 10.2): one `ViewportResult` per content
 * kind present (a mixed pane yields detail and overview). A content is complete
 * only when it holds no unknown/error record; otherwise it carries the frame's
 * non-complete state and error, so no loading record is published as complete.
 * `ViewportResult` has no slot for the exact active-cell record (spec §11.1), so
 * it is carried beside the results; the worker `viewport` reply must publish it
 * (inspector/status/accessible name) under the same owner. */
export function toViewportResults(outcome: StoreViewport): { results: ViewportResult[]; active?: ViewportCell } {
  return { results: viewportResults(outcome), ...(outcome.active ? { active: outcome.active } : {}) }
}
function viewportResults(outcome: StoreViewport): ViewportResult[] {
  const results: ViewportResult[] = []
  const pendingCells = outcome.cells.some(item => item.kind === 'unknown' || item.kind === 'error')
  const pendingGroups = outcome.groups.some(group => group.content === 'unknown')
  const state = (pending: boolean): 'complete' | 'loading' | 'incomplete' =>
    outcome.completeness === 'complete' && !pending ? 'complete' : outcome.completeness === 'complete' ? 'loading' : outcome.completeness
  const error = outcome.error ? { error: outcome.error } : {}
  if (outcome.cells.length || !outcome.groups.length) {
    const completeness = state(pendingCells)
    results.push(completeness === 'complete'
      ? { kind: 'viewport', completeness, content: { kind: 'detail', cells: outcome.cells as Extract<CompleteViewportContent, { kind: 'detail' }>['cells'] } }
      : { kind: 'viewport', completeness, content: { kind: 'detail', cells: outcome.cells }, ...error })
  }
  if (outcome.groups.length) {
    const completeness = state(pendingGroups)
    results.push(completeness === 'complete'
      ? { kind: 'viewport', completeness, content: { kind: 'overview', groups: outcome.groups as Extract<CompleteViewportContent, { kind: 'overview' }>['groups'] } }
      : { kind: 'viewport', completeness, content: { kind: 'overview', groups: outcome.groups }, ...error })
  }
  return results
}
