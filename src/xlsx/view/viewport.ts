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
import type { XlsxCell, XlsxDrawing, XlsxMergeRange, XlsxRow, XlsxSheet } from '../types'
import type { DrawingContentShape } from '../../drawing/content'
import { resolveGeometry } from '../../drawing/geometry'
import type { AxisRun, AxisSizeInterval, IntervalSet, LogicalViewport, RowOrder, RowVisibility, WorkerDiagnostic } from './contracts'
import { MAX_COLS, MAX_ROWS, isValidCoordinate, type SourceCell } from './coordinates'
import { AxisMetricsError, createAxisMetrics, type AxisMetricsErrorCode, type SparseAxisMetrics } from './axes'
import { createSheetMetrics, type ProjectedSheetMetrics } from './projection'
import { RectangleIndex } from './structure'
import { createIntervalSet } from './visibility'

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

/** Proposed §11.1 per-frame limits: individual records (cells plus drawings) and
 * gridline instances on desktop (D) and mobile (M). Beyond them a frame is coarse
 * before any cell is read; the coarse layer itself belongs to S5.1. */
export const VIEWPORT_DETAIL_LIMITS = Object.freeze({
  D: Object.freeze({ records: 20_000, gridlines: 4096 }),
  M: Object.freeze({ records: 10_000, gridlines: 2048 }),
})
export type ViewportProfile = keyof typeof VIEWPORT_DETAIL_LIMITS

const EXPLICIT_REF = /^[A-Z]+\d+$/
/** Source row of a cell. An explicit cell reference wins over its row record:
 * `<row>` without `r` parses at index 0, but `<c r="A16386">` still names row
 * 16,385. Cells without a reference inherit the row record's index. */
export function cellSourceRow(row: XlsxRow, cell: XlsxCell): number {
  return EXPLICIT_REF.test(cell.ref) ? cell.row : row.index
}

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

type Matrix = [number, number, number, number, number, number]
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]
function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]]
}
/** Painted extent of drawing nodes in worksheet CSS px, including rotation, flips,
 * group transforms, custom geometry and diagram children. Empty input yields
 * infinite bounds (left/top +Infinity, right/bottom -Infinity). */
export function drawingExtent(nodes: readonly XlsxDrawing[]): { left: number; top: number; right: number; bottom: number } {
  const bounds = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity }
  type BoundsNode = XlsxDrawing | DrawingContentShape
  const walk = (items: readonly BoundsNode[], parent: Matrix): void => {
    for (const node of items) {
      const x = node.xEmu / 9525, y = node.yEmu / 9525, w = node.widthEmu / 9525, h = node.heightEmu / 9525
      const deg = node.rotationDeg ?? 0
      if ('transformValid' in node && node.transformValid === false || ![x, y, w, h, deg].every(Number.isFinite) || w < 0 || h < 0) continue
      const rad = deg * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad)
      const fx = node.flipH ? -1 : 1, fy = node.flipV ? -1 : 1
      const cx = x + w / 2, cy = y + h / 2
      const local: Matrix = [cos * fx, sin * fx, -sin * fy, cos * fy, cx - cos * fx * w / 2 + sin * fy * h / 2, cy - sin * fx * w / 2 - cos * fy * h / 2]
      const matrix = multiply(parent, local)
      const include = (px: number, py: number) => {
        const tx = matrix[0] * px + matrix[2] * py + matrix[4]
        const ty = matrix[1] * px + matrix[3] * py + matrix[5]
        if (Number.isFinite(tx) && Number.isFinite(ty)) {
          bounds.left = Math.min(bounds.left, tx); bounds.top = Math.min(bounds.top, ty)
          bounds.right = Math.max(bounds.right, tx); bounds.bottom = Math.max(bounds.bottom, ty)
        }
      }
      if (node.group) {
        const g = node.group, sx = g.ext.width / g.chExt.width, sy = g.ext.height / g.chExt.height
        if (![sx, sy, g.chOff.x, g.chOff.y].every(Number.isFinite) || sx < 0 || sy < 0) continue
        const childMatrix = multiply(matrix, [sx, 0, 0, sy, -sx * g.chOff.x / 9525, -sy * g.chOff.y / 9525])
        walk((node.children ?? []) as BoundsNode[], childMatrix)
      } else {
        include(0, 0); include(w, 0); include(w, h); include(0, h)
        if (node.drawingGeometry && !node.drawingGeometry.preset) {
          for (const path of resolveGeometry(node.drawingGeometry, w, h).paths) {
            for (const command of path.commands) {
              if (command[0] === 'moveTo' || command[0] === 'lnTo') include(command[1], command[2])
              else if (command[0] === 'quadBezTo') { include(command[1], command[2]); include(command[3], command[4]) }
              else if (command[0] === 'cubicBezTo') { include(command[1], command[2]); include(command[3], command[4]); include(command[5], command[6]) }
              else if (command[0] === 'arcTo') {
                const [, cx, cy, rx, ry] = command
                include(cx - rx, cy - ry); include(cx + rx, cy - ry)
                include(cx - rx, cy + ry); include(cx + rx, cy + ry)
              }
            }
          }
        }
        if ('content' in node && node.content?.kind === 'diagram') walk(node.content.shapes as BoundsNode[], matrix)
      }
    }
  }
  walk(nodes, IDENTITY)
  return bounds
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
}
export type ViewportPlan =
  | ViewportDetail
  | { kind: 'pending' | 'diagnostic' | 'coarse'; diagnostics: readonly WorkerDiagnostic[] }

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

/** Plan one frame. Admission happens before any cell is read: window counts
 * (O(log n)), projected device size per run, distinct boundary instances,
 * per-row column-index counts, and merge/drawing index queries. Only an admitted
 * frame resolves its source cells; source rows outside the window are never scanned. */
export function planViewport(geometry: SheetGeometryResult | undefined, viewport: LogicalViewport, profile: ViewportProfile = 'D'): ViewportPlan {
  if (!geometry) return { kind: 'pending', diagnostics: [{ code: 'geometry-pending', message: 'sheet geometry has not been published' }] }
  if (geometry.kind !== 'ready') return { kind: 'diagnostic', diagnostics: geometry.diagnostics }
  try {
    return detailPlan(geometry, viewport, VIEWPORT_DETAIL_LIMITS[profile])
  } catch (error) {
    if (!(error instanceof AxisMetricsError)) throw error
    return { kind: 'diagnostic', diagnostics: [...geometry.diagnostics, axisDiagnostic(error, 'viewport')] }
  }
}

/** Boundary positions of the window's runs, coalescing coincident positions
 * (zero-extent rows/columns between runs share a boundary). */
function boundaries(runs: Iterable<AxisRun>, start: (index: number) => number, scale: number, limit: number): { lines: number[]; tooSmall: boolean } | null {
  const lines: number[] = []
  let tooSmall = false
  const push = (value: number) => { if (lines[lines.length - 1] !== value) lines.push(value) }
  for (const run of runs) {
    if (belowDetail(run.sizeCss, scale)) tooSmall = true
    for (let index = run.start; index <= run.end; index++) {
      push(start(index))
      if (lines.length > limit) return null
    }
    push(start(run.end + 1))
  }
  return { lines, tooSmall }
}

function detailPlan(geometry: SheetGeometry, viewport: LogicalViewport, limits: { records: number; gridlines: number }): ViewportPlan {
  const { metrics } = geometry
  const { x, y, width, height } = viewport.box
  if (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0) {
    throw new AxisMetricsError('INVALID_QUERY', 'Viewport size must be finite and nonnegative')
  }
  const scale = viewport.zoom * viewport.deviceScale
  if (!Number.isFinite(viewport.zoom) || viewport.zoom <= 0 || !Number.isFinite(viewport.deviceScale) || viewport.deviceScale <= 0 || !Number.isFinite(scale)) {
    throw new AxisMetricsError('INVALID_QUERY', 'Viewport zoom and device scale must be finite and positive')
  }
  const rowWindow = metrics.rows.visibleBetween(y, y + height)
  const colWindow = metrics.cols.visibleBetween(x, x + width)
  const coarse = (message: string): ViewportPlan => ({ kind: 'coarse', diagnostics: [...geometry.diagnostics, { code: 'detail-limit', message }] })
  // Necessary bound from O(log n) counts: each visible index owns one boundary, plus one closing each axis.
  if (rowWindow.count + colWindow.count + 2 > limits.gridlines) {
    return coarse(`${rowWindow.count} rows x ${colWindow.count} columns exceed ${limits.gridlines} gridline instances`)
  }
  const rowBounds = boundaries(rowWindow.runs, index => metrics.rows.start(index), scale, limits.gridlines)
  const colBounds = boundaries(colWindow.runs, index => metrics.cols.start(index), scale, limits.gridlines)
  if (!rowBounds || !colBounds || rowBounds.lines.length + colBounds.lines.length > limits.gridlines) {
    return coarse(`boundary instances exceed ${limits.gridlines}`)
  }
  // TODO(S5.1): coarse is per frame here: one sub-detail run makes the whole pane
  // coarse. The S5.1 coarse layer should coalesce only those runs and keep detail
  // for the rest; until it lands, a mixed pane shows the coarse notice.
  if (rowBounds.tooSmall || colBounds.tooSmall) {
    return coarse(`visible rows/columns project below ${MIN_DETAIL_DEVICE_PX} device pixels at zoom ${viewport.zoom} x DPR ${viewport.deviceScale}`)
  }
  const rowLines = rowBounds.lines, colLines = colBounds.lines
  const extent = { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
  const diagnostics = [...geometry.diagnostics]
  const slots: number[] = []
  for (const run of rowWindow.runs) for (let slot = run.start; slot <= run.end; slot++) slots.push(slot)
  const cells: Array<CellPlacement & { slot: number }> = []
  const merges: MergePlacement[] = []
  let visibleC0 = MAX_COLS, visibleC1 = -1
  for (const run of colWindow.runs) { visibleC0 = Math.min(visibleC0, run.start); visibleC1 = Math.max(visibleC1, run.end) }
  if (!slots.length || visibleC1 < 0) return { kind: 'detail', cells, merges, drawings: [], rowLines, colLines, extent, diagnostics }
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

  // Merges never span a permutation band (the S3/S7 sort guard refuses that), so
  // outside bands display slots equal source rows and the slot envelope is the
  // source query. A merge that does touch a band is painted cell by cell.
  const page = geometry.merges.query({ r0: slots[0], r1: slots[slots.length - 1], c0, c1 }, { limit: limits.records })
  if (page.nextCursor !== undefined) return coarse(`more than ${limits.records} merged ranges intersect the frame`)
  // Admission from indexes only: per-row column counts by binary search, no cell content.
  let records = page.items.length
  for (const slot of slots) {
    const rowCells = geometry.cellsOf(metrics.order.displayToSource(slot))
    if (rowCells) records += firstCellAtOrAfter(rowCells, c1 + 1) - firstCellAtOrAfter(rowCells, c0)
    if (records > limits.records) return coarse(`more than ${limits.records} cell records in the frame`)
  }
  const drawings = geometry.drawingsIn(x, y, x + width, y + height, limits.records - records)
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
      const colWidth = metrics.cols.size(col)
      // Collapsed rows/columns produce no fill, text or hit target (spec §5.3).
      if (colWidth <= 0 || locator.at(source, col)) continue
      cells.push({ cell: item, source: { row: source, col }, slot, x: metrics.cols.start(col), y: top, width: colWidth, height: rowHeight })
    }
  }
  cells.sort((a, b) => a.slot - b.slot || a.source.col - b.source.col)
  return { kind: 'detail', cells, merges, drawings, rowLines, colLines, extent, diagnostics }
}
