/**
 * View-only row/column resizing (spec §5.6 amendment). Session sizes are a
 * sparse overlay on the authored axis metadata: they live per sheet and per
 * viewer session, are never written to the file, and never expand the
 * 1,048,576 × 16,384 axes into per-index arrays.
 *
 * - Columns are keyed by source column (the view never permutes columns).
 * - Rows are keyed by SOURCE row: the projection looks each display slot's
 *   height up through the row order, so a resized row keeps its height when a
 *   sort moves it and loses it to the mask when a filter hides it.
 * - Authored-hidden rows/columns stay hidden: a session size never replaces a
 *   zero-size authored interval, and `collapsed` intervals win in the axis.
 *
 * Pure and lazy: only the interactive viewer chunk reaches this module.
 */
import type { AxisMetrics, AxisSizeInterval } from './contracts'
import type { XlsxDrawing } from '../types'
import { retainDrawingCopyOwners } from './drawingOwners'
import { cellFontString } from '../cellFont'

export type ResizeAxis = 'row' | 'col'
/** Whole-CSS-pixel bounds. The maxima follow Excel's own limits (255
 * characters ≈ 1790 px wide, 409 pt ≈ 545 px tall); the minima keep a
 * resized line above the §11.1 detail rule at 50% zoom. Clamping keeps
 * session values themselves well inside Float64 axis precision; beside extreme
 * authored sizes a session size can still be unrepresentable, and then the
 * builders keep the authored geometry with a typed diagnostic. */
export const RESIZE_LIMITS = Object.freeze({ col: Object.freeze({ min: 8, max: 1790 }), row: Object.freeze({ min: 8, max: 546 }) })
/** At most this many session sizes per axis per sheet (a deliberate gesture
 * each); a further new index is refused with a notice, never silently dropped. */
export const MAX_SESSION_SIZES = 4096
/** Retained bytes per session entry: the Map entry plus its published snapshot copy. */
export const SESSION_SIZE_ENTRY_BYTES = 128
/** One repositioned drawing/index (320 B) and at most three new ownership
 * links (64 B each): copied node, copied group and the result array's share. */
export const SIZED_DRAWING_BYTES = 512
/** Upper bound of the bytes building one sparse axis allocates for
 * `intervals` authored/hidden/session intervals (two segments each with typed
 * summaries and JS records, plus normalization copies; axes.ts). */
export const axisUpperBytes = (intervals: number): number => 272 * intervals + 400
/** Keyboard step in CSS px (Alt+Shift+Arrow). */
export const RESIZE_KEY_STEP = 8
/** Auto-fit reads at most this many source rows for a column: the rows visible
 * in the pane first, then rows the source already holds on the main thread
 * (the resident cell index or the worker row cache), never a full-sheet scan
 * and never a source read. Known rows are visited at most 4× this many times. */
export const AUTOFIT_MAX_ROWS = 2048
/** Auto-fit measures at most this many cells of a row: the visible columns,
 * then the row's populated cells. */
export const AUTOFIT_MAX_CELLS = 1024
// Duplicated from viewport.ts / render.ts (importing them would form a cycle or
// add exports to the eager entry): default geometry and horizontal cell padding.
const DEFAULT_ROW_CSS = 20
const DEFAULT_COL_CSS = 64
const CELL_PADDING_CSS = 6
// Calibrated to Excel's auto-fit: Calibri 11 pt → 20 px, 20 pt → 35 px.
const LINE_HEIGHT = 1.25
const ROW_PADDING_CSS = 2

export function clampSize(axis: ResizeAxis, sizeCss: number): number {
  const { min, max } = RESIZE_LIMITS[axis]
  return Number.isFinite(sizeCss) ? Math.min(max, Math.max(min, Math.round(sizeCss))) : min
}

/** An immutable published view of one sheet's session sizes. */
export interface SessionSizes { readonly rows: ReadonlyMap<number, number>; readonly cols: ReadonlyMap<number, number> }
export const NO_SESSION_SIZES: SessionSizes = Object.freeze({ rows: new Map<number, number>(), cols: new Map<number, number>() })
export const hasSessionSizes = (sizes: SessionSizes | undefined): sizes is SessionSizes => Boolean(sizes && (sizes.rows.size || sizes.cols.size))

/** One sheet's session sizes. `gen` changes with every effective change so the
 * viewer can rebuild geometry and refuse plans made for an older layout. */
export class SheetSizes {
  private readonly rows = new Map<number, number>()
  private readonly cols = new Map<number, number>()
  private version = 0
  private published: SessionSizes | undefined
  get gen(): number { return this.version }
  get empty(): boolean { return !this.rows.size && !this.cols.size }
  get count(): number { return this.rows.size + this.cols.size }
  get bytes(): number { return this.count * SESSION_SIZE_ENTRY_BYTES }
  get(axis: ResizeAxis, index: number): number | undefined { return (axis === 'row' ? this.rows : this.cols).get(index) }
  /** Clamped; false when the axis cap refuses a new index. */
  set(axis: ResizeAxis, index: number, sizeCss: number): boolean {
    const map = axis === 'row' ? this.rows : this.cols
    if (!map.has(index) && map.size >= MAX_SESSION_SIZES) return false
    const value = clampSize(axis, sizeCss)
    if (map.get(index) === value) return true
    map.set(index, value)
    this.changed()
    return true
  }
  /** Restores the workbook sizes; false when there was nothing to reset. */
  /** The snapshot this store would publish with one entry set (clamped), or
   * null when the axis cap refuses a new index; nothing changes. */
  preview(axis: ResizeAxis, index: number, sizeCss: number): SessionSizes | null {
    const map = axis === 'row' ? this.rows : this.cols
    if (!map.has(index) && map.size >= MAX_SESSION_SIZES) return null
    const next = new Map(map).set(index, clampSize(axis, sizeCss))
    return Object.freeze(axis === 'row' ? { rows: next, cols: new Map(this.cols) } : { rows: new Map(this.rows), cols: next })
  }
  /** Puts one entry back to `previous` (undefined: removes it), e.g. when the
   * geometry refused a size it cannot represent. */
  restore(axis: ResizeAxis, index: number, previous: number | undefined): void {
    const map = axis === 'row' ? this.rows : this.cols
    if (map.get(index) === previous) return
    if (previous === undefined) map.delete(index)
    else map.set(index, previous)
    this.changed()
  }
  reset(): boolean {
    if (this.empty) return false
    this.rows.clear()
    this.cols.clear()
    this.changed()
    return true
  }
  snapshot(): SessionSizes {
    return this.published ??= Object.freeze({ rows: new Map(this.rows), cols: new Map(this.cols) })
  }
  private changed(): void { this.version++; this.published = undefined }
}

function lowerBound(keys: readonly number[], value: number): number {
  let low = 0, high = keys.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (keys[mid] < value) low = mid + 1
    else high = mid
  }
  return low
}

/** Authored intervals with session sizes applied: authored intervals are split
 * around session indexes (an overlap with a different size would be
 * contradictory axis metadata), a session index inside a zero-size authored
 * interval is ignored, and the rest become one-index intervals. O((A + S) log S). */
export function mergeAxisOverrides(base: readonly AxisSizeInterval[], session: ReadonlyMap<number, number>): AxisSizeInterval[] {
  if (!session.size) return [...base]
  const keys = [...session.keys()].sort((a, b) => a - b)
  const hidden = new Set<number>()
  const out: AxisSizeInterval[] = []
  for (const interval of [...base].sort((a, b) => a.start - b.start || a.end - b.end)) {
    let cursor = interval.start
    for (let at = lowerBound(keys, interval.start); at < keys.length && keys[at] <= interval.end; at++) {
      const key = keys[at]
      if (interval.sizeCss === 0) { hidden.add(key); continue }
      if (cursor < key) out.push({ start: cursor, end: key - 1, sizeCss: interval.sizeCss })
      cursor = key + 1
    }
    if (interval.sizeCss === 0) out.push({ ...interval })
    else if (cursor <= interval.end) out.push({ start: cursor, end: interval.end, sizeCss: interval.sizeCss })
  }
  for (const key of keys) if (!hidden.has(key)) out.push({ start: key, end: key, sizeCss: session.get(key)! })
  return out.sort((a, b) => a.start - b.start || a.end - b.end)
}

// ---- drawings anchored to cells -------------------------------------------

const EMU_PER_CSS = 9525
export interface PlacementAxes { rows: AxisMetrics; cols: AxisMetrics }
/** A logical position inside cell i keeps its offset into that cell, clamped to
 * the cell's session size; past the axis end it shifts by the total change. */
function mapPoint(base: AxisMetrics, sized: AxisMetrics, css: number): number {
  if (!(css > 0)) return css
  const index = base.locate(css)
  if (index === null) return css + sized.totalCss() - base.totalCss()
  return sized.start(index) + Math.min(css - base.start(index), sized.size(index))
}
/** Top-level drawings repositioned from authored to session geometry: a
 * two-cell anchor moves and stretches with its cells, a one-cell anchor moves
 * with its `from` cell and keeps its size, an absolute anchor stays. Nodes are
 * shallow copies (group children keep their child coordinate space, which the
 * group transform maps onto the new box); unchanged nodes are returned as is. */
export function remapDrawings(drawings: readonly XlsxDrawing[], base: PlacementAxes, sized: PlacementAxes): XlsxDrawing[] {
  const result = drawings.map(node => {
    const anchor = node.source?.treePath?.split('/')[0] ?? ''
    const two = anchor.startsWith('twoCellAnchor')
    if (!two && !anchor.startsWith('oneCellAnchor')) return node
    const x = node.xEmu / EMU_PER_CSS, y = node.yEmu / EMU_PER_CSS, w = node.widthEmu / EMU_PER_CSS, h = node.heightEmu / EMU_PER_CSS
    if (![x, y, w, h].every(Number.isFinite)) return node
    const x0 = mapPoint(base.cols, sized.cols, x), y0 = mapPoint(base.rows, sized.rows, y)
    const width = two ? mapPoint(base.cols, sized.cols, x + w) - x0 : w
    const height = two ? mapPoint(base.rows, sized.rows, y + h) - y0 : h
    if (x0 === x && y0 === y && width === w && height === h) return node
    const moved = { ...node, xEmu: x0 * EMU_PER_CSS, yEmu: y0 * EMU_PER_CSS, widthEmu: Math.max(0, width) * EMU_PER_CSS, heightEmu: Math.max(0, height) * EMU_PER_CSS }
    // A group paints its children through ext/chExt: the copied group keeps its
    // child coordinate space (chOff/chExt) and its ext follows the new box, so
    // children move and stretch with the anchor.
    if (node.group) {
      const rx = w > 0 ? moved.widthEmu / node.widthEmu : 1, ry = h > 0 ? moved.heightEmu / node.heightEmu : 1
      moved.group = { ...node.group, off: { x: node.group.off.x + moved.xEmu - node.xEmu, y: node.group.off.y + moved.yEmu - node.yEmu },
        ext: { width: node.group.ext.width * rx, height: node.group.ext.height * ry } }
    }
    return moved
  })
  retainDrawingCopyOwners(result, drawings)
  return result
}

// ---- auto-fit ---------------------------------------------------------------

export type Measure = (font: string, text: string) => number
/** Widest formatted line plus cell padding and the gridline; the default width
 * when nothing was measured. The caller clamps. */
export function fitColumnWidth(cells: Iterable<{ text: string; font: string }>, measure: Measure): number {
  let widest = -1
  for (const cell of cells) for (const line of cell.text.split('\n')) if (line) widest = Math.max(widest, measure(cell.font, line))
  return widest < 0 ? DEFAULT_COL_CSS : Math.ceil(widest + CELL_PADDING_CSS + 1)
}
function wrappedLines(text: string, font: string, inner: number, measure: Measure): number {
  let lines = 0
  for (const paragraph of text.split('\n')) {
    let line = ''
    lines++
    for (const word of paragraph.split(' ')) {
      const candidate = line ? `${line} ${word}` : word
      if (line && measure(font, candidate) > inner) { lines++; line = word }
      else line = candidate
    }
  }
  return lines
}
/** Tallest cell: line count (wrapped at the column width when the cell wraps)
 * × font line height plus padding; the default height when nothing was
 * measured. `heightPx` is a wrapped body's height from the canonical painter
 * layout (§11.2): when present it decides, plus the vertical cell padding. */
export function fitRowHeight(cells: Iterable<{ text: string; font: string; sizePx: number; wrap: boolean; widthCss: number; heightPx?: number }>, measure: Measure): number {
  let tallest = -1
  for (const cell of cells) {
    if (!cell.text) continue
    if (cell.heightPx !== undefined && Number.isFinite(cell.heightPx)) { tallest = Math.max(tallest, Math.ceil(cell.heightPx + CELL_PADDING_CSS)); continue }
    const lines = cell.wrap ? wrappedLines(cell.text, cell.font, Math.max(1, cell.widthCss - CELL_PADDING_CSS), measure) : 1
    tallest = Math.max(tallest, Math.round(lines * cell.sizePx * LINE_HEIGHT + ROW_PADDING_CSS))
  }
  return tallest < 0 ? DEFAULT_ROW_CSS : tallest
}
/** The bounded source rows a column auto-fit reads (see AUTOFIT_MAX_ROWS). */
export function autoFitRowCandidates(visible: Iterable<number>, known: Iterable<number>, isVisible: (row: number) => boolean, limit = AUTOFIT_MAX_ROWS): number[] {
  const seen = new Set<number>()
  for (const row of visible) { if (seen.size >= limit) break; if (isVisible(row)) seen.add(row) }
  let visits = 0
  for (const row of known) {
    if (seen.size >= limit || ++visits > 4 * limit) break
    if (isVisible(row)) seen.add(row)
  }
  return [...seen]
}
/** The painter's font string for a cell style (render.ts layoutCellBody). */
export function cellFont(style: { bold?: boolean; italic?: boolean; fontSizePt?: number; fontName?: string } | undefined): { font: string; sizePx: number } {
  const sizePt = style?.fontSizePt ?? 10
  return { font: cellFontString(style, sizePt), sizePx: sizePt * 96 / 72 }
}
