/**
 * Contour charts (`surfaceChart`, the 2D surface; spec 4.1 contour row and 9 work bounds; lazy chunk).
 *
 * - The z grid is series (rows) x categories (columns). `contourGrid` is the one dimension rule, shared by the
 *   parse and fill diagnostics, the axis graph (domain and axis slots), paint admission and the painter: at most
 *   CONTOUR_ROW_LIMIT rows (the chart's series bound) and CONTOUR_COLUMN_LIMIT columns. A clamp in either
 *   dimension is a visible rendering bound: diagnosed (`chart-contour-grid-limit`) and the model marked truncated.
 * - Bands are the intervals between the resolved value axis's ticks (its min and major unit, automatic or
 *   authored; stage 1 resolves them once, `axes.ts`), so the band count is Office's. A value below the first
 *   tick is in band 0, one at or past the last tick in the last band. Missing cells are never interpolated.
 * - Band colours: `c:bandFmts` when present (a noFill band is not painted), otherwise the document's accent cycle
 *   resolved at parse time.
 * - `wireframe=1` strokes only band boundaries: every neighbouring cell pair (left-right, below-above) is visited
 *   once, and an edge is drawn when its two cells are in different bands, a missing cell counting as its own
 *   band (two missing cells share no boundary). A boundary takes the higher band's line colour (the band whose
 *   lower edge it is; next to a missing cell, the populated cell's band).
 * - Row 0 is the bottom row (the series axis runs up from the category axis), column 0 the left column; a
 *   reversed axis mirrors its direction.
 */
import type { ChartDiagnostic, ChartGroup, ChartModel } from '../model'

/** Rows (series) and columns (categories) drawn (spec 9: contour grid <= 256 x 256; rows are the 64-series bound). */
export const CONTOUR_ROW_LIMIT = 64
export const CONTOUR_COLUMN_LIMIT = 256
export const CONTOUR_GRID_LIMIT = 'chart-contour-grid-limit'
/** The engine's default accent cycle, used only for a hand-built model without parse-time band accents. */
const DEFAULT_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

export interface ContourGrid { rows: number; cols: number; series: number; columns: number; clipped: boolean }
/**
 * The grid dimensions of `series` series over `categories` categories (no categories: the longest value
 * series, `valueExtent`). `clipped` when either dimension is past its limit.
 */
export function contourGrid(series: number, categories: number, valueExtent = 0): ContourGrid {
  const columns = categories > 0 ? categories : valueExtent
  const rows = Math.max(0, Math.min(series, CONTOUR_ROW_LIMIT)), cols = Math.max(0, Math.min(columns, CONTOUR_COLUMN_LIMIT))
  return { rows, cols, series, columns, clipped: series > rows || columns > cols }
}
/** The grid of a model's contour group; `series` overrides the group's (parse-time) series count with the authored one. */
export function contourGridOf(model: Pick<ChartModel, 'categories'>, group: Pick<ChartGroup, 'series'>, series = group.series.length): ContourGrid {
  let extent = 0
  for (const s of group.series) extent = Math.max(extent, s.values.length)
  return contourGrid(series, model.categories.length, extent)
}
export const contourGridDiagnostic = (grid: ContourGrid): ChartDiagnostic => ({
  code: CONTOUR_GRID_LIMIT,
  reason: `Contour grid of ${grid.series} series x ${grid.columns} categories drawn as ${grid.rows} x ${grid.cols} (at most ${CONTOUR_ROW_LIMIT} x ${CONTOUR_COLUMN_LIMIT})`,
})
/** Paint scratch of a contour group: one Int16 band per grid cell (plus the typed array's header). */
export const contourPaintBytes = (model: Pick<ChartModel, 'categories'>, group: Pick<ChartGroup, 'series'>): number => {
  const grid = contourGridOf(model, group)
  return 64 + grid.rows * grid.cols * 2
}

/** The band of value `z` on ascending `ticks`: the last tick at or below it, clamped to the bands (band count ticks - 1). */
export function contourBand(ticks: readonly number[], z: number): number {
  const last = ticks.length - 2
  if (last <= 0 || !(z >= ticks[1])) return 0
  let lo = 1, hi = last
  while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (ticks[mid] <= z) lo = mid; else hi = mid - 1 }
  return lo
}
/**
 * The band edges of a resolved value axis: its ticks, plus its maximum when the scale ends past the last tick (an
 * authored minimum or maximum off the major grid leaves a last partial band). Band count: edges - 1.
 */
export function contourEdges(axis: { ticks: readonly number[]; min: number; max: number }): number[] {
  const ticks = axis.ticks.length ? [...axis.ticks] : [axis.min]
  if (axis.max > ticks[ticks.length - 1] + Math.abs(axis.max) * 1e-12) ticks.push(axis.max)
  return ticks
}

type BandStyled = Pick<ChartGroup, 'bandColors' | 'bandLines' | 'bandAccents'>
const accent = (group: BandStyled, band: number) => { const cycle = group.bandAccents?.length ? group.bandAccents : DEFAULT_ACCENTS; return cycle[band % cycle.length] }
/** A band's fill colour (hex), or undefined for an authored noFill band. */
export function contourBandFill(group: BandStyled, band: number): string | undefined {
  const authored = group.bandColors?.[band]
  return authored === null ? undefined : authored ?? accent(group, band)
}
/** A band's wireframe line colour (hex): its authored line, else its authored fill, else the accent; undefined for an authored noFill line. */
export function contourBandLine(group: BandStyled, band: number): string | undefined {
  const line = group.bandLines?.[band]
  return line === null ? undefined : line ?? group.bandColors?.[band] ?? accent(group, band)
}
/** Units kept per contour legend name (the chart label limit, spec 7); a longer name is cut and diagnosed. */
export const CONTOUR_LEGEND_NAME_LIMIT = 255
export const CONTOUR_LEGEND_TRUNCATED = 'chart-legend-truncated'
/**
 * The bytes one legend entry can need while it is built and kept: its record, its name (at most the limit) and
 * one transient endpoint cut of at most the limit, priced as UTF-16 plus 32 (`stringBytes`), plus its measured width.
 */
export const CONTOUR_LEGEND_ENTRY_BYTES = 64 + 2 * (CONTOUR_LEGEND_NAME_LIMIT * 2 + 32) + (64 + CONTOUR_LEGEND_NAME_LIMIT * 2)
/**
 * Legend entries: one per band, named by its tick interval (the value axis's tick labels), in its band colour.
 * At most `limit` entries (the rows a legend can show) are built; each name is cut to CONTOUR_LEGEND_NAME_LIMIT
 * units before it is concatenated, and `truncated` is called when a name was cut.
 */
export function contourLegendEntries(group: BandStyled & Pick<ChartGroup, 'wireframe'>, ticks: readonly number[], labels: readonly string[], format: (edge: number) => string = String, limit = Infinity, truncated?: () => void): Array<{ name: string; color?: string }> {
  const out: Array<{ name: string; color?: string }> = []
  const max = CONTOUR_LEGEND_NAME_LIMIT
  for (let band = 0; band < ticks.length - 1 && out.length < limit; band++) {
    const color = group.wireframe ? contourBandLine(group, band) : contourBandFill(group, band)
    const low = labels[band] ?? format(ticks[band]), high = labels[band + 1] ?? format(ticks[band + 1])
    let name: string
    if (low.length + 1 + high.length <= max) name = `${low}-${high}`
    else { truncated?.(); name = low.length >= max ? low.slice(0, max) : `${low}-${high.slice(0, max - 1 - low.length)}` }
    out.push({ name, ...(color ? { color } : {}) })
  }
  return out
}

interface Box { x: number; y: number; w: number; h: number }
interface ContourSeries { value(slot: number): number | undefined }
export interface ContourPaint {
  group: BandStyled & Pick<ChartGroup, 'wireframe'> & { series: readonly ContourSeries[] }
  grid: ContourGrid
  /** Display slots on the category axis (hidden categories leave it). */
  slots: number
  ticks: readonly number[]
  categoryReversed: boolean
  seriesReversed: boolean
  plot: Box
}
/** The rows and columns a paint draws (within the grid, so admission by `contourPaintBytes` covers its scratch). */
export const contourPainted = (grid: ContourGrid, seriesShown: number, slots: number) => ({ rows: Math.min(grid.rows, seriesShown), cols: Math.max(1, Math.min(grid.cols, slots)) })

export function paintContour(ctx: CanvasRenderingContext2D, p: ContourPaint): void {
  const { group, plot, ticks } = p
  const { rows, cols } = contourPainted(p.grid, group.series.length, p.slots)
  if (!rows || !cols) return
  // One band per cell (-1: missing), admitted as paint scratch before the paint.
  const bands = new Int16Array(rows * cols)
  for (let r = 0; r < rows; r++) {
    const s = group.series[r]
    for (let c = 0; c < cols; c++) { const v = s.value(c); bands[r * cols + c] = v !== undefined && Number.isFinite(v) ? contourBand(ticks, v) : -1 }
  }
  const cw = plot.w / cols, ch = plot.h / rows
  /** Cell edges: column edge k (0..cols) and row edge k (0..rows, row 0 at the bottom). */
  const xEdge = (k: number) => plot.x + (p.categoryReversed ? cols - k : k) * cw
  const yEdge = (k: number) => plot.y + plot.h - (p.seriesReversed ? rows - k : k) * ch
  if (!group.wireframe) {
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const band = bands[r * cols + c]
      if (band < 0) continue
      const color = contourBandFill(group, band)
      if (!color) continue
      ctx.fillStyle = `#${color}`
      ctx.fillRect(Math.min(xEdge(c), xEdge(c + 1)), Math.min(yEdge(r), yEdge(r + 1)), cw, ch)
    }
    return
  }
  ctx.lineWidth = 1
  const edge = (a: number, b: number, x0: number, y0: number, x1: number, y1: number) => {
    if (a === b) return
    const color = contourBandLine(group, Math.max(a, b))
    if (!color) return
    ctx.strokeStyle = `#${color}`
    ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.stroke()
  }
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    const band = bands[r * cols + c]
    // The right neighbour and the neighbour above: each pair once.
    if (c + 1 < cols) edge(band, bands[r * cols + c + 1], xEdge(c + 1), yEdge(r), xEdge(c + 1), yEdge(r + 1))
    if (r + 1 < rows) edge(band, bands[(r + 1) * cols + c], xEdge(c), yEdge(r + 1), xEdge(c + 1), yEdge(r + 1))
  }
}
