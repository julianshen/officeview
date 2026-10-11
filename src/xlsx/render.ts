/**
 * XLSX rendering: paint sheets as a grid onto a canvas 2D context.
 * Layout mimics Excel defaults: default col width 8.43 chars (~64px), row
 * height 15pt (20px). Cell geometry is fixed; no reflow needed.
 */
import type { XlsxCell, XlsxCellStyle, XlsxDrawing, XlsxMergeRange, XlsxSheet } from './types'
import type { PptxTextBody } from '../pptx/types'
import { layoutTextBody, type MeasureText, type TextLayout } from '../drawing/text-layout'
import { resolveColor } from '../core/color'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { paintScene } from '../drawing/scene-paint'
import { paintDrawingContent } from '../drawing/content-paint'
import type { ContentPaintAssets } from '../drawing/content'
import { createTextBodyMeasurer, paintTextBody } from '../drawing/text-paint'
import { graphemes } from '../core/text-recording'
import { isFormulaError, isEvaluationError } from './formula/evaluator'
import { formatCommon } from './formula/format'
import { decodeSerial } from './formula/serial'
import type { ResolvedSemantics } from './formula/types'
import type { WorkerDiagnostic } from './view/contracts'
import { MAX_COLS, isValidCoordinate } from './view/coordinates'
import { RectangleIndex } from './view/structure'
import { cellSourceRow, drawingExtent } from './view/sheetCells'
import type { CellBodyLayout } from './view/textLayout'

export interface GridMetrics {
  colWidthsPx: number[]
  rowHeightsPx: number[]
  widthPx: number
  heightPx: number
  requestedDrawingBounds?: { right: number; bottom: number }
  retainedDrawingBounds?: { right: number; bottom: number }
  /** Typed static-raster refusals: populated cells outside the retained grid
   * (spec §11.5). Present only when the static bitmap cannot hold the sheet. */
  exportLimits?: readonly StaticExportLimit[]
  /** Metadata resolved conservatively (for example conflicting duplicate rows). */
  diagnostics?: readonly WorkerDiagnostic[]
}

/** One explicit static-output budget that a request exceeded (spec §11.5). */
export interface StaticExportLimit { code: 'large-export-limit'; resource: 'rows' | 'cols' | 'pixels' | 'raster-bytes' | 'pages'; requested: number; limit: number; message: string }
/** Raised before painting when a static raster/print request exceeds an explicit
 * budget. Static output never aliases or silently drops logical coordinates. */
export class StaticExportLimitError extends RangeError {
  readonly code = 'large-export-limit'
  constructor(readonly limits: readonly StaticExportLimit[]) {
    super(limits.map(limit => limit.message).join('; '))
    this.name = 'StaticExportLimitError'
  }
}

const DEFAULT_COL_PX = 64
const DEFAULT_ROW_PX = 20
const CHAR_PX = 7
/** Hard caps so a pathological sheet can never allocate an impossible canvas. */
const MAX_GRID_COLS = 4096
const MAX_GRID_ROWS = 16384
/** Map an OOXML alignment textRotation to canvas paint behavior. Values 1-90 are
 * counterclockwise degrees; 91-180 encode clockwise 1-90 (180 is 90 degrees
 * clockwise, not a half turn); 255 stacks glyphs with RTL column progression.
 * Absent, 0 and invalid values stay undefined (horizontal). */
export function cellTextRotation(rotation: number | undefined): { radians: number } | { stacked: true } | undefined {
  if (rotation === undefined || rotation === 0) return undefined
  if (rotation === 255) return { stacked: true }
  if (!Number.isInteger(rotation) || rotation < 1 || rotation > 180) return undefined
  return { radians: rotation <= 90 ? (-rotation * Math.PI) / 180 : ((rotation - 90) * Math.PI) / 180 }
}

/** Measured glyph bands also cover the canonical recording band, so fitting
 * cell text has the same physical and selectable bounds on every Canvas host. */
function cellTextMeasurer(ctx: CanvasRenderingContext2D): MeasureText {
  const native = createTextBodyMeasurer(ctx, family => family)
  return (text, style) => {
    const m = native(text, style), size = (style.fontSizePt ?? 10) * 96 / 72
    const ascent = Math.max(m.ascent ?? 0, size * .85), descent = Math.max(m.descent ?? 0, size * .25)
    return { width: m.width, ascent, descent, normalHeight: Math.max(m.normalHeight ?? 0, ascent + descent) }
  }
}
function rotatedLayout(layout: TextLayout, radians: number): TextLayout {
  const a = Math.cos(radians), b = Math.sin(radians)
  return { ...layout, lines: layout.lines.map(line => ({ ...line, segments: line.segments.map(segment => {
    const t = segment.transform ?? { a: 1, b: 0, c: 0, d: 1, e: segment.x, f: line.baseline }
    return { ...segment, transform: { a: a * t.a - b * t.b, b: b * t.a + a * t.b,
      c: a * t.c - b * t.d, d: b * t.c + a * t.d, e: a * t.e - b * t.f, f: b * t.e + a * t.f } }
  }) })) }
}
function textLayoutBounds(layout: TextLayout, measure: MeasureText): { x: number; y: number; width: number; height: number } {
  const points: Array<{ x: number; y: number }> = []
  for (const line of layout.lines) for (const segment of line.segments) {
    if (segment.text === '\n') continue
    const t = segment.transform ?? { a: 1, b: 0, c: 0, d: 1, e: segment.x, f: line.baseline }
    const m = measure('Mg', segment.style), ascent = m.ascent ?? 0, descent = m.descent ?? 0
    for (const [x, y] of [[0, -ascent], [segment.width, -ascent], [segment.width, descent], [0, descent]])
      points.push({ x: t.a * x + t.c * y + t.e, y: t.b * x + t.d * y + t.f })
  }
  if (!points.length) return { x: 0, y: 0, width: 0, height: 0 }
  const x = Math.min(...points.map(p => p.x)), y = Math.min(...points.map(p => p.y))
  return { x, y, width: Math.max(...points.map(p => p.x)) - x, height: Math.max(...points.map(p => p.y)) - y }
}
function alignedLayout(layout: TextLayout, bounds: ReturnType<typeof textLayoutBounds>, box: { x: number; y: number; w: number; h: number }, horizontal: string, vertical: string): TextLayout {
  // Preserve font size. If a block fits, relax only nominal padding where the
  // em-cell capacity fits but its measured glyph band uses the last fraction.
  const px = Math.min(PADDING_L, Math.max(0, (box.w - bounds.width) / 2))
  const py = Math.min(PADDING_L, Math.max(0, (box.h - bounds.height) / 2))
  const x = horizontal === 'right' ? box.x + box.w - px - bounds.width : horizontal === 'center' ? box.x + (box.w - bounds.width) / 2 : box.x + px
  const y = vertical === 'bottom' ? box.y + box.h - py - bounds.height : vertical === 'center' || vertical === 'middle' ? box.y + (box.h - bounds.height) / 2 : box.y + py
  const dx = x - bounds.x, dy = y - bounds.y
  return { ...layout, lines: layout.lines.map(line => ({ ...line, segments: line.segments.map(segment => ({ ...segment,
    transform: segment.transform ? { ...segment.transform, e: segment.transform.e + dx, f: segment.transform.f + dy } : { a: 1, b: 0, c: 0, d: 1, e: segment.x + dx, f: line.baseline + dy } })) })) }
}
/** Wrapped text chooses a measured local multiline box that fits its rotated
 * cell interior. No-wrap or physically oversized text keeps authored font size
 * and clips to the cell; source remains complete, with clip-aware indexing.
 * Layout and paint are separate so the viewer can cache the layout (computed
 * in a box at the origin) and replay it; the static path runs both back to back
 * at absolute coordinates, exactly as before. */
function layoutCellText(ctx: CanvasRenderingContext2D, text: string, style: XlsxCellStyle | undefined, numeric: boolean, x: number, y: number, w: number, h: number, color: string, center = false): Extract<NonNullable<CellBodyLayout['text']>, { kind: 'layout' }> {
  const mapped = cellTextRotation(style?.textRotation), stacked = mapped && 'stacked' in mapped
  const radians = mapped && 'radians' in mapped ? mapped.radians : 0
  const defaultHorizontal = numeric ? 'right' : center ? 'center' : 'left'
  const horizontal = style?.horizontal === undefined || style.horizontal === 'general' ? defaultHorizontal : style.horizontal
  const vertical = style?.vertical ?? 'center'
  const wrap = style?.wrapText ?? !!stacked
  const body: PptxTextBody = { direction: stacked ? 'wordArtVertRtl' : 'horz', anchor: 't', wrap,
    insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
    paragraphs: [{ runs: [{ text, fontSizePt: style?.fontSizePt ?? 10, fontFamily: 'Calibri', bold: style?.bold, italic: style?.italic, color }], align: 'left', level: 0,
      defaultProperties: { fontSizePt: style?.fontSizePt ?? 10, fontFamily: 'Calibri', bold: style?.bold, italic: style?.italic, color }, wrapWhitespace: true,
      ...(stacked ? {} : { wordLineSpacing: { rule: 'auto', value: 240 } }) }] }
  ctx.save()
  try {
    const measure = cellTextMeasurer(ctx)
    let layout: TextLayout
    if (stacked) {
      const run = body.paragraphs[0].runs[0]
      const metrics = measure('Mg', run)
      const pitch = Math.max((run.fontSizePt ?? 10) * 96 / 72, metrics.normalHeight ?? 0, ...graphemes(text).map(g => measure(g.text, run).width))
      body.paragraphs[0].lineSpacing = { kind: 'points', value: pitch * 72 / 96 }
      layout = layoutTextBody(body, w, h, measure)
    } else {
      const cosine = Math.abs(Math.cos(radians)), sine = Math.abs(Math.sin(radians))
      const innerW = Math.max(1, w - PADDING_L - PADDING_R), innerH = Math.max(1, h - 2 * PADDING_L)
      const maxFlow = Math.min(cosine > 1e-10 ? innerW / cosine : Infinity, sine > 1e-10 ? innerH / sine : Infinity)
      const advance = measure(text, body.paragraphs[0].runs[0]).width
      const metrics = measure('Mg', body.paragraphs[0].runs[0])
      const bandHeight = (metrics.ascent ?? 0) + (metrics.descent ?? 0), lineHeight = metrics.normalHeight ?? bandHeight
      const maxLocalHeight = Math.min(sine > 1e-10 ? innerW / sine : Infinity, cosine > 1e-10 ? innerH / cosine : Infinity)
      const count = Math.max(1, Math.min(graphemes(text).length, 1 + Math.floor(Math.max(0, maxLocalHeight - bandHeight) / Math.max(1, lineHeight))))
      // Finite candidates derived from source length/advance, not a giant flow
      // cap or calibrated font-size adjustment. Select the largest fitting box.
      const widths = wrap ? [...new Set([maxFlow, ...Array.from({ length: count }, (_, i) => {
        const localHeight = i * lineHeight + bandHeight
        return Math.min(cosine > 1e-10 ? (innerW - sine * localHeight) / cosine : maxFlow,
          sine > 1e-10 ? (innerH - cosine * localHeight) / sine : maxFlow)
      }), ...Array.from({ length: count }, (_, i) => Math.min(maxFlow, advance / (i + 1)))])].filter(width => width > 0).sort((a, b) => b - a) : [Math.max(1, advance)]
      layout = rotatedLayout(layoutTextBody(body, widths[0], h, measure), radians)
      for (const width of widths) {
        const candidate = rotatedLayout(layoutTextBody(body, width, h, measure), radians)
        const bounds = textLayoutBounds(candidate, measure)
        if (bounds.width <= innerW + 1e-8 && bounds.height <= innerH + 1e-8) { layout = candidate; break }
      }
    }
    const bounds = textLayoutBounds(layout, measure)
    const placed = alignedLayout(layout, bounds, { x, y, w, h }, horizontal, vertical)
    // Coordinates in the placed layout are box-local to (x, y) and already include
    // the final merged bounds. Paint and recording replay this exact geometry.
    return { kind: 'layout', text, font: '', color, body, layout: placed, clip: { x, y, width: w, height: h },
      ...(stacked ? { stacked: true } : radians ? { rotation: radians } : {}), wrap, lines: placed.lines.length }
  } finally { ctx.restore() }
}
function paintCellText(ctx: CanvasRenderingContext2D, text: Extract<NonNullable<CellBodyLayout['text']>, { kind: 'layout' }>): void {
  ctx.save()
  try {
    paintTextBody(text.body, ctx, 0, 0, text.clip.width, text.clip.height, family => family, undefined, { layout: text.layout, clip: text.clip })
  } finally { ctx.restore() }
}
/** How much wider than the used range a <col> declaration may be and still
 * count as real formatting rather than noise. */
const MAX_DECLARED_COL_SLACK = 64
/** The legacy drawing-path raster clamp: 16,384 px per axis and 16,777,216 px
 * (64 MiB RGBA). A request beyond it used to be cropped silently and is now a typed
 * refusal. Drawing-free sheets were never clamped by pixels: within the legacy grid
 * caps they keep rendering in full (spec §11.5 preserves admitted static output). */
const MAX_RASTER_SIDE = 16384
const MAX_RASTER_BYTES = 64 * 1024 * 1024
const MAX_DRAWING_SIDE = MAX_RASTER_SIDE
const MAX_DRAWING_AREA = MAX_RASTER_BYTES / 4
/** A full-sheet print run beyond this many pages needs an explicit bounded range. */
const MAX_PRINT_PAGES = 10_000
/** Explicit static (full-sheet bitmap/print) budgets. The interactive viewer is not
 * bound by them: it paints a bounded viewport over the full logical axes. */
export const STATIC_SHEET_BUDGET = Object.freeze({
  maxGridRows: MAX_GRID_ROWS,
  maxGridCols: MAX_GRID_COLS,
  maxDrawingSidePx: MAX_DRAWING_SIDE,
  maxDrawingAreaPx: MAX_DRAWING_AREA,
  maxPrintPages: MAX_PRINT_PAGES,
})

function validMerge(r: XlsxMergeRange): boolean {
  return isValidCoordinate(r.minRow, r.minCol) && isValidCoordinate(r.maxRow, r.maxCol) && r.minRow <= r.maxRow && r.minCol <= r.maxCol
}

/** Populated cells and merge extents the retained static grid cannot place. Before
 * this check the renderer drew cells at `rowY[row] ?? 0` / `colX[col] ?? 0`, i.e.
 * over A1, and clipped merges to the bitmap edge. */
function gridCoverageLimits(sheet: XlsxSheet, rows: number, cols: number): StaticExportLimit[] {
  let maxRow = -1, maxCol = -1
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      const source = cellSourceRow(row, cell)
      if (!Number.isSafeInteger(source) || source < 0) continue
      maxRow = Math.max(maxRow, source)
      if (Number.isSafeInteger(cell.col) && cell.col >= 0) maxCol = Math.max(maxCol, cell.col)
    }
  }
  for (const r of sheet.mergeRanges ?? []) {
    if (validMerge(r)) { maxRow = Math.max(maxRow, r.maxRow); maxCol = Math.max(maxCol, r.maxCol) }
  }
  const limits: StaticExportLimit[] = []
  if (maxRow + 1 > rows) limits.push({ code: 'large-export-limit', resource: 'rows', requested: maxRow + 1, limit: rows, message: `Row ${maxRow + 1} is beyond the ${rows}-row static grid; render a bounded viewport instead` })
  if (maxCol + 1 > cols) limits.push({ code: 'large-export-limit', resource: 'cols', requested: maxCol + 1, limit: cols, message: `Column ${maxCol + 1} is beyond the ${cols}-column static grid; render a bounded viewport instead` })
  return limits
}

/** Pixel/byte admission for one full-sheet drawing-bearing raster: exactly the
 * requests the legacy clamp would have truncated. */
function rasterLimits(width: number, height: number): StaticExportLimit[] {
  const w = Math.ceil(width), h = Math.ceil(height)
  const limits: StaticExportLimit[] = []
  const side = Math.max(w, h)
  if (side > MAX_RASTER_SIDE) limits.push({ code: 'large-export-limit', resource: 'pixels', requested: side, limit: MAX_RASTER_SIDE, message: `A ${w}x${h}px sheet image exceeds the ${MAX_RASTER_SIDE}px static raster side; render a bounded viewport instead` })
  if (w * h * 4 > MAX_RASTER_BYTES) limits.push({ code: 'large-export-limit', resource: 'raster-bytes', requested: w * h * 4, limit: MAX_RASTER_BYTES, message: `A ${w}x${h}px sheet image needs ${w * h * 4} bytes, over the ${MAX_RASTER_BYTES}-byte static raster budget` })
  return limits
}
/** Raster limits of a full-sheet request; only a clamped drawing bitmap can truncate. */
function requestedRasterLimits(m: GridMetrics): StaticExportLimit[] {
  return m.requestedDrawingBounds ? rasterLimits(m.requestedDrawingBounds.right, m.requestedDrawingBounds.bottom) : []
}

function transformedDrawingBounds(nodes: XlsxDrawing[]): { right: number; bottom: number } {
  const extent = drawingExtent(nodes)
  return { right: Math.max(0, extent.right), bottom: Math.max(0, extent.bottom) }
}

export function computeMetrics(sheet: XlsxSheet): GridMetrics {
  const merges = (sheet.mergeRanges ?? []).filter(validMerge)
  let maxCol = 0
  for (const row of sheet.rows) for (const c of row.cells) maxCol = Math.max(maxCol, c.col)
  // Merged ranges are part of the painted grid even where only the anchor is populated.
  for (const r of merges) maxCol = Math.max(maxCol, r.maxCol)
  // Real spreadsheets declare column formatting far wider than the used range
  // (a sheet with data in A1:C10 may carry <col max="1025"/>). Sizing the grid
  // from such a declared range alone produced a 66,000px canvas and threw, so
  // trust the used range when the declared one is implausibly wider. A modestly
  // wider declaration is real formatting (a styled but empty column) and stays.
  const drawing = !!sheet.drawings?.length
  const declaredMax = sheet.cols.reduce((m, c) => drawing && !Number.isSafeInteger(c.max) ? m : Math.max(m, c.max), maxCol)
  const gridMax = declaredMax > maxCol + MAX_DECLARED_COL_SLACK ? maxCol : declaredMax
  const nCols = Math.min(Math.max(gridMax + 1, (drawing ? (sheet.drawingMarkers?.maxCol ?? 0) + 1 : 1), 1), MAX_GRID_COLS)
  const colWidthsPx = new Array<number>(nCols).fill(DEFAULT_COL_PX)
  for (const spec of sheet.cols) {
    // Authored-hidden columns collapse unconditionally (spec §5.3), not only on
    // drawing-bearing sheets.
    if (spec.widthChars !== undefined || spec.hidden) {
      const w = spec.hidden ? 0 : drawing && (spec.widthChars === undefined || !Number.isFinite(spec.widthChars) || spec.widthChars < 0) ? DEFAULT_COL_PX : Math.round(spec.widthChars! * CHAR_PX + 5)
      for (let c = Math.max(0, spec.min); c <= spec.max && c < nCols; c++) colWidthsPx[c] = w
    }
  }
  const diagnostics: WorkerDiagnostic[] = []
  let maxRow = 0, inconsistent = 0
  for (const row of sheet.rows) if (!drawing || Number.isSafeInteger(row.index) && row.index >= 0) maxRow = Math.max(row.index, maxRow)
  // An explicit cell reference names its row even when the row record does not.
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      const source = cellSourceRow(row, cell)
      if (source === row.index || !Number.isSafeInteger(source) || source < 0) continue
      inconsistent++
      maxRow = Math.max(maxRow, source)
    }
  }
  if (inconsistent) diagnostics.push({ code: 'inconsistent-row-coordinate', message: `${inconsistent} cell(s) carry a reference outside their row record; placed by the cell reference` })
  for (const r of merges) maxRow = Math.max(maxRow, r.maxRow)
  const nRows = Math.min(Math.max(maxRow + 1, drawing ? (sheet.drawingMarkers?.maxRow ?? 0) + 1 : 1), MAX_GRID_ROWS)
  const rowHeightsPx = new Array<number>(nRows).fill(DEFAULT_ROW_PX)
  // Duplicate records for one index are invalid OOXML (rows without `r` currently
  // all parse at index 0). Collapse only when every record is hidden so a stray
  // hidden duplicate never hides visible data; conflicts are diagnosed.
  const rowHidden = new Map<number, boolean>()
  for (const row of sheet.rows) rowHidden.set(row.index, (rowHidden.get(row.index) ?? true) && row.hidden === true)
  const conflicting = new Set(sheet.rows.filter(row => row.hidden === true && rowHidden.get(row.index) === false).map(row => row.index))
  if (conflicting.size) diagnostics.push({ code: 'conflicting-row-metadata', message: `${conflicting.size} row index(es) carry both hidden and visible records; kept visible` })
  for (const row of sheet.rows) {
    if (row.index >= 0 && row.index < nRows) {
      if (rowHidden.get(row.index)) rowHeightsPx[row.index] = 0
      else if (row.heightPt !== undefined && row.customHeight) rowHeightsPx[row.index] = drawing && (!Number.isFinite(row.heightPt) || row.heightPt < 0) ? DEFAULT_ROW_PX : Math.round(row.heightPt * (96 / 72))
    }
  }
  const gridWidth = colWidthsPx.reduce((a, b) => a + b, 0), gridHeight = rowHeightsPx.reduce((a, b) => a + b, 0)
  const extras = (raster: StaticExportLimit[]) => {
    const exportLimits = [...gridCoverageLimits(sheet, nRows, nCols), ...raster]
    return {
      ...(exportLimits.length ? { exportLimits } : {}),
      ...(diagnostics.length ? { diagnostics } : {}),
    }
  }
  if (!drawing) return { colWidthsPx, rowHeightsPx, widthPx: gridWidth, heightPx: gridHeight, ...extras([]) }
  const visual = transformedDrawingBounds(sheet.drawings!)
  const requested = { right: Math.max(gridWidth, visual.right), bottom: Math.max(gridHeight, visual.bottom) }
  // The retained bitmap stays clamped (parse-time coverage reporting relies on
  // it), but a clamped request is refused, never painted as if complete.
  let width = Math.max(1, Math.min(MAX_DRAWING_SIDE, Math.ceil(requested.right)))
  let height = Math.max(1, Math.min(MAX_DRAWING_SIDE, Math.ceil(requested.bottom)))
  if (width * height > MAX_DRAWING_AREA) height = Math.max(1, Math.floor(MAX_DRAWING_AREA / width))
  return {
    colWidthsPx,
    rowHeightsPx,
    widthPx: width,
    heightPx: height,
    requestedDrawingBounds: requested,
    retainedDrawingBounds: { right: width, bottom: height },
    ...extras(rasterLimits(requested.right, requested.bottom)),
  }
}

/** OOXML paper size ids to portrait width/height in inches (Letter default).
 * A/B sizes are JIS (B4 257x364mm, B5 182x257mm); Folio/Quarto are US
 * 8.5x13in and 215x275mm. ECMA-376 names ids only, so dimensions follow
 * printer convention; ids beyond 15 fall back to Letter. */
export const PAPER_SIZES_IN: Record<number, [number, number]> = {
  1: [8.5, 11], 2: [8.5, 11], 3: [11, 17], 4: [17, 11], 5: [8.5, 14],
  6: [5.5, 8.5], 7: [7.25, 10.5], 8: [11.69, 16.54], 9: [8.27, 11.69],
  10: [8.27, 11.69], 11: [5.83, 8.27], 12: [10.12, 14.33], 13: [7.17, 10.12],
  14: [8.5, 13], 15: [8.47, 10.83],
}
export interface PrintMetrics {
  paperPx: { width: number; height: number }
  printable: { x: number; y: number; width: number; height: number }
  /** Effective content scale (explicit scale, possibly reduced by fit-to-page). */
  scale: number
  pagesWide: number
  pagesTall: number
  /** Full-run refusals (grid coverage or page quota). A caller printing every page
   * must honor them; a single bounded page is admitted on its own coverage. */
  exportLimits?: readonly StaticExportLimit[]
}
/**
 * Print geometry for one sheet: paper rect, printable rect inside the
 * margins, effective scale and page counts. Excel defaults (Letter, portrait,
 * 100%, 0.7/0.75in margins) apply wherever the sheet authors nothing.
 */
export function computePrintMetrics(sheet: XlsxSheet, grid: GridMetrics, dpi = 96): PrintMetrics {
  const setup = sheet.pageSetup
  const [pwIn, phIn] = PAPER_SIZES_IN[setup?.paperSizeId ?? 1] ?? PAPER_SIZES_IN[1]
  const landscape = setup?.orientation === 'landscape'
  const paperPx = {
    width: (landscape ? phIn : pwIn) * dpi,
    height: (landscape ? pwIn : phIn) * dpi,
  }
  const margins = sheet.pageMargins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
  const printable = {
    x: margins.left * dpi,
    y: margins.top * dpi,
    width: Math.max(0, paperPx.width - (margins.left + margins.right) * dpi),
    height: Math.max(0, paperPx.height - (margins.top + margins.bottom) * dpi),
  }
  const base = Math.min(4, Math.max(0.1, (setup?.scale ?? 100) / 100))
  let scale = base
  if (setup?.fitToPage) {
    const fw = setup.fitToWidth ?? 0, fh = setup.fitToHeight ?? 0
    const sx = fw > 0 && grid.widthPx > 0 ? (printable.width * fw) / grid.widthPx : Infinity
    const sy = fh > 0 && grid.heightPx > 0 ? (printable.height * fh) / grid.heightPx : Infinity
    scale = Math.min(base, sx, sy)
    if (!Number.isFinite(scale) || scale <= 0) scale = base
  }
  const pagesWide = Math.max(1, Math.ceil((grid.widthPx * scale) / Math.max(1, printable.width)))
  const pagesTall = Math.max(1, Math.ceil((grid.heightPx * scale) / Math.max(1, printable.height)))
  const pages = pagesWide * pagesTall
  // A print run is many bounded pages, not one raster: only grid coverage and the
  // page quota apply. Individual pages are admitted by renderPrintPage.
  const exportLimits = [...(grid.exportLimits ?? []).filter(limit => limit.resource === 'rows' || limit.resource === 'cols'),
    ...(pages > MAX_PRINT_PAGES ? [{ code: 'large-export-limit' as const, resource: 'pages' as const, requested: pages, limit: MAX_PRINT_PAGES,
      message: `Printing ${pages} pages exceeds the ${MAX_PRINT_PAGES}-page static budget; print a bounded range instead` }] : [])]
  return { paperPx, printable, scale, pagesWide, pagesTall, ...(exportLimits.length ? { exportLimits } : {}) }
}
/**
 * Paint one print page: paper-white sheet, content scaled into the printable
 * area, clipped. Page (pageCol, pageRow) selects its sheet-coordinate window;
 * content beyond one page clips (multi-page paintables are a follow-up).
 */
export interface RenderSheetAssets {
  images?: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont?: (family: string) => string
  imageFor?: ContentPaintAssets['imageFor']
}

export function renderPrintPage(
  sheet: XlsxSheet,
  ctx: CanvasRenderingContext2D,
  metrics: GridMetrics,
  print: PrintMetrics,
  pageCol = 0,
  pageRow = 0,
  watermark?: WatermarkOptions | ResolvedWatermark,
  prepared?: RenderSheetAssets,
): void {
  // One page is a bounded request: renderSheet refuses it only when its window
  // reaches coordinates the static grid cannot place (spec §11.5).
  if (!Number.isSafeInteger(pageCol) || !Number.isSafeInteger(pageRow) || pageCol < 0 || pageRow < 0
    || pageCol >= print.pagesWide || pageRow >= print.pagesTall) throw new RangeError('Print page is outside the page grid')
  const page = {
    x: (pageCol * print.printable.width) / print.scale,
    y: (pageRow * print.printable.height) / print.scale,
    width: print.printable.width / print.scale,
    height: print.printable.height / print.scale,
  }
  const limits = staticAdmission(sheet, metrics, page)
  if (limits.length) throw new StaticExportLimitError(limits)
  ctx.save()
  try {
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, print.paperPx.width, print.paperPx.height)
    // Watermark belongs to the paper, not the sheet: paint it here in paper
    // coordinates (under the cells) instead of forwarding it into sheet space,
    // where the viewport translation would anchor it to the sheet.
    if (watermark) paintWatermark(ctx, { widthPx: print.paperPx.width, heightPx: print.paperPx.height }, watermark)
    const { printable, scale } = print
    ctx.beginPath()
    ctx.rect(printable.x, printable.y, printable.width, printable.height)
    ctx.clip()
    ctx.translate(printable.x, printable.y)
    ctx.scale(scale, scale)
    renderSheet(sheet, ctx, metrics, undefined, prepared, page)
  } finally {
    ctx.restore()
  }
}

/** Default render semantics when a hand-built sheet carries none. */
const DEFAULT_RENDER_SEMANTICS: ResolvedSemantics = {
  dateSystem: '1900',
  unicode: { version: 2, source: 'standalone-default' },
  locale: 'en-US',
  timeZone: 'UTC',
  epochNowMs: 0,
}

/** One private fallback shared by the default numFmt path and the
 * unsupported-format fallback. `Math.round(value * 100) / 100` can overflow to
 * Infinity for a finite source, so when that intermediate is not finite we keep
 * the source magnitude (rawString) rather than emitting Infinity/NaN. */
function formatFiniteFallback(value: number): string {
  const rounded = Math.round(value * 100) / 100
  return Number.isFinite(rounded) ? String(rounded) : String(value)
}

/** Native General controls use at most eleven unsigned display characters.
 * Short exact decimals retain their zeros; longer small numbers and numbers
 * with twelve integer digits use scientific notation. Cell painting supplies
 * a physical width test so the display can lose precision without changing
 * the stored value. This finite profile does not implement locale General. */
function formatGeneral(value: number, fits: (text: string) => boolean = () => true): string {
  if (!Number.isFinite(value)) return String(value)
  if (value === 0) return '0'
  const negative = value < 0, absolute = Math.abs(value)
  const sign = negative ? '-' : ''
  const raw = String(absolute)
  // Expand only small exponents which could fit the plain-character budget.
  const plain = raw.includes('e') && absolute < 1 && absolute >= 1e-9
    ? '0.' + '0'.repeat(-Number(raw.split('e')[1]) - 1) + raw.split('e')[0].replace('.', '')
    : raw
  const scientific = (digits: number) => {
    const [mantissa, exponent] = absolute.toExponential(digits).split('e')
    const compact = mantissa.includes('.') ? mantissa.replace(/0+$/, '').replace(/\.$/, '') : mantissa
    const e = Number(exponent)
    return `${sign}${compact}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`
  }
  const useScientific = absolute >= 1e11 || (absolute < 1e-4 && (plain.includes('e') || plain.length > 11))
  if (!useScientific) {
    if (plain.length <= 11 && fits(sign + plain)) return sign + plain
    const integerDigits = absolute >= 1 ? Math.floor(Math.log10(absolute)) + 1 : 1
    for (let decimals = Math.max(0, 10 - integerDigits); decimals >= 0; decimals--) {
      const fixed = absolute.toFixed(decimals)
      const compact = fixed.includes('.') ? fixed.replace(/0+$/, '').replace(/\.$/, '') : fixed
      // A nonzero value must not disappear through display rounding.
      if (compact.length <= 11 && Number(compact) !== 0 && fits(sign + compact)) return sign + compact
    }
  }
  for (let decimals = 5; decimals >= 0; decimals--) {
    const text = scientific(decimals)
    if (text.length - sign.length <= 11 && fits(text)) return text
  }
  return '#'
}

/** Format a numeric value for common built-in number formats. */
export function formatValue(
  value: string | number | boolean | null,
  numFmtId: number,
  semantics?: ResolvedSemantics,
): string {
  if (value === null) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  const sem = semantics ?? DEFAULT_RENDER_SEMANTICS
  switch (numFmtId) {
    case 0: // General
      return formatGeneral(value)
    case 2: // 0.00
      return formatText(value, '0.00', sem)
    case 9: // 0%
      return formatText(value, '0%', sem)
    case 10: // 0.00%
      return formatText(value, '0.00%', sem)
    case 3: // #,##0
      return formatText(value, '#,##0', sem)
    case 4: // #,##0.00
      return formatText(value, '#,##0.00', sem)
    default:
      // Built-in date formats keep the established ISO product convention but
      // now decode through the SHARED serial model (1900 day-0/fictitious day,
      // explicit 1904) instead of the removed JS-Date epoch path.
      if (numFmtId >= 18 && numFmtId <= 22) {
        const formats = ['h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yy h:mm']
        return formatText(value, formats[numFmtId - 18], sem)
      }
      if (numFmtId >= 14 && numFmtId <= 17) return formatDateSerial(value, sem.dateSystem)
      return formatFiniteFallback(value)
  }
}

/** One shared formatter call; unsupported classes fall back to plain numbers. */
function formatText(value: number, format: string, semantics: ResolvedSemantics): string {
  const result = formatCommon(value, format, semantics)
  if (result.kind === 'text') return result.text
  if (isEvaluationError(result)) return ''
  return formatFiniteFallback(value)
}

export function formatDateSerial(serial: number, dateSystem: '1900' | '1904' = '1900'): string {
  const parts = decodeSerial(serial, dateSystem)
  if (isEvaluationError(parts)) return ''
  const { year, month, day } = parts.civil
  const y = year < 0 ? `-${String(-year).padStart(4, '0')}` : String(year).padStart(4, '0')
  return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
}

/**
 * Cell text for the shared renderer: an authored custom numFmt formatCode uses
 * the SAME finite common formatter as TEXT (numeric AND text/boolean values, so
 * the fourth text section applies). Unsupported authored classes fall back to
 * the built-in/default output and record a specific sheet diagnostic.
 */
function renderCellText(
  value: string | number | boolean | null,
  style: XlsxCellStyle | undefined,
  semantics: ResolvedSemantics | undefined,
  sheet: XlsxSheet,
  fitsGeneral?: (text: string) => boolean,
): string {
  const code = style?.formatCode
  if (code !== undefined && value !== null) {
    const result = formatCommon(value, code, semantics ?? DEFAULT_RENDER_SEMANTICS)
    if (result.kind === 'text') return result.text
    if (!isEvaluationError(result)) noteFormatUnsupported(sheet, code, result.reason)
  }
  if (typeof value === 'number' && (style?.numFmtId ?? 0) === 0) return formatGeneral(value, fitsGeneral)
  return formatValue(value, style?.numFmtId ?? 0, semantics)
}

/**
 * Record an unsupported authored-format diagnostic on the existing sheet
 * channel (same shape as doc.diagnostics), deduped across repeated paints.
 * No product warning UI is produced.
 */
function noteFormatUnsupported(sheet: XlsxSheet, format: string, reason: string): void {
  const holder = sheet as unknown as { diagnostics?: Array<{ kind: string; feature: string; message: string }> }
  holder.diagnostics ??= []
  const message = `Format '${format}' is outside the finite common-format grammar: ${reason}`
  if (!holder.diagnostics.some((d) => d.feature === 'format-unsupported' && d.message === message)) {
    holder.diagnostics.push({ kind: 'unsupported-format', feature: 'format-unsupported', message })
  }
}

function prefixSums(widths: number[]): number[] {
  const out = new Array<number>(widths.length + 1)
  out[0] = 0
  for (let i = 0; i < widths.length; i++) out[i + 1] = out[i] + widths[i]
  return out
}

/** Render a sheet grid. ctx state: 1 unit = 1 px, (0,0) top-left of sheet. */
export interface SheetViewport {
  /** Sheet-coordinate rectangle rendered at the output origin (canvas clips the rest). */
  x: number
  y: number
  width: number
  height: number
}
export function renderSheet(
  sheet: XlsxSheet,
  ctx: CanvasRenderingContext2D,
  metrics?: GridMetrics,
  watermark?: WatermarkOptions | ResolvedWatermark,
  prepared?: RenderSheetAssets,
  viewport?: SheetViewport,
): void {
  const m = metrics ?? computeMetrics(sheet)
  const { colWidthsPx, rowHeightsPx } = m
  const limits = staticAdmission(sheet, m, viewport)
  if (limits.length) throw new StaticExportLimitError(limits)
  const colX = prefixSums(colWidthsPx)
  const rowY = prefixSums(rowHeightsPx)
  const gridWidth = colX[colWidthsPx.length]
  const gridHeight = rowY[rowHeightsPx.length]
  ctx.save()
  try {
  const ranges = sheet.mergeRanges ?? []
  // Viewport culling: paint only intersecting cells/boundaries, translated to
  // the output origin. Absent viewport paints the whole sheet exactly as
  // before (byte-identical: metrics extent, not just the cell grid, because
  // drawings may expand the bitmap beyond the cells). Per-frame text
  // measurement is cached by face+text.
  const vp = viewport ?? { x: 0, y: 0, width: m.widthPx, height: m.heightPx }
  const measured = textMeasurer(ctx)
  if (viewport) ctx.translate(-vp.x, -vp.y)
  const intersects = (x: number, y: number, w: number, h: number): boolean =>
    x < vp.x + vp.width && x + w > vp.x && y < vp.y + vp.height && y + h > vp.y
  const merges = staticMergeIndex(ranges)

  ctx.fillStyle = '#ffffff'
  ctx.fillRect(vp.x, vp.y, vp.width, vp.height)
  // The watermark goes on straight after this fill — drawing it before would be
  // erased here — and under the cells, matching how the Word path does it.
  if (watermark) paintWatermark(ctx, { widthPx: m.widthPx, heightPx: m.heightPx }, watermark)

  const paintBorders: Array<() => void> = []
  // Row bottoms extended by merge ranges anchored in the row: a row whose own
  // box sits above the viewport still paints when its merged range reaches in.
  const rowBottom = new Map<number, number>()
  const edges = mergeEdges(sheet, m, colX, rowY)
  for (const r of ranges) {
    const bottom = edges.row(r.maxRow + 1)
    if (bottom > (rowBottom.get(r.minRow) ?? Number.NEGATIVE_INFINITY)) rowBottom.set(r.minRow, bottom)
  }
  const extent = { width: m.widthPx, height: m.heightPx }
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      // Only coordinates inside the grid have geometry; admission refused any
      // populated cell beyond it that this request needs, so nothing here falls
      // back to the origin. An explicit cell reference names its own row.
      const source = cellSourceRow(row, cell)
      if (!Number.isInteger(source) || source < 0 || source >= rowHeightsPx.length) continue
      if (!Number.isInteger(cell.col) || cell.col < 0 || cell.col >= colWidthsPx.length) continue
      const y = rowY[source]
      const h = rowHeightsPx[source]
      const bottom = Math.max(y + h, rowBottom.get(source) ?? Number.NEGATIVE_INFINITY)
      if (bottom <= vp.y || y >= vp.y + vp.height) continue
      if (merges.covers(cell.row, cell.col)) continue // inside a merge, not its anchor — nothing to paint
      const rng = merges.anchorAt(cell.row, cell.col)
      // geometry: anchor cells expand across their merge range
      const x = colX[cell.col]
      const w = rng ? edges.col(rng.maxCol + 1) - x : colWidthsPx[cell.col]
      const hh = rng ? edges.row(rng.maxRow + 1) - y : h
      // Collapsed (authored-hidden) rows/columns produce no fill or text (spec §5.3).
      if (!(w > 0 && hh > 0) || !intersects(x, y, w, hh)) continue
      const border = paintCellBody(ctx, sheet, cell, x, y, w, hh, measured, extent)
      if (border) paintBorders.push(border)
    }
  }

  // grid lines, with merged interiors masked out. A vertical boundary at
  // column index c (left edge of column c) is hidden for the rows of a merge
  // when c is strictly inside the merge's column span; same for horizontal.
  ctx.strokeStyle = '#d0d0d0'
  ctx.lineWidth = 1
  ctx.beginPath()
  const firstCol = Math.max(0, colX.findIndex((x, i) => x + (colWidthsPx[i] ?? 0) > vp.x))
  const lastCol = colX.findIndex(x => x >= vp.x + vp.width)
  const c0 = firstCol < 0 ? 0 : firstCol, c1 = lastCol < 0 ? colWidthsPx.length : lastCol
  for (let c = c0; c <= c1; c++) {
    const x = colX[c]
    for (const [y1, y2] of visibleSegments(c, ranges, rowY, gridHeight, true)) {
      if (y2 <= vp.y || y1 >= vp.y + vp.height) continue
      ctx.moveTo(x + 0.5, y1 + 0.5)
      ctx.lineTo(x + 0.5, y2 + 0.5)
    }
  }
  const firstRow = Math.max(0, rowY.findIndex((y, i) => y + (rowHeightsPx[i] ?? 0) > vp.y))
  const lastRow = rowY.findIndex(y => y >= vp.y + vp.height)
  const r0 = firstRow < 0 ? 0 : firstRow, r1 = lastRow < 0 ? rowHeightsPx.length : lastRow
  for (let r = r0; r <= r1; r++) {
    const y = rowY[r]
    for (const [x1, x2] of visibleSegments(r, ranges, colX, gridWidth, false)) {
      if (x2 <= vp.x || x1 >= vp.x + vp.width) continue
      ctx.moveTo(x1 + 0.5, y + 0.5)
      ctx.lineTo(x2 + 0.5, y + 0.5)
    }
  }
  ctx.stroke()
  // Explicit formatting wins over both gridlines and neighboring cell fills.
  for (const paint of paintBorders) paint()
  paintSheetDrawings(sheet, sheet.drawings ?? [], ctx, prepared)
  } finally { ctx.restore() }
}

/** Static admission (spec §11.5, plan S2.3), always recomputed from the sheet and
 * the metrics actually used, never trusted from caller-supplied refusal metadata.
 * A full-sheet raster must cover every populated cell and merge and fit the
 * legacy drawing raster clamp. A bounded window is admitted unless it reaches past the grid on
 * an axis where data continues. Nothing uncovered is ever drawn over A1. */
function staticAdmission(sheet: XlsxSheet, m: GridMetrics, viewport?: SheetViewport): StaticExportLimit[] {
  const coverage = gridCoverageLimits(sheet, m.rowHeightsPx.length, m.colWidthsPx.length)
  if (!viewport) return [...coverage, ...requestedRasterLimits(m)]
  const sum = (sizes: number[]) => sizes.reduce((a, b) => a + b, 0)
  const gridWidth = sum(m.colWidthsPx), gridHeight = sum(m.rowHeightsPx)
  // Merges the grid truncates are laid out at their true extent (mergeEdges), so
  // only a window reaching past the grid where data or merges continue is refused.
  const limits = coverage.filter(limit => limit.resource === 'rows' ? viewport.y + viewport.height > gridHeight : viewport.x + viewport.width > gridWidth)
  return limits
}

/** Grid boundary positions for merge geometry. Inside the retained grid these are
 * the prefix sums; past it (a merge the static grid truncates) the extent follows
 * the same authored row/column rules, so an anchor lays out against its true
 * merged box instead of the bitmap edge. Invalid edges keep the bitmap clamp. */
/** @internal Exported for work-count tests; not part of the public API. */
export function mergeEdges(sheet: XlsxSheet, m: GridMetrics, colX: number[] = prefixSums(m.colWidthsPx), rowY: number[] = prefixSums(m.rowHeightsPx)) {
  let work = 0
  const nRows = m.rowHeightsPx.length, nCols = m.colWidthsPx.length
  let rowsBeyond: Map<number, number> | undefined, colsBeyond: Map<number, number> | undefined
  const beyondRows = () => {
    if (!rowsBeyond) {
      rowsBeyond = new Map()
      const hidden = new Map<number, boolean>()
      for (const row of sheet.rows) if (row.index >= nRows) hidden.set(row.index, (hidden.get(row.index) ?? true) && row.hidden === true)
      for (const row of sheet.rows) {
        if (!(row.index >= nRows) || !Number.isSafeInteger(row.index)) continue
        if (hidden.get(row.index)) rowsBeyond.set(row.index, 0)
        else if (row.customHeight && row.heightPt !== undefined && Number.isFinite(row.heightPt) && row.heightPt >= 0) rowsBeyond.set(row.index, Math.round(row.heightPt * (96 / 72)))
      }
    }
    return rowsBeyond
  }
  const beyondCols = () => {
    if (!colsBeyond) {
      colsBeyond = new Map()
      for (const spec of sheet.cols) {
        if (!(spec.widthChars !== undefined || spec.hidden) || !Number.isSafeInteger(spec.min) || !Number.isSafeInteger(spec.max)) continue
        const w = spec.hidden ? 0 : !Number.isFinite(spec.widthChars) || spec.widthChars! < 0 ? DEFAULT_COL_PX : Math.round(spec.widthChars! * CHAR_PX + 5)
        for (let c = Math.max(nCols, spec.min); c <= Math.min(spec.max, MAX_COLS - 1); c++) colsBeyond.set(c, w)
      }
    }
    return colsBeyond
  }
  // Each beyond-grid axis becomes sorted positions plus a prefix of size deltas,
  // built once; an edge is then a binary search (O(B + T log B) for T merges).
  const sorted = (beyond: Map<number, number>, size: number) => {
    const entries = [...beyond].sort((a, b) => a[0] - b[0])
    work += entries.length
    const at = Float64Array.from(entries, entry => entry[0])
    const delta = new Float64Array(entries.length + 1)
    entries.forEach(([, value], i) => { delta[i + 1] = delta[i] + value - size })
    return { at, delta }
  }
  let rowPrefix: ReturnType<typeof sorted> | undefined, colPrefix: ReturnType<typeof sorted> | undefined
  const edge = (index: number, prefix: number[], n: number, size: number, deltas: () => ReturnType<typeof sorted>, clamp: number): number => {
    if (index <= n) return prefix[index] ?? clamp
    if (!Number.isSafeInteger(index)) return clamp
    const { at, delta } = deltas()
    let low = 0, high = at.length
    while (low < high) {
      work++
      const mid = Math.floor((low + high) / 2)
      if (at[mid] < index) low = mid + 1
      else high = mid
    }
    return prefix[n] + (index - n) * size + delta[low]
  }
  return {
    row: (index: number) => edge(index, rowY, nRows, DEFAULT_ROW_PX, () => (rowPrefix ??= sorted(beyondRows(), DEFAULT_ROW_PX)), m.heightPx),
    col: (index: number) => edge(index, colX, nCols, DEFAULT_COL_PX, () => (colPrefix ??= sorted(beyondCols(), DEFAULT_COL_PX)), m.widthPx),
    /** Beyond-grid entries visited so far (test seam for the O(B + T log B) bound). */
    work: () => work,
  }
}

/** @internal Shared with the interactive viewport painter (xlsx/view/renderViewport). */
export type MeasureCached = (font: string, text: string) => number
/** @internal */
export function textMeasurer(ctx: CanvasRenderingContext2D): MeasureCached {
  const measureCache = new Map<string, number>()
  return (font: string, text: string): number => {
    const key = `${font}\n${text}`
    let w = measureCache.get(key)
    if (w === undefined) {
      ctx.font = font
      w = ctx.measureText(text).width
      measureCache.set(key, w)
    }
    return w
  }
}

/** Merge lookups without expanding every covered cell (a giant merge used to
 * allocate one map entry per covered cell). Decisions match the old maps:
 * any covering range that is not anchored here wins; the last range anchored
 * at a cell supplies its merged extent. */
function staticMergeIndex(ranges: readonly XlsxMergeRange[]) {
  const anchor = new Map<string, XlsxMergeRange>()
  const entries: Array<{ rect: { r0: number; r1: number; c0: number; c1: number }; value: XlsxMergeRange }> = []
  for (const r of ranges) {
    anchor.set(`${r.minRow}:${r.minCol}`, r)
    if (isValidCoordinate(r.minRow, r.minCol) && isValidCoordinate(r.maxRow, r.maxCol) && r.minRow <= r.maxRow && r.minCol <= r.maxCol) {
      entries.push({ rect: { r0: r.minRow, r1: r.maxRow, c0: r.minCol, c1: r.maxCol }, value: r })
    }
  }
  const index = RectangleIndex.build(entries)
  return {
    anchorAt: (row: number, col: number) => anchor.get(`${row}:${col}`),
    covers: (row: number, col: number) => isValidCoordinate(row, col)
      && index.firstWhere({ r0: row, r1: row, c0: col, c1: col }, r => r.minRow !== row || r.minCol !== col) !== undefined,
  }
}

/** Fill, text and (deferred) border of one cell box, shared by the static and
 * interactive paths so both lay a cell out identically. Borders stay inside
 * `extent` (the bitmap for static output, the logical sheet for the viewer). */
function paintCellBody(
  ctx: CanvasRenderingContext2D,
  sheet: XlsxSheet,
  cell: XlsxCell,
  x: number, y: number, w: number, hh: number,
  measured: MeasureCached,
  extent: { width: number; height: number },
): (() => void) | undefined {
  return paintCellLayout(ctx, layoutCellBody(ctx, sheet, cell, x, y, w, hh, measured), extent)
}

/** Host formatting and text layout of one cell in the box (x, y, w, hh): the
 * single source of a cell's display text and geometry for both paths. */
/** @internal */
export function layoutCellBody(
  ctx: CanvasRenderingContext2D,
  sheet: XlsxSheet,
  cell: XlsxCell,
  x: number, y: number, w: number, hh: number,
  measured: MeasureCached,
): CellBodyLayout {
  const out: CellBodyLayout = { x, y, w, h: hh, formattedText: '' }
  if (cell.style?.fillColor) out.fill = resolveColor(cell.style.fillColor)
  // text (right-align numbers, left-align strings; alignment spans the merge)
  const font = `${cell.style?.italic ? 'italic ' : ''}${cell.style?.bold ? 'bold ' : ''}${cell.style?.fontSizePt ?? 10}pt "Calibri"`
  const fitsGeneral = cellTextRotation(cell.style?.textRotation) === undefined
    ? (text: string) => measured(font, text) <= Math.max(0, w - PADDING_L - PADDING_R)
    : undefined
  const text = renderCellText(cell.value, cell.style, sheet.semantics, sheet, fitsGeneral)
  out.formattedText = text
  if (text !== '') {
    const color = cell.style?.color ? resolveColor(cell.style.color) : '#000000'
    // OOXML often authors default alignment fields on every cell. Those
    // retain ordinary placement; real alignment/wrapping/rotation needs
    // the measured local layout (including explicit no-wrap for255).
    const explicit = (cell.style?.horizontal !== undefined && cell.style.horizontal !== 'general')
      || (cell.style?.vertical !== undefined && cell.style.vertical !== 'bottom')
      || cell.style?.wrapText === true || cellTextRotation(cell.style?.textRotation) !== undefined
    const isCenter = cell.valueIsError ?? (cell.formula !== undefined && isFormulaError(cell.value))
    if (explicit) {
      ctx.fillStyle = color
      ctx.font = font
      out.text = { ...layoutCellText(ctx, text, cell.style, typeof cell.value === 'number', x, y, w, hh, color, isCenter), font }
    } else {
      // Preserve accepted ordinary default horizontal layout exactly.
      ctx.font = font
      const textW = measured(ctx.font, text)
      const tx = typeof cell.value === 'number'
        ? x + w - PADDING_R - textW
        : isCenter
          ? x + (w - textW) / 2
          : x + PADDING_L
      const ty = y + hh - (hh - (cell.style?.fontSizePt ?? 10) * (96 / 72)) / 2
      out.text = { kind: 'plain', text, font, color, tx, ty, width: textW }
    }
  }
  if (cell.style?.borders) out.borders = cell.style.borders
  return out
}

/** Replay one cell layout at its own coordinates; returns the deferred border. */
/** @internal */
export function paintCellLayout(ctx: CanvasRenderingContext2D, layout: CellBodyLayout, extent: { width: number; height: number }): (() => void) | undefined {
  const { x, y, w, h: hh } = layout
  if (layout.fill) {
    ctx.fillStyle = layout.fill
    ctx.fillRect(x, y, w, hh)
  }
  const text = layout.text
  if (text) {
    ctx.fillStyle = text.color
    ctx.font = text.font
    if (text.kind === 'layout') paintCellText(ctx, text)
    else { ctx.textBaseline = 'alphabetic'; ctx.fillText(text.text, text.tx, text.ty) }
  }
  // borders (anchor draws the merged rect's outline)
  const b = layout.borders
  if (!b) return undefined
  return () => {
    ctx.lineWidth = 1
    const draw = (side: string | undefined, x1: number, yy1: number, x2: number, yy2: number) => {
      if (!side) return
      ctx.strokeStyle = '#000000'
      ctx.beginPath()
      // Keep outer borders within the worksheet bitmap instead of
      // clipping the right/bottom stroke completely off the canvas.
      ctx.moveTo(Math.min(x1 + 0.5, extent.width - 0.5), Math.min(yy1 + 0.5, extent.height - 0.5))
      ctx.lineTo(Math.min(x2 + 0.5, extent.width - 0.5), Math.min(yy2 + 0.5, extent.height - 0.5))
      ctx.stroke()
    }
    draw(b.left, x, y, x, y + hh)
    draw(b.right, x + w, y, x + w, y + hh)
    draw(b.top, x, y, x + w, y)
    draw(b.bottom, x, y + hh, x + w, y + hh)
  }
}

/** Drawings are positioned in worksheet CSS pixels by the parser; both paths
 * paint them in the same (translated) sheet space. */
/** @internal */
export function paintSheetDrawings(sheet: XlsxSheet, drawings: readonly XlsxDrawing[], ctx: CanvasRenderingContext2D, prepared?: RenderSheetAssets): void {
  if (drawings.length) paintScene(drawings, ctx, {
    images: prepared?.images,
    paintContent(node, context, w, h) {
      if (!node.content) return
      paintDrawingContent(node.content, context, w, h, {
        assets: prepared,
        fontFamilyCss: family => JSON.stringify(prepared?.resolveFont?.(family) ?? family),
        paintDiagramText(shape, c, width, height) { if (shape.textBody) paintTextBody(shape.textBody, c, 0, 0, width, height, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
        paintTextbox(content, c, width, height) { paintTextBody({ paragraphs: content.paragraphs, anchor: 't', wrap: true, insetLeftEmu: content.insets.left, insetTopEmu: content.insets.top, insetRightEmu: content.insets.right, insetBottomEmu: content.insets.bottom }, c, 0, 0, width, height, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
      })
    },
    paintText(node, context, w, h) { if (node.textBody) paintTextBody(node.textBody, context, 0, 0, w, h, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
  })
}

/**
 * Visible segments of one grid boundary. For vertical boundaries, `index` is
 * the column index and hidden intervals come from merges whose row span
 * covers this boundary's column-crossing. `vertical=true` → x fixed at
 * boundary `index`, hidden y intervals derived from row spans.
 */
function visibleSegments(
  index: number,
  ranges: XlsxMergeRange[],
  starts: number[],
  total: number,
  vertical: boolean,
): Array<[number, number]> {
  // hidden intervals along the boundary's axis
  const hidden: Array<[number, number]> = []
  for (const r of ranges) {
    const interior = vertical
      ? index > r.minCol && index <= r.maxCol
      : index > r.minRow && index <= r.maxRow
    if (!interior) continue
    // A merge starting beyond the grid has no boundary inside it; it must not
    // mask from the origin. Its far edge still clamps to the grid extent.
    const first = starts[vertical ? r.minRow : r.minCol]
    if (first === undefined) continue
    hidden.push([first, starts[(vertical ? r.maxRow : r.maxCol) + 1] ?? total])
  }
  if (hidden.length === 0) return [[0, total]]
  hidden.sort((a, b) => a[0] - b[0])
  const segs: Array<[number, number]> = []
  let pos = 0
  for (const [h1, h2] of hidden) {
    if (h1 > pos) segs.push([pos, h1])
    pos = Math.max(pos, h2)
  }
  if (pos < total) segs.push([pos, total])
  return segs
}

const PADDING_L = 3
const PADDING_R = 3
