/**
 * XLSX rendering: paint sheets as a grid onto a canvas 2D context.
 * Layout mimics Excel defaults: default col width 8.43 chars (~64px), row
 * height 15pt (20px). Cell geometry is fixed; no reflow needed.
 */
import type { XlsxCellStyle, XlsxDrawing, XlsxMergeRange, XlsxSheet } from './types'
import type { PptxTextBody } from '../pptx/types'
import { layoutTextBody, type MeasureText, type TextLayout } from '../drawing/text-layout'
import { resolveColor } from '../core/color'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import { paintScene } from '../drawing/scene-paint'
import { paintDrawingContent } from '../drawing/content-paint'
import { createTextBodyMeasurer, paintTextBody } from '../drawing/text-paint'
import { graphemes } from '../core/text-recording'
import { resolveGeometry } from '../drawing/geometry'
import type { DrawingContentShape } from '../drawing/content'
import { isFormulaError, isEvaluationError } from './formula/evaluator'
import { formatCommon } from './formula/format'
import { decodeSerial } from './formula/serial'
import type { ResolvedSemantics } from './formula/types'

export interface GridMetrics {
  colWidthsPx: number[]
  rowHeightsPx: number[]
  widthPx: number
  heightPx: number
  requestedDrawingBounds?: { right: number; bottom: number }
  retainedDrawingBounds?: { right: number; bottom: number }
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
 * and clips to the cell; source remains complete, with clip-aware indexing. */
function paintCellText(ctx: CanvasRenderingContext2D, text: string, style: XlsxCellStyle | undefined, numeric: boolean, x: number, y: number, w: number, h: number, color: string, center = false): void {
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
    // Coordinates in the placed layout are worksheet-local and already include
    // the final merged bounds. Paint and recording replay this exact geometry.
    paintTextBody(body, ctx, 0, 0, w, h, family => family, undefined, { layout: placed, clip: { x, y, width: w, height: h } })
  } finally { ctx.restore() }
}
/** How much wider than the used range a <col> declaration may be and still
 * count as real formatting rather than noise. */
const MAX_DECLARED_COL_SLACK = 64
const MAX_DRAWING_SIDE = 16384
const MAX_DRAWING_AREA = 16777216

type Matrix = [number, number, number, number, number, number]
const identity: Matrix = [1, 0, 0, 1, 0, 0]
function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]]
}
function transformedDrawingBounds(nodes: XlsxDrawing[]): { right: number; bottom: number } {
  const bounds = { right: 0, bottom: 0 }
  type BoundsNode = XlsxDrawing | DrawingContentShape
  const walk = (items: BoundsNode[], parent: Matrix): void => {
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
        if (Number.isFinite(tx) && Number.isFinite(ty)) { bounds.right = Math.max(bounds.right, tx); bounds.bottom = Math.max(bounds.bottom, ty) }
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
  walk(nodes, identity)
  return bounds
}

export function computeMetrics(sheet: XlsxSheet): GridMetrics {
  let maxCol = 0
  for (const row of sheet.rows) for (const c of row.cells) maxCol = Math.max(maxCol, c.col)
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
    if (drawing ? spec.widthChars !== undefined || spec.hidden : spec.widthChars !== undefined && !spec.hidden) {
      const w = drawing && spec.hidden ? 0 : drawing && (spec.widthChars === undefined || !Number.isFinite(spec.widthChars) || spec.widthChars < 0) ? DEFAULT_COL_PX : Math.round(spec.widthChars! * CHAR_PX + 5)
      for (let c = Math.max(0, spec.min); c <= spec.max && c < nCols; c++) colWidthsPx[c] = w
    }
  }
  let maxRow = 0
  for (const row of sheet.rows) if (!drawing || Number.isSafeInteger(row.index) && row.index >= 0) maxRow = Math.max(row.index, maxRow)
  const nRows = Math.min(Math.max(maxRow + 1, drawing ? (sheet.drawingMarkers?.maxRow ?? 0) + 1 : 1), MAX_GRID_ROWS)
  const rowHeightsPx = new Array<number>(nRows).fill(DEFAULT_ROW_PX)
  for (const row of sheet.rows) {
    if (row.index >= 0 && row.index < nRows) {
      if (drawing && row.hidden) rowHeightsPx[row.index] = 0
      else if (row.heightPt !== undefined && row.customHeight) rowHeightsPx[row.index] = drawing && (!Number.isFinite(row.heightPt) || row.heightPt < 0) ? DEFAULT_ROW_PX : Math.round(row.heightPt * (96 / 72))
    }
  }
  const gridWidth = colWidthsPx.reduce((a, b) => a + b, 0), gridHeight = rowHeightsPx.reduce((a, b) => a + b, 0)
  if (!drawing) return { colWidthsPx, rowHeightsPx, widthPx: gridWidth, heightPx: gridHeight }
  const visual = transformedDrawingBounds(sheet.drawings!)
  const requested = { right: Math.max(gridWidth, visual.right), bottom: Math.max(gridHeight, visual.bottom) }
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
  return { paperPx, printable, scale, pagesWide, pagesTall }
}
/**
 * Paint one print page: paper-white sheet, content scaled into the printable
 * area, clipped. Page (pageCol, pageRow) selects its sheet-coordinate window;
 * content beyond one page clips (multi-page paintables are a follow-up).
 */
export function renderPrintPage(
  sheet: XlsxSheet,
  ctx: CanvasRenderingContext2D,
  metrics: GridMetrics,
  print: PrintMetrics,
  pageCol = 0,
  pageRow = 0,
  watermark?: WatermarkOptions | ResolvedWatermark,
  prepared?: { images?: ReadonlyArray<CanvasImageSource | undefined>; resolveFont?: (family: string) => string },
): void {
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
    renderSheet(sheet, ctx, metrics, undefined, prepared, {
      x: (pageCol * printable.width) / scale,
      y: (pageRow * printable.height) / scale,
      width: printable.width / scale,
      height: printable.height / scale,
    })
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
  prepared?: { images?: ReadonlyArray<CanvasImageSource | undefined>; resolveFont?: (family: string) => string },
  viewport?: SheetViewport,
): void {
  ctx.save()
  try {
  const m = metrics ?? computeMetrics(sheet)
  const { colWidthsPx, rowHeightsPx } = m
  const colX = prefixSums(colWidthsPx)
  const rowY = prefixSums(rowHeightsPx)
  const gridWidth = colX[colWidthsPx.length]
  const gridHeight = rowY[rowHeightsPx.length]
  const ranges = sheet.mergeRanges ?? []
  // Viewport culling: paint only intersecting cells/boundaries, translated to
  // the output origin. Absent viewport paints the whole sheet exactly as
  // before (byte-identical: metrics extent, not just the cell grid, because
  // drawings may expand the bitmap beyond the cells). Per-frame text
  // measurement is cached by face+text.
  const vp = viewport ?? { x: 0, y: 0, width: m.widthPx, height: m.heightPx }
  const measureCache = new Map<string, number>()
  const measured = (font: string, text: string): number => {
    const key = `${font}\n${text}`
    let w = measureCache.get(key)
    if (w === undefined) {
      ctx.font = font
      w = ctx.measureText(text).width
      measureCache.set(key, w)
    }
    return w
  }
  if (viewport) ctx.translate(-vp.x, -vp.y)
  const intersects = (x: number, y: number, w: number, h: number): boolean =>
    x < vp.x + vp.width && x + w > vp.x && y < vp.y + vp.height && y + h > vp.y

  // map "row:col" -> range for every covered (non-anchor) cell
  const covered = new Map<string, XlsxMergeRange>()
  const anchor = new Map<string, XlsxMergeRange>()
  for (const r of ranges) {
    for (let row = r.minRow; row <= r.maxRow; row++) {
      for (let col = r.minCol; col <= r.maxCol; col++) {
        const key = `${row}:${col}`
        if (row === r.minRow && col === r.minCol) anchor.set(key, r)
        else covered.set(key, r)
      }
    }
  }

  ctx.fillStyle = '#ffffff'
  ctx.fillRect(vp.x, vp.y, vp.width, vp.height)
  // The watermark goes on straight after this fill — drawing it before would be
  // erased here — and under the cells, matching how the Word path does it.
  if (watermark) paintWatermark(ctx, { widthPx: m.widthPx, heightPx: m.heightPx }, watermark)

  const paintBorders: Array<() => void> = []
  // Row bottoms extended by merge ranges anchored in the row: a row whose own
  // box sits above the viewport still paints when its merged range reaches in.
  const rowBottom = new Map<number, number>()
  for (const r of ranges) {
    const bottom = rowY[r.maxRow + 1] ?? m.heightPx
    if (bottom > (rowBottom.get(r.minRow) ?? Number.NEGATIVE_INFINITY)) rowBottom.set(r.minRow, bottom)
  }
  for (const row of sheet.rows) {
    const y = rowY[row.index] ?? 0
    const h = rowHeightsPx[row.index] ?? DEFAULT_ROW_PX
    const bottom = Math.max(y + h, rowBottom.get(row.index) ?? Number.NEGATIVE_INFINITY)
    if (bottom <= vp.y || y >= vp.y + vp.height) continue
    for (const cell of row.cells) {
      const cov = covered.get(`${cell.row}:${cell.col}`)
      if (cov) continue // inside a merge, not its anchor — nothing to paint
      const rng = anchor.get(`${cell.row}:${cell.col}`)
      // geometry: anchor cells expand across their merge range
      const x = colX[cell.col] ?? 0
      const w = rng ? (colX[rng.maxCol + 1] ?? m.widthPx) - x : (colWidthsPx[cell.col] ?? DEFAULT_COL_PX)
      const hh = rng ? (rowY[rng.maxRow + 1] ?? m.heightPx) - y : h
      if (!intersects(x, y, w, hh)) continue
      // fill
      const fill = cell.style?.fillColor
      if (fill) {
        ctx.fillStyle = resolveColor(fill)
        ctx.fillRect(x, y, w, hh)
      }
      // text (right-align numbers, left-align strings; alignment spans the merge)
      const font = `${cell.style?.italic ? 'italic ' : ''}${cell.style?.bold ? 'bold ' : ''}${cell.style?.fontSizePt ?? 10}pt "Calibri"`
      const fitsGeneral = cellTextRotation(cell.style?.textRotation) === undefined
        ? (text: string) => measured(font, text) <= Math.max(0, w - PADDING_L - PADDING_R)
        : undefined
      const text = renderCellText(cell.value, cell.style, sheet.semantics, sheet, fitsGeneral)
      if (text !== '') {
        const color = cell.style?.color ? resolveColor(cell.style.color) : '#000000'
        ctx.fillStyle = color
        ctx.font = font
        // OOXML often authors default alignment fields on every cell. Those
        // retain ordinary placement; real alignment/wrapping/rotation needs
        // the measured local layout (including explicit no-wrap for255).
        const explicit = (cell.style?.horizontal !== undefined && cell.style.horizontal !== 'general')
          || (cell.style?.vertical !== undefined && cell.style.vertical !== 'bottom')
          || cell.style?.wrapText === true || cellTextRotation(cell.style?.textRotation) !== undefined
        const isCenter = cell.valueIsError ?? (cell.formula !== undefined && isFormulaError(cell.value))
        if (explicit) paintCellText(ctx, text, cell.style, typeof cell.value === 'number', x, y, w, hh, color, isCenter)
        else {
          // Preserve accepted ordinary default horizontal layout exactly.
          const textW = measured(ctx.font, text)
          const tx = typeof cell.value === 'number'
            ? x + w - PADDING_R - textW
            : isCenter
              ? x + (w - textW) / 2
              : x + PADDING_L
          const ty = y + hh - (hh - (cell.style?.fontSizePt ?? 10) * (96 / 72)) / 2
          ctx.textBaseline = 'alphabetic'; ctx.fillText(text, tx, ty)
        }
      }
      // borders (anchor draws the merged rect's outline)
      const b = cell.style?.borders
      if (b) paintBorders.push(() => {
        ctx.lineWidth = 1
        const draw = (side: string | undefined, x1: number, yy1: number, x2: number, yy2: number) => {
          if (!side) return
          ctx.strokeStyle = '#000000'
          ctx.beginPath()
          // Keep outer borders within the worksheet bitmap instead of
          // clipping the right/bottom stroke completely off the canvas.
          ctx.moveTo(Math.min(x1 + 0.5, m.widthPx - 0.5), Math.min(yy1 + 0.5, m.heightPx - 0.5))
          ctx.lineTo(Math.min(x2 + 0.5, m.widthPx - 0.5), Math.min(yy2 + 0.5, m.heightPx - 0.5))
          ctx.stroke()
        }
        draw(b.left, x, y, x, y + hh)
        draw(b.right, x + w, y, x + w, y + hh)
        draw(b.top, x, y, x + w, y)
        draw(b.bottom, x, y + hh, x + w, y + hh)
      })
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
  if (sheet.drawings?.length) paintScene(sheet.drawings, ctx, {
    images: prepared?.images,
    paintContent(node, context, w, h) {
      if (!node.content) return
      paintDrawingContent(node.content, context, w, h, {
        fontFamilyCss: family => JSON.stringify(prepared?.resolveFont?.(family) ?? family),
        paintDiagramText(shape, c, width, height) { if (shape.textBody) paintTextBody(shape.textBody, c, 0, 0, width, height, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
        paintTextbox(content, c, width, height) { paintTextBody({ paragraphs: content.paragraphs, anchor: 't', wrap: true, insetLeftEmu: content.insets.left, insetTopEmu: content.insets.top, insetRightEmu: content.insets.right, insetBottomEmu: content.insets.bottom }, c, 0, 0, width, height, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
      })
    },
    paintText(node, context, w, h) { if (node.textBody) paintTextBody(node.textBody, context, 0, 0, w, h, prepared?.resolveFont ?? (f => f), sheet.drawingTheme) },
  })
  } finally { ctx.restore() }
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
    if (vertical) {
      const y1 = starts[r.minRow] ?? 0
      const y2 = starts[r.maxRow + 1] ?? total
      hidden.push([y1, y2])
    } else {
      const x1 = starts[r.minCol] ?? 0
      const x2 = starts[r.maxCol + 1] ?? total
      hidden.push([x1, x2])
    }
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
