/**
 * XLSX rendering: paint sheets as a grid onto a canvas 2D context.
 * Layout mimics Excel defaults: default col width 8.43 chars (~64px), row
 * height 15pt (20px). Cell geometry is fixed; no reflow needed.
 */
import type { XlsxMergeRange, XlsxSheet } from './types'
import { resolveColor } from '../core/color'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'

export interface GridMetrics {
  colWidthsPx: number[]
  rowHeightsPx: number[]
  widthPx: number
  heightPx: number
}

const DEFAULT_COL_PX = 64
const DEFAULT_ROW_PX = 20
const CHAR_PX = 7
/** Hard caps so a pathological sheet can never allocate an impossible canvas. */
const MAX_GRID_COLS = 4096
const MAX_GRID_ROWS = 16384
/** How much wider than the used range a <col> declaration may be and still
 * count as real formatting rather than noise. */
const MAX_DECLARED_COL_SLACK = 64

export function computeMetrics(sheet: XlsxSheet): GridMetrics {
  let maxCol = 0
  for (const row of sheet.rows) for (const c of row.cells) maxCol = Math.max(maxCol, c.col)
  // Real spreadsheets declare column formatting far wider than the used range
  // (a sheet with data in A1:C10 may carry <col max="1025"/>). Sizing the grid
  // from such a declared range alone produced a 66,000px canvas and threw, so
  // trust the used range when the declared one is implausibly wider. A modestly
  // wider declaration is real formatting (a styled but empty column) and stays.
  const declaredMax = sheet.cols.reduce((m, c) => Math.max(m, c.max), maxCol)
  const gridMax = declaredMax > maxCol + MAX_DECLARED_COL_SLACK ? maxCol : declaredMax
  const nCols = Math.min(Math.max(gridMax + 1, 1), MAX_GRID_COLS)
  const colWidthsPx = new Array<number>(nCols).fill(DEFAULT_COL_PX)
  for (const spec of sheet.cols) {
    if (spec.widthChars !== undefined && !spec.hidden) {
      const w = Math.round(spec.widthChars * CHAR_PX + 5)
      for (let c = spec.min; c <= spec.max && c < nCols; c++) colWidthsPx[c] = w
    }
  }
  let maxRow = 0
  for (const row of sheet.rows) maxRow = Math.max(row.index, maxRow)
  const nRows = Math.min(maxRow + 1, MAX_GRID_ROWS)
  const rowHeightsPx = new Array<number>(nRows).fill(DEFAULT_ROW_PX)
  for (const row of sheet.rows) {
    if (row.heightPt !== undefined && row.customHeight && row.index < nRows) {
      rowHeightsPx[row.index] = Math.round(row.heightPt * (96 / 72))
    }
  }
  return {
    colWidthsPx,
    rowHeightsPx,
    widthPx: colWidthsPx.reduce((a, b) => a + b, 0),
    heightPx: rowHeightsPx.reduce((a, b) => a + b, 0),
  }
}

/** Format a numeric value for common built-in number formats. */
export function formatValue(value: string | number | boolean | null, numFmtId: number): string {
  if (value === null) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  switch (numFmtId) {
    case 2: // 0.00
      return value.toFixed(2)
    case 9: return `${Math.round(value * 100)}%`
    case 10: return `${(value * 100).toFixed(2)}%`
    case 3: return Math.round(value).toLocaleString('en-US')
    case 4: return value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    default:
      if (numFmtId >= 14 && numFmtId <= 22) return formatDateSerial(value)
      return String(Math.round(value * 100) / 100)
  }
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30)

export function formatDateSerial(serial: number): string {
  const d = new Date(EXCEL_EPOCH + serial * 86400000)
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
  const dd = String(d.getUTCDate()).padStart(2, '0')
  return `${d.getUTCFullYear()}-${mm}-${dd}`
}

function prefixSums(widths: number[]): number[] {
  const out = new Array<number>(widths.length + 1)
  out[0] = 0
  for (let i = 0; i < widths.length; i++) out[i + 1] = out[i] + widths[i]
  return out
}

/** Render a sheet grid. ctx state: 1 unit = 1 px, (0,0) top-left of sheet. */
export function renderSheet(
  sheet: XlsxSheet,
  ctx: CanvasRenderingContext2D,
  metrics?: GridMetrics,
  watermark?: WatermarkOptions | ResolvedWatermark,
): void {
  const m = metrics ?? computeMetrics(sheet)
  const { colWidthsPx, rowHeightsPx } = m
  const colX = prefixSums(colWidthsPx)
  const rowY = prefixSums(rowHeightsPx)
  const ranges = sheet.mergeRanges ?? []

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
  ctx.fillRect(0, 0, m.widthPx, m.heightPx)
  // The watermark goes on straight after this fill — drawing it before would be
  // erased here — and under the cells, matching how the Word path does it.
  if (watermark) paintWatermark(ctx, { widthPx: m.widthPx, heightPx: m.heightPx }, watermark)

  const paintBorders: Array<() => void> = []
  for (const row of sheet.rows) {
    const y = rowY[row.index] ?? 0
    const h = rowHeightsPx[row.index] ?? DEFAULT_ROW_PX
    for (const cell of row.cells) {
      const cov = covered.get(`${cell.row}:${cell.col}`)
      if (cov) continue // inside a merge, not its anchor — nothing to paint
      const rng = anchor.get(`${cell.row}:${cell.col}`)
      // geometry: anchor cells expand across their merge range
      const x = colX[cell.col] ?? 0
      const w = rng ? (colX[rng.maxCol + 1] ?? m.widthPx) - x : (colWidthsPx[cell.col] ?? DEFAULT_COL_PX)
      const hh = rng ? (rowY[rng.maxRow + 1] ?? m.heightPx) - y : h
      // fill
      const fill = cell.style?.fillColor
      if (fill) {
        ctx.fillStyle = resolveColor(fill)
        ctx.fillRect(x, y, w, hh)
      }
      // text (right-align numbers, left-align strings; alignment spans the merge)
      const text = formatValue(cell.value, cell.style?.numFmtId ?? 0)
      if (text !== '') {
        ctx.fillStyle = cell.style?.color ? resolveColor(cell.style.color) : '#000000'
        ctx.font = `${cell.style?.italic ? 'italic ' : ''}${cell.style?.bold ? 'bold ' : ''}${cell.style?.fontSizePt ?? 10}pt "Calibri"`
        const numeric = typeof cell.value === 'number'
        const textW = ctx.measureText(text).width
        const tx = numeric ? x + w - PADDING_R - textW : x + PADDING_L
        const ty = y + hh - (hh - (cell.style?.fontSizePt ?? 10) * (96 / 72)) / 2
        ctx.textBaseline = 'alphabetic'
        ctx.fillText(text, tx, ty)
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
  for (let c = 0; c <= colWidthsPx.length; c++) {
    const x = c === colWidthsPx.length ? m.widthPx : colX[c]
    for (const [y1, y2] of visibleSegments(c, ranges, rowY, m.heightPx, true)) {
      ctx.moveTo(x + 0.5, y1 + 0.5)
      ctx.lineTo(x + 0.5, y2 + 0.5)
    }
  }
  for (let r = 0; r <= rowHeightsPx.length; r++) {
    const y = r === rowHeightsPx.length ? m.heightPx : rowY[r]
    for (const [x1, x2] of visibleSegments(r, ranges, colX, m.widthPx, false)) {
      ctx.moveTo(x1 + 0.5, y + 0.5)
      ctx.lineTo(x2 + 0.5, y + 0.5)
    }
  }
  ctx.stroke()
  // Explicit formatting wins over both gridlines and neighboring cell fills.
  for (const paint of paintBorders) paint()
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
