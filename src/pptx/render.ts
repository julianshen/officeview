import { paintDrawingContent } from '../drawing/content-paint'
import type { FontResolver } from '../core/fonts/register'
/**
 * PPTX rendering: paint slides onto a canvas 2D context. 1 unit = 1 px;
 * EMU geometry converted with emuToPx at 96dpi default.
 */
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import type { PptxDocument, PptxShape, PptxSlide, PptxTable, PptxTableCell } from './types'
import { emuToPx } from '../core/geometry'
import { resolveColor } from '../core/color'
import { paintTextBody } from './text-paint'
import { applyTableTextDefaults } from './text-layout'
export { layoutTextBody, resolveTextFamily } from './text-layout'
import { paintScene } from '../drawing/scene-paint'
import type { DrawingFill, DrawingLine, ThemeContext } from '../drawing/style'
import { paintGeometry } from '../drawing/paint'

export interface SlideMetrics {
  widthPx: number
  heightPx: number
}

export function slideMetrics(doc: PptxDocument): SlideMetrics {
  return {
    widthPx: Math.round(emuToPx(doc.slideWidthEmu)),
    heightPx: Math.round(emuToPx(doc.slideHeightEmu)),
  }
}

/** Render one slide. ctx: 1 unit = 1 px, slide fills the given size. */
export function renderSlide(
  slide: PptxSlide,
  ctx: CanvasRenderingContext2D,
  metrics?: SlideMetrics,
  images?: Array<CanvasImageSource | undefined>,
  watermark?: WatermarkOptions | ResolvedWatermark,
  resolveFont: FontResolver = family => family,
): void {
  ctx.save()
  try {
    const m = metrics ?? { widthPx: Math.round(emuToPx(slide.widthEmu)), heightPx: Math.round(emuToPx(slide.heightEmu)) }
    ctx.fillStyle = slide.background ?? '#ffffff'
    ctx.fillRect(0, 0, m.widthPx, m.heightPx)
    // after the fill so it is not erased, before the shapes so it sits under them
    if (watermark) paintWatermark(ctx, { widthPx: m.widthPx, heightPx: m.heightPx, background: slide.background }, watermark)
    const missingImage = (shape: PptxShape): void => {
      const message = 'Embedded image could not be decoded; static representation is unavailable'
      if (shape.diagnostics?.some(issue => issue.kind === 'missing-image' && issue.message === message)) return
      const issue = { kind: 'missing-image' as const, message, feature: 'blipFill', source: shape.source }
      shape.diagnostics ??= []
      shape.diagnostics.push(issue)
      slide.diagnostics ??= []
      slide.diagnostics.push(issue)
    }
    paintScene(slide.shapes, ctx, {
      images, missingImage,
      paintContent(shape, context, w, h) {
        if (shape.content) paintDrawingContent(shape.content, context, w, h, {
          fontFamilyCss: family => JSON.stringify(resolveFont(family)),
          paintDiagramText(node, c, width, height) {
            // Cacheless-SmartArt text boxes carry no intrinsic extents; paint
            // their text across the enclosing frame box instead of a degenerate
            // one (which paintTextBody early-returns on, losing all visible ink).
            if (node.textBody) paintTextBody(node.textBody, c, 0, 0, width > 0 ? width : w, height > 0 ? height : h, resolveFont, slide.theme)
          },
          paintTextbox(drawing, c, width, height) {
            paintTextBody({ paragraphs: drawing.paragraphs, anchor: 't', wrap: true,
              insetLeftEmu: drawing.insets.left, insetRightEmu: drawing.insets.right,
              insetTopEmu: drawing.insets.top, insetBottomEmu: drawing.insets.bottom,
            }, c, 0, 0, width, height, resolveFont, slide.theme)
          },
        })
        if (shape.table) paintTable(shape.table, context, 0, 0, w, h, resolveFont, slide.theme)
      },
      paintText(shape, context, w, h) {
        if (shape.textBody) paintTextBody(shape.textBody, context, 0, 0, w, h, resolveFont, slide.theme)
      },
    })
  } finally { ctx.restore() }
}

const CELL_PAD_EMU = 91440 // 0.1" default PowerPoint cell inset
const CELL_VPAD_EMU = 45720

/**
 * Paint an a:tbl inside the frame's rect. Column widths come from the grid;
 * row heights come from a:tr@h when present, otherwise the remaining height
 * is divided evenly (PowerPoint auto-grows rows).
 */
function paintTable(
  table: PptxTable,
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  resolveFont: FontResolver,
  theme?: ThemeContext,
): void {
  const cols = table.colWidthsEmu.length > 0 ? table.colWidthsEmu : [w / 96]
  const colWidthsPx = cols.map((c) => emuToPx(c))
  const totalW = colWidthsPx.reduce((a, b) => a + b, 0) || w
  // scale the grid to the frame if the declared widths disagree
  const colScale = totalW > 0 ? w / totalW : 1
  const widths = colWidthsPx.map((c) => c * colScale)

  // row heights: explicit h, else split what is left evenly
  const explicitHeights = table.rows.map((r) => (r.heightEmu !== undefined ? emuToPx(r.heightEmu) : 0))
  const explicitTotal = explicitHeights.reduce((a, b) => a + b, 0)
  const autoCount = explicitHeights.filter((v) => v === 0).length
  const fallbackRowH = autoCount > 0 ? Math.max(8, (h - explicitTotal) / autoCount) : 0
  const heights = explicitHeights.map((v) => (v === 0 ? fallbackRowH : v))

  const colX: number[] = [0]
  for (const cw of widths) colX.push(colX[colX.length - 1] + cw)
  const rowY: number[] = [0]
  for (const rh of heights) rowY.push(rowY[rowY.length - 1] + rh)

  ctx.save()
  ctx.translate(x, y)
  ctx.beginPath()
  ctx.rect(0, 0, w, h)
  ctx.clip()

  // fills + text per cell; merged cells are painted by their anchor
  const owners = new Map<string, string>()
  const ownerCells = new Map<string, PptxTableCell>()
  const styleFills = table.styleFills
  // resolve the style-provided fill for a cell, lowest priority last
  const styleFillFor = (ri: number, ci: number): string | undefined => {
    if (!styleFills) return undefined
    if (table.firstRow && styleFills.firstRow && ri === 0) return styleFills.firstRow
    if (table.bandRow) {
      // banding counts data rows only (row 0 is the header when firstRow)
      const band = table.firstRow ? ri - 1 : ri
      if (band >= 0) return (band % 2 === 0 ? styleFills.band1 : styleFills.band2) ?? styleFills.wholeTable
    }
    if (styleFills.wholeTable) return styleFills.wholeTable
    void ci
    return undefined
  }
  table.rows.forEach((row, ri) => {
    let ci = 0
    // A physical row has one XML cell per grid column, including placeholders.
    // Its gridSpan controls anchor coverage, never the slot's column advance.
    const physicalColumns = row.cells.length === widths.length
    for (const cell of row.cells) {
      if (cell.merged) {
        // skip: an absorbed cell has no content of its own
        ci += physicalColumns ? 1 : Math.max(1, cell.gridSpan)
        continue
      }
      const colSpan = Math.max(1, cell.gridSpan)
      const rowSpan = Math.max(1, cell.rowSpan)
      const cx = colX[ci] ?? 0
      const cy = rowY[ri] ?? 0
      const cw = (colX[ci + colSpan] ?? w) - cx
      const ch = (rowY[ri + rowSpan] ?? h) - cy
      // explicit cell fill wins over the table style's banding
      const fill = cell.fill ?? styleFillFor(ri, ci)
      if (cell.drawingFill?.kind === 'solid') {
        const { r, g, b, a } = cell.drawingFill.color
        ctx.fillStyle = `rgba(${r},${g},${b},${a})`
        ctx.fillRect(cx, cy, cw, ch)
      } else if (cell.drawingFill) {
        ctx.save()
        ctx.translate(cx, cy)
        paintGeometry(ctx, [{ fill: 'norm', stroke: false, commands: [['moveTo', 0, 0], ['lnTo', cw, 0], ['lnTo', cw, ch], ['lnTo', 0, ch], ['close']] }], { fill: cell.drawingFill, issues: [] }, cw, ch)
        ctx.restore()
      } else if (fill) {
        ctx.fillStyle = resolveColor(fill)
        ctx.fillRect(cx, cy, cw, ch)
      }
      if (cell.paragraphs.length > 0) {
        // header rows in a styled table often switch to light text
        const textColor = table.firstRow && ri === 0 ? table.firstRowTextColor : undefined
        const headerBold = table.firstRow && ri === 0 ? table.firstRowBold : undefined
        const paragraphs = textColor || headerBold !== undefined
          ? cell.paragraphs.map(p => applyTableTextDefaults(p, {
              ...(textColor !== undefined ? { color: textColor } : {}),
              ...(headerBold !== undefined ? { bold: headerBold } : {}),
            }))
          : cell.paragraphs
        paintTextBody(
          {
            paragraphs,
            anchor: cell.anchor ?? 'ctr',
            insetLeftEmu: cell.margins?.leftEmu ?? CELL_PAD_EMU,
            insetRightEmu: cell.margins?.rightEmu ?? CELL_PAD_EMU,
            insetTopEmu: cell.margins?.topEmu ?? CELL_VPAD_EMU,
            insetBottomEmu: cell.margins?.bottomEmu ?? CELL_VPAD_EMU,
            wrap: true,
          },
          ctx,
          cx,
          cy,
          cw,
          ch,
          resolveFont, theme,
        )
      }
      ownerCells.set(`${ri}:${ci}`, cell)
      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < colSpan; dc++) owners.set(`${ri + dr}:${ci + dc}`, `${ri}:${ci}`)
      }
      ci += physicalColumns ? 1 : colSpan
    }
  })

  // grid: thin lines on every boundary, skipping interiors of merged cells
  const sameOwner = (a: string, b: string) => owners.has(a) && owners.get(a) === owners.get(b)
  const directBorder = (row: number, col: number, side: 'left' | 'right' | 'top' | 'bottom'): DrawingLine | undefined => {
    const owner = owners.get(`${row}:${col}`)
    return owner === undefined ? undefined : ownerCells.get(owner)?.drawingBorders?.[side]
  }
  const draw = (side: 'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV',
    row: number, col: number, x1: number, y1: number, x2: number, y2: number) => {
    const headerSide = side === 'insideH' && row === 1 ? 'bottom' : side
    const headerBoundary = row === 0 || (side === 'insideH' && row === 1)
    const border = table.firstRow && headerBoundary ? table.firstRowBorders?.[headerSide] ?? table.styleBorders?.[side]
      : table.styleBorders?.[side]
    // At a shared boundary the later cell's left/top declaration wins.
    const direct = side === 'insideV' ? directBorder(row, col, 'left') ?? directBorder(row, col - 1, 'right')
      : side === 'insideH' ? directBorder(row, col, 'top') ?? directBorder(row - 1, col, 'bottom')
      : directBorder(side === 'bottom' ? row - 1 : row, side === 'right' ? col - 1 : col, side)
    if (direct) {
      const hex = /^#([0-9a-f]{6})$/i.exec(border?.color ?? '')?.[1]
      const fill: DrawingFill = { kind: 'solid', color: hex ? { r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4), 16), a: 1 } : { r: 0, g: 0, b: 0, a: .35 } }
      const line: DrawingLine = { ...direct, width: direct.width ?? (border ? emuToPx(border.widthEmu ?? 12700) : 1), fill: direct.fill ?? fill }
      paintGeometry(ctx, [{ fill: 'none', stroke: true, commands: [['moveTo', x1, y1], ['lnTo', x2, y2]] }], { line, issues: [] }, w, h)
      return
    }
    ctx.strokeStyle = border?.color ?? 'rgba(0,0,0,0.35)'
    ctx.lineWidth = border ? emuToPx(border.widthEmu ?? 12700) : 1
    ctx.beginPath()
    ctx.moveTo(x1, y1)
    ctx.lineTo(x2, y2)
    ctx.stroke()
  }
  const gridOffset = table.styleBorders ? 0 : 0.5
  for (let c = 0; c <= widths.length; c++) {
    const gx = colX[c] ?? w
    for (let r = 0; r < heights.length; r++) {
      if (c > 0 && c < widths.length && sameOwner(`${r}:${c - 1}`, `${r}:${c}`)) continue
      const y1 = rowY[r] ?? 0
      const y2 = rowY[r + 1] ?? h
      draw(c === 0 ? 'left' : c === widths.length ? 'right' : 'insideV', r, c, gx + gridOffset, y1, gx + gridOffset, y2)
    }
  }
  for (let r = 0; r <= heights.length; r++) {
    const gy = rowY[r] ?? h
    for (let c = 0; c < widths.length; c++) {
      if (r > 0 && r < heights.length && sameOwner(`${r - 1}:${c}`, `${r}:${c}`)) continue
      const x1 = colX[c] ?? 0
      const x2 = colX[c + 1] ?? w
      draw(r === 0 ? 'top' : r === heights.length ? 'bottom' : 'insideH', r, c, x1, gy + gridOffset, x2, gy + gridOffset)
    }
  }
  ctx.restore()
}
