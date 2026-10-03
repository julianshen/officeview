/**
 * PPTX rendering: paint slides onto a canvas 2D context. 1 unit = 1 px;
 * EMU geometry converted with emuToPx at 96dpi default.
 */
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../core/watermark'
import type { PptxDocument, PptxShape, PptxSlide, PptxTable, PptxTextBody } from './types'
import { emuToPx } from '../core/geometry'
import { resolveColor } from '../core/color'

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
): void {
  const m = metrics ?? { widthPx: Math.round(emuToPx(slide.widthEmu)), heightPx: Math.round(emuToPx(slide.heightEmu)) }
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, m.widthPx, m.heightPx)
  // after the fill so it is not erased, before the shapes so it sits under them
  if (watermark) paintWatermark(ctx, { widthPx: m.widthPx, heightPx: m.heightPx }, watermark)
  for (const shape of slide.shapes) {
    paintShape(shape, ctx, images)
  }
}

function paintShape(shape: PptxShape, ctx: CanvasRenderingContext2D, images?: Array<CanvasImageSource | undefined>): void {
  const x = emuToPx(shape.xEmu)
  const y = emuToPx(shape.yEmu)
  const w = emuToPx(shape.widthEmu)
  const h = emuToPx(shape.heightEmu)
  ctx.save()
  if (shape.rotationDeg) {
    ctx.translate(x + w / 2, y + h / 2)
    ctx.rotate((shape.rotationDeg * Math.PI) / 180)
    ctx.translate(-(x + w / 2), -(y + h / 2))
  }
  // fill + outline
  const path = () => {
    ctx.beginPath()
    if (shape.geometry === 'ellipse') {
      ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
    } else if (shape.geometry === 'roundRect') {
      const r = Math.min(w, h) * 0.15
      ctx.roundRect(x, y, w, h, r)
    } else {
      ctx.rect(x, y, w, h)
    }
  }
  if (shape.fill) {
    path()
    ctx.fillStyle = resolveColor(shape.fill)
    ctx.fill()
  }
  if (shape.line) {
    path()
    ctx.strokeStyle = resolveColor(shape.line.color)
    ctx.lineWidth = Math.max(1, emuToPx(shape.line.widthEmu ?? 12700))
    ctx.stroke()
  }
  // table (p:graphicFrame/a:tbl): fills, grid, then cell text
  if (shape.table) {
    paintTable(shape.table, ctx, x, y, w, h)
  }
  // picture (p:pic): drawn over the fill, beneath text
  if (images && shape.imageIndex !== undefined) {
    const img = images[shape.imageIndex]
    if (img) {
      drawImageFitted(img as CanvasImageSource, shape, ctx, x, y, w, h)
    }
  }
  // text
  if (shape.textBody) paintTextBody(shape.textBody, ctx, x, y, w, h)
  ctx.restore()
}

/** Draw a picture into the shape rect, honoring a:srcRect crop. */
function drawImageFitted(
  img: CanvasImageSource,
  shape: PptxShape,
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  const src = shape.image?.srcRect
  if (!src) {
    ctx.drawImage(img, x, y, w, h)
    return
  }
  const iw = naturalWidth(img)
  const ih = naturalHeight(img)
  if (iw <= 0 || ih <= 0) {
    ctx.drawImage(img, x, y, w, h)
    return
  }
  const sx = iw * src.l
  const sy = ih * src.t
  const sw = iw * (1 - src.l - src.r)
  const sh = ih * (1 - src.t - src.b)
  if (sw <= 0 || sh <= 0) return
  ctx.drawImage(img, sx, sy, sw, sh, x, y, w, h)
}

function naturalWidth(img: CanvasImageSource): number {
  const anyImg = img as { width?: number; naturalWidth?: number }
  return anyImg.naturalWidth ?? anyImg.width ?? 0
}

function naturalHeight(img: CanvasImageSource): number {
  const anyImg = img as { height?: number; naturalHeight?: number }
  return anyImg.naturalHeight ?? anyImg.height ?? 0
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
      if (fill) {
        ctx.fillStyle = resolveColor(fill)
        ctx.fillRect(cx, cy, cw, ch)
      }
      if (cell.paragraphs.length > 0) {
        // header rows in a styled table often switch to light text
        const textColor = table.firstRow && ri === 0 ? table.firstRowTextColor : undefined
        const headerBold = table.firstRow && ri === 0 ? table.firstRowBold : undefined
        const paragraphs = textColor || headerBold !== undefined
          ? cell.paragraphs.map((p) => ({
              ...p,
              runs: p.runs.map((r) => ({ ...r, color: r.color ?? textColor, bold: r.bold ?? headerBold })),
            }))
          : cell.paragraphs
        paintTextBody(
          {
            paragraphs,
            anchor: 'ctr',
            insetLeftEmu: CELL_PAD_EMU,
            insetRightEmu: CELL_PAD_EMU,
            insetTopEmu: CELL_VPAD_EMU,
            insetBottomEmu: CELL_VPAD_EMU,
            wrap: true,
          },
          ctx,
          cx,
          cy,
          cw,
          ch,
        )
      }
      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < colSpan; dc++) owners.set(`${ri + dr}:${ci + dc}`, `${ri}:${ci}`)
      }
      ci += physicalColumns ? 1 : colSpan
    }
  })

  // grid: thin lines on every boundary, skipping interiors of merged cells
  const sameOwner = (a: string, b: string) => owners.has(a) && owners.get(a) === owners.get(b)
  const draw = (side: 'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV',
    row: number, x1: number, y1: number, x2: number, y2: number) => {
    const headerSide = side === 'insideH' && row === 1 ? 'bottom' : side
    const headerBoundary = row === 0 || (side === 'insideH' && row === 1)
    const border = table.firstRow && headerBoundary ? table.firstRowBorders?.[headerSide] ?? table.styleBorders?.[side]
      : table.styleBorders?.[side]
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
      draw(c === 0 ? 'left' : c === widths.length ? 'right' : 'insideV', r, gx + gridOffset, y1, gx + gridOffset, y2)
    }
  }
  for (let r = 0; r <= heights.length; r++) {
    const gy = rowY[r] ?? h
    for (let c = 0; c < widths.length; c++) {
      if (r > 0 && r < heights.length && sameOwner(`${r - 1}:${c}`, `${r}:${c}`)) continue
      const x1 = colX[c] ?? 0
      const x2 = colX[c + 1] ?? w
      draw(r === 0 ? 'top' : r === heights.length ? 'bottom' : 'insideH', r, x1, gy + gridOffset, x2, gy + gridOffset)
    }
  }
  ctx.restore()
}

function paintTextBody(body: PptxTextBody, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): void {
  const padL = emuToPx(body.insetLeftEmu)
  const padR = emuToPx(body.insetRightEmu)
  const padT = emuToPx(body.insetTopEmu)
  const padB = emuToPx(body.insetBottomEmu)
  const textW = w - padL - padR
  const textH = h - padT - padB
  if (textW <= 0 || textH <= 0) return

  // measure paragraphs into wrapped lines first (for vertical anchoring)
  interface LaidLine { text: string; fontSizePt: number; heightPx: number; x: number; align: string; bullet: boolean; color?: string; bold?: boolean; italic?: boolean; fontFamily?: string }
  const laid: LaidLine[] = []
  for (const para of body.paragraphs) {
    if (para.runs.length === 0) {
      laid.push({ text: '', fontSizePt: 12, heightPx: 12 * (96 / 72), x: padL, align: para.align, bullet: false })
      continue
    }
    for (const run of para.runs) {
      if (run.text === '\n') {
        laid.push({ text: '', fontSizePt: run.fontSizePt ?? 12, heightPx: (run.fontSizePt ?? 12) * (96 / 72), x: padL, align: para.align, bullet: false })
        continue
      }
      const sizePt = run.fontSizePt ?? 12
      ctx.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${sizePt}pt "${run.fontFamily ?? 'Calibri'}"`
      // After an explicit a:br the next run often starts with a source-format
      // space — trim it so the new line doesn't carry a stray leading gap.
      const words = run.text.replace(/^\s+/, '').split(/(\s+)/).filter((s) => s !== '')
      let line = ''
      let lineW = 0
      for (const word of words) {
        const ww = ctx.measureText(word).width
        if (lineW + ww > textW && line !== '') {
          laid.push({ text: line, fontSizePt: sizePt, heightPx: sizePt * (96 / 72) * 1.2, x: padL, align: para.align, bullet: !!para.bullet, color: run.color, bold: run.bold, italic: run.italic, fontFamily: run.fontFamily })
          line = ''
          lineW = 0
        }
        line += word
        lineW += ww
      }
      if (line) {
        laid.push({ text: line, fontSizePt: sizePt, heightPx: sizePt * (96 / 72) * 1.2, x: padL, align: para.align, bullet: !!para.bullet, color: run.color, bold: run.bold, italic: run.italic, fontFamily: run.fontFamily })
      }
      // measure once more for alignment widths
      void run
    }
  }
  // bullet marking: prefix first line of bulleted paragraph — approximate by
  // tagging lines that start a paragraph; we simplified by bullets per line.
  const totalH = laid.reduce((a, l) => a + l.heightPx, 0)
  let ty: number
  if (body.anchor === 'ctr') ty = y + padT + (textH - totalH) / 2
  else if (body.anchor === 'b') ty = y + padT + textH - totalH
  else ty = y + padT

  ctx.textBaseline = 'alphabetic'
  let lx = 0
  for (const line of laid) {
    const lineH = line.heightPx
    if (line.text) {
      ctx.font = `${line.italic ? 'italic ' : ''}${line.bold ? 'bold ' : ''}${line.fontSizePt}pt "${line.fontFamily ?? 'Calibri'}"`
      const tw = ctx.measureText(line.text).width
      let px = x + padL
      if (line.align === 'center') px = x + padL + (textW - tw) / 2
      else if (line.align === 'right') px = x + padL + textW - tw
      // honor the run color (white text on a dark cell fill is common)
      ctx.fillStyle = line.color ? resolveColor(line.color) : '#000000'
      ctx.fillText(line.text, px, ty + lineH * 0.8)
    }
    ty += lineH
    lx += 1
  }
  void lx
}
