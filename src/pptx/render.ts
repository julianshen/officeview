/**
 * PPTX rendering: paint slides onto a canvas 2D context. 1 unit = 1 px;
 * EMU geometry converted with emuToPx at 96dpi default.
 */
import type { PptxDocument, PptxShape, PptxSlide, PptxTextBody } from './types'
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
): void {
  const m = metrics ?? { widthPx: Math.round(emuToPx(slide.widthEmu)), heightPx: Math.round(emuToPx(slide.heightEmu)) }
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, m.widthPx, m.heightPx)
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

function paintTextBody(body: PptxTextBody, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): void {
  const padL = emuToPx(body.insetLeftEmu)
  const padR = emuToPx(body.insetRightEmu)
  const padT = emuToPx(body.insetTopEmu)
  const padB = emuToPx(body.insetBottomEmu)
  const textW = w - padL - padR
  const textH = h - padT - padB
  if (textW <= 0 || textH <= 0) return

  // measure paragraphs into wrapped lines first (for vertical anchoring)
  interface LaidLine { text: string; fontSizePt: number; heightPx: number; x: number; align: string; bullet: boolean }
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
          laid.push({ text: line, fontSizePt: sizePt, heightPx: sizePt * (96 / 72) * 1.2, x: padL, align: para.align, bullet: !!para.bullet })
          line = ''
          lineW = 0
        }
        line += word
        lineW += ww
      }
      if (line) {
        laid.push({ text: line, fontSizePt: sizePt, heightPx: sizePt * (96 / 72) * 1.2, x: padL, align: para.align, bullet: !!para.bullet })
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
      ctx.font = `${line.fontSizePt}pt "Calibri"`
      const tw = ctx.measureText(line.text).width
      let px = x + padL
      if (line.align === 'center') px = x + padL + (textW - tw) / 2
      else if (line.align === 'right') px = x + padL + textW - tw
      ctx.fillStyle = '#000000'
      ctx.fillText(line.text, px, ty + lineH * 0.8)
    }
    ty += lineH
    lx += 1
  }
  void lx
}
