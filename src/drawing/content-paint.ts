import { emuToPx } from '../core/geometry'
import { resolveGeometry } from './geometry'
import { paintGeometry } from './paint'
import type { DrawingContent, DrawingContentShape, ContentPaintAssets } from './content'
import type { XlsxChartModel } from '../xlsx/chartModel'
type XlsxChartPainter = (chart: XlsxChartModel, ctx: CanvasRenderingContext2D, width: number, height: number, familyCss?: (family: string) => string) => void
/** Set by the lazy chart engine when it loads (src/charts). */
let xlsxChartPainter: XlsxChartPainter | undefined
export function registerXlsxChartPainter(painter: XlsxChartPainter): void { xlsxChartPainter = painter }
/** A chart that is not drawn: a card with its title and the reason, never blank space. */
export function paintChartPlaceholder(chart: XlsxChartModel, ctx: CanvasRenderingContext2D, w: number, h: number, familyCss: (family: string) => string = JSON.stringify): void {
  ctx.save()
  try {
    ctx.beginPath(); ctx.rect(0, 0, w, h); ctx.clip()
    ctx.fillStyle = '#F7F7F7'; ctx.fillRect(0, 0, w, h)
    ctx.strokeStyle = '#BFBFBF'; ctx.lineWidth = 1; ctx.setLineDash([4, 3]); ctx.strokeRect(0.5, 0.5, w - 1, h - 1); ctx.setLineDash([])
    const s = Math.min(24, w / 6, h / 6), x = w / 2 - s / 2, y = h / 2 - s - 6
    ctx.fillStyle = '#A6A6A6'
    for (const [i, f] of [0.45, 0.8, 0.6].entries()) ctx.fillRect(x + i * s / 3 + 1, y + s * (1 - f), s / 3 - 2, s * f)
    ctx.fillStyle = '#595959'; ctx.textAlign = 'center'; ctx.textBaseline = 'top'
    ctx.font = `bold 12px ${familyCss(chart.fontFamily)}`
    ctx.fillText(chart.title ?? 'Chart', w / 2, h / 2, w - 8)
    ctx.font = `11px ${familyCss(chart.fontFamily)}`
    ctx.fillText(chart.unsupported ?? 'Chart not loaded', w / 2, h / 2 + 16, w - 8)
  } finally { ctx.restore() }
}
export interface ContentPaintOptions<Paragraph, Text> {
  paintDiagramText?: (shape: DrawingContentShape<Text>, ctx: CanvasRenderingContext2D, width: number, height: number) => void
  paintTextbox?: (drawing: Extract<DrawingContent<Paragraph, Text>, { kind: 'textbox' }>, ctx: CanvasRenderingContext2D, width: number, height: number, assets?: ContentPaintAssets) => void
  fontFamilyCss?: (family: string) => string
  assets?: ContentPaintAssets
}
export function paintDrawingContent<Paragraph, Text>(drawing: DrawingContent<Paragraph, Text>, ctx: CanvasRenderingContext2D, width: number, height: number, options: ContentPaintOptions<Paragraph, Text> = {}): void {
  if (drawing.kind === 'ink') paintInk(drawing, ctx, width, height)
  else if (drawing.kind === 'diagram') {
    const paintShapes = (shapes: DrawingContentShape<Text>[]): void => {
    for (const s of shapes) {
      const x = emuToPx(s.xEmu),
        y = emuToPx(s.yEmu),
        w = emuToPx(s.widthEmu),
        h = emuToPx(s.heightEmu)
      if (![x, y, w, h, s.rotationDeg ?? 0].every(Number.isFinite) || w < 0 || h < 0) continue
      ctx.save()
      try {
      ctx.translate(x + w / 2, y + h / 2)
      ctx.rotate(((s.rotationDeg ?? 0) * Math.PI) / 180)
      if (s.flipH || s.flipV) ctx.scale(s.flipH ? -1 : 1, s.flipV ? -1 : 1)
      ctx.translate(-w / 2, -h / 2)
      if (s.group) {
        const g = s.group
        const sx = g.ext.width / g.chExt.width, sy = g.ext.height / g.chExt.height
        const cx = emuToPx(g.chOff.x), cy = emuToPx(g.chOff.y)
        if (![sx, sy, cx, cy].every(Number.isFinite) || g.ext.width < 0 || g.ext.height < 0 || g.chExt.width <= 0 || g.chExt.height <= 0) continue
        ctx.scale(sx, sy)
        ctx.translate(-cx, -cy)
        paintShapes(s.children ?? [])
        continue
      }
      if (s.image) {
        const img = options.assets?.imageFor?.(s.image)
        if (img && w > 0 && h > 0) ctx.drawImage(img, 0, 0, w, h)
      } else if (s.drawingGeometry) {
        paintGeometry(ctx, resolveGeometry(s.drawingGeometry, w, h), s.drawingStyle ?? { issues: [] }, w, h)
      } else {
        ctx.beginPath()
        if (s.geometry === 'ellipse') ctx.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
        else if (s.geometry === 'rightArrow') {
          ctx.moveTo(0, h * 0.2)
          ctx.lineTo(w * 0.5, h * 0.2)
          ctx.lineTo(w * 0.5, 0)
          ctx.lineTo(w, h / 2)
          ctx.lineTo(w * 0.5, h)
          ctx.lineTo(w * 0.5, h * 0.8)
          ctx.lineTo(0, h * 0.8)
          ctx.closePath()
        } else ctx.rect(0, 0, w, h)
        if (s.fill) {
          ctx.fillStyle = `#${s.fill}`
          ctx.fill()
        }
        if (s.line) {
          ctx.strokeStyle = `#${s.line.color}`
          ctx.lineWidth = emuToPx(s.line.widthEmu)
          ctx.stroke()
        }
      }
      // Text-only fallback diagrams carry no intrinsic extents; offer the
      // enclosing diagram box instead of a degenerate one (which text
      // painters early-return on, losing all visible ink). Gated on the
      // fallback marker — every other diagram, including degenerate cached
      // shapes, keeps authored boxes byte-identically.
      const fallbackBox = drawing.textOnly === true
      const textW = w > 0 || !fallbackBox ? w : width
      const textH = h > 0 || !fallbackBox ? h : height
      try { options.paintDiagramText?.(s, ctx, textW, textH) }
      catch { /* A malformed text body cannot suppress later shapes. */ }
      } finally { ctx.restore() }
    }
    }
    paintShapes(drawing.shapes)
  } else if (drawing.kind === 'xlsxChart') (drawing.chart.unsupported || !xlsxChartPainter ? paintChartPlaceholder : xlsxChartPainter)(drawing.chart, ctx, width, height, options.fontFamilyCss)
  else {
    if (drawing.fill) {
      ctx.fillStyle = `#${drawing.fill}`
      ctx.fillRect(0, 0, width, height)
    }
    if (drawing.line) {
      ctx.strokeStyle = `#${drawing.line.color}`
      ctx.lineWidth = emuToPx(drawing.line.widthEmu)
      ctx.strokeRect(0, 0, width, height)
    }
    options.paintTextbox?.(drawing, ctx, width, height, options.assets)
  }
}

export function paintInk(
  ink: Extract<DrawingContent, { kind: 'ink' }>,
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number
): void {
  const sx = ink.widthEmu > 0 ? width / ink.widthEmu : 1 / 9525
  const sy = ink.heightEmu > 0 ? height / ink.heightEmu : 1 / 9525
  ctx.lineCap = 'round'
  ctx.lineJoin = 'round'
  for (const stroke of ink.strokes) {
    ctx.strokeStyle = ctx.fillStyle = `#${stroke.color}`
    ctx.lineWidth = (stroke.widthEmu * (sx + sy)) / 2
    ctx.beginPath()
    stroke.points.forEach(([x, y], index) => {
      if (index === 0) ctx.moveTo(x * sx, y * sy)
      else ctx.lineTo(x * sx, y * sy)
    })
    if (stroke.points.length === 1) {
      const [x, y] = stroke.points[0]
      ctx.arc(x * sx, y * sy, ctx.lineWidth / 2, 0, Math.PI * 2)
      ctx.fill()
    } else ctx.stroke()
  }
}
