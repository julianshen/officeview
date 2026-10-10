import { emuToPx } from '../core/geometry'
import { resolveGeometry } from './geometry'
import { paintGeometry } from './paint'
import type { DrawingContent, DrawingContentShape, ContentTextRun, ContentPaintAssets } from './content'
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
      if (s.drawingGeometry) {
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
  } else if (drawing.kind === 'chart') paintChart(drawing, ctx, width, height, options.fontFamilyCss)
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

function font(run: ContentTextRun, family: string, size: number, familyCss: (family: string) => string = JSON.stringify): string {
  return `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSizePt ?? size}pt ${familyCss(run.fontFamily ?? family)}`
}
function paintChart(
  chart: Extract<DrawingContent, { kind: 'chart' }>,
  ctx: CanvasRenderingContext2D,
  w: number,
  h: number,
  familyCss?: (family: string) => string
): void {
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, w, h)
  ctx.strokeStyle = '#D9D9D9'
  ctx.lineWidth = 1
  ctx.strokeRect(0, 0, w, h)
  const left = w * 0.046,
    right = w * 0.026,
    top = h * (chart.title?.length ? 0.152 : 0.09),
    bottom = h * (chart.legend ? 0.185 : 0.14)
  const pw = w - left - right,
    ph = h - top - bottom
  const values = chart.series.flatMap((s) => s.values).filter((v): v is number => v !== undefined)
  const min = chart.min ?? Math.min(0, ...values),
    maximum = Math.max(1, ...values)
  const raw = ((maximum - min) * 1.1) / 6,
    base = 10 ** Math.floor(Math.log10(raw)),
    step =
      chart.majorUnit && chart.majorUnit > 0
        ? chart.majorUnit
        : ([1, 2, 5, 10].map((v) => v * base).find((v) => v >= raw) ?? base * 10)
  const max = chart.max ?? Math.ceil(((maximum - min) * 1.1) / step) * step + min,
    range = Math.max(1, max - min)
  ctx.font = font({ text: '' }, chart.fontFamily, chart.fontSizePt, familyCss)
  ctx.fillStyle = '#595959'
  ctx.textAlign = 'right'
  ctx.textBaseline = 'middle'
  for (let value = min, count = 0; value <= max + step / 1000 && count < 100; value += step, count++) {
    const y = top + ph - ((value - min) / range) * ph
    ctx.strokeStyle = '#D9D9D9'
    ctx.beginPath()
    ctx.moveTo(left, y)
    ctx.lineTo(w - right, y)
    ctx.stroke()
    ctx.fillText(String(Number(value.toPrecision(8))), left - 8, y)
  }
  const categories = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1),
    group = pw / categories
  const stride = 1 - chart.overlap / 100,
    span = 1 + (Math.max(1, chart.series.length) - 1) * stride
  const bar = group / (span + chart.gapWidth / 100),
    zero = top + ph - ((0 - min) / range) * ph
  chart.series.forEach((s, si) =>
    s.values.forEach((value, ci) => {
      if (value === undefined) return
      const x = left + group * ci + (group - bar * span) / 2 + bar * si * stride,
        y = top + ph - ((value - min) / range) * ph
      ctx.fillStyle = `#${s.color}`
      ctx.fillRect(x, Math.min(y, zero), bar, Math.abs(zero - y))
    })
  )
  ctx.fillStyle = '#595959'
  ctx.textAlign = 'center'
  ctx.textBaseline = 'top'
  chart.categories.forEach((label, i) => ctx.fillText(label, left + group * (i + 0.5), top + ph + 10))
  if (chart.title) {
    ctx.font = font({ text: '' }, chart.fontFamily, 14, familyCss)
    ctx.fillText(chart.title, w / 2, 14)
  }
  if (chart.legend) {
    ctx.font = font({ text: '' }, chart.fontFamily, chart.fontSizePt, familyCss)
    const sizes = chart.series.map((s) => ctx.measureText(s.name).width + 17),
      total = sizes.reduce((a, b) => a + b, 0)
    let x = (w - total) / 2
    const y = h - 24
    chart.series.forEach((s, i) => {
      ctx.fillStyle = `#${s.color}`
      ctx.fillRect(x, y + 3, 6, 6)
      ctx.fillStyle = '#595959'
      ctx.textAlign = 'left'
      ctx.fillText(s.name, x + 9, y)
      x += sizes[i]
    })
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
