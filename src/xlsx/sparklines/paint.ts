import type { XlsxSheet } from '../types'
import { sparklineGeometry, type SparkPrimitive, type SparkScale } from './geometry'
import type { ResolvedSparkline } from './resolve'
import type { Rect } from '../view/textLayout'

export type SparklinePainter = (ctx: CanvasRenderingContext2D, sheet: XlsxSheet, cellRect: (row: number, col: number) => Rect | undefined, visible: { r0: number; c0: number; r1: number; c1: number }, scale: SparkScale) => boolean

function cssColor(color: string): string {
  const hex = color.replace(/^#/, '')
  return `#${hex.length === 8 ? hex.slice(2) : hex}`
}

/** Apply the viewport's device-space translation after geometry has snapped on the sheet-global grid. */
export function translateSparkPrimitive(primitive: SparkPrimitive, offsetX: number, offsetY: number): SparkPrimitive {
  if (primitive.kind === 'polyline') return { ...primitive, points: primitive.points.map(([x, y]) => [x + offsetX, y + offsetY] as [number, number]), ...(primitive.band ? { band: { y0: primitive.band.y0 + offsetY, y1: primitive.band.y1 + offsetY } } : {}) }
  if (primitive.kind === 'axis') return { ...primitive, x0: primitive.x0 + offsetX, x1: primitive.x1 + offsetX, y: primitive.y + offsetY }
  if (primitive.kind === 'rect') return { ...primitive, x: primitive.x + offsetX, y: primitive.y + offsetY }
  return { ...primitive, x: primitive.x + offsetX, y: primitive.y + offsetY }
}

export function sparklinePrimitivesForViewport(sparkline: ResolvedSparkline, cell: Rect, scale: SparkScale, offsetX: number, offsetY: number): SparkPrimitive[] {
  return sparklineGeometry(sparkline, { x: cell.x, y: cell.y, w: cell.width, h: cell.height }, scale).map(primitive => translateSparkPrimitive(primitive, offsetX, offsetY))
}

const IDENTITY = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } as DOMMatrix

function paintPrimitive(ctx: CanvasRenderingContext2D, primitive: SparkPrimitive, transform: DOMMatrix, deviceSpace: boolean): void {
  ctx.fillStyle = cssColor(primitive.color)
  ctx.strokeStyle = cssColor(primitive.color)
  const scale = deviceSpace ? 1 : Math.hypot(transform.a, transform.b)
  const point = (x: number, y: number): [number, number] => deviceSpace
    ? [x + transform.e, y + transform.f]
    : [transform.a * x + transform.c * y + transform.e, transform.b * x + transform.d * y + transform.f]
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  if (primitive.kind === 'polyline') {
    if (primitive.points.length >= 2) {
      if (primitive.band) {
        // lineClipRule: the stroke is clipped to the plot box rows (the axis-aligned band survives any scale transform).
        const y0 = point(0, primitive.band.y0)[1], y1 = point(0, primitive.band.y1)[1]
        ctx.beginPath(); ctx.rect(-1e6, Math.min(y0, y1), 2e6, Math.abs(y1 - y0)); ctx.clip()
      }
      ctx.lineWidth = primitive.widthPx * scale; ctx.lineJoin = primitive.join; ctx.lineCap = primitive.cap
      ctx.beginPath()
      const first = point(...primitive.points[0]); ctx.moveTo(...first)
      for (let i = 1; i < primitive.points.length; i++) { const next = point(...primitive.points[i]); ctx.lineTo(...next) }
      ctx.stroke()
    }
  } else if (primitive.kind === 'axis') {
    const a = point(primitive.x0, primitive.y), b = point(primitive.x1, primitive.y)
    ctx.beginPath(); ctx.moveTo(...a); ctx.lineTo(...b); ctx.lineWidth = primitive.widthPx * scale; ctx.stroke()
  } else if (primitive.kind === 'rect') {
    const [x, y] = point(primitive.x, primitive.y); ctx.fillRect(x, y, primitive.w * scale, primitive.h * scale)
  } else {
    const [x, y] = point(primitive.x, primitive.y), size = primitive.sizePx * scale
    if (primitive.shape === 'circle') { ctx.beginPath(); ctx.arc(x, y, size, 0, Math.PI * 2); ctx.fill() }
    else if (primitive.shape === 'square') ctx.fillRect(x - size, y - size, size * 2, size * 2)
    else { ctx.beginPath(); ctx.moveTo(x, y - size); ctx.lineTo(x + size, y); ctx.lineTo(x, y + size); ctx.lineTo(x - size, y); ctx.closePath(); ctx.fill() }
  }
  ctx.restore()
}

/** Lazy-chunk implementation: geometry remains outside the eager renderer. */
export const paintSparklinesIn: SparklinePainter = (ctx, sheet, cellRect, visible, scale) => {
  for (const [row, columns] of sheet.resolvedSparklines ?? []) {
    if (row < visible.r0 || row > visible.r1) continue
    for (const [col, sparkline] of columns) {
      if (col < visible.c0 || col > visible.c1) continue
      const rect = cellRect(row, col)
      if (!rect || !(rect.width > 0 && rect.height > 0)) continue
      const transform = ctx.getTransform()
      ctx.save()
      try {
        ctx.beginPath(); ctx.rect(rect.x, rect.y, rect.width, rect.height); ctx.clip()
        const deviceSpace = scale.output === 'device'
        const primitives = deviceSpace
          ? sparklinePrimitivesForViewport(sparkline as ResolvedSparkline, rect, scale, transform.e, transform.f)
          : sparklineGeometry(sparkline as ResolvedSparkline, { x: rect.x, y: rect.y, w: rect.width, h: rect.height }, scale)
        for (const primitive of primitives) paintPrimitive(ctx, primitive, deviceSpace ? IDENTITY : transform, deviceSpace)
      } finally { ctx.restore() }
    }
  }
  return true
}
