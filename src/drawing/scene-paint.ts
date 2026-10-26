import { resolveColor } from '../core/color'
import { emuToPx } from '../core/geometry'
import { resolveGeometry } from './geometry'
import { paintGeometry, tracePath } from './paint'
import type { SceneNode } from './scene'
import { paintChartPlaceholder } from './content-paint'

/** Recursive structural constraint preserves the adapter's node type in callbacks. */
type SceneTree<Node> = SceneNode & { children?: Node[] }

export interface ScenePaintOptions<Node> {
  images?: ReadonlyArray<CanvasImageSource | undefined>
  /** Format-specific content (such as tables), after geometry and before pictures. */
  paintContent?: (node: Node, ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number) => void
  /** Adapters retain their existing text painter; no parallel text layout here. */
  paintText?: (node: Node, ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number) => void
  missingImage?: (node: Node) => void
  /** The page background, in the coordinate space `root` names: `useBgFill` nodes show it through their filled geometry. */
  backdrop?: readonly Node[]
  root?: DOMMatrix
  /** PowerPoint's device-pixel rule for geometry (measured from its movie export); see `snapToPixels`. */
  pixelSnap?: boolean
}

const roundHalfUp = (v: number) => Math.floor(v + 0.5 + 1e-7)
/**
 * Maps an axis-aligned `w`×`h` box onto whole device pixels, edges rounded half up, as PowerPoint
 * rasterises shapes: fills end crisply on pixel boundaries, and `paintGeometry`'s `pixelSnap` then puts
 * a W px line on pixels [s - floor(W/2), s - floor(W/2) + W) around each snapped edge s. An axis
 * whose extent rounds to nothing keeps its scale; rotated or skewed transforms are left alone.
 */
function snapToPixels(ctx: CanvasRenderingContext2D, w: number, h: number): void {
  const m = ctx.getTransform()
  if (m.b !== 0 || m.c !== 0) return
  const axis = (origin: number, scale: number, size: number): [number, number] => {
    const a = roundHalfUp(origin), b = roundHalfUp(origin + scale * size)
    return size > 0 && a !== b ? [(b - a) / size, a] : [scale, size > 0 ? origin : a]
  }
  const [a, e] = axis(m.e, m.a, w), [d, f] = axis(m.f, m.d, h)
  ctx.setTransform(a, 0, 0, d, e, f)
}

/** Paint source-ordered objects onto the existing caller surface and transform.
 * Does not clear a background, paint watermarks, decode assets or assign indices.
 */
export function paintScene<Node extends SceneTree<Node>>(
  nodes: readonly Node[],
  ctx: CanvasRenderingContext2D,
  options: ScenePaintOptions<Node> = {},
): void {
  for (const node of nodes) paintNode(node, ctx, options)
}

function paintNode<Node extends SceneTree<Node>>(shape: Node, ctx: CanvasRenderingContext2D, options: ScenePaintOptions<Node>): void {
  const x = emuToPx(shape.xEmu), y = emuToPx(shape.yEmu)
  const w = emuToPx(shape.widthEmu), h = emuToPx(shape.heightEmu)
  if (shape.transformValid === false || ![x, y, w, h, shape.rotationDeg ?? 0].every(Number.isFinite) || w < 0 || h < 0) return
  ctx.save()
  try {
    // Placement composes with the parent's child coordinate system. Flips precede rotation.
    ctx.translate(x + w / 2, y + h / 2)
    if (shape.rotationDeg) ctx.rotate(shape.rotationDeg * Math.PI / 180)
    if (shape.flipH || shape.flipV) ctx.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1)
    ctx.translate(-w / 2, -h / 2)
    if (shape.group) {
      const group = shape.group
      if (![group.ext.width, group.ext.height, group.chExt.width, group.chExt.height].every(Number.isFinite)
        || group.ext.width < 0 || group.ext.height < 0 || group.chExt.width <= 0 || group.chExt.height <= 0) return
      const sx = group.ext.width / group.chExt.width, sy = group.ext.height / group.chExt.height
      const cx = emuToPx(group.chOff.x), cy = emuToPx(group.chOff.y)
      if (![sx, sy, cx, cy].every(Number.isFinite)) return
      ctx.scale(sx, sy)
      ctx.translate(-cx, -cy)
      paintScene<Node>(shape.children ?? [], ctx, options)
      return
    }
    // useBgFill (ECMA-376 p:sp@useBgFill) is the shape's fill: it replaces an authored solid, gradient or picture fill.
    const backdropped = !!(shape.useBgFill && shape.drawingGeometry && options.backdrop && options.root)
    if (shape.drawingGeometry) {
      const geometry = resolveGeometry(shape.drawingGeometry, w, h)
      const style = shape.drawingStyle ?? { issues: [] }
      const snap = !!options.pixelSnap
      ctx.save()
      try {
        if (snap) snapToPixels(ctx, w, h)
        if (backdropped) {
          // The shape's silhouette casts its authored shadow (an opaque stand-in fill; only the shadow pass paints).
          if (style.shadow) paintGeometry(ctx, geometry, { issues: [], shadow: style.shadow, line: style.line, fill: { kind: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } } }, w, h, 'shadow')
          // The background shows through the shape as it lies on the page: clip here, paint it in page space.
          ctx.save()
          try {
            ctx.beginPath()
            for (const path of geometry.paths) if (path.fill !== 'none') tracePath(ctx, path.commands)
            ctx.clip()
            ctx.setTransform(options.root!)
            paintScene(options.backdrop!, ctx, { images: options.images, missingImage: options.missingImage, pixelSnap: options.pixelSnap })
          } finally { ctx.restore() }
          paintGeometry(ctx, geometry, { ...style, fill: undefined, shadow: undefined }, w, h, undefined, snap)
        } else paintGeometry(ctx, geometry, style, w, h, undefined, snap)
      } finally { ctx.restore() }
    } else {
      // Legacy handwritten models keep their existing simple geometry contract.
      const path = () => {
        ctx.beginPath()
        if (shape.geometry === 'ellipse') ctx.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
        else if (shape.geometry === 'roundRect') ctx.roundRect(0, 0, w, h, Math.min(w, h) * .15)
        else ctx.rect(0, 0, w, h)
      }
      if (shape.fill) { path(); ctx.fillStyle = resolveColor(shape.fill); ctx.fill() }
      if (shape.line) { path(); ctx.strokeStyle = resolveColor(shape.line.color); ctx.lineWidth = Math.max(1, emuToPx(shape.line.widthEmu ?? 12700)); ctx.stroke() }
    }
    options.paintContent?.(shape, ctx, w, h)
    if (options.images && shape.imageIndex !== undefined && !backdropped) {
      const img = options.images[shape.imageIndex]
      if (img) {
        const opacity = shape.image?.opacity ?? 1
        if (Number.isFinite(opacity) && opacity >= 0 && opacity <= 1) {
          ctx.save()
          try { ctx.globalAlpha *= opacity; drawImageFitted(img, shape, ctx, 0, 0, w, h) }
          finally { ctx.restore() }
        }
      }
      // A chart's selected fallback picture that could not be decoded paints the chart's card (spec 6.3).
      else if (shape.chartFallback) paintChartPlaceholder(shape.chartFallback, ctx, w, h)
      else options.missingImage?.(shape)
    }
    if (shape.textBody) options.paintText?.(shape, ctx, w, h)
  } finally { ctx.restore() }
}

/** Draw a picture into the local rect, honoring source crop and destination fillRect. */
function drawImageFitted(img: CanvasImageSource, shape: SceneNode, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): void {
  const fill = shape.image?.fillRect
  if (fill) {
    x += w * fill.l
    y += h * fill.t
    w *= 1 - fill.l - fill.r
    h *= 1 - fill.t - fill.b
  }
  if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) return
  const src = shape.image?.srcRect
  if (src && ![src.l, src.t, src.r, src.b].every(Number.isFinite)) return
  if (!src) { ctx.drawImage(img, x, y, w, h); return }
  const iw = naturalWidth(img), ih = naturalHeight(img)
  if (iw <= 0 || ih <= 0) { ctx.drawImage(img, x, y, w, h); return }
  const sx = iw * src.l, sy = ih * src.t
  const sw = iw * (1 - src.l - src.r), sh = ih * (1 - src.t - src.b)
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
