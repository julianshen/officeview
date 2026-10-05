import { resolveColor } from '../core/color'
import { emuToPx } from '../core/geometry'
import { resolveGeometry } from './geometry'
import { paintGeometry } from './paint'
import type { SceneNode } from './scene'

/** Recursive structural constraint preserves the adapter's node type in callbacks. */
type SceneTree<Node> = SceneNode & { children?: Node[] }

export interface ScenePaintOptions<Node> {
  images?: ReadonlyArray<CanvasImageSource | undefined>
  /** Format-specific content (such as tables), after geometry and before pictures. */
  paintContent?: (node: Node, ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number) => void
  /** Adapters retain their existing text painter; no parallel text layout here. */
  paintText?: (node: Node, ctx: CanvasRenderingContext2D, widthPx: number, heightPx: number) => void
  missingImage?: (node: Node) => void
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
    if (shape.drawingGeometry) {
      const geometry = resolveGeometry(shape.drawingGeometry, w, h)
      paintGeometry(ctx, geometry, shape.drawingStyle ?? { issues: [] }, w, h)
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
    if (options.images && shape.imageIndex !== undefined) {
      const img = options.images[shape.imageIndex]
      if (img) {
        const opacity = shape.image?.opacity ?? 1
        if (Number.isFinite(opacity) && opacity >= 0 && opacity <= 1) {
          ctx.save()
          try { ctx.globalAlpha *= opacity; drawImageFitted(img, shape, ctx, 0, 0, w, h) }
          finally { ctx.restore() }
        }
      }
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
