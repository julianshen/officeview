/**
 * The terminal fallback painter (slideshow spec 7, ladder step 3): the
 * current render state painted straight into a 2-D output canvas in one
 * pass, with no layer bitmaps, no textures and no compositor.
 *
 * Visibility, opacity, transform and motion are honoured. Masks and colour
 * matrices coarsen to cuts: the state's `visible` flag already carries the
 * step, and colour changes need a re-raster this path does not do.
 */
import type { ContentPaintAssets } from '../../drawing/content'
import { paintScene } from '../../drawing/scene-paint'
import type { FontResolver } from '../../core/fonts/register'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../../core/watermark'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import type { AllocationOwner, BuildOwnership, DiagnosticSink, Lease, RenderState, TargetKey } from '../types'
import { buildRenderPlan, type PlanLayer } from './plan'
import { imageIndices, sceneOptions, slideSize } from './scene'

export interface DirectAssets {
  images: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont: FontResolver
  /** Device pixels per slide px, composed over the context's current transform. */
  scale: number
  /** Decode/paint scratch: `terminal.scratch` for the terminal path, the presenter lease owner for presenter work. Never another owner. */
  scratch: AllocationOwner
  /** Additive to the brief: painted last, above every shape, unmoved by any state. */
  watermark?: WatermarkOptions | ResolvedWatermark
  /** Additive to the brief: scratch-overflow diagnostics. */
  sink?: DiagnosticSink
  /** Additive to the brief: the slide's build records. Absent, paragraph ranges are read back from the state's `sp:<id>/p:<st>-<end>` keys. */
  builds?: readonly BuildOwnership[]
  contentAssets?: ContentPaintAssets
}

const RANGE_KEY = /^sp:([^#/]+)\/p:(\d+)-(\d+)$/

function buildsFromState(state: RenderState): BuildOwnership[] {
  const bySpid = new Map<string, Array<{ key: TargetKey; st: number; end: number }>>()
  for (const key of state.keys()) {
    const m = RANGE_KEY.exec(key)
    if (!m) continue
    const list = bySpid.get(m[1]) ?? []
    list.push({ key, st: Number(m[2]), end: Number(m[3]) })
    bySpid.set(m[1], list)
  }
  return [...bySpid].map(([enc, ranges]) => {
    let spid = enc
    try { spid = decodeURIComponent(enc) } catch { /* keep the encoded form */ }
    const body = `sp:${enc}`
    return { spid, mode: 'paragraph', ranges: ranges.sort((a, b) => a.st - b.st), animBg: state.has(body), body }
  })
}

function imageBytes(img: CanvasImageSource): number {
  const any = img as { naturalWidth?: number; naturalHeight?: number; width?: number | { baseVal?: { value: number } }; height?: number | { baseVal?: { value: number } } }
  const w = any.naturalWidth ?? (typeof any.width === 'number' ? any.width : 0)
  const h = any.naturalHeight ?? (typeof any.height === 'number' ? any.height : 0)
  return Math.max(1, Math.ceil(w)) * Math.max(1, Math.ceil(h)) * 4
}

export function paintDirect(ctx: CanvasRenderingContext2D, slide: PptxSlide, state: RenderState, assets: DirectAssets): void {
  const plan = buildRenderPlan(slide, new Set(state.keys()), new Map(), assets.builds ?? buildsFromState(state))
  const { w, h } = slideSize(slide)
  const leases: Lease[] = []
  // Images are admitted against the scratch owner as they are first painted; a refused one stays undefined.
  const admitted: Array<CanvasImageSource | undefined> = []
  const decided = new Set<number>()
  const withheld = new Set<number>()
  const admit = (nodes: readonly PptxShape[]): void => {
    for (const index of imageIndices(nodes)) {
      if (decided.has(index)) continue
      decided.add(index)
      const img = assets.images[index]
      if (!img) continue
      const lease = assets.scratch.lease(imageBytes(img), `direct:image:${index}`)
      if (lease) { leases.push(lease); admitted[index] = img; continue }
      withheld.add(index)
      assets.sink?.add({
        code: 'scratch-image', slideIndex: slide.index, effect: `image:${index}`,
        message: `Image ${index} does not fit the direct-paint scratch allowance; painted as the missing-image placeholder`,
      })
    }
  }
  const options = sceneOptions(slide, {
    images: admitted, resolveFont: assets.resolveFont, contentAssets: assets.contentAssets,
    withheld: shape => shape.imageIndex !== undefined && withheld.has(shape.imageIndex),
  })

  const paint = (layer: PlanLayer): void => {
    if (layer.kind === 'watermark') return
    if (layer.kind === 'static') { admit(layer.nodes); paintScene(layer.nodes, ctx, options); return }
    const s = state.get(layer.key)
    if (s && !s.visible) return
    ctx.save()
    try {
      if (s) {
        if (s.motion && Number.isFinite(s.motion.dx) && Number.isFinite(s.motion.dy)) ctx.translate(s.motion.dx, s.motion.dy)
        if (s.transform.every(Number.isFinite)) ctx.transform(...s.transform)
        if (Number.isFinite(s.opacity)) ctx.globalAlpha *= Math.min(1, Math.max(0, s.opacity))
      }
      admit(layer.nodes)
      paintScene(layer.nodes, ctx, options)
      for (const sub of layer.sub ?? []) paint(sub)
    } finally { ctx.restore() }
  }

  ctx.save()
  try {
    ctx.scale(assets.scale, assets.scale)
    ctx.fillStyle = slide.background ?? '#ffffff'
    ctx.fillRect(0, 0, w, h)
    for (const layer of plan.layers) paint(layer)
    // Always last, outside every target's transform.
    if (assets.watermark) paintWatermark(ctx, { widthPx: w, heightPx: h, background: slide.background }, assets.watermark)
  } finally {
    ctx.restore()
    for (const lease of leases) lease.release()
  }
}
