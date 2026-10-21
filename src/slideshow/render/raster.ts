/**
 * The layer rasteriser (slideshow spec 6.3): one plan layer into one cropped
 * 2-D bitmap, charged to the workspace before anything is allocated.
 */
import type { ContentPaintAssets } from '../../drawing/content'
import { paintScene } from '../../drawing/scene-paint'
import type { FontResolver } from '../../core/fonts/register'
import { paintWatermark, type ResolvedWatermark, type WatermarkOptions } from '../../core/watermark'
import type { PptxSlide } from '../../pptx/types'
import type { CanvasFactory, DiagnosticSink, LayerBitmap, Lease, SlideShowWorkspace } from '../types'
import type { PlanLayer } from './plan'
import { hasText, paintedBounds, sceneOptions, slideSize, type Box } from './scene'

export interface RasterContext {
  /** Device pixels per slide px. */
  scale: number
  images: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont: FontResolver
  canvas: CanvasFactory
  workspace: SlideShowWorkspace
  maxTextureSize: number
  /** Additive to the brief: the watermark layer's mark. Absent means the watermark layer is empty. */
  watermark?: WatermarkOptions | ResolvedWatermark
  /** Additive to the brief: where reduced-scale and refusal diagnostics go. */
  sink?: DiagnosticSink
  contentAssets?: ContentPaintAssets
}

/** A painted layer and the lease that charges its bitmap; the caller hands the lease to `Compositor.upload` or releases it. */
export interface RasteredLayer extends LayerBitmap { lease: Lease }

/** `empty`: the layer paints nothing (a group target with `sub` paints through its sub-layers). */
export type RasterResult = RasteredLayer | { refused: 'workspace' | 'empty' }

type Ctx2D = CanvasRenderingContext2D

function tagOf(layer: PlanLayer): string {
  return layer.kind === 'target' ? `layer:${layer.key}${layer.part ? `:${layer.part}` : ''}` : `layer:${layer.kind}`
}

export function rasterLayer(layer: PlanLayer, slide: PptxSlide, ctx: RasterContext): RasterResult {
  const { w: slideW, h: slideH } = slideSize(slide)
  const tag = tagOf(layer)
  const refuse = (): RasterResult => {
    ctx.sink?.add({ code: 'layer-workspace', slideIndex: slide.index, effect: tag, message: `Layer ${tag} did not fit the slideshow workspace` })
    return { refused: 'workspace' }
  }

  let bounds: Box | undefined
  if (layer.kind === 'watermark') bounds = ctx.watermark ? { minX: 0, minY: 0, maxX: slideW, maxY: slideH } : undefined
  else if (layer.kind === 'static' && layer.background) bounds = { minX: 0, minY: 0, maxX: slideW, maxY: slideH }
  else if (layer.nodes.length > 0) {
    let measure: Ctx2D | undefined
    let measureLease: Lease | undefined
    if (hasText(layer.nodes)) {
      // Text can overflow its box; measuring needs a context, which is charged first like any other.
      measureLease = ctx.workspace.lease(4, `${tag}:measure`)
      if (!measureLease) return refuse()
      try { measure = ctx.canvas(1, 1).getContext('2d') as unknown as Ctx2D }
      catch (error) { measureLease.release(); throw error }
    }
    try { bounds = paintedBounds(layer.nodes, slide, ctx.resolveFont, measure) }
    finally { measureLease?.release() }
  }
  if (!bounds) return { refused: 'empty' }

  // Keep a slide's worth of margin for content that animates in from off-slide; drop the rest.
  bounds = {
    minX: Math.max(bounds.minX, -slideW), minY: Math.max(bounds.minY, -slideH),
    maxX: Math.min(bounds.maxX, 2 * slideW), maxY: Math.min(bounds.maxY, 2 * slideH),
  }
  if (!(bounds.maxX > bounds.minX && bounds.maxY > bounds.minY)) return { refused: 'empty' }

  // Integer device-pixel grid: compositing at whole pixels never resamples.
  let scale = ctx.scale
  const grid = (s: number) => {
    const x0 = Math.floor(bounds!.minX * s), y0 = Math.floor(bounds!.minY * s)
    return { x0, y0, w: Math.max(1, Math.ceil(bounds!.maxX * s) - x0), h: Math.max(1, Math.ceil(bounds!.maxY * s) - y0) }
  }
  let px = grid(scale)
  const limit = Math.max(1, Math.floor(ctx.maxTextureSize))
  if (px.w > limit || px.h > limit) {
    const requested = { w: px.w, h: px.h }
    for (let i = 0; i < 8 && (px.w > limit || px.h > limit); i++) {
      scale *= (limit - 1) / Math.max(px.w, px.h)
      px = grid(scale)
    }
    ctx.sink?.add({
      code: 'layer-scale', slideIndex: slide.index, effect: tag,
      message: `Layer ${tag} (${requested.w}×${requested.h} px) exceeds the maximum texture size ${limit}; rasterised at reduced scale ${scale.toFixed(3)} and scaled up when composited`,
    })
    if (px.w > limit || px.h > limit) return refuse()
  }

  const bytes = px.w * px.h * 4
  const lease = ctx.workspace.lease(bytes, tag)
  if (!lease) return refuse()
  let canvas: HTMLCanvasElement | OffscreenCanvas
  try { canvas = ctx.canvas(px.w, px.h) }
  catch (error) { lease.release(); throw error }
  const c2d = canvas.getContext('2d') as unknown as Ctx2D
  c2d.setTransform(scale, 0, 0, scale, -px.x0, -px.y0)

  if (layer.kind === 'watermark') {
    paintWatermark(c2d, { widthPx: slideW, heightPx: slideH, background: slide.background }, ctx.watermark)
  } else {
    const options = sceneOptions(slide, { images: ctx.images, resolveFont: ctx.resolveFont, contentAssets: ctx.contentAssets })
    if (layer.kind === 'static' && layer.background) {
      c2d.fillStyle = slide.background ?? '#ffffff'
      c2d.fillRect(0, 0, slideW, slideH)
      // The plan put the underlay at the head of this run's nodes.
    }
    paintScene(layer.nodes, c2d, options)
  }
  return { canvas, bounds: { x: px.x0 / scale, y: px.y0 / scale, w: px.w / scale, h: px.h / scale }, scale, bytes, lease }
}
