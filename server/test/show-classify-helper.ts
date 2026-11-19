/**
 * Test helpers for the show-plan eligibility classifier (piece 4 Task 5):
 * classify a fixture deck through the real server backend, then apply the
 * classification (drops left out, masks multiplied into the rasters) and
 * compare the reconstructed composite against the original at a history.
 */
import { canvasBackend } from '../../src/core/canvas-backend'
import { decodeImageAsset } from '../../src/core/images'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import type { MediaClockRecord } from '../../src/slideshow/engine/media-ends'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../../src/slideshow/engine/steps'
import { prepareSlide, rasterPreparedLayer, type PreparedSlide } from '../../src/slideshow/prepare'
import type { PlanLayer } from '../../src/slideshow/render/plan'
import { createRng } from '../../src/slideshow/rng'
import type { CanvasFactory, LayerBitmap, Lease, SlideShowWorkspace } from '../../src/slideshow/types'
import { HIDDEN_INK, loadShowDeck } from '../../src/testdata/show-decks'
import { classifiedRasters, createSlideComposer, drawsAt, slideStateAt } from '../src/show/compose'
import { classifySlide, type Classification } from '../src/show/eligibility'

function roomyWorkspace(): SlideShowWorkspace {
  return {
    limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined,
    lease: (bytes, tag) => {
      const lease: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => lease, release() {} }
      return lease
    },
  }
}
export const factory = (): CanvasFactory => (w, h) => canvasBackend().createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never

export interface Run { ps: PreparedSlide; c: Classification; rasters: Map<PlanLayer, LayerBitmap>; timeline: SlideTimeline; scale: number; ms: number }

/** Prepares, rasterises (decoding the deck's images when `images`) and classifies slide 0 of `bytes`. */
export async function classify(bytes: Uint8Array, opts: { scale?: number; strictMarks?: boolean; images?: boolean; media?: { records: readonly MediaClockRecord[] } } = {}): Promise<Run> {
  const scale = opts.scale ?? 0.25
  const d = await loadShowDeck(bytes)
  const meta = await d.meta(0)
  const ps = prepareSlide({ index: 0, slide: d.doc.slides[0], meta, watermark: false })
  const backend = canvasBackend()
  const images = opts.images
    ? await Promise.all(d.doc.images.map(a => decodeImageAsset(a as never, (b, mime) => backend.decodeImage(b, mime) as Promise<CanvasImageSource>)))
    : []
  const sink = createDiagnosticSink()
  const rasters = new Map<PlanLayer, LayerBitmap>()
  for (const layer of ps.layers) {
    const r = rasterPreparedLayer(ps, layer, { scale, images, resolveFont: f => f, canvas: factory(), workspace: roomyWorkspace(), maxTextureSize: 8192, sink })
    if (!('refused' in r)) rasters.set(layer, r)
  }
  const timeline = createSlideTimeline(meta.timing, { showAnimation: true, reducedMotion: false })
  const t0 = performance.now()
  const c = classifySlide({ ps, timeline, scale, rasters, strictMarks: opts.strictMarks ?? false, triggers: timeline.triggers(), ...(opts.media ? { media: opts.media } : {}) })
  return { ps, c, rasters, timeline, scale, ms: performance.now() - t0 }
}

const ink = [parseInt(HIDDEN_INK.slice(0, 2), 16), parseInt(HIDDEN_INK.slice(2, 4), 16), parseInt(HIDDEN_INK.slice(4, 6), 16)]
/** Pixels within ΔRGB ≤ 6 of the hidden ink, with α > 0. */
export function inkPixels(px: Uint8ClampedArray): number {
  let n = 0
  for (let i = 0; i < px.length; i += 4) {
    if (px[i + 3] > 0 && Math.abs(px[i] - ink[0]) <= 6 && Math.abs(px[i + 1] - ink[1]) <= 6 && Math.abs(px[i + 2] - ink[2]) <= 6) n++
  }
  return n
}

export interface Reconstruction { diff: number; inkOriginal: number; inkClassified: number }

/**
 * The slide at `cursor`, time `t`, composed from the original rasters and from the classified ones (every
 * dropped layer left out, every mask applied): how many device pixels differ, and the hidden ink in each.
 */
export function reconstruct(run: Run, cursor: SlideCursor, t: number): Reconstruction {
  if (run.c.fallback) throw new Error(`slide falls back to its ${run.c.fallback}: no layer reconstructs it`)
  const draws = drawsAt(run.ps, slideStateAt(run.ps, run.timeline, cursor, t, createRng(1), createDiagnosticSink()))
  const a = createSlideComposer(run.ps, run.rasters, run.scale)
  const b = createSlideComposer(run.ps, classifiedRasters(run.rasters, run.c.drop, run.c.masks), run.scale)
  try {
    const pa = a.compose(draws), pb = b.compose(draws)
    let diff = 0
    for (let i = 0; i < pa.length; i += 4) if (pa[i] !== pb[i] || pa[i + 1] !== pb[i + 1] || pa[i + 2] !== pb[i + 2] || pa[i + 3] !== pb[i + 3]) diff++
    return { diff, inkOriginal: inkPixels(pa), inkClassified: inkPixels(pb) }
  } finally {
    a.dispose()
    b.dispose()
  }
}
