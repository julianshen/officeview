/**
 * Keyframe tracks for `composited` steps (piece 4 spec §6 rule 6, plan D3, D4): full frames of the slide,
 * composed on the server from its original layers, at 15 fps, with the visible mark painted once (as a unit is).
 * Layer-1 is applied per session at send time, as for units. A track is fetched on demand (one charge, D4).
 *
 * A step of `count` frames plays frames 1 … count − 1 at `k / 15` s from its click, then its end frame (the
 * settled checkpoint after the step). One frame is the end frame alone; a step longer than its frames (the
 * 90-frame cap, or the slide's 512) plays its first frames and then cuts to the end (`step-truncated`).
 */
import { canvasBackend } from '../../../src/core/canvas-backend'
import { createDiagnosticSink } from '../../../src/slideshow/diagnostics'
import { createRng } from '../../../src/slideshow/rng'
import { encodePng } from '../mark/png'
import { buildShowPlan, markFrame, TRACK_ASSET_BASE, trackInputs, type ShowBuildOptions, type ShowHost, type ShowPlanJob } from './build'
import { createSlideComposer, drawsAt, slideStateAt } from './compose'
import { stepWindow, TRACK_FPS, TRACK_MAX_FRAMES } from './eligibility'

/** Frame `i` (0-based) of a `count`-frame track over a step of `durMs`: ms from the step's click; the last is the end. */
export function trackTimes(durMs: number, count: number): number[] {
  return Array.from({ length: count }, (_, i) => (i === count - 1 ? durMs : Math.min(durMs, ((i + 1) * 1000) / TRACK_FPS)))
}

/** Track frame `i` of step `step`: its asset id, unique within the slide and apart from the bundle's (entry 0, layers 1 …). */
export const trackAssetId = (step: number, i: number) => TRACK_ASSET_BASE + step * TRACK_MAX_FRAMES + i

/**
 * The keyframe track of step `job.step` (0-based), as the bundle's plan allocates it. A step the plan plays from
 * layers, or a slide the plan refuses, has no track (`[]`).
 */
export function renderTrack(doc: ShowHost, job: ShowPlanJob & { step: number }, opts: ShowBuildOptions = {}): Array<{ assetId: number; w: number; h: number; png: Uint8Array }> {
  const show = doc.show
  if (!show) throw new Error('show: not a deck')
  const built = buildShowPlan(doc, job, opts)
  if ('refused' in built) return []
  const step = built.plan.steps[job.step]
  if (!step || step.eligibility !== 'composited' || !step.track) return []
  const e = trackInputs(show, job)
  const { ps, timeline, rasters } = e
  const composer = createSlideComposer(ps, rasters, job.scale)
  try {
    const win = stepWindow(timeline, job.step)
    const rng = createRng(0x5eed + ps.index), sink = createDiagnosticSink()
    const draw = { measure: canvasBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D, resolveFont: (f: string) => f }
    const W = composer.width, H = composer.height
    const page = { w: built.plan.slide.w, h: built.plan.slide.h }
    return trackTimes(win.dur, step.track.count).map((t, i, all) => {
      const state = i === all.length - 1
        ? slideStateAt(ps, timeline, timeline.checkpoint(job.step + 1), 0, rng, sink)
        : slideStateAt(ps, timeline, win.cursor, win.t0 + t, rng, sink)
      const rgba = markFrame(new Uint8ClampedArray(composer.compose(drawsAt(ps, state, draw))), W, H, job.scale, page, job.mark)
      return { assetId: trackAssetId(job.step, i), w: W, h: H, png: encodePng(W, H, rgba) }
    })
  } finally {
    composer.dispose()
  }
}
