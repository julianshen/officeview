/**
 * The show-plan eligibility classifier (piece 4 spec §6, plan D3, D8, D10).
 *
 * Before any asset of a slide is sent, the server proves that no layer it
 * sends carries ink the composed show never displays. A step it cannot prove
 * plays `composited`, as server-rendered keyframes. Whenever a check is in
 * doubt, the answer is `composited`, never `layers`.
 *
 * - **Rule 1, the observable envelope (D10).** The slide is composed at every
 *   step boundary, at the sampled times of the pre-click interval, and at the
 *   sampled times of each step that may play from layers (`sampleTimes`).
 *   For each layer, each sample is composed with and without it. Every device
 *   pixel the layer changes is mapped back through the layer's draw matrix
 *   into layer-local px. The envelope is then dilated by 1 px. A tinted draw
 *   (a colour part under a non-identity matrix) does not observe the layer's
 *   raw ink (controller ruling 3), so it adds nothing to the envelope.
 * - **Rules 2 and 3.** A layer none of whose inked pixels is observed is
 *   dropped (`drop`). A layer some of whose inked pixels are never observed
 *   gets a layer-local alpha mask (`masks`: 255 keeps a pixel, 0 removes it).
 *   Transparent pixels are kept, since they carry no ink.
 * - **Rule 4, the proof subset (`proveItems`).** A step is `layers` only if
 *   all of these hold:
 *   - none of its draws is tinted;
 *   - none of its draws is masked;
 *   - no mover's swept hull (`sweptHull`) crosses an inked pixel that rule 3
 *     removed from a layer below it.
 *   Visibility and opacity changes pass.
 * - **Then D8 and `strictMarks`.**
 *   - `strictMarks` makes every moving step `composited` (`strict-marks`).
 *   - A mover that starts a step away from its first rest plays that step
 *     `composited` (`mark-rest`).
 *   - A step with an effect still running and never ending at its end is
 *     `composited` (`emphasis`).
 *   - So is a step during which the compositor degraded an effect.
 * - **Rule 5, the budget.** Compositions are counted per slide. Before each
 *   one, the count is checked against `CLASSIFY_BUDGET`. Once the budget is
 *   spent, every step not yet classified is `composited` (`budget`, with a
 *   `show-budget` diagnostic).
 * - **Rule 6, frames (D3).** Each composited step reserves its end frame
 *   first. The rest of the slide's 512 is then shared in step order
 *   (`allocateFrames`).
 * - **Rule 6a, triggers.** A trigger plays from layers only if, at every one
 *   of its samples, it changes nothing but visibility and opacity. Its end
 *   state is then part of the envelope. Otherwise it is `cut`, with a
 *   `trigger-cut` diagnostic.
 * - **Limits (D3).**
 *   - Over 64 steps, the slide plays as its still: every step is `limits`,
 *     every layer is dropped, and every trigger is cut (`show-limits`).
 *   - Over 128 layers, every step is `composited` (`limits`).
 *
 * Steps are proven in order, against the envelope gathered so far. A later
 * sample only adds observed pixels, so the removed set a proof sees always
 * contains the final one: an early proof is conservative.
 */
import { invert, stateMatrix } from '../../../src/slideshow/affine'
import type { SchedulerSnapshot } from '../../../src/slideshow/engine/scheduler'
import type { SlideCursor, SlideTimeline } from '../../../src/slideshow/engine/steps'
import type { FontResolver } from '../../../src/core/fonts/register'
import type { PreparedSlide } from '../../../src/slideshow/prepare'
import type { PlanLayer } from '../../../src/slideshow/render/plan'
import { createRng } from '../../../src/slideshow/rng'
import { targetKey } from '../../../src/slideshow/target'
import type { Affine, BehaviourNode, DiagnosticSink, TimeNode, LayerBitmap, RenderState, TargetState } from '../../../src/slideshow/types'
import { SHOW_PLAN_LIMITS, type Rect } from '../../../src/stream/show-plan'
import { canvasBackend } from '../../../src/core/canvas-backend'
import { createSlideComposer, deviceMatrix, drawsAt, slideStateAt, type DrawnLayer } from './compose'

export const SAMPLES_PER_INTERVAL = 8, CLASSIFY_BUDGET = 4096, TRACK_FPS = 15, TRACK_MAX_FRAMES = 90

export interface ClassifyInput {
  ps: PreparedSlide
  timeline: SlideTimeline
  /** Device px per slide px: the shipped bucket the rasters were drawn at. */
  scale: number
  rasters: Map<PlanLayer, LayerBitmap>
  strictMarks: boolean
  triggers: readonly string[]
  /** Additive to the brief: the font resolver the rasters used (it measures a paragraph range's mask box). Default: identity. */
  resolveFont?: FontResolver
}
export interface StepClass { eligibility: 'layers' | 'composited'; reason?: 'colour' | 'mask' | 'emphasis' | 'budget' | 'swept' | 'strict-marks' | 'mark-rest' | 'limits'; frames?: number }
export interface Classification {
  drop: Set<PlanLayer>
  masks: Map<PlanLayer, Uint8Array /* layer-local alpha mask */>
  steps: StepClass[]
  triggers: Map<string, 'layers' | 'cut'>
  compositions: number
  diagnostics: Array<{ code: string; step?: number }>
}

type Reason = NonNullable<StepClass['reason']>

// ---------------------------------------------------------------------------
// Intervals and samples
// ---------------------------------------------------------------------------

interface Interval { cursor: SlideCursor; t0: number; dur: number }

/** The slide time click steps start at: after every finite effect of the pre-click interval. */
function startTime(timeline: SlideTimeline): number {
  const b = timeline.busyUntil(timeline.initial())
  return Number.isFinite(b) && b > 0 ? b : 0
}

/** Step `step` (0-based) entered from its settled checkpoint, and how long its finite effects run. */
function stepInterval(timeline: SlideTimeline, step: number): Interval {
  const t0 = startTime(timeline)
  const cursor = timeline.next(timeline.checkpoint(step), t0)
  const end = timeline.busyUntil(cursor)
  return { cursor, t0, dur: Number.isFinite(end) && end > t0 ? end - t0 : 0 }
}

/** `SAMPLES_PER_INTERVAL` times evenly spaced over `[0, dur]` (both ends included), plus `extra`, sorted and distinct. */
function samplesOver(dur: number, extra: readonly number[]): number[] {
  const out = new Set<number>()
  for (let i = 0; i < SAMPLES_PER_INTERVAL; i++) out.add(dur > 0 ? (i * dur) / (SAMPLES_PER_INTERVAL - 1) : 0)
  for (const t of extra) if (t >= 0 && t <= dur) out.add(t)
  return [...out].sort((a, b) => a - b)
}

/**
 * The sample times of step `step` (0-based), in ms from the step's click:
 * - 8 evenly spaced times over the step, both ends included;
 * - each timing keyframe time and each behaviour begin and end time within the step (§6 rule 1, controller ruling 1).
 */
export function sampleTimes(timeline: SlideTimeline, step: number): number[] {
  const { cursor, t0, dur } = stepInterval(timeline, step)
  return samplesOver(dur, timeline.keyTimes(cursor, t0, t0 + dur).map(t => t - t0))
}

// ---------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------

const corners = (m: Affine, b: Rect): Array<[number, number]> =>
  [[b.x, b.y], [b.x + b.w, b.y], [b.x, b.y + b.h], [b.x + b.w, b.y + b.h]].map(([x, y]) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]])

/**
 * The area a mover sweeps, in slide px: the axis-aligned box over its bounds at every sample (a superset of
 * their convex hull), widened on every side by the largest displacement of any corner between consecutive
 * samples. One rect.
 */
export function sweptHull(states: readonly TargetState[], bounds: Rect): Rect[] {
  if (states.length === 0) return []
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity, step = 0
  let prev: Array<[number, number]> | undefined
  for (const s of states) {
    const c = corners(stateMatrix(s), bounds)
    for (const [x, y] of c) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y) }
    if (prev) for (let i = 0; i < 4; i++) step = Math.max(step, Math.hypot(c[i][0] - prev[i][0], c[i][1] - prev[i][1]))
    prev = c
  }
  if (![x0, y0, x1, y1, step].every(Number.isFinite)) return [{ x: -1e9, y: -1e9, w: 2e9, h: 2e9 }]
  return [{ x: x0 - step, y: y0 - step, w: x1 - x0 + 2 * step, h: y1 - y0 + 2 * step }]
}

const ID_COLOR = [1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0]
function isIdentityColor(m: readonly number[] | undefined): boolean {
  return !m || (m.length === 20 && m.every((v, i) => Math.abs(v - ID_COLOR[i]) < 1e-9))
}
/** Whether a draw shows anything other than its raster's own colours. */
function tinted(s: TargetState): boolean {
  const c = s.color
  return !!c && !(isIdentityColor(c.fill) && isIdentityColor(c.line) && isIdentityColor(c.text))
}

const drawn = (s: TargetState | undefined): s is TargetState => !!s && s.visible && s.opacity > 0

/** Two draw matrices differ: any linear difference, or a translation over `tol` slide px. */
function differs(a: Affine, b: Affine, tol: number): boolean {
  for (let i = 0; i < 4; i++) if (Math.abs(a[i] - b[i]) > 1e-6) return true
  return Math.abs(a[4] - b[4]) > tol || Math.abs(a[5] - b[5]) > tol
}

// ---------------------------------------------------------------------------
// Rule 4: the proof subset
// ---------------------------------------------------------------------------

/** One layer of a proof, in z order (bottom-most first). */
export interface ProofItem {
  /** The bitmap's bounds, slide px. */
  bounds: Rect
  /** The bitmap's size, device px. */
  size: { w: number; h: number }
  /** Rule 3's layer-local alpha mask (0 = an inked pixel removed); absent when nothing is removed. */
  mask?: Uint8Array
  /** The draw state at each sample of the interval; `undefined` where it is not drawn. */
  states: ReadonlyArray<TargetState | undefined>
  /**
   * The timing moves this item during the interval, whether or not its samples differ (a motion that leaves and
   * returns between samples). Its swept hull is then taken over its sampled bounds: anything it uncovers between
   * samples, it covered at a sample.
   */
  moves?: boolean
}

export type Proof = { ok: true; movers: number[] } | { ok: false; reason: 'colour' | 'mask' | 'swept' }

/**
 * §6 rule 4 over one interval's items (z order), at `scale` device px per slide px.
 * - A tinted draw fails with `colour`, and a masked draw with `mask`.
 * - A mover fails with `swept` when its swept hull covers a pixel that rule 3 removed from an item below it,
 *   taken at any of that item's drawn states.
 * - Visibility and opacity changes pass.
 * - On success, `movers` lists the indices of the items that move.
 */
export function proveItems(items: readonly ProofItem[], scale: number): Proof {
  for (const it of items) {
    for (const s of it.states) {
      if (!drawn(s)) continue
      if (tinted(s)) return { ok: false, reason: 'colour' }
      if (s.mask) return { ok: false, reason: 'mask' }
    }
  }
  const movers: number[] = []
  for (let i = 0; i < items.length; i++) {
    const states = items[i].states.filter(drawn)
    if (states.length === 0) continue
    const m0 = stateMatrix(states[0])
    if (!items[i].moves && !states.some(s => differs(stateMatrix(s), m0, 1e-6))) continue
    movers.push(i)
    const [hull] = sweptHull(states, items[i].bounds)
    for (let j = 0; j < i; j++) if (removedUnder(items[j], hull, scale)) return { ok: false, reason: 'swept' }
  }
  return { ok: true, movers }
}

/** Whether a pixel rule 3 removed from `item` (at any of its drawn states) lies inside `hull` (slide px). */
function removedUnder(item: ProofItem, hull: Rect, scale: number): boolean {
  const mask = item.mask
  if (!mask || !mask.includes(0)) return false
  const { w, h } = item.size
  const b = item.bounds
  const seen: Affine[] = []
  for (const s of item.states) {
    if (!drawn(s)) continue
    const m = deviceMatrix(scale, s)
    if (seen.some(x => !differs(x, m, 1e-6))) continue
    seen.push(m)
    const inv = invert(m)
    if (!inv) return true
    // The hull and the item's device box, intersected, in device px.
    const box = corners(m, b)
    const dx0 = Math.max(Math.floor(hull.x * scale), Math.floor(Math.min(...box.map(p => p[0]))))
    const dy0 = Math.max(Math.floor(hull.y * scale), Math.floor(Math.min(...box.map(p => p[1]))))
    const dx1 = Math.min(Math.ceil((hull.x + hull.w) * scale), Math.ceil(Math.max(...box.map(p => p[0]))))
    const dy1 = Math.min(Math.ceil((hull.y + hull.h) * scale), Math.ceil(Math.max(...box.map(p => p[1]))))
    for (let y = dy0; y < dy1; y++) {
      for (let x = dx0; x < dx1; x++) {
        const sx = inv[0] * (x + 0.5) + inv[2] * (y + 0.5) + inv[4], sy = inv[1] * (x + 0.5) + inv[3] * (y + 0.5) + inv[5]
        const u = Math.floor(((sx - b.x) * w) / b.w), v = Math.floor(((sy - b.y) * h) / b.h)
        if (u >= 0 && v >= 0 && u < w && v < h && mask[v * w + u] === 0) return true
      }
    }
  }
  return false
}

// ---------------------------------------------------------------------------
// Rule 6: frames (D3)
// ---------------------------------------------------------------------------

/**
 * D3 frame allocation over a slide's steps, in step order. Every composited step reserves its end frame. The
 * rest of `SHOW_PLAN_LIMITS.slideKeyframes` (512) is then shared at `min(90, ceil(durMs · 15 / 1000))` frames
 * per step. `truncated` marks a step that got fewer frames than its duration needs, by the 90-frame cap or by
 * the slide pool. A non-composited step gets 0.
 */
export function allocateFrames(steps: ReadonlyArray<{ composited: boolean; durMs: number }>): Array<{ frames: number; truncated: boolean }> {
  let rest = SHOW_PLAN_LIMITS.slideKeyframes - steps.filter(s => s.composited).length
  return steps.map(s => {
    if (!s.composited) return { frames: 0, truncated: false }
    const need = Math.max(1, Math.ceil((Math.max(0, Number.isFinite(s.durMs) ? s.durMs : 0) * TRACK_FPS) / 1000))
    const want = Math.min(TRACK_MAX_FRAMES, need)
    const extra = Math.max(0, Math.min(want - 1, rest))
    rest -= extra
    return { frames: 1 + extra, truncated: 1 + extra < need }
  })
}

// ---------------------------------------------------------------------------
// The classifier
// ---------------------------------------------------------------------------

interface LayerInfo {
  layer: PlanLayer
  bmp: LayerBitmap
  w: number
  h: number
  /** The raster's alpha (inked where > 0). */
  alpha: Uint8Array
  /** Observed layer-local px (before dilation). */
  seen: Uint8Array
}

const NOOP_SINK: DiagnosticSink = { add() {}, list: () => [] }

class OverBudget extends Error { constructor() { super('classify budget') } }
class Refused extends Error { constructor() { super('classify raster refused') } }

export function classifySlide(input: ClassifyInput): Classification {
  const { ps, timeline, scale, rasters, strictMarks } = input
  const n = timeline.steps
  const out: Classification = { drop: new Set(), masks: new Map(), steps: [], triggers: new Map(), compositions: 0, diagnostics: [] }
  const shipped = ps.layers.filter(l => l.kind !== 'watermark')
  const cutTriggers = () => {
    for (const spid of input.triggers) out.triggers.set(spid, 'cut')
    if (input.triggers.length) out.diagnostics.push({ code: 'trigger-cut' })
  }

  // D3 limits: over 64 steps the slide is its still; over 128 layers every step is composited.
  if (n > SHOW_PLAN_LIMITS.steps) {
    out.steps = Array.from({ length: n }, () => ({ eligibility: 'composited', reason: 'limits' }))
    for (const l of shipped) out.drop.add(l)
    out.diagnostics.push({ code: 'show-limits' })
    cutTriggers()
    return out
  }
  if (shipped.length > SHOW_PLAN_LIMITS.layers) {
    const frames = allocateFrames(Array.from({ length: n }, (_, k) => ({ composited: true, durMs: stepInterval(timeline, k).dur })))
    out.steps = frames.map(f => ({ eligibility: 'composited', reason: 'limits', frames: f.frames }))
    for (const l of shipped) out.drop.add(l)
    out.steps.forEach((_, k) => out.diagnostics.push({ code: 'step-composited', step: k }))
    frames.forEach((f, k) => { if (f.truncated) out.diagnostics.push({ code: 'step-truncated', step: k }) })
    cutTriggers()
    return out
  }

  const infos = new Map<PlanLayer, LayerInfo>()
  for (const layer of shipped) {
    const bmp = rasters.get(layer)
    if (!bmp) continue
    const w = bmp.canvas.width, h = bmp.canvas.height
    const data = (bmp.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, w, h).data
    const alpha = new Uint8Array(w * h)
    for (let i = 0; i < alpha.length; i++) alpha[i] = data[i * 4 + 3]
    infos.set(layer, { layer, bmp, w, h, alpha, seen: new Uint8Array(w * h) })
  }
  const composer = createSlideComposer(ps, rasters, scale)
  const W = composer.width, H = composer.height
  const rng = createRng(0x5eed + ps.index)
  const measure = canvasBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
  const drawOpts = { measure, ...(input.resolveFont ? { resolveFont: input.resolveFont } : {}) }
  const stateAt = (cursor: SlideCursor, t: number) => slideStateAt(ps, timeline, cursor, t, rng, NOOP_SINK)
  const draws = (state: RenderState) => drawsAt(ps, state, drawOpts)
  const byId = behavioursById(ps)

  const composeCounted = (list: readonly DrawnLayer[], without?: PlanLayer) => {
    if (out.compositions >= CLASSIFY_BUDGET) throw new OverBudget()
    out.compositions++
    return composer.compose(list, without)
  }

  /** Rule 1 at one state: every untinted drawn layer's changed pixels, mapped into layer-local px. */
  const observe = (list: readonly DrawnLayer[]) => {
    let full: Uint8ClampedArray | undefined
    for (const d of list) {
      const info = infos.get(d.layer)
      if (!info || !drawn(d.state) || tinted(d.state)) continue
      const m = deviceMatrix(scale, d.state)
      const inv = invert(m)
      const box = corners(m, info.bmp.bounds)
      const x0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[0])))), y0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[1]))))
      const x1 = Math.min(W, Math.ceil(Math.max(...box.map(p => p[0])))), y1 = Math.min(H, Math.ceil(Math.max(...box.map(p => p[1]))))
      if (!inv || !(x1 > x0 && y1 > y0)) continue
      full ??= composeCounted(list)
      const without = composeCounted(list, d.layer)
      const b = info.bmp.bounds, kx = info.w / b.w, ky = info.h / b.h
      const local = (x: number, y: number): [number, number] => [
        ((inv[0] * x + inv[2] * y + inv[4]) - b.x) * kx, ((inv[1] * x + inv[3] * y + inv[5]) - b.y) * ky,
      ]
      for (let y = y0; y < y1; y++) {
        for (let x = x0; x < x1; x++) {
          const i = (y * W + x) * 4
          if (full[i] === without[i] && full[i + 1] === without[i + 1] && full[i + 2] === without[i + 2] && full[i + 3] === without[i + 3]) continue
          // The device pixel's square, mapped into the bitmap: every bitmap px it touches is observed.
          const c = [local(x, y), local(x + 1, y), local(x, y + 1), local(x + 1, y + 1)]
          const u0 = Math.max(0, Math.floor(Math.min(...c.map(p => p[0])))), v0 = Math.max(0, Math.floor(Math.min(...c.map(p => p[1]))))
          const u1 = Math.min(info.w, Math.ceil(Math.max(...c.map(p => p[0])))), v1 = Math.min(info.h, Math.ceil(Math.max(...c.map(p => p[1]))))
          for (let v = v0; v < v1; v++) info.seen.fill(1, v * info.w + u0, v * info.w + Math.max(u0, u1))
        }
      }
    }
  }

  /** Rule 3's mask of one layer from the envelope so far (dilated by 1 px); `undefined` when no inked pixel is removed. */
  const maskOf = (info: LayerInfo): Uint8Array | undefined => {
    const { w, h, seen, alpha } = info
    let removed = false
    const mask = new Uint8Array(w * h).fill(255)
    for (let v = 0; v < h; v++) {
      for (let u = 0; u < w; u++) {
        const i = v * w + u
        if (alpha[i] === 0 || seen[i]) continue
        let near = false
        for (let dv = -1; dv <= 1 && !near; dv++) {
          for (let du = -1; du <= 1; du++) {
            const uu = u + du, vv = v + dv
            if (uu >= 0 && vv >= 0 && uu < w && vv < h && seen[vv * w + uu]) { near = true; break }
          }
        }
        if (!near) { mask[i] = 0; removed = true }
      }
    }
    return removed ? mask : undefined
  }
  const observedInk = (info: LayerInfo): boolean => {
    for (let i = 0; i < info.seen.length; i++) if (info.seen[i] && info.alpha[i] > 0) return true
    return false
  }

  const checkpointEnvelope = () => new Map([...infos.values()].map(i => [i, i.seen.slice()]))
  const restoreEnvelope = (saved: Map<LayerInfo, Uint8Array>) => { for (const [i, seen] of saved) i.seen.set(seen) }

  /** The draw state of each info layer at each of `lists` (undefined where not drawn). */
  const statesPerLayer = (lists: readonly (readonly DrawnLayer[])[]) => {
    const per = new Map<PlanLayer, Array<TargetState | undefined>>()
    for (const info of infos.values()) per.set(info.layer, lists.map(() => undefined))
    lists.forEach((list, k) => { for (const d of list) { const row = per.get(d.layer); if (row) row[k] = d.state } })
    return per
  }

  const steps: Array<StepClass | undefined> = Array.from({ length: n }, () => undefined)
  const durs: number[] = []
  let overBudget = false
  try {
    // A raster the compositor refuses cannot be composed, so nothing about the slide can be proven.
    if (composer.refused.length) throw new Refused()
    // Rule 1 at every step boundary, and over the pre-click interval.
    const boundaries = Array.from({ length: n + 1 }, (_, k) => draws(stateAt(timeline.checkpoint(k), 0)))
    for (const list of boundaries) observe(list)
    const pre = startTime(timeline)
    if (pre > 0) {
      const init = timeline.initial()
      for (const t of samplesOver(pre, timeline.keyTimes(init, 0, pre))) observe(draws(stateAt(init, t)))
    }
    // A degraded effect at a boundary or before the first click: the envelope itself is unproven.
    if (composer.takeDiagnostics().length) throw new Refused()
    // D8: each layer's first rest (its draw matrix at the first boundary it is visible at).
    const boundaryStates = statesPerLayer(boundaries)
    const firstRest = new Map<PlanLayer, Affine>()
    for (const [layer, row] of boundaryStates) { const s = row.find(drawn); if (s) firstRest.set(layer, stateMatrix(s)) }

    for (let k = 0; k < n; k++) {
      const iv = stepInterval(timeline, k)
      durs[k] = iv.dur
      const times = samplesOver(iv.dur, timeline.keyTimes(iv.cursor, iv.t0, iv.t0 + iv.dur).map(t => t - iv.t0))
      const lists = times.map(t => draws(stateAt(iv.cursor, iv.t0 + t)))
      const per = statesPerLayer(lists)
      const timed = timedEffects(ps, byId, timeline.stateAt(timeline.checkpoint(k), 0), times.map(t => timeline.stateAt(iv.cursor, iv.t0 + t)))
      const moved = timed.moved
      const items: ProofItem[] = [...infos.values()].map(info => ({
        bounds: info.bmp.bounds, size: { w: info.w, h: info.h }, states: per.get(info.layer)!, ...(moved.has(info.layer) ? { moves: true } : {}),
      }))
      // State-only rule 4 checks first: a tinted or masked step needs no compositions.
      let early: Reason | undefined = timed.colour ? 'colour' : timed.mask ? 'mask' : undefined
      for (const it of items) for (const s of it.states) if (drawn(s) && !early) early = tinted(s) ? 'colour' : s.mask ? 'mask' : undefined
      if (!early) {
        const still = timeline.stateAt(iv.cursor, iv.t0 + iv.dur + 1).active.some(e => e.node.kind !== 'audio' && e.node.kind !== 'video')
        if (still) early = 'emphasis'
      }
      if (early) { steps[k] = { eligibility: 'composited', reason: early }; continue }
      const saved = checkpointEnvelope()
      for (const list of lists) observe(list)
      const degraded = composer.takeDiagnostics()
      if (degraded.length) {
        // What a degraded composition showed is not what the client shows: none of it is observed.
        restoreEnvelope(saved)
        steps[k] = { eligibility: 'composited', reason: degraded.some(c => c.startsWith('mask')) ? 'mask' : 'emphasis' }
        continue
      }
      // Rule 4, with each layer's mask from the envelope so far.
      const proof = proveItems([...infos.values()].map((info, i) => {
        const mask = maskOf(info)
        return { ...items[i], ...(mask ? { mask } : {}) }
      }), scale)
      if (!proof.ok) { steps[k] = { eligibility: 'composited', reason: proof.reason }; continue }
      if (proof.movers.length && strictMarks) { steps[k] = { eligibility: 'composited', reason: 'strict-marks' }; continue }
      // D8: a mover starting this step away from its first rest (over 0.5 device px, or any rotation or scale).
      const layers = [...infos.keys()]
      const away = proof.movers.some(i => {
        const start = boundaryStates.get(layers[i])![k]
        const first = firstRest.get(layers[i])
        return drawn(start) && !!first && differs(stateMatrix(start), first, 0.5 / scale)
      })
      if (away) { steps[k] = { eligibility: 'composited', reason: 'mark-rest' }; out.diagnostics.push({ code: 'mark-rest', step: k }); continue }
      steps[k] = { eligibility: 'layers' }
    }

    // Rule 6a: triggers, from the slide's first checkpoint.
    const base = timeline.checkpoint(0)
    const baseList = boundaries[0]
    const t0 = startTime(timeline)
    for (const spid of input.triggers) {
      const cur = timeline.trigger(base, spid, t0)
      if (!cur) { out.triggers.set(spid, 'cut'); continue }
      const keys = timeline.keyTimes(cur, t0, Number.POSITIVE_INFINITY)
      const dur = keys.length ? keys[keys.length - 1] - t0 : 0
      const keyed = samplesOver(dur, keys.map(t => t - t0))
      const lists = keyed.map(t => draws(stateAt(cur, t0 + t)))
      const timed = timedEffects(ps, byId, timeline.stateAt(base, 0), keyed.map(t => timeline.stateAt(cur, t0 + t)))
      const ok = !timed.colour && !timed.mask && timed.moved.size === 0 && lists.every(list => toggleOnly(baseList, list))
      if (!ok) { out.triggers.set(spid, 'cut'); continue }
      const saved = checkpointEnvelope()
      observe(lists[lists.length - 1])
      if (composer.takeDiagnostics().length) { restoreEnvelope(saved); out.triggers.set(spid, 'cut'); continue }
      out.triggers.set(spid, 'layers')
    }
  } catch (e) {
    if (e instanceof OverBudget) overBudget = true
    else if (e instanceof Refused) {
      // Nothing about the slide can be proven: every step is composited and no layer ships.
      for (let k = 0; k < n; k++) steps[k] = { eligibility: 'composited', reason: 'limits' }
      for (const info of infos.values()) info.seen.fill(0)
    } else throw e
  } finally {
    composer.dispose()
  }

  if (overBudget) out.diagnostics.push({ code: 'show-budget' })
  for (let k = 0; k < n; k++) {
    steps[k] ??= { eligibility: 'composited', reason: 'budget' }
    durs[k] ??= stepInterval(timeline, k).dur
  }
  for (const spid of input.triggers) if (!out.triggers.has(spid)) out.triggers.set(spid, 'cut')
  if ([...out.triggers.values()].includes('cut')) out.diagnostics.push({ code: 'trigger-cut' })

  // Rule 6: frames for the composited steps.
  const frames = allocateFrames(steps.map((s, k) => ({ composited: s!.eligibility === 'composited', durMs: durs[k] })))
  out.steps = steps.map((s, k) => (s!.eligibility === 'composited' ? { ...s!, frames: frames[k].frames } : s!))
  out.steps.forEach((s, k) => {
    if (s.eligibility !== 'composited') return
    out.diagnostics.push({ code: 'step-composited', step: k })
    if (frames[k].truncated) out.diagnostics.push({ code: 'step-truncated', step: k })
  })

  // Rules 2 and 3 over the final envelope.
  for (const info of infos.values()) {
    if (!observedInk(info)) { out.drop.add(info.layer); continue }
    const mask = maskOf(info)
    if (mask) out.masks.set(info.layer, mask)
  }
  // A layer that never rasterised paints nothing: it ships nothing.
  for (const l of shipped) if (!infos.has(l)) out.drop.add(l)
  return out
}

/** Attributes that only show, hide or fade a target. */
const TOGGLES = new Set(['style.visibility', 'style.opacity'])
/** Attributes that recolour a part (`behaviours.ts` colour channels). */
const COLOURS = new Set(['fillcolor', 'stroke.color', 'style.color'])

type SnapshotLike = SchedulerSnapshot

/** What the timing does in an interval, whatever its samples show (§6: "sampling alone is never the proof"). */
interface TimedEffects {
  /** Layers moved by motion, rotation, scale or a geometry attribute (sub-layers of a moved group included). */
  moved: Set<PlanLayer>
  /** A recolour (`animClr`, or a colour attribute) runs. */
  colour: boolean
  /** A filter other than a plain fade runs (a mask). */
  mask: boolean
}

/** Every behaviour node of the slide's timing tree (the scheduled one, text builds fanned out), by `cTn` id. */
function behavioursById(ps: PreparedSlide): Map<number, BehaviourNode> {
  const out = new Map<number, BehaviourNode>()
  const timing = ps.meta.timing
  const root = ps.textTiming?.tree.root ?? (timing && !('refused' in timing) ? timing.root : undefined)
  const walk = (n: TimeNode, depth: number) => {
    if (depth > 256) return
    if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
    out.set(n.ctn.id, n)
  }
  if (root) walk(root, 0)
  return out
}

/**
 * The behaviours that run in an interval: those active at any of its samples, plus those frozen or ended at
 * its end that were not at its start (so a behaviour whose begin was never sampled still counts), classified
 * by kind and attribute.
 */
function timedEffects(ps: PreparedSlide, byId: ReadonlyMap<number, BehaviourNode>, start: SnapshotLike, snaps: readonly SnapshotLike[]): TimedEffects {
  const nodes = new Set<BehaviourNode>()
  for (const snap of snaps) for (const { node } of snap.active) nodes.add(node)
  const last = snaps[snaps.length - 1]
  if (last) {
    const before = new Set([...start.frozen, ...start.active].map(e => e.node))
    for (const { node } of last.frozen) if (!before.has(node)) nodes.add(node)
    for (const id of last.ended) if (!start.ended.has(id)) { const n = byId.get(id); if (n) nodes.add(n) }
  }
  const keys = new Set<string>()
  let colour = false, mask = false
  for (const node of nodes) {
    if (node.kind === 'cmd' || node.kind === 'audio' || node.kind === 'video') continue
    if (node.kind === 'animClr' || ((node.kind === 'anim' || node.kind === 'set') && node.attrNames.some(a => COLOURS.has(a)))) { colour = true; continue }
    if (node.kind === 'animEffect') { if ((node.filter ?? 'fade').trim() !== 'fade') mask = true; continue }
    const moves = node.kind === 'animMotion' || node.kind === 'animRot' || node.kind === 'animScale' ||
      node.attrNames.some(a => !TOGGLES.has(a))
    if (moves && 'spid' in node.target) keys.add(targetKey(node.target))
  }
  const moved = new Set<PlanLayer>()
  const hit = (key: string) => [...keys].some(k => key === k || key.startsWith(`${k}/`))
  const walk = (layers: readonly PlanLayer[], inMoved: boolean) => {
    for (const l of layers) {
      const m = inMoved || (l.kind === 'target' && hit(l.key))
      if (m) moved.add(l)
      if (l.kind === 'target' && l.sub) walk(l.sub, m)
    }
  }
  walk(ps.plan.layers, false)
  return { moved, colour, mask }
}

/** A trigger's draws differ from the base only in which layers are drawn and their opacity (§6 rule 6a). */
function toggleOnly(base: readonly DrawnLayer[], list: readonly DrawnLayer[]): boolean {
  const before = new Map(base.map(d => [d.layer, d.state]))
  for (const d of list) {
    const s = d.state
    if (!drawn(s)) continue
    if (tinted(s) || s.mask) return false
    const b = before.get(d.layer)
    if (b && differs(stateMatrix(b), stateMatrix(s), 1e-6)) return false
    if (!b && differs(stateMatrix(s), [1, 0, 0, 1, 0, 0], 1e-6)) return false
  }
  return true
}
