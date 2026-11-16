/**
 * The show-plan eligibility classifier (piece 4 spec §6, plan D3, D8, D10; fix round 1 rulings).
 *
 * Before any asset of a slide is sent, the server proves that no layer it sends carries ink the composed show
 * never displays, and that the layers it sends reproduce every state they play. What it cannot prove plays
 * `composited` (keyframes), or the whole slide plays its piece-3 still (`fallback: 'still'`). Never `layers`
 * on doubt.
 *
 * 1. **Plan the states (no composition).** The observed states are:
 *    - every step boundary;
 *    - the autoplay interval before the first click;
 *    - every trigger subset, applied in a fixed order from every checkpoint (at most `MAX_TRIGGERS`
 *      triggers, more plays the still);
 *    - the samples (`sampleTimes`) of each click step that may play from layers.
 *    Checks that need no composition run first, by sample and by timing (`timedEffects`):
 *    - a tint or mask gives `colour` or `mask`;
 *    - an effect still running at the end gives `emphasis`;
 *    - a mover with `strictMarks` gives `strict-marks`;
 *    - a mover away from its first rest gives `mark-rest` (D8).
 *    Unproven autoplay plays the still.
 * 2. **Budget (rule 5).** Each state costs a with/without pair per observable layer, plus a reconstruction
 *    pair. The fixed states must fit `CLASSIFY_BUDGET`, or the slide plays the still. Click steps are then
 *    admitted in order while they fit; the rest are `budget`.
 * 3. **Envelopes (rules 1–3), top to bottom.** Each layer's with/without pairs are composed with every layer
 *    above it already dropped or masked, so two layers that each hide the other's pixels cannot both lose
 *    them. Changed device px map into layer-local px, dilated by 1 px. Tinted draws observe nothing. A layer
 *    with no observed inked pixel is dropped; otherwise its unobserved inked pixels are masked out.
 * 4. **Rule 4 (`proveItems`)** runs over every interval with the final masks; a dropped layer counts as fully
 *    removed. A mover's swept hull must not cross removed pixels below it.
 * 5. **Reconstruction.** Every observed state, composed from the classified layers, must equal the original:
 *    - a mismatch in a step's sample composites that step;
 *    - a mismatch at a boundary composites both steps meeting there (on a slide without triggers);
 *    - any other mismatch plays the still.
 * 6. **Frames (D3, `allocateFrames`) and triggers (6a).** A trigger plays from layers only if it is
 *    visibility and opacity only, by sample and by timing, in every combination.
 *    - Fix round 2: a slide with any trigger plays from layers only if nothing is dropped or masked and
 *      every click step is `layers`. Otherwise it plays the still.
 *    - The trigger combinations above still feed the envelope: they are reachable histories, and they can
 *      show a layer's full ink.
 */
import { invert, stateMatrix } from '../../../src/slideshow/affine'
import type { SchedulerSnapshot } from '../../../src/slideshow/engine/scheduler'
import type { SlideCursor, SlideTimeline } from '../../../src/slideshow/engine/steps'
import type { FontResolver } from '../../../src/core/fonts/register'
import type { PreparedSlide } from '../../../src/slideshow/prepare'
import type { PlanLayer } from '../../../src/slideshow/render/plan'
import { createRng } from '../../../src/slideshow/rng'
import { targetKey } from '../../../src/slideshow/target'
import type { Affine, BehaviourNode, Condition, ContainerNode, DiagnosticSink, TimeNode, LayerBitmap, RenderState, TargetState } from '../../../src/slideshow/types'
import { SHOW_PLAN_LIMITS, type Rect } from '../../../src/stream/show-plan'
import { showCanvasBackend } from './canvas'
import { createSlideComposer, deviceMatrix, drawsAt, maskedRaster, slideStateAt, type DrawnLayer } from './compose'

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
export interface StepClass { eligibility: 'layers' | 'composited'; reason?: 'colour' | 'mask' | 'emphasis' | 'budget' | 'swept' | 'strict-marks' | 'mark-rest' | 'limits' | 'media'; frames?: number }
export interface Classification {
  drop: Set<PlanLayer>
  masks: Map<PlanLayer, Uint8Array /* layer-local alpha mask */>
  steps: StepClass[]
  triggers: Map<string, 'layers' | 'cut'>
  compositions: number
  diagnostics: Array<{ code: string; step?: number }>
  /**
   * Fix round 1 (controller ruling): the slide cannot be proven (unproven autoplay, too many or too costly
   * trigger combinations, over 64 steps, a reconstruction mismatch outside a click step). It plays as its
   * piece-3 still: no layer ships, so `drop` and `masks` are empty and every trigger is cut.
   */
  fallback?: 'still'
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

/** Additive (Task 6): each click step's duration in ms (its finite effects), for D3 frame allocation and tracks. */
export function stepDurations(timeline: SlideTimeline): number[] {
  return Array.from({ length: timeline.steps }, (_, k) => stepInterval(timeline, k).dur)
}

/** Additive (Task 6): step `step` (0-based) entered from its settled checkpoint: its cursor, start time and duration. */
export function stepWindow(timeline: SlideTimeline, step: number): { cursor: SlideCursor; t0: number; dur: number } {
  return stepInterval(timeline, step)
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
const EMPTY_SNAPSHOT: SchedulerSnapshot = { active: [], frozen: [], ended: new Set() }

/** The most triggers whose combinations are classified (controller ruling, fix round 1): more plays the still. */
export const MAX_TRIGGERS = 4

class OverBudget extends Error { constructor() { super('classify budget') } }
/** The slide cannot be proven: it plays its still (`Classification.fallback`). */
class Still extends Error { constructor(readonly reason: Reason) { super(`classify: still (${reason})`) } }

/** Who displays an observed state: a boundary, the autoplay interval, a click step's sample, or a trigger combination. */
type Owner = { kind: 'boundary'; at: number } | { kind: 'auto' } | { kind: 'step'; step: number } | { kind: 'trigger'; spids: readonly string[] }
interface Obs { list: readonly DrawnLayer[]; owner: Owner }

/** One interval (the autoplay one or a click step) planned before any composition. */
interface Planned {
  lists: DrawnLayer[][]
  items: ProofItem[]
  /** Fails rule 4 or 5 without a composition: tinted, masked, or still running at its end. */
  early?: Reason
  /** Passes only if rule 4 does, then is composited all the same (D8, strictMarks). */
  post?: 'strict-marks' | 'mark-rest'
}

export function classifySlide(input: ClassifyInput): Classification {
  const { ps, timeline, scale, rasters, strictMarks } = input
  const n = timeline.steps
  const out: Classification = { drop: new Set(), masks: new Map(), steps: [], triggers: new Map(), compositions: 0, diagnostics: [] }
  const shipped = ps.layers.filter(l => l.kind !== 'watermark')
  const durOf = (k: number) => stepInterval(timeline, k).dur
  const finish = (steps: StepClass[]) => {
    const frames = allocateFrames(steps.map((s, k) => ({ composited: s.eligibility === 'composited', durMs: durOf(k) })))
    out.steps = steps.map((s, k) => (s.eligibility === 'composited' ? { ...s, frames: frames[k].frames } : s))
    out.steps.forEach((s, k) => {
      if (s.eligibility !== 'composited') return
      out.diagnostics.push({ code: 'step-composited', step: k })
      if (frames[k].truncated) out.diagnostics.push({ code: 'step-truncated', step: k })
    })
    for (const spid of input.triggers) if (!out.triggers.has(spid)) out.triggers.set(spid, 'cut')
    if ([...out.triggers.values()].includes('cut')) out.diagnostics.push({ code: 'trigger-cut' })
    return out
  }
  /** Fix round 1: the slide plays its piece-3 still. Nothing ships, every step is composited, every trigger cut. */
  const still = (reason: Reason) => {
    out.fallback = 'still'
    out.drop.clear()
    out.masks.clear()
    out.triggers.clear()
    out.diagnostics.push({ code: 'step-composited' })
    if (reason === 'budget') out.diagnostics.push({ code: 'show-budget' })
    out.steps = Array.from({ length: n }, () => ({ eligibility: 'composited', reason }))
    for (const spid of input.triggers) out.triggers.set(spid, 'cut')
    if (input.triggers.length) out.diagnostics.push({ code: 'trigger-cut' })
    return out
  }

  // D3 limits: over 64 steps the slide is its still; over 128 layers every step is composited.
  if (n > SHOW_PLAN_LIMITS.steps) {
    out.diagnostics.push({ code: 'show-limits' })
    return still('limits')
  }
  if (shipped.length > SHOW_PLAN_LIMITS.layers) {
    for (const l of shipped) out.drop.add(l)
    return finish(Array.from({ length: n }, () => ({ eligibility: 'composited', reason: 'limits' })))
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
  const order = [...infos.values()]
  const composer = createSlideComposer(ps, rasters, scale)
  const W = composer.width, H = composer.height
  const rng = createRng(0x5eed + ps.index)
  const measure = showCanvasBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
  const drawOpts = { measure, ...(input.resolveFont ? { resolveFont: input.resolveFont } : {}) }
  const stateAt = (cursor: SlideCursor, t: number) => slideStateAt(ps, timeline, cursor, t, rng, NOOP_SINK)
  const draws = (state: RenderState) => drawsAt(ps, state, drawOpts)
  const byId = behavioursById(ps)

  const composeCounted = (list: readonly DrawnLayer[], without: PlanLayer | undefined, classified: boolean) => {
    if (out.compositions >= CLASSIFY_BUDGET) throw new OverBudget()
    out.compositions++
    return composer.compose(list, without, classified)
  }
  /** Layers of `list` whose raw ink a draw would observe (drawn, untinted, with a raster). */
  const observable = (list: readonly DrawnLayer[]) => list.filter(d => infos.has(d.layer) && drawn(d.state) && !tinted(d.state))
  /** Compositions one observed state costs: a with/without pair per observable layer, then the reconstruction pair. */
  const cost = (list: readonly DrawnLayer[]) => 2 * observable(list).length + 2

  /**
   * Rule 1 for one layer at one state (fix round 1, ruling 3): with and without the layer, every layer above it
   * already classified (dropped or masked), every layer below it raw. Changed device px are mapped into
   * layer-local px.
   */
  const observeLayer = (info: LayerInfo, list: readonly DrawnLayer[]) => {
    const d = list.find(x => x.layer === info.layer)
    if (!d || !drawn(d.state) || tinted(d.state)) return
    const m = deviceMatrix(scale, d.state)
    const inv = invert(m)
    const box = corners(m, info.bmp.bounds)
    const x0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[0])))), y0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[1]))))
    const x1 = Math.min(W, Math.ceil(Math.max(...box.map(p => p[0])))), y1 = Math.min(H, Math.ceil(Math.max(...box.map(p => p[1]))))
    if (!inv || !(x1 > x0 && y1 > y0)) return
    const full = composeCounted(list, undefined, true)
    const without = composeCounted(list, info.layer, true)
    const b = info.bmp.bounds, kx = info.w / b.w, ky = info.h / b.h
    const [i0, i1, i2, i3, i4, i5] = inv
    // Task 6 (cost, same arithmetic): the device pixel square's four corners mapped into the bitmap as scalars
    // (no per-pixel arrays): every bitmap px the square touches is observed.
    const lu = (x: number, y: number) => ((i0 * x + i2 * y + i4) - b.x) * kx
    const lv = (x: number, y: number) => ((i1 * x + i3 * y + i5) - b.y) * ky
    for (let y = y0; y < y1; y++) {
      for (let x = x0; x < x1; x++) {
        const i = (y * W + x) * 4
        if (full[i] === without[i] && full[i + 1] === without[i + 1] && full[i + 2] === without[i + 2] && full[i + 3] === without[i + 3]) continue
        const ua = lu(x, y), ub = lu(x + 1, y), uc = lu(x, y + 1), ud = lu(x + 1, y + 1)
        const va = lv(x, y), vb = lv(x + 1, y), vc = lv(x, y + 1), vd = lv(x + 1, y + 1)
        const u0 = Math.max(0, Math.floor(Math.min(ua, ub, uc, ud))), v0 = Math.max(0, Math.floor(Math.min(va, vb, vc, vd)))
        const u1 = Math.min(info.w, Math.ceil(Math.max(ua, ub, uc, ud))), v1 = Math.min(info.h, Math.ceil(Math.max(va, vb, vc, vd)))
        for (let v = v0; v < v1; v++) info.seen.fill(1, v * info.w + u0, v * info.w + Math.max(u0, u1))
      }
    }
  }

  /** Rule 3's mask of one layer from its envelope (dilated by 1 px); `undefined` when no inked pixel is removed. */
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

  /** The draw state of each info layer at each of `lists` (undefined where not drawn). */
  const statesPerLayer = (lists: readonly (readonly DrawnLayer[])[]) => {
    const per = new Map<PlanLayer, Array<TargetState | undefined>>()
    for (const info of order) per.set(info.layer, lists.map(() => undefined))
    lists.forEach((list, k) => { for (const d of list) { const row = per.get(d.layer); if (row) row[k] = d.state } })
    return per
  }

  try {
    // A raster the compositor refuses cannot be composed, so nothing about the slide can be proven.
    if (composer.refused.length) throw new Still('limits')
    // PR #78 round 12: a streamed show plays no media, so the client ends every audio/video node on the first frame
    // it is active (`syntheticEnds` in the shared driver), at a frame-dependent time and only on a live cursor.
    // Every history built here has no media end. Round 12c: when a media end can reach drawing (a structural
    // dependency analysis, `mediaRefused`), what that end starts or cuts short is unproven: the still.
    if (mediaRefused(ps)) throw new Still('media')

    // ---- Plan every state, before any composition -------------------------------------------------------
    const boundaries = Array.from({ length: n + 1 }, (_, k) => draws(stateAt(timeline.checkpoint(k), 0)))
    const boundaryStates = statesPerLayer(boundaries)
    // D8: each layer's first rest (its draw matrix at the first boundary it is visible at).
    const firstRest = new Map<PlanLayer, Affine>()
    for (const [layer, row] of boundaryStates) { const s = row.find(drawn); if (s) firstRest.set(layer, stateMatrix(s)) }
    // D8, strengthened (Task 6 follow-up, controller ruling): one marked asset carries one rest. A layer drawn at a
    // boundary away from its first rest has its mark off the page there, so every step from the last boundary it
    // was drawn at its first rest onward plays composited (`mark-rest`): no boundary shows it away from its mark.
    let restCut = Number.POSITIVE_INFINITY
    for (const [layer, row] of boundaryStates) {
      const first = firstRest.get(layer)
      if (!first) continue
      let lastAtRest = -1
      for (let j = 0; j < row.length; j++) {
        const s = row[j]
        if (!drawn(s)) continue
        if (!differs(stateMatrix(s), first, 0.5 / scale)) { lastAtRest = j; continue }
        restCut = Math.min(restCut, Math.max(0, lastAtRest))
        break
      }
    }

    /** One interval's lists, proof items and the checks that need no composition. */
    const plan = (cursor: SlideCursor, t0: number, dur: number, start: SchedulerSnapshot, restAt: number | undefined): Planned => {
      const times = samplesOver(dur, timeline.keyTimes(cursor, t0, t0 + dur).map(t => t - t0))
      const lists = times.map(t => draws(stateAt(cursor, t0 + t)))
      const per = statesPerLayer(lists)
      const timed = timedEffects(ps, byId, start, times.map(t => timeline.stateAt(cursor, t0 + t)))
      const items: ProofItem[] = order.map(info => ({
        bounds: info.bmp.bounds, size: { w: info.w, h: info.h }, states: per.get(info.layer)!, ...(timed.moved.has(info.layer) ? { moves: true } : {}),
      }))
      let early: Reason | undefined = timed.colour ? 'colour' : timed.mask ? 'mask' : undefined
      for (const it of items) for (const s of it.states) if (drawn(s) && !early) early = tinted(s) ? 'colour' : s.mask ? 'mask' : undefined
      if (!early && timeline.stateAt(cursor, t0 + dur + 1).active.some(e => e.node.kind !== 'audio' && e.node.kind !== 'video')) early = 'emphasis'
      if (early) return { lists, items, early }
      const moving = items.map((it, i) => {
        const ds = it.states.filter(drawn)
        return it.moves || (ds.length > 1 && ds.some(s => differs(stateMatrix(s), stateMatrix(ds[0]), 1e-6))) ? i : -1
      }).filter(i => i >= 0)
      if (moving.length && strictMarks) return { lists, items, post: 'strict-marks' }
      // D8: a mover starting the interval away from its first rest (over 0.5 device px, or any rotation or scale).
      const away = moving.some(i => {
        const layer = order[i].layer
        const startState = restAt === undefined ? per.get(layer)![0] : boundaryStates.get(layer)![restAt]
        const first = firstRest.get(layer)
        return drawn(startState) && !!first && differs(stateMatrix(startState), first, 0.5 / scale)
      })
      return away ? { lists, items, post: 'mark-rest' } : { lists, items }
    }

    // Autoplay (fix round 1, ruling 1): the interval before the first click is proven like a step, or the slide is its still.
    const init = timeline.initial()
    const pre = startTime(timeline)
    const init0 = timeline.stateAt(init, 0)
    const autoKeys = timeline.keyTimes(init, 0, Number.POSITIVE_INFINITY)
    const auto = pre > 0 || autoKeys.length > 0 || init0.active.length > 0 || init0.frozen.length > 0
      ? plan(init, 0, pre, EMPTY_SNAPSHOT, undefined) : undefined
    if (auto?.early) throw new Still(auto.early)
    if (auto?.post) throw new Still(auto.post)

    // Click steps.
    const plans = Array.from({ length: n }, (_, k) => {
      const iv = stepInterval(timeline, k)
      const p = plan(iv.cursor, iv.t0, iv.dur, timeline.stateAt(timeline.checkpoint(k), 0), k)
      return k >= restCut && !p.early && !p.post ? { ...p, post: 'mark-rest' as const } : p
    })

    // Triggers (fix round 1, ruling 2): every subset, in a fixed order, from every checkpoint; over 4, the still.
    const triggers = input.triggers
    if (triggers.length > MAX_TRIGGERS) throw new Still('limits')
    const t0 = startTime(timeline)
    const triggerOk = new Map<string, boolean>()
    for (const spid of triggers) {
      // Its own effect, from the first checkpoint: visibility and opacity only, at every sample and in its timing.
      const base = timeline.checkpoint(0)
      const cur = timeline.trigger(base, spid, t0)
      if (!cur) { triggerOk.set(spid, false); continue }
      const keys = timeline.keyTimes(cur, t0, Number.POSITIVE_INFINITY)
      const dur = keys.length ? keys[keys.length - 1] - t0 : 0
      const keyed = samplesOver(dur, keys.map(t => t - t0))
      const timed = timedEffects(ps, byId, timeline.stateAt(base, 0), keyed.map(t => timeline.stateAt(cur, t0 + t)))
      triggerOk.set(spid, !timed.colour && !timed.mask && timed.moved.size === 0 && keyed.every(t => toggleOnly(boundaries[0], draws(stateAt(cur, t0 + t)))))
    }
    const combos: Obs[] = []
    for (let set = 1; set < 1 << triggers.length; set++) {
      const spids = triggers.filter((_, i) => set & (1 << i))
      for (let k = 0; k <= n; k++) {
        let cur: SlideCursor | undefined = timeline.checkpoint(k)
        let t = t0
        for (const spid of spids) {
          cur = cur && timeline.trigger(cur, spid, t)
          if (!cur) break
          const keys = timeline.keyTimes(cur, t, Number.POSITIVE_INFINITY)
          t = keys.length ? Math.max(t, keys[keys.length - 1]) : t
        }
        if (!cur) { for (const s of spids) triggerOk.set(s, false); continue }
        const list = draws(stateAt(cur, t))
        if (!toggleOnly(boundaries[k], list)) for (const s of spids) triggerOk.set(s, false)
        combos.push({ list, owner: { kind: 'trigger', spids } })
      }
    }

    // ---- Budget (rule 5): the fixed states first, then click steps in order while they fit ---------------
    const obs: Obs[] = [
      ...boundaries.map((list, k) => ({ list, owner: { kind: 'boundary', at: k } as Owner })),
      ...(auto ? auto.lists.map(list => ({ list, owner: { kind: 'auto' } as Owner })) : []),
      ...combos,
    ]
    let room = CLASSIFY_BUDGET - obs.reduce((a, o) => a + cost(o.list), 0)
    if (room < 0) throw new Still('budget')
    const budgetCut = new Set<number>()
    for (let k = 0; k < n; k++) {
      const p = plans[k]
      if (p.early || p.post) continue
      const c = budgetCut.size ? Infinity : p.lists.reduce((a, l) => a + cost(l), 0)
      if (c > room) { budgetCut.add(k); continue }
      room -= c
      for (const list of p.lists) obs.push({ list, owner: { kind: 'step', step: k } })
    }

    // ---- Rules 1–3: envelopes top to bottom, each against the layers above already classified -------------
    for (const info of [...order].reverse()) {
      for (const o of obs) observeLayer(info, o.list)
      if (!observedInk(info)) { out.drop.add(info.layer); composer.classify(info.layer, null); continue }
      const mask = maskOf(info)
      if (mask) { out.masks.set(info.layer, mask); composer.classify(info.layer, maskedRaster(info.bmp, mask)) }
    }
    // A degraded effect anywhere: what was composed is not what the client shows.
    if (composer.takeDiagnostics().length) throw new Still('mask')

    // ---- Rule 4 with the final masks (a dropped layer has every inked pixel removed) ----------------------
    const proofMask = (info: LayerInfo) => out.drop.has(info.layer) ? info.alpha.map(a => (a > 0 ? 0 : 255)) : out.masks.get(info.layer)
    const prove = (p: Planned) => proveItems(p.items.map((it, i) => { const mask = proofMask(order[i]); return mask ? { ...it, mask } : it }), scale)
    if (auto) { const proof = prove(auto); if (!proof.ok) throw new Still(proof.reason) }
    const steps: StepClass[] = plans.map((p, k) => {
      if (p.early) return { eligibility: 'composited', reason: p.early }
      if (budgetCut.has(k)) return { eligibility: 'composited', reason: 'budget' }
      const proof = prove(p)
      if (!proof.ok) return { eligibility: 'composited', reason: proof.reason }
      if (p.post) {
        if (p.post === 'mark-rest') out.diagnostics.push({ code: 'mark-rest', step: k })
        return { eligibility: 'composited', reason: p.post }
      }
      return { eligibility: 'layers' }
    })

    // ---- Reconstruction (ruling 3): every observed state, classified, equals the original -----------------
    for (const o of obs) {
      if (o.owner.kind === 'step' && steps[o.owner.step].eligibility === 'composited') continue
      const a = composeCounted(o.list, undefined, false)
      const b = composeCounted(o.list, undefined, true)
      let same = a.length === b.length
      for (let i = 0; same && i < a.length; i++) if (a[i] !== b[i]) same = false
      if (same) continue
      // The classified layers cannot show this state. A click step's sample: that step is composited. A
      // boundary: both steps meeting there are (their keyframe tracks show it), and no trigger can fire from
      // layers. Autoplay or a trigger combination: the still.
      const composite = (k: number) => {
        if (k >= 0 && k < n && steps[k].eligibility === 'layers') steps[k] = { eligibility: 'composited', reason: o.list.some(d => drawn(d.state) && tinted(d.state)) ? 'colour' : 'mask' }
      }
      if (o.owner.kind === 'step') composite(o.owner.step)
      else if (o.owner.kind === 'boundary' && triggers.length === 0) { composite(o.owner.at - 1); composite(o.owner.at) }
      else throw new Still('mask')
    }
    if (composer.takeDiagnostics().length) throw new Still('mask')

    // Rule 6a (fix round 2, controller ruling): trigger histories interleave with the main sequence and with each
    // other without bound, so no finite sample proves them. A slide with triggers plays from layers only when
    // every shipped layer is its full raster (no drop, no mask: any history then composes from complete layers
    // exactly as locally) and every click step plays from layers (a keyframe track cannot carry trigger effects).
    if (triggers.length && (out.drop.size > 0 || out.masks.size > 0 || steps.some(st => st.eligibility !== 'layers'))) throw new Still('mask')
    for (const spid of triggers) out.triggers.set(spid, triggerOk.get(spid) ? 'layers' : 'cut')
    if (budgetCut.size) out.diagnostics.push({ code: 'show-budget' })
    return finish(steps)
  } catch (e) {
    if (e instanceof Still) return still(e.reason)
    if (e instanceof OverBudget) return still('budget')
    throw e
  } finally {
    composer.dispose()
  }
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

/**
 * PR #78 round 12: the slide's scheduled timing holds an audio or video node. A streamed show plays no media, so the
 * client ends each one on its first active frame (`syntheticEnds`), which no server history carries: such a slide
 * plays its still, and no Morph pair starts or ends on it. A superset of the driver's own walk (`mediaNodes`, by kind,
 * bounded at depth 64): a deeper node the client never ends is refused here too.
 */
export function holdsMedia(ps: PreparedSlide): boolean {
  const timing = ps.meta.timing
  // By kind, not by id (a duplicate `cTn` id must not hide a media node), over the tree the client walks (the
  // streamed timing) and the scheduled one (text builds fanned out).
  const roots = [timing && !('refused' in timing) ? timing.root : undefined, ps.textTiming?.tree.root]
  const media = (n: TimeNode, depth: number): boolean => {
    if (depth > 256) return false
    if ('children' in n) return n.children.some(c => media(c, depth + 1))
    return n.kind === 'audio' || n.kind === 'video'
  }
  return roots.some(r => !!r && media(r, 0))
}

/**
 * PR #78 round 12c: whether a media node's end can change a drawn behaviour's activation, duration or visibility,
 * by a conservative structural dependency analysis of `root` (no timing is sampled: the client's synthetic end
 * lands at a frame-dependent time, and end-condition semantics are not monotonic in it).
 *
 * Per node, three flags, grown to a fixpoint:
 * - `B`, its begin may depend on a media end: its parent's does; a start condition names a node whose begin or
 *   end depends (whatever the event: a node that begins again begins after its end); in a `seq`, it is chained to its previous sibling (a
 *   finite delay, or no condition) whose end depends; its parent repeats or reverses and the parent's end depends;
 *   in an `excl`, a sibling's begin or end depends.
 * - `C`, it may be cut short at such a time: it is media (the synthetic end); an end condition or its `endSync`
 *   names a dependent node; its parent is cut short; its parent's `endSync` ends it with a dependent child; an `excl` sibling depends.
 * - `E`, its end may depend: `B` or `C`; or it is a container whose end follows its children (no explicit `dur`,
 *   or an `endSync`) and a child's end depends; or its `endSync` names a dependent node.
 *
 * An implicit end is the latest child end, so it never cuts a child short; an explicit `dur` blocks propagation
 * from below. The media end then reaches drawing when a drawn behaviour (not media, not `cmd`) has `B` or `C`, or a
 * container with drawn descendants has `E` and either lets go of them at its end (any `fill` but `hold` or
 * `freeze`), can begin again (an event start condition, or `restart` other than `never`), or holds a non-media
 * descendant whose begin is not a fixed offset (an event, a `tn` or an indefinite delay: it could begin after the
 * client's end of the container while the server still has it active).
 *
 * Fail closed (true): a reference to a node not in the tree, a cycle among the references, a tree deeper than 256
 * or over 20,000 nodes, and (round 12d) any timing on a dependent node outside the allowlist (`timingModelled`).
 */
export function mediaEndAffectsDrawing(root: TimeNode): boolean {
  const nodes: TimeNode[] = []
  const parent = new Map<TimeNode, ContainerNode | undefined>()
  const byId = new Map<number, TimeNode[]>()
  let over = false
  const walk = (n: TimeNode, p: ContainerNode | undefined, depth: number) => {
    if (over) return
    if (depth > 256 || nodes.length >= 20_000) { over = true; return }
    nodes.push(n)
    parent.set(n, p)
    const same = byId.get(n.ctn.id)
    if (same) same.push(n); else byId.set(n.ctn.id, [n])
    if ('children' in n) for (const c of n.children) walk(c, n, depth + 1)
  }
  walk(root, undefined, 0)
  if (over) return true
  const conds = (n: TimeNode): Condition[] => [...n.ctn.start, ...n.ctn.end, ...(n.ctn.endSync ? [n.ctn.endSync] : [])]
  // Unresolved references, then cycles among them (a node waits on another's begin or end).
  for (const n of nodes) for (const c of conds(n)) if (c.tn !== undefined && !byId.has(c.tn)) return true
  const state = new Map<TimeNode, 1 | 2>()
  const cyclic = (n: TimeNode): boolean => {
    const st = state.get(n)
    if (st === 1) return true
    if (st === 2) return false
    state.set(n, 1)
    for (const c of conds(n)) if (c.tn !== undefined) for (const m of byId.get(c.tn)!) if (cyclic(m)) return true
    state.set(n, 2)
    return false
  }
  for (const n of nodes) if (cyclic(n)) return true

  const B = new Set<TimeNode>(), C = new Set<TimeNode>(), E = new Set<TimeNode>()
  const isMedia = (n: TimeNode) => n.kind === 'audio' || n.kind === 'video'
  // Any reference, whatever its event: a node that begins again has later begins that follow its end.
  const refDep = (c: Condition) => c.tn !== undefined && byId.get(c.tn)!.some(m => B.has(m) || E.has(m))
  const repeats = (n: TimeNode) => n.ctn.repeatCount !== undefined || n.ctn.repeatDurMs !== undefined || n.ctn.autoRev
  const chained = (n: TimeNode) => n.ctn.start.length === 0 || n.ctn.start.some(c => c.evt === undefined && c.tn === undefined && c.delayMs !== 'indefinite')
  const add = (set: Set<TimeNode>, n: TimeNode, yes: boolean) => { if (yes && !set.has(n)) { set.add(n); return true } return false }
  for (let changed = true; changed;) {
    changed = false
    for (const n of nodes) {
      const p = parent.get(n)
      const sibs = p ? p.children : []
      const k = sibs.indexOf(n)
      const prev = p?.kind === 'seq' && k > 0 ? sibs[k - 1] : undefined
      const exclDep = p?.kind === 'excl' && sibs.some(o => o !== n && (B.has(o) || E.has(o)))
      const b = (!!p && B.has(p)) || n.ctn.start.some(refDep) || (!!prev && E.has(prev) && chained(n)) ||
        (!!p && repeats(p) && E.has(p)) || exclDep
      const c = isMedia(n) || n.ctn.end.some(refDep) || (!!n.ctn.endSync && refDep(n.ctn.endSync)) || (!!p && C.has(p)) ||
        (!!p && !!p.ctn.endSync && sibs.some(o => E.has(o))) || exclDep
      const follows = 'children' in n && (n.ctn.durMs === undefined || !!n.ctn.endSync) && n.children.some(o => E.has(o))
      const e = b || c || follows || (!!n.ctn.endSync && refDep(n.ctn.endSync))
      if (add(B, n, b)) changed = true
      if (add(C, n, c)) changed = true
      if (add(E, n, e)) changed = true
    }
  }
  // Round 12d: a media-dependent node carries only timing the model handles (`timingModelled`).
  for (const n of nodes) if ((B.has(n) || C.has(n) || E.has(n)) && !timingModelled(n)) return true
  const drawn = (n: TimeNode) => !('children' in n) && !isMedia(n) && n.kind !== 'cmd'
  const drawnBelow = (n: TimeNode): boolean => 'children' in n ? n.children.some(o => drawn(o) || drawnBelow(o)) : false
  // A begin that is not a fixed offset into its parent: it may come after the container's client-side end.
  const late = (n: TimeNode) => n.ctn.start.some(c => c.evt !== undefined || c.tn !== undefined || c.delayMs === 'indefinite')
  const lateBelow = (n: TimeNode): boolean => 'children' in n ? n.children.some(o => (!isMedia(o) && late(o)) || lateBelow(o)) : false
  for (const n of nodes) {
    if (drawn(n) && (B.has(n) || C.has(n))) return true
    if (!('children' in n) || !E.has(n) || !drawnBelow(n)) continue
    if (n.ctn.fill !== 'hold' && n.ctn.fill !== 'freeze') return true
    if (n.ctn.start.some(c => c.evt !== undefined) || (n.ctn.restart !== undefined && n.ctn.restart !== 'never')) return true
    if (lateBelow(n)) return true
  }
  return false
}

/**
 * PR #78 round 12d: the `p:cTn` timing attributes, from what the scheduler (`src/slideshow/engine/scheduler.ts`,
 * `steps.ts`, `behaviours.ts`) reads. On a media-dependent node (`B`, `C` or `E` in `mediaEndAffectsDrawing`) only
 * these are accepted; any other key, or a non-default value of a `TIMING_DEFAULT_ONLY` one, refuses the slide.
 *
 * - `TIMING_MODELLED`, any value: the analysis models them. `id`, `start`, `end`, `endSync` (the references and
 *   their propagation), `durMs` (an explicit one blocks propagation), `fill` and `restart` (the final checks),
 *   and the role and preset fields the scheduler matches on but that set no local time (`nodeType`, `presetClass`,
 *   `presetId`, `presetSubtype`, `grpId`).
 * - `TIMING_DEFAULT_ONLY`: they change a node's local time or activity in ways the analysis does not model, so
 *   they must hold their defaults. `accel`, `decel` (0: once a dependent duration turns finite, easing reshapes the
 *   children's progress, codex round 12d), `autoRev` (false), `repeatCount`, `repeatDurMs` and `iterate` (absent).
 * - `TIMING_IGNORED`: the reader carries them, and nothing in the engine acts on them: `syncBehavior`,
 *   `masterRel`, `afterEffect`.
 *
 * A dependent container must be a `par`, or a `seq` with the reader's default options (not concurrent, no
 * `nextAc`/`prevAc`): `excl` and the seek and concurrency rules of a dependent sequence are not modelled.
 */
export const TIMING_MODELLED: ReadonlySet<string> = new Set(['id', 'start', 'end', 'endSync', 'durMs', 'fill', 'restart', 'nodeType', 'presetClass', 'presetId', 'presetSubtype', 'grpId'])
export const TIMING_DEFAULT_ONLY: Readonly<Record<string, (v: unknown) => boolean>> = {
  accel: v => v === undefined || v === 0,
  decel: v => v === undefined || v === 0,
  autoRev: v => v === undefined || v === false,
  repeatCount: v => v === undefined,
  repeatDurMs: v => v === undefined,
  iterate: v => v === undefined,
}
export const TIMING_IGNORED: ReadonlySet<string> = new Set(['syncBehavior', 'masterRel', 'afterEffect'])

function timingModelled(n: TimeNode): boolean {
  for (const [k, v] of Object.entries(n.ctn)) {
    if (TIMING_MODELLED.has(k) || TIMING_IGNORED.has(k)) continue
    const ok = TIMING_DEFAULT_ONLY[k]
    if (!ok || !ok(v)) return false
  }
  if (n.kind === 'excl') return false
  if (n.kind === 'seq' && n.seq && (n.seq.concurrent || n.seq.nextAc !== 'none' || n.seq.prevAc !== 'none')) return false
  return true
}

/**
 * PR #78 round 12c: the slide plays its still for media (`classifySlide`, reason `media`; `buildMorphPlan`, the
 * captured-frame fade): it holds media whose end can reach drawing in the streamed timing or the scheduled one.
 */
export function mediaRefused(ps: PreparedSlide): boolean {
  if (!holdsMedia(ps)) return false
  const timing = ps.meta.timing
  const roots = [timing && !('refused' in timing) ? timing.root : undefined, ps.textTiming?.tree.root]
  return roots.some(r => !!r && mediaEndAffectsDrawing(r))
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
