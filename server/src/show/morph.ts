/**
 * The Morph pair build (piece 4 spec §5, §6; plan D1, D7, D11): one `(from, to, fromStep)` Morph transition as
 * the client will play it, over the shared DOM-free preparation (`planMorph`, Task 2).
 *
 * 1. Both slides are prepared as the bundle build prepares them (`prepareShowSlide`). The outgoing slide is
 *    re-planned at `checkpoint(fromStep)` (D11: the cursor's step when the slide was left, settled), the incoming
 *    at its arrival state.
 * 2. `planMorph` runs with `MorphEnv.raster` supplied (Task 2 ruling): each ref is rasterised at the bucket, and a
 *    ref that paints nothing drops out. A string from `planMorph` is the only refusal: `morph-fallback`.
 * 3. **Only the refs an item side names ship** (D11): a from-side object hidden at `fromStep` is in `plan.refs`
 *    (it painted) but in no item, and gets no asset.
 * 4. **The proof subset** (§6 rules 1–5) over the Morph's samples (`SAMPLES_PER_INTERVAL` + 1 evenly spaced
 *    progress values, the ends included): each shipped ref's envelope (with and without it, every other ref raw)
 *    gives its rule-3 mask; `proveItems` (rule 4) runs over the items in z order with those masks; the composite
 *    from the masked rasters must equal the original at every sample (reconstruction); the compositions must fit
 *    `CLASSIFY_BUDGET` (rule 5). `strictMarks` with any mover, a rest that is not an axis-aligned box (D8), an asset
 *    over its cap or a record the codec refuses also fail. Any failure plays the whole Morph composited.
 *    The marks are then validated at progress 0 and 1 (§7, D7, D9: `validateMorph`), as the bundle's are at its
 *    step boundaries; a failure plays composited (`marks`).
 * 5. **Assets** are the masked rasters with the visible mark burned at their rest (`source-atop`, D7): a from side at
 *    progress 0, a to side at progress 1. Layer-1 is applied per session at send time with each asset's `rest`.
 * 6. **Composited** (D11, `morph-composited`): `min(90, ceil(durMs · 15 / 1000))` full frames of the Morph,
 *    composed by `morphFrame` on the server Canvas 2D compositor from the original rasters (displayed ink only)
 *    and marked once as a keyframe is. A longer Morph plays its first 6 s, then cuts (`step-truncated`).
 *
 * Every asset carries `unit = to` on the wire (Task 4). Asset ids are disjoint from the `to` slide's bundle and
 * tracks: `MORPH_ASSET_BASE + ref` for layers, `MORPH_FRAME_BASE + i` for frames.
 */
import { invert, stateMatrix } from '../../../src/slideshow/affine'
import { canvasBackend } from '../../../src/core/canvas-backend'
import { createDiagnosticSink } from '../../../src/slideshow/diagnostics'
import { morphFrame, type MorphFrameItem, type MorphFrameSpec } from '../../../src/slideshow/morph/render'
import { rasterPreparedLayer } from '../../../src/slideshow/prepare'
import { planMorph, type MorphLayerRef, type MorphPlan } from '../../../src/slideshow/prepare/morph'
import { createCanvas2DCompositor } from '../../../src/slideshow/render/canvas2d'
import { slideSize } from '../../../src/slideshow/render/scene'
import { createRng } from '../../../src/slideshow/rng'
import type { Affine, CanvasFactory, Compositor, LayerBitmap, LayerDraw, LayerHandle, Lease, SlideShowWorkspace, TargetState } from '../../../src/slideshow/types'
import { LAYER_LIMIT } from '../../../src/slideshow/workspace'
import { encodeShowMessage, PlanRejected, SHOW_PLAN_LIMITS, type MorphPlanRecord, type MorphSideRecord, type Rect } from '../../../src/stream/show-plan'
import { createHash } from 'node:crypto'
import { tauFor } from '../detect/tau'
import { basis } from '../mark/basis'
import { extendedCodeword } from '../mark/bch'
import { detect } from '../mark/detect'
import { markLayerRaster } from '../mark/mark-raster'
import { messageFor } from '../mark/message'
import { DEFAULT_PARAMS } from '../mark/params'
import { decodePng, encodePng } from '../mark/png'
import type { MarkSpec } from '../mark/visible'
import { markFrame, prepareShowSlide, SHOW_BUNDLE_BUDGET, VALIDATION_MESSAGE_KEY, type ShowHost } from './build'
import { slideStateAt } from './compose'
import { CLASSIFY_BUDGET, proveItems, SAMPLES_PER_INTERVAL, TRACK_FPS, TRACK_MAX_FRAMES, type ProofItem } from './eligibility'
import { burnVisibleMark, validateMarks } from './marks'
import { trackTimes } from './tracks'

export interface MorphJob { from: number; to: number; fromStep: number; scale: 1 | 2; motion: 'full' | 'reduced'; mark: MarkSpec; strictMarks?: boolean }
/** `rest` (additive to the brief): the Layer-1 crop, as `ShowAsset.rest` (page device px at the bucket; the frames rest at 0, 0). */
export interface MorphAsset { assetId: number; w: number; h: number; png: Uint8Array; rest: { x: number; y: number; scale: number } }
/**
 * Additive to the brief: `diagnostics`, D21 codes the wire carries as `diag` text (the Morph record has no
 * diagnostics field); `page`, the slide's device size at the bucket (the Layer-1 page the rests crop).
 */
export type BuiltMorph =
  | { record: MorphPlanRecord; assets: MorphAsset[]; diagnostics: string[]; page: { w: number; h: number } }
  | { record: MorphPlanRecord & { composited: { fps: 15; count: number } }; frames: MorphAsset[]; diagnostics: string[]; page: { w: number; h: number }; reason: MorphUnproven }
  | { refused: 'morph-fallback' }

/** The request names no Morph this deck plays (`to` not a Morph, `fromStep` past the from-slide's steps): the host answers `unit-range`. */
export class MorphInvalid extends Error { constructor(why: string) { super(`show-morph: ${why}`); this.name = 'MorphInvalid' } }

export const MORPH_ASSET_BASE = 0x20000
export const MORPH_FRAME_BASE = 0x30000
/** The most layers one Morph rasterises for itself (both slides), as local mode (`stage.ts` `MORPH_MAX_LAYERS`). */
const MORPH_MAX_LAYERS = 2 * LAYER_LIMIT
const MAX_UNIT_PX = 8192

const open = (bytes: number, tag: string): Lease => { const l: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => l, release() {} }; return l }
const OPEN_WORKSPACE: SlideShowWorkspace = { limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined, lease: open }
const factory: CanvasFactory = (w, h) => canvasBackend().createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never
const resolveFont = (f: string) => f

/** Why a Morph plays composited: rule 4 (`colour`, `mask`, `swept`), rule 5 (`budget`), the reconstruction, `strictMarks`, a rest that is not a box (D8), a cap, the §7 mark validation (`marks`). */
export type MorphUnproven = 'colour' | 'mask' | 'swept' | 'budget' | 'reconstruction' | 'strict-marks' | 'mark-rest' | 'limits' | 'marks'
type Unproven = MorphUnproven

export function buildMorphPlan(doc: ShowHost, job: MorphJob): BuiltMorph {
  const show = doc.show
  if (!show) throw new Error('show: not a deck')
  const from = prepareShowSlide(show, job.from, job.motion)
  const to = prepareShowSlide(show, job.to, job.motion)
  const spec = to.ps.meta.transition
  if (!spec || spec.element !== 'morph') throw new MorphInvalid('the slide does not enter by Morph')
  if (!Number.isInteger(job.fromStep) || job.fromStep < 0 || job.fromStep > from.timeline.steps) throw new MorphInvalid('fromStep is past the slide')
  const scale = job.scale
  const sink = createDiagnosticSink()
  const fromState = slideStateAt(from.ps, from.timeline, from.timeline.checkpoint(job.fromStep), 0, createRng(0x5eed + job.from), sink)
  const live = slideStateAt(to.ps, to.timeline, to.timeline.initial(), 0, createRng(0x5eed + job.to), sink)

  const bitmaps = new Map<MorphLayerRef, LayerBitmap>()
  const plan = planMorph({ ps: from.ps, state: fromState }, to.ps, live, spec, {
    sink, workspace: OPEN_WORKSPACE, canvas: factory, resolveFont, maxLayers: MORPH_MAX_LAYERS, scale, watermark: false,
    raster: (refs) => {
      bitmaps.clear()
      const painted = new Set<number>()
      for (let i = 0; i < refs.length; i++) {
        const r = refs[i]!
        const ps = r.side === 'from' ? from.ps : to.ps
        const out = rasterPreparedLayer(ps, r.layer, { scale: r.scale, images: show.images ?? [], resolveFont, canvas: factory, workspace: OPEN_WORKSPACE, maxTextureSize: MAX_UNIT_PX, sink })
        if ('refused' in out) {
          if (out.refused === 'empty') continue
          bitmaps.clear()
          return 'pressure'
        }
        bitmaps.set(r, { canvas: out.canvas, bounds: out.bounds, scale: out.scale, bytes: out.bytes })
        painted.add(i)
      }
      return painted
    },
  })
  if (typeof plan === 'string') return { refused: 'morph-fallback' }
  try {
    return assemble(plan, bitmaps, job, slideSize(to.ps.slide), spec.durationMs, show.docKey)
  } finally {
    for (const l of plan.leases) l.release()
  }
}

function assemble(plan: MorphPlan, bitmaps: ReadonlyMap<MorphLayerRef, LayerBitmap>, job: MorphJob, page: { w: number; h: number }, durationMs: number, docKey: string): BuiltMorph {
  const scale = job.scale
  const W = Math.max(1, Math.ceil(page.w * scale)), H = Math.max(1, Math.ceil(page.h * scale))
  const dur = Number.isFinite(durationMs) && durationMs > 0 ? durationMs : 0
  const bmpOf = (ref: number) => bitmaps.get(plan.refs[ref]!)
  // The refs an item side (or `over`) names: only these ship (D11).
  const used = new Set<number>()
  for (const it of plan.items) { if (it.from) used.add(it.from.ref); if (it.to) used.add(it.to.ref) }
  for (const r of plan.over) used.add(r)
  const handles = new Map<number, LayerHandle>()
  for (const r of used) handles.set(r, { id: r })
  const spec: MorphFrameSpec = {
    pairs: plan.items.map((item): MorphFrameItem => {
      const { from: f, to: t, ...rest } = item
      return { ...rest, ...(f ? { from: { ...f, handle: handles.get(f.ref)! } } : {}), ...(t ? { to: { ...t, handle: handles.get(t.ref)! } } : {}) }
    }),
    ...(plan.over.length ? { over: plan.over.map(r => ({ handle: handles.get(r)!, state: { visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] as Affine } })) } : {}),
  }
  const samples = Array.from({ length: SAMPLES_PER_INTERVAL + 1 }, (_, k) => k / SAMPLES_PER_INTERVAL)
  const lists = samples.map(p => drawsAtProgress(spec, p))

  const composited = (why: Unproven): BuiltMorph => compositeMorph(plan, bitmaps, used, spec, job, page, W, H, dur, why)
  // z order (bottom-first) as `morphFrame` paints: each item's from side, then its to side; then `over`.
  const order: number[] = []
  for (const it of plan.items) { if (it.from) order.push(it.from.ref); if (it.to) order.push(it.to.ref) }
  for (const r of plan.over) order.push(r)
  const proved = prove(order, bmpOf, used, lists, scale, W, H)
  if (typeof proved === 'string') return composited(proved)
  if (job.strictMarks && proved.movers > 0) return composited('strict-marks')

  // Rests (D8): a from side at progress 0, a to side at progress 1, each an axis-aligned box.
  const sideOf = new Map<number, 'from' | 'to'>()
  for (const it of plan.items) { if (it.from) sideOf.set(it.from.ref, 'from'); if (it.to) sideOf.set(it.to.ref, 'to') }
  const restOf = (ref: number): Rect | undefined => {
    const list = sideOf.get(ref) === 'to' ? lists[lists.length - 1]! : lists[0]!
    const d = flatten(list).find(x => x.handle.id === ref)
    const bmp = bmpOf(ref)!
    // No draw at its rest (a side not shown there): its mark has no page anchor (fail closed).
    if (!d) return undefined
    const m = stateMatrix(d.state)
    if (Math.abs(m[1]) > 1e-6 || Math.abs(m[2]) > 1e-6 || !(m[0] > 1e-6) || !(m[3] > 1e-6)) return undefined
    return { x: m[0] * bmp.bounds.x + m[4], y: m[3] * bmp.bounds.y + m[5], w: m[0] * bmp.bounds.w, h: m[3] * bmp.bounds.h }
  }
  const ids = [...used].sort((a, b) => a - b)
  let total = 0
  const assets: MorphAsset[] = []
  for (const ref of ids) {
    const rest = restOf(ref)
    if (!rest) return composited('mark-rest')
    const bmp = proved.rasters.get(ref)!
    const bytes = bmp.canvas.width * bmp.canvas.height * 4
    if (bytes > SHOW_PLAN_LIMITS.assetDecoded) return composited('limits')
    total += bytes
    if (total > SHOW_BUNDLE_BUDGET) return composited('limits')
    const marked = burnVisibleMark(bmp, rest, job.mark, page)
    const w = marked.canvas.width, h = marked.canvas.height
    const rgba = (marked.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, w, h).data
    assets.push({ assetId: MORPH_ASSET_BASE + ref, w, h, png: encodePng(w, h, rgba), rest: { x: Math.round(rest.x * scale), y: Math.round(rest.y * scale), scale: w / (rest.w * scale) } })
  }
  const side = (s: NonNullable<MorphPlan['items'][number]['from']>): MorphSideRecord => {
    const { ref, geometry, state, part, key } = s
    return { assetId: MORPH_ASSET_BASE + ref, geometry: { x: geometry.x, y: geometry.y, w: geometry.w, h: geometry.h, rot: geometry.rot }, ...(state ? { state } : {}), ...(part ? { part } : {}), ...(key ? { key } : {}) }
  }
  const record: MorphPlanRecord = {
    v: 1, from: job.from, to: job.to, fromStep: job.fromStep,
    items: plan.items.map(item => {
      const { from: f, to: t, ...rest } = item
      return { ...rest, ...(f ? { from: side(f) } : {}), ...(t ? { to: side(t) } : {}) }
    }),
    over: plan.over.map(r => MORPH_ASSET_BASE + r),
    pairs: plan.pairs.map(p => ({ fromKey: p.fromKey, toKey: p.toKey, rule: p.rule })),
  }
  try { encodeShowMessage({ tag: 5, morph: record }) } catch (e) { if (e instanceof PlanRejected) return composited('limits'); throw e }
  // §7 (D7, D9) at the Morph's two boundaries: the shipped assets composed against the masked slide marked once.
  if (!validateMorph(ids.map((ref, i) => ({ ref, asset: assets[i]! })), proved.rasters, [lists[0]!, lists[lists.length - 1]!], { scale, W, H, page, job, docKey })) return composited('marks')
  return { record, assets, diagnostics: [], page: { w: W, h: H } }
}

// ---------------------------------------------------------------------------
// Sampling
// ---------------------------------------------------------------------------

/** The draws `morphFrame` composites at `progress`, recorded (handle ids are ref indices). */
function drawsAtProgress(spec: MorphFrameSpec, progress: number): LayerDraw[] {
  let out: LayerDraw[] = []
  const rec = { frame: (draws: readonly LayerDraw[]) => { out = [...draws] } } as unknown as Compositor
  morphFrame(spec, progress, rec)
  return out
}
/** A frame's draws bottom-first, a summed pair's two sides in order (the outgoing side, then the incoming). */
const flatten = (list: readonly LayerDraw[]): LayerDraw[] => list.flatMap(d => (d.sum ? [d, d.sum.with] : [d]))
/** `list` without ref `ref` (a summed pair keeps its other side, drawn on its own). */
function without(list: readonly LayerDraw[], ref: number): LayerDraw[] {
  const out: LayerDraw[] = []
  for (const d of list) {
    if (d.sum) {
      const { sum, ...a } = d
      if (d.handle.id === ref) out.push(sum.with)
      else if (sum.with.handle.id === ref) out.push(a)
      else out.push(d)
      continue
    }
    if (d.handle.id !== ref) out.push(d)
  }
  return out
}
const drawn = (s: TargetState | undefined): s is TargetState => !!s && s.visible && s.opacity > 0

function corners(m: Affine, b: Rect): Array<[number, number]> {
  return [[b.x, b.y], [b.x + b.w, b.y], [b.x, b.y + b.h], [b.x + b.w, b.y + b.h]].map(([x, y]) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as [number, number])
}

/** A Canvas 2D compositor at the bucket over `rasters` (by ref); `compose` returns the frame's RGBA. */
function composerOf(rasters: ReadonlyMap<number, LayerBitmap>, scale: number, W: number, H: number) {
  const canvas = factory(W, H)
  const c = createCanvas2DCompositor(canvas, factory, OPEN_WORKSPACE, createDiagnosticSink())
  c.setViewport({ matrix: [scale, 0, 0, scale, 0, 0], deviceWidth: W, deviceHeight: H })
  const real = new Map<number, LayerHandle>()
  let refused = false
  for (const [ref, bmp] of rasters) {
    const up = c.upload(bmp, open(0, 'morph'))
    if ('refused' in up) refused = true
    else real.set(ref, up)
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const map = (d: LayerDraw): LayerDraw | undefined => {
    const handle = real.get(d.handle.id)
    if (!handle) return undefined
    if (!d.sum) return { ...d, handle }
    const other = map(d.sum.with)
    return other ? { ...d, handle, sum: { ...d.sum, with: other } } : { ...d, handle, sum: undefined }
  }
  return {
    refused,
    compose(list: readonly LayerDraw[]): Uint8ClampedArray {
      c.frame(list.flatMap(d => { const m = map(d); return m ? [m] : [] }))
      return ctx.getImageData(0, 0, W, H).data
    },
    dispose: () => c.dispose(),
  }
}

// ---------------------------------------------------------------------------
// The proof subset (§6 rules 1–5)
// ---------------------------------------------------------------------------

interface Proved { rasters: Map<number, LayerBitmap>; movers: number }

function prove(order: readonly number[], bmpOf: (ref: number) => LayerBitmap | undefined, used: ReadonlySet<number>, lists: readonly LayerDraw[][], scale: number, W: number, H: number): Proved | Unproven {
  const rasters = new Map<number, LayerBitmap>()
  for (const r of used) { const b = bmpOf(r); if (!b) return 'limits'; rasters.set(r, b) }
  // Rule 5: one full composite per sample, plus one without each drawn ref, plus the reconstruction.
  const cost = lists.reduce((n, l) => n + 2 + flatten(l).filter(d => drawn(d.state)).length, 0)
  if (cost > CLASSIFY_BUDGET) return 'budget'
  const composer = composerOf(rasters, scale, W, H)
  try {
    if (composer.refused) return 'limits'
    // Rules 1–3: each ref's envelope over the samples.
    const seen = new Map<number, Uint8Array>()
    for (const [ref, bmp] of rasters) seen.set(ref, new Uint8Array(bmp.canvas.width * bmp.canvas.height))
    const originals: Uint8ClampedArray[] = []
    for (const list of lists) {
      const full = new Uint8ClampedArray(composer.compose(list))
      originals.push(full)
      for (const d of flatten(list)) {
        if (!drawn(d.state)) continue
        const ref = d.handle.id, bmp = rasters.get(ref)!, mark = seen.get(ref)!
        const sm = stateMatrix(d.state)
        const m: Affine = [scale * sm[0], scale * sm[1], scale * sm[2], scale * sm[3], scale * sm[4], scale * sm[5]]
        const inv = invert(m)
        if (!inv) continue
        const box = corners(m, bmp.bounds)
        const x0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[0])))), y0 = Math.max(0, Math.floor(Math.min(...box.map(p => p[1]))))
        const x1 = Math.min(W, Math.ceil(Math.max(...box.map(p => p[0])))), y1 = Math.min(H, Math.ceil(Math.max(...box.map(p => p[1]))))
        if (!(x1 > x0 && y1 > y0)) continue
        const less = composer.compose(without(list, ref))
        const bw = bmp.canvas.width, bh = bmp.canvas.height, b = bmp.bounds, kx = bw / b.w, ky = bh / b.h
        const lu = (x: number, y: number) => ((inv[0] * x + inv[2] * y + inv[4]) - b.x) * kx
        const lv = (x: number, y: number) => ((inv[1] * x + inv[3] * y + inv[5]) - b.y) * ky
        for (let y = y0; y < y1; y++) {
          for (let x = x0; x < x1; x++) {
            const i = (y * W + x) * 4
            if (full[i] === less[i] && full[i + 1] === less[i + 1] && full[i + 2] === less[i + 2] && full[i + 3] === less[i + 3]) continue
            // The device pixel square's four corners in the bitmap, as scalars (no per-pixel arrays).
            const ua = lu(x, y), ub = lu(x + 1, y), uc = lu(x, y + 1), ud = lu(x + 1, y + 1)
            const va = lv(x, y), vb = lv(x + 1, y), vc = lv(x, y + 1), vd = lv(x + 1, y + 1)
            const u0 = Math.max(0, Math.floor(Math.min(ua, ub, uc, ud))), v0 = Math.max(0, Math.floor(Math.min(va, vb, vc, vd)))
            const u1 = Math.min(bw, Math.ceil(Math.max(ua, ub, uc, ud))), v1 = Math.min(bh, Math.ceil(Math.max(va, vb, vc, vd)))
            if (u1 > u0) for (let v = v0; v < v1; v++) mark.fill(1, v * bw + u0, v * bw + u1)
          }
        }
      }
    }
    // Rule 3: unobserved inked pixels (dilated by 1 px) are removed.
    const masks = new Map<number, Uint8Array>()
    for (const [ref, bmp] of rasters) {
      const m = maskOf(bmp, seen.get(ref)!)
      if (m) masks.set(ref, m)
    }
    // Rule 4 over the items in z order (a summed pair's sides in order), with the final masks.
    // `order` is the paint order at every sample; a side absent at a sample (weight 0) has an undefined state there.
    const states = (ref: number) => lists.map(l => flatten(l).find(d => d.handle.id === ref)?.state)
    const items: ProofItem[] = order.map(ref => {
      const bmp = rasters.get(ref)!
      const mask = masks.get(ref)
      return { bounds: bmp.bounds, size: { w: bmp.canvas.width, h: bmp.canvas.height }, ...(mask ? { mask } : {}), states: states(ref) }
    })
    const proof = proveItems(items, scale)
    if (!proof.ok) return proof.reason
    // Reconstruction: the masked rasters compose every sample exactly.
    if (masks.size) {
      const masked = new Map(rasters)
      for (const [ref, m] of masks) masked.set(ref, maskedOf(rasters.get(ref)!, m))
      const again = composerOf(masked, scale, W, H)
      try {
        if (again.refused) return 'limits'
        for (let k = 0; k < lists.length; k++) {
          const got = again.compose(lists[k]!), want = originals[k]!
          for (let i = 0; i < want.length; i++) if (got[i] !== want[i]) return 'reconstruction'
        }
      } finally { again.dispose() }
      return { rasters: masked, movers: proof.movers.length }
    }
    return { rasters, movers: proof.movers.length }
  } finally {
    composer.dispose()
  }
}

const bitmapOf = (w: number, h: number, rgba: Uint8ClampedArray, like: LayerBitmap): LayerBitmap => {
  const canvas = canvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const img = ctx.createImageData(w, h)
  img.data.set(rgba)
  ctx.putImageData(img, 0, 0)
  return { canvas, bounds: { ...like.bounds }, scale: like.scale, bytes: w * h * 4 }
}

/**
 * §7 at progress 0 and 1 (as `validate` in `build.ts` does at step boundaries): the shipped PNGs (visible mark at
 * rest, the validation code's Layer-1 at each `rest` crop) composed, against the masked rasters composed once,
 * marked visibly once and with Layer-1 once; the visible mark's own contribution over its ink; and the validation
 * code attributed on the composite (`detect`).
 */
function validateMorph(shipped: ReadonlyArray<{ ref: number; asset: MorphAsset }>, rasters: ReadonlyMap<number, LayerBitmap>, ends: readonly LayerDraw[][], env: { scale: number; W: number; H: number; page: { w: number; h: number }; job: MorphJob; docKey: string }): boolean {
  const { scale, W, H, page, job } = env
  const kBasis = createHash('sha256').update(`officeview show-morph validation basis@1\n${env.docKey}\n${job.to}\n${scale}`).digest()
  const codes = [1, 2, 3, 4].map(t => extendedCodeword(messageFor(VALIDATION_MESSAGE_KEY, 0, t, 0)))
  const code = codes[0]!
  let chips: ReturnType<typeof basis> | undefined
  try { chips = basis(kBasis, W, H, DEFAULT_PARAMS) } catch { chips = undefined }
  const sent = new Map<number, LayerBitmap>(), visible = new Map<number, LayerBitmap>()
  for (const { ref, asset } of shipped) {
    const like = rasters.get(ref)!
    const v = decodePng(asset.png)
    visible.set(ref, bitmapOf(v.w, v.h, v.rgba, like))
    const r = markLayerRaster({ png: asset.png, w: asset.w, h: asset.h, kBasis, code, crop: { x: asset.rest.x, y: asset.rest.y, pageW: W, pageH: H, scale: asset.rest.scale } })
    const d = decodePng(r.png)
    sent.set(ref, bitmapOf(d.w, d.h, d.rgba, like))
  }
  const layered = composerOf(sent, scale, W, H), visibleOnly = composerOf(visible, scale, W, H), reference = composerOf(rasters, scale, W, H)
  try {
    if (layered.refused || visibleOnly.refused || reference.refused) return false
    const image = (data: Uint8ClampedArray) => ({ width: W, height: H, data, colorSpace: 'srgb' }) as unknown as ImageData
    const visibleRefs = new Map<number, Uint8ClampedArray>()
    const visibleRef = (k: number) => {
      let v = visibleRefs.get(k)
      if (!v) { v = markFrame(new Uint8ClampedArray(reference.compose(ends[k]!)), W, H, scale, page, job.mark); visibleRefs.set(k, v) }
      return v
    }
    const tau = tauFor(codes.length)
    const verdict = validateMarks({ plan: { steps: [0] } },
      (k) => {
        const r = markLayerRaster({ png: encodePng(W, H, visibleRef(k)), w: W, h: H, kBasis, code, crop: { x: 0, y: 0, pageW: W, pageH: H } })
        return image(decodePng(r.png).rgba)
      },
      (k) => image(new Uint8ClampedArray(layered.compose(ends[k]!))),
      {
        ...(chips ? {
          attributes: (k: number, composite: ImageData) => {
            const bmp = (d: Uint8ClampedArray) => ({ width: W, height: H, data: d })
            const decision = detect(bmp(new Uint8ClampedArray(composite.data)), bmp(visibleRef(k)), chips!, codes, DEFAULT_PARAMS, tau)
            return decision.kind === 'attributed' && decision.index === 0
          },
        } : {}),
        contribution: {
          composeMarked: (k) => image(new Uint8ClampedArray(visibleOnly.compose(ends[k]!))),
          composeUnmarked: (k) => image(new Uint8ClampedArray(reference.compose(ends[k]!))),
          referenceMarked: (k) => image(visibleRef(k)),
          referenceUnmarked: (k) => image(new Uint8ClampedArray(reference.compose(ends[k]!))),
        },
      })
    return verdict.ok
  } finally {
    layered.dispose(); visibleOnly.dispose(); reference.dispose()
  }
}

function maskOf(bmp: LayerBitmap, seen: Uint8Array): Uint8Array | undefined {
  const w = bmp.canvas.width, h = bmp.canvas.height
  const data = (bmp.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, w, h).data
  const mask = new Uint8Array(w * h).fill(255)
  let removed = false
  for (let v = 0; v < h; v++) {
    for (let u = 0; u < w; u++) {
      const i = v * w + u
      if (data[i * 4 + 3] === 0 || seen[i]) continue
      let near = false
      for (let dv = -1; dv <= 1 && !near; dv++) for (let du = -1; du <= 1; du++) {
        const uu = u + du, vv = v + dv
        if (uu >= 0 && vv >= 0 && uu < w && vv < h && seen[vv * w + uu]) { near = true; break }
      }
      if (!near) { mask[i] = 0; removed = true }
    }
  }
  return removed ? mask : undefined
}

function maskedOf(bmp: LayerBitmap, mask: Uint8Array): LayerBitmap {
  const w = bmp.canvas.width, h = bmp.canvas.height
  const canvas = canvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(bmp.canvas as CanvasImageSource, 0, 0)
  const img = ctx.getImageData(0, 0, w, h)
  for (let i = 0; i < w * h; i++) if (mask[i] === 0) img.data[i * 4 + 3] = 0
  ctx.putImageData(img, 0, 0)
  return { canvas, bounds: { ...bmp.bounds }, scale: bmp.scale, bytes: w * h * 4 }
}

// ---------------------------------------------------------------------------
// The composited Morph (D11)
// ---------------------------------------------------------------------------

function compositeMorph(plan: MorphPlan, bitmaps: ReadonlyMap<MorphLayerRef, LayerBitmap>, used: ReadonlySet<number>, spec: MorphFrameSpec, job: MorphJob, page: { w: number; h: number }, W: number, H: number, dur: number, why: Unproven): BuiltMorph {
  const need = Math.max(1, Math.ceil((dur * TRACK_FPS) / 1000))
  const count = Math.min(TRACK_MAX_FRAMES, need)
  const rasters = new Map<number, LayerBitmap>()
  for (const r of used) { const b = bitmaps.get(plan.refs[r]!); if (b) rasters.set(r, b) }
  const composer = composerOf(rasters, job.scale, W, H)
  try {
    const frames = trackTimes(dur, count).map((t, i, all): MorphAsset => {
      const progress = i === all.length - 1 || dur === 0 ? 1 : t / dur
      const rgba = markFrame(new Uint8ClampedArray(composer.compose(drawsAtProgress(spec, progress))), W, H, job.scale, page, job.mark)
      return { assetId: MORPH_FRAME_BASE + i, w: W, h: H, png: encodePng(W, H, rgba), rest: { x: 0, y: 0, scale: 1 } }
    })
    const record_: MorphPlanRecord & { composited: { fps: 15; count: number } } = {
      v: 1, from: job.from, to: job.to, fromStep: job.fromStep, items: [], over: [],
      pairs: plan.pairs.map(p => ({ fromKey: p.fromKey, toKey: p.toKey, rule: p.rule })).filter(p => p.fromKey.length <= SHOW_PLAN_LIMITS.str && p.toKey.length <= SHOW_PLAN_LIMITS.str).slice(0, SHOW_PLAN_LIMITS.morphItems),
      composited: { fps: 15, count },
    }
    return { record: record_, frames, diagnostics: ['morph-composited', ...(need > count ? ['step-truncated'] : [])], page: { w: W, h: H }, reason: why }
  } finally {
    composer.dispose()
  }
}
