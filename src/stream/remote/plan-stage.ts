/**
 * PlanStage (streaming Stage 3 piece 4, Task 8; spec §4–§6, plan D11, D14,
 * D15, D16, D17, D22, D23, D25): the driver stage of a streamed show. It plays
 * a server show plan (the OVSP record, its layer bitmaps, keyframe tracks and
 * Morph records) on the shared compositor, behaviours, transitions and
 * `morphFrame`, with no package, no shapes and no text runs.
 *
 * - **State.** The slide's render state is the local stage's (`stateOf`):
 *   `renderStateAt` over the plan's base geometry, relative to the base
 *   rotation each layer is painted at, a target with an entrance hidden until
 *   that entrance begins (`timingTargets`, over the tree the driver schedules).
 *   Text by word or letter fans out here (`planIterations`, D22) from the
 *   items the server sent, so the fan is local mode's by construction.
 * - **Layers.** The plan's records, in z order, are drawn as the local
 *   `entriesOf` draws its plan: a container record's state composes over its
 *   children; a range layer's mask sweeps the record's `maskBox`; a split
 *   layer takes its part's colour.
 * - **Composited steps** (§6 rule 5, 6). Once a `composited` step has begun no
 *   layer is drawn for it: its track plays frame `min(count − 1, ⌊t·15⌋)`
 *   (`t` in s from its click), then holds its end frame until a `layers` step
 *   begins. A frame not yet decoded holds the previous picture (D23). A track
 *   not there when the step begins (`'pending'`) makes the step a cut to its
 *   end frame, once that arrives (`step-deferred`).
 * - **Triggers** (§6a): a trigger the server cut (`cut: true`) draws the
 *   targets of its sequences at their end state once fired; the others
 *   animate from the layers.
 * - **Transitions.** The outgoing frame is what was on screen, composited by
 *   `flattenCurrent` with a Canvas 2D compositor into a workspace-leased
 *   canvas (a still slide's from-frame is its still). The incoming frame is
 *   the entry frame (or, entering at a later checkpoint, that checkpoint
 *   composed the same way). Reduced motion uses `reducedTransition`.
 * - **Morph** (D11): the pair `(from, to, fromStep)` is asked for when the
 *   slide is entered; the outgoing frame holds for up to `MORPH_WAIT_MS` on
 *   the show clock, polled per frame. A layered record plays `morphFrame`; a
 *   composited one its track; nothing in time plays the fade (`morph-late`).
 * - **Compositor** (D17): `createPlanCompositor`; a lost WebGL2 context
 *   switches to Canvas 2D and every bitmap is re-uploaded. Under workspace
 *   pressure the slide holds its entry frame, static (`plan-stage-pressure`).
 * - **Leases.** The bitmaps are the provider's (charged by the bundle); each
 *   upload takes a zero-byte lease that the compositor grows by its texture.
 *   `dispose` releases every lease the stage or its compositors took.
 *
 * Package-free (D25): value imports are the engine, the affine helpers, the
 * transitions, `morph/render` and the painter-free compositor only.
 */
import { apply, invert, mul, stateMatrix } from '../../slideshow/affine'
import { expandIterations, renderStateAt, type BaseGeometry, type IteratedTiming, type IterateFan } from '../../slideshow/engine/behaviours'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../../slideshow/engine/steps'
import { timingTargets } from '../../slideshow/engine/targets'
import { morphFrame, type MorphFrameItem, type MorphFrameSpec, type MorphLayerFrame } from '../../slideshow/morph/render'
import { createCanvas2DCompositor } from '../../slideshow/render/canvas2d'
import { createPlanCompositor, type PlanCompositor } from '../../slideshow/render/plan-compositor'
import { dispatchTransition, transition3DDraw } from '../../slideshow/render/three/framework'
import { reducedTransition, resolveTransition2D, transitionDraw } from '../../slideshow/render/transitions2d'
import { createRng } from '../../slideshow/rng'
import { targetKey } from '../../slideshow/target'
import { surfaceBytes } from '../../slideshow/workspace'
import type { DriverStage, HitResult } from '../../slideshow/player/driver'
import type {
  Affine, BehaviourNode, CanvasFactory, Compositor, DiagnosticSink, LayerDraw, LayerHandle, Lease, RenderState, Rng,
  ShowSlide, SlideShowWorkspace, TargetKey, TargetState, TimeNode, TimingTree, TransitionSpec,
} from '../../slideshow/types'
import type { MorphPlanRecord, PlanContainerRecord, PlanLayerRecord, Rect, ShowPlanRecord } from '../show-plan'
import type { ImageBitmapLike } from './show-fetch'

export interface PlanAssets { bitmap(assetId: number): ImageBitmapLike | undefined; bytes(assetId: number): number }
export interface PlanTrack { readonly fps: 15; readonly count: number; frame(i: number): ImageBitmapLike | undefined /* D23: undefined → hold previous */; want(i: number): void /* moves the 8-frame decode window */ }
export interface PlanSlideData { plan: ShowPlanRecord; assets: PlanAssets; track(step: number): PlanTrack | 'pending' | undefined }
export interface PlanStageDeps {
  root: HTMLElement
  output: HTMLCanvasElement
  workspace: SlideShowWorkspace
  sink: DiagnosticSink
  reduced: boolean
  showNow(): number
  data(index: number): PlanSlideData | undefined
  still(index: number): ImageBitmapLike | undefined
  morph(from: number, to: number, fromStep: number): Promise<{ record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack } | undefined>
  seed?: number
  /** Additive: start on the Canvas 2D compositor (tests, the forced path). */
  force?: 'canvas2d'
}
export interface PlanStage extends DriverStage {
  composeAt(index: number, cursor: SlideCursor, canvas: HTMLCanvasElement): boolean
  resize(css: { w: number; h: number }, dpr: number): void
  dispose(): void
  /** @internal Additive: what the last frame drew, the running transition and the backend (tests, the debug harness). */
  inspect(): PlanStageInspection
}
export interface PlanStageInspection {
  current: number | undefined
  /** `layer:<id>`, `entry`, `still`, `track:<step>:<i>`, `from` (the outgoing frame), `transition`, `morph`, `morph-track:<i>`. */
  lastDraws: string[]
  transition: { effect: string; duration: number; from: 'flat' | 'black' | 'none'; to: 'entry' | 'flat' } | undefined
  backend: PlanCompositor['state']
}

/** D11: how long the outgoing frame holds for a Morph record or track before the fade plays (`morph-late`). */
export const MORPH_WAIT_MS = 500
/** D23: decoded frames a track holds at once. */
export const TRACK_WINDOW = 8
const TRACK_FPS = 15
/** Container nesting walked for a hit test (the codec's records name only earlier containers, so no cycle forms). */
const SHOW_LAYER_DEPTH = 128

/**
 * D22: the slide's text items fanned out over `tree` exactly as local mode
 * fans them (`expandIterations` over the items and word groups the server
 * computed from the package).
 */
export function planIterations(tree: TimingTree, iterate: ShowPlanRecord['iterate'], options: { charge?: (bytes: number, tag: string) => boolean; sink?: DiagnosticSink; slideIndex?: number } = {}): IteratedTiming {
  const keys = new Map<TargetKey, readonly TargetKey[]>(iterate.keys.map(([k, items]) => [k, items]))
  const groups = new Map(iterate.groups.map(([k, g]) => [k, { ...(g.wd ? { wd: g.wd } : {}), ...(g.lt ? { lt: g.lt } : {}) }]))
  return expandIterations(tree, keys, { ...options, groups })
}

/**
 * D23: a keyframe track decoded on a rolling window of `TRACK_WINDOW` frames.
 * Each decoded frame is charged `frameBytes` to `workspace` before it is
 * decoded and released when it leaves the window (its bitmap closed). The
 * encoded frames are the caller's (`decode(i)`).
 */
export function createTrackWindow(o: { count: number; frameBytes: number; workspace: SlideShowWorkspace; decode(i: number): Promise<ImageBitmapLike | undefined> }): PlanTrack & { dispose(): void } {
  const held = new Map<number, { lease: Lease; bmp?: ImageBitmapLike }>()
  let lo = 0
  let disposed = false
  const close = (b: ImageBitmapLike | undefined) => { try { b?.close?.() } catch { /* already released */ } }
  const drop = (i: number) => { const h = held.get(i); if (!h) return; held.delete(i); close(h.bmp); h.lease.release() }
  return {
    fps: 15,
    count: o.count,
    frame: (i) => held.get(i)?.bmp,
    want(i) {
      if (disposed) return
      lo = Math.max(0, Math.min(o.count - 1, Math.floor(i)))
      for (const k of [...held.keys()]) if (k < lo || k >= lo + TRACK_WINDOW) drop(k)
      for (let k = lo; k < Math.min(o.count, lo + TRACK_WINDOW); k++) {
        if (held.has(k)) continue
        const lease = o.workspace.lease(o.frameBytes, 'plan:track-frame')
        if (!lease) break
        const slot: { lease: Lease; bmp?: ImageBitmapLike } = { lease }
        held.set(k, slot)
        o.decode(k).then(
          (bmp) => { if (held.get(k) === slot) slot.bmp = bmp; else close(bmp) },
          () => { if (held.get(k) === slot) drop(k) },
        )
      }
    },
    dispose() {
      if (disposed) return
      disposed = true
      for (const k of [...held.keys()]) drop(k)
    },
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
const PLAIN: TargetState = Object.freeze({ visible: true, opacity: 1, transform: IDENTITY })
const HIDDEN: TargetState = Object.freeze({ visible: false, opacity: 1, transform: IDENTITY })
const EMPTY_STATE: RenderState = new Map()
/** Slide time far past every finite effect: a settled state. */
const SETTLED_MS = 1e9

/** The group state composed over a sub-layer's own (the local stage's `compose`). */
function compose(parent: TargetState | undefined, own: TargetState | undefined): TargetState {
  if (!parent) return own ?? PLAIN
  const p = stateMatrix(parent), o = stateMatrix(own)
  return {
    visible: parent.visible && (own?.visible ?? true),
    opacity: parent.opacity * (own?.opacity ?? 1),
    transform: mul(p, o),
    ...(own?.mask ? { mask: own.mask, ...(own.maskBox ? { maskBox: own.maskBox } : {}), ...(own.maskTransform ? { maskTransform: mul(p, own.maskTransform) } : {}) } : {}),
    ...(own?.color ?? parent.color ? { color: own?.color ?? parent.color } : {}),
  }
}

/** `transform ∘ rotate(−base.rot about the centre)`: layers are painted at their base rotation already. */
function relativeTo(state: TargetState, g: BaseGeometry | undefined): TargetState {
  if (!g || !g.rot) return state
  const a = -g.rot * Math.PI / 180, c = Math.cos(a), s = Math.sin(a)
  const cx = g.x + g.w / 2, cy = g.y + g.h / 2
  const back: Affine = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
  return { ...state, transform: mul(state.transform, back) }
}

const isContainer = (l: PlanLayerRecord | PlanContainerRecord): l is PlanContainerRecord => 'container' in l && l.container === true
const inRect = (r: Rect, x: number, y: number) => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h
const nexts = (c: SlideCursor) => c.events.reduce((n, e) => n + (e.kind === 'next' ? 1 : 0), 0)

/** Per trigger shape (spid): the target keys its interactive sequences animate, with their text items. */
function triggerTargets(tree: TimingTree | undefined, items: ReadonlyMap<TargetKey, readonly TargetKey[]> | undefined): Map<string, Set<TargetKey>> {
  const out = new Map<string, Set<TargetKey>>()
  if (!tree) return out
  const walk = (n: TimeNode, under: readonly string[], depth: number) => {
    if (depth > 128) return
    if ('children' in n) {
      let here = under
      if (n.ctn.nodeType === 'interactiveSeq') {
        const spids = n.ctn.start.flatMap(c => (c.target && 'spid' in c.target ? [c.target.spid] : []))
        here = [...under, ...spids]
      }
      for (const c of n.children) walk(c, here, depth + 1)
      return
    }
    if (!under.length || !('spid' in n.target)) return
    const k = targetKey(n.target)
    for (const spid of under) {
      let set = out.get(spid)
      if (!set) out.set(spid, set = new Set())
      set.add(k)
      for (const i of items?.get(k) ?? []) set.add(i)
    }
  }
  walk(tree.root, [], 0)
  return out
}

// ---------------------------------------------------------------------------
// The stage
// ---------------------------------------------------------------------------

/** A bitmap the stage draws: its key (unique per slide), the provider's bitmap and where it is placed. */
interface Src { key: string; bmp: ImageBitmapLike; bounds: Rect; scale: number }
/** One draw of a picture: a source at a state. */
interface Pic { src: Src; state: TargetState; part?: 'fill' | 'line' | 'text'; kind: string }
interface Up { handle: LayerHandle; epoch: number }
/**
 * A full-output frame the stage composited for itself (the outgoing frame, a flattened incoming frame, black):
 * its canvas (leased at its surface bytes) is kept, and uploaded (a texture lease of its own) under the viewport in
 * force when it is drawn, so a slide of another size or a backend switch re-uploads it.
 */
interface Flat { canvas: HTMLCanvasElement; lease: Lease; up?: { handle: LayerHandle; epoch: number; key: string } }

interface Slide {
  index: number
  data: PlanSlideData | undefined
  W: number
  H: number
  base: Map<TargetKey, BaseGeometry>
  hidden: Set<TargetKey>
  entrances: Map<TargetKey, Set<TimeNode>>
  fan: ReadonlyMap<BehaviourNode, IterateFan> | undefined
  /** Per trigger spid, the targets its sequences animate (`cut` triggers draw them at their end state). */
  triggerKeys: Map<string, Set<TargetKey>>
  rng: Rng
  /** Uploaded sources by key, at an epoch. */
  ups: Map<string, Up>
  pressure: boolean
  /** Track activations (`step|t0`) that began without their track: cuts to their end frame. */
  deferred: Set<string>
  /** The last picture drawn (held while a track frame is missing; the outgoing frame's content). */
  held: Pic[]
  timeline: SlideTimeline | undefined
  zOf: Map<TargetKey, number>
  /** `trigger-cut` was diagnosed (once per slide visit, never per frame). */
  cutNoted: boolean
}

interface MorphRun {
  spec?: MorphFrameSpec
  track?: PlanTrack
  handles: LayerHandle[]
  /** The source each layered handle was uploaded from (an interrupted Morph is recomposed from them). */
  srcOf: Map<LayerHandle, Src>
  epoch: number
}

/** What the last frame showed, so `flattenCurrent` captures exactly that (review fix round 1). */
type Shown =
  | { kind: 'pics' }
  | { kind: 'from' }
  | { kind: 'transition'; progress: number; params: Readonly<Record<string, unknown>> }
  | { kind: 'morph'; progress: number }

/** The displayed track frame, copied out of the decode window (which may close it) and workspace-charged. */
interface HeldFrame { key: string; canvas: HTMLCanvasElement; lease: Lease }

interface Running {
  spec: TransitionSpec
  three: boolean
  from?: Flat
  /** The incoming frame: the slide's entry frame, or the checkpoint it was entered at, flattened. */
  to: 'entry' | Flat
  start: number
  duration: number
  rng: Rng
  morph?: MorphRun
  fromKind: 'flat' | 'black' | 'none'
  toKind: 'entry' | 'flat'
}

/** A Morph asked for and not yet started. */
interface Waiting {
  p: Slide
  spec: TransitionSpec
  duration: number
  asked: number
  /** The cursor the slide was entered at (its incoming frame if the fade plays). */
  cursor: SlideCursor
  result?: { record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack } | 'refused'
}

export function createPlanStage(deps: PlanStageDeps): PlanStage {
  const { root, output, workspace, sink } = deps
  const doc = root.ownerDocument
  const factory: CanvasFactory = (w, h) => {
    const c = doc.createElement('canvas')
    c.width = Math.max(1, Math.ceil(w)); c.height = Math.max(1, Math.ceil(h))
    return c
  }
  const container = doc.createElement('div')
  container.setAttribute('data-officeview-show-stage', '')
  output.after(container)

  let epoch = 0
  let disposed = false
  let current: Slide | undefined
  let origin = Number.POSITIVE_INFINITY
  /** The outgoing frame (`flattenCurrent`), drawn between slides and as the transition's source. */
  let outgoing: Flat | undefined
  let running: Running | undefined
  let waiting: Waiting | undefined
  let captured = false
  let shownBefore = false
  let left: { index: number; step: number } | undefined
  let lastStep = 0
  let kinds: string[] = []
  /** One track frame uploaded at a time: its slide-local key. */
  let trackKey: string | undefined
  let heldFrame: HeldFrame | undefined
  let shown: Shown = { kind: 'pics' }
  /** Fanned text timing per `ShowSlide` (handed to the driver, so its node identities are the timeline's). */
  const texts = new WeakMap<ShowSlide, IteratedTiming | null>()

  const manager: PlanCompositor = createPlanCompositor(container, {
    output, workspace, sink, ...(deps.force ? { force: deps.force } : {}),
    rebuild: (c) => {
      // Every handle belonged to the old backend, which released its leases.
      epoch++
      // The stage's own frames keep their canvases and re-upload when next drawn; a Morph's layers go (a cut).
      if (running?.morph) { releaseFlat(running.from); running.morph.handles = []; running = undefined }
      if (outgoing) outgoing.up = undefined
      if (running) { if (running.from) running.from.up = undefined; if (running.to !== 'entry') running.to.up = undefined }
      trackKey = undefined
      const p = current
      if (!p) return
      p.ups.clear()
      c.setSlideIndex(p.index)
      setView(p)
      load(p)
    },
  })
  const compositor = (): Compositor => manager.current

  // -------------------------------------------------------------------
  // Viewport
  // -------------------------------------------------------------------

  /** The viewport last set on the compositor (the stage's own frames are uploaded under it). */
  let view: { matrix: Affine } = { matrix: IDENTITY }
  function setView(p: Slide | undefined) {
    const v = viewport(p)
    view = v
    return compositor().setViewport(v)
  }
  const sizeOf = (p: Slide | undefined) => (p ? { w: p.W, h: p.H } : { w: 960, h: 540 })
  function viewport(p: Slide | undefined, w = output.width, h = output.height) {
    const s0 = sizeOf(p)
    const s = Math.min(w / s0.w, h / s0.h)
    return { matrix: [s, 0, 0, s, (w - s0.w * s) / 2, (h - s0.h * s) / 2] as Affine, deviceWidth: w, deviceHeight: h, scale: s }
  }

  // -------------------------------------------------------------------
  // Slides
  // -------------------------------------------------------------------

  function slideOf(index: number, meta: ShowSlide | undefined, own?: IteratedTiming): Slide {
    const data = deps.data(index)
    const plan = data?.plan
    const text = own ?? (meta ? texts.get(meta) ?? undefined : undefined)
    const timing = meta?.timing ?? plan?.show.timing
    const tree = text?.tree ?? (timing && !('refused' in timing) ? timing : undefined)
    const t = timingTargets({ index, hidden: false, timing: tree, notes: [], links: [], media: [], builds: [] })
    if (text) {
      // Every text item is animated, hidden and entered as its iterated target is (the local `targetsOf`).
      for (const [target, keys] of text.items) {
        for (const k of keys) {
          if (t.hidden.has(target)) {
            t.hidden.add(k)
            t.entrances.set(k, t.entrances.get(target) ?? new Set())
          }
        }
      }
    }
    const zOf = new Map<TargetKey, number>()
    for (const l of plan?.layers ?? []) if (l.key !== undefined) zOf.set(l.key, Math.max(zOf.get(l.key) ?? -Infinity, l.z))
    const still = plan ? undefined : deps.still(index)
    return {
      index, data, W: plan?.slide.w ?? still?.width ?? 960, H: plan?.slide.h ?? still?.height ?? 540,
      base: new Map(plan?.base ?? []), hidden: t.hidden, entrances: t.entrances, fan: text?.fan,
      triggerKeys: triggerTargets(tree, text?.items),
      // The server's unit index seeds the rng, as its tracks do (`0x5eed + unit`).
      rng: createRng((deps.seed ?? 0x5eed) + (plan?.slide.index ?? index)),
      ups: new Map(), pressure: false, deferred: new Set(), held: [], timeline: undefined, zOf, cutNoted: false,
    }
  }

  /** Uploads `src` to the compositor (zero-byte lease, grown by the texture); `undefined` when refused. */
  function upload(src: Src, tag: string, c: Compositor = compositor()): LayerHandle | undefined {
    const lease = workspace.lease(0, tag)
    if (!lease) return undefined
    const up = c.upload({ canvas: src.bmp as unknown as HTMLCanvasElement, bounds: src.bounds, scale: src.scale, bytes: 0 }, lease)
    if ('refused' in up) { lease.release(); return undefined }
    return up
  }

  /** The handle of `src` on `p` (uploaded on first use). */
  function handleOf(p: Slide, src: Src): LayerHandle | undefined {
    const have = p.ups.get(src.key)
    if (have && have.epoch === epoch) return have.handle
    const h = upload(src, `plan:${p.index}:${src.key}`)
    if (h) p.ups.set(src.key, { handle: h, epoch })
    return h
  }

  function releaseUp(p: Slide, key: string) {
    const u = p.ups.get(key)
    if (!u) return
    p.ups.delete(key)
    if (u.epoch === epoch) compositor().release(u.handle)
  }

  function releaseSlide(p: Slide) {
    for (const key of [...p.ups.keys()]) releaseUp(p, key)
    p.held = []
    heldFrame?.lease.release()
    heldFrame = undefined
    trackKey = undefined
  }

  /** A charged copy of a track frame: the decode window may close the provider's bitmap once it moves on (D23). */
  function copyFrame(src: Src): HeldFrame | undefined {
    const w = src.bmp.width, h = src.bmp.height
    if (!(w > 0 && h > 0)) return undefined
    const lease = workspace.lease(surfaceBytes(w, h), 'plan:held-frame')
    if (!lease) return undefined
    const canvas = factory(w, h) as HTMLCanvasElement
    const ctx = canvas.getContext('2d')
    if (!ctx) { lease.release(); return undefined }
    ctx.drawImage(src.bmp as unknown as CanvasImageSource, 0, 0)
    return { key: src.key, canvas, lease }
  }

  const entrySrc = (p: Slide): Src | undefined => {
    const bmp = p.data ? p.data.assets.bitmap(p.data.plan.entry) : deps.still(p.index)
    if (bmp && !p.data && (p.W !== bmp.width || p.H !== bmp.height) && bmp.width > 0 && bmp.height > 0) {
      // A still slide is its bitmap: its px are the slide's (it may arrive after the slide is entered).
      p.W = bmp.width
      p.H = bmp.height
      if (p === current) setView(p)
    }
    return bmp ? { key: p.data ? 'entry' : 'still', bmp, bounds: { x: 0, y: 0, w: p.W, h: p.H }, scale: bmp.width / p.W } : undefined
  }

  function layerSrc(p: Slide, l: PlanLayerRecord): Src | undefined {
    const bmp = p.data?.assets.bitmap(l.assetId)
    return bmp ? { key: `a:${l.assetId}`, bmp, bounds: l.bounds, scale: l.scale } : undefined
  }

  const pressure = (p: Slide, message: string) =>
    sink.add({ code: 'plan-stage-pressure', slideIndex: p.index, message: `${message}; the slide holds its entry frame, static` })

  /** Uploads the entry frame and every layer of `p`; on a refusal the layers go and the entry frame alone stays. */
  function load(p: Slide): void {
    const entry = entrySrc(p)
    if (entry && !handleOf(p, entry)) { p.pressure = true; pressure(p, 'The entry frame does not fit the slideshow workspace') }
    if (!p.data || p.pressure) return
    for (const l of p.data.plan.layers) {
      if (isContainer(l)) continue
      const src = layerSrc(p, l)
      if (src && handleOf(p, src)) continue
      for (const key of [...p.ups.keys()]) if (key.startsWith('a:')) releaseUp(p, key)
      p.pressure = true
      pressure(p, src ? 'The layers do not fit the slideshow workspace' : `Layer asset ${l.assetId} is missing`)
      return
    }
  }

  function stateOf(p: Slide, cursor: SlideCursor | undefined, slideT: number, timeline: SlideTimeline | undefined): RenderState {
    if (!cursor || !timeline || cursor.mode === 'static') return EMPTY_STATE
    const snapshot = timeline.stateAt(cursor, Math.max(0, slideT))
    const raw = renderStateAt(snapshot, p.base, p.rng, sink, { w: p.W, h: p.H }, p.fan)
    const state = new Map<TargetKey, TargetState>()
    const baseOf = (k: TargetKey) => p.base.get(k) ?? (k.includes('/') ? p.base.get(k.slice(0, k.indexOf('/'))) : undefined)
    for (const [k, v] of raw) state.set(k, relativeTo(v, baseOf(k)))
    if (p.hidden.size) {
      // Hidden until one of its own entrance behaviours has begun.
      const entered = new Set<TargetKey>()
      for (const e of [...snapshot.active, ...snapshot.frozen]) {
        if (!('spid' in e.node.target)) continue
        const key = targetKey(e.node.target)
        if (p.entrances.get(key)?.has(e.node)) entered.add(key)
        for (const item of p.fan?.get(e.node)?.items ?? []) if (p.entrances.get(item.key)?.has(e.node)) entered.add(item.key)
      }
      for (const k of p.hidden) if (!entered.has(k)) state.set(k, HIDDEN)
    }
    return state
  }

  /** The live state, the targets of each fired `cut` trigger at their end state (§6a); `layers` triggers animate. */
  function liveState(p: Slide, cursor: SlideCursor | undefined, slideT: number, timeline: SlideTimeline | undefined): RenderState {
    const state = stateOf(p, cursor, slideT, timeline)
    if (!cursor || !p.data) return state
    const keys = new Set<TargetKey>()
    for (const trig of p.data.plan.triggers) {
      if (!trig.cut || !cursor.events.some(e => e.kind === 'trigger' && e.spid === trig.spid)) continue
      for (const k of p.triggerKeys.get(trig.spid) ?? []) keys.add(k)
    }
    if (!keys.size) return state
    if (!p.cutNoted) {
      p.cutNoted = true
      sink.add({ code: 'trigger-cut', slideIndex: p.index, message: 'A triggered effect the layers cannot prove jumps to its end state' })
    }
    const settled = stateOf(p, cursor, SETTLED_MS, timeline)
    const out = new Map(state)
    for (const k of keys) {
      const s = settled.get(k)
      if (s) out.set(k, s)
      else out.delete(k)
    }
    return out
  }

  /** `key`'s state composed through its container records, as `layerPics` draws it (§5.1 hit tests). */
  function effectiveState(p: Slide, state: RenderState, key: TargetKey): TargetState | undefined {
    const layers = p.data?.plan.layers ?? []
    let rec = layers.find(l => l.key === key)
    if (!rec) return state.get(key)
    const byId = new Map(layers.map(l => [l.id, l]))
    const chain: Array<PlanLayerRecord | PlanContainerRecord> = []
    for (let depth = 0; rec && depth <= SHOW_LAYER_DEPTH; depth++) {
      chain.unshift(rec)
      rec = rec.parent === undefined ? undefined : byId.get(rec.parent)
    }
    let s: TargetState | undefined
    for (const r of chain) s = compose(s, r.key === undefined ? undefined : state.get(r.key))
    return s
  }

  /** The layer draws of `p` at `state`, in paint order (the local `entriesOf` over the flat records). */
  function layerPics(p: Slide, state: RenderState): Pic[] {
    const out: Pic[] = []
    const groups = new Map<number, TargetState | null>()
    for (const l of p.data?.plan.layers ?? []) {
      const parent = l.parent === undefined ? undefined : groups.get(l.parent)
      // A record whose group is invisible is not drawn (nor are its own children).
      if (parent === null) { if (isContainer(l)) groups.set(l.id, null); continue }
      if (isContainer(l)) {
        const s = compose(parent, state.get(l.key))
        groups.set(l.id, s.visible ? s : null)
        continue
      }
      const src = layerSrc(p, l)
      if (!src) continue
      if (l.key === undefined) { out.push({ src, state: compose(parent, undefined), kind: `layer:${l.id}` }); continue }
      let s = compose(parent, state.get(l.key))
      if (!s.visible) continue
      if (l.part === 'range' && s.mask && !s.maskBox && l.maskBox) s = { ...s, maskBox: l.maskBox }
      const part = l.part === 'fill' || l.part === 'line' || l.part === 'text' ? l.part : l.part === 'range' && s.color?.text ? 'text' : undefined
      out.push({ src, state: s, ...(part ? { part } : {}), kind: `layer:${l.id}` })
    }
    return out
  }

  /** The track frame `i` of step `step` on `p`, as a picture source. */
  function trackSrc(p: Slide, step: number, i: number, bmp: ImageBitmapLike): Src {
    return { key: `t:${step}:${i}`, bmp, bounds: { x: 0, y: 0, w: p.W, h: p.H }, scale: bmp.width / p.W }
  }

  const entryPics = (p: Slide): Pic[] => { const e = entrySrc(p); return e ? [{ src: e, state: PLAIN, kind: e.key }] : [] }

  /**
   * What `p` shows at `cursor`/`slideT`: the entry frame (or still) under pressure or without a plan, a
   * composited step's track frame, or the layers. `undefined`: hold the previous picture.
   */
  function picture(p: Slide, cursor: SlideCursor | undefined, slideT: number, timeline: SlideTimeline | undefined, state: RenderState, live: boolean): Pic[] | undefined {
    const entryPic = entryPics(p)
    if (!p.data || p.pressure) return entryPic
    const k = cursor?.step ?? 0
    const step = k >= 1 ? p.data.plan.steps[k - 1] : undefined
    if (!cursor || !step || step.eligibility !== 'composited') {
      // A slide without raster layers (transition-only, or every step composited) shows its entry frame.
      return p.data.plan.layers.some(l => !isContainer(l)) ? layerPics(p, state) : entryPic
    }
    // A composited step: its track, never its layers.
    const events = cursor.events
    let t0 = 0
    for (let e = events.length - 1; e >= 0; e--) if (events[e]!.kind === 'next') { t0 = events[e]!.t; break }
    const playing = cursor.mode === 'live' && nexts(cursor) === 2 * k && !!timeline && slideT < timeline.busyUntil(cursor)
    const activation = `${k - 1}|${cursor.mode === 'live' ? t0 : 'settled'}`
    const track = p.data.track(k - 1)
    if (track === undefined || track === 'pending') {
      if (live && playing && !p.deferred.has(activation)) {
        p.deferred.add(activation)
        sink.add({ code: 'step-deferred', slideIndex: p.index, effect: `step:${k - 1}`, message: `The keyframe track of step ${k} is not here yet; the step cuts to its end` })
      }
      return undefined
    }
    const cut = p.deferred.has(activation)
    const i = playing && !cut ? Math.min(track.count - 1, Math.max(0, Math.floor(((slideT - t0) * TRACK_FPS) / 1000))) : track.count - 1
    if (live) track.want(i)
    const bmp = track.frame(i)
    if (!bmp) return undefined
    return [{ src: trackSrc(p, k - 1, i, bmp), state: PLAIN, kind: `track:${k - 1}:${i}` }]
  }

  /** Draws `pics` of `p` on the compositor; `false` when a source could not be uploaded. */
  function draw(p: Slide, pics: Pic[]): boolean {
    const list: LayerDraw[] = []
    const drawn: Pic[] = []
    let fresh: HeldFrame | undefined
    for (const pic of pics) {
      let src = pic.src
      const isTrack = src.key.startsWith('t:') || src.key.startsWith('mt:')
      if (isTrack) {
        // The displayed track frame is drawn from the stage's own copy (flattened and re-uploaded from it).
        if (heldFrame?.key === src.key) src = { ...src, bmp: heldFrame.canvas }
        else if ((fresh = copyFrame(src))) src = { ...src, bmp: fresh.canvas }
      }
      const h = handleOf(p, src)
      if (!h) { fresh?.lease.release(); return false }
      // One track frame uploaded at a time: the new one is in, the previous goes.
      if (isTrack && trackKey !== src.key) { if (trackKey) releaseUp(p, trackKey); trackKey = src.key }
      list.push({ handle: h, state: pic.state, ...(pic.part ? { part: pic.part } : {}) })
      drawn.push({ ...pic, src })
    }
    if (fresh) { heldFrame?.lease.release(); heldFrame = fresh }
    compositor().frame(list)
    kinds = pics.map(x => x.kind)
    p.held = drawn
    shown = { kind: 'pics' }
    return true
  }

  // -------------------------------------------------------------------
  // Flattening (outgoing frames, incoming checkpoints, Presenter "Next")
  // -------------------------------------------------------------------

  type Paint = (c2: Compositor, up: (src: Src) => LayerHandle | undefined, m: Affine) => boolean

  /** Paints on a transient Canvas 2D compositor into `canvas` (its viewport fits `p`); `false` when refused. */
  function composeWith(p: Slide, canvas: HTMLCanvasElement, paint: Paint): boolean {
    const c2 = createCanvas2DCompositor(canvas, factory, workspace, sink)
    try {
      const v = viewport(p, canvas.width, canvas.height)
      c2.setViewport(v)
      return paint(c2, (src) => upload(src, 'plan:compose', c2), v.matrix)
    } finally {
      c2.dispose()
    }
  }

  const paintPics = (pics: Pic[]): Paint => (c2, up) => {
    const list: LayerDraw[] = []
    for (const pic of pics) {
      const h = up(pic.src)
      if (!h) return false
      list.push({ handle: h, state: pic.state, ...(pic.part ? { part: pic.part } : {}) })
    }
    c2.frame(list)
    return true
  }

  const composeInto = (p: Slide, pics: Pic[], canvas: HTMLCanvasElement) => composeWith(p, canvas, paintPics(pics))

  /** A full-output frame painted by `paint`, in a leased canvas; `undefined` when refused. */
  function flattenWith(p: Slide, tag: string, paint: Paint): Flat | undefined {
    const cw = Math.max(1, output.width), ch = Math.max(1, output.height)
    const lease = workspace.lease(surfaceBytes(cw, ch), tag)
    if (!lease) return undefined
    const canvas = factory(cw, ch) as HTMLCanvasElement
    if (!composeWith(p, canvas, paint)) { lease.release(); return undefined }
    const f: Flat = { canvas, lease }
    if (!flatHandle(f)) { releaseFlat(f); return undefined }
    return f
  }

  const flatten = (p: Slide, pics: Pic[], tag: string) => flattenWith(p, tag, paintPics(pics))

  /** A flat frame as a source under the slide-px → device matrix `m` (the whole output). */
  const flatSrc = (f: Flat, m: Affine): Src => {
    const s = m[0]
    return { key: 'flat', bmp: f.canvas, bounds: { x: -m[4] / s, y: -m[5] / s, w: f.canvas.width / s, h: f.canvas.height / s }, scale: s }
  }

  /** `f` uploaded under the viewport in force (the whole output in slide px), re-uploaded when that changed. */
  function flatHandle(f: Flat): LayerHandle | undefined {
    const m = view.matrix
    const key = `${epoch}|${m.join(',')}`
    if (f.up?.key === key) return f.up.handle
    if (f.up && f.up.epoch === epoch) compositor().release(f.up.handle)
    f.up = undefined
    const h = upload(flatSrc(f, m), 'plan:frame')
    if (h) f.up = { handle: h, epoch, key }
    return h
  }

  /** What the last frame showed, flattened (the outgoing frame); `undefined` when refused or nothing was shown. */
  function capture(p: Slide): Flat | undefined {
    const r = running
    if (shown.kind === 'from') {
      // The held outgoing frame (a Morph waiting, or a composited Morph before its first frame): it is still that frame.
      if (waiting && outgoing) { const f = outgoing; outgoing = undefined; return f }
      if (r?.from) { const f = r.from; r.from = undefined; return f }
    }
    if (shown.kind === 'transition' && r && !r.morph) {
      const { progress, params } = shown
      return flattenWith(p, 'plan:from', (c2, up, m) => {
        const entry = r.to === 'entry' ? entrySrc(p) : undefined
        const to = r.to === 'entry' ? entry && up(entry) : up(flatSrc(r.to, m))
        const from = r.from ? up(flatSrc(r.from, m)) : to
        if (!to || !from) return false
        c2.frame([], { from, to, effect: r.spec.element, progress, params, rng: r.rng })
        return true
      })
    }
    if (shown.kind === 'morph' && r?.morph?.spec) {
      const { spec, srcOf } = r.morph
      const { progress } = shown
      return flattenWith(p, 'plan:from', (c2, up) => {
        const mine = new Map<LayerHandle, LayerHandle>()
        const swap = (side: MorphLayerFrame | undefined): MorphLayerFrame | undefined | false => {
          if (!side) return undefined
          let h = mine.get(side.handle)
          if (!h) { const src = srcOf.get(side.handle); h = src && up(src); if (!h) return false; mine.set(side.handle, h) }
          return { ...side, handle: h }
        }
        const pairs: MorphFrameItem[] = []
        for (const { from, to, ...rest } of spec.pairs) {
          const a = swap(from), b = swap(to)
          if (a === false || b === false) return false
          pairs.push({ ...rest, ...(a ? { from: a } : {}), ...(b ? { to: b } : {}) })
        }
        morphFrame({ pairs }, progress, c2)
        return true
      })
    }
    // What was on screen: the last picture (a still's, a held track frame, or the entry frame).
    const pics = p.held.length ? p.held : entryPics(p)
    return pics.length ? flatten(p, pics, 'plan:from') : undefined
  }

  function blackFrame(): Flat | undefined {
    const cw = Math.max(1, output.width), ch = Math.max(1, output.height)
    const lease = workspace.lease(surfaceBytes(cw, ch), 'plan:black')
    if (!lease) return undefined
    const canvas = factory(cw, ch) as HTMLCanvasElement
    const ctx = canvas.getContext('2d')
    if (!ctx) { lease.release(); return undefined }
    ctx.fillStyle = '#000'
    ctx.fillRect(0, 0, cw, ch)
    const f: Flat = { canvas, lease }
    if (!flatHandle(f)) { releaseFlat(f); return undefined }
    return f
  }

  function releaseFlat(f: Flat | undefined) {
    if (!f) return
    if (f.up && f.up.epoch === epoch) compositor().release(f.up.handle)
    f.up = undefined
    f.lease.release()
  }

  function releaseMorph(m: MorphRun) {
    if (m.epoch === epoch) for (const h of m.handles) compositor().release(h)
    m.handles = []
  }

  function endRunning() {
    if (!running) return
    releaseFlat(running.from)
    if (running.to !== 'entry') releaseFlat(running.to)
    if (running.morph) releaseMorph(running.morph)
    running = undefined
  }

  // -------------------------------------------------------------------
  // Morph (D11)
  // -------------------------------------------------------------------

  function morphRun(r: { record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack }): MorphRun | string {
    if (r.record.composited) return r.track ? { track: r.track, handles: [], srcOf: new Map(), epoch } : 'its keyframe track is not here'
    const handles: LayerHandle[] = []
    const srcOf = new Map<LayerHandle, Src>()
    const byAsset = new Map<number, LayerHandle>()
    const fail = (why: string) => { for (const h of handles) compositor().release(h); return why }
    const side = (s: NonNullable<MorphPlanRecord['items'][number]['from']>): MorphLayerFrame | string => {
      let h = byAsset.get(s.assetId)
      if (!h) {
        const bmp = r.assets.bitmap(s.assetId)
        if (!bmp) return `asset ${s.assetId} is missing`
        const src: Src = { key: `m:${s.assetId}`, bmp, bounds: s.bounds, scale: bmp.width / Math.max(1e-6, s.bounds.w) }
        h = upload(src, 'plan:morph')
        if (!h) return 'its layers do not fit the slideshow workspace'
        handles.push(h)
        srcOf.set(h, src)
        byAsset.set(s.assetId, h)
      }
      return { handle: h, geometry: s.geometry, ...(s.state ? { state: s.state } : {}), ...(s.part ? { part: s.part } : {}), ...(s.key ? { key: s.key } : {}) }
    }
    const pairs: MorphFrameItem[] = []
    for (const { from, to, ...rest } of r.record.items) {
      const a = from ? side(from) : undefined
      if (typeof a === 'string') return fail(a)
      const b = to ? side(to) : undefined
      if (typeof b === 'string') return fail(b)
      pairs.push({ ...rest, ...(a ? { from: a } : {}), ...(b ? { to: b } : {}) })
    }
    // `over` (the incoming watermark) is not drawn: marks are burned into every asset (the server builds with watermark off).
    return { spec: { pairs }, handles, srcOf, epoch }
  }

  /** Whether the Morph `r` has its initial content: every layered asset, or its track's first frame (asked for). */
  function morphReady(r: { record: MorphPlanRecord; assets: PlanAssets; track?: PlanTrack }): boolean {
    if (r.record.composited) {
      if (!r.track) return false
      r.track.want(0)
      return r.track.frame(0) !== undefined
    }
    return r.record.items.every(it => [it.from, it.to].every(side => !side || r.assets.bitmap(side.assetId) !== undefined))
  }

  /** Starts the transition of `w.p` now: the Morph `r`, or (`why`) the fade look-alike. */
  function startMorph(w: Waiting, r: Waiting['result'], why?: { code: string; message: string }) {
    waiting = undefined
    const p = w.p
    let run: MorphRun | string | undefined
    if (r && r !== 'refused') run = morphRun(r)
    if (why) sink.add({ code: why.code, slideIndex: p.index, effect: 'morph', message: why.message })
    else if (r === 'refused' || r === undefined) sink.add({ code: 'transition-approximate', slideIndex: p.index, effect: 'morph-fallback', message: 'Morph could not run (the server refused the pair); the fade look-alike plays' })
    else if (typeof run === 'string') sink.add({ code: 'morph-late', slideIndex: p.index, effect: 'morph', message: `Morph could not run (${run}); the fade look-alike plays` })
    const start = deps.showNow()
    if (run && typeof run !== 'string') {
      // A composited Morph keeps the outgoing frame for its frames not yet decoded; a layered one draws its own.
      const from = run.track ? outgoing : undefined
      if (!run.track) releaseFlat(outgoing)
      outgoing = undefined
      running = { spec: w.spec, three: false, ...(from ? { from } : {}), to: 'entry', start, duration: w.duration, rng: createRng((deps.seed ?? 0x5eed) ^ (p.index * 2654435761)), morph: run, fromKind: from ? 'flat' : 'none', toKind: 'entry' }
    } else startFlat(p, w.spec, 'lookalike', w.duration, start, w.cursor)
    origin = running ? start + w.duration : start
  }

  /** Starts a flat transition (from the outgoing frame to the incoming one) at `start`. */
  function startFlat(p: Slide, spec: TransitionSpec, route: 'shader' | 'lookalike' | 'morph', duration: number, start: number, cursor: SlideCursor | undefined) {
    let from = outgoing
    let fromKind: Running['fromKind'] = from ? 'flat' : 'none'
    outgoing = undefined
    if (!from && !shownBefore) {
      from = blackFrame()
      fromKind = from ? 'black' : 'none'
      if (!from) sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The black source frame did not fit the slideshow workspace; the first slide is a cut' })
    }
    // The incoming frame: the entry frame on entry, else the checkpoint entered at, composed.
    let to: Running['to'] | undefined
    let toKind: Running['toKind'] = 'entry'
    const entry = entrySrc(p)
    // A later checkpoint whose composited end frame is not decoded yet starts from the entry frame, as `frame` holds it.
    const pics = cursor && cursor.step > 0 ? picture(p, cursor, 0, p.timeline, liveState(p, cursor, 0, p.timeline), false) : undefined
    if (pics) {
      const f = flatten(p, pics, 'plan:to')
      if (f) { to = f; toKind = 'flat' }
    } else if (entry && handleOf(p, entry)) to = 'entry'
    if (!from || !to) {
      if (from || to) sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The transition frames did not fit the slideshow workspace; the slide change is a cut' })
      releaseFlat(from)
      if (to && to !== 'entry') releaseFlat(to)
      return
    }
    running = { spec, three: route === 'shader', from, to, start, duration, rng: createRng((deps.seed ?? 0x5eed) ^ (p.index * 2654435761)), fromKind, toKind }
  }

  // -------------------------------------------------------------------
  // The stage
  // -------------------------------------------------------------------

  const stage: PlanStage = {
    inspect() {
      return {
        current: current?.index,
        lastDraws: [...kinds],
        transition: running && { effect: running.three || running.morph || running.spec.element === 'random' ? running.spec.element : resolveTransition2D(running.spec).spec.element, duration: running.duration, from: running.fromKind, to: running.toKind },
        backend: manager.state,
      }
    },

    prepare() {
      // The source fetches the neighbours' bundles; layers are uploaded when a slide is entered.
    },

    flattenCurrent() {
      if (captured) return
      captured = true
      const p = current
      left = p && !disposed ? { index: p.index, step: lastStep } : undefined
      // What the audience saw is captured before anything it was drawn from is released.
      const f = p && !disposed ? capture(p) : undefined
      waiting = undefined
      endRunning()
      releaseFlat(outgoing)
      outgoing = f
      if (p && !disposed && !f && (p.held.length || entrySrc(p) || shown.kind !== 'pics')) {
        sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The outgoing frame did not fit the slideshow workspace; the slide change is a cut' })
      }
    },

    leave() {
      if (!captured) stage.flattenCurrent()
      captured = false
      if (current) releaseSlide(current)
      trackKey = undefined
      current = undefined
      origin = Number.POSITIVE_INFINITY
    },

    textTiming(index, meta, charge) {
      const have = texts.get(meta)
      if (have !== undefined) return have && have.fan.size ? { tree: have.tree, orderOf: have.orderOf } : undefined
      const plan = deps.data(index)?.plan
      const timing = meta.timing
      if (!plan || !plan.iterate.keys.length || !timing || 'refused' in timing) { texts.set(meta, null); return undefined }
      const x = planIterations(timing, plan.iterate, { charge, sink, slideIndex: index })
      texts.set(meta, x.fan.size ? x : null)
      return x.fan.size ? { tree: x.tree, orderOf: x.orderOf } : undefined
    },

    async enter(model, t) {
      if (disposed) return
      captured = false
      endRunning()
      waiting = undefined
      if (current) releaseSlide(current)
      const out = left
      left = undefined
      const p = slideOf(model.index, model.slide)
      p.timeline = model.timeline
      current = p
      lastStep = model.cursor.step
      const c = compositor()
      c.setSlideIndex(p.index)
      setView(p)
      load(p)
      let spec = model.slide.transition
      if (spec && deps.reduced) spec = reducedTransition(spec)
      const drawn = spec ? resolveTransition2D(spec) : undefined
      const route = spec ? dispatchTransition(spec, c, sink, p.index) : 'lookalike'
      const duration = spec && drawn && drawn.spec.element !== 'cut' && Number.isFinite(spec.durationMs) ? Math.max(0, spec.durationMs) : 0
      const start = Math.max(t, deps.showNow())
      if (spec && duration > 0 && route === 'morph') {
        if (!shownBefore || !out) {
          sink.add({ code: 'transition-approximate', slideIndex: p.index, effect: 'morph-fallback', message: `Morph could not run (${!shownBefore ? 'the first slide enters from black' : 'no slide was left'}); the fade look-alike plays` })
          startFlat(p, spec, 'lookalike', duration, start, model.cursor)
          origin = running ? start + duration : start
        } else {
          // D11: ask for the pair now; the outgoing frame holds until it lands or MORPH_WAIT_MS passes (polled per frame).
          const w: Waiting = { p, spec, duration, asked: deps.showNow(), cursor: model.cursor }
          waiting = w
          deps.morph(out.index, p.index, out.step).then(
            (r) => { if (waiting === w && current === p && !disposed) w.result = r ?? 'refused' },
            () => { /* late: the fade plays at the deadline */ },
          )
          origin = Number.POSITIVE_INFINITY
        }
      } else {
        if (spec && duration > 0) startFlat(p, spec, route, duration, start, model.cursor)
        origin = running ? start + duration : start
      }
      if (!running && !waiting) { releaseFlat(outgoing); outgoing = undefined }
      // Until its first frame, what is on screen is the outgoing frame (a leave now captures it).
      shown = waiting || running?.from ? { kind: 'from' } : { kind: 'pics' }
      shownBefore = true
    },

    slideTime(t) { return t - origin },

    frame(cursor, t) {
      if (disposed) return undefined
      const p = current
      const c = compositor()
      if (!p) {
        const h = outgoing && flatHandle(outgoing)
        if (h) { c.frame([{ handle: h, state: PLAIN }]); kinds = ['from']; shown = { kind: 'from' } }
        return undefined
      }
      if (cursor) lastStep = cursor.step
      if (waiting) {
        const w = waiting
        const now = deps.showNow()
        if (w.result === 'refused') startMorph(w, w.result)
        else if (w.result && morphReady(w.result)) startMorph(w, w.result)
        else if (now - w.asked >= MORPH_WAIT_MS) {
          startMorph(w, undefined, { code: 'morph-late', message: w.result ? `The Morph's ${w.result.record.composited ? 'first keyframe' : 'layers'} did not arrive within ${MORPH_WAIT_MS} ms; the fade look-alike plays` : `No Morph record within ${MORPH_WAIT_MS} ms; the fade look-alike plays` })
        } else {
          const h = outgoing && flatHandle(outgoing)
          c.frame(h ? [{ handle: h, state: PLAIN }] : [])
          kinds = ['from']
          shown = { kind: 'from' }
          return EMPTY_STATE
        }
      }
      if (running) {
        const progress = running.duration > 0 ? (t - running.start) / running.duration : 1
        if (progress < 1) {
          const m = running.morph
          if (m && m.epoch === epoch) {
            if (m.spec) { morphFrame(m.spec, Math.max(0, progress), c); kinds = ['morph']; shown = { kind: 'morph', progress: Math.max(0, progress) }; return EMPTY_STATE }
            const track = m.track!
            const i = Math.min(track.count - 1, Math.max(0, Math.floor(((t - running.start) * TRACK_FPS) / 1000)))
            track.want(i)
            const bmp = track.frame(i)
            if (bmp) {
              const pics = [{ src: { key: `mt:${i}`, bmp, bounds: { x: 0, y: 0, w: p.W, h: p.H }, scale: bmp.width / p.W }, state: PLAIN, kind: `morph-track:${i}` }]
              if (draw(p, pics)) return EMPTY_STATE
            }
            // Undecoded (D23): hold the previous frame, or the outgoing one before the first.
            if (p.held[0]?.src.key.startsWith('mt:') && draw(p, p.held)) return EMPTY_STATE
            const h = running.from && flatHandle(running.from)
            c.frame(h ? [{ handle: h, state: PLAIN }] : [])
            kinds = ['from']
            shown = { kind: 'from' }
            return EMPTY_STATE
          }
          const entry = running.to === 'entry' ? entrySrc(p) : undefined
          const to = running.to === 'entry' ? entry && handleOf(p, entry) : flatHandle(running.to)
          const from = running.from ? flatHandle(running.from) : to
          if (!m && to && from) {
            const aspect = p.W / p.H
            const params = running.three ? transition3DDraw(running.spec, progress, running.rng, aspect) : transitionDraw(running.spec, progress, running.rng, aspect)
            c.frame([], { from, to, effect: running.spec.element, progress, params, rng: running.rng })
            kinds = ['transition']
            shown = { kind: 'transition', progress, params }
            return EMPTY_STATE
          }
        }
        if (trackKey?.startsWith('mt:')) { releaseUp(p, trackKey); trackKey = undefined }
        endRunning()
      }
      const slideT = t - origin
      const state = liveState(p, cursor, slideT, p.timeline)
      const pics = picture(p, cursor, slideT, p.timeline, state, true)
      // Undefined (a missing track frame): hold the previous picture; nothing yet, the entry frame.
      const held = pics ?? (p.held.length ? p.held : entryPics(p))
      if (!draw(p, held)) {
        // A track frame that does not fit: hold what was drawn.
        if (p.held.length && held !== p.held) draw(p, p.held)
        else c.frame([])
      }
      return state
    },

    hitTest(clientX, clientY, cursor, t, on) {
      const p = current
      const plan = p?.data?.plan
      if (!p || !plan) return undefined
      const r = root.getBoundingClientRect()
      if (!(r.width > 0 && r.height > 0)) return undefined
      // Every surface canvas is letterboxed (object-fit: contain) to the backing store's aspect.
      const k = Math.min(r.width / output.width, r.height / output.height)
      const dx = (clientX - (r.left + (r.width - output.width * k) / 2)) / k
      const dy = (clientY - (r.top + (r.height - output.height * k) / 2)) / k
      const inv = invert(viewport(p).matrix)
      if (!inv) return undefined
      const [x, y] = apply(inv, dx, dy)
      const state = liveState(p, cursor, t - origin, p.timeline)
      // The topmost trigger (by its layer's z), its rects moved with its live state (§5.1).
      let best: { spid: string; z: number } | undefined
      for (const trig of plan.triggers) {
        if (!on.triggers.includes(trig.spid)) continue
        const key = targetKey({ spid: trig.spid })
        const s = effectiveState(p, state, key)
        if (s && (!s.visible || !(s.opacity > 0))) continue
        const m = invert(stateMatrix(s))
        if (!m) continue
        const [u, v] = apply(m, x, y)
        if (!trig.rects.some(rc => inRect(rc, u, v))) continue
        const z = p.zOf.get(key) ?? Number.NEGATIVE_INFINITY
        if (!best || z > best.z) best = { spid: trig.spid, z }
      }
      if (best) return { kind: 'trigger', spid: best.spid }
      // Link regions carry no target: they are static rects (D14). `withheld` consumes the click.
      for (let i = plan.linkRegions.length - 1; i >= 0; i--) {
        const region = plan.linkRegions[i]!
        if (!inRect(region.rect, x, y)) continue
        const a = region.action
        if (a.kind === 'withheld') return { kind: 'withheld' }
        return { kind: 'link', link: { spid: '', action: a.kind === 'slide' ? { kind: 'slide', index: a.index } : { kind: 'jump', verb: a.verb as 'nextslide' } } } satisfies HitResult
      }
      return undefined
    },

    composeAt(index, cursor, canvas) {
      if (disposed) return false
      const p = current?.index === index ? current : (() => {
        const data = deps.data(index)
        if (!data) return undefined
        // Another slide (Presenter "Next" across slides): its own fan and timeline, as the driver would schedule them.
        const timing = data.plan.show.timing
        const x = data.plan.iterate.keys.length && timing && !('refused' in timing) ? planIterations(timing, data.plan.iterate, { sink, slideIndex: index }) : undefined
        const text = x?.fan.size ? x : undefined
        const q = slideOf(index, undefined, text)
        q.timeline = createSlideTimeline(text?.tree ?? timing, { showAnimation: data.plan.options.showAnimation, reducedMotion: deps.reduced, ...(text ? { orderOf: text.orderOf } : {}) })
        return q
      })()
      if (!p) return false
      // A settled checkpoint: the layers at it, or a composited step's end frame.
      const state = liveState(p, cursor, SETTLED_MS, p.timeline)
      const pics = picture(p, cursor.mode === 'checkpoint' ? cursor : p.timeline?.checkpoint(Math.min(cursor.step, p.timeline.steps)) ?? cursor, SETTLED_MS, p.timeline, state, false)
      if (!pics) return false
      return composeInto(p, pics, canvas)
    },

    resize(css, dpr) {
      if (disposed) return
      const w = Math.max(1, Math.round(css.w * dpr)), h = Math.max(1, Math.round(css.h * dpr))
      if (w === output.width && h === output.height) return
      output.width = w
      output.height = h
      setView(current)
    },

    dispose() {
      if (disposed) return
      waiting = undefined
      endRunning()
      releaseFlat(outgoing)
      outgoing = undefined
      if (current) releaseSlide(current)
      current = undefined
      disposed = true
      manager.dispose()
      container.remove()
    },
  }
  setView(undefined)
  return stage
}
