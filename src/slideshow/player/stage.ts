/**
 * The per-slide engine (slideshow plan S1-20; spec 6.2 slide states, 6.3
 * rendering, transitions, resize and DPR, watermark; 7 residency and the
 * fallback ladder; checkpoints 2–8 integrated).
 *
 * The session (session.ts) owns navigation: which slide, its timeline and
 * its cursor. The stage owns everything drawn for it:
 *
 * - **Model → plan.** The slide's timing tree gives the animated targets,
 *   the colour-animated ones (split into fill/line/text part layers so each
 *   colour matrix applies to its part alone) and the initially hidden ones
 *   (spec 6.2 "initial": a target whose first behaviour in the tree is an
 *   entrance stays hidden until that behaviour begins); `buildRenderPlan`
 *   turns them, with the slide's build records, into ordered layers.
 * - **Residency.** `createResidency` keeps current, previous and next
 *   resident; the stage prices each rung from its plan (`cost`) and
 *   rasterises and uploads through the rung's sub-owner (`build`). One
 *   `ensure` per slide runs at a time. A rung change releases the superseded
 *   rung's handles once the new ones are uploaded.
 * - **Ladder.** `full` and `reduced` composite layers; `flattened` paints one
 *   full-slide layer per checkpoint (effects cut at their steps, transitions
 *   crossfade); `terminal` hides the compositor, releases every layer and
 *   paints the live state straight into the permanent reserve's output
 *   canvas (effects and slide changes are cuts). An upload refused while the
 *   context is lost is not pressure (`manager.state === 'lost'`): the live
 *   painter covers it and `rebuildLayers` re-uploads.
 * - **Handles.** Every handle is recorded with the backend epoch it was
 *   uploaded in; a backend switch (`rebuildLayers`) bumps the epoch, drops
 *   every record without releasing it (the old backend already released its
 *   leases) and re-uploads the current slide to the new compositor. Only
 *   handles of the current epoch are ever drawn (ids collide across
 *   backends).
 * - **Transitions.** The outgoing frame is flattened when its slide is left
 *   and keeps painting until the incoming slide is ready; then the incoming
 *   initial frame is flattened and the two are composited with
 *   `transitionDraw` (called before `frame`, with one rng for the whole
 *   effect). Reduced motion uses `reducedTransition` (never adds motion).
 *   S3-1: `dispatchTransition` sends a registered 3-D row on WebGL2 to its
 *   shader (`transition3DDraw` params); anything else draws `transitionDraw`,
 *   an approximated row diagnosed once.
 * - **Resize and DPR.** A new output takes its reserve before the old is
 *   released (`residency.resize`); the output and the viewport swap, then
 *   `commitResize`. Current, previous and next are re-rasterised layer by
 *   layer at the new scale, at most 8 per frame, while the old layers keep
 *   painting (their slide-px bounds scale with the viewport).
 */
import { acquireFonts, type FontLease } from '../../core/fonts/register'
import { decodeImage, decodeImageAsset, releaseDecodedImage, type ImageDecodeFn } from '../../core/images'
import type { WatermarkOptions } from '../../core/watermark'
import { layoutTextBody } from '../../drawing/text-layout'
import { createTextBodyMeasurer } from '../../drawing/text-paint'
import type { PptxShape, PptxSlide } from '../../pptx/types'
import { renderStateAt, type BaseGeometry } from '../engine/behaviours'
import type { SlideCursor, SlideTimeline } from '../engine/steps'
import { createBackendManager, createLivePainter, type BackendManager } from '../render/backend'
import { paintDirect } from '../render/direct'
import { buildRenderPlan, shapeKey, type ColourPart, type PlanLayer, type RenderPlan } from '../render/plan'
import { rasterLayer } from '../render/raster'
import { imageDrawSize, rasteriseSvg } from '../render/svg-raster'
import { imageIndices, paintedBounds, slideSize, textFilter, underlayOf } from '../render/scene'
import { reducedTransition, resolveTransition2D, transitionDraw } from '../render/transitions2d'
import { dispatchTransition, transition3DDraw } from '../render/three/framework'
import { createResidency, type LadderLevel, type Residency } from '../residency'
import { createRng } from '../rng'
import { textBases, textTimingFor, type TextTiming } from './text-builds'
import { targetKey } from '../target'
import type {
  Affine, AllocationOwner, CanvasFactory, Compositor, CompositorViewport, DiagnosticSink, LayerDraw, LayerHandle, Lease, MetadataHandle,
  RenderState, Rng, ShowSlide, SlideMetadataStore, SlideShowWorkspace, TargetKey, TargetState, TerminalReserve, TimeNode, TimingTree, TransitionSpec,
} from '../types'
import { createFrameLimits, LAYER_LIMIT, RERASTER_LIMIT, surfaceBytes, type Backing } from '../workspace'
import { apply, groupChildMatrix, invert, mul, nodeMatrix, stateMatrix } from './hit-test'

export interface StageDeps {
  doc: { slides: readonly PptxSlide[]; images?: unknown; fonts?: unknown }
  /** The audience surface root: the stage mounts its compositor container in it. */
  root: HTMLElement
  /** The terminal output canvas (the permanent reserve's surface). */
  output: HTMLCanvasElement
  workspace: SlideShowWorkspace
  terminal: TerminalReserve
  backing: Backing
  store: SlideMetadataStore
  sink: DiagnosticSink
  /** Reduced motion is in force (spec 4.5, after `slideShowMotion`). */
  reduced: boolean
  watermark?: WatermarkOptions
  /** Start on the Canvas 2D compositor (tests, the forced path). */
  force?: 'canvas2d'
  /** Monotonic ms (`performance.now`) for the context-restore deadline. */
  now(): number
  /**
   * The show clock. `enter` reads it once the slide is drawable, so a transition starts when its incoming
   * frame can be drawn, not when the change was requested (PR #18 review P1). Absent, `enter`'s `t` is used.
   */
  showNow?(): number
  schedule(cb: () => void, ms: number): unknown
  /** The slide's metadata through the session (so the target resolver sees its underlay). */
  metadata(index: number): Promise<MetadataHandle>
  /** Seed for the show's `rand` and random transitions. */
  seed?: number
}

export interface StageModel {
  index: number
  slide: ShowSlide
  timeline: SlideTimeline
  /** The cursor the slide is entered at (initial, or the last checkpoint when entered by Previous). */
  cursor: SlideCursor
}

export interface Stage {
  /** The shown slide is being left: its current frame becomes the outgoing transition frame. */
  leave(): void
  /**
   * Shows `model`; resolves once it is drawable and its transition (if any) has started. `t` is the show time
   * of the request; the transition starts at the show clock's reading once the slide is prepared.
   */
  enter(model: StageModel, t: number, neighbours: readonly number[]): Promise<void>
  /** Show time minus the moment the slide's timeline began (its transition's end); `-Infinity` before. */
  slideTime(t: number): number
  /** Draws one frame of the shown slide at show time `t`; returns the render state it drew (undefined between slides). */
  frame(cursor: SlideCursor | undefined, t: number): RenderState | undefined
  /** The shown slide's render state at `cursor`/`t` (hit tests, accessibility). */
  stateAt(cursor: SlideCursor | undefined, t: number): RenderState
  readonly plan: RenderPlan | undefined
  /** Client px → slide px of the shown slide. */
  toSlide(clientX: number, clientY: number): { x: number; y: number } | undefined
  /**
   * Additive (S2-7): shape `spid`'s base box on the shown slide in CSS px
   * relative to the surface root (letterbox included), for the video
   * overlay and the "Click to play" button; `undefined` without layout.
   */
  mediaBox?(spid: string): { left: number; top: number; width: number; height: number } | undefined
  /**
   * Additive (S2-8): the timing to schedule for slide `index` when it builds
   * text by word or letter: a working copy fanned out over the items, with
   * the source pre-order `orderOf` for `createSlideTimeline`. `charge`
   * reserves its bytes (the slide's metadata handle). `undefined`: schedule
   * the slide's own timing.
   */
  textTiming?(index: number, meta: ShowSlide, charge: (bytes: number, tag: string) => boolean): { tree: TimingTree; orderOf: ReadonlyMap<TimeNode, number> } | undefined
  /**
   * Additive (S2-5): what Presenter View needs to paint slide `index` at
   * `cursor` on `timeline` (a settled checkpoint for Next; none for the grid's
   * static state): the slide, its render state with the stage's own base
   * geometry and entrance-hidden targets, and the decoded images and fonts.
   * Only for the shown slide and its prepared neighbours; `undefined` for
   * any other slide.
   */
  previewOf?(index: number, cursor: SlideCursor | undefined, timeline: SlideTimeline | undefined): StagePreview | undefined
  /**
   * The text run (paragraph, run index) laid out at local px (u, v) of `shape`, whose box is `w` × `h`
   * (PR #18 review 3: hyperlinks on runs). `null` when no run is there; `undefined` when the text cannot be
   * located (vertical or warped text).
   */
  runAt(shape: PptxShape, u: number, v: number, w: number, h: number): { p: number; r: number } | null | undefined
  resize(css: { w: number; h: number }, dpr: number): void
  readonly level: LadderLevel | undefined
  readonly backend: string
  readonly backingStore: Backing
  /** The running transition, if any (what is drawn and for how long). */
  readonly transition: { effect: string; duration: number } | undefined
  /** @internal Per-slide residency and layer bookkeeping, and the last frame's draws (tests, the debug harness). */
  inspect(): StageInspection
  dispose(): void
}

/** Additive (S2-5): a slide's paint inputs for Presenter View (`Stage.previewOf`). */
export interface StagePreview {
  slide: PptxSlide
  state: RenderState
  images: ReadonlyArray<CanvasImageSource | undefined>
  resolveFont(family: string): string
  builds: ShowSlide['builds']
}

export interface StageInspection {
  current: number | undefined
  slides: Array<{ index: number; level: LadderLevel | undefined; scale: number | undefined; layers: number; uploaded: number; pending: number }>
  /** The plan-layer kind of each draw in the last composited frame (`target:<key>` for targets). */
  lastDraws: string[]
}

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
const PLAIN: TargetState = Object.freeze({ visible: true, opacity: 1, transform: IDENTITY })
const HIDDEN: TargetState = Object.freeze({ visible: false, opacity: 1, transform: IDENTITY })
const EMPTY_STATE: RenderState = new Map()
const NEXT: Record<Exclude<LadderLevel, 'terminal'>, LadderLevel> = { full: 'reduced', reduced: 'flattened', flattened: 'terminal' }

class Pressure extends Error { constructor() { super('pressure') } }

// ---------------------------------------------------------------------------
// Model → plan inputs
// ---------------------------------------------------------------------------

function isContainer(n: TimeNode): n is Extract<TimeNode, { children: TimeNode[] }> {
  return n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'
}

const COLOUR_ATTR = new Set(['fillcolor', 'stroke.color', 'style.color'])

/**
 * Decoded pixel size from a PNG, GIF, JPEG, WebP, BMP or TIFF header (S1-20 review I1, PR #18 review 2: the
 * decode is charged before it happens). `undefined` for anything else (SVG, unknown, truncated): the caller
 * does not decode it.
 */
export function imageSize(b: Uint8Array): { w: number; h: number } | undefined {
  const u16 = (o: number, le = false) => le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1]
  const u32 = (o: number) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0
  const ok = (w: number, h: number) => (w > 0 && h > 0 ? { w, h } : undefined)
  if (b.length >= 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return ok(u32(16), u32(20))
  if (b.length >= 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return ok(u16(6, true), u16(8, true))
  if (b.length >= 30 && b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) {
    const kind = String.fromCharCode(b[12], b[13], b[14], b[15])
    if (kind === 'VP8 ') return ok(u16(26, true) & 0x3fff, u16(28, true) & 0x3fff)
    if (kind === 'VP8L') return ok(1 + (((b[22] & 0x3f) << 8) | b[21]), 1 + (((b[24] & 0x0f) << 10) | (b[23] << 2) | ((b[22] & 0xc0) >> 6)))
    if (kind === 'VP8X') return ok(1 + (b[24] | (b[25] << 8) | (b[26] << 16)), 1 + (b[27] | (b[28] << 8) | (b[29] << 16)))
    return undefined
  }
  // BMP: BITMAPCOREHEADER (12 bytes) has 16-bit sizes at 18 and 20; the others have int32 at 18 and 22,
  // the height negative for a top-down bitmap.
  if (b.length >= 26 && b[0] === 0x42 && b[1] === 0x4d) {
    const i32 = (o: number) => b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)
    const header = i32(14)
    if (header === 12) return ok(u16(18, true), u16(20, true))
    if (header >= 40) return ok(Math.abs(i32(18)), Math.abs(i32(22)))
    return undefined
  }
  // TIFF: ImageWidth (256) and ImageLength (257) in IFD0, SHORT or LONG.
  if (b.length >= 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 42 && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 42))) {
    const le = b[0] === 0x49
    const t16 = (o: number) => u16(o, le)
    const t32 = (o: number) => (le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : u32(o))
    const ifd = t32(4)
    if (ifd + 2 > b.length) return undefined
    let w = 0, h = 0
    for (let k = 0, n = t16(ifd); k < n; k++) {
      const e = ifd + 2 + k * 12
      if (e + 12 > b.length) break
      const tag = t16(e), type = t16(e + 2)
      const value = type === 3 ? t16(e + 8) : type === 4 ? t32(e + 8) : 0
      if (tag === 256) w = value
      else if (tag === 257) h = value
    }
    return ok(w, h)
  }
  if (b.length >= 4 && b[0] === 0xff && b[1] === 0xd8) {
    // Walk the segments to the first start-of-frame (SOF0..SOF15 except DHT, JPG and DAC); bounded by the byte length.
    let o = 2
    while (o + 9 < b.length) {
      if (b[o] !== 0xff) return undefined
      const m = b[o + 1]
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return ok(u16(o + 7), u16(o + 5))
      o += 2 + u16(o + 2)
    }
  }
  return undefined
}

/**
 * Animated targets, colour-animated targets, and the initially hidden ones with their entrance behaviours.
 *
 * Initial hiding (spec 6.2 "initial"; S1-20 review I2): a target with any behaviour under an entrance preset is
 * hidden until one of those entrance behaviours begins, unless a non-entrance behaviour on it runs before its
 * first entrance in time. "Before in time" is read at click-step granularity: a non-entrance behaviour in a
 * strictly earlier main-sequence step (or before the first click) reveals the target from the start, the way
 * PowerPoint shows a shape whose first effect is an emphasis. Behaviours in the same step run concurrently, and
 * interactive sequences run only when triggered, so neither reveals the target before its entrance.
 */
export function timingTargets(slide: ShowSlide): {
  animated: Set<TargetKey>; colour: Set<TargetKey>; hidden: Set<TargetKey>; entrances: Map<TargetKey, Set<TimeNode>>
} {
  const animated = new Set<TargetKey>(), colour = new Set<TargetKey>(), hidden = new Set<TargetKey>()
  const entrances = new Map<TargetKey, Set<TimeNode>>()
  const timing = slide.timing
  if (!timing || 'refused' in timing) return { animated, colour, hidden, entrances }
  const firstEntrance = new Map<TargetKey, number>()
  const firstOther = new Map<TargetKey, number>()
  const walk = (n: TimeNode, cls: string | undefined, step: number, depth: number): void => {
    if (depth > 128) return
    const here = n.ctn.presetClass ?? cls
    if (isContainer(n)) {
      // The main sequence's children are its steps, in order; everything off it (interactive) is "never before".
      const main = n.kind === 'seq' && n.ctn.nodeType === 'mainSeq'
      n.children.forEach((c, i) => walk(c, here, main ? i : n.ctn.nodeType === 'interactiveSeq' ? Infinity : step, depth + 1))
      return
    }
    if (!('spid' in n.target)) return
    const key = targetKey(n.target)
    animated.add(key)
    if (n.kind === 'animClr' || n.attrNames.some(a => COLOUR_ATTR.has(a))) colour.add(key)
    const first = here === 'entr' ? firstEntrance : firstOther
    first.set(key, Math.min(first.get(key) ?? Infinity, step))
    if (here === 'entr') {
      let set = entrances.get(key)
      if (!set) entrances.set(key, set = new Set())
      set.add(n)
    }
  }
  walk(timing.root, undefined, Infinity, 0)
  for (const [key, at] of firstEntrance) if (!((firstOther.get(key) ?? Infinity) < at)) hidden.add(key)
  for (const key of [...entrances.keys()]) if (!hidden.has(key)) entrances.delete(key)
  return { animated, colour, hidden, entrances }
}

/** Base geometry (slide px, group transforms composed) per shape key, groups included; ranges and parts share their shape's. */
export function baseGeometry(slide: PptxSlide): Map<TargetKey, BaseGeometry> {
  const out = new Map<TargetKey, BaseGeometry>()
  const walk = (nodes: readonly PptxShape[], parent: Affine, rot: number, depth: number) => {
    if (depth > 64) return
    for (const shape of nodes) {
      const local = nodeMatrix(shape, parent)
      if (!local) continue
      const r = rot + (shape.rotationDeg ?? 0)
      const key = shapeKey(shape)
      if (key !== undefined) out.set(key, geometryOf(local.m, local.w, local.h, r))
      if (shape.group) {
        const child = groupChildMatrix(shape, local.m)
        if (child) walk(shape.children ?? [], child, r, depth + 1)
      }
    }
  }
  walk(slide.shapes, IDENTITY, 0, 0)
  return out
}

function geometryOf(m: Affine, w: number, h: number, rot = 0): BaseGeometry {
  const [cx, cy] = apply(m, w / 2, h / 2)
  const sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3])
  const ww = w * sx, hh = h * sy
  return { x: cx - ww / 2, y: cy - hh / 2, w: ww, h: hh, rot: Number.isFinite(rot) ? rot : 0 }
}

/** `transform ∘ rotate(−base.rot about the centre)`: layers are painted at their base rotation already. */
function relativeTo(state: TargetState, g: BaseGeometry | undefined): TargetState {
  if (!g || !g.rot) return state
  const a = -g.rot * Math.PI / 180, c = Math.cos(a), s = Math.sin(a)
  const cx = g.x + g.w / 2, cy = g.y + g.h / 2
  const back: Affine = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
  return { ...state, transform: mul(state.transform, back) }
}

/** The group state composed over a sub-layer's own. */
function compose(parent: TargetState | undefined, own: TargetState | undefined): TargetState {
  if (!parent) return own ?? PLAIN
  const p = stateMatrix(parent), o = stateMatrix(own)
  return {
    visible: parent.visible && (own?.visible ?? true),
    opacity: parent.opacity * (own?.opacity ?? 1),
    transform: mul(p, o),
    ...(own?.mask ? { mask: own.mask } : {}),
    ...(own?.color ?? parent.color ? { color: own?.color ?? parent.color } : {}),
  }
}

/** Every plan layer that rasterises on its own, in paint order (sub-layers included; a group-with-sub paints through them). */
function rasterLayers(plan: RenderPlan, watermark: boolean): PlanLayer[] {
  const out: PlanLayer[] = []
  const walk = (layers: readonly PlanLayer[]) => {
    for (const l of layers) {
      if (l.kind === 'watermark') { if (watermark) out.push(l); continue }
      if (l.kind === 'target' && l.sub) { walk(l.sub); continue }
      out.push(l)
    }
  }
  walk(plan.layers)
  return out
}

/** Slide time far past every finite effect: the settled state a flattened frame shows (PR #18 review 8). */
const SETTLED_MS = 1e9

function stateKey(state: RenderState): string {
  let s = ''
  for (const [k, v] of state) s += `${k}:${v.visible ? 1 : 0}:${v.opacity.toFixed(3)}:${v.transform.map(n => n.toFixed(2)).join(',')}:${v.motion ? `${v.motion.dx.toFixed(1)},${v.motion.dy.toFixed(1)}` : ''};`
  return s
}

// ---------------------------------------------------------------------------
// The stage
// ---------------------------------------------------------------------------

interface Built {
  level: Exclude<LadderLevel, 'terminal'>
  epoch: number
  /** Device px per slide px the layers were rasterised at. */
  scale: number
  handles: Map<PlanLayer, LayerHandle>
  /** Layers to (re)rasterise at `scale` (resize, a backend switch). */
  pending: PlanLayer[]
  /** Flattened rung: the current checkpoint's full-slide frame. */
  /** `cursor`: the cursor the frame was settled for (it is re-checked only when the cursor changes). */
  flat?: { key: string; handle: LayerHandle; cursor?: SlideCursor }
}

interface Prepared {
  index: number
  slide: PptxSlide
  meta: ShowSlide
  plan: RenderPlan
  layers: PlanLayer[]
  base: Map<TargetKey, BaseGeometry>
  hidden: Set<TargetKey>
  /** The entrance behaviours of each hidden target: it shows once one of them has begun. */
  entrances: Map<TargetKey, Set<TimeNode>>
  /** The show's sink, scoped to this slide: behaviour diagnostics carry no slide of their own (S1-22). */
  sink: DiagnosticSink
  images: Array<CanvasImageSource | undefined>
  /** The workspace charge of every decoded image above (released with them). */
  imageLeases: Lease[]
  ready: Promise<void>
  /** The stage's own metadata reference (neighbours); the shown slide's is the session's. */
  handle?: MetadataHandle
  built?: Built
  level?: LadderLevel
  /** One `ensure` at a time per slide. */
  chain: Promise<unknown>
  dead: boolean
}

interface Frame { handle: LayerHandle; epoch: number; lease?: Lease }

interface Running {
  spec: TransitionSpec
  /** S3-1: dispatched to its registered 3-D shader (`transition3DDraw` params), else `transitionDraw`. */
  three: boolean
  from?: Frame
  to: Frame
  start: number
  duration: number
  rng: Rng
}

export function createStage(deps: StageDeps): Stage {
  const { root, output, workspace, terminal, sink, store } = deps
  const doc = root.ownerDocument
  const factory: CanvasFactory = (w, h) => {
    const c = doc.createElement('canvas')
    c.width = Math.max(1, Math.ceil(w)); c.height = Math.max(1, Math.ceil(h))
    return c
  }
  const container = doc.createElement('div')
  container.setAttribute('data-officeview-show-stage', '')
  container.hidden = true
  output.after(container)

  const limits = createFrameLimits(sink)
  let backing: Backing = deps.backing
  let epoch = 0
  let disposed = false
  let fonts: Promise<FontLease | undefined> | undefined
  let resolveFont: (family: string) => string = (f) => f
  const lease = () => fonts ??= acquireFonts(deps.doc as never).then(l => { resolveFont = l.resolve; return l }, () => undefined)

  const prepared = new Map<number, Prepared>()
  let current: Prepared | undefined
  let origin = Number.POSITIVE_INFINITY
  let outgoing: Frame | undefined
  let running: Running | undefined
  let lastDirect = ''
  let rngs = new Map<number, Rng>()
  let liveState: RenderState = EMPTY_STATE
  /** Whether any slide has been entered yet (the first one transitions from black). */
  let shownBefore = false
  let kinds: string[] = []
  /** A 1 × 1 canvas whose context measures text for `runAt`. */
  let measurer: ReturnType<CanvasFactory> | undefined

  // -------------------------------------------------------------------
  // Viewport
  // -------------------------------------------------------------------

  const slidePx = (p: Prepared | undefined) => {
    const slide = p?.slide ?? deps.doc.slides[0]
    // A deck without slides shows only the end-of-show slide: any valid size will do for its empty viewport.
    return slide ? slideSize(slide) : { w: 960, h: 540 }
  }
  function viewport(p: Prepared | undefined): CompositorViewport & { scale: number } {
    const { w, h } = slidePx(p)
    const s = Math.min(backing.w / w, backing.h / h)
    return { matrix: [s, 0, 0, s, (backing.w - w * s) / 2, (backing.h - h * s) / 2], deviceWidth: backing.w, deviceHeight: backing.h, scale: s }
  }
  const layerScale = (level: Exclude<LadderLevel, 'terminal'>, p: Prepared) => {
    const s = viewport(p).scale
    return level === 'full' ? s : s / Math.max(1, backing.dpr)
  }

  // -------------------------------------------------------------------
  // Backend
  // -------------------------------------------------------------------

  const manager: BackendManager = createBackendManager(container, {
    // The Canvas 2D fallback composites on the terminal output, which the permanent reserve pays for (PR #18 review P2).
    output,
    now: deps.now,
    schedule: deps.schedule,
    paintLive: createLivePainter(terminal, () => current && {
      slide: current.slide, state: liveState,
      assets: { images: current.images, resolveFont, sink, builds: current.meta.builds, ...(deps.watermark ? { watermark: deps.watermark } : {}) },
    }),
    rebuildLayers: async (c) => {
      // A backend switch: every recorded handle belonged to the old backend, which released its leases.
      epoch++
      outgoing = undefined
      running = undefined
      for (const p of prepared.values()) {
        if (!p.built) continue
        p.built.handles.clear()
        p.built.flat = undefined
        p.built.epoch = epoch
        p.built.pending = p.built.level === 'flattened' ? [] : [...p.layers]
      }
      if (current?.built) {
        c.setSlideIndex(current.index)
        c.setViewport(viewport(current))
        drain(current, Infinity, c)
      }
    },
    workspace, terminal, sink,
    ...(deps.force ? { force: deps.force } : {}),
  })
  const compositor = (): Compositor => manager.current
  compositor().setViewport(viewport(undefined))

  // -------------------------------------------------------------------
  // Residency
  // -------------------------------------------------------------------

  const asWorkspace = (owner: AllocationOwner): SlideShowWorkspace => ({
    limit: Number.POSITIVE_INFINITY,
    lease: (bytes, tag) => owner.lease(bytes, tag),
    reserveTerminal: () => undefined,
    used: () => 0,
  })

  function price(slide: number, level: Exclude<LadderLevel, 'terminal'>): number {
    const p = prepared.get(slide)
    const { w, h } = slidePx(p)
    const frameAt = (k: number) => surfaceBytes(w * k, h * k)
    if (!p) return 3 * surfaceBytes(backing.w, backing.h)
    const s = layerScale(level === 'flattened' ? 'reduced' : level, p)
    // Transition frames (incoming initial) are drawn from the slide's own owner: a frame, its upload peak.
    const transition = 2 * frameAt(s)
    if (level === 'flattened') return 2 * frameAt(s) + transition
    let sum = 0, peak = 0
    for (const layer of p.layers) {
      let box = { minX: 0, minY: 0, maxX: w, maxY: h }
      if (layer.kind === 'target' || (layer.kind === 'static' && !layer.background)) {
        const b = paintedBounds(layer.nodes, p.slide, resolveFont)
        if (!b) continue
        box = { minX: Math.max(b.minX, -w), minY: Math.max(b.minY, -h), maxX: Math.min(b.maxX, 2 * w), maxY: Math.min(b.maxY, 2 * h) }
      }
      // Text can overflow its box: a quarter more, plus one device px of grid rounding on each side.
      const bytes = surfaceBytes((box.maxX - box.minX) * s * 1.25 + 2, (box.maxY - box.minY) * s * 1.25 + 2)
      sum += bytes
      peak = Math.max(peak, bytes)
    }
    return sum + peak + transition
  }

  const residency: Residency = createResidency(workspace, store, terminal, {
    sink,
    backing,
    cost: (slide, level) => price(slide, level),
    build: async (slide, level, owner) => build(slide, level, owner),
    // An evicted neighbour's uploaded layers and image leases go with it (PR #18 review 6).
    teardown: (slide) => { const p = prepared.get(slide); if (p && p !== current) drop(p) },
  })

  /** Rasterises one layer through `owner` and uploads it to `c`. `undefined`: nothing to draw (empty, or deferred while lost). */
  function rasterOne(p: Prepared, layer: PlanLayer, scale: number, owner: AllocationOwner, c: Compositor): LayerHandle | undefined {
    const r = rasterLayer(layer, p.slide, {
      scale, images: p.images, resolveFont, canvas: factory, workspace: asWorkspace(owner),
      maxTextureSize: c.maxTextureSize, sink, ...(deps.watermark ? { watermark: deps.watermark } : {}),
    })
    if ('refused' in r) {
      if (r.refused === 'empty') return undefined
      throw new Pressure()
    }
    const up = c.upload(r, r.lease)
    if ('refused' in up) {
      r.lease.release()
      // Lost context: not pressure; the live painter covers it and rebuildLayers re-uploads.
      if (manager.state === 'lost') return undefined
      throw new Pressure()
    }
    return up
  }

  function releaseBuilt(b: Built | undefined) {
    if (!b) return
    if (b.epoch === epoch) {
      for (const h of b.handles.values()) compositor().release(h)
      if (b.flat) compositor().release(b.flat.handle)
    }
    b.handles.clear()
    b.flat = undefined
  }

  /** Residency's builder: rasterise every layer through the rung's owner. A same-rung rebuild (resize) is incremental. */
  function build(slide: number, level: Exclude<LadderLevel, 'terminal'>, owner: AllocationOwner): void {
    const p = prepared.get(slide)
    if (!p || p.dead || disposed) throw new Error('cancelled')
    const scale = layerScale(level, p)
    const old = p.built
    if (old && old.level === level && old.epoch === epoch) {
      // Same rung (a resize): re-rasterise layer by layer, old ones painting until replaced.
      old.scale = scale
      old.flat = old.flat && level === 'flattened' ? old.flat : old.flat
      if (level === 'flattened') { if (old.flat) { compositor().release(old.flat.handle); old.flat = undefined } }
      else old.pending = [...p.layers]
      owners.set(slide, owner)
      return
    }
    const next: Built = { level, epoch, scale, handles: new Map(), pending: [] }
    owners.set(slide, owner)
    if (level !== 'flattened') {
      if (p.layers.length > LAYER_LIMIT) {
        limits.check(slide, { layers: p.layers.length, reRasters: 0 })
        throw new Pressure()
      }
      const c = compositor()
      if (current === p) c.setSlideIndex(slide)
      try {
        for (const layer of p.layers) {
          const h = rasterOne(p, layer, scale, owner, c)
          if (h) next.handles.set(layer, h)
          else if (manager.state === 'lost') next.pending.push(layer)
        }
      } catch (e) {
        for (const h of next.handles.values()) c.release(h)
        throw e
      }
    }
    // The superseded rung's handles go once the new ones are in (their leases refund the retired owner).
    releaseBuilt(old)
    p.built = next
  }
  const owners = new Map<number, AllocationOwner>()

  /** Re-rasterises up to `budget` pending layers of `p` (resize, backend switch). */
  function drain(p: Prepared, budget: number, c: Compositor = compositor()): number {
    const b = p.built
    const owner = owners.get(p.index)
    if (!b || !owner || b.level === 'flattened' || b.epoch !== epoch) return 0
    let done = 0
    while (b.pending.length > 0 && done < budget) {
      const layer = b.pending[0]
      let h: LayerHandle | undefined
      try {
        h = rasterOne(p, layer, b.scale, owner, c)
      } catch (e) {
        if (!(e instanceof Pressure)) throw e
        // No room for both: let the old one go first, then try once more.
        const old = b.handles.get(layer)
        if (!old) { void stepDown(p); return done }
        c.release(old)
        b.handles.delete(layer)
        try { h = rasterOne(p, layer, b.scale, owner, c) } catch { void stepDown(p); return done }
      }
      if (h === undefined && manager.state === 'lost') return done
      b.pending.shift()
      const old = b.handles.get(layer)
      if (old) c.release(old)
      if (h) b.handles.set(layer, h)
      else b.handles.delete(layer)
      done++
    }
    return done
  }

  function ensure(p: Prepared, level: LadderLevel): Promise<LadderLevel | undefined> {
    const run = p.chain.then(async () => {
      if (p.dead || disposed) return undefined
      let got: LadderLevel
      try {
        got = await residency.ensure(p.index, level)
      } catch (e) {
        if (e instanceof Pressure && level !== 'terminal') {
          const next = NEXT[level as Exclude<LadderLevel, 'terminal'>]
          if (current === p) sink.add({ code: `ladder-${next}`, slideIndex: p.index, message: next === 'terminal' ? 'the slide is painted directly into the terminal reserve; effects and slide changes become cuts' : next === 'reduced' ? 'the slide does not fit: its layers are rebuilt at reduced resolution' : 'the slide does not fit: it is flattened into one layer per checkpoint, with cuts and a crossfade' })
          p.chain = Promise.resolve()
          return ensure(p, next)
        }
        // Left before it finished, or a failed read: nothing is shown for it.
        return undefined
      }
      if (p.dead) return undefined
      if (got === 'terminal') { releaseBuilt(p.built); p.built = undefined }
      if ((got === 'terminal' || got === 'flattened') && textTimings.get(p.meta)?.splits.size) {
        sink.add({ code: 'text-build-whole', slideIndex: p.index, effect: got, message: `Text by word or letter: at the ${got} rung the split shapes paint whole, appearing with their first item` })
      }
      p.level = got
      return got
    })
    p.chain = run.catch(() => undefined)
    return run
  }

  function stepDown(p: Prepared): Promise<unknown> {
    const at = p.level
    if (!at || at === 'terminal') return Promise.resolve()
    const next = NEXT[at]
    if (current === p) sink.add({ code: `ladder-${next}`, slideIndex: p.index, message: `the slide no longer fits at the ${at} rung; it steps down to ${next}` })
    return ensure(p, next)
  }

  // -------------------------------------------------------------------
  // Slides
  // -------------------------------------------------------------------

  /** Text builds per slide metadata (S2-8): one working copy per `ShowSlide`, shared by the timeline and the plan. */
  const textTimings = new WeakMap<ShowSlide, TextTiming | null>()
  function textTimingOf(index: number, meta: ShowSlide, charge: ((bytes: number, tag: string) => boolean) | undefined): TextTiming | undefined {
    const have = textTimings.get(meta)
    if (have !== undefined) return have ?? undefined
    const timing = meta.timing
    const slide = deps.doc.slides[index]
    if (!timing || 'refused' in timing || !slide) { textTimings.set(meta, null); return undefined }
    // Layers the items may take: the layer bound less what the slide plans without them (and a margin per split).
    const t = timingTargets(meta)
    const plain = buildRenderPlan(slide, t.animated, new Map(), meta.builds).layers.length
    const tt = textTimingFor(slide, timing, { sink, slideIndex: index, charge: charge ?? (() => false), layers: LAYER_LIMIT - plain - 4 })
    textTimings.set(meta, tt ?? null)
    return tt
  }

  /** Timing targets of the scheduled tree, with every text item hidden and entered as its iterated target is. */
  function targetsOf(meta: ShowSlide, tt: TextTiming | undefined) {
    const t = timingTargets(tt ? { ...meta, timing: tt.tree } : meta)
    if (tt) {
      for (const [target, keys] of tt.items) {
        for (const k of keys) {
          t.animated.add(k)
          if (t.hidden.has(target)) {
            t.hidden.add(k)
            t.entrances.set(k, t.entrances.get(target) ?? new Set())
          }
        }
      }
    }
    return t
  }

  function prepare(index: number, meta: ShowSlide, handle?: MetadataHandle): Prepared {
    const have = prepared.get(index)
    if (have && !have.dead) {
      if (handle && !have.handle) have.handle = handle
      else handle?.release()
      if (have.meta !== meta) {
        // Another metadata object for the same slide (a re-read): its text timing is the one now scheduled.
        const tt = textTimingOf(index, meta, have.handle ? (b, t) => have.handle!.attach(b, t) : undefined)
        if (tt || textTimings.get(have.meta)) {
          const t = targetsOf(meta, tt)
          have.hidden = t.hidden
          have.entrances = t.entrances
        }
      }
      have.meta = meta
      return have
    }
    const slide = deps.doc.slides[index]
    const tt = textTimingOf(index, meta, handle ? (b, t) => handle.attach(b, t) : undefined)
    const { animated, colour, hidden, entrances } = targetsOf(meta, tt)
    // Any colour animation splits its shape into fill, line and text layers (a two-part set forces the split).
    const parts = new Map<TargetKey, ReadonlySet<ColourPart>>([...colour].map(k => [k, new Set<ColourPart>(['fill', 'line'])]))
    const plan = buildRenderPlan(slide, animated, parts, meta.builds, tt?.splits)
    const images: Array<CanvasImageSource | undefined> = []
    const p: Prepared = {
      index, slide, meta, plan, layers: rasterLayers(plan, !!deps.watermark), base: baseGeometry(slide), hidden, entrances, images,
      sink: { add: (d) => sink.add(d.slideIndex === undefined ? { ...d, slideIndex: index } : d), list: () => sink.list() },
      ready: Promise.resolve(), handle, chain: Promise.resolve(), dead: false, imageLeases: [],
    }
    const used = [...new Set(imageIndices([...underlayOf(slide), ...(slide.shapes ?? [])]))]
    const assets = (deps.doc as { images?: Array<unknown> }).images
    // Decoded bitmaps are charged to the workspace for as long as the slide holds them (S1-20 review I1), and
    // every charge is taken before its decoder allocates (PR #18 review 2): the representation decodeImageAsset
    // selects is sized first. A raster is sized from its header; an SVG is drawn at the size it is painted at
    // (the full rung's device px), whatever its declared intrinsic size. A raster whose header gives no size is
    // not decoded.
    const refuse = (i: number, w: number, h: number) => sink.add({
      code: 'image-workspace', slideIndex: index, effect: `image:${i}`,
      message: `Image ${i} (${w}×${h} px decoded) does not fit the slideshow workspace; painted as the missing-image placeholder`,
    })
    const nodes = [...underlayOf(slide), ...(slide.shapes ?? [])]
    p.ready = Promise.all([lease(), ...used.map(async (i) => {
      const asset = assets?.[i] as object | undefined
      if (!asset) return
      const tag = `image:${index}:${i}`
      let held: Lease | undefined
      let refused: { w: number; h: number } | 'unsized' | undefined
      const take = (w: number, h: number) => {
        held = workspace.lease(surfaceBytes(w, h), tag)
        if (!held) refused = { w, h }
        return !!held
      }
      const unhold = () => { held?.release(); held = undefined }
      // One call per representation tried (an accepted SVG, then its raster fallback): reserve, then decode.
      const decode: ImageDecodeFn = async (bytes, mime) => {
        unhold()
        if (mime === 'image/svg+xml') {
          const drawn = imageDrawSize(nodes, i)
          const v = viewport(p)
          const cap = 2 * Math.max(backing.w, backing.h)
          const w = Math.max(1, Math.min(cap, Math.ceil((drawn?.w ?? slidePx(p).w) * v.scale)))
          const h = Math.max(1, Math.min(cap, Math.ceil((drawn?.h ?? slidePx(p).h) * v.scale)))
          if (!take(w, h)) return undefined
          const canvas = await rasteriseSvg(bytes, w, h, factory)
          if (!canvas) unhold()
          return canvas
        }
        const known = imageSize(bytes)
        if (!known) { refused = 'unsized'; return undefined }
        if (!take(known.w, known.h)) return undefined
        const img = await decodeImage(bytes, mime).catch(() => undefined)
        if (!img) unhold()
        return img
      }
      const img = await decodeImageAsset(asset as never, decode).catch(() => undefined)
      if (img) {
        // Always reconcile with the decoded dimensions before keeping the bitmap: grow (or trim) the charge,
        // or close the bitmap and refuse when the workspace cannot hold it (a header that understated its size).
        const d = img as { width?: number; height?: number }
        const w = Number(d.width) || 1, h = Number(d.height) || 1
        const bytes = surfaceBytes(w, h)
        const l = held as Lease | undefined
        if (!l) held = workspace.lease(bytes, tag)
        else if (bytes > l.bytes && !l.grow(bytes - l.bytes)) unhold()
        else if (bytes < l.bytes) l.shrink(l.bytes - bytes)
        if (!held) { releaseDecodedImage(img); refuse(i, w, h); return }
      } else {
        unhold()
        if (refused === 'unsized') sink.add({ code: 'image-workspace', slideIndex: index, effect: `image:${i}`, message: `Image ${i} has no readable size to reserve before decoding; painted as the missing-image placeholder` })
        else if (refused) refuse(i, refused.w, refused.h)
        return
      }
      if (p.dead) { releaseDecodedImage(img); unhold(); return }
      images[i] = img
      if (held) p.imageLeases.push(held)
    })]).then(() => {
      // Item geometry (S2-8) needs the real fonts: laid out once they are in, on a charged measuring context.
      if (!tt || tt.splits.size === 0 || p.dead) return
      const l = workspace.lease(4, `text-bases:${index}`)
      if (!l) { sink.add({ code: 'text-build-refused', slideIndex: index, message: 'Text by word or letter: no room to measure the text; the items move as their shape does' }); return }
      try {
        const ctx = factory(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D | null
        if (ctx) for (const [k, g] of textBases(slide, tt.splits, p.base, ctx, resolveFont)) p.base.set(k, g)
      } finally {
        l.release()
      }
    })
    prepared.set(index, p)
    return p
  }

  function drop(p: Prepared) {
    if (p.dead) return
    p.dead = true
    releaseBuilt(p.built)
    p.built = undefined
    owners.delete(p.index)
    rngs.delete(p.index)
    p.handle?.release()
    for (const img of new Set(p.images)) releaseDecodedImage(img)
    for (const l of p.imageLeases.splice(0)) l.release()
    if (prepared.get(p.index) === p) prepared.delete(p.index)
  }

  const rngOf = (i: number) => {
    let r = rngs.get(i)
    if (!r) rngs.set(i, r = createRng((deps.seed ?? 0x5eed) + i))
    return r
  }

  function stateOf(p: Prepared, cursor: SlideCursor | undefined, slideT: number, timeline: SlideTimeline | undefined): RenderState {
    if (!cursor || !timeline || cursor.mode === 'static') return EMPTY_STATE
    const snapshot = timeline.stateAt(cursor, Math.max(0, slideT))
    const fan = textTimings.get(p.meta)?.fan
    const raw = renderStateAt(snapshot, p.base, rngOf(p.index), p.sink, slidePx(p), fan)
    const state = new Map<TargetKey, TargetState>()
    // Text items without their own geometry share their shape's.
    const baseOf = (k: TargetKey) => p.base.get(k) ?? (k.includes('/') ? p.base.get(k.slice(0, k.indexOf('/'))) : undefined)
    for (const [k, v] of raw) state.set(k, relativeTo(v, baseOf(k)))
    if (p.hidden.size) {
      // Hidden until one of its own entrance behaviours has begun (an earlier-listed emphasis does not count).
      const entered = new Set<TargetKey>()
      for (const e of [...snapshot.active, ...snapshot.frozen]) {
        if (!('spid' in e.node.target)) continue
        const key = targetKey(e.node.target)
        if (p.entrances.get(key)?.has(e.node)) entered.add(key)
        // A fanned behaviour enters all its items; renderStateAt hides those whose turn has not come.
        for (const item of fan?.get(e.node)?.items ?? []) if (p.entrances.get(item.key)?.has(e.node)) entered.add(item.key)
      }
      for (const k of p.hidden) if (!entered.has(k)) state.set(k, HIDDEN)
    }
    return state
  }

  function draws(p: Prepared, state: RenderState): LayerDraw[] {
    const b = p.built
    if (!b || b.epoch !== epoch) return []
    const out: LayerDraw[] = []
    kinds = []
    const walk = (layers: readonly PlanLayer[], parent: TargetState | undefined) => {
      for (const l of layers) {
        if (l.kind === 'watermark') {
          const h = b.handles.get(l)
          if (h) { out.push({ handle: h, state: PLAIN }); kinds.push('watermark') }
          continue
        }
        if (l.kind === 'static') {
          const h = b.handles.get(l)
          if (h) { out.push({ handle: h, state: compose(parent, undefined) }); kinds.push('static') }
          continue
        }
        const own = state.get(l.key)
        const s = compose(parent, own)
        if (!s.visible) continue
        if (l.sub) { walk(l.sub, s); continue }
        const h = b.handles.get(l)
        if (!h) continue
        const part = l.part === 'fill' || l.part === 'line' || l.part === 'text' ? l.part : undefined
        out.push({ handle: h, state: s, ...(part ? { part } : {}) })
        kinds.push(`target:${l.key}${l.part ? `:${l.part}` : ''}`)
      }
    }
    walk(p.plan.layers, undefined)
    return out
  }

  /** A full-slide frame of `p` at `state`, painted directly and uploaded; charged to `owner` (refused: undefined, diagnosed by the caller). */
  function flatten(p: Prepared, state: RenderState, owner: AllocationOwner, scale: number, tag: string): Frame | undefined {
    if (manager.state === 'lost') return undefined
    const { w, h } = slidePx(p)
    const cw = Math.max(1, Math.ceil(w * scale)), ch = Math.max(1, Math.ceil(h * scale))
    const l = owner.lease(surfaceBytes(cw, ch), tag)
    if (!l) return undefined
    let canvas: HTMLCanvasElement
    try {
      canvas = factory(cw, ch) as HTMLCanvasElement
      const ctx = canvas.getContext('2d')
      if (!ctx) { l.release(); return undefined }
      paintDirect(ctx, p.slide, state, { images: p.images, resolveFont, scale, scratch: terminal.scratch, sink, builds: p.meta.builds })
    } catch (e) {
      l.release()
      throw e
    }
    const up = compositor().upload({ canvas, bounds: { x: 0, y: 0, w: cw / scale, h: ch / scale }, scale, bytes: cw * ch * 4 }, l)
    if ('refused' in up) { l.release(); return undefined }
    return { handle: up, epoch }
  }

  /** A black full-slide frame (the first slide's transition source), charged to the workspace. */
  function blackFrame(p: Prepared, scale: number): Frame | undefined {
    const { w, h } = slidePx(p)
    const cw = Math.max(1, Math.ceil(w * scale)), ch = Math.max(1, Math.ceil(h * scale))
    const l = workspace.lease(surfaceBytes(cw, ch), 'transition:black')
    if (!l) return undefined
    const canvas = factory(cw, ch) as HTMLCanvasElement
    const ctx = canvas.getContext('2d')
    if (!ctx) { l.release(); return undefined }
    ctx.fillStyle = '#000'
    ctx.fillRect(0, 0, cw, ch)
    const up = compositor().upload({ canvas, bounds: { x: 0, y: 0, w: cw / scale, h: ch / scale }, scale, bytes: cw * ch * 4 }, l)
    if ('refused' in up) { l.release(); return undefined }
    return { handle: up, epoch }
  }

  const releaseFrame = (f: Frame | undefined) => { if (f && f.epoch === epoch) compositor().release(f.handle) }

  // -------------------------------------------------------------------
  // Painting
  // -------------------------------------------------------------------

  function paintTerminal(p: Prepared, state: RenderState) {
    const key = `${p.index}|${output.width}x${output.height}|${stateKey(state)}`
    if (key === lastDirect) return
    lastDirect = key
    const ctx = output.getContext('2d')
    if (!ctx) return
    const v = viewport(p)
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.fillStyle = '#000'
    ctx.fillRect(0, 0, output.width, output.height)
    ctx.translate(v.matrix[4], v.matrix[5])
    try {
      paintDirect(ctx, p.slide, state, {
        images: p.images, resolveFont, scale: v.scale, scratch: terminal.scratch, sink, builds: p.meta.builds,
        ...(deps.watermark ? { watermark: deps.watermark } : {}),
      })
    } catch (e) {
      sink.add({ code: 'paint-failed', slideIndex: p.index, message: `The slide could not be painted: ${e instanceof Error ? e.message : String(e)}` })
    } finally {
      ctx.setTransform(1, 0, 0, 1, 0, 0)
    }
  }

  function clearOutput() {
    lastDirect = ''
    const ctx = output.getContext('2d')
    if (!ctx) return
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.fillStyle = '#000'
    ctx.fillRect(0, 0, output.width, output.height)
  }

  let timelineOf: SlideTimeline | undefined

  const stage: Stage = {
    get plan() { return current?.plan },
    get level() { return current?.level },
    get backend() { return manager.state },
    get backingStore() { return backing },
    // What is drawn: a catalogued stand-in reports its 2-D look-alike (S1-22); `random` draws its pick per rng.
    get transition() { return running && { effect: running.spec.element === 'random' ? 'random' : resolveTransition2D(running.spec).spec.element, duration: running.duration } },
    inspect() {
      return {
        current: current?.index,
        slides: [...prepared.values()].sort((a, b) => a.index - b.index).map(p => ({
          index: p.index, level: p.level, scale: p.built?.scale, layers: p.layers.length,
          uploaded: p.built && p.built.epoch === epoch ? p.built.handles.size : 0, pending: p.built?.pending.length ?? 0,
        })),
        lastDraws: [...kinds],
      }
    },

    leave() {
      const p = current
      running && (releaseFrame(running.from), releaseFrame(running.to))
      running = undefined
      releaseFrame(outgoing)
      outgoing = undefined
      if (p && p.level && p.level !== 'terminal' && !disposed) {
        // The outgoing frame comes from the transition allowance the slide's rung already priced (its own
        // incoming frame was released when its transition ended), so it is never charged twice (S1-22 fix 1).
        // An eviction while it is drawn is safe: a retired rung keeps its charge until its last lease goes.
        // Only a slide without an open rung falls back to the workspace.
        const s = viewport(p).scale
        const scale = p.level === 'full' ? s : s / Math.max(1, backing.dpr)
        const own = owners.get(p.index)
        const f = (own && flatten(p, liveState, own, scale, 'transition:from')) || flatten(p, liveState, workspace, scale, 'transition:from')
        if (f) outgoing = f
        else sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The outgoing frame did not fit the slideshow workspace; the slide change is a cut' })
      }
      current = undefined
      timelineOf = undefined
      origin = Number.POSITIVE_INFINITY
    },

    textTiming(index, meta, charge) {
      const tt = textTimingOf(index, meta, charge)
      return tt && tt.fan.size > 0 ? { tree: tt.tree, orderOf: tt.orderOf } : undefined
    },

    async enter(model, t, neighbours) {
      if (disposed) return
      const p = prepare(model.index, model.slide)
      current = p
      timelineOf = model.timeline
      // The window is the slide and its visible neighbours, which hidden slides can put further than ±1 (PR #18 review 6).
      residency.show(model.index, neighbours)
      for (const other of [...prepared.values()]) if (other !== p && !neighbours.includes(other.index)) drop(other)
      compositor().setSlideIndex(model.index)
      compositor().setViewport(viewport(p))
      await p.ready
      if (current !== p || disposed) return
      await ensure(p, 'full')
      if (current !== p || disposed) return
      const level = p.level ?? 'terminal'
      liveState = stateOf(p, model.cursor, 0, model.timeline)
      // Transition: flattened rungs crossfade; terminal and lost backends cut.
      let spec = model.slide.transition
      if (spec && deps.reduced) spec = reducedTransition(spec)
      const drawn = spec ? resolveTransition2D(spec) : undefined
      // S3-1: a registered 3-D row on WebGL2 runs its shader; otherwise its 2-D look-alike, diagnosed (after reduced motion, so it never adds motion).
      const route = spec ? dispatchTransition(spec, compositor(), sink, p.index) : 'lookalike'
      // Sounds are Stage 2 (spec 13): the transition plays silently, diagnosed once for the slide.
      const sound = model.slide.transition?.sound
      if (sound?.start) sink.add({ code: 'media-stand-in', slideIndex: p.index, effect: 'transition-sound', message: 'The transition sound is pending Stage 2: the transition plays silently' })
      if (sound?.end) sink.add({ code: 'media-stand-in', slideIndex: p.index, effect: 'transition-sound-end', message: 'Stopping sounds at the transition is pending Stage 2: nothing plays to stop' })
      const duration = spec && drawn && drawn.spec.element !== 'cut' && Number.isFinite(spec.durationMs) ? Math.max(0, spec.durationMs) : 0
      // Preparation (image decoding, layer construction) took show time: the transition, and the slide's
      // timeline after it, start now that the incoming frame can be drawn (PR #18 review P1).
      const start = deps.showNow ? Math.max(t, deps.showNow()) : t
      // The show's first slide transitions from black, as PowerPoint does (S1-20 review M1).
      if (spec && duration > 0 && !outgoing && !shownBefore && level !== 'terminal' && manager.state !== 'lost') {
        const s = viewport(p).scale
        outgoing = blackFrame(p, level === 'full' ? s : s / Math.max(1, backing.dpr))
        if (!outgoing) sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The black source frame did not fit the slideshow workspace; the first slide is a cut' })
      }
      shownBefore = true
      if (spec && duration > 0 && outgoing && level !== 'terminal' && manager.state !== 'lost') {
        const owner = owners.get(p.index) ?? workspace
        const s = viewport(p).scale
        const to = flatten(p, liveState, owner, level === 'full' ? s : s / Math.max(1, backing.dpr), 'transition:to')
        if (to) {
          running = { spec: level === 'flattened' ? { ...spec, element: 'fade', ns: 'p', attrs: {} } : spec, three: route === 'shader' && level !== 'flattened', from: outgoing, to, start, duration, rng: createRng((deps.seed ?? 0x5eed) ^ (p.index * 2654435761)) }
          outgoing = undefined
        } else sink.add({ code: 'transition-workspace', slideIndex: p.index, message: 'The incoming frame did not fit the slideshow workspace; the slide change is a cut' })
      }
      origin = running ? start + duration : start
      if (!running) { releaseFrame(outgoing); outgoing = undefined }
      container.hidden = level === 'terminal'
      if (level === 'terminal') lastDirect = ''
      // Prefetch the neighbours in the background (catching 'cancelled').
      for (const n of neighbours) {
        if (n === p.index || prepared.has(n)) continue
        void deps.metadata(n).then((h) => {
          if (disposed || current !== p) { h.release(); return }
          const q = prepare(n, h.slide, h)
          return q.ready.then(() => ensure(q, 'full'))
        }).catch(() => undefined)
      }
    },

    slideTime(t) { return t - origin },

    stateAt(cursor, t) {
      const p = current
      if (!p) return EMPTY_STATE
      return stateOf(p, cursor, t - origin, timelineOf)
    },

    previewOf(index, cursor, timeline) {
      const p = current?.index === index ? current : prepared.get(index)
      if (!p || p.dead || disposed) return undefined
      // A settled checkpoint reads the same at any time; the grid's static state has no cursor.
      return { slide: p.slide, state: stateOf(p, cursor, 0, timeline), images: p.images, resolveFont, builds: p.meta.builds }
    },

    frame(cursor, t) {
      if (disposed) return
      const p = current
      if (!p) {
        // Between slides: the outgoing frame keeps painting until the next one is ready.
        if (outgoing && outgoing.epoch === epoch) compositor().frame([{ handle: outgoing.handle, state: PLAIN }])
        return
      }
      const state = liveState = stateOf(p, cursor, t - origin, timelineOf)
      const level = p.level
      if (level === undefined) {
        if (outgoing && outgoing.epoch === epoch) compositor().frame([{ handle: outgoing.handle, state: PLAIN }])
        return state
      }
      if (level === 'terminal') { container.hidden = true; paintTerminal(p, state); return state }
      container.hidden = false
      // The Canvas 2D compositor may draw on the output: a later terminal paint must not skip as unchanged.
      lastDirect = ''
      // Re-rasterisation work (resize, backend switch), bounded per frame: current first, then neighbours.
      let budget = RERASTER_LIMIT
      budget -= drain(p, budget)
      for (const q of prepared.values()) if (q !== p && budget > 0) budget -= drain(q, budget)
      const c = compositor()
      if (running) {
        const progress = running.duration > 0 ? (t - running.start) / running.duration : 1
        if (progress < 1 && running.to.epoch === epoch && (!running.from || running.from.epoch === epoch)) {
          // transitionDraw first (a random pick draws from the rng), then frame with the same rng.
          const aspect = slidePx(p).w / slidePx(p).h
          const params = running.three ? transition3DDraw(running.spec, progress, running.rng, aspect) : transitionDraw(running.spec, progress, running.rng, aspect)
          const wm = p.plan.layers[p.plan.layers.length - 1]
          const mark = wm && p.built?.epoch === epoch ? p.built.handles.get(wm) : undefined
          c.frame(mark ? [{ handle: mark, state: PLAIN }] : [], {
            from: running.from?.handle ?? running.to.handle, to: running.to.handle,
            effect: running.spec.element, progress, params, rng: running.rng,
          })
          return state
        }
        releaseFrame(running.from)
        releaseFrame(running.to)
        running = undefined
      }
      if (level === 'flattened') {
        // Checkpoint granularity (the checkpoint-cut fallback; PR #18 review 8): the frame shows the settled state
        // the cursor is heading to, and is re-rasterised only when that changes, never per animation frame.
        const b = p.built
        const owner = owners.get(p.index)
        if (b && owner && b.epoch === epoch && (b.flat?.cursor !== cursor || !b.flat)) {
          const settled = stateOf(p, cursor, SETTLED_MS, timelineOf)
          const key = `${cursor?.step ?? 0}|${stateKey(settled)}`
          if (b.flat?.key === key) b.flat.cursor = cursor
          else {
            if (b.flat) { c.release(b.flat.handle); b.flat = undefined }
            const f = flatten(p, settled, owner, b.scale, 'flattened')
            if (f) b.flat = { key, handle: f.handle, cursor }
            else void stepDown(p)
          }
        }
        const flat = p.built?.flat
        c.frame(flat ? [{ handle: flat.handle, state: PLAIN }] : [])
        return state
      }
      const list = draws(p, state)
      limits.check(p.index, { layers: list.length, reRasters: RERASTER_LIMIT - budget })
      c.frame(list)
      return state
    },

    toSlide(clientX, clientY) {
      const p = current
      if (!p) return undefined
      const r = root.getBoundingClientRect()
      if (!(r.width > 0 && r.height > 0)) return undefined
      // Every surface canvas is letterboxed (object-fit: contain) to the backing store's aspect.
      const k = Math.min(r.width / backing.w, r.height / backing.h)
      const dx = (clientX - (r.left + (r.width - backing.w * k) / 2)) / k
      const dy = (clientY - (r.top + (r.height - backing.h * k) / 2)) / k
      const inv = invert(viewport(p).matrix)
      if (!inv) return undefined
      const [x, y] = apply(inv, dx, dy)
      return { x, y }
    },

    mediaBox(spid) {
      const p = current
      const g = p?.base.get(targetKey({ spid }))
      if (!p || !g) return undefined
      const r = root.getBoundingClientRect()
      if (!(r.width > 0 && r.height > 0)) return undefined
      // Slide px → device px (the viewport) → CSS px of the letterboxed canvas (object-fit: contain), as toSlide inverts.
      const k = Math.min(r.width / backing.w, r.height / backing.h)
      const m = viewport(p).matrix
      const ox = (r.width - backing.w * k) / 2, oy = (r.height - backing.h * k) / 2
      return { left: ox + (m[0] * g.x + m[4]) * k, top: oy + (m[3] * g.y + m[5]) * k, width: m[0] * g.w * k, height: m[3] * g.h * k }
    },

    runAt(shape, u, v, w, h) {
      const body = shape.textBody
      if (!body || body.paragraphs.length === 0) return null
      if ((body.direction && body.direction !== 'horz') || body.textWarp) return undefined
      const ctx = (measurer ??= factory(1, 1)).getContext('2d')
      if (!ctx) return undefined
      const layout = layoutTextBody(body, w, h, createTextBodyMeasurer(ctx as CanvasRenderingContext2D, resolveFont), current?.slide.theme)
      // A paragraph-build clone paints only its paragraphs (as rasterisation does): runs elsewhere are not on it (PR #18 review 6).
      const keep = textFilter.get(shape)
      for (const line of layout.lines) {
        if ((keep && !keep(line.paragraphIndex)) || v < line.y || v > line.y + line.height) continue
        for (const seg of line.segments) {
          if (seg.transform || u < seg.x || u > seg.x + seg.width) continue
          return { p: line.paragraphIndex, r: seg.runIndex }
        }
      }
      return null
    },

    resize(css, dpr) {
      if (disposed) return
      const before = backing
      const r = residency.resize(css, dpr)
      if (r !== 'resized') return
      backing = residency.backing
      // Swap the output: the terminal surface and the compositor's backing store, then refund the old reserve.
      output.width = backing.w
      output.height = backing.h
      lastDirect = ''
      // The compositor's output is part of the transaction: when it keeps its old size (WebGL2 under pressure),
      // the old backing store, layer scale and viewport stay and the new reserve is refunded (PR #18 review 4).
      if (compositor().setViewport(viewport(current)) === false) {
        residency.abortResize()
        backing = residency.backing
        output.width = backing.w
        output.height = backing.h
        compositor().setViewport(viewport(current))
        sink.add({ code: 'resize-kept-old', message: 'the compositor could not take the new output size; the old backing store is kept and scaled' })
        return
      }
      residency.commitResize()
      // A DPR-only change keeps the pixel size but moves the reduced rung's layer scale (S1-20 review M2).
      if (before.w === backing.w && before.h === backing.h && before.dpr === backing.dpr) return
      // Current, previous and next re-rasterise at the new size; the old layers paint scaled meanwhile.
      for (const p of [...prepared.values()].sort((a, b) => (a === current ? -1 : b === current ? 1 : 0))) {
        if (p.level && p.level !== 'terminal') void ensure(p, p.level)
      }
    },

    dispose() {
      if (disposed) return
      disposed = true
      running && (releaseFrame(running.from), releaseFrame(running.to))
      running = undefined
      releaseFrame(outgoing)
      outgoing = undefined
      for (const p of [...prepared.values()]) drop(p)
      residency.dispose()
      manager.dispose()
      container.remove()
      void fonts?.then((l) => l?.dispose())
    },
  }
  clearOutput()
  return stage
}
