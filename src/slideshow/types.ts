/**
 * Slideshow foundation contracts (slideshow plan S1-2; spec 4.2, 6.1–6.4, 7).
 *
 * Types only: nothing here emits code. Later tasks consume these names
 * verbatim; extend them, never rename them.
 */
import type { XmlNode } from '../core/xml'
import type { Transition3D } from './render/three/registry'

// ---------------------------------------------------------------------------
// Random numbers
// ---------------------------------------------------------------------------

/** A seeded generator returning values in [0, 1); see `createRng` (rng.ts). */
export type Rng = () => number

// ---------------------------------------------------------------------------
// Diagnostics
// ---------------------------------------------------------------------------

export interface SlideShowDiagnostic { code: string; slideIndex?: number; effect?: string; message: string }

/** Deduplicated by `(code, slideIndex, effect)`; never fed per frame. */
export interface DiagnosticSink { add(d: SlideShowDiagnostic): void; list(): readonly SlideShowDiagnostic[] }

// ---------------------------------------------------------------------------
// Targets
// ---------------------------------------------------------------------------

export type TargetRef =
  | { spid: string; text?: { kind: 'p' | 'char'; st: number; end: number }; sub?: 'bg' | 'subSp' | 'chart' | 'graphic' | 'ink' }
  | { bg: true }
  | { sld: true }
  | { snd: string }

/** The string identity of a `TargetRef`, from `targetKey` (target.ts). */
export type TargetKey = string

/** Answers target-resolution questions for one slide, built from the parsed `PptxSlide` (after compatibility selection) and its underlay. */
export interface TargetResolver {
  shape(spid: string): 'slide' | 'master-only' | 'missing'
  paragraphs(spid: string): ReadonlyArray<{ level: number; runs: number; empty: boolean }>
}

// ---------------------------------------------------------------------------
// Timing tree (spec 6.1 table; units are normalised by the reader)
// ---------------------------------------------------------------------------

export interface Condition {
  /** Milliseconds, or `'indefinite'`. */
  delayMs: number | 'indefinite'
  /** `onBegin`, `onEnd`, `onClick`, `onNext`, `onPrev`, `begin`, `end`, `onMouseOver`, `onMouseOut`, `onStopAudio`. */
  evt?: string
  tn?: number
  target?: TargetRef
  rtn?: string
}

export interface Iterate { type: 'el' | 'wd' | 'lt'; backwards: boolean; tmAbsMs?: number; tmPct?: number }

/** `p:cTn`. Every attribute of the spec 6.1 table has a field; fields the engine does not act on are still carried. */
export interface CTn {
  id: number
  presetClass?: string
  presetId?: number
  presetSubtype?: number
  nodeType?: string
  grpId?: number
  durMs: number | 'indefinite' | undefined
  /** In iterations (the reader converts the file's thousandths). */
  repeatCount?: number | 'indefinite'
  repeatDurMs?: number | 'indefinite'
  /** Fraction of the duration, 0–1. */
  accel: number
  /** Fraction of the duration, 0–1. */
  decel: number
  autoRev: boolean
  fill?: 'hold' | 'remove' | 'freeze' | 'transition'
  restart?: 'always' | 'whenNotActive' | 'never'
  syncBehavior?: 'canSlip' | 'locked'
  masterRel?: 'sameClick' | 'lastClick' | 'nextClick'
  afterEffect?: boolean
  start: Condition[]
  end: Condition[]
  endSync?: Condition
  iterate?: Iterate
}

export interface Keyframe { tm: number | 'indefinite'; value: string | number; fmla?: string }

/** Produced by S1-6's `parseMotionPath`; coordinates are fractions of the slide. */
export interface MotionPath { segments: ReadonlyArray<{ cmd: 'M' | 'L' | 'C' | 'Z'; pts: readonly number[] }>; length: number }

export interface ContainerNode {
  kind: 'par' | 'seq' | 'excl'
  ctn: CTn
  children: TimeNode[]
  seq?: { concurrent: boolean; nextAc: string; prevAc: string }
}

/** Fields shared by every behaviour (`p:cBhvr`). */
export interface BehaviourBase {
  ctn: CTn
  target: TargetRef
  attrNames: string[]
  additive: 'base' | 'sum' | 'repl' | 'mult' | 'none'
  accumulate: 'always' | 'none'
  override?: 'normal' | 'childStyle'
  xfrmType?: 'pt' | 'img'
  rctx?: string
}

/** An animation value (`p:tav`/`p:to` variant): string, number, or colour as written. */
type AnimValue = string | number
export interface Point { x: number; y: number }

export interface SetBehaviour extends BehaviourBase { kind: 'set'; to?: AnimValue }
export interface AnimBehaviour extends BehaviourBase {
  kind: 'anim'
  keys: Keyframe[]
  calcmode?: 'discrete' | 'lin' | 'fmla'
  valueType?: 'str' | 'num' | 'clr'
  by?: string
  from?: string
  to?: string
}
export interface AnimEffectBehaviour extends BehaviourBase {
  kind: 'animEffect'
  transition?: 'in' | 'out' | 'none'
  filter?: string
  prLst?: string
  progress?: AnimValue
}
export interface AnimMotionBehaviour extends BehaviourBase {
  kind: 'animMotion'
  path?: MotionPath
  origin?: 'parent' | 'layout'
  /** Degrees (the reader converts 1/60000°). */
  rAngDeg?: number
  ptsTypes?: string
  by?: Point
  from?: Point
  to?: Point
  rCtr?: Point
}
export interface AnimRotBehaviour extends BehaviourBase {
  kind: 'animRot'
  /** Degrees (the reader converts 1/60000°). */
  by?: number
  from?: number
  to?: number
}
export interface AnimScaleBehaviour extends BehaviourBase {
  kind: 'animScale'
  /** Scale factors, 1 = 100 % (the reader converts 1/1000 %). */
  by?: Point
  from?: Point
  to?: Point
  zoomContents?: boolean
}
export interface AnimClrBehaviour extends BehaviourBase {
  kind: 'animClr'
  clrSpc: 'rgb' | 'hsl'
  dir: 'cw' | 'ccw'
  /** Colours as `#rrggbb`, or a by-offset as written (`hsl(h,s,l)` deltas). */
  by?: string
  from?: string
  to?: string
}
export interface CmdBehaviour extends BehaviourBase { kind: 'cmd'; cmdType: 'evt' | 'call' | 'verb'; cmd: string }
/** Additive (S2-7): what `p:cMediaNode` carries for playback. */
interface MediaNodeFields {
  /** `numSld`: how many slides the media plays across (cross-slide sound, spec 7). */
  numSld?: number
  /** `vol`, 0–1 (the reader converts 1/1000 %). */
  volume?: number
  muted?: boolean
}
export interface AudioBehaviour extends BehaviourBase, MediaNodeFields { kind: 'audio'; isNarration?: boolean }
export interface VideoBehaviour extends BehaviourBase, MediaNodeFields { kind: 'video'; fullScrn?: boolean }

export type BehaviourNode =
  | SetBehaviour
  | AnimBehaviour
  | AnimEffectBehaviour
  | AnimMotionBehaviour
  | AnimRotBehaviour
  | AnimScaleBehaviour
  | AnimClrBehaviour
  | CmdBehaviour
  | AudioBehaviour
  | VideoBehaviour

export type TimeNode = ContainerNode | BehaviourNode

export type BuildTemplate =
  | { kind: 'bldP'; spid: string; grpId: number; build: 'whole' | 'p' | 'cust' | 'allAtOnce'; bldLvl: number; animBg: boolean; rev: boolean }
  | { kind: 'bldGraphic' | 'bldOleChart' | 'bldDgm'; spid: string; grpId: number; bld?: string }

export interface TimingTree { root: TimeNode; build: BuildTemplate[] }

/** The normalised per-shape build record the render planner consumes. */
export interface BuildOwnership {
  spid: string
  mode: 'whole' | 'paragraph'
  ranges: ReadonlyArray<{ key: TargetKey; st: number; end: number }>
  animBg: boolean
  body: TargetKey
}

// ---------------------------------------------------------------------------
// Slide model
// ---------------------------------------------------------------------------

export interface TransitionSpec {
  element: string
  ns: 'p' | 'p14' | 'p15' | 'p159'
  preset?: string
  attrs: Readonly<Record<string, string>>
  durationMs: number
  advClick: boolean
  advTmMs?: number
  sound?: { start?: { rId: string; loop: boolean }; end?: true }
}

export interface NoteParagraph { level: number; bullet?: string; text: string }

export interface LinkSpec {
  spid: string
  run?: { p: number; r: number }
  action:
    | { kind: 'slide'; index: number }
    | { kind: 'jump'; verb: 'nextslide' | 'previousslide' | 'firstslide' | 'lastslide' | 'endshow' }
    | { kind: 'url'; href: string }
    | { kind: 'refused'; reason: string }
}

export interface MediaRecord { rId: string; partPath: string; contentType: string; declaredBytes: number; spid?: string }

export interface ShowSlide {
  index: number
  hidden: boolean
  transition?: TransitionSpec
  timing: TimingTree | { refused: string } | undefined
  notes: NoteParagraph[]
  links: LinkSpec[]
  media: MediaRecord[]
  /** Retained and charged with the rest of the slide's metadata handle. */
  builds: readonly BuildOwnership[]
}

export interface ShowSettings {
  loop: boolean
  showAnimation: boolean
  useTimings: boolean
  showNarration: boolean
  type: 'present' | 'browse' | 'kiosk'
  kioskRestartMs: number
  range?: { st: number; end: number }
}

export type SlideTimingReader = (
  slideXml: XmlNode,
  slideIndex: number,
  resolve: TargetResolver,
  sink: DiagnosticSink,
) => { transition?: TransitionSpec; timing: TimingTree | { refused: string } | undefined; builds: BuildOwnership[] }

/** The cheap deck-level part of the model. */
export interface ShowIndex {
  slides: ReadonlyArray<{ index: number; hidden: boolean; partPath?: string }>
  settings: ShowSettings
  diagnostics: SlideShowDiagnostic[]
}

/**
 * One slide's metadata, charged to one lease. Consumers hold a handle for as
 * long as they use the slide; `release` is idempotent. `attach` reserves
 * growth against the store's 64 MiB aggregate and the workspace **before**
 * the consumer allocates; `false` means refused, with no allocation.
 */
export interface MetadataHandle { readonly slide: ShowSlide; attach(bytes: number, tag: string): boolean; release(): void }

/**
 * `evict(i)` drops the cache entry; the charge stays until the last handle is
 * released, and a later `acquire(i)` re-reads into a fresh entry. `close()`
 * cancels in-flight reads (they reject with `'cancelled'` and charge nothing).
 */
export interface SlideMetadataStore {
  acquire(index: number, signal?: AbortSignal): Promise<MetadataHandle>
  evict(index: number): void
  resident(): readonly number[]
  close(): void
}

export interface ShowModel { index: ShowIndex; store: SlideMetadataStore }

// ---------------------------------------------------------------------------
// Render state
// ---------------------------------------------------------------------------

export type Affine = readonly [a: number, b: number, c: number, d: number, e: number, f: number]

/** A 4×5 colour matrix, row-major (20 numbers), applied to un-premultiplied colour. */
export type ColorMatrix = readonly number[]

/**
 * Every mask the catalogue names (spec 6.3: the wipe-like filters and the 2-D
 * transition masks). Geometry and params are in `render/masks.ts`; `dissolve`
 * and `randomBar` are noise rasters, every other kind is a path.
 */
export type MaskKind =
  | 'wipe' | 'split' | 'circle' | 'diamond' | 'plus' | 'wedge' | 'wheel' | 'barn' | 'checkerboard' | 'dissolve'
  | 'blinds' | 'box' | 'comb' | 'randomBar' | 'strips'

export interface MaskState { kind: MaskKind; progress: number; params: Readonly<Record<string, number | string>> }

export interface TargetState {
  visible: boolean
  opacity: number
  transform: Affine
  mask?: MaskState
  color?: { fill?: ColorMatrix; line?: ColorMatrix; text?: ColorMatrix }
  motion?: { dx: number; dy: number }
}

export type RenderState = ReadonlyMap<TargetKey, TargetState>

// ---------------------------------------------------------------------------
// Events and clock
// ---------------------------------------------------------------------------

/** External scheduler inputs; `onBegin`/`onEnd`/`begin`/`end` are internal to the scheduler. */
export type ShowEvent =
  | { kind: 'next'; t: number }
  | { kind: 'previous'; t: number }
  | { kind: 'trigger'; spid: string; t: number }
  | { kind: 'mouseOver' | 'mouseOut'; spid: string; t: number }
  | { kind: 'mediaEnd'; nodeId: number; t: number; synthetic: boolean }
  | { kind: 'stopAudio'; t: number }

export interface ShowClock { now(): number; paused: boolean }

// ---------------------------------------------------------------------------
// Workspace (spec 7). Every producer reserves before it allocates.
// ---------------------------------------------------------------------------

/** `release` is idempotent and returns the bytes exactly once; `transfer` moves ownership without a second charge. */
export interface Lease {
  readonly tag: string
  readonly bytes: number
  grow(bytes: number): boolean
  shrink(bytes: number): void
  transfer(tag: string): Lease
  release(): void
}

/** The minimal owner interface; `SlideShowWorkspace` and `SubOwner` both implement it. */
export interface AllocationOwner { lease(bytes: number, tag: string): Lease | undefined }

/** A bounded sub-owner carved from bytes already reserved: its leases never touch the parent ledger. */
export interface SubOwner extends AllocationOwner {}

export interface TerminalReserve {
  readonly outputBytes: number
  readonly scratch: SubOwner
  resize(outputBytes: number): boolean
  release(): void
}

export interface SlideShowWorkspace extends AllocationOwner {
  readonly limit: number
  lease(bytes: number, tag: string): Lease | undefined
  reserveTerminal(outputBytes: number): TerminalReserve | undefined
  used(): number
}

// ---------------------------------------------------------------------------
// Compositor (spec 6.3)
// ---------------------------------------------------------------------------

/** Injected so node-canvas drives tests. */
export type CanvasFactory = (w: number, h: number) => HTMLCanvasElement | OffscreenCanvas

/** Bounds are padded for stroke, shadow and glow. */
export interface LayerBitmap {
  canvas: HTMLCanvasElement | OffscreenCanvas
  bounds: { x: number; y: number; w: number; h: number }
  scale: number
  bytes: number
}

/** Opaque. */
export interface LayerHandle { readonly id: number }

export interface LayerDraw { handle: LayerHandle; state: TargetState; part?: 'fill' | 'line' | 'text' }

export interface TransitionDraw {
  from: LayerHandle
  to: LayerHandle
  effect: string
  progress: number
  params: Readonly<Record<string, unknown>>
  rng: Rng
}

/**
 * Where the slide lands on the output (S1-13 fix round 1, controller ruling).
 * `matrix` maps slide px to output device px (letterbox offset and
 * backing-store scale, DPR included); `deviceWidth` × `deviceHeight` is the
 * output backing store in device px.
 */
export interface CompositorViewport { matrix: Affine; deviceWidth: number; deviceHeight: number }

/**
 * `TargetState.transform` and `motion` are in slide px (the renderer knows
 * nothing about the viewport). A draw is composited at
 * `viewport.matrix ∘ translate(motion) ∘ state.transform`. Transition affines
 * are in unit slide space (the layer's bounds are [0, 1]²) and transition
 * masks are computed in device px over the slide's device rect.
 */
export interface Compositor {
  readonly backend: 'webgl2' | 'canvas2d'
  readonly maxTextureSize: number
  upload(layer: LayerBitmap, lease: Lease): LayerHandle | { refused: 'workspace' | 'texture-size' }
  release(handle: LayerHandle): void
  frame(draws: readonly LayerDraw[], transition?: TransitionDraw): void
  resize(widthPx: number, heightPx: number): void
  /**
   * Additive (S1-13 fix round 1). Resizes the output when the device size differs. Default: identity, current size.
   * Returns false when the output could not take the new size and keeps its old one (PR #18 review 4).
   */
  setViewport(view: CompositorViewport): boolean | void
  /** Additive (S1-13 fix round 1): the slide that compositor diagnostics are attributed to. */
  setSlideIndex(index: number | undefined): void
  /**
   * Additive (S3-1): WebGL2 only. Prepares `t` (leases its grids, compiles its
   * programs) so a transition dispatched to its shader can start; `false`
   * when it cannot (lost context, refused lease, failed compile).
   */
  prepareTransition3D?(t: Transition3D): boolean
  dispose(): void
}

// ---------------------------------------------------------------------------
// Effect catalogue (spec 6.4)
// ---------------------------------------------------------------------------

/** Frame-compare tolerance, as fractions of 255. */
export interface Tolerance { meanAbs: number; rowMax: number }

export interface CatalogueRow {
  id: string
  kind: 'transition' | 'animation' | 'behaviour'
  element?: string
  ns?: string
  preset?: { cls: string; id: number; subtype?: number }
  uiName: string
  implementation: 'file' | 'preset-table' | 'mask' | 'transform' | 'shader' | 're-raster' | 'stand-in'
  status: 'exact' | 'approximate' | 'pending-stage-2' | 'pending-stage-3' | 'never'
  verified: boolean
  lookalike?: string
  tolerance: Tolerance
  stage: 1 | 2 | 3
}
