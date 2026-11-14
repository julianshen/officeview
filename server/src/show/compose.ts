/**
 * Server composition of a prepared slide (piece 4 §6): the slide's render
 * state at a cursor and time, its layer draws in paint order, and the
 * composite of those draws on the Canvas 2D compositor
 * (`createCanvas2DCompositor`), over the `@napi-rs/canvas` backend.
 *
 * The state and the draw list mirror the local stage exactly (`stateOf` and
 * `entriesOf` in `src/slideshow/player/stage.ts`, both private there): a
 * hidden target stays hidden until one of its own entrance behaviours has
 * begun, states are relative to the base rotation the layer is painted at,
 * group states compose over their sub-layers, a paragraph range's mask
 * sweeps its text box, and a split layer takes its part's colour. If the two
 * diverged, the eligibility envelope would describe another show than the
 * one the client plays.
 *
 * The watermark layer is never composed here: marks are applied to the
 * classified assets later (D7), so classification sees unmarked ink only.
 */
import { showCanvasBackend } from './canvas'
import type { FontResolver } from '../../../src/core/fonts/register'
import { mul, stateMatrix } from '../../../src/slideshow/affine'
import { renderStateAt, type BaseGeometry } from '../../../src/slideshow/engine/behaviours'
import type { SlideCursor, SlideTimeline } from '../../../src/slideshow/engine/steps'
import { compose } from '../../../src/slideshow/prepare/morph'
import type { PreparedSlide } from '../../../src/slideshow/prepare'
import { createCanvas2DCompositor } from '../../../src/slideshow/render/canvas2d'
import type { PlanLayer } from '../../../src/slideshow/render/plan'
import { slideSize, textRangeBox } from '../../../src/slideshow/render/scene'
import { targetKey } from '../../../src/slideshow/target'
import type {
  Affine, CanvasFactory, Compositor, DiagnosticSink, LayerBitmap, LayerHandle, Lease, RenderState, Rng, SlideShowWorkspace, TargetKey, TargetState,
} from '../../../src/slideshow/types'

const IDENTITY: Affine = [1, 0, 0, 1, 0, 0]
const HIDDEN: TargetState = Object.freeze({ visible: false, opacity: 1, transform: IDENTITY })

/** `transform ∘ rotate(−base.rot about the centre)`: layers are painted at their base rotation already (stage `relativeTo`). */
function relativeTo(state: TargetState, g: BaseGeometry | undefined): TargetState {
  if (!g || !g.rot) return state
  const a = -g.rot * Math.PI / 180, c = Math.cos(a), s = Math.sin(a)
  const cx = g.x + g.w / 2, cy = g.y + g.h / 2
  const back: Affine = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
  return { ...state, transform: mul(state.transform, back) }
}

/** The slide's render state at `cursor` and slide time `t` (stage `stateOf`). */
export function slideStateAt(ps: PreparedSlide, timeline: SlideTimeline, cursor: SlideCursor, t: number, rng: Rng, sink: DiagnosticSink): RenderState {
  if (cursor.mode === 'static') return new Map()
  const snapshot = timeline.stateAt(cursor, Math.max(0, t))
  const fan = ps.textTiming?.fan
  const raw = renderStateAt(snapshot, ps.base, rng, sink, slideSize(ps.slide), fan)
  const state = new Map<TargetKey, TargetState>()
  const baseOf = (k: TargetKey) => ps.base.get(k) ?? (k.includes('/') ? ps.base.get(k.slice(0, k.indexOf('/'))) : undefined)
  for (const [k, v] of raw) state.set(k, relativeTo(v, baseOf(k)))
  if (ps.hidden.size) {
    const entered = new Set<TargetKey>()
    for (const e of [...snapshot.active, ...snapshot.frozen]) {
      if (!('spid' in e.node.target)) continue
      const key = targetKey(e.node.target)
      if (ps.entrances.get(key)?.has(e.node)) entered.add(key)
      for (const item of fan?.get(e.node)?.items ?? []) if (ps.entrances.get(item.key)?.has(e.node)) entered.add(item.key)
    }
    for (const k of ps.hidden) if (!entered.has(k)) state.set(k, HIDDEN)
  }
  return state
}

/** One layer draw: the raster layer, the state it is drawn at (group states composed) and its colour part. */
export interface DrawnLayer { layer: PlanLayer; state: TargetState; part?: 'fill' | 'line' | 'text' }

export interface DrawOptions {
  /** Measures a paragraph range's text box (its mask's sweep box); absent, no range box is measured. */
  measure?: CanvasRenderingContext2D
  resolveFont?: FontResolver
}

/** Every layer draw of `ps` at `state`, in paint order; invisible targets and the watermark are left out (stage `entriesOf`). */
export function drawsAt(ps: PreparedSlide, state: RenderState, opts: DrawOptions = {}): DrawnLayer[] {
  const out: DrawnLayer[] = []
  const walk = (layers: readonly PlanLayer[], parent: TargetState | undefined) => {
    for (const l of layers) {
      if (l.kind === 'watermark') continue
      if (l.kind === 'static') { out.push({ layer: l, state: compose(parent, undefined) }); continue }
      let s = compose(parent, state.get(l.key))
      if (!s.visible) continue
      if (l.part === 'range' && s.mask && !s.maskBox && opts.measure) {
        const b = textRangeBox(l.nodes, ps.slide, opts.resolveFont ?? (f => f), opts.measure)
        if (b) s = { ...s, maskBox: { x: b.minX, y: b.minY, w: b.maxX - b.minX, h: b.maxY - b.minY } }
      }
      if (l.sub) { walk(l.sub, s); continue }
      const part = l.part === 'fill' || l.part === 'line' || l.part === 'text' ? l.part : l.part === 'range' && s.color?.text ? 'text' : undefined
      out.push({ layer: l, state: s, ...(part ? { part } : {}) })
    }
  }
  walk(ps.plan.layers, undefined)
  return out
}

/** The device matrix a draw composites at (`viewport ∘ translate(motion) ∘ transform`). */
export function deviceMatrix(scale: number, state: TargetState): Affine {
  return mul([scale, 0, 0, scale, 0, 0], stateMatrix(state))
}

/** A ledger that grants every lease: the classifier's own surfaces are bounded by `CLASSIFY_BUDGET` and the slide size. */
function openLease(bytes: number, tag: string): Lease {
  const l: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => l, release() {} }
  return l
}
const OPEN_WORKSPACE: SlideShowWorkspace = { limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined, lease: openLease }

export interface SlideComposer {
  /** Output size in device px (`ceil(slide × scale)`). */
  readonly width: number
  readonly height: number
  /** Layers whose raster the compositor refused (over its texture size): they cannot be composed. */
  readonly refused: readonly PlanLayer[]
  /**
   * Composes `draws` (optionally leaving `without` out) and returns the RGBA pixels. Layers without a raster are
   * skipped. With `classified`, each layer given a classified raster (`classify`) is drawn from it instead, or
   * not at all when it was dropped.
   */
  compose(draws: readonly DrawnLayer[], without?: PlanLayer, classified?: boolean): Uint8ClampedArray
  /** Records a layer's classification: its masked raster, or `null` when it is dropped. */
  classify(layer: PlanLayer, raster: LayerBitmap | null): void
  /** The codes the compositor diagnosed since the last call (degraded effects: a mask cut, a skipped recolour). */
  takeDiagnostics(): string[]
  dispose(): void
}

export function createSlideComposer(ps: PreparedSlide, rasters: ReadonlyMap<PlanLayer, LayerBitmap>, scale: number): SlideComposer {
  const { w, h } = slideSize(ps.slide)
  const width = Math.max(1, Math.ceil(w * scale)), height = Math.max(1, Math.ceil(h * scale))
  const backend = showCanvasBackend()
  const factory: CanvasFactory = (cw, ch) => backend.createCanvas(Math.max(1, Math.ceil(cw)), Math.max(1, Math.ceil(ch))) as never
  const canvas = factory(width, height)
  const codes: string[] = []
  const sink: DiagnosticSink = { add: (d) => { codes.push(d.code) }, list: () => [] }
  const c: Compositor = createCanvas2DCompositor(canvas, factory, OPEN_WORKSPACE, sink)
  c.setViewport({ matrix: [scale, 0, 0, scale, 0, 0], deviceWidth: width, deviceHeight: height })
  const handles = new Map<PlanLayer, LayerHandle>()
  const refused: PlanLayer[] = []
  for (const [layer, bmp] of rasters) {
    if (layer.kind === 'watermark') continue
    const up = c.upload(bmp, openLease(0, 'classify'))
    if ('refused' in up) refused.push(layer)
    else handles.set(layer, up)
  }
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const alt = new Map<PlanLayer, LayerHandle | null>()
  return {
    width, height, refused,
    compose(draws, without, classified = false) {
      const list = []
      for (const d of draws) {
        if (d.layer === without) continue
        const handle = classified && alt.has(d.layer) ? alt.get(d.layer) : handles.get(d.layer)
        if (handle) list.push({ handle, state: d.state, ...(d.part ? { part: d.part } : {}) })
      }
      c.frame(list)
      return ctx.getImageData(0, 0, width, height).data
    },
    classify(layer, raster) {
      const prev = alt.get(layer)
      if (prev) c.release(prev)
      if (!raster) { alt.set(layer, null); return }
      const up = c.upload(raster, openLease(0, 'classify'))
      // A refused classified raster is drawn as nothing: the reconstruction check then fails closed.
      alt.set(layer, 'refused' in up ? null : up)
    },
    takeDiagnostics: () => codes.splice(0),
    dispose: () => c.dispose(),
  }
}

/**
 * The rasters a classification ships (§6 rules 2 and 3): every dropped layer left out, and each masked layer
 * redrawn with its alpha multiplied by its layer-local mask (0 removes a pixel, 255 keeps it).
 */
export function classifiedRasters(rasters: ReadonlyMap<PlanLayer, LayerBitmap>, drop: ReadonlySet<PlanLayer>, masks: ReadonlyMap<PlanLayer, Uint8Array>): Map<PlanLayer, LayerBitmap> {
  const out = new Map<PlanLayer, LayerBitmap>()
  for (const [layer, bmp] of rasters) {
    if (drop.has(layer)) continue
    const mask = masks.get(layer)
    out.set(layer, mask ? maskedRaster(bmp, mask) : bmp)
  }
  return out
}

/** `bmp` with its alpha multiplied by `mask` (layer-local, one byte per bitmap px), as a new bitmap. */
export function maskedRaster(bmp: LayerBitmap, mask: Uint8Array): LayerBitmap {
  const w = bmp.canvas.width, h = bmp.canvas.height
  const canvas = showCanvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(bmp.canvas as CanvasImageSource, 0, 0)
  const img = ctx.getImageData(0, 0, w, h)
  const d = img.data
  for (let i = 0; i < w * h; i++) {
    const m = mask[i] ?? 255
    if (m !== 255) d[i * 4 + 3] = Math.round((d[i * 4 + 3] * m) / 255)
  }
  ctx.putImageData(img, 0, 0)
  return { canvas, bounds: { ...bmp.bounds }, scale: bmp.scale, bytes: w * h * 4 }
}
