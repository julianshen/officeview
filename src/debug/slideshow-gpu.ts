/**
 * Slideshow plan S1-14 GPU harness (dev only; src/debug is outside the
 * library build). Exposes `window.__slideshowGpu` for
 * tests/browser/slideshow-gpu.spec.ts:
 *
 * - `parity(group)`: one shared scene list rendered on BOTH backends (the
 *   WebGL2 compositor and the Canvas 2D compositor, each on a real browser
 *   canvas), composited over the same opaque backdrop and compared with
 *   `frameDiff` against the scene's tolerance;
 * - `three.*` (S3-1, tests/browser/slideshow-3d.spec.ts): the 3-D transition
 *   framework with a SAMPLE shader (src/debug/slideshow-3d-sample.ts, a card
 *   flip registered for `p14:vortex` in this page only): its ends against the
 *   flat frames, its middle against the look-alike, the Canvas 2D look-alike
 *   path, and recovery from a real context loss;
 * - `start`/`frame`/`probe`/`lose`/`restore`/`advance`/`tick`/`exit`: a
 *   `BackendManager` showing a slide whose red square moves with a manual
 *   show clock, with a manual monotonic clock and scheduler for the restore
 *   deadline. Nothing here reads wall time.
 */
import { createBackendManager, createLivePainter, type BackendManager } from '../slideshow/render/backend'
import { createCanvas2DCompositor } from '../slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../slideshow/render/webgl2'
import { buildRenderPlan, type PlanLayer } from '../slideshow/render/plan'
import { rasterLayer, type RasterContext } from '../slideshow/render/raster'
import { MASK_KINDS } from '../slideshow/render/masks'
import { transitionDraw } from '../slideshow/render/transitions2d'
import { CATALOGUE } from '../slideshow/catalogue'
import { ManualClock } from '../slideshow/clock'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { createRng } from '../slideshow/rng'
import type {
  Affine, CanvasFactory, ColorMatrix, Compositor, LayerBitmap, LayerDraw, LayerHandle, Lease, MaskKind, MaskState, SlideShowDiagnostic,
  SlideShowWorkspace, SubOwner, TargetState, TerminalReserve, Tolerance, TransitionDraw, TransitionSpec,
} from '../slideshow/types'
import type { PptxShape, PptxSlide } from '../pptx/types'
import { frameDiff, withinTolerance } from './slideshow-parity'
import { createSampleTransition3D } from './slideshow-3d-sample'
import { registerTransition3D } from '../slideshow/render/three/registry'
import { dispatchTransition, transition3DDraw } from '../slideshow/render/three/framework'
import { morphFrame } from '../slideshow/morph/render'

// ---------------------------------------------------------------------------
// Workspace ledger (S1-16 owns the real one)
// ---------------------------------------------------------------------------
interface Pool { used: number; limit: number }
function poolLease(pool: Pool, bytes: number, tag: string): Lease | undefined {
  if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
  pool.used += bytes
  let held = bytes, live = true
  const lease: Lease = {
    get tag() { return tag },
    get bytes() { return live ? held : 0 },
    grow(extra) { if (!live || pool.used + extra > pool.limit) return false; pool.used += extra; held += extra; return true },
    shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
    transfer() { return lease },
    release() { if (!live) return; live = false; pool.used -= held },
  }
  return lease
}
function makeWorkspace(limit: number): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  return {
    pool, limit,
    lease: (bytes, tag) => poolLease(pool, bytes, tag),
    used: () => pool.used,
    reserveTerminal(outputBytes) {
      const scratchBytes = 64 * 1024 * 1024
      const whole = poolLease(pool, outputBytes + scratchBytes, 'terminal')
      if (!whole) return undefined
      const sub: Pool = { used: 0, limit: scratchBytes }
      const scratch: SubOwner = { lease: (bytes, tag) => poolLease(sub, bytes, tag) }
      const reserve: TerminalReserve = { outputBytes, scratch, resize: () => false, release: () => whole.release() }
      return reserve
    },
  }
}

const factory: CanvasFactory = (w, h) => {
  const c = document.createElement('canvas')
  c.width = w; c.height = h
  return c
}

// ---------------------------------------------------------------------------
// Geometry: a 48 × 32 slide at 2 device px per slide px, letterboxed by 4 px.
// ---------------------------------------------------------------------------
const SLIDE_W = 48, SLIDE_H = 32
const DEV_W = 96, DEV_H = 72
const VIEW: Affine = [2, 0, 0, 2, 0, 4]
const SCALE = 2
const EXACT: Tolerance = { meanAbs: 0.02, rowMax: 0.10 }
const I: Affine = [1, 0, 0, 1, 0, 0]

type Variant = 'bg' | 'a' | 'b' | 'from' | 'to'
type Bounds = LayerBitmap['bounds']

/** Content with a vertical and horizontal asymmetry, so a flip or a transposition shows. */
function paintVariant(ctx: CanvasRenderingContext2D, w: number, h: number, v: Variant) {
  const g = ctx.createLinearGradient(0, 0, w, h)
  const stops: Record<Variant, [string, string]> = {
    bg: ['#f4f0e8', '#8090a0'], a: ['#ff3020', '#2040ff'], b: ['rgba(20,200,90,0.7)', 'rgba(240,200,0,0.5)'],
    from: ['#204080', '#80c0ff'], to: ['#a02020', '#ffd060'],
  }
  g.addColorStop(0, stops[v][0]); g.addColorStop(1, stops[v][1])
  ctx.fillStyle = g
  ctx.fillRect(0, 0, w, h)
  // Top-left marker and a lower-right disc.
  ctx.fillStyle = v === 'b' ? 'rgba(0,0,0,0.6)' : '#10e010'
  ctx.fillRect(w * 0.08, h * 0.1, w * 0.25, h * 0.2)
  ctx.fillStyle = v === 'bg' ? '#304050' : '#ffffff'
  ctx.beginPath(); ctx.arc(w * 0.7, h * 0.72, Math.min(w, h) * 0.18, 0, Math.PI * 2); ctx.fill()
}

function bitmap(bounds: Bounds, v: Variant): LayerBitmap {
  const w = Math.ceil(bounds.w * SCALE), h = Math.ceil(bounds.h * SCALE)
  const canvas = factory(w, h) as HTMLCanvasElement
  paintVariant(canvas.getContext('2d')!, w, h, v)
  return { canvas, bounds, scale: SCALE, bytes: w * h * 4 }
}

// ---------------------------------------------------------------------------
// The shared parity scenes
// ---------------------------------------------------------------------------
interface Scene {
  name: string
  group: 'static' | 'transform' | 'mask-layer' | 'mask-transition' | 'transition' | 'transition-2d' | 'three'
  tolerance: Tolerance
  layers: Array<{ bounds: Bounds; v: Variant }>
  draws(h: LayerHandle[]): LayerDraw[]
  transition?(h: LayerHandle[]): TransitionDraw
}

const FULL: Bounds = { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H }
const A: Bounds = { x: 6, y: 4, w: 20, h: 14 }
const B: Bounds = { x: 20, y: 12, w: 22, h: 16 }
const st = (over: Partial<TargetState> = {}): TargetState => ({ visible: true, opacity: 1, transform: I, ...over })
/** An affine about the centre of `b`. */
function about(b: Bounds, m: Affine): Affine {
  const cx = b.x + b.w / 2, cy = b.y + b.h / 2
  const [a, bb, c, d] = m
  return [a, bb, c, d, cx - a * cx - c * cy + m[4], cy - bb * cx - d * cy + m[5]]
}
const rot = (deg: number): Affine => { const r = (deg * Math.PI) / 180; return [Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0] }

const MASK_PARAMS: Partial<Record<MaskKind, MaskState['params']>> = {
  wheel: { spokes: 4 }, dissolve: { cell: 3 }, randomBar: { bar: 2 }, split: { orient: 'vert', dir: 'in' },
  circle: { dir: 'in' }, wipe: { from: 'b' }, strips: { from: 'ru' }, blinds: { orient: 'vert' }, checkerboard: { cells: 6 },
}
const SEPIA: ColorMatrix = [0.39, 0.77, 0.19, 0, 0, 0.35, 0.69, 0.17, 0, 0, 0.27, 0.53, 0.13, 0, 0, 0, 0, 0, 1, 0]
const FADE_TINT: ColorMatrix = [1, 0, 0, 0, 0.1, 0, 0.5, 0, 0, 0, 0, 0, 1, 0, 0.2, 0, 0, 0, 0.6, 0.05]

function scenes(): Scene[] {
  const list: Scene[] = []
  const bgDraw = (h: LayerHandle[]): LayerDraw => ({ handle: h[0], state: st() })
  list.push({
    name: 'static', group: 'static', tolerance: EXACT,
    layers: [{ bounds: FULL, v: 'bg' }, { bounds: A, v: 'a' }, { bounds: B, v: 'b' }],
    draws: h => [bgDraw(h), { handle: h[1], state: st() }, { handle: h[2], state: st() }],
  })
  const one = (name: string, state: TargetState, part?: LayerDraw['part'], v: Variant = 'a', bounds = A): Scene => ({
    name, group: 'transform', tolerance: EXACT,
    layers: [{ bounds: FULL, v: 'bg' }, { bounds, v }],
    draws: h => [bgDraw(h), { handle: h[1], state, part }],
  })
  list.push(one('rotate+motion', st({ transform: about(A, rot(25)), motion: { dx: 3, dy: -2 } })))
  list.push(one('scale', st({ transform: about(A, [1.4, 0, 0, 1.4, 0, 0]) })))
  list.push(one('skew-flip', st({ transform: about(A, [-1, 0, 0.3, 1, 0, 0]) })))
  list.push(one('opacity', st({ opacity: 0.45 }), undefined, 'b', B))
  list.push(one('colour-sepia', st({ color: { fill: SEPIA } }), 'fill'))
  list.push(one('colour-alpha-offset', st({ color: { text: FADE_TINT } }), 'text', 'b', B))
  list.push(one('colour-whole-layer-ignored', st({ color: { fill: SEPIA } })))
  list.push(one('opacity+colour+rotate', st({ opacity: 0.6, transform: about(A, rot(-40)), color: { line: SEPIA } }), 'line'))
  for (const kind of MASK_KINDS) {
    const params = MASK_PARAMS[kind] ?? {}
    list.push({
      name: `layer:${kind}`, group: 'mask-layer', tolerance: EXACT,
      layers: [{ bounds: FULL, v: 'bg' }, { bounds: A, v: 'a' }],
      draws: h => [bgDraw(h), { handle: h[1], state: st({ mask: { kind, progress: 0.43, params: { ...params, seed: 5 } }, motion: { dx: 4, dy: 2 } }) }],
    })
    list.push({
      name: `transition:${kind}`, group: 'mask-transition', tolerance: EXACT,
      layers: [{ bounds: FULL, v: 'from' }, { bounds: FULL, v: 'to' }],
      draws: () => [],
      transition: h => ({ from: h[0], to: h[1], effect: kind, progress: 0.43, params: { mask: { kind, progress: 0.43, params } }, rng: createRng(11) }),
    })
  }
  const tr = (name: string, params: Record<string, unknown>, extra?: Variant): Scene => ({
    name, group: 'transition', tolerance: EXACT,
    layers: [{ bounds: FULL, v: 'from' }, { bounds: FULL, v: 'to' }, ...(extra ? [{ bounds: B, v: extra }] : [])],
    draws: h => (extra ? [{ handle: h[2], state: st() }] : []),
    transition: h => ({ from: h[0], to: h[1], effect: name, progress: 0.35, params, rng: createRng(3) }),
  })
  list.push(tr('push', { from: [1, 0, 0, 1, -0.35, 0], to: [1, 0, 0, 1, 0.65, 0] }))
  list.push(tr('cover-zoom', { to: [0.6, 0, 0, 0.6, 0.2, 0.25], toOpacity: 0.7 }))
  list.push(tr('fade', { toOpacity: 0.4 }))
  list.push(tr('wipe+draws-on-top', { mask: { kind: 'wipe', progress: 0.35, params: { from: 'r' } } }, 'b'))
  // S1-15: every 2-D catalogue row (and the attribute variants with their own geometry) through
  // transitionDraw at 0/25/50/75/100 %, compared at the row's own tolerance.
  const specOf = (element: string, ns: string, attrs: Record<string, string> = {}): TransitionSpec =>
    ({ element, ns: ns as TransitionSpec['ns'], attrs, durationMs: 500, advClick: true })
  const rows = CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 1)
  const variants: Array<{ name: string; spec: TransitionSpec; tolerance: Tolerance }> = rows.map(r => ({ name: r.id, spec: specOf(r.element!, r.ns!), tolerance: r.tolerance }))
  const tol = (element: string) => rows.find(r => r.element === element)!.tolerance
  for (const [element, attrs] of [
    ['cut', { thruBlk: '1' }], ['fade', { thruBlk: '1' }], ['zoom', { dir: 'in' }], ['pull', { dir: 'ld' }], ['cover', { dir: 'ru' }],
    ['split', { orient: 'vert', dir: 'in' }], ['wipe', { dir: 'd' }], ['strips', { dir: 'rd' }], ['checker', { dir: 'vert' }],
  ] as const) variants.push({ name: `transition:p:${element} ${JSON.stringify(attrs)}`, spec: specOf(element, 'p', attrs), tolerance: tol(element) })
  for (const v of variants) {
    for (const p of [0, 0.25, 0.5, 0.75, 1]) {
      list.push({
        name: `${v.name} @${p}`, group: 'transition-2d', tolerance: v.tolerance,
        layers: [{ bounds: FULL, v: 'from' }, { bounds: FULL, v: 'to' }],
        draws: () => [],
        transition: h => {
          const rng = createRng(17)
          return { from: h[0], to: h[1], effect: v.spec.element, progress: p, params: transitionDraw(v.spec, p, rng, SLIDE_W / SLIDE_H), rng }
        },
      })
    }
  }
  return list
}

function readBack(out: HTMLCanvasElement): Uint8ClampedArray {
  const r = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const ctx = r.getContext('2d', { willReadFrequently: true })!
  ctx.fillStyle = '#202020'
  ctx.fillRect(0, 0, DEV_W, DEV_H)
  // Same task as the frame: the WebGL drawing buffer is still there without preserveDrawingBuffer.
  ctx.drawImage(out, 0, 0)
  return ctx.getImageData(0, 0, DEV_W, DEV_H).data
}

function render(backend: 'webgl2' | 'canvas2d', scene: Scene): { pixels: Uint8ClampedArray; diagnostics: SlideShowDiagnostic[] } {
  const ws = makeWorkspace(256 * 1024 * 1024)
  const sink = createDiagnosticSink()
  const out = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const comp: Compositor | undefined = backend === 'webgl2' ? createWebGL2Compositor(out, ws, { sink }) : createCanvas2DCompositor(out, factory, ws, sink)
  if (!comp) throw new Error(`no ${backend} compositor`)
  try {
    comp.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
    const handles = scene.layers.map(l => {
      const b = bitmap(l.bounds, l.v)
      const lease = ws.lease(b.bytes, 'layer')!
      const r = comp.upload(b, lease)
      if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
      return r
    })
    comp.frame(scene.draws(handles), scene.transition?.(handles))
    return { pixels: readBack(out), diagnostics: [...sink.list()] }
  } finally {
    comp.dispose()
    if (ws.used() !== 0) throw new Error(`${backend} leaked ${ws.used()} bytes`)
  }
}

interface ParityResult { name: string; meanAbs: number; rowMax: number; ok: boolean; tolerance: Tolerance; diagnostics: string[] }

function compare(scene: Scene): ParityResult {
  const gl = render('webgl2', scene)
  const c2 = render('canvas2d', scene)
  const d = frameDiff(gl.pixels, c2.pixels, DEV_W, DEV_H)
  const codes = (l: SlideShowDiagnostic[]) => l.map(x => `${x.code}:${x.effect}`).sort().join(',')
  const sameDiagnostics = codes(gl.diagnostics) === codes(c2.diagnostics)
  return {
    name: scene.name, meanAbs: d.meanAbs, rowMax: d.rowMax, tolerance: scene.tolerance,
    ok: withinTolerance(d, scene.tolerance) && sameDiagnostics, diagnostics: [codes(gl.diagnostics), codes(c2.diagnostics)],
  }
}

/** A deliberately wrong pairing (a vertically flipped scene) proves the comparison can fail. */
function flippedControl(): ParityResult {
  const base = scenes().find(s => s.name === 'static')!
  const flipped: Scene = { ...base, draws: h => base.draws(h).map(d => ({ ...d, state: { ...d.state, transform: [1, 0, 0, -1, 0, SLIDE_H] as Affine } })) }
  const gl = render('webgl2', flipped)
  const c2 = render('canvas2d', base)
  const d = frameDiff(gl.pixels, c2.pixels, DEV_W, DEV_H)
  return { name: 'flipped-control', meanAbs: d.meanAbs, rowMax: d.rowMax, tolerance: EXACT, ok: withinTolerance(d, EXACT), diagnostics: [] }
}

// ---------------------------------------------------------------------------
// A real slide through the S1-12 rasteriser
// ---------------------------------------------------------------------------
const PX = 9525
function shape(id: number, geometry: PptxShape['geometry'], x: number, y: number, w: number, h: number, fill: string, more: Partial<PptxShape> = {}): PptxShape {
  return {
    xEmu: x * PX, yEmu: y * PX, widthEmu: w * PX, heightEmu: h * PX, geometry, fill,
    source: { partPath: 'ppt/slides/slide1.xml', treePath: `spTree/sp[${id}]`, element: 'sp', id: String(id), representation: 'native' }, ...more,
  }
}
const STATIC_SLIDE: PptxSlide = {
  index: 0, widthEmu: SLIDE_W * PX, heightEmu: SLIDE_H * PX, background: '#f6f2ea',
  shapes: [
    shape(2, 'rect', 3, 3, 26, 12, '#3366cc'),
    shape(3, 'ellipse', 18, 8, 20, 16, '#cc3333'),
    shape(4, 'roundRect', 8, 18, 30, 10, '#33aa55', { rotationDeg: 12 }),
  ],
}

function renderSlideOn(backend: 'webgl2' | 'canvas2d', slide: PptxSlide, animated: ReadonlySet<string>): Uint8ClampedArray {
  const ws = makeWorkspace(256 * 1024 * 1024)
  const out = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const comp: Compositor | undefined = backend === 'webgl2' ? createWebGL2Compositor(out, ws) : createCanvas2DCompositor(out, factory, ws)
  if (!comp) throw new Error(`no ${backend} compositor`)
  try {
    comp.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
    const draws = uploadPlan(comp, ws, slide, animated).map(l => ({ handle: l.handle, state: st() }))
    comp.frame(draws)
    return readBack(out)
  } finally { comp.dispose() }
}

function uploadPlan(comp: Compositor, ws: SlideShowWorkspace, slide: PptxSlide, animated: ReadonlySet<string>): Array<{ layer: PlanLayer; handle: LayerHandle }> {
  const plan = buildRenderPlan(slide, animated, new Map(), [])
  const ctx: RasterContext = { scale: SCALE, images: [], resolveFont: f => f, canvas: factory, workspace: ws, maxTextureSize: comp.maxTextureSize }
  const out: Array<{ layer: PlanLayer; handle: LayerHandle }> = []
  for (const layer of plan.layers) {
    const r = rasterLayer(layer, slide, ctx)
    if ('refused' in r) { if (r.refused === 'empty') continue; throw new Error('raster refused') }
    const u = comp.upload(r, r.lease)
    if ('refused' in u) { r.lease.release(); throw new Error(`upload refused ${u.refused}`) }
    out.push({ layer, handle: u })
  }
  return out
}

// ---------------------------------------------------------------------------
// The backend manager under a manual show clock and monotonic clock
// ---------------------------------------------------------------------------
const MOVER_X = 2, MOVER_Y = 13, MOVER = 6
const MOVING_SLIDE: PptxSlide = {
  index: 0, widthEmu: SLIDE_W * PX, heightEmu: SLIDE_H * PX, background: '#ffffff',
  shapes: [shape(2, 'rect', 0, 24, SLIDE_W, 6, '#808080'), shape(3, 'rect', MOVER_X, MOVER_Y, MOVER, MOVER, '#ff0000')],
}
/** Slide px the square has moved at show time `t` ms: 1 px per 50 ms, wrapping before the edge. */
const moverDx = (t: number) => Math.floor(t / 50) % 36

interface Session {
  mgr: BackendManager
  show: ManualClock
  layers: Array<{ layer: PlanLayer; handle: LayerHandle }>
  /** A full-slide frame of the "previous slide" to transition from. */
  prev?: LayerHandle
  transition: boolean
  perf: { now: number; queue: Array<{ at: number; cb: () => void }> }
  sink: ReturnType<typeof createDiagnosticSink>
  ext: WEBGL_lose_context | null
  glCanvas: HTMLCanvasElement | null
  events: { lost: number; restored: number }
  errors: string[]
}
let session: Session | undefined

function moverState(show: ManualClock): Map<string, TargetState> {
  return new Map([['sp:3', st({ transform: [1, 0, 0, 1, moverDx(show.now()), 0] })]])
}

async function build(c: Compositor, s: { show: ManualClock; layers: Session['layers']; prev?: LayerHandle }, ws: SlideShowWorkspace) {
  s.layers = uploadPlan(c, ws, MOVING_SLIDE, new Set(['sp:3']))
  const b = bitmap(FULL, 'from')
  const lease = ws.lease(b.bytes, 'prev')!
  const r = c.upload(b, lease)
  if ('refused' in r) { lease.release(); s.prev = undefined } else s.prev = r
}

function start(opts: { force?: 'canvas2d' } = {}) {
  session?.mgr.dispose()
  const host = document.getElementById('stage')!
  host.replaceChildren()
  const ws = makeWorkspace(256 * 1024 * 1024)
  const terminal = ws.reserveTerminal(DEV_W * DEV_H * 4)!
  // The terminal reserve's output canvas, under the backend's canvases (as on the audience surface): the
  // Canvas 2D fallback composites on it (PR #18 review P2).
  const output = factory(DEV_W, DEV_H) as HTMLCanvasElement
  output.style.cssText = 'position: absolute; left: 0; top: 0; width: 100%; height: 100%; display: block'
  host.append(output)
  const show = new ManualClock()
  const sink = createDiagnosticSink()
  const perf = { now: 0, queue: [] as Session['perf']['queue'] }
  const partial = { show, layers: [] as Session['layers'], prev: undefined as LayerHandle | undefined }
  const paintLive = createLivePainter(terminal, () => ({ slide: MOVING_SLIDE, state: moverState(show), assets: { images: [], resolveFont: (f: string) => f } }))
  const mgr = createBackendManager(host, {
    now: () => perf.now,
    schedule: (cb, ms) => { perf.queue.push({ at: perf.now + ms, cb }); return perf.queue.length },
    paintLive,
    rebuildLayers: async c => { await build(c, partial, ws); if (session) { session.layers = partial.layers; session.prev = partial.prev } },
    workspace: ws, terminal, output, sink, force: opts.force,
  })
  mgr.current.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
  void build(mgr.current, partial, ws)
  const glCanvas = host.querySelector<HTMLCanvasElement>('canvas[data-slideshow-backend="webgl2"]')
  const gl = glCanvas?.getContext('webgl2') ?? null
  const events = { lost: 0, restored: 0 }
  glCanvas?.addEventListener('webglcontextlost', () => { events.lost++ })
  glCanvas?.addEventListener('webglcontextrestored', () => { events.restored++ })
  session = { mgr, show, layers: partial.layers, prev: partial.prev, transition: false, perf, sink, ext: gl?.getExtension('WEBGL_lose_context') ?? null, glCanvas, events, errors: [] }
  return info()
}

function frame() {
  const s = session!
  const state = moverState(s.show)
  const draws: LayerDraw[] = s.layers.map(l => ({ handle: l.handle, state: l.layer.kind === 'target' ? state.get(l.layer.key) ?? st() : st() }))
  let transition: TransitionDraw | undefined
  if (s.transition && s.prev && draws.length > 0) {
    // The previous slide wipes away to the background layer; the rest draws on top.
    const p = Math.min(1, s.show.now() / 4000)
    transition = { from: s.prev, to: draws[0].handle, effect: 'wipe', progress: p, params: { mask: { kind: 'wipe', progress: p, params: {} } }, rng: createRng(1) }
    draws.shift()
  }
  s.mgr.current.frame(draws, transition)
}

/** Frame, then read the visible canvas in the same task; `redX` is the square's left edge in device px. */
function probe() {
  const s = session!
  frame()
  const host = document.getElementById('stage')!
  // The backends' canvases, bottom to top: the output (while Canvas 2D draws on it), then the WebGL one.
  const canvases = [...host.querySelectorAll<HTMLCanvasElement>('canvas[data-slideshow-backend]')]
  // The visible one: the output while it is drawn on (a lost WebGL canvas over it is hidden), else the WebGL canvas.
  const top = canvases.find(c => c.dataset.slideshowBackend === 'canvas2d') ?? canvases[canvases.length - 1]
  let redX = -1
  let grey: number[] = []
  if (top) {
    const r = factory(top.width, top.height) as HTMLCanvasElement
    const ctx = r.getContext('2d', { willReadFrequently: true })!
    ctx.drawImage(top, 0, 0)
    const y = Math.round(VIEW[3] * (MOVER_Y + MOVER / 2) + VIEW[5])
    const row = ctx.getImageData(0, y, top.width, 1).data
    for (let x = 0; x < top.width; x++) {
      if (row[x * 4] > 200 && row[x * 4 + 1] < 60 && row[x * 4 + 2] < 60) { redX = x; break }
    }
    grey = [...ctx.getImageData(10, Math.round(VIEW[3] * 27 + VIEW[5]), 1, 1).data]
  }
  return {
    state: s.mgr.state, backend: s.mgr.current.backend, canvases: canvases.map(c => c.dataset.slideshowBackend!),
    redX, expectedX: Math.round(VIEW[0] * (MOVER_X + moverDx(s.show.now())) + VIEW[4]), grey, showTime: s.show.now(), events: { ...s.events },
  }
}

function info() {
  const s = session
  let renderer = ''
  const gl = s?.glCanvas?.getContext('webgl2')
  if (gl) {
    const dbg = gl.getExtension('WEBGL_debug_renderer_info')
    renderer = String(dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER))
  }
  return { state: s?.mgr.state, backend: s?.mgr.current.backend, maxTextureSize: s?.mgr.current.maxTextureSize, renderer, hasLoseContext: !!s?.ext }
}

function advance(ms: number) {
  const p = session!.perf
  p.now += ms
  for (;;) {
    const due = p.queue.filter(q => q.at <= p.now).sort((a, b) => a.at - b.at)[0]
    if (!due) break
    p.queue.splice(p.queue.indexOf(due), 1)
    due.cb()
  }
  return p.now
}

// ---------------------------------------------------------------------------
// S3-1: the 3-D transition framework with the sample shader
// ---------------------------------------------------------------------------
const sample3D = createSampleTransition3D('vortex')
const VORTEX: TransitionSpec = { element: 'vortex', ns: 'p14', attrs: { dir: 'r' }, durationMs: 1000, advClick: true }
const DISSOLVE: TransitionSpec = { element: 'dissolve', ns: 'p', attrs: {}, durationMs: 1000, advClick: true }

/** One transition frame over the from / to frames, with params made from the effect rng (seed 21). */
function transitionScene(name: string, progress: number, make: (rng: ReturnType<typeof createRng>) => Record<string, unknown>): Scene {
  return {
    name, group: 'three', tolerance: EXACT,
    layers: [{ bounds: FULL, v: 'from' }, { bounds: FULL, v: 'to' }],
    draws: () => [],
    transition: h => { const rng = createRng(21); return { from: h[0], to: h[1], effect: 'vortex', progress, params: make(rng), rng } },
  }
}
const shaderAt = (p: number) => transitionScene(`vortex 3-D @${p}`, p, rng => transition3DDraw(VORTEX, p, rng, SLIDE_W / SLIDE_H))
const lookalikeAt = (p: number) => transitionScene(`dissolve @${p}`, p, rng => transitionDraw(DISSOLVE, p, rng, SLIDE_W / SLIDE_H))
/** The flat outgoing frame (the incoming one hidden) and the flat incoming frame. */
const flat = (which: 'from' | 'to') => transitionScene(`flat ${which}`, 0, () => (which === 'from' ? { toOpacity: 0 } : {}))

function px(pixels: Uint8ClampedArray, x: number, y: number): number[] {
  return [...pixels.slice((y * DEV_W + x) * 4, (y * DEV_W + x) * 4 + 4)]
}

const three = {
  /** The sample shader at and near its ends against the flat frames (WebGL2). */
  ends() {
    registerTransition3D(sample3D.transition)
    return ([[0, 'from'], [1e-3, 'from'], [0.999, 'to'], [1, 'to']] as const).map(([p, which]) => {
      const before = sample3D.calls.length
      const got = render('webgl2', shaderAt(p))
      const want = render('webgl2', flat(which))
      const d = frameDiff(got.pixels, want.pixels, DEV_W, DEV_H)
      return { progress: p, which, meanAbs: d.meanAbs, rowMax: d.rowMax, ok: withinTolerance(d, EXACT), shaderCalls: sample3D.calls.length - before, diagnostics: got.diagnostics.map(x => x.code) }
    })
  },
  /** The sample shader mid-way: a turned card (black where it no longer covers), not the look-alike. */
  middle() {
    registerTransition3D(sample3D.transition)
    return [0.25, 0.75].map((p) => {
      const before = sample3D.calls.length
      const got = render('webgl2', shaderAt(p))
      const look = render('webgl2', lookalikeAt(p))
      const fromFlat = render('webgl2', flat('from')).pixels, toFlat = render('webgl2', flat('to')).pixels
      const d = frameDiff(got.pixels, look.pixels, DEV_W, DEV_H)
      const y = DEV_H / 2
      return {
        progress: p, shaderCalls: sample3D.calls.length - before, vsLookalike: d.meanAbs,
        // Device px: the slide is x 0..96, y 4..68. The flip's far edge leaves the right of the slide black.
        left: px(got.pixels, 6, y), right: px(got.pixels, 90, y), centre: px(got.pixels, 40, y), fromCentre: px(fromFlat, 40, y), toCentre: px(toFlat, 40, y),
        letterbox: px(got.pixels, 40, 1), diagnostics: got.diagnostics.map(x => x.code),
      }
    })
  },
  /** Canvas 2D: dispatch says look-alike with a diagnostic, and the shader params draw the look-alike (parity with WebGL2's own look-alike). */
  canvas2d() {
    registerTransition3D(sample3D.transition)
    const ws = makeWorkspace(64 * 1024 * 1024)
    const sink = createDiagnosticSink()
    const c2 = createCanvas2DCompositor(factory(DEV_W, DEV_H), factory, ws, sink)
    const route = dispatchTransition(VORTEX, c2, sink, 4)
    c2.dispose()
    const glSink = createDiagnosticSink()
    const glOut = factory(DEV_W, DEV_H) as HTMLCanvasElement
    const gl = createWebGL2Compositor(glOut, makeWorkspace(64 * 1024 * 1024), { sink: glSink })
    const glRoute = gl ? dispatchTransition(VORTEX, gl, glSink, 4) : 'none'
    gl?.dispose()
    const results = [0.25, 0.5, 0.75].map((p) => {
      const before = sample3D.calls.length
      const c = render('canvas2d', shaderAt(p))
      const w = render('webgl2', lookalikeAt(p))
      const d = frameDiff(c.pixels, w.pixels, DEV_W, DEV_H)
      return { progress: p, meanAbs: d.meanAbs, rowMax: d.rowMax, ok: withinTolerance(d, EXACT), shaderCalls: sample3D.calls.length - before }
    })
    return { route, glRoute, diagnostics: sink.list().map(d => ({ code: d.code, slideIndex: d.slideIndex, effect: d.effect, message: d.message })), glDiagnostics: glSink.list().length, results }
  },
  /** A real context loss in the middle of a 3-D transition: the pool is dropped, then rebuilt, and the frame is the same. */
  async contextLoss() {
    registerTransition3D(sample3D.transition)
    const ws = makeWorkspace(64 * 1024 * 1024)
    const sink = createDiagnosticSink()
    const out = factory(DEV_W, DEV_H) as HTMLCanvasElement
    document.getElementById('stage')!.replaceChildren(out)
    out.addEventListener('webglcontextlost', e => e.preventDefault())
    const steps: string[] = []
    // A failure guard only (the spec never waits on time): a missing browser event fails with the steps reached.
    const within = (what: string, p: Promise<void>) => {
      let timer: ReturnType<typeof setTimeout> | undefined
      const guard = new Promise<void>((_, reject) => { timer = setTimeout(() => reject(new Error(`${what} timed out after ${steps.join(' > ')}`)), 20000) })
      return Promise.race([p, guard]).finally(() => clearTimeout(timer))
    }
    const comp = createWebGL2Compositor(out, ws, { sink })
    if (!comp) throw new Error('no webgl2 compositor')
    const ext = out.getContext('webgl2')!.getExtension('WEBGL_lose_context')
    if (!ext) throw new Error('no WEBGL_lose_context')
    comp.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
    const scene = shaderAt(0.3)
    const upload = () => scene.layers.map((l) => {
      const b = bitmap(l.bounds, l.v)
      const r = comp.upload(b, ws.lease(b.bytes, 'layer')!)
      if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
      return r
    })
    const shoot = () => {
      const h = upload()
      comp.frame([], scene.transition!(h))
      const pixels = readBack(out)
      for (const x of h) comp.release(x)
      return pixels
    }
    const routeBefore = dispatchTransition(VORTEX, comp, sink, 0)
    const before = shoot()
    const heldBefore = ws.used()
    const lost = new Promise<void>(r => out.addEventListener('webglcontextlost', () => r(), { once: true }))
    steps.push('lose')
    ext.loseContext()
    await within('webglcontextlost', lost)
    steps.push('lost')
    comp.contextLost()
    const heldLost = ws.used()
    const routeLost = dispatchTransition(VORTEX, comp, sink, 0)
    const restored = new Promise<void>(r => out.addEventListener('webglcontextrestored', () => r(), { once: true }))
    // Restore from a later task, as a page would (Chromium ignores a restore requested inside the loss dispatch).
    await new Promise<void>(r => setTimeout(r, 0))
    ext.restoreContext()
    steps.push('restore')
    await within('webglcontextrestored', restored)
    const rebuilt = comp.contextRestored()
    const routeAfter = dispatchTransition(VORTEX, comp, sink, 1)
    const after = shoot()
    const d = frameDiff(before, after, DEV_W, DEV_H)
    comp.dispose()
    out.remove()
    return { routeBefore, routeLost, routeAfter, rebuilt, heldBefore, heldLost, heldEnd: ws.used(), meanAbs: d.meanAbs, rowMax: d.rowMax, same: withinTolerance(d, EXACT) }
  },
}

/**
 * PR #20 review 2: a same-outline Morph pair of two identical translucent
 * shapes (authored 50 % alpha, then opaque bitmaps at 50 % state opacity),
 * at p = 0.25, 0.5 and 0.9, on a transparent output: the alpha inside the
 * shape (device 28,16; the top band) and below it (device 28,50), raw.
 */
function morphSum(backend: 'webgl2' | 'canvas2d') {
  const ws = makeWorkspace(256 * 1024 * 1024)
  const sink = createDiagnosticSink()
  const out = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const comp: Compositor | undefined = backend === 'webgl2' ? createWebGL2Compositor(out, ws, { sink }) : createCanvas2DCompositor(out, factory, ws, sink)
  if (!comp) throw new Error(`no ${backend} compositor`)
  const results: Array<{ authored: number; opacity: number; p: number; inside: number; below: number }> = []
  try {
    comp.setViewport({ matrix: VIEW, deviceWidth: DEV_W, deviceHeight: DEV_H })
    const bounds = { x: 4, y: 2, w: 20, h: 8 }
    const g = { ...bounds, rot: 0 }
    const up = (alpha: number): LayerHandle => {
      const c = factory(bounds.w * SCALE, bounds.h * SCALE) as HTMLCanvasElement
      const x = c.getContext('2d')!
      x.globalAlpha = alpha
      x.fillStyle = 'rgb(40,120,200)'
      x.fillRect(0, 0, c.width, c.height)
      const bytes = c.width * c.height * 4
      const r = comp.upload({ canvas: c, bounds, scale: SCALE, bytes }, ws.lease(bytes, 'layer')!)
      if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
      return r
    }
    for (const [authored, opacity] of [[0.5, 1], [1, 0.5]] as const) {
      const fh = up(authored), th = up(authored)
      const state = (): TargetState => ({ visible: true, opacity, transform: I })
      for (const p of [0.25, 0.5, 0.9]) {
        morphFrame({ pairs: [{ from: { handle: fh, geometry: g, state: state() }, to: { handle: th, geometry: g, state: state() }, under: true }] }, p, comp)
        const r = factory(DEV_W, DEV_H) as HTMLCanvasElement
        const ctx = r.getContext('2d', { willReadFrequently: true })!
        ctx.drawImage(out, 0, 0)
        const a = (x: number, y: number) => ctx.getImageData(x, y, 1, 1).data[3]!
        results.push({ authored, opacity, p, inside: a(28, 16), below: a(28, 50) })
      }
      comp.release(fh)
      comp.release(th)
    }
    return { results, diagnostics: sink.list().map(d => `${d.code}:${d.effect}`) }
  } finally {
    comp.dispose()
    if (ws.used() !== 0) throw new Error(`${backend} leaked ${ws.used()} bytes`)
  }
}

declare global {
  interface Window { __slideshowGpu?: typeof api }
}
const api = {
  groups: () => [...new Set(scenes().map(s => s.group))],
  parity: (group: Scene['group']) => scenes().filter(s => s.group === group).map(compare),
  flippedControl,
  morphSum,
  staticSlide() {
    const animated = new Set(['sp:3'])
    const gl = renderSlideOn('webgl2', STATIC_SLIDE, animated)
    const c2 = renderSlideOn('canvas2d', STATIC_SLIDE, animated)
    const d = frameDiff(gl, c2, DEV_W, DEV_H)
    const opaque = [...gl].filter((_v, i) => i % 4 === 3).every(a => a === 255)
    // A sample inside the blue rectangle proves something was painted on the GPU path at all.
    const at = (x: number, y: number) => [...gl.slice((y * DEV_W + x) * 4, (y * DEV_W + x) * 4 + 4)]
    return { ...d, ok: withinTolerance(d, EXACT), opaque, blue: at(2 * 5, 2 * 5 + 4), tolerance: EXACT }
  },
  three,
  start, frame, probe, info, advance,
  tick: (ms: number) => { session!.show.advance(ms); return session!.show.now() },
  pause: () => session!.show.pause(),
  resume: () => session!.show.resume(),
  beginTransition: () => { session!.transition = true },
  lose: () => { session!.ext!.loseContext() },
  restore: () => { session!.ext!.restoreContext() },
  exit: () => { session!.mgr.dispose(); return { state: session!.mgr.state, canvases: document.getElementById('stage')!.querySelectorAll('canvas[data-slideshow-backend]').length } },
  diagnostics: () => session!.sink.list().map(d => `${d.code}:${d.effect}`),
}
window.__slideshowGpu = api

export type SlideshowGpu = typeof api
