/**
 * Slideshow plan S1-14 GPU harness (dev only; src/debug is outside the
 * library build). Exposes `window.__slideshowGpu` for
 * tests/browser/slideshow-gpu.spec.ts:
 *
 * - `parity(group)`: one shared scene list rendered on BOTH backends (the
 *   WebGL2 compositor and the Canvas 2D compositor, each on a real browser
 *   canvas), composited over the same opaque backdrop and compared with
 *   `frameDiff` against the scene's tolerance;
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
import { ManualClock } from '../slideshow/clock'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { createRng } from '../slideshow/rng'
import type {
  Affine, CanvasFactory, ColorMatrix, Compositor, LayerBitmap, LayerDraw, LayerHandle, Lease, MaskKind, MaskState, SlideShowDiagnostic,
  SlideShowWorkspace, SubOwner, TargetState, TerminalReserve, Tolerance, TransitionDraw,
} from '../slideshow/types'
import type { PptxShape, PptxSlide } from '../pptx/types'
import { frameDiff, withinTolerance } from './slideshow-parity'

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
  group: 'static' | 'transform' | 'mask-layer' | 'mask-transition' | 'transition'
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
    workspace: ws, terminal, sink, force: opts.force,
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
  const canvases = [...host.querySelectorAll('canvas')]
  const top = canvases[canvases.length - 1]
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
    state: s.mgr.state, backend: s.mgr.current.backend, canvases: canvases.map(c => c.dataset.slideshowBackend ?? '?'),
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

declare global {
  interface Window { __slideshowGpu?: typeof api }
}
const api = {
  groups: () => [...new Set(scenes().map(s => s.group))],
  parity: (group: Scene['group']) => scenes().filter(s => s.group === group).map(compare),
  flippedControl,
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
  start, frame, probe, info, advance,
  tick: (ms: number) => { session!.show.advance(ms); return session!.show.now() },
  pause: () => session!.show.pause(),
  resume: () => session!.show.resume(),
  beginTransition: () => { session!.transition = true },
  lose: () => { session!.ext!.loseContext() },
  restore: () => { session!.ext!.restoreContext() },
  exit: () => { session!.mgr.dispose(); return { state: session!.mgr.state, canvases: document.getElementById('stage')!.querySelectorAll('canvas').length } },
  diagnostics: () => session!.sink.list().map(d => `${d.code}:${d.effect}`),
}
window.__slideshowGpu = api

export type SlideshowGpu = typeof api
