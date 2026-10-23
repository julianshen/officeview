/**
 * Slideshow plan S3-4 harness (dev only; src/debug is outside the library
 * build): the deformation 3-D transitions (src/slideshow/render/three/deform.ts)
 * on a real WebGL2 compositor, for tests/browser/slideshow-3d-deform.spec.ts.
 * Exposes `window.__slideshowDeform`:
 *
 * - `ends(name, attrs)`: the shader just inside its ends against the flat frames;
 * - `fractions(name, attrs, p)`: solid frames (from blue, to red), pixels classified as to / back / from;
 * - `probe(name, attrs, p, points)`: solid frames, the RGBA at NDC points;
 * - `differs(name, a, b, p)`: whether two attribute sets / effect seeds draw different frames;
 * - `canvas2d(name)`: dispatch on Canvas 2D (the look-alike, diagnosed);
 * - `strip(name, attrs)`: 0 / 25 / 50 / 75 / 100 % frames into `#strip` (Mac goldens);
 * - `perf(name, frames)`: informal per-frame times, never asserted.
 *
 * The deformation rows are registered in this page only (`registerDeform3D`).
 */
import { createCanvas2DCompositor } from '../slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../slideshow/render/webgl2'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { createRng } from '../slideshow/rng'
import { effectSeed } from '../slideshow/render/three/mesh'
import { dispatchTransition, transition3DDraw } from '../slideshow/render/three/framework'
import { transition3DFor } from '../slideshow/render/three/registry'
import { DEFORM_ROWS, registerDeform3D } from '../slideshow/render/three/deform'
import type { CanvasFactory, Compositor, LayerBitmap, Lease, Rng, SlideShowWorkspace, Tolerance, TransitionSpec } from '../slideshow/types'
import { frameDiff, withinTolerance } from './slideshow-parity'

registerDeform3D()

interface Pool { used: number; limit: number }
function makeWorkspace(limit = 256 * 1024 * 1024): SlideShowWorkspace {
  const pool: Pool = { used: 0, limit }
  const lease = (bytes: number, tag: string): Lease | undefined => {
    if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
    pool.used += bytes
    let held = bytes, live = true
    const l: Lease = {
      get tag() { return tag },
      get bytes() { return live ? held : 0 },
      grow(extra) { if (!live || pool.used + extra > pool.limit) return false; pool.used += extra; held += extra; return true },
      shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
      transfer() { return l },
      release() { if (!live) return; live = false; pool.used -= held },
    }
    return l
  }
  return { limit, lease, used: () => pool.used, reserveTerminal: () => undefined }
}
const factory: CanvasFactory = (w, h) => { const c = document.createElement('canvas'); c.width = w; c.height = h; return c }

/** A 16 : 9 slide filling a 320 × 180 device canvas. */
const DEV_W = 320, DEV_H = 180
const EXACT: Tolerance = { meanAbs: 0.02, rowMax: 0.10 }
type Paint = 'blue' | 'red' | 'from' | 'to'

function paint(ctx: CanvasRenderingContext2D, w: number, h: number, v: Paint) {
  if (v === 'blue' || v === 'red') { ctx.fillStyle = v === 'blue' ? '#0000ff' : '#ff0000'; ctx.fillRect(0, 0, w, h); return }
  const g = ctx.createLinearGradient(0, 0, w, h)
  const stops = v === 'from' ? ['#204080', '#80c0ff'] : ['#a02020', '#ffd060']
  g.addColorStop(0, stops[0]); g.addColorStop(1, stops[1])
  ctx.fillStyle = g
  ctx.fillRect(0, 0, w, h)
  ctx.fillStyle = '#10e010'
  ctx.fillRect(w * 0.08, h * 0.1, w * 0.25, h * 0.2)
  ctx.fillStyle = '#ffffff'
  ctx.beginPath(); ctx.arc(w * 0.7, h * 0.72, Math.min(w, h) * 0.18, 0, Math.PI * 2); ctx.fill()
}
function bitmap(v: Paint): LayerBitmap {
  const canvas = factory(DEV_W, DEV_H) as HTMLCanvasElement
  paint(canvas.getContext('2d')!, DEV_W, DEV_H, v)
  return { canvas, bounds: { x: 0, y: 0, w: DEV_W, h: DEV_H }, scale: 1, bytes: DEV_W * DEV_H * 4 }
}

function specOf(name: string, attrs: Record<string, string> = {}): TransitionSpec {
  const row = DEFORM_ROWS.find(r => (r.preset ?? r.element) === name)
  if (!row) throw new Error(`not a deformation row: ${name}`)
  return { element: row.element, ns: row.element === 'prstTrans' ? 'p15' : 'p14', attrs, durationMs: 1000, advClick: true, ...(row.preset !== undefined ? { preset: row.preset } : {}) }
}

function readBack(out: HTMLCanvasElement): Uint8ClampedArray {
  const r = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const ctx = r.getContext('2d', { willReadFrequently: true })!
  ctx.drawImage(out, 0, 0)
  return ctx.getImageData(0, 0, DEV_W, DEV_H).data
}

interface Shot { pixels: Uint8ClampedArray; diagnostics: string[]; shaderCalls: number; seed?: number; out: HTMLCanvasElement }
/** One frame of `spec` at `p` (`flat`: the flat outgoing / incoming frame instead), on a fresh WebGL2 compositor. */
function shoot(spec: TransitionSpec, p: number, paints: [Paint, Paint], opts: { flat?: 'from' | 'to'; rngSeed?: number; keep?: boolean } = {}): Shot {
  const ws = makeWorkspace()
  const sink = createDiagnosticSink()
  const out = factory(DEV_W, DEV_H) as HTMLCanvasElement
  const comp = createWebGL2Compositor(out, ws, { sink })
  if (!comp) throw new Error('no webgl2 compositor')
  const t = transition3DFor(spec)!
  let shaderCalls = 0
  const draw = t.draw
  t.draw = function (...args) { shaderCalls++; return draw.apply(this, args) }
  try {
    comp.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: DEV_W, deviceHeight: DEV_H })
    const route = opts.flat ? 'flat' : dispatchTransition(spec, comp, sink, 0)
    if (!opts.flat && route !== 'shader') throw new Error(`${spec.preset ?? spec.element} dispatched to ${route}: ${sink.list().map(d => d.message).join('; ')}`)
    const [from, to] = paints.map((v) => {
      const b = bitmap(v)
      const r = comp.upload(b, ws.lease(b.bytes, 'layer')!)
      if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
      return r
    })
    // Records the effect rng's values: the compositor seeds the shader rng from the one it draws.
    const values: number[] = []
    const base = createRng(opts.rngSeed ?? 21)
    const rng: Rng = () => { const v = base(); values.push(v); return v }
    const params = opts.flat ? (opts.flat === 'from' ? { toOpacity: 0 } : {}) : transition3DDraw(spec, p, rng, DEV_W / DEV_H)
    comp.frame([], { from, to, effect: spec.element, progress: opts.flat ? 0 : p, params, rng })
    const pixels = readBack(out)
    const seed = values.length ? effectSeed(createRng(Math.floor(values[values.length - 1] * 0x100000000))) : undefined
    comp.release(from); comp.release(to)
    return { pixels, diagnostics: sink.list().map(d => d.code), shaderCalls, ...(seed !== undefined ? { seed } : {}), out }
  } finally {
    t.draw = draw
    comp.dispose()
    if (ws.used() !== 0) throw new Error(`leaked ${ws.used()} bytes`)
  }
}

const at = (px: Uint8ClampedArray, x: number, y: number) => [...px.slice((y * DEV_W + x) * 4, (y * DEV_W + x) * 4 + 4)]
const ndcPx = (nx: number, ny: number): [number, number] => [
  Math.min(DEV_W - 1, Math.max(0, Math.floor(((nx + 1) / 2) * DEV_W))),
  Math.min(DEV_H - 1, Math.max(0, Math.floor(((1 - ny) / 2) * DEV_H))),
]

const api = {
  rows: () => DEFORM_ROWS.map(r => r.preset ?? r.element),
  ends(name: string, attrs: Record<string, string> = {}) {
    const spec = specOf(name, attrs)
    return ([[1e-3, 'from'], [0.999, 'to']] as const).map(([p, which]) => {
      const got = shoot(spec, p, ['from', 'to'])
      const want = shoot(spec, 0, ['from', 'to'], { flat: which })
      const d = frameDiff(got.pixels, want.pixels, DEV_W, DEV_H)
      return { progress: p, which, meanAbs: d.meanAbs, rowMax: d.rowMax, ok: withinTolerance(d, EXACT), shaderCalls: got.shaderCalls, diagnostics: got.diagnostics }
    })
  },
  fractions(name: string, attrs: Record<string, string>, p: number) {
    const got = shoot(specOf(name, attrs), p, ['blue', 'red'])
    let to = 0, back = 0, from = 0, black = 0
    for (let i = 0; i < got.pixels.length; i += 4) {
      const [r, g, b] = [got.pixels[i], got.pixels[i + 1], got.pixels[i + 2]]
      if (g > 40) back++
      else if (r > 40 && r >= b) to++
      else if (b > 40) from++
      else black++
    }
    const n = DEV_W * DEV_H
    return { to: to / n, back: back / n, from: from / n, black: black / n, diagnostics: got.diagnostics, shaderCalls: got.shaderCalls }
  },
  probe(name: string, attrs: Record<string, string>, p: number, points: Array<[number, number]>) {
    const got = shoot(specOf(name, attrs), p, ['blue', 'red'])
    return { colors: points.map(([x, y]) => at(got.pixels, ...ndcPx(x, y))), seed: got.seed, diagnostics: got.diagnostics }
  },
  differs(name: string, a: Record<string, string>, b: Record<string, string>, p: number, seeds: [number, number] = [21, 21]) {
    const x = shoot(specOf(name, a), p, ['from', 'to'], { rngSeed: seeds[0] })
    const y = shoot(specOf(name, b), p, ['from', 'to'], { rngSeed: seeds[1] })
    const d = frameDiff(x.pixels, y.pixels, DEV_W, DEV_H)
    let same = true
    for (let i = 0; i < x.pixels.length; i++) if (x.pixels[i] !== y.pixels[i]) { same = false; break }
    return { meanAbs: d.meanAbs, identical: same }
  },
  canvas2d(name: string) {
    const ws = makeWorkspace(), sink = createDiagnosticSink()
    const c2: Compositor = createCanvas2DCompositor(factory(DEV_W, DEV_H), factory, ws, sink)
    const route = dispatchTransition(specOf(name), c2, sink, 2)
    c2.dispose()
    return { route, diagnostics: sink.list().map(d => ({ code: d.code, slideIndex: d.slideIndex, effect: d.effect, message: d.message })) }
  },
  strip(name: string, attrs: Record<string, string> = {}) {
    const host = document.getElementById('strip')!
    host.replaceChildren()
    const W = DEV_W / 2, H = DEV_H / 2
    for (const p of [0, 0.25, 0.5, 0.75, 1]) {
      const shot = p === 0 || p === 1 ? shoot(specOf(name, attrs), 0, ['from', 'to'], { flat: p === 0 ? 'from' : 'to' }) : shoot(specOf(name, attrs), p, ['from', 'to'])
      const c = factory(W, H) as HTMLCanvasElement
      const ctx = c.getContext('2d')!
      const full = factory(DEV_W, DEV_H) as HTMLCanvasElement
      full.getContext('2d')!.putImageData(new ImageData(new Uint8ClampedArray(shot.pixels), DEV_W, DEV_H), 0, 0)
      ctx.drawImage(full, 0, 0, W, H)
      host.appendChild(c)
    }
    return host.children.length
  },
  /** Informal: per-frame CPU+GPU time (gl.finish-synchronised) over a sweep, at `size` device px. Never asserted. */
  perf(name: string, frames = 120, w = 1920, h = 1080) {
    const ws = makeWorkspace(512 * 1024 * 1024), sink = createDiagnosticSink()
    const out = factory(w, h) as HTMLCanvasElement
    const comp = createWebGL2Compositor(out, ws, { sink })
    if (!comp) throw new Error('no webgl2 compositor')
    const gl = out.getContext('webgl2')!
    const spec = specOf(name)
    comp.setViewport({ matrix: [w / DEV_W, 0, 0, h / DEV_H, 0, 0], deviceWidth: w, deviceHeight: h })
    if (dispatchTransition(spec, comp, sink, 0) !== 'shader') throw new Error('not dispatched to the shader')
    const big = (v: Paint) => { const c = factory(w, h) as HTMLCanvasElement; paint(c.getContext('2d')!, w, h, v); return { canvas: c, bounds: { x: 0, y: 0, w: DEV_W, h: DEV_H }, scale: w / DEV_W, bytes: w * h * 4 } }
    const [from, to] = (['from', 'to'] as const).map(v => { const b = big(v); const r = comp.upload(b, ws.lease(b.bytes, 'layer')!); if ('refused' in r) throw new Error(r.refused); return r })
    const rng = createRng(3)
    const px = new Uint8Array(4)
    const times: number[] = []
    for (let i = 0; i < frames; i++) {
      const p = (i + 0.5) / frames
      const t0 = performance.now()
      comp.frame([], { from, to, effect: spec.element, progress: p, params: transition3DDraw(spec, p, rng, DEV_W / DEV_H), rng })
      gl.readPixels(0, 0, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, px)
      times.push(performance.now() - t0)
    }
    comp.release(from); comp.release(to); comp.dispose()
    const sorted = [...times].sort((a, b) => a - b)
    const renderer = (() => { const dbg = gl.getExtension('WEBGL_debug_renderer_info'); return String(dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER)) })()
    return { name, frames, p50: sorted[Math.floor(frames * 0.5)], p95: sorted[Math.min(frames - 1, Math.floor(frames * 0.95))], max: sorted[frames - 1], renderer, diagnostics: sink.list().map(d => d.code) }
  },
}

declare global {
  interface Window { __slideshowDeform?: typeof api }
}
window.__slideshowDeform = api
export type SlideshowDeform = typeof api
