/**
 * Slideshow plan S3-2 harness (dev only; src/debug is outside the library
 * build): the geometric 3-D transitions on the real WebGL2 compositor, for
 * tests/browser/slideshow-3d-geometric.spec.ts through
 * `window.__slideshow3dGeometric`. Its own page, so the S3-1 harness
 * (slideshow-gpu.ts) and the other group tasks' harnesses stay untouched.
 *
 * - `ends(element, attrs)`: just inside 0 and 1 against the flat frames;
 * - `middle(element, attrs)`: 25/50/75 % against the look-alike, the
 *   letterbox, and pixel probes on the top piece against the model
 *   (`geometricPieces`, held to the sheet's oracles by the unit tests);
 * - `canvas2d(element)`: dispatch and parity of the Canvas 2D look-alike;
 * - `frame(element, attrs, p)`: one frame as a PNG data URL (goldens);
 * - `bench(...)`: informal frame timings (the spec runs it only on request).
 *
 * Frames are "Office-shaped synthetic". Nothing here reads wall time except
 * `bench`, which measures it.
 */
import { createCanvas2DCompositor } from '../slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../slideshow/render/webgl2'
import { transitionDraw } from '../slideshow/render/transitions2d'
import { rowForTransition } from '../slideshow/catalogue'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { createRng } from '../slideshow/rng'
import { DEFAULT_FOV_DEG, mul4, project, slideCamera, slidePoint } from '../slideshow/render/three/mesh'
import { transition3DFor } from '../slideshow/render/three/registry'
import { dispatchTransition, transition3DDraw } from '../slideshow/render/three/framework'
import { BACK_SHADE, GEOMETRIC_ELEMENTS, geometricPieces, registerGeometric3D } from '../slideshow/render/three/geometric'
import type { Affine, CanvasFactory, Compositor, LayerBitmap, LayerHandle, Lease, SlideShowWorkspace, Tolerance, TransitionDraw, TransitionSpec } from '../slideshow/types'
import { frameDiff, withinTolerance } from './slideshow-parity'

registerGeometric3D()

// ---------------------------------------------------------------------------
// Workspace, canvases, frames
// ---------------------------------------------------------------------------
function makeWorkspace(limit = 512 * 1024 * 1024): SlideShowWorkspace {
  let used = 0
  const lease = (bytes: number, tag: string): Lease | undefined => {
    if (bytes < 0 || used + bytes > limit) return undefined
    used += bytes
    let held = bytes, live = true
    const l: Lease = {
      get tag() { return tag },
      get bytes() { return live ? held : 0 },
      grow(extra) { if (!live || used + extra > limit) return false; used += extra; held += extra; return true },
      shrink(less) { const n = Math.min(less, held); held -= n; used -= n },
      transfer() { return l },
      release() { if (!live) return; live = false; used -= held },
    }
    return l
  }
  return { limit, lease, used: () => used, reserveTerminal: () => undefined }
}
const factory: CanvasFactory = (w, h) => {
  const c = document.createElement('canvas')
  c.width = w; c.height = h
  return c
}

/** A 160 × 90 slide at 1 device px per slide px, letterboxed by 4 px above and below. */
interface Geometry { slideW: number; slideH: number; scale: number; devW: number; devH: number; top: number }
const SMALL: Geometry = { slideW: 160, slideH: 90, scale: 1, devW: 160, devH: 98, top: 4 }
const EXACT: Tolerance = { meanAbs: 0.02, rowMax: 0.10 }

/** Asymmetric content (a gradient, a top-left marker, a lower-right disc), so a flip or mirror shows. */
function paint(ctx: CanvasRenderingContext2D, w: number, h: number, which: 'from' | 'to') {
  const g = ctx.createLinearGradient(0, 0, w, h)
  const [a, b] = which === 'from' ? ['#204080', '#80c0ff'] : ['#a02020', '#ffd060']
  g.addColorStop(0, a); g.addColorStop(1, b)
  ctx.fillStyle = g
  ctx.fillRect(0, 0, w, h)
  ctx.fillStyle = '#10e010'
  ctx.fillRect(w * 0.08, h * 0.1, w * 0.25, h * 0.2)
  ctx.fillStyle = '#ffffff'
  ctx.beginPath(); ctx.arc(w * 0.7, h * 0.72, Math.min(w, h) * 0.18, 0, Math.PI * 2); ctx.fill()
}

function bitmap(geo: Geometry, which: 'from' | 'to'): LayerBitmap {
  const w = Math.ceil(geo.slideW * geo.scale), h = Math.ceil(geo.slideH * geo.scale)
  const canvas = factory(w, h) as HTMLCanvasElement
  paint(canvas.getContext('2d')!, w, h, which)
  return { canvas, bounds: { x: 0, y: 0, w: geo.slideW, h: geo.slideH }, scale: geo.scale, bytes: w * h * 4 }
}

function readBack(out: HTMLCanvasElement, geo: Geometry): Uint8ClampedArray {
  const r = factory(geo.devW, geo.devH) as HTMLCanvasElement
  const ctx = r.getContext('2d', { willReadFrequently: true })!
  ctx.fillStyle = '#202020'
  ctx.fillRect(0, 0, geo.devW, geo.devH)
  ctx.drawImage(out, 0, 0)
  return ctx.getImageData(0, 0, geo.devW, geo.devH).data
}

const spec = (element: string, attrs: Record<string, string> = {}): TransitionSpec => ({ element, ns: 'p14', attrs, durationMs: 1000, advClick: true })

type Make = (rng: ReturnType<typeof createRng>, aspect: number) => Record<string, unknown>
const shader = (element: string, attrs: Record<string, string>, p: number): Make => (rng, aspect) => transition3DDraw(spec(element, attrs), p, rng, aspect)
const lookalike = (element: string, attrs: Record<string, string>, p: number): Make => (rng, aspect) => transitionDraw(spec(element, attrs), p, rng, aspect)
const flatFrom: Make = () => ({ toOpacity: 0 })
const flatTo: Make = () => ({})

interface Rendered { pixels: Uint8ClampedArray; diagnostics: string[]; shaderCalls: number; out: HTMLCanvasElement }

/** One transition frame on `backend`; `element` names the effect for the compositor. */
function render(backend: 'webgl2' | 'canvas2d', element: string, progress: number, make: Make, geo = SMALL): Rendered {
  const ws = makeWorkspace()
  const sink = createDiagnosticSink()
  const out = factory(geo.devW, geo.devH) as HTMLCanvasElement
  const comp: Compositor | undefined = backend === 'webgl2' ? createWebGL2Compositor(out, ws, { sink }) : createCanvas2DCompositor(out, factory, ws, sink)
  if (!comp) throw new Error(`no ${backend} compositor`)
  const t = transition3DFor(spec(element))
  let shaderCalls = 0
  const original = t?.draw
  if (t && original) t.draw = function (...args) { shaderCalls++; return original.apply(this, args) }
  try {
    const view: Affine = [geo.scale, 0, 0, geo.scale, 0, geo.top]
    comp.setViewport({ matrix: view, deviceWidth: geo.devW, deviceHeight: geo.devH })
    const up = (which: 'from' | 'to'): LayerHandle => {
      const b = bitmap(geo, which)
      const r = comp.upload(b, ws.lease(b.bytes, 'layer')!)
      if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
      return r
    }
    const from = up('from'), to = up('to')
    const rng = createRng(21)
    const tr: TransitionDraw = { from, to, effect: element, progress, params: make(rng, geo.slideW / geo.slideH), rng }
    comp.frame([], tr)
    return { pixels: readBack(out, geo), diagnostics: sink.list().map(d => d.code), shaderCalls, out }
  } finally {
    if (t && original) t.draw = original
    comp.dispose()
    if (ws.used() !== 0) throw new Error(`${backend} leaked ${ws.used()} bytes`)
  }
}

const px = (pixels: Uint8ClampedArray, geo: Geometry, x: number, y: number): number[] => {
  const i = (Math.round(y) * geo.devW + Math.round(x)) * 4
  return [...pixels.slice(i, i + 4)]
}

// ---------------------------------------------------------------------------
// The API
// ---------------------------------------------------------------------------
const api = {
  elements: () => [...GEOMETRIC_ELEMENTS],

  /** Just inside the ends against the flat frames (and exactly at them, where the shader is not called). */
  ends(element: string, attrs: Record<string, string> = {}) {
    return ([[0, 'from'], [1e-3, 'from'], [0.999, 'to'], [1, 'to']] as const).map(([p, which]) => {
      const got = render('webgl2', element, p, shader(element, attrs, p))
      const want = render('webgl2', element, 0, which === 'from' ? flatFrom : flatTo)
      const d = frameDiff(got.pixels, want.pixels, SMALL.devW, SMALL.devH)
      return { progress: p, which, meanAbs: d.meanAbs, rowMax: d.rowMax, ok: withinTolerance(d, EXACT), shaderCalls: got.shaderCalls, diagnostics: got.diagnostics }
    })
  },

  /**
   * Mid-way: the shader runs, the frame is not the look-alike, the letterbox is untouched, and
   * interior points of the top piece land where the model puts them, showing its frame there.
   */
  middle(element: string, attrs: Record<string, string> = {}) {
    const geo = SMALL
    const fromFlat = render('webgl2', element, 0, flatFrom).pixels, toFlat = render('webgl2', element, 0, flatTo).pixels
    return [0.25, 0.5, 0.75].map((p) => {
      const got = render('webgl2', element, p, shader(element, attrs, p))
      const look = render('webgl2', element, p, lookalike(element, attrs, p))
      const a = geo.slideW / geo.slideH
      const camera = slideCamera(DEFAULT_FOV_DEG, a)
      const pieces = geometricPieces(element, p, attrs, a)
      const toDevice = (n: readonly number[]) => [(n[0] + 1) / 2 * geo.devW, geo.top + (1 - n[1]) / 2 * geo.slideH * geo.scale]
      // The last piece drawn that faces the camera, or is two-sided, is on top wherever it lands.
      let probes = 0, matched = 0, worst = 0
      for (let k = pieces.length - 1; k >= 0; k--) {
        const pc = pieces[k]
        const mvp = mul4(camera, pc.model)
        const r = pc.uvRect ?? [0, 0, 1, 1]
        const ndc = (u: number, v: number) => project(mvp, slidePoint(r[0] + u * r[2], r[1] + v * r[3], a))
        const [tl, tr, bl] = [ndc(0, 0), ndc(1, 0), ndc(0, 1)]
        const area = (bl[0] - tl[0]) * (tr[1] - tl[1]) - (bl[1] - tl[1]) * (tr[0] - tl[0])
        // Edge-on (flip at exactly ½) covers no pixels to probe; a culled back face shows nothing.
        if (Math.abs(area) < 1e-3) continue
        const front = area > 0
        if (!front && pc.sides !== 'both') continue
        const shade = (0.7 + 0.3 * Math.min(1, Math.abs(pc.model[10]))) * (front ? 1 : BACK_SHADE)
        const flat = pc.face === 'from' ? fromFlat : toFlat
        for (let i = 1; i < 5; i++) {
          for (let j = 1; j < 5; j++) {
            const u = i / 5, v = j / 5
            const n = ndc(u, v)
            if (Math.abs(n[0]) > 0.97 || Math.abs(n[1]) > 0.95 || n[2] < -1 || n[2] > 1) continue
            const [x, y] = toDevice(n)
            const su = r[0] + u * r[2], sv = r[1] + v * r[3]
            const want = px(flat, geo, su * geo.slideW * geo.scale, geo.top + sv * geo.slideH * geo.scale).slice(0, 3).map(c => c * shade)
            const have = px(got.pixels, geo, x, y)
            const dist = Math.max(...want.map((c, ch) => Math.abs(c - have[ch])))
            probes++
            worst = Math.max(worst, dist)
            if (dist <= 48) matched++
          }
        }
        // A piece mostly swung off the slide leaves the next one's points to probe (no S3-2 row's pieces overlap there).
        if (probes >= 4) break
      }
      const d = frameDiff(got.pixels, look.pixels, geo.devW, geo.devH)
      return {
        progress: p, shaderCalls: got.shaderCalls, diagnostics: got.diagnostics, vsLookalike: d.meanAbs,
        letterbox: px(got.pixels, geo, geo.devW / 2, 1), centre: px(got.pixels, geo, geo.devW * 0.3, geo.top + geo.slideH * geo.scale * 0.4), probes, matched, worst,
        digest: Array.from(got.pixels.filter((_v, i) => i % 97 === 0)).join(','),
      }
    })
  },

  /** Canvas 2D: dispatch says look-alike with a diagnostic, and the shader params draw exactly WebGL2's look-alike. */
  canvas2d(element: string) {
    const sink = createDiagnosticSink()
    const c2 = createCanvas2DCompositor(factory(SMALL.devW, SMALL.devH), factory, makeWorkspace(), sink)
    const route = dispatchTransition(spec(element), c2, sink, 4)
    c2.dispose()
    const glSink = createDiagnosticSink()
    const gl = createWebGL2Compositor(factory(SMALL.devW, SMALL.devH) as HTMLCanvasElement, makeWorkspace(), { sink: glSink })
    const glRoute = gl ? dispatchTransition(spec(element), gl, glSink, 4) : 'none'
    gl?.dispose()
    const results = [0.25, 0.5, 0.75].map((p) => {
      const c = render('canvas2d', element, p, shader(element, {}, p))
      const w = render('webgl2', element, p, lookalike(element, {}, p))
      const d = frameDiff(c.pixels, w.pixels, SMALL.devW, SMALL.devH)
      return { progress: p, meanAbs: d.meanAbs, rowMax: d.rowMax, ok: withinTolerance(d, EXACT), shaderCalls: c.shaderCalls }
    })
    const lookalikeId = rowForTransition('p14', element)?.lookalike ?? ''
    return { route, glRoute, lookalike: lookalikeId.split(':').pop(), diagnostics: sink.list().map(d => ({ code: d.code, slideIndex: d.slideIndex, effect: d.effect, message: d.message })), glDiagnostics: glSink.list().length, results }
  },

  /** One WebGL2 frame as RGBA (for the goldens) and as a PNG data URL (to record one). */
  frame(element: string, p: number, attrs: Record<string, string> = {}) {
    const got = render('webgl2', element, p, shader(element, attrs, p))
    const c = factory(SMALL.devW, SMALL.devH) as HTMLCanvasElement
    c.getContext('2d')!.putImageData(new ImageData(new Uint8ClampedArray(got.pixels), SMALL.devW, SMALL.devH), 0, 0)
    return { width: SMALL.devW, height: SMALL.devH, png: c.toDataURL('image/png'), diagnostics: got.diagnostics, shaderCalls: got.shaderCalls }
  },

  /** Compares RGBA `pixels` with a PNG data URL (a golden) in the page. */
  async compare(element: string, p: number, golden: string) {
    const img = new Image()
    img.src = golden
    await img.decode()
    const c = factory(SMALL.devW, SMALL.devH) as HTMLCanvasElement
    const ctx = c.getContext('2d', { willReadFrequently: true })!
    ctx.drawImage(img, 0, 0)
    const want = ctx.getImageData(0, 0, SMALL.devW, SMALL.devH).data
    const got = render('webgl2', element, p, shader(element, {}, p)).pixels
    const d = frameDiff(got, want, SMALL.devW, SMALL.devH)
    return { size: [img.width, img.height], meanAbs: d.meanAbs, rowMax: d.rowMax }
  },

  /**
   * Informal frame timings: `frames` transition frames of `element` on one WebGL2 compositor at
   * `devW × devH`, each timed from `frame()` to a 1-pixel `readPixels` (which waits for the GPU).
   */
  bench(element: string, devW: number, devH: number, frames = 120) {
    const geo: Geometry = { slideW: 160, slideH: 90, scale: devW / 160, devW, devH, top: (devH - 90 * (devW / 160)) / 2 }
    const ws = makeWorkspace(1024 * 1024 * 1024)
    const out = factory(devW, devH) as HTMLCanvasElement
    document.body.append(out)
    const comp = createWebGL2Compositor(out, ws)
    if (!comp) throw new Error('no webgl2 compositor')
    const gl = out.getContext('webgl2')!
    const dbg = gl.getExtension('WEBGL_debug_renderer_info')
    const renderer = String(dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER))
    comp.setViewport({ matrix: [geo.scale, 0, 0, geo.scale, 0, geo.top], deviceWidth: devW, deviceHeight: devH })
    const up = (which: 'from' | 'to') => { const b = bitmap(geo, which); const r = comp.upload(b, ws.lease(b.bytes, 'layer')!); if ('refused' in r) throw new Error(r.refused); return r }
    const from = up('from'), to = up('to')
    const s = spec(element)
    const route = dispatchTransition(s, comp)
    const one = new Uint8Array(4)
    const times: number[] = []
    for (let i = 0; i < frames + 10; i++) {
      const p = ((i % frames) + 0.5) / frames
      const rng = createRng(3)
      const t0 = performance.now()
      comp.frame([], { from, to, effect: element, progress: p, params: transition3DDraw(s, p, rng, 16 / 9), rng })
      gl.readPixels(0, 0, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, one)
      const dt = performance.now() - t0
      if (i >= 10) times.push(dt) // the first frames compile and warm up
    }
    comp.dispose()
    out.remove()
    const sorted = [...times].sort((x, y) => x - y)
    const q = (f: number) => sorted[Math.min(sorted.length - 1, Math.floor(f * sorted.length))]
    return { element, route, renderer, devW, devH, frames: times.length, p50: q(0.5), p95: q(0.95), max: sorted[sorted.length - 1] }
  },
}

declare global {
  interface Window { __slideshow3dGeometric?: typeof api }
}
window.__slideshow3dGeometric = api

export type Slideshow3DGeometric = typeof api
