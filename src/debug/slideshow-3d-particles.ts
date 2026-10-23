/**
 * Slideshow plan S3-3 harness (dev only; src/debug is outside the library
 * build). Exposes `window.__slideshowParticles` for
 * tests/browser/slideshow-3d-particles.spec.ts:
 *
 * - `render`: one frame of a tile or particle row at a progress and effect
 *   seed on the WebGL2 or the Canvas 2D compositor (a real browser canvas),
 *   read back as RGBA, with the dispatch route and the diagnostics;
 * - `probe`: the row's own vertex shader run under transform feedback on
 *   chosen vertices, returning clip position, world position, local time,
 *   opacity and shade, so the spec can hold the GLSL to the algorithm sheet's
 *   reference models cell by cell;
 * - `bench`: frame times of a row at a device size, each frame forced to
 *   completion (`readPixels`), for the p95 budget on a real GPU.
 */
import { createCanvas2DCompositor } from '../slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../slideshow/render/webgl2'
import { createDiagnosticSink } from '../slideshow/diagnostics'
import { createRng } from '../slideshow/rng'
import { dispatchTransition, transition3DDraw } from '../slideshow/render/three/framework'
import { effectSeed } from '../slideshow/render/three/mesh'
import { PARTICLE_SHADERS, particlePlan, registerParticle3D, setParticleUniforms, type ShaderKey } from '../slideshow/render/three/particles'
import type { Affine, CanvasFactory, Compositor, LayerHandle, Lease, SlideShowWorkspace, TransitionSpec } from '../slideshow/types'

// A 192 × 108 slide at 1 device px per slide px, on a 208 × 124 output letterboxed by 8 px.
export const SLIDE_W = 192, SLIDE_H = 108
export const DEV_W = 208, DEV_H = 124
const VIEW: Affine = [1, 0, 0, 1, 8, 8]

interface Pool { used: number; limit: number }
function makeWorkspace(limit: number): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  return {
    pool, limit, used: () => pool.used, reserveTerminal: () => undefined,
    lease(bytes, tag): Lease | undefined {
      if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
      pool.used += bytes
      let held = bytes, live = true
      const lease: Lease = {
        get tag() { return tag }, get bytes() { return live ? held : 0 },
        grow(extra) { if (!live || pool.used + extra > pool.limit) return false; pool.used += extra; held += extra; return true },
        shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
        transfer() { return lease },
        release() { if (!live) return; live = false; pool.used -= held },
      }
      return lease
    },
  }
}

const factory: CanvasFactory = (w, h) => {
  const c = document.createElement('canvas')
  c.width = w; c.height = h
  return c
}

/** Every row of the group by the key the spec uses. */
const SPECS: Record<string, Pick<TransitionSpec, 'ns' | 'element' | 'preset'>> = {
  vortex: { ns: 'p14', element: 'vortex' },
  shred: { ns: 'p14', element: 'shred' },
  fracture: { ns: 'p15', element: 'prstTrans', preset: 'fracture' },
  crush: { ns: 'p15', element: 'prstTrans', preset: 'crush' },
  glitter: { ns: 'p14', element: 'glitter' },
  honeycomb: { ns: 'p14', element: 'honeycomb' },
  ripple: { ns: 'p14', element: 'ripple' },
}
const specOf = (key: string, attrs: Record<string, string>): TransitionSpec => ({ ...SPECS[key], attrs, durationMs: 1000, advClick: true })

type Variant = 'from' | 'to'
/** A smooth 2-D ramp, so every position and orientation has its own colour: from is (x, y, 40), to is (1 - x, 1 - y, 220). */
function paint(w: number, h: number, v: Variant): HTMLCanvasElement {
  const canvas = factory(w, h) as HTMLCanvasElement
  const ctx = canvas.getContext('2d')!
  const img = ctx.createImageData(w, h)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const fx = (x + 0.5) / w, fy = (y + 0.5) / h
      const i = (y * w + x) * 4
      img.data[i] = Math.round(255 * (v === 'from' ? fx : 1 - fx))
      img.data[i + 1] = Math.round(255 * (v === 'from' ? fy : 1 - fy))
      img.data[i + 2] = v === 'from' ? 40 : 220
      img.data[i + 3] = 255
    }
  }
  ctx.putImageData(img, 0, 0)
  return canvas
}

interface Session {
  comp: Compositor
  out: HTMLCanvasElement
  ws: ReturnType<typeof makeWorkspace>
  sink: ReturnType<typeof createDiagnosticSink>
  from: LayerHandle
  to: LayerHandle
  devW: number
  devH: number
  slideW: number
  slideH: number
}

function open(backend: 'webgl2' | 'canvas2d', devW = DEV_W, devH = DEV_H, slideW = SLIDE_W, slideH = SLIDE_H, view: Affine = VIEW): Session {
  const ws = makeWorkspace(1024 * 1024 * 1024)
  const sink = createDiagnosticSink()
  const out = factory(devW, devH) as HTMLCanvasElement
  const comp = backend === 'webgl2' ? createWebGL2Compositor(out, ws, { sink }) : createCanvas2DCompositor(out, factory, ws, sink)
  if (!comp) throw new Error(`no ${backend} compositor`)
  comp.setViewport({ matrix: view, deviceWidth: devW, deviceHeight: devH })
  const up = (v: Variant): LayerHandle => {
    const canvas = paint(slideW, slideH, v)
    const lease = ws.lease(slideW * slideH * 4, 'layer')!
    const r = comp.upload({ canvas, bounds: { x: 0, y: 0, w: slideW, h: slideH }, scale: 1, bytes: slideW * slideH * 4 }, lease)
    if ('refused' in r) throw new Error(`upload refused: ${r.refused}`)
    return r
  }
  return { comp, out, ws, sink, from: up('from'), to: up('to'), devW, devH, slideW, slideH }
}

function close(s: Session) {
  s.comp.dispose()
  if (s.ws.used() !== 0) throw new Error(`leaked ${s.ws.used()} bytes`)
}

function readBack(s: Session): Uint8ClampedArray {
  const r = factory(s.devW, s.devH) as HTMLCanvasElement
  const ctx = r.getContext('2d', { willReadFrequently: true })!
  ctx.fillStyle = '#202020'
  ctx.fillRect(0, 0, s.devW, s.devH)
  // Same task as the frame: the WebGL drawing buffer is still there without preserveDrawingBuffer.
  ctx.drawImage(s.out, 0, 0)
  return ctx.getImageData(0, 0, s.devW, s.devH).data
}

export interface RenderOptions { key: string; p: number; attrs?: Record<string, string>; backend?: 'webgl2' | 'canvas2d'; rngSeed?: number }
export interface RenderResult { pixels: number[]; route: string; /** The effect seed the shader derived from the effect rng. */ seed: number; diagnostics: Array<{ code: string; effect?: string; message: string }>; width: number; height: number }

function frameOn(s: Session, o: RenderOptions): { route: string; seed: number } {
  const spec = specOf(o.key, o.attrs ?? {})
  const route = dispatchTransition(spec, s.comp, s.sink, 0)
  const aspect = s.slideW / s.slideH
  const rng = createRng(o.rngSeed ?? 21)
  const params = transition3DDraw(spec, o.p, rng, aspect)
  // The compositor seeds the shader's rng from the effect rng's next value (webgl2.ts `seedOf`), then the row reads its first value.
  const twin = createRng(o.rngSeed ?? 21)
  transition3DDraw(spec, o.p, twin, aspect)
  const seed = effectSeed(createRng(Math.floor(twin() * 0x100000000)))
  s.comp.frame([], { from: s.from, to: s.to, effect: spec.element, progress: o.p, params, rng })
  return { route, seed }
}

function render(o: RenderOptions): RenderResult {
  const s = open(o.backend ?? 'webgl2')
  try {
    const { route, seed } = frameOn(s, o)
    return {
      pixels: Array.from(readBack(s)), route, seed, width: s.devW, height: s.devH,
      diagnostics: s.sink.list().map(d => ({ code: d.code, effect: d.effect, message: d.message })),
    }
  } finally { close(s) }
}

export interface ProbeVertex { uv: [number, number]; cell: [number, number] }
export interface ProbeOptions { key: ShaderKey; element: string; preset?: string; attrs?: Record<string, string>; p: number; aspect: number; seed: number; vertices: ProbeVertex[]; pass?: number }
export interface ProbeResult { clip: [number, number, number, number]; world: [number, number, number]; t: number; opacity: number; shade: number }

/** Runs the row's vertex shader on `vertices` under transform feedback (no rasterisation). */
function probe(o: ProbeOptions): ProbeResult[] {
  const plan = particlePlan(o.element, o.preset, o.attrs ?? {}, o.p, o.aspect, o.seed)
  const canvas = document.createElement('canvas')
  const gl = canvas.getContext('webgl2')
  if (!gl) throw new Error('no webgl2')
  const sources = PARTICLE_SHADERS[o.key]
  const compile = (type: number, src: string) => {
    const sh = gl.createShader(type)!
    gl.shaderSource(sh, src)
    gl.compileShader(sh)
    if (!gl.getShaderParameter(sh, gl.COMPILE_STATUS)) throw new Error(`shader: ${gl.getShaderInfoLog(sh)}`)
    return sh
  }
  const prog = gl.createProgram()!
  gl.attachShader(prog, compile(gl.VERTEX_SHADER, sources.vertex))
  gl.attachShader(prog, compile(gl.FRAGMENT_SHADER, sources.fragment))
  gl.transformFeedbackVaryings(prog, ['v_clip', 'v_world', 'v_t', 'v_opacity', 'v_shade'], gl.INTERLEAVED_ATTRIBS)
  gl.linkProgram(prog)
  if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) throw new Error(`link: ${gl.getProgramInfoLog(prog)}`)
  const n = o.vertices.length
  const data = new Float32Array(n * 4)
  o.vertices.forEach((v, i) => data.set([v.uv[0], v.uv[1], v.cell[0], v.cell[1]], i * 4))
  const vao = gl.createVertexArray()!
  gl.bindVertexArray(vao)
  const vbo = gl.createBuffer()!
  gl.bindBuffer(gl.ARRAY_BUFFER, vbo)
  gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW)
  gl.enableVertexAttribArray(0)
  gl.vertexAttribPointer(0, 2, gl.FLOAT, false, 16, 0)
  gl.enableVertexAttribArray(1)
  gl.vertexAttribPointer(1, 2, gl.FLOAT, false, 16, 8)
  const FLOATS = 10
  const tf = gl.createBuffer()!
  gl.bindBuffer(gl.TRANSFORM_FEEDBACK_BUFFER, tf)
  gl.bufferData(gl.TRANSFORM_FEEDBACK_BUFFER, n * FLOATS * 4, gl.STATIC_READ)
  gl.bindBufferBase(gl.TRANSFORM_FEEDBACK_BUFFER, 0, tf)
  gl.useProgram(prog)
  setParticleUniforms(gl, { uniform: name => gl.getUniformLocation(prog, name) }, plan.passes[o.pass ?? 0].uniforms)
  gl.enable(gl.RASTERIZER_DISCARD)
  gl.beginTransformFeedback(gl.POINTS)
  gl.drawArrays(gl.POINTS, 0, n)
  gl.endTransformFeedback()
  gl.disable(gl.RASTERIZER_DISCARD)
  const out = new Float32Array(n * FLOATS)
  gl.getBufferSubData(gl.TRANSFORM_FEEDBACK_BUFFER, 0, out)
  const err = gl.getError()
  if (err !== 0) throw new Error(`GL error ${err}`)
  return Array.from({ length: n }, (_v, i) => {
    const b = i * FLOATS
    return {
      clip: [out[b], out[b + 1], out[b + 2], out[b + 3]], world: [out[b + 4], out[b + 5], out[b + 6]],
      t: out[b + 7], opacity: out[b + 8], shade: out[b + 9],
    }
  })
}

export interface BenchOptions { key: string; attrs?: Record<string, string>; devW: number; devH: number; frames: number }

/** Frame times (ms) of the row across progress 0..1, each forced to complete on the GPU. */
function bench(o: BenchOptions): number[] {
  const s = open('webgl2', o.devW, o.devH, o.devW, o.devH, [1, 0, 0, 1, 0, 0])
  try {
    const gl = s.out.getContext('webgl2')!
    const pixel = new Uint8Array(4)
    const spec = specOf(o.key, o.attrs ?? {})
    if (dispatchTransition(spec, s.comp, s.sink, 0) !== 'shader') throw new Error('not dispatched to the shader')
    const times: number[] = []
    for (let i = 0; i < o.frames + 5; i++) {
      const p = (i % o.frames + 0.5) / o.frames
      const rng = createRng(21)
      const params = transition3DDraw(spec, p, rng, o.devW / o.devH)
      const t0 = performance.now()
      s.comp.frame([], { from: s.from, to: s.to, effect: spec.element, progress: p, params, rng })
      gl.readPixels(0, 0, 1, 1, gl.RGBA, gl.UNSIGNED_BYTE, pixel)
      if (i >= 5) times.push(performance.now() - t0)
    }
    if (s.sink.list().some(d => d.code === 'transition-shader-fallback')) throw new Error('shader fell back')
    return times
  } finally { close(s) }
}

registerParticle3D()

declare global {
  interface Window { __slideshowParticles?: typeof api }
}
const api = {
  geometry: { slideW: SLIDE_W, slideH: SLIDE_H, devW: DEV_W, devH: DEV_H, originX: VIEW[4], originY: VIEW[5] },
  keys: () => Object.keys(SPECS),
  render, probe, bench,
  /** WebGL renderer string, to record what the numbers were measured on. */
  renderer(): string {
    const gl = document.createElement('canvas').getContext('webgl2')
    if (!gl) return 'none'
    const dbg = gl.getExtension('WEBGL_debug_renderer_info')
    return String(dbg ? gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) : gl.getParameter(gl.RENDERER))
  },
}
window.__slideshowParticles = api
