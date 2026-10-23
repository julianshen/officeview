/**
 * S3-1: the 3-D transition framework (slideshow spec 6.3 "3-D transitions
 * run as GPU compositor shaders; without WebGL2 each uses its catalogued 2-D
 * look-alike, with a diagnostic").
 *
 * Node has no WebGL2, so the WebGL2 compositor runs on a recording fake GL
 * (the S1-14 pattern): these tests pin dispatch, leases, GL state and the
 * pure camera maths, never GPU pixels. Pixels (the sample shader at its ends
 * matching the flat frames, the Canvas 2D look-alike) are checked in Chromium
 * by tests/browser/slideshow-3d.spec.ts. Canvas 2D pixels come from
 * node-canvas.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { CATALOGUE } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createRng } from '../src/slideshow/rng'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../src/slideshow/render/webgl2'
import { reducedTransition, transitionDraw } from '../src/slideshow/render/transitions2d'
import {
  DEFAULT_FOV_DEG, GRID_MAX, Transition3DError, cameraDistance, cellRandom, createMeshPool, ease, effectSeed, gridBytes, hash32, mul4, perspective, project, slideCamera, slidePoint,
} from '../src/slideshow/render/three/mesh'
import { THREE_D_IDS, registerTransition3D, registeredTransitions3D, resetTransitions3D, transition3DFor } from '../src/slideshow/render/three/registry'
import { dispatchTransition, shader3DOf, transition3DDraw } from '../src/slideshow/render/three/framework'
import { createSampleTransition3D, sampleFace } from '../src/debug/slideshow-3d-sample'
import type { CanvasFactory, Compositor, LayerHandle, Lease, SlideShowWorkspace, TransitionSpec } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace ledger, factory and a recording fake WebGL2 context
// ---------------------------------------------------------------------------
interface Pool { used: number; limit: number }
function poolLease(pool: Pool, bytes: number, tag: string, log?: Array<{ tag: string; bytes: number }>): Lease | undefined {
  if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
  pool.used += bytes
  log?.push({ tag, bytes })
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
function makeWorkspace(limit = 256 * 1024 * 1024, log?: Array<{ tag: string; bytes: number }>): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  return { pool, limit, lease: (b, t) => poolLease(pool, b, t, log), used: () => pool.used, reserveTerminal: () => undefined }
}
const factory: CanvasFactory = (w, h) => createCanvas(w, h) as unknown as HTMLCanvasElement

interface GlCall { name: string; args: unknown[] }
function fakeGL() {
  const calls: GlCall[] = []
  const constants = new Map<string, number>()
  let serial = 0
  let lost = false
  const constant = (name: string) => {
    let v = constants.get(name)
    if (v === undefined) { v = 0x1000 + constants.size; constants.set(name, v) }
    return v
  }
  const overrides: Record<string, unknown> = {
    isContextLost: () => lost,
    getParameter: (p: number) => (p === constant('MAX_TEXTURE_SIZE') ? 8192 : 0),
    getShaderParameter: () => true,
    getProgramParameter: () => true,
    getShaderInfoLog: () => '',
    getProgramInfoLog: () => '',
    getUniformLocation: (_p: unknown, name: string) => ({ uniform: name }),
    getAttribLocation: () => 0,
    getExtension: () => null,
    getError: () => 0,
  }
  const gl = new Proxy({}, {
    get(_t, prop) {
      if (typeof prop !== 'string') return undefined
      if (prop in overrides) return overrides[prop]
      if (/^[A-Z][A-Z0-9_]*$/.test(prop)) return constant(prop)
      return (...args: unknown[]) => {
        calls.push({ name: prop, args })
        if (prop.startsWith('create')) return { kind: prop.slice(6), serial: ++serial }
        return undefined
      }
    },
  }) as WebGL2RenderingContext
  return {
    gl, calls, constant,
    count: (name: string) => calls.filter(c => c.name === name).length,
    lose() { lost = true },
    restore() { lost = false },
  }
}

let currentGl: ReturnType<typeof fakeGL> | null = null
let realGetContext: typeof HTMLCanvasElement.prototype.getContext
beforeEach(() => {
  resetTransitions3D()
  realGetContext = HTMLCanvasElement.prototype.getContext
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement, type: string, ...rest: unknown[]) {
    if (type === 'webgl2') return (currentGl?.gl ?? null) as never
    return (realGetContext as (...a: unknown[]) => unknown).call(this, type, ...rest) as never
  } as never)
})
afterEach(() => { vi.restoreAllMocks(); currentGl = null; resetTransitions3D() })

const W = 64, H = 36
function spec(element: string, attrs: Record<string, string> = {}, ns: TransitionSpec['ns'] = 'p14', preset?: string): TransitionSpec {
  return { element, ns, attrs, durationMs: 500, advClick: true, ...(preset !== undefined ? { preset } : {}) }
}

/** Two solid full-slide frames uploaded to `comp`. */
function frames(comp: Compositor, ws: SlideShowWorkspace, w = W, h = H): { from: LayerHandle; to: LayerHandle } {
  const layer = (color: string): LayerHandle => {
    const c = createCanvas(w, h)
    const x = c.getContext('2d'); x.fillStyle = color; x.fillRect(0, 0, w, h)
    const r = comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w, h }, scale: 1, bytes: w * h * 4 }, ws.lease(w * h * 4, 'layer')!)
    if ('refused' in r) throw new Error(r.refused)
    return r
  }
  return { from: layer('#ff0000'), to: layer('#0000ff') }
}

function webgl(ws = makeWorkspace(), w = W, h = H) {
  currentGl = fakeGL()
  const fake = currentGl
  const sink = createDiagnosticSink()
  const canvas = document.createElement('canvas')
  canvas.width = w; canvas.height = h
  const comp = createWebGL2Compositor(canvas, ws, { sink, factory })!
  comp.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: w, deviceHeight: h })
  return { comp, fake, sink, ws }
}

function canvasPixels(params: Record<string, unknown>, progress: number, element: string) {
  const ws = makeWorkspace()
  const out = createCanvas(W, H)
  const sink = createDiagnosticSink()
  const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws, sink)
  const { from, to } = frames(comp, ws)
  comp.frame([], { from, to, effect: element, progress, params, rng: createRng(5) })
  const data = [...out.getContext('2d').getImageData(0, 0, W, H).data]
  comp.dispose()
  return data
}

// ---------------------------------------------------------------------------
// The brief's tests
// ---------------------------------------------------------------------------
describe('3-D transition framework', () => {
  it('a registered 3-D row dispatches to its shader on webgl2 and to its lookalike on canvas2d with a diagnostic', () => {
    const { transition, calls } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const vortex = spec('vortex', { dir: 'r' })

    // WebGL2: the shader, no diagnostic.
    const g = webgl()
    expect(dispatchTransition(vortex, g.comp, g.sink, 2)).toBe('shader')
    expect(g.sink.list()).toEqual([])
    const { from, to } = frames(g.comp, g.ws)
    expect(g.fake.count('createTexture')).toBe(2)
    const rng = createRng(9)
    const params = transition3DDraw(vortex, 0.25, rng, W / H)
    expect(shader3DOf(params)).toEqual({ ns: 'p14', element: 'vortex', attrs: { dir: 'r' } })
    const drawSpy = vi.spyOn(transition, 'draw')
    g.comp.frame([], { from, to, effect: 'vortex', progress: 0.25, params, rng })
    expect(drawSpy).toHaveBeenCalledTimes(1)
    const [gl, fromTex, toTex, progress, attrs, shaderRng, mesh] = drawSpy.mock.calls[0]
    expect(gl).toBe(g.fake.gl)
    expect((fromTex as { serial: number }).serial).toBeLessThan((toTex as { serial: number }).serial)
    expect(progress).toBe(0.25)
    expect(attrs).toEqual({ dir: 'r' })
    expect(typeof shaderRng).toBe('function')
    expect(mesh.aspect).toBeCloseTo(W / H)
    expect(calls).toHaveLength(1)
    // The standard program drew the 1 × 1 grid: one indexed draw of 6 indices.
    expect(g.fake.calls.filter(c => c.name === 'drawElements').map(c => c.args[1])).toEqual([6])
    expect(g.sink.list()).toEqual([])
    g.comp.dispose()
    expect(g.ws.used()).toBe(0)

    // Canvas 2D: the catalogued look-alike (vortex → dissolve), diagnosed once.
    const ws = makeWorkspace()
    const sink = createDiagnosticSink()
    const c2 = createCanvas2DCompositor(createCanvas(W, H) as unknown as HTMLCanvasElement, factory, ws, sink)
    expect(dispatchTransition(vortex, c2, sink, 2)).toBe('lookalike')
    expect(dispatchTransition(vortex, c2, sink, 2)).toBe('lookalike')
    expect(sink.list()).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 2, effect: 'vortex' })])
    expect(sink.list()[0].message).toMatch(/dissolve/)
    c2.dispose()
    // The same params draw exactly the look-alike on Canvas 2D (the shader reference is ignored there).
    const lookalike = transitionDraw(spec('dissolve', {}, 'p'), 0.5, createRng(1), W / H)
    expect(canvasPixels(transition3DDraw(vortex, 0.5, createRng(1), W / H), 0.5, 'vortex')).toEqual(canvasPixels(lookalike, 0.5, 'dissolve'))
    expect(calls).toHaveLength(1)
  })

  it('mesh pool reuses grids and charges the workspace', () => {
    const fake = fakeGL()
    const log: Array<{ tag: string; bytes: number }> = []
    const ws = makeWorkspace(256 * 1024 * 1024, log)
    const owner = createMeshPool(fake.gl, ws)
    const a = owner.pool.grid(4, 3)
    expect(a.count).toBe(4 * 3 * 6)
    expect(ws.used()).toBe(gridBytes(4, 3))
    expect(log).toEqual([{ tag: 'webgl2:mesh', bytes: gridBytes(4, 3) }])
    // Leased before any GL object exists for it.
    const vaos = fake.count('createVertexArray'), buffers = fake.count('createBuffer')
    expect(vaos).toBe(1)
    expect(buffers).toBe(2)
    // Reused: the same handle, nothing created or charged again.
    expect(owner.pool.grid(4, 3)).toBe(a)
    expect(fake.count('createVertexArray')).toBe(1)
    expect(ws.used()).toBe(gridBytes(4, 3))
    // A different grid is its own charge.
    const b = owner.pool.grid(2, 2)
    expect(b).not.toBe(a)
    expect(ws.used()).toBe(gridBytes(4, 3) + gridBytes(2, 2))
    // A refused lease creates nothing and holds nothing.
    const tight = makeWorkspace(gridBytes(1, 1) - 1)
    const before = fake.count('createVertexArray')
    const refused = createMeshPool(fake.gl, tight)
    expect(() => refused.pool.grid(1, 1)).toThrow(Transition3DError)
    expect(fake.count('createVertexArray')).toBe(before)
    expect(tight.used()).toBe(0)
    // Out-of-range grids are rejected before leasing.
    expect(() => owner.pool.grid(0, 1)).toThrow(RangeError)
    expect(() => owner.pool.grid(GRID_MAX + 1, 1)).toThrow(RangeError)
    expect(() => owner.pool.grid(1.5, 2)).toThrow(RangeError)
    // Dispose deletes and refunds everything.
    owner.dispose(true)
    expect(ws.used()).toBe(0)
    expect(fake.count('deleteVertexArray')).toBe(2)
    expect(fake.count('deleteBuffer')).toBe(4)
  })

  it('progress 0 and 1 reproduce the outgoing and incoming frames for a sample shader', () => {
    const aspect = W / H
    const corners = [[0, 0], [1, 0], [0, 1], [1, 1]] as const
    const ndc = (u: number, v: number) => [2 * u - 1, 1 - 2 * v]
    // The sample's own model: at 0 the outgoing frame and at 1 the incoming frame land exactly on the slide rect.
    for (const [p, face] of [[0, 'from'], [1, 'to']] as const) {
      const s = sampleFace(p, aspect)
      expect(s.face).toBe(face)
      for (const [u, v] of corners) {
        const [x, y] = project(s.mvp, slidePoint(u, v, aspect))
        expect(x).toBeCloseTo(ndc(u, v)[0], 6)
        expect(y).toBeCloseTo(ndc(u, v)[1], 6)
      }
    }
    // Near the ends it is still that frame, within a hair of the rect (the browser spec compares pixels).
    const near = project(sampleFace(1e-4, aspect).mvp, slidePoint(1, 0, aspect))
    expect(near[0]).toBeCloseTo(1, 3)

    // The compositor at exactly 0 and 1 draws the flat outgoing / incoming frame and never calls the shader.
    const { transition, calls } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const g = webgl()
    g.fake.calls.length = 0
    const { from, to } = frames(g.comp, g.ws)
    // The two layer textures, in upload order.
    const [fromTex, toTex] = g.fake.calls.filter(c => c.name === 'bindTexture').map(c => c.args[1])
    for (const [p, expected] of [[0, fromTex], [1, toTex]] as const) {
      g.fake.calls.length = 0
      const rng = createRng(3)
      g.comp.frame([], { from, to, effect: 'vortex', progress: p, params: transition3DDraw(spec('vortex'), p, rng, aspect), rng })
      expect(calls).toHaveLength(0)
      // Flat quads only, no indexed (mesh) draw; the topmost quad is the expected frame, unmasked and opaque.
      expect(g.fake.count('drawElements')).toBe(0)
      let unit = -1
      const onUnit0: unknown[] = []
      let drawn: unknown
      for (const c of g.fake.calls) {
        if (c.name === 'activeTexture') unit = c.args[0] as number
        if (c.name === 'bindTexture' && unit === g.fake.constant('TEXTURE0')) onUnit0.push(c.args[1])
        if (c.name === 'drawArrays') drawn = onUnit0[onUnit0.length - 1]
      }
      expect(drawn).toBe(expected)
      const last = <T,>(name: string) => g.fake.calls.filter(c => c.name === name).at(-1)!.args as T
      expect(last<[unknown, number]>('uniform1f')[1]).toBe(1)
    }
    g.comp.dispose()
  })
})

// ---------------------------------------------------------------------------
// Contracts beyond the brief's names
// ---------------------------------------------------------------------------
describe('3-D transition framework contracts', () => {
  it('the registry holds exactly the catalogued 3-D rows, each with a look-alike', () => {
    const rows = CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 3 && r.ns !== 'p159')
    expect(THREE_D_IDS).toEqual(rows.map(r => r.id))
    expect(THREE_D_IDS).toHaveLength(26)
    for (const r of rows) expect(r.lookalike, r.id).toBeTruthy()
    // Non-3-D rows and unknown presets are refused; a duplicate is refused; re-registering the same object is idempotent.
    expect(() => registerTransition3D({ element: 'fade', draw() {} })).toThrow()
    expect(() => registerTransition3D({ element: 'morph', draw() {} })).toThrow()
    expect(() => registerTransition3D({ element: 'prstTrans', preset: 'nope', draw() {} })).toThrow()
    const t = { element: 'prstTrans', preset: 'curtains', draw() {} }
    registerTransition3D(t)
    registerTransition3D(t)
    expect(() => registerTransition3D({ element: 'prstTrans', preset: 'curtains', draw() {} })).toThrow()
    expect(registeredTransitions3D()).toEqual([t])
    expect(transition3DFor(spec('prstTrans', {}, 'p15', 'curtains'))).toBe(t)
    expect(transition3DFor(spec('prstTrans', {}, 'p15', 'airplane'))).toBeUndefined()
  })

  it('an unregistered 3-D row draws its look-alike even on webgl2, with a diagnostic', () => {
    const g = webgl()
    expect(dispatchTransition(spec('vortex'), g.comp, g.sink, 0)).toBe('lookalike')
    expect(g.sink.list().map(d => d.code)).toEqual(['transition-approximate'])
    // A native 2-D row is not approximate: no diagnostic.
    const sink = createDiagnosticSink()
    expect(dispatchTransition(spec('wipe', {}, 'p'), g.comp, sink, 0)).toBe('lookalike')
    expect(sink.list()).toEqual([])
    g.comp.dispose()
  })

  it('reduced motion never reaches a shader', () => {
    const { transition } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const g = webgl()
    const reduced = reducedTransition(spec('vortex'))
    expect(reduced.element).toBe('fade')
    expect(reduced.durationMs).toBeLessThanOrEqual(200)
    expect(dispatchTransition(reduced, g.comp, g.sink, 0)).toBe('lookalike')
    expect(g.sink.list()).toEqual([])
    g.comp.dispose()
  })

  it('the shader gets the same seeded rng on every frame, derived from the effect rng', () => {
    const { transition, calls } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const g = webgl()
    const { from, to } = frames(g.comp, g.ws)
    const rng = createRng(11)
    for (const p of [0.2, 0.4, 0.6]) g.comp.frame([], { from, to, effect: 'vortex', progress: p, params: transition3DDraw(spec('vortex'), p, rng, W / H), rng })
    expect(calls.map(c => c.progress)).toEqual([0.2, 0.4, 0.6])
    expect(new Set(calls.map(c => c.first)).size).toBe(1)
    // Another effect rng with the same seed gives the same draw; a different seed a different one.
    const again = webgl()
    const f2 = frames(again.comp, again.ws)
    const r2 = createRng(11)
    again.comp.frame([], { ...f2, effect: 'vortex', progress: 0.5, params: transition3DDraw(spec('vortex'), 0.5, r2, W / H), rng: r2 })
    const r3 = createRng(12)
    again.comp.frame([], { ...f2, effect: 'vortex', progress: 0.5, params: transition3DDraw(spec('vortex'), 0.5, r3, W / H), rng: r3 })
    expect(calls[3].first).toBe(calls[0].first)
    expect(calls[4].first).not.toBe(calls[0].first)
    g.comp.dispose(); again.comp.dispose()
  })

  it('the 3-D draw is viewport-mapped to the slide rect and the compositor state is restored after it', () => {
    const { transition } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const g = webgl(makeWorkspace(), 100, 80)
    // A letterboxed slide: 64 × 36 slide px at 1.5 device px each, offset (2, 13).
    g.comp.setViewport({ matrix: [1.5, 0, 0, 1.5, 2, 13], deviceWidth: 100, deviceHeight: 80 })
    const { from, to } = frames(g.comp, g.ws)
    g.fake.calls.length = 0
    const rng = createRng(1)
    g.comp.frame([], { from, to, effect: 'vortex', progress: 0.3, params: transition3DDraw(spec('vortex'), 0.3, rng, W / H), rng })
    const viewports = g.fake.calls.filter(c => c.name === 'viewport').map(c => c.args)
    // Output, then the slide's device rect (GL origin bottom-left: 80 - (13 + 54) = 13), then the output again.
    expect(viewports).toEqual([[0, 0, 100, 80], [2, 13, 96, 54], [0, 0, 100, 80]])
    const names = g.fake.calls.map(c => c.name)
    const drawAt = names.indexOf('drawElements')
    const after = g.fake.calls.slice(drawAt)
    expect(after.some(c => c.name === 'useProgram')).toBe(true)
    expect(after.some(c => c.name === 'bindVertexArray' && c.args[0] !== null)).toBe(true)
    expect(after.some(c => c.name === 'disable' && c.args[0] === g.fake.constant('CULL_FACE'))).toBe(true)
    expect(after.some(c => c.name === 'disable' && c.args[0] === g.fake.constant('SCISSOR_TEST'))).toBe(true)
    expect(after.some(c => c.name === 'blendEquation' && c.args[0] === g.fake.constant('FUNC_ADD'))).toBe(true)
    // The scissor is re-issued on the slide rect (GL origin bottom-left), whatever the shader set.
    expect(after.filter(c => c.name === 'scissor').map(c => c.args)).toEqual([[2, 13, 96, 54]])
    // An opaque black clear under the 3-D scene, inside the scissor.
    const clears = g.fake.calls.filter(c => c.name === 'clearColor').map(c => c.args)
    expect(clears).toContainEqual([0, 0, 0, 1])
    g.comp.dispose()
  })

  it('a refused mesh at frame time falls back to the look-alike with a diagnostic and holds nothing', () => {
    // No prepare: its largest grid is leased inside the frame.
    const calls: number[] = []
    registerTransition3D({ element: 'vortex', draw(_gl, _f, _t, p, _a, _r, mesh) { calls.push(p); mesh.grid(GRID_MAX, GRID_MAX) } })
    const ws = makeWorkspace()
    const g = webgl(ws)
    const { from, to } = frames(g.comp, g.ws)
    // Fill the workspace so the mesh lease is refused (the look-alike's small mask still fits).
    const filler = ws.lease(ws.limit - ws.used() - 512 * 1024, 'filler')!
    g.fake.calls.length = 0
    const rng = createRng(2)
    g.comp.frame([], { from, to, effect: 'vortex', progress: 0.5, params: transition3DDraw(spec('vortex'), 0.5, rng, W / H), rng })
    expect(calls).toHaveLength(1)
    expect(g.fake.count('createVertexArray')).toBe(0)
    expect(g.sink.list().map(d => `${d.code}:${d.effect}`)).toEqual(['transition-shader-fallback:vortex'])
    // The look-alike drew instead (flat quads).
    expect(g.fake.count('drawArrays')).toBeGreaterThan(0)
    filler.release()
    g.comp.dispose()
    expect(ws.used()).toBe(0)
  })

  it('a prepare failure dispatches to the look-alike', () => {
    registerTransition3D({ element: 'vortex', prepare() { throw new Transition3DError('compile', 'no') }, draw() {} })
    const g = webgl()
    expect(dispatchTransition(spec('vortex'), g.comp, g.sink, 1)).toBe('lookalike')
    expect(g.sink.list().map(d => d.code)).toEqual(['transition-approximate'])
    g.comp.dispose()
  })

  it('context loss drops the mesh pool and a restored context rebuilds it on demand', () => {
    const { transition } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const ws = makeWorkspace()
    const g = webgl(ws)
    const base = ws.used()
    expect(dispatchTransition(spec('vortex'), g.comp, g.sink, 0)).toBe('shader')
    expect(ws.used()).toBe(base + gridBytes(1, 1))
    g.fake.lose()
    g.comp.contextLost()
    expect(ws.used()).toBe(0)
    expect(dispatchTransition(spec('vortex'), g.comp, g.sink, 0)).toBe('lookalike')
    g.fake.restore()
    expect(g.comp.contextRestored()).toBe(true)
    const programs = g.fake.count('createProgram')
    const { from, to } = frames(g.comp, ws)
    const rng = createRng(4)
    g.comp.frame([], { from, to, effect: 'vortex', progress: 0.5, params: transition3DDraw(spec('vortex'), 0.5, rng, W / H), rng })
    // The standard 3-D program and the grid were recreated after the restore.
    expect(g.fake.count('createProgram')).toBe(programs + 1)
    expect(g.fake.count('drawElements')).toBe(1)
    g.comp.dispose()
    expect(ws.used()).toBe(0)
  })

  it('at 4K the 3-D path leases only its small meshes, never a surface', () => {
    const { transition } = createSampleTransition3D('vortex')
    registerTransition3D(transition)
    const log: Array<{ tag: string; bytes: number }> = []
    const ws = makeWorkspace(512 * 1024 * 1024, log)
    const g = webgl(ws, 3840, 2160)
    const { from, to } = frames(g.comp, ws, 64, 36)
    g.comp.setViewport({ matrix: [60, 0, 0, 60, 0, 0], deviceWidth: 3840, deviceHeight: 2160 })
    log.length = 0
    const rng = createRng(6)
    g.comp.frame([], { from, to, effect: 'vortex', progress: 0.5, params: transition3DDraw(spec('vortex'), 0.5, rng, 64 / 36), rng })
    expect(log.map(l => l.tag)).toEqual(['webgl2:mesh'])
    expect(log[0].bytes).toBe(gridBytes(1, 1))
    expect(gridBytes(GRID_MAX, GRID_MAX)).toBeLessThan(2 * 1024 * 1024)
    g.comp.dispose()
  })

  it('the slide camera fills the viewport with the z = 0 slide plane', () => {
    for (const aspect of [16 / 9, 4 / 3, 1]) {
      const m = slideCamera(DEFAULT_FOV_DEG, aspect)
      const at = (u: number, v: number) => project(m, slidePoint(u, v, aspect))
      expect(at(0, 0)[0]).toBeCloseTo(-1, 6); expect(at(0, 0)[1]).toBeCloseTo(1, 6)
      expect(at(1, 1)[0]).toBeCloseTo(1, 6); expect(at(1, 1)[1]).toBeCloseTo(-1, 6)
      expect(at(0.5, 0.5)[0]).toBeCloseTo(0, 6); expect(at(0.5, 0.5)[1]).toBeCloseTo(0, 6)
    }
    expect(cameraDistance(90)).toBeCloseTo(0.5)
    // perspective(fov, aspect) alone: a point at z = -d on the frustum edge lands on NDC ±1.
    const d = cameraDistance(DEFAULT_FOV_DEG)
    const p = perspective(DEFAULT_FOV_DEG, 2)
    expect(project(p, [1, 0.5, -d])[0]).toBeCloseTo(1)
    expect(project(mul4(p, p), [0, 0, -1]).every(Number.isFinite)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// The algorithm sheet (docs/validation/slideshow-3d-algorithms.md) and its seeded helpers
// ---------------------------------------------------------------------------
describe('3-D algorithm sheet', () => {
  it('has one entry and one generated oracle per 3-D row, matching the reference script', async () => {
    const { readFileSync } = await import('node:fs')
    const { ROWS } = await import('../scripts/slideshow-3d-oracles')
    const doc = readFileSync('docs/validation/slideshow-3d-algorithms.md', 'utf8')
    const headings = [...doc.matchAll(/^### `(transition:[^`]+)`/gm)].map(m => m[1])
    expect([...headings].sort()).toEqual([...THREE_D_IDS].sort())
    expect(ROWS.map(r => r.id).sort()).toEqual([...THREE_D_IDS].sort())
    // Labels exactly; numbers within 2e-4 (libm last-ulp differences between hosts move a 4th decimal or a boundary sample).
    const nums = (t: string) => (t.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number)
    const lines = doc.split('\n')
    for (const r of ROWS) {
      const want = `| ${r.id} | ${r.task} | ${r.attrs} | ${Object.entries(r.oracle(0.5)).map(([k, v]) => `${k} = ${v}`).join('; ')} |`
      const got = lines.find(l => l.startsWith(`| ${r.id} | ${r.task} |`))
      expect(got, r.id).toBeDefined()
      expect(got!.replace(/-?\d+(?:\.\d+)?/g, '#'), r.id).toBe(want.replace(/-?\d+(?:\.\d+)?/g, '#'))
      const a = nums(got!), b = nums(want)
      expect(a.length, r.id).toBe(b.length)
      a.forEach((v, i) => expect(Math.abs(v - b[i]), `${r.id} #${i}`).toBeLessThanOrEqual(2e-4))
    }
  })

  it('every reference model is continuous into its ends', async () => {
    const o = await import('../scripts/slideshow-3d-oracles')
    const near = (m: Float32Array, eps = 2e-3) => {
      const id = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]
      return [...m].every((v, i) => Math.abs(v - id[i]) < eps)
    }
    for (const f of [o.prism, o.flip, o.gallery, o.switchRow, o.conveyor, o.ferris]) {
      expect(near(f(1e-4).from), f.name).toBe(true)
      expect(near(f(1 - 1e-4).to), f.name).toBe(true)
    }
    expect(near(o.prism(1e-4, 'u', true, true).from)).toBe(true)
    expect(near(o.doors(1e-4).first) && near(o.doors(1 - 1e-4).to)).toBe(true)
    expect(near(o.fallOver(1e-4).from) && near(o.prestige(1 - 1e-4).to)).toBe(true)
    expect(o.vortexTile(0, 3, 4).t).toBe(0)
    expect(o.vortexTile(1, 3, 4).opacity).toBe(0)
    expect(o.fold(0, 1, 0.5, o.curlSingle()[0])).toBe('from')
    expect(o.fold(1, 1, 0.5, o.curlSingle()[0])).toBe('to')
    expect(o.warpK(0, 0.7)).toBe(1)
    expect(o.curtainX(1, 0.5, 0)).toBeLessThan(-o.ASPECT / 2)
    expect(o.drapeY(1, 0)).toBeLessThan(-0.5)
  })

  it('cellRandom is a deterministic integer hash in [0, 1)', () => {
    expect(hash32(0)).toBe(0)
    expect(hash32(1)).toBe(hash32(1))
    const seen = new Set<number>()
    for (let i = 0; i < 64; i++) for (let j = 0; j < 4; j++) {
      const v = cellRandom(2 ** 31, i, j, 0)
      expect(v).toBeGreaterThanOrEqual(0)
      expect(v).toBeLessThan(1)
      seen.add(v)
    }
    expect(seen.size).toBe(256)
    expect(cellRandom(1, 2, 3, 4)).toBe(cellRandom(1, 2, 3, 4))
    expect(cellRandom(1, 2, 3, 4)).not.toBe(cellRandom(2, 2, 3, 4))
    expect(effectSeed(() => 0.5)).toBe(2 ** 31)
    expect(ease(0.5)).toBe(0.5)
  })
})
