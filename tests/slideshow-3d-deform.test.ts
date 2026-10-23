/**
 * S3-4: the deformation 3-D transitions (`src/slideshow/render/three/deform.ts`):
 * pageCurlSingle, pageCurlDouble, peelOff, airplane, origami, curtains, drape,
 * wind, fallOver, prestige and warp (docs/validation/slideshow-3d-algorithms.md).
 *
 * Node has no WebGL2, so the shaders run on the S3-1 recording fake GL: these
 * tests pin registration, dispatch, leases, the uniforms and matrices each row
 * sends, and its pure model against the algorithm sheet's 50 % oracles (the
 * reference script `scripts/slideshow-3d-oracles.ts` wins over the prose).
 * GPU pixels (the GLSL against the same models) are checked in Chromium by
 * tests/browser/slideshow-3d-deform.spec.ts.
 */
import { existsSync, readFileSync } from 'node:fs'
import { resolve as resolvePath } from 'node:path'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { readTransition } from '../src/slideshow/read/transition'
import { createRng } from '../src/slideshow/rng'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { reducedTransition } from '../src/slideshow/render/transitions2d'
import { createWebGL2Compositor } from '../src/slideshow/render/webgl2'
import { DEFAULT_FOV_DEG, ease, gridBytes, mul4, project, slideCamera, slidePoint, type Mat4 } from '../src/slideshow/render/three/mesh'
import { registeredTransitions3D, resetTransitions3D, transition3DFor } from '../src/slideshow/render/three/registry'
import { dispatchTransition, transition3DDraw } from '../src/slideshow/render/three/framework'
import {
  DEFORM_ROWS, WIND, airplaneModel, curtainVertex, drapeVertex, fallOverModel, foldClassify, foldPages, origamiModel, prestigeModel, registerDeform3D, warpK, windTile,
} from '../src/slideshow/render/three/deform'
import * as oracle from '../scripts/slideshow-3d-oracles'
import type { CanvasFactory, Compositor, LayerHandle, Lease, SlideShowWorkspace, TransitionSpec } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace, factory and the recording fake WebGL2 context (the S3-1 pattern)
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
  const constant = (name: string) => {
    let v = constants.get(name)
    if (v === undefined) { v = 0x1000 + constants.size; constants.set(name, v) }
    return v
  }
  const overrides: Record<string, unknown> = {
    isContextLost: () => false,
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
        calls.push({ name: prop, args: args.map(a => (a instanceof Float32Array ? [...a] : a)) })
        if (prop.startsWith('create')) return { kind: prop.slice(6), serial: ++serial }
        return undefined
      }
    },
  }) as WebGL2RenderingContext
  return { gl, calls, constant, count: (name: string) => calls.filter(c => c.name === name).length }
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
const A = oracle.ASPECT
const C = slideCamera(DEFAULT_FOV_DEG, A)

function specOf(row: { element: string; preset?: string }, attrs: Record<string, string> = {}): TransitionSpec {
  return { element: row.element, ns: row.element === 'prstTrans' ? 'p15' : 'p14', attrs, durationMs: 1000, advClick: true, ...(row.preset !== undefined ? { preset: row.preset } : {}) }
}
const rowNamed = (name: string) => DEFORM_ROWS.find(r => (r.preset ?? r.element) === name)!

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

/** Runs one shader frame of `spec` at `progress` (aspect 16 : 9) and returns the GL calls the row's `draw` made. */
function drawCalls(spec: TransitionSpec, progress: number, seed = 11): GlCall[] {
  const ws = makeWorkspace()
  const g = webgl(ws, 160, 90)
  expect(dispatchTransition(spec, g.comp, g.sink, 0)).toBe('shader')
  const { from, to } = frames(g.comp, g.ws, 160, 90)
  const rng = createRng(seed)
  const params = transition3DDraw(spec, progress, rng, A)
  const t = transition3DFor(spec)!
  const calls: GlCall[] = []
  const original = t.draw.bind(t)
  const spy = vi.spyOn(t, 'draw').mockImplementation((...args: Parameters<typeof t.draw>) => {
    const start = g.fake.calls.length
    try { original(...args) } finally { calls.push(...g.fake.calls.slice(start)) }
  })
  g.comp.frame([], { from, to, effect: spec.element, progress, params, rng })
  spy.mockRestore()
  expect(calls.length).toBeGreaterThan(0)
  expect(g.sink.list()).toEqual([])
  g.comp.dispose()
  expect(ws.used()).toBe(0)
  return calls
}

/** The `u_mvp` matrices a draw sent, in draw order (one per `drawElements`). */
function mvps(calls: GlCall[]): Array<{ mvp: number[]; opacity: number; uvRect: number[] }> {
  const out: Array<{ mvp: number[]; opacity: number; uvRect: number[] }> = []
  let mvp: number[] = [], opacity = 1, uvRect: number[] = [0, 0, 1, 1]
  for (const c of calls) {
    const name = (c.args[0] as { uniform?: string } | null)?.uniform
    if (c.name === 'uniformMatrix4fv' && name === 'u_mvp') mvp = c.args[2] as number[]
    if (c.name === 'uniform1f' && name === 'u_opacity') opacity = c.args[1] as number
    if (c.name === 'uniform4f' && name === 'u_uvRect') uvRect = c.args.slice(1) as number[]
    if (c.name === 'drawElements') out.push({ mvp, opacity, uvRect })
  }
  return out
}
const uniform = (calls: GlCall[], name: string) => calls.filter(c => (c.args[0] as { uniform?: string } | null)?.uniform === name).map(c => c.args.slice(1))
const near = (got: readonly number[], want: readonly number[], tol = 1e-3) => {
  expect(got.length).toBe(want.length)
  got.forEach((v, i) => expect(Math.abs(v - want[i]), `[${i}] ${v} vs ${want[i]}`).toBeLessThan(tol))
}
const ndcOf = (model: Mat4, u: number, v: number) => project(mul4(C, model), slidePoint(u, v, A))
const worldOf = (model: Mat4, u: number, v: number) => {
  const p = slidePoint(u, v, A)
  return [0, 1, 2].map(r => model[r] * p[0] + model[4 + r] * p[1] + model[8 + r] * p[2] + model[12 + r])
}
/** The oracle table row for `id`, parsed from the algorithm sheet. */
function sheetOracle(id: string): Record<string, number[]> {
  const doc = readFileSync(resolvePath('docs/validation/slideshow-3d-algorithms.md'), 'utf8')
  const line = doc.split('\n').find(l => l.startsWith(`| ${id} |`))!
  const cell = line.split('|')[4]
  const out: Record<string, number[]> = {}
  for (const part of cell.split(';')) {
    const [k, v] = part.split(' = ')
    out[k.trim()] = v.replace(/[()]/g, '').split(',').map(Number)
  }
  return out
}
function foldFractions(p: number, pages: ReturnType<typeof foldPages>, aspect = A) {
  const N = 640, M = 360
  let to = 0, back = 0
  for (let j = 0; j < M; j++) for (let i = 0; i < N; i++) {
    const r = foldClassify(p, ((i + 0.5) / N) * aspect, (j + 0.5) / M, pages)
    if (r === 'to') to++
    if (r === 'back') back++
  }
  return { to: to / (N * M), back: back / (N * M) }
}

const NAMES = ['pageCurlSingle', 'pageCurlDouble', 'peelOff', 'airplane', 'origami', 'curtains', 'drape', 'wind', 'fallOver', 'prestige', 'warp']

// ---------------------------------------------------------------------------
// The brief's tests
// ---------------------------------------------------------------------------
describe('deformation 3-D transitions', () => {
  it('every row in the group has a registered shader and a lookalike', () => {
    expect(DEFORM_ROWS.map(r => r.preset ?? r.element)).toEqual(NAMES)
    expect(registeredTransitions3D()).toEqual([])
    registerDeform3D()
    registerDeform3D() // idempotent: the same objects
    expect(registeredTransitions3D()).toHaveLength(11)
    for (const row of DEFORM_ROWS) {
      const spec = specOf(row)
      const t = transition3DFor(spec)
      expect(t, row.preset ?? row.element).toBeDefined()
      const cat = rowForTransition(spec.ns, spec.element, spec.preset)!
      expect(cat.lookalike, cat.id).toBeTruthy()
      // WebGL2: the shader, no diagnostic.
      const g = webgl()
      expect(dispatchTransition(spec, g.comp, g.sink, 3)).toBe('shader')
      expect(g.sink.list()).toEqual([])
      g.comp.dispose()
      expect(g.ws.used()).toBe(0)
      // Canvas 2D: the catalogued look-alike, diagnosed.
      const ws = makeWorkspace(), sink = createDiagnosticSink()
      const c2 = createCanvas2DCompositor(createCanvas(W, H) as unknown as HTMLCanvasElement, factory, ws, sink)
      expect(dispatchTransition(spec, c2, sink, 3)).toBe('lookalike')
      const look = cat.lookalike!.split(':').at(-1)!
      expect(sink.list()).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 3 })])
      expect(sink.list()[0].message).toContain(`look-alike ${look}`)
      c2.dispose()
    }
    // airplane falls back to fade (S1-22), curtains to split.
    expect(rowForTransition('p15', 'prstTrans', 'airplane')!.lookalike).toBe('transition:p:fade')
    expect(rowForTransition('p15', 'prstTrans', 'curtains')!.lookalike).toBe('transition:p:split')
  })

  it('direction and option attributes change the output', () => {
    registerDeform3D()
    const cases: Array<[string, Record<string, string>, number]> = [
      ['pageCurlSingle', { invX: '1' }, 0.5], ['pageCurlSingle', { invY: '1' }, 0.5],
      ['pageCurlDouble', { invY: '1' }, 0.5], ['peelOff', { invX: '1' }, 0.5],
      ['airplane', { invX: '1' }, 0.75], ['origami', { invX: '1' }, 0.5],
      ['wind', { invX: '1' }, 0.5], ['fallOver', { invX: '1' }, 0.5],
      ['warp', { dir: 'out' }, 0.3],
    ]
    for (const [name, attrs, p] of cases) {
      const row = rowNamed(name)
      const base = drawCalls(specOf(row), p)
      const changed = drawCalls(specOf(row, attrs), p)
      expect(JSON.stringify(changed), `${name} ${JSON.stringify(attrs)}`).not.toEqual(JSON.stringify(base))
      // "true" means the same as "1"; an unknown value is the default.
      const key = Object.keys(attrs)[0]
      if (attrs[key] === '1') {
        expect(JSON.stringify(drawCalls(specOf(row, { [key]: 'true' }), p))).toEqual(JSON.stringify(changed))
        expect(JSON.stringify(drawCalls(specOf(row, { [key]: 'maybe' }), p))).toEqual(JSON.stringify(base))
      }
    }
    expect(JSON.stringify(drawCalls(specOf(rowNamed('warp'), { dir: 'in' }), 0.3))).toEqual(JSON.stringify(drawCalls(specOf(rowNamed('warp')), 0.3)))
    // Mirrors: invX flips the page-curl start corner, invY its row; peelOff's invX starts bottom-right.
    expect(foldPages('pageCurlSingle', {}, A)).toEqual(oracle.curlSingle())
    expect(foldPages('pageCurlSingle', { invX: '1', invY: '1' }, A)).toEqual(oracle.curlSingle(true, true))
    expect(foldPages('pageCurlDouble', { invY: '1' }, A)).toEqual(oracle.curlDouble(true))
    expect(foldPages('peelOff', { invX: '1' }, A)).toEqual(oracle.peel(true))
    // origami's invX and fallOver's invX are the oracle script's mirrors; airplane's flies left.
    for (const k of [0, 1, 2, 3]) near([...origamiModel(0.5, A, true).panels[k]], [...oracle.origami(0.5, true).panels[k]], 1e-9)
    near([...fallOverModel(0.5, true).from], [...oracle.fallOver(0.5, true).from], 1e-9)
    near([...airplaneModel(0.75, A, true).top], [...oracle.airplane(0.75, true).top], 1e-9)
    // wind's invX is an exact mirror image: cell (i, j) mirrored is cell (cols-1-i, j) with x negated.
    const w = windTile(0.5, 2, 5, 99, A, false), m = windTile(0.5, WIND.cols - 1 - 2, 5, 99, A, true)
    expect(m.t).toBe(w.t)
    const c = slidePoint(2.5 / WIND.cols, 5.5 / WIND.rows, A), cm = slidePoint((WIND.cols - 2.5) / WIND.cols, 5.5 / WIND.rows, A)
    const pw = [0, 1, 2].map(r => w.model[r] * c[0] + w.model[4 + r] * c[1] + w.model[8 + r] * c[2] + w.model[12 + r])
    const pm = [0, 1, 2].map(r => m.model[r] * cm[0] + m.model[4 + r] * cm[1] + m.model[8 + r] * cm[2] + m.model[12 + r])
    near(pm, [-pw[0], pw[1], pw[2]], 1e-9)
  })

  it('seeded effects are deterministic', () => {
    registerDeform3D()
    const wind = specOf(rowNamed('wind'))
    const random = vi.spyOn(Math, 'random')
    const a = drawCalls(wind, 0.4, 7), b = drawCalls(wind, 0.4, 7), other = drawCalls(wind, 0.4, 8)
    expect(JSON.stringify(a)).toEqual(JSON.stringify(b))
    expect(JSON.stringify(other)).not.toEqual(JSON.stringify(a))
    // The seed is the shader rng's first value, sent as one uint; the cell randoms come from cellRandom on the GPU.
    const seedOf = (calls: GlCall[]) => uniform(calls, 'u_seed').map(x => x[0])
    expect(seedOf(a)).toHaveLength(1)
    expect(seedOf(a)).toEqual(seedOf(b))
    expect(seedOf(other)).not.toEqual(seedOf(a))
    // The same seed gives the same frame at any progress order (no per-frame random).
    drawCalls(wind, 0.9, 7)
    expect(JSON.stringify(drawCalls(wind, 0.4, 7))).toEqual(JSON.stringify(a))
    // The CPU model too.
    expect(windTile(0.5, 3, 4, 1234, A, false)).toEqual(windTile(0.5, 3, 4, 1234, A, false))
    expect(windTile(0.5, 3, 4, 1234, A, false).model).not.toEqual(windTile(0.5, 3, 4, 4321, A, false).model)
    // The unseeded rows draw nothing from the rng and send no seed.
    for (const name of NAMES.filter(n => n !== 'wind')) {
      const calls = drawCalls(specOf(rowNamed(name)), 0.4, 7)
      expect(JSON.stringify(calls), name).toEqual(JSON.stringify(drawCalls(specOf(rowNamed(name)), 0.4, 8)))
      expect(seedOf(calls)).toEqual([])
    }
    expect(random).not.toHaveBeenCalled()
  })

  it('each shader matches its algorithm-sheet oracle at 50 %', () => {
    registerDeform3D()
    const p = 0.5

    // Fold rows: the per-pixel rule, counted over the sheet's 640 × 360 samples, and the uniforms the shader gets.
    for (const [name, id, pages] of [
      ['pageCurlSingle', 'transition:p15:prstTrans:pageCurlSingle', oracle.curlSingle()],
      ['pageCurlDouble', 'transition:p15:prstTrans:pageCurlDouble', oracle.curlDouble()],
      ['peelOff', 'transition:p15:prstTrans:peelOff', oracle.peel()],
    ] as const) {
      const want = sheetOracle(id)
      const got = foldFractions(p, foldPages(name, {}, A))
      expect(Math.abs(got.to - want['fraction to'][0]), name).toBeLessThan(1e-4)
      expect(Math.abs(got.back - want['fraction back'][0]), name).toBeLessThan(1e-4)
      const calls = drawCalls(specOf(rowNamed(name)), p)
      expect(uniform(calls, 'u_e')).toEqual([[ease(p)]])
      expect(uniform(calls, 'u_aspect')[0][0]).toBeCloseTo(A, 6)
      expect(uniform(calls, 'u_pages')).toEqual([[pages.length]])
      const lines = uniform(calls, 'u_line[0]')[0][0] as number[], ranges = uniform(calls, 'u_range[0]')[0][0] as number[]
      pages.forEach((pg, k) => {
        near(lines.slice(4 * k, 4 * k + 4), [...pg.corner, ...pg.target], 1e-6)
        near(ranges.slice(4 * k, 4 * k + 2), [pg.x0, pg.x1], 1e-6)
      })
      expect(uniform(calls, 'u_peel')).toEqual([[name === 'peelOff' ? 1 : 0]])
      expect(calls.filter(c => c.name === 'drawElements')).toHaveLength(1)
    }

    // airplane: the halves' matrices, as sent, against the reference model and the sheet's numbers.
    {
      const want = sheetOracle('transition:p15:prstTrans:airplane')
      const model = airplaneModel(p, A, false)
      near([...model.top], [...oracle.airplane(p).top], 1e-9)
      near([...model.bottom], [...oracle.airplane(p).bottom], 1e-9)
      near(ndcOf(model.top, 0, 0).slice(0, 2), want['top half uv(0,0) NDC'])
      near(worldOf(model.top, 0, 0), want['top half uv(0,0) world'])
      near(ndcOf(airplaneModel(0.75, A, false).top, 0, 0).slice(0, 2), want['p 0.75 top uv(0,0) NDC'])
      const sent = mvps(drawCalls(specOf(rowNamed('airplane')), p))
      expect(sent).toHaveLength(3) // to flat, then both halves
      near(sent[0].mvp, [...C], 1e-6)
      const halves = sent.slice(1)
      const top = halves.find(h => h.uvRect[1] === 0 && h.uvRect[3] === 0.5)!, bottom = halves.find(h => h.uvRect[1] === 0.5)!
      near(top.mvp, [...mul4(C, model.top)])
      near(bottom.mvp, [...mul4(C, model.bottom)])
      expect(top.opacity).toBeCloseTo(model.opacity, 9)
    }

    // origami: four panels, farthest first.
    {
      const want = sheetOracle('transition:p15:prstTrans:origami')
      const model = origamiModel(p, A, false)
      for (const k of [0, 1, 2, 3]) near([...model.panels[k]], [...oracle.origami(p).panels[k]], 1e-9)
      near(worldOf(model.panels[2], 0.5, 0), want['panel 2 uv(0.5,0) world'])
      near(ndcOf(model.panels[3], 1, 0).slice(0, 2), want['panel 3 uv(1,0) NDC'])
      const sent = mvps(drawCalls(specOf(rowNamed('origami')), p))
      expect(sent).toHaveLength(5)
      const panels = sent.slice(1)
      for (const k of [0, 1, 2, 3]) {
        const got = panels.find(s => Math.abs(s.uvRect[0] - k / 4) < 1e-9)!
        near(got.mvp, [...mul4(C, model.panels[k])])
      }
      // Painter's order: each panel's centre is no nearer than the next one's.
      const depth = (s: { uvRect: number[] }) => { const k = Math.round(s.uvRect[0] * 4); return worldOf(model.panels[k], (k + 0.5) / 4, 0.5)[2] }
      for (let i = 1; i < panels.length; i++) expect(depth(panels[i - 1])).toBeLessThanOrEqual(depth(panels[i]) + 1e-12)
    }

    // curtains, drape, warp: the vertex formulas the GLSL mirrors.
    {
      const want = sheetOracle('transition:p15:prstTrans:curtains')
      expect(curtainVertex(p, 0.5, 0.5, A)[0]).toBeCloseTo(want['x at uv(0.5,0.5) (left half)'][0], 3)
      expect(curtainVertex(p, 0.25, 1, A)[0]).toBeCloseTo(want['x at uv(0.25,1)'][0], 3)
      expect(curtainVertex(1, 0.5, 0, A)[0]).toBeCloseTo(want['p 1 x at uv(0.5,0)'][0], 3)
      for (const [u, v] of [[0.1, 0.2], [0.5, 0.5], [0.7, 0.9], [1, 0]]) expect(curtainVertex(0.37, u, v, A)[0]).toBeCloseTo(oracle.curtainX(0.37, u, v), 9)
      const calls = drawCalls(specOf(rowNamed('curtains')), p)
      expect(uniform(calls, 'u_side')).toEqual([[-1], [1]])
      expect(uniform(calls, 'u_p')).toEqual([[p], [p]])
    }
    {
      const want = sheetOracle('transition:p15:prstTrans:drape')
      expect(drapeVertex(p, 0.5, 0, A)[1]).toBeCloseTo(want['y at uv(0.5,0)'][0], 3)
      expect(drapeVertex(p, 0.5, 1, A)[1]).toBeCloseTo(want['y at uv(0.5,1)'][0], 3)
      for (const v of [0, 0.3, 1]) expect(drapeVertex(0.62, 0.4, v, A)[1]).toBeCloseTo(oracle.drapeY(0.62, v), 9)
      expect(uniform(drawCalls(specOf(rowNamed('drape')), p), 'u_p')).toEqual([[p]])
    }
    {
      const want = sheetOracle('transition:p14:warp')
      expect(warpK(p, A / 2)).toBeCloseTo(want['k at uv(1,0.5)'][0], 4)
      expect((A / 2) * warpK(p, A / 2)).toBeCloseTo(want['world x at uv(1,0.5)'][0], 4)
      expect(warpK(p, 0)).toBeCloseTo(want['k at uv(0.5,0.5)'][0], 4)
      // in: to flat at S(0.7 + 0.3e) first, then the warped from; out: the frames swap and time runs backwards.
      const sent = mvps(drawCalls(specOf(rowNamed('warp')), p))
      expect(sent).toHaveLength(2)
      const s = 0.7 + 0.3 * ease(p)
      near(sent[0].mvp, [...mul4(C, new Float32Array([s, 0, 0, 0, 0, s, 0, 0, 0, 0, s, 0, 0, 0, 0, 1]))], 1e-6)
      const out = drawCalls(specOf(rowNamed('warp'), { dir: 'out' }), 0.3)
      expect(uniform(out, 'u_kp')).toEqual([[1 - 0.3]])
      expect(uniform(drawCalls(specOf(rowNamed('warp')), 0.3), 'u_kp')).toEqual([[0.3]])
    }

    // wind: the tile template (rest, done, cell (7, 0)) at seed 2^31.
    {
      const want = sheetOracle('transition:p15:prstTrans:wind')
      const seed = oracle.SEED
      let rest = 0, done = 0
      for (let j = 0; j < WIND.rows; j++) for (let i = 0; i < WIND.cols; i++) {
        const t = windTile(p, i, j, seed, A, false).t
        if (t === 0) rest++
        if (t === 1) done++
      }
      expect(rest / (WIND.cols * WIND.rows)).toBeCloseTo(want.rest[0], 4)
      expect(done / (WIND.cols * WIND.rows)).toBeCloseTo(want.done[0], 4)
      const cell = windTile(p, 7, 0, seed, A, false)
      expect(cell.t).toBeCloseTo(want['cell (7,0) t'][0], 4)
      near(ndcOf(cell.model, 7.5 / 8, 0.5 / 18).slice(0, 2), want['cell (7,0) centre NDC'])
      near([...cell.model], [...oracle.windTile(p, 7, 0).model], 1e-9)
      // A shader rng whose first value is 0.5 sends the oracle's seed.
      const t = transition3DFor(specOf(rowNamed('wind')))!
      currentGl = fakeGL()
      const half = () => 0.5
      const pool = { aspect: A, grid: () => ({ vao: {}, count: 6 }), program: (key: string) => ({ program: { key }, uniform: (n: string) => ({ uniform: n }) }), draw() {}, drawMesh() {} }
      t.draw(currentGl.gl, {} as WebGLTexture, {} as WebGLTexture, p, {}, half, pool as never)
      expect(uniform(currentGl.calls, 'u_seed')).toEqual([[seed]])
      expect(uniform(currentGl.calls, 'u_grid')).toEqual([[WIND.cols, WIND.rows, WIND.W, WIND.J]])
    }

    // fallOver and prestige: the matrices as sent.
    {
      const want = sheetOracle('transition:p15:prstTrans:fallOver')
      const model = fallOverModel(p, false)
      near(ndcOf(model.from, 0.5, 0).slice(0, 2), want['from uv(0.5,0) NDC'])
      near(worldOf(model.from, 0.5, 0), want['from uv(0.5,0) world'])
      const sent = mvps(drawCalls(specOf(rowNamed('fallOver')), p))
      expect(sent).toHaveLength(2)
      near(sent[0].mvp, [...C], 1e-6)
      near(sent[1].mvp, [...mul4(C, model.from)])
      expect(sent[1].opacity).toBeCloseTo(model.opacity, 9)
    }
    {
      const want = sheetOracle('transition:p15:prstTrans:prestige')
      const model = prestigeModel(p)
      near(ndcOf(model.from, 0.5, 0).slice(0, 2), want['from uv(0.5,0) NDC'])
      near(ndcOf(model.to, 0.5, 0).slice(0, 2), want['to uv(0.5,0) NDC'])
      expect(model.fromOpacity).toBeCloseTo(want['from opacity'][0], 4)
      const sent = mvps(drawCalls(specOf(rowNamed('prestige')), p))
      expect(sent).toHaveLength(2) // from, then to over it
      near(sent[0].mvp, [...mul4(C, model.from)])
      expect(sent[0].opacity).toBeCloseTo(0.5, 9)
      near(sent[1].mvp, [...mul4(C, model.to)])
      expect(sent[1].opacity).toBe(1)
    }
  })
})

// ---------------------------------------------------------------------------
// Contracts
// ---------------------------------------------------------------------------
describe('deformation 3-D transition contracts', () => {
  it('prepare leases every grid before creating it and keeps mesh resolution bounded', () => {
    registerDeform3D()
    for (const row of DEFORM_ROWS) {
      const log: Array<{ tag: string; bytes: number }> = []
      const ws = makeWorkspace(256 * 1024 * 1024, log)
      const g = webgl(ws)
      const vaosBefore = g.fake.count('createVertexArray')
      expect(dispatchTransition(specOf(row), g.comp, g.sink, 0)).toBe('shader')
      const meshes = log.filter(l => l.tag === 'webgl2:mesh')
      expect(meshes.length, row.preset ?? row.element).toBeGreaterThan(0)
      // One VAO per leased grid, and no grid above 32 × 18.
      expect(g.fake.count('createVertexArray') - vaosBefore).toBe(meshes.length)
      for (const m of meshes) expect(m.bytes).toBeLessThanOrEqual(gridBytes(32, 18))
      g.comp.dispose()
      expect(ws.used()).toBe(0)
    }
  })

  it('a refused mesh lease sends the row to its look-alike, holding nothing', () => {
    registerDeform3D()
    for (const row of DEFORM_ROWS) {
      // Room for the compositor's own state but not the row's grids (each at least gridBytes(1, 1)).
      const ws = makeWorkspace()
      const g = webgl(ws)
      const filler = ws.lease(ws.limit - ws.used() - (gridBytes(1, 1) - 1), 'filler')!
      const before = g.fake.count('createVertexArray')
      expect(dispatchTransition(specOf(row), g.comp, g.sink, 0)).toBe('lookalike')
      expect(g.fake.count('createVertexArray')).toBe(before)
      expect(g.sink.list().map(d => d.code)).toEqual(['transition-approximate'])
      expect(g.sink.list()[0].message).toMatch(/could not be prepared/)
      filler.release()
      g.comp.dispose()
      expect(ws.used()).toBe(0)
    }
  })

  it('reduced motion never reaches a deformation shader', () => {
    registerDeform3D()
    for (const row of DEFORM_ROWS) {
      const reduced = reducedTransition(specOf(row))
      const g = webgl()
      expect(dispatchTransition(reduced, g.comp, g.sink, 0)).toBe('lookalike')
      expect(transition3DFor(reduced)).toBeUndefined()
      expect(reduced.durationMs).toBeLessThanOrEqual(200)
      g.comp.dispose()
    }
  })

  it('every model is continuous into its ends', () => {
    const eps = 1e-4
    const rest = (model: Mat4, u: number, v: number) => ndcOf(model, u, v).slice(0, 2)
    const flat = (u: number, v: number) => [2 * u - 1, 1 - 2 * v]
    for (const [u, v] of [[0, 0], [1, 0], [0.3, 0.8], [1, 1]]) {
      near(rest(airplaneModel(eps, A, false).top, u, v / 2), flat(u, v / 2))
      for (let k = 0; k < 4; k++) near(rest(origamiModel(eps, A, false).panels[k], u, v), flat(u, v))
      near(rest(fallOverModel(eps, false).from, u, v), flat(u, v))
      near(rest(prestigeModel(eps).from, u, v), flat(u, v))
      near(rest(prestigeModel(1 - eps).to, u, v), flat(u, v))
      near(curtainVertex(eps, u, v, A).slice(0, 2), slidePoint(u, v, A).slice(0, 2))
      near(drapeVertex(eps, u, v, A).slice(0, 2), slidePoint(u, v, A).slice(0, 2))
    }
    // Gone (or transparent) at the far end.
    expect(airplaneModel(1 - eps, A, false).opacity).toBeLessThan(1e-3)
    expect(origamiModel(1 - eps, A, false).opacity).toBeLessThan(1e-3)
    expect(fallOverModel(1 - eps, false).opacity).toBeLessThan(1e-3)
    expect(prestigeModel(1 - eps).fromOpacity).toBeLessThan(1e-3)
    expect(curtainVertex(1, 0.5, 0, A)[0]).toBeLessThan(-A / 2)
    expect(curtainVertex(1, 0.5001, 0, A)[0]).toBeGreaterThan(A / 2)
    expect(drapeVertex(1, 0.5, 0, A)[1]).toBeLessThan(-0.5)
    expect(warpK(eps, 1)).toBeGreaterThan(0.999)
    expect(warpK(1 - eps, 0)).toBeLessThan(1e-3)
    for (const name of ['pageCurlSingle', 'pageCurlDouble', 'peelOff']) {
      expect(foldFractions(eps, foldPages(name, {}, A)).to).toBeLessThan(1e-3)
      expect(foldFractions(1 - eps, foldPages(name, {}, A)).to).toBeGreaterThan(0.999)
    }
    for (let j = 0; j < WIND.rows; j++) for (let i = 0; i < WIND.cols; i++) {
      expect(windTile(0, i, j, 5, A, false).t).toBe(0)
      expect(windTile(1, i, j, 5, A, false).opacity).toBe(0)
    }
  })
})

// ---------------------------------------------------------------------------
// The corpus airplane deck
// ---------------------------------------------------------------------------
const corpus = resolvePath('corpus') + '/'
const corpusAbsent = !existsSync(corpus + 'lo-airplane-transition.pptx')
describe.skipIf(corpusAbsent)(`corpus airplane${corpusAbsent ? ' — run bun scripts/fetch-corpus.ts' : ''}`, () => {
  it('airplane-transition.pptx plays the airplane shader', async () => {
    const zip = await JSZip.loadAsync(readFileSync(corpus + 'lo-airplane-transition.pptx'))
    const sink = createDiagnosticSink()
    const spec = readTransition(parseXmlOrdered(await zip.file('ppt/slides/slide2.xml')!.async('string')), 1, sink)!
    expect(spec).toMatchObject({ element: 'prstTrans', ns: 'p15', preset: 'airplane' })
    registerDeform3D()
    const airplane = transition3DFor(spec)!
    expect(airplane.preset).toBe('airplane')

    // WebGL2: dispatched to the shader, which draws every in-between frame with the effect's rng.
    const g = webgl(makeWorkspace(), 160, 90)
    expect(dispatchTransition(spec, g.comp, g.sink, 1)).toBe('shader')
    const { from, to } = frames(g.comp, g.ws, 160, 90)
    const draw = vi.spyOn(airplane, 'draw')
    const rng = createRng(0x5eed)
    for (const p of [0, 0.25, 0.5, 0.75, 1]) {
      g.fake.calls.length = 0
      g.comp.frame([], { from, to, effect: spec.element, progress: p, params: transition3DDraw(spec, p, rng, 160 / 90), rng })
      // to flat plus both halves in between; the flat frame alone at the ends.
      expect(g.fake.count('drawElements'), `p ${p}`).toBe(p === 0 || p === 1 ? 0 : 3)
    }
    expect(draw.mock.calls.map(c => c[3])).toEqual([0.25, 0.5, 0.75])
    expect(g.sink.list()).toEqual([])
    g.comp.dispose()
    expect(g.ws.used()).toBe(0)

    // Canvas 2D: its catalogued look-alike, fade, with one diagnostic.
    const ws = makeWorkspace(), c2sink = createDiagnosticSink()
    const c2 = createCanvas2DCompositor(createCanvas(W, H) as unknown as HTMLCanvasElement, factory, ws, c2sink)
    expect(dispatchTransition(spec, c2, c2sink, 1)).toBe('lookalike')
    expect(c2sink.list()).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 1, effect: 'prstTrans:airplane' })])
    expect(c2sink.list()[0].message).toContain('look-alike fade')
    c2.dispose()
  })
})
