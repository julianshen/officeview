/**
 * S3-2: the geometric 3-D transitions (slideshow spec 6.3; algorithm sheet
 * docs/validation/slideshow-3d-algorithms.md, "S3-2: geometric rows"):
 * prism (cube / box / rotate / orbit variants), flip, gallery, doors,
 * window, switch, conveyor and ferris. `pan` is a 2-D row (S1-15), not here.
 *
 * Node has no WebGL2, so the compositor runs on a recording fake GL (the
 * S1-14 / S3-1 pattern). These tests pin the models against the sheet's 50 %
 * oracles and the reference script, the matrices actually sent to the GPU,
 * leases, dispatch, fallbacks and reduced motion, and drive one geometric
 * row end to end through `runShow` and the real stage. Pixels are checked in
 * Chromium by tests/browser/slideshow-3d-geometric.spec.ts.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { CATALOGUE, rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createRng } from '../src/slideshow/rng'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../src/slideshow/render/webgl2'
import { reducedTransition } from '../src/slideshow/render/transitions2d'
import {
  DEFAULT_FOV_DEG, createMeshPool, gridBytes, mul4, project, slideCamera, slidePoint, type Mat4,
} from '../src/slideshow/render/three/mesh'
import { registeredTransitions3D, resetTransitions3D, transition3DFor, type Transition3D } from '../src/slideshow/render/three/registry'
import { dispatchTransition, transition3DDraw } from '../src/slideshow/render/three/framework'
import { GEOMETRIC_ELEMENTS, geometricPieces, registerGeometric3D, type GeometricPiece } from '../src/slideshow/render/three/geometric'
import { runShow } from '../src/slideshow'
import { sessionDeps } from '../src/slideshow/player/session'
import { ManualClock } from '../src/slideshow/clock'
import type { ShowStart } from '../src/components/present'
import { buildPptxDeck } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx, pptxPackageHandle } from '../src/pptx/parse'
import type { CanvasFactory, Lease, SlideShowWorkspace, TransitionSpec } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace, factory and a recording fake WebGL2 context (as in slideshow-3d-framework.test.ts)
// ---------------------------------------------------------------------------
interface Pool { used: number; limit: number }
function poolLease(pool: Pool, bytes: number, tag: string, log?: Array<{ tag: string; bytes: number; at: number }>, at = () => 0): Lease | undefined {
  if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
  pool.used += bytes
  log?.push({ tag, bytes, at: at() })
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
function makeWorkspace(limit = 256 * 1024 * 1024, log?: Array<{ tag: string; bytes: number; at: number }>, at?: () => number): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  return { pool, limit, lease: (b, t) => poolLease(pool, b, t, log, at), used: () => pool.used, reserveTerminal: () => undefined }
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
        calls.push({ name: prop, args })
        if (prop.startsWith('create')) return { kind: prop.slice(6), serial: ++serial }
        return undefined
      }
    },
  }) as WebGL2RenderingContext
  return { gl, calls, constant, count: (name: string) => calls.filter(c => c.name === name).length }
}

/** One indexed draw as the GPU received it: the bound texture, the uniforms in force and the cull state. */
interface Drawn { tex: unknown; mvp: Mat4; uvRect: [number, number, number, number]; shade: number; opacity: number; cull: 'none' | 'back' | 'front' }
function drawsOf(fake: ReturnType<typeof fakeGL>, from = 0): Drawn[] {
  const out: Drawn[] = []
  let tex: unknown, mvp = new Float32Array(16), uvRect: Drawn['uvRect'] = [0, 0, 1, 1], shade = 1, opacity = 1
  let culling = false, mode = fake.constant('BACK')
  for (const c of fake.calls.slice(from)) {
    const loc = (c.args[0] as { uniform?: string } | null)?.uniform
    if (c.name === 'bindTexture') tex = c.args[1]
    else if (c.name === 'uniformMatrix4fv' && loc === 'u_mvp') mvp = new Float32Array(c.args[2] as Float32Array)
    else if (c.name === 'uniform4f' && loc === 'u_uvRect') uvRect = c.args.slice(1, 5) as Drawn['uvRect']
    else if (c.name === 'uniform1f' && loc === 'u_shade') shade = c.args[1] as number
    else if (c.name === 'uniform1f' && loc === 'u_opacity') opacity = c.args[1] as number
    else if (c.name === 'enable' && c.args[0] === fake.constant('CULL_FACE')) culling = true
    else if (c.name === 'disable' && c.args[0] === fake.constant('CULL_FACE')) culling = false
    else if (c.name === 'cullFace') mode = c.args[0] as number
    else if (c.name === 'drawElements') out.push({ tex, mvp, uvRect, shade, opacity, cull: !culling ? 'none' : mode === fake.constant('FRONT') ? 'front' : 'back' })
  }
  return out
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
afterEach(() => {
  vi.restoreAllMocks()
  currentGl = null
  sessionDeps.backend = undefined
  resetTransitions3D()
  document.body.replaceChildren()
})

const A = 16 / 9
const C = slideCamera(DEFAULT_FOV_DEG, A)
function spec(element: string, attrs: Record<string, string> = {}, ns: TransitionSpec['ns'] = 'p14'): TransitionSpec {
  return { element, ns, attrs, durationMs: 1000, advClick: true }
}
const ndcOf = (m: Mat4, u: number, v: number) => project(mul4(C, m), slidePoint(u, v, A))
const worldOf = (m: Mat4, u: number, v: number) => {
  const p = slidePoint(u, v, A)
  return [0, 1, 2].map(r => m[r] * p[0] + m[4 + r] * p[1] + m[8 + r] * p[2] + m[12 + r])
}
const piece = (pieces: GeometricPiece[], role: GeometricPiece['role']) => {
  const p = pieces.find(x => x.role === role)
  if (!p) throw new Error(`no ${role} piece`)
  return p
}
const near = (got: readonly number[], want: readonly number[], digits = 3) => {
  expect(got.length).toBeGreaterThanOrEqual(want.length)
  want.forEach((w, i) => expect(got[i], `component ${i} of [${got.map(x => x.toFixed(4))}] vs [${want}]`).toBeCloseTo(w, digits))
}

/** Runs `t.draw` on a fake GL with a real mesh pool at aspect `aspect`; returns the draws. */
function drawOnce(t: Transition3D, progress: number, attrs: Record<string, string> = {}, aspect = A, seed = 7) {
  const fake = fakeGL()
  const ws = makeWorkspace()
  const owner = createMeshPool(fake.gl, ws)
  owner.setAspect(aspect)
  const from = { tex: 'from' } as unknown as WebGLTexture, to = { tex: 'to' } as unknown as WebGLTexture
  t.prepare?.(owner.pool)
  const mark = fake.calls.length
  t.draw(fake.gl, from, to, progress, attrs, createRng(seed), owner.pool)
  return { fake, ws, owner, draws: drawsOf(fake, mark), calls: fake.calls.slice(mark) }
}

const IDS = GEOMETRIC_ELEMENTS.map(e => `transition:p14:${e}`)

// ---------------------------------------------------------------------------
// The brief's tests
// ---------------------------------------------------------------------------
describe('geometric 3-D transitions (S3-2)', () => {
  it('every row in the group has a registered shader and a lookalike', () => {
    expect([...GEOMETRIC_ELEMENTS].sort()).toEqual(['conveyor', 'doors', 'ferris', 'flip', 'gallery', 'prism', 'switch', 'window'])
    registerGeometric3D()
    registerGeometric3D() // idempotent
    expect(registeredTransitions3D().map(t => t.element).sort()).toEqual([...GEOMETRIC_ELEMENTS].sort())
    for (const id of IDS) {
      const row = CATALOGUE.find(r => r.id === id)
      expect(row, id).toBeDefined()
      expect(row!.stage, id).toBe(3)
      expect(row!.lookalike, id).toBeTruthy()
      const t = transition3DFor(spec(row!.element!))
      expect(t, id).toBeDefined()
      expect(typeof t!.prepare).toBe('function')
    }
    // pan is a 2-D row (S1-15): no shader, and the registry refuses one.
    expect(rowForTransition('p14', 'pan')!.stage).not.toBe(3)
    expect(transition3DFor(spec('pan'))).toBeUndefined()
  })

  it('direction and option attributes change the output', () => {
    registerGeometric3D()
    const mvps = (element: string, attrs: Record<string, string>, p = 0.4) =>
      JSON.stringify(drawOnce(transition3DFor(spec(element))!, p, attrs).draws.map(d => [...d.mvp, ...d.uvRect, d.cull]))
    const variants: Array<[string, Array<Record<string, string>>]> = [
      ['prism', [{}, { dir: 'r' }, { dir: 'u' }, { dir: 'd' }, { isInverted: '1' }, { isContent: '1' }, { isInverted: 'true', isContent: 'true' }]],
      ['flip', [{}, { dir: 'r' }]],
      ['gallery', [{}, { dir: 'r' }]],
      ['doors', [{}, { dir: 'horz' }]],
      ['window', [{}, { dir: 'horz' }]],
      ['switch', [{}, { dir: 'r' }]],
      ['conveyor', [{}, { dir: 'r' }]],
      ['ferris', [{}, { dir: 'r' }]],
    ]
    for (const [element, list] of variants) {
      const seen = list.map(a => mvps(element, a))
      expect(new Set(seen).size, `${element}: ${list.map(a => JSON.stringify(a)).join(' ')}`).toBe(list.length)
    }
    // Defaults: an omitted or unknown value is the sheet's default; "0" / "false" are false.
    expect(mvps('prism', { dir: 'l' })).toBe(mvps('prism', {}))
    expect(mvps('prism', { dir: 'zz' })).toBe(mvps('prism', {}))
    expect(mvps('prism', { isInverted: '0', isContent: 'false' })).toBe(mvps('prism', {}))
    expect(mvps('doors', { dir: 'vert' })).toBe(mvps('doors', {}))
    expect(mvps('flip', { dir: 'l' })).toBe(mvps('flip', {}))
    // window is not doors: its halves part and its incoming frame starts further back.
    expect(mvps('window', {})).not.toBe(mvps('doors', {}))
    // Mirrored directions mirror x.
    const l = geometricPieces('flip', 0.3, {}, A), r = geometricPieces('flip', 0.3, { dir: 'r' }, A)
    expect(ndcOf(piece(r, 'from').model, 1, 0.5)[0]).toBeCloseTo(-ndcOf(piece(l, 'from').model, 0, 0.5)[0], 5)
  })

  it('seeded effects are deterministic', () => {
    // No geometric row is seeded (sheet: "Seeded parameters: none"): a frame is a pure function of
    // progress, attributes and aspect. Any rng, any number of times, gives the same GPU calls.
    registerGeometric3D()
    for (const element of GEOMETRIC_ELEMENTS) {
      const t = transition3DFor(spec(element))!
      const strip = (calls: GlCall[]) => JSON.stringify(calls.map(c => [c.name, c.args.map(a => (a instanceof Float32Array ? [...a] : a))]))
      const a = strip(drawOnce(t, 0.37, {}, A, 1).calls)
      const b = strip(drawOnce(t, 0.37, {}, A, 99).calls)
      expect(b, element).toBe(a)
      expect(strip(drawOnce(t, 0.37, {}, A, 1).calls), element).toBe(a)
      // The shader draws no values from the rng (a later effect would see a shifted stream otherwise).
      const rng = createRng(5)
      const spy = vi.fn(rng)
      const fake = fakeGL()
      const owner = createMeshPool(fake.gl, makeWorkspace())
      t.draw(fake.gl, {} as WebGLTexture, {} as WebGLTexture, 0.5, {}, spy as never, owner.pool)
      expect(spy, element).not.toHaveBeenCalled()
    }
  })

  it('each shader matches its algorithm-sheet oracle at 50 %', () => {
    // The literal numbers of the sheet's generated oracle table (aspect 16 : 9, fov 45°, p = ½).
    const P = 0.5
    const g = (element: string, attrs: Record<string, string> = {}) => geometricPieces(element, P, attrs, A)
    near(ndcOf(piece(g('prism'), 'from').model, 1, 0), [0, 1.4389])
    near(ndcOf(piece(g('prism'), 'to').model, 1, 1), [0.8145, -0.5759])
    near(ndcOf(piece(g('prism', { isInverted: '1' }), 'from').model, 1, 0), [0, 0.7663])
    near(ndcOf(piece(g('prism', { dir: 'u' }), 'from').model, 0, 0), [-0.7071, 1])
    near(ndcOf(piece(g('flip'), 'from').model, 0, 0), [0, 0.465])
    near(worldOf(piece(g('flip'), 'from').model, 0, 0), [0, 0.5, -1.3889])
    near(ndcOf(piece(g('gallery'), 'from').model, 1, 0.5), [-0.1291, 0])
    near(ndcOf(piece(g('gallery'), 'to').model, 0, 0.5), [0.0918, 0])
    near(ndcOf(piece(g('doors'), 'first').model, 0.5, 0), [-0.6111, 2.0864])
    near(ndcOf(piece(g('doors'), 'to').model, 0, 0), [-0.8009, 0.8009])
    near(ndcOf(piece(g('doors', { dir: 'horz' }), 'first').model, 0, 0.5), [-1.4142, 0.4142])
    near(ndcOf(piece(g('window'), 'first').model, 0.5, 0), [-1.1327, 2.0864])
    near(ndcOf(piece(g('window'), 'to').model, 0, 0), [-0.668, 0.668])
    near(ndcOf(piece(g('switch'), 'from').model, 0.5, 0.5), [-0.7348, 0])
    near(ndcOf(piece(g('switch'), 'to').model, 0.5, 0.5), [0.7348, 0])
    near(ndcOf(piece(g('switch'), 'from').model, 1, 0), [-0.254, 0.9248])
    near(ndcOf(piece(g('conveyor'), 'from').model, 1, 0.5), [-0.0885, 0])
    near(ndcOf(piece(g('conveyor'), 'to').model, 0, 0.5), [0.0711, 0])
    near(ndcOf(piece(g('ferris'), 'from').model, 0.5, 0.5), [-0.7208, -0.3434])
    near(ndcOf(piece(g('ferris'), 'to').model, 0.5, 0.5), [0.7208, -0.3434])

    // And the GPU receives exactly those matrices: the u_mvp of each draw projects the oracle points.
    registerGeometric3D()
    const drawn = (element: string, attrs: Record<string, string> = {}) => drawOnce(transition3DFor(spec(element))!, P, attrs).draws
    const at = (d: Drawn, u: number, v: number) => project(d.mvp, slidePoint(u, v, A))
    const tex = (d: Drawn) => (d.tex as { tex: string }).tex
    const half = (ds: Drawn[], rect: number[]) => ds.find(d => tex(d) === 'from' && d.uvRect.every((x, i) => Math.abs(x - rect[i]) < 1e-9))!
    near(at(drawn('prism').find(d => tex(d) === 'from')!, 1, 0), [0, 1.4389])
    near(at(drawn('flip').find(d => tex(d) === 'from')!, 0, 0), [0, 0.465])
    near(at(half(drawn('doors'), [0, 0, 0.5, 1]), 0.5, 0), [-0.6111, 2.0864])
    near(at(half(drawn('doors', { dir: 'horz' }), [0, 0, 1, 0.5]), 0, 0.5), [-1.4142, 0.4142])
    near(at(drawn('window').find(d => tex(d) === 'to')!, 0, 0), [-0.668, 0.668])
    near(at(drawn('ferris').find(d => tex(d) === 'to')!, 0.5, 0.5), [0.7208, -0.3434])
  })
})

// ---------------------------------------------------------------------------
// Models against the reference script, ends, draw order
// ---------------------------------------------------------------------------
describe('geometric models', () => {
  it('match the reference script (scripts/slideshow-3d-oracles.ts) at every progress and attribute', async () => {
    const ref = await import('../scripts/slideshow-3d-oracles')
    const close = (a: Mat4, b: Mat4, what: string) => { for (let i = 0; i < 16; i++) expect(a[i], `${what}[${i}]`).toBeCloseTo(b[i], 6) }
    for (const p of [0.05, 0.2, 0.5, 0.63, 0.9]) {
      for (const dir of ['l', 'r', 'u', 'd']) {
        for (const [inv, con] of [[false, false], [true, false], [false, true], [true, true]]) {
          const attrs = { dir, ...(inv ? { isInverted: '1' } : {}), ...(con ? { isContent: '1' } : {}) }
          const want = ref.prism(p, dir, inv, con), got = geometricPieces('prism', p, attrs, ref.ASPECT)
          close(piece(got, 'from').model, want.from, `prism ${JSON.stringify(attrs)} from @${p}`)
          close(piece(got, 'to').model, want.to, `prism ${JSON.stringify(attrs)} to @${p}`)
        }
      }
      for (const dir of ['l', 'r']) {
        for (const [name, fn] of [['flip', ref.flip], ['gallery', ref.gallery], ['switch', ref.switchRow], ['conveyor', ref.conveyor], ['ferris', ref.ferris]] as const) {
          const want = fn(p, dir), got = geometricPieces(name, p, { dir }, ref.ASPECT)
          close(piece(got, 'from').model, want.from, `${name} ${dir} from @${p}`)
          close(piece(got, 'to').model, want.to, `${name} ${dir} to @${p}`)
        }
      }
      for (const dir of ['vert', 'horz']) {
        for (const windowed of [false, true]) {
          const want = ref.doors(p, dir, windowed), got = geometricPieces(windowed ? 'window' : 'doors', p, { dir }, ref.ASPECT)
          close(piece(got, 'first').model, want.first, `${windowed ? 'window' : 'doors'} ${dir} first @${p}`)
          close(piece(got, 'second').model, want.second, `${windowed ? 'window' : 'doors'} ${dir} second @${p}`)
          close(piece(got, 'to').model, want.to, `${windowed ? 'window' : 'doors'} ${dir} to @${p}`)
        }
      }
    }
  })

  it('every row is continuous into its ends: the outgoing frame at rest as p → 0, the incoming one as p → 1', () => {
    const corners: Array<[number, number]> = [[0, 0], [1, 0], [0, 1], [1, 1], [0.5, 0.5]]
    const rest = (u: number, v: number) => project(C, slidePoint(u, v, A))
    const attrsOf: Record<string, Array<Record<string, string>>> = {
      prism: [{}, { dir: 'r' }, { dir: 'u' }, { dir: 'd' }, { isInverted: '1' }, { isContent: '1' }],
      doors: [{}, { dir: 'horz' }], window: [{}, { dir: 'horz' }],
    }
    for (const element of GEOMETRIC_ELEMENTS) {
      for (const attrs of attrsOf[element] ?? [{}, { dir: 'r' }]) {
        for (const [p, face] of [[1e-4, 'from'], [1 - 1e-4, 'to']] as const) {
          const pieces = geometricPieces(element, p, attrs, A).filter(x => x.face === face && (face === 'to' || x.role !== 'to'))
          for (const pc of pieces) {
            const r = pc.uvRect ?? [0, 0, 1, 1]
            for (const [cu, cv] of corners) {
              const u = r[0] + cu * r[2], v = r[1] + cv * r[3]
              near(ndcOf(pc.model, u, v).slice(0, 2), rest(u, v).slice(0, 2), 2)
            }
          }
        }
      }
    }
  })

  it('draws back to front with the sheet\'s culling: flat backdrop first, farther faces first', () => {
    registerGeometric3D()
    const tex = (d: Drawn) => (d.tex as { tex: string }).tex
    // doors / window: the incoming frame (behind) first, then both halves, each two-sided (front, then a shaded back).
    for (const element of ['doors', 'window']) {
      const ds = drawOnce(transition3DFor(spec(element))!, 0.5).draws
      expect(tex(ds[0]), element).toBe('to')
      expect(ds.slice(1).every(d => tex(d) === 'from'), element).toBe(true)
      expect(ds.slice(1).map(d => d.cull), element).toEqual(['back', 'front', 'back', 'front'])
      expect(ds[2].shade).toBeLessThan(ds[1].shade)
    }
    // prism and flip cull back faces.
    for (const element of ['prism', 'flip']) {
      for (const d of drawOnce(transition3DFor(spec(element))!, 0.3).draws) expect(d.cull, element).toBe('back')
    }
    // switch: the frame farther back first; on the tie at ½, from first.
    const order = (p: number) => drawOnce(transition3DFor(spec('switch'))!, p).draws.map(tex)
    expect(order(0.2)).toEqual(['to', 'from'])
    expect(order(0.5)).toEqual(['from', 'to'])
    expect(order(0.8)).toEqual(['from', 'to'])
    // prism isInverted: the farther face (by the view-space z of its centre) first.
    const inv = geometricPieces('prism', 0.3, { isInverted: '1' }, A)
    const zs = inv.map(pc => worldOf(pc.model, 0.5, 0.5)[2])
    expect(zs[0]).toBeLessThanOrEqual(zs[1])
    // The camera follows the slide's aspect, never a fixed 16 : 9.
    const wide = drawOnce(transition3DFor(spec('ferris'))!, 0.5, {}, 2.4).draws.find(d => tex(d) === 'to')!
    const want = project(mul4(slideCamera(DEFAULT_FOV_DEG, 2.4), piece(geometricPieces('ferris', 0.5, {}, 2.4), 'to').model), slidePoint(0.5, 0.5, 2.4))
    near(project(wide.mvp, slidePoint(0.5, 0.5, 2.4)), want, 5)
  })
})

// ---------------------------------------------------------------------------
// Compositor: leases, dispatch, fallbacks, reduced motion
// ---------------------------------------------------------------------------
function webgl(ws: SlideShowWorkspace = makeWorkspace()) {
  currentGl = fakeGL()
  const fake = currentGl
  const sink = createDiagnosticSink()
  const canvas = document.createElement('canvas')
  canvas.width = 64; canvas.height = 36
  const comp = createWebGL2Compositor(canvas, ws, { sink, factory })!
  comp.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: 64, deviceHeight: 36 })
  return { comp, fake, sink, ws }
}

describe('geometric rows on the compositor', () => {
  it('lease their one mesh before creating any GPU object, and a refused lease sends the row to its look-alike', () => {
    registerGeometric3D()
    for (const element of GEOMETRIC_ELEMENTS) {
      const log: Array<{ tag: string; bytes: number; at: number }> = []
      let g: ReturnType<typeof webgl> | undefined
      const ws = makeWorkspace(256 * 1024 * 1024, log, () => g?.fake.count('createVertexArray') ?? -1)
      g = webgl(ws)
      const vaosBefore = g.fake.count('createVertexArray')
      expect(dispatchTransition(spec(element), g.comp, g.sink, 1), element).toBe('shader')
      const mesh = log.filter(l => l.tag === 'webgl2:mesh')
      expect(mesh, element).toEqual([{ tag: 'webgl2:mesh', bytes: gridBytes(1, 1), at: vaosBefore }])
      expect(g.fake.count('createVertexArray'), element).toBe(vaosBefore + 1)
      expect(g.sink.list(), element).toEqual([])
      g.comp.dispose()
      expect(ws.used(), element).toBe(0)
    }
    // A workspace with no room: dispatch is the look-alike, diagnosed, and nothing is held.
    const roomy = makeWorkspace()
    const tight: SlideShowWorkspace = { ...roomy, lease: (b, t) => (t === 'webgl2:mesh' ? undefined : roomy.lease(b, t)) }
    const g = webgl(tight)
    const vaoCount = g.fake.count('createVertexArray')
    expect(dispatchTransition(spec('doors'), g.comp, g.sink, 3)).toBe('lookalike')
    expect(g.sink.list()).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 3, effect: 'doors' })])
    expect(g.sink.list()[0].message).toMatch(/split/)
    expect(g.fake.count('createVertexArray')).toBe(vaoCount)
    g.comp.dispose()
    expect(roomy.used()).toBe(0)
  })

  it('without WebGL2 each row draws its catalogued 2-D look-alike, with a diagnostic', () => {
    registerGeometric3D()
    for (const element of GEOMETRIC_ELEMENTS) {
      const sink = createDiagnosticSink()
      const c2 = createCanvas2DCompositor(createCanvas(64, 36) as unknown as HTMLCanvasElement, factory, makeWorkspace(), sink)
      expect(dispatchTransition(spec(element), c2, sink, 2), element).toBe('lookalike')
      const want = rowForTransition('p14', element)!.lookalike!.split(':').pop()!
      expect(want, element).toBe(element === 'doors' || element === 'window' ? 'split' : 'push')
      expect(sink.list(), element).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 2, effect: element })])
      expect(sink.list()[0].message, element).toContain(want)
      c2.dispose()
    }
  })

  it('reduced motion never reaches a geometric shader and never adds motion', () => {
    registerGeometric3D()
    const g = webgl()
    for (const element of GEOMETRIC_ELEMENTS) {
      const t = transition3DFor(spec(element))!
      const draw = vi.spyOn(t, 'draw')
      const reduced = reducedTransition(spec(element))
      expect(reduced.element, element).toBe('fade')
      expect(reduced.durationMs, element).toBeLessThanOrEqual(200)
      expect(dispatchTransition(reduced, g.comp, g.sink, 0), element).toBe('lookalike')
      expect(draw, element).not.toHaveBeenCalled()
    }
    expect(g.sink.list()).toEqual([])
    g.comp.dispose()
  })

  it('a dispatched row draws through the compositor at mid-progress and not at its ends', () => {
    registerGeometric3D()
    const g = webgl()
    const layer = (color: string) => {
      const c = createCanvas(64, 36); const x = c.getContext('2d'); x.fillStyle = color; x.fillRect(0, 0, 64, 36)
      const r = g.comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: 64, h: 36 }, scale: 1, bytes: 64 * 36 * 4 }, g.ws.lease(64 * 36 * 4, 'layer')!)
      if ('refused' in r) throw new Error(r.refused)
      return r
    }
    const from = layer('#f00'), to = layer('#00f')
    for (const element of GEOMETRIC_ELEMENTS) {
      const s = spec(element)
      expect(dispatchTransition(s, g.comp, g.sink, 0)).toBe('shader')
      const t = transition3DFor(s)!
      const draw = vi.spyOn(t, 'draw')
      for (const p of [0, 0.5, 1]) {
        const rng = createRng(3)
        g.comp.frame([], { from, to, effect: element, progress: p, params: transition3DDraw(s, p, rng, 64 / 36), rng })
      }
      expect(draw.mock.calls.map(c => c[3]), element).toEqual([0.5])
    }
    expect(g.sink.list()).toEqual([])
    g.comp.dispose()
    expect(g.ws.used()).toBe(0)
  })
})

// ---------------------------------------------------------------------------
// End to end: runShow → the real stage → dispatch → the WebGL2 compositor → the geometric shader
// ---------------------------------------------------------------------------
describe('a geometric row through the stage (S3-1 concern: no shader was driven through stage.ts)', () => {
  const clock = new ManualClock()
  const flush = async () => { for (let i = 0; i < 20; i++) await new Promise((r) => setTimeout(r, 0)) }
  const FLIP = '<p:transition spd="slow" p14:dur="1000"><p14:flip dir="r"/></p:transition>'

  async function play(opts: { motion?: 'auto' | 'reduced'; canvas2d?: boolean } = {}) {
    vi.spyOn(sessionDeps, 'clock').mockImplementation(() => clock)
    vi.spyOn(sessionDeps, 'frames').mockImplementation(() => () => {})
    vi.spyOn(window, 'innerWidth', 'get').mockImplementation(() => 640)
    vi.spyOn(window, 'innerHeight', 'get').mockImplementation(() => 360)
    vi.spyOn(window, 'devicePixelRatio', 'get').mockImplementation(() => 1)
    if (opts.canvas2d) sessionDeps.backend = 'canvas2d'
    currentGl = fakeGL()
    const doc = await parsePptx(await OfficePackage.load(await buildPptxDeck({
      slideSize: [9144000, 5143500],
      slides: [
        { shapes: [{ id: 2, off: ['0', '0'], ext: ['914400', '914400'], fill: '336699' }] },
        { shapes: [{ id: 2, off: ['914400', '914400'], ext: ['914400', '914400'], fill: 'cc3322' }], transition: { xml: FLIP, wrapInAlternateContent: { requires: 'p14', fallbackXml: '<p:transition spd="slow"><p:push/></p:transition>' } } },
      ],
    } as never)))
    expect(doc.slides).toHaveLength(2)
    const host = document.createElement('div')
    host.setAttribute('data-officeview-show', '')
    document.body.append(host)
    const ac = new AbortController()
    const diags: Array<{ code: string; effect?: string; message: string }> = []
    const s = await runShow({
      doc, handle: pptxPackageHandle(doc), host,
      token: { release() {}, released: false, owns: () => false },
      platform: { hasActivation: () => true, requestFullscreen: async () => {}, exitFullscreen: async () => {}, fullscreenElement: () => null, open: () => null },
      fullscreen: Promise.resolve('entered'), retryFullscreen: async () => 'entered',
      from: 'beginning', motion: opts.motion ?? 'auto',
      callbacks: { currentSlide: () => 0, restoreFocus() {}, scrollToSlide() {}, onDiagnostic: (d: { code: string; effect?: string; message: string }) => diags.push(d) },
      signal: ac.signal, generation: 1, isCurrent: () => !ac.signal.aborted, cleanup: () => ac.abort(),
    } as unknown as ShowStart)
    await flush()
    s.tick()
    return { s, diags }
  }

  it('plays p14:flip with its shader on WebGL2, with the matrices of the model', async () => {
    registerGeometric3D()
    const t = transition3DFor(spec('flip'))!
    const draw = vi.spyOn(t, 'draw')
    const { s, diags } = await play()
    expect(s.stage!.backend).toBe('webgl2')
    s.next()
    await flush()
    expect(s.state.slide).toBe(1)
    expect(s.stage!.transition).toMatchObject({ effect: 'flip', duration: 1000 })
    const fake = currentGl!
    const mark = fake.calls.length
    clock.advance(250)
    s.tick()
    expect(draw).toHaveBeenCalledTimes(1)
    const [gl, , , progress, attrs, , mesh] = draw.mock.calls[0]
    expect(gl).toBe(fake.gl)
    expect(progress).toBeCloseTo(0.25, 6)
    expect(attrs).toEqual({ dir: 'r' })
    // The draws the GPU got are the model's: flip dir r at 25 %, under the slide camera at the stage's aspect.
    const want = geometricPieces('flip', progress, { dir: 'r' }, mesh.aspect)
    const drawn = drawsOf(fake, mark)
    const visible = drawn.filter(d => d.mvp.some(x => x !== 0))
    expect(visible.length).toBeGreaterThanOrEqual(2)
    const cam = slideCamera(DEFAULT_FOV_DEG, mesh.aspect)
    const fromDraw = visible[visible.length - 2]
    const fromModel = piece(want, 'from').model
    near(project(fromDraw.mvp, slidePoint(0, 0, mesh.aspect)), project(mul4(cam, fromModel), slidePoint(0, 0, mesh.aspect)), 4)
    expect(mesh.aspect).toBeCloseTo(16 / 9, 2)
    // Mid-way the shader plays; past the end the stage settles on the slide and the shader stops.
    clock.advance(500)
    s.tick()
    expect(draw).toHaveBeenCalledTimes(2)
    clock.advance(400)
    s.tick()
    expect(draw).toHaveBeenCalledTimes(2)
    expect(diags.filter(d => /transition/.test(d.code))).toEqual([])
    s.exit()
  })

  it('on the forced Canvas 2D path the same deck plays the push look-alike, diagnosed', async () => {
    registerGeometric3D()
    const draw = vi.spyOn(transition3DFor(spec('flip'))!, 'draw')
    const { s, diags } = await play({ canvas2d: true })
    expect(s.stage!.backend).toBe('canvas2d')
    s.next()
    await flush()
    clock.advance(250)
    s.tick()
    expect(draw).not.toHaveBeenCalled()
    expect(diags.filter(d => d.code === 'transition-approximate')).toEqual([expect.objectContaining({ effect: 'flip' })])
    expect(diags.find(d => d.code === 'transition-approximate')!.message).toMatch(/push/)
    s.exit()
  })

  it('under reduced motion the deck crossfades and the shader never runs', async () => {
    registerGeometric3D()
    const draw = vi.spyOn(transition3DFor(spec('flip'))!, 'draw')
    const { s, diags } = await play({ motion: 'reduced' })
    expect(s.stage!.backend).toBe('webgl2')
    s.next()
    await flush()
    expect(s.stage!.transition?.effect).toBe('fade')
    expect(s.stage!.transition!.duration).toBeLessThanOrEqual(200)
    clock.advance(100)
    s.tick()
    expect(draw).not.toHaveBeenCalled()
    expect(diags.filter(d => d.code === 'transition-approximate')).toEqual([])
    s.exit()
  })
})
