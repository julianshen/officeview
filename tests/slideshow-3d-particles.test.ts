/**
 * S3-3: the tile and particle 3-D transitions (vortex, shred, fracture,
 * crush, glitter, honeycomb, ripple; `three/particles.ts`).
 *
 * `flash` and `reveal` are in the S3-3 line of the plan, but the catalogue
 * draws them as exact 2-D rows (stage 1, no look-alike: S1-15) and
 * `registerTransition3D` rejects them, so they are asserted below to stay
 * 2-D and have no shader.
 *
 * Node has no WebGL2: these tests drive the pool and the shaders' draw calls
 * on a recording fake GL (the S3-1 pattern) and pin the pure plan each row
 * hands the GPU, never GLSL pixels. The shaders' own numbers (every cell's
 * time, position and opacity) are checked against the algorithm sheet's
 * reference models (scripts/slideshow-3d-oracles.ts) on a real GPU by
 * tests/browser/slideshow-3d-particles.spec.ts.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { CATALOGUE, rowForTransition } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createRng } from '../src/slideshow/rng'
import { reducedTransition, resolveTransition2D, transitionDraw } from '../src/slideshow/render/transitions2d'
import { GRID_MAX, Transition3DError, createMeshPool, gridBytes } from '../src/slideshow/render/three/mesh'
import { THREE_D_IDS, registerTransition3D, registeredTransitions3D, resetTransitions3D, transition3DFor } from '../src/slideshow/render/three/registry'
import { dispatchTransition } from '../src/slideshow/render/three/framework'
import { PARTICLE_ROWS, particlePlan, registerParticle3D, type ParticlePlan } from '../src/slideshow/render/three/particles'
import { ROWS, tileTime } from '../scripts/slideshow-3d-oracles'
import type { AllocationOwner, Compositor, Lease, TransitionSpec } from '../src/slideshow/types'

interface Pool { used: number; limit: number }
function makeOwner(limit = 256 * 1024 * 1024, log: Array<{ tag: string; bytes: number }> = []): AllocationOwner & { pool: Pool; log: typeof log } {
  const pool: Pool = { used: 0, limit }
  return {
    pool, log,
    lease(bytes, tag): Lease | undefined {
      if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
      pool.used += bytes
      log.push({ tag, bytes })
      let live = true
      const lease: Lease = {
        tag, get bytes() { return live ? bytes : 0 },
        grow: () => false, shrink: () => undefined, transfer: () => lease,
        release() { if (live) { live = false; pool.used -= bytes } },
      }
      return lease
    },
  } as AllocationOwner & { pool: Pool; log: typeof log }
}

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
    getShaderParameter: () => true,
    getProgramParameter: () => true,
    getUniformLocation: (_p: unknown, name: string) => ({ uniform: name }),
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
  const named = (name: string) => calls.filter(c => c.name === name)
  const uniform = (name: string) => calls.filter(c => c.name.startsWith('uniform') && (c.args[0] as { uniform?: string } | null)?.uniform === name)
  return { gl, calls, constant, named, uniform, clear() { calls.length = 0 } }
}

const ASPECT = 16 / 9
const spec = (element: string, attrs: Record<string, string> = {}, ns: TransitionSpec['ns'] = 'p14', preset?: string): TransitionSpec =>
  ({ element, ns, attrs, durationMs: 500, advClick: true, ...(preset !== undefined ? { preset } : {}) })

/** The seven rows this group draws, with the catalogue look-alike each must name. */
const GROUP: Array<{ id: string; spec: TransitionSpec; lookalike: string }> = [
  { id: 'transition:p14:vortex', spec: spec('vortex'), lookalike: 'dissolve' },
  { id: 'transition:p14:shred', spec: spec('shred'), lookalike: 'dissolve' },
  { id: 'transition:p15:prstTrans:fracture', spec: spec('prstTrans', {}, 'p15', 'fracture'), lookalike: 'fade' },
  { id: 'transition:p15:prstTrans:crush', spec: spec('prstTrans', {}, 'p15', 'crush'), lookalike: 'fade' },
  { id: 'transition:p14:glitter', spec: spec('glitter'), lookalike: 'dissolve' },
  { id: 'transition:p14:honeycomb', spec: spec('honeycomb'), lookalike: 'checker' },
  { id: 'transition:p14:ripple', spec: spec('ripple'), lookalike: 'dissolve' },
]
const planOf = (s: TransitionSpec, p = 0.5, seed = 2 ** 31): ParticlePlan => particlePlan(s.element, s.preset, s.attrs, p, ASPECT, seed)

beforeEach(() => { resetTransitions3D() })
afterEach(() => { vi.restoreAllMocks(); resetTransitions3D() })

describe('S3-3 tile and particle 3-D transitions', () => {
  it('every row in the group has a registered shader and a lookalike', () => {
    registerParticle3D()
    registerParticle3D() // idempotent
    expect(PARTICLE_ROWS.map(r => r.preset ?? r.element).sort()).toEqual(['crush', 'fracture', 'glitter', 'honeycomb', 'ripple', 'shred', 'vortex'])
    for (const g of GROUP) {
      expect(THREE_D_IDS, g.id).toContain(g.id)
      const t = transition3DFor(g.spec)
      expect(t, g.id).toBeDefined()
      expect(typeof t!.prepare, g.id).toBe('function')
      const look = resolveTransition2D(g.spec)
      expect(look.approximated, g.id).toBe(true)
      expect(look.spec.element, g.id).toBe(g.lookalike)
      expect(rowForTransition(g.spec.ns, g.spec.element, g.spec.preset)?.lookalike, g.id).toBe(`transition:p:${g.lookalike}`)
    }
    expect(registeredTransitions3D()).toHaveLength(GROUP.length)
    // flash and reveal are exact 2-D rows (S1-15): never 3-D, never registered.
    for (const element of ['flash', 'reveal']) {
      const row = CATALOGUE.find(r => r.id === `transition:p14:${element}`)!
      expect(row.stage, element).toBe(1)
      expect(THREE_D_IDS).not.toContain(row.id)
      expect(() => registerTransition3D({ element, draw: () => undefined })).toThrow(/not a catalogued 3-D/)
      expect(transition3DFor(spec(element))).toBeUndefined()
      expect(resolveTransition2D(spec(element)).approximated).toBe(false)
    }
  })

  it('direction and option attributes change the output', () => {
    const key = (p: ParticlePlan) => JSON.stringify(p, (_k, v) => (v instanceof Float32Array ? Array.from(v) : v))
    const differ = (label: string, a: TransitionSpec, b: TransitionSpec) => expect(key(planOf(a)), label).not.toBe(key(planOf(b)))
    const same = (label: string, a: TransitionSpec, b: TransitionSpec) => expect(key(planOf(a)), label).toBe(key(planOf(b)))
    // vortex dir: four distinct headings; omitted and unknown take `l`.
    const dirs = ['l', 'u', 'r', 'd'].map(dir => spec('vortex', { dir }))
    for (let i = 0; i < 4; i++) for (let j = i + 1; j < 4; j++) differ(`vortex ${i} ${j}`, dirs[i], dirs[j])
    same('vortex default', spec('vortex'), spec('vortex', { dir: 'l' }))
    same('vortex unknown', spec('vortex', { dir: 'x' }), spec('vortex', { dir: 'l' }))
    // shred: pattern changes the grid, dir `in` swaps the frames and runs time backwards.
    const strip = planOf(spec('shred', { pattern: 'strip' })), rect = planOf(spec('shred', { pattern: 'rectangle' }))
    expect(strip.grid).toEqual([24, 1])
    expect(rect.grid).toEqual([24, 14])
    differ('shred in', spec('shred'), spec('shred', { dir: 'in' }))
    const inn = planOf(spec('shred', { dir: 'in' }))
    expect(planOf(spec('shred')).flat).toBe('to')
    expect(inn.flat).toBe('from')
    expect(inn.passes[0].texture).toBe('to')
    expect(planOf(spec('shred')).passes[0].texture).toBe('from')
    expect(inn.passes[0].uniforms.u_p).toBeCloseTo(1 - 0.5, 12)
    expect(planOf(spec('shred', { dir: 'in' }), 0.25).passes[0].uniforms.u_p).toBeCloseTo(0.75, 12)
    // glitter: pattern selects the grid, dir the turning axis and sense.
    expect(planOf(spec('glitter', { pattern: 'diamond' })).grid).toEqual([32, 18])
    expect(planOf(spec('glitter', { pattern: 'hexagon' })).grid).toEqual([24, 14])
    differ('glitter pattern', spec('glitter', { pattern: 'diamond' }), spec('glitter', { pattern: 'hexagon' }))
    const gd = ['l', 'u', 'r', 'd'].map(dir => spec('glitter', { dir }))
    for (let i = 0; i < 4; i++) for (let j = i + 1; j < 4; j++) differ(`glitter ${i} ${j}`, gd[i], gd[j])
    same('glitter default', spec('glitter'), spec('glitter', { pattern: 'diamond', dir: 'l' }))
    // ripple dir: five origins.
    const origins = ['center', 'lu', 'ru', 'ld', 'rd'].map(dir => spec('ripple', { dir }))
    for (let i = 0; i < 5; i++) for (let j = i + 1; j < 5; j++) differ(`ripple ${i} ${j}`, origins[i], origins[j])
    same('ripple default', spec('ripple'), spec('ripple', { dir: 'center' }))
    // The rows without attributes ignore them.
    same('honeycomb', spec('honeycomb'), spec('honeycomb', { dir: 'r' }))
    // The plan follows progress and aspect.
    expect(key(planOf(spec('vortex'), 0.25))).not.toBe(key(planOf(spec('vortex'), 0.75)))
    expect(key(particlePlan('vortex', undefined, {}, 0.5, 1, 1))).not.toBe(key(particlePlan('vortex', undefined, {}, 0.5, 2, 1)))
  })

  it('seeded effects are deterministic', () => {
    const mathRandom = vi.spyOn(Math, 'random')
    registerParticle3D()
    const seedsOf = (g: typeof GROUP[number], seedRng: number, runs = 2) => {
      const out: number[][] = []
      for (let k = 0; k < runs; k++) {
        const fake = fakeGL()
        const owner = makeOwner()
        const pool = createMeshPool(fake.gl, owner)
        pool.setAspect(ASPECT)
        const t = transition3DFor(g.spec)!
        t.prepare!(pool.pool)
        const tex = { tex: 1 } as unknown as WebGLTexture
        t.draw(fake.gl, tex, tex, 0.5, g.spec.attrs, createRng(seedRng), pool.pool)
        out.push(fake.uniform('u_seed').map(c => Number(c.args[1])))
        pool.dispose(true)
        expect(owner.pool.used).toBe(0)
      }
      return out
    }
    for (const g of GROUP) {
      const [a, b] = seedsOf(g, 7)
      expect(a.length, g.id).toBeGreaterThan(0)
      expect(a, g.id).toEqual(b)
      // A different effect rng gives different cell values.
      expect(seedsOf(g, 8, 1)[0], g.id).not.toEqual(a)
    }
    // The plan is a pure function of its inputs.
    const s = spec('glitter', { dir: 'u', pattern: 'hexagon' })
    expect(JSON.stringify(planOf(s, 0.3, 123))).toBe(JSON.stringify(planOf(s, 0.3, 123)))
    expect(mathRandom).not.toHaveBeenCalled()
  })

  it('each shader matches its algorithm-sheet oracle at 50 %', () => {
    // The plan's tile constants reproduce the sheet's rest / done fractions with the sheet's own `tileTime`.
    const sheet = (id: string) => ROWS.find(r => r.id === id)!.oracle(0.5)
    const orderOf = (plan: ParticlePlan) => (uc: number, vc: number) => {
      const u = plan.passes[0].uniforms
      if (u.u_centre === 1) return Math.min(1, Math.hypot((uc - 0.5) * ASPECT, vc - 0.5) / Math.hypot(ASPECT / 2, 0.5))
      const gw = u.u_gw as readonly number[]
      return gw[0] * uc + gw[1] * vc + (u.u_g0 as number)
    }
    const fractions = (plan: ParticlePlan, seed = 2 ** 31) => {
      const [cols, rows] = plan.grid
      const [W, J] = plan.passes[0].uniforms.u_wj as readonly number[]
      const g = orderOf(plan)
      let rest = 0, done = 0
      for (let j = 0; j < rows; j++) for (let i = 0; i < cols; i++) {
        const { t } = tileTime({ cols, rows, W, J, g: (_i, _j, uc, vc) => g(uc, vc) }, seed, i, j, plan.progress)
        if (t === 0) rest++
        if (t === 1) done++
      }
      return { rest: Math.round(rest / (cols * rows) * 1e4) / 1e4, done: Math.round(done / (cols * rows) * 1e4) / 1e4 }
    }
    const table: Array<[string, TransitionSpec, string]> = [
      ['transition:p14:vortex', spec('vortex'), ''],
      ['transition:p14:shred', spec('shred'), ''],
      ['transition:p15:prstTrans:fracture', spec('prstTrans', {}, 'p15', 'fracture'), ''],
      ['transition:p15:prstTrans:crush', spec('prstTrans', {}, 'p15', 'crush'), ''],
      ['transition:p14:glitter', spec('glitter'), ''],
      ['transition:p14:honeycomb', spec('honeycomb'), ''],
    ]
    for (const [id, s] of table) {
      const o = sheet(id)
      const f = fractions(planOf(s))
      expect(f.rest, `${id} rest`).toBeCloseTo(o.rest as number, 4)
      expect(f.done, `${id} done`).toBeCloseTo(o.done as number, 4)
    }
    const rect = fractions(planOf(spec('shred', { pattern: 'rectangle' })))
    expect(rect.rest).toBeCloseTo(sheet('transition:p14:shred')['rectangle rest'] as number, 4)
    expect(rect.done).toBeCloseTo(sheet('transition:p14:shred')['rectangle done'] as number, 4)
    // Literal numbers from the sheet (docs/validation/slideshow-3d-algorithms.md).
    expect(fractions(planOf(spec('vortex')))).toEqual({ rest: 0.059, done: 0.059 })
    expect(fractions(planOf(spec('prstTrans', {}, 'p15', 'fracture')))).toEqual({ rest: 0.0714, done: 0.0238 })
    expect(fractions(planOf(spec('glitter')))).toEqual({ rest: 0.1372, done: 0.1319 })
    expect(fractions(planOf(spec('honeycomb')))).toEqual({ rest: 0.0455, done: 0.0318 })
    // Grids and timing constants, row by row.
    const grid = (s: TransitionSpec) => planOf(s).grid
    expect(grid(spec('vortex'))).toEqual([32, 18])
    expect(grid(spec('prstTrans', {}, 'p15', 'fracture'))).toEqual([12, 7])
    expect(grid(spec('prstTrans', {}, 'p15', 'crush'))).toEqual([16, 9])
    expect(grid(spec('honeycomb'))).toEqual([20, 11])
    expect(grid(spec('ripple'))).toEqual([64, 36])
    expect(planOf(spec('vortex')).passes[0].uniforms.u_wj).toEqual([0.6, 0.35])
    expect(planOf(spec('shred')).passes[0].uniforms.u_wj).toEqual([0.5, 0.2])
    expect(planOf(spec('shred', { pattern: 'rectangle' })).passes[0].uniforms.u_wj).toEqual([0.5, 0.6])
    expect(planOf(spec('prstTrans', {}, 'p15', 'fracture')).passes[0].uniforms.u_wj).toEqual([0.6, 0.4])
    expect(planOf(spec('prstTrans', {}, 'p15', 'crush')).passes[0].uniforms.u_wj).toEqual([0.3, 1])
    expect(planOf(spec('glitter')).passes[0].uniforms.u_wj).toEqual([0.65, 0.4])
    expect(planOf(spec('honeycomb')).passes[0].uniforms.u_wj).toEqual([0.6, 0.3])
    // Glitter: a back pass and a front pass, culled, and the back pass reads the incoming frame.
    const glitter = planOf(spec('glitter'))
    expect(glitter.flat).toBeUndefined()
    expect(glitter.passes.map(p => [p.texture, p.cull])).toEqual([['to', 'front'], ['from', 'back']])
    // Ripple: the incoming frame flat, the outgoing one displaced with opacity 1 - e(p); the sheet's z at 50 %.
    const ripple = planOf(spec('ripple'))
    expect(ripple.flat).toBe('to')
    expect(ripple.passes[0].texture).toBe('from')
    expect(ripple.passes[0].uniforms.u_p).toBe(0.5)
    expect(sheet('transition:p14:ripple')['from opacity']).toBe(0.5)
  })

  it('leases every grid before creating it, and keeps particle counts bounded', () => {
    registerParticle3D()
    for (const g of GROUP) {
      const t = transition3DFor(g.spec)!
      const fake = fakeGL()
      const owner = makeOwner()
      const pool = createMeshPool(fake.gl, owner)
      t.prepare!(pool.pool)
      expect(owner.log.length, g.id).toBeGreaterThan(0)
      expect(owner.log.every(l => l.tag === 'webgl2:mesh'), g.id).toBe(true)
      // The first GL object is created only after its lease.
      expect(owner.log.length, g.id).toBe(fake.named('createVertexArray').length)
      expect(fake.named('createProgram').length, g.id).toBeGreaterThan(0)
      pool.dispose(true)
      expect(owner.pool.used, g.id).toBe(0)
      // A refused lease throws holding nothing, and creates no GL object.
      const refuse = fakeGL()
      const small = makeOwner(1)
      const p2 = createMeshPool(refuse.gl, small)
      expect(() => t.prepare!(p2.pool), g.id).toThrow(Transition3DError)
      expect(refuse.named('createVertexArray'), g.id).toHaveLength(0)
      expect(small.pool.used).toBe(0)
    }
    // Every plan a row can make stays within the mesh bound and a small cell count.
    for (const g of GROUP) {
      for (const attrs of [{}, { pattern: 'rectangle', dir: 'in' }, { pattern: 'hexagon', dir: 'u' }, { dir: 'rd' }] as Array<Record<string, string>>) {
        const plan = particlePlan(g.spec.element, g.spec.preset, attrs, 0.5, ASPECT, 1)
        const [c, r] = plan.grid
        expect(c, g.id).toBeLessThanOrEqual(GRID_MAX)
        expect(r, g.id).toBeLessThanOrEqual(GRID_MAX)
        expect(c * r, g.id).toBeLessThanOrEqual(64 * 36)
        expect(gridBytes(c, r), g.id).toBeLessThan(256 * 1024)
      }
    }
    // The glitter and shred variants are all leased up front, since prepare sees no attributes.
    const fake = fakeGL()
    const owner = makeOwner()
    const pool = createMeshPool(fake.gl, owner)
    transition3DFor(spec('glitter'))!.prepare!(pool.pool)
    expect(owner.log.map(l => l.bytes).sort()).toEqual([gridBytes(24, 14), gridBytes(32, 18)].sort())
  })

  it('without WebGL2 each row draws its look-alike with a diagnostic, and reduced motion never reaches a shader', () => {
    registerParticle3D()
    const canvas2d = { backend: 'canvas2d' } as unknown as Compositor
    for (const g of GROUP) {
      const sink = createDiagnosticSink()
      expect(dispatchTransition(g.spec, canvas2d, sink, 3), g.id).toBe('lookalike')
      expect(sink.list(), g.id).toEqual([expect.objectContaining({ code: 'transition-approximate', slideIndex: 3 })])
      // The look-alike is what transitionDraw draws.
      expect(transitionDraw(g.spec, 0.5, createRng(1), ASPECT)).toBeDefined()
      // Reduced motion: a fade of at most 200 ms, which is no 3-D row.
      const reduced = reducedTransition({ ...g.spec, durationMs: 4000 })
      expect(transition3DFor(reduced), g.id).toBeUndefined()
      expect(reduced.durationMs, g.id).toBeLessThanOrEqual(200)
    }
  })

  it('draws the row with the expected passes on the fake GL', () => {
    registerParticle3D()
    const run = (s: TransitionSpec, p = 0.5) => {
      const fake = fakeGL()
      const pool = createMeshPool(fake.gl, makeOwner())
      pool.setAspect(ASPECT)
      const t = transition3DFor(s)!
      t.prepare!(pool.pool)
      fake.clear()
      const tex = { tex: 1 } as unknown as WebGLTexture
      t.draw(fake.gl, tex, tex, p, s.attrs, createRng(3), pool.pool)
      return { draws: fake.named('drawElements'), fake }
    }
    // A flat backdrop plus one tile draw (vortex ... honeycomb, ripple); glitter draws two culled passes.
    for (const g of GROUP) {
      const { draws, fake } = run(g.spec)
      expect(draws.length, g.id).toBe(2)
      if (g.id.endsWith('glitter')) expect(fake.named('cullFace').map(c => c.args[0])).toEqual([fake.constant('FRONT'), fake.constant('BACK')])
      expect(fake.named('depthFunc'), g.id).toHaveLength(0)
    }
  })
})
