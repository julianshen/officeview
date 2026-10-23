/**
 * S3-7, the Stage 3 close: every 3-D row is wired (`three/index.ts`), the
 * show's backend registers them before its WebGL2 compositor exists, and the
 * mesh pool caches every grid all three groups use, so playing them back to
 * back never rebuilds a grid it already built (carry item 1).
 *
 * Node has no WebGL2: the pool runs on a recording fake GL (the S3-1
 * pattern), and a grid build is one `createVertexArray` call.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { CATALOGUE } from '../src/slideshow/catalogue'
import { createRng } from '../src/slideshow/rng'
import { MAX_CACHED_GRIDS, createMeshPool, gridBytes } from '../src/slideshow/render/three/mesh'
import { THREE_D_IDS, registeredTransitions3D, resetTransitions3D, transition3DFor, type Transition3D } from '../src/slideshow/render/three/registry'
import { registerTransitions3D } from '../src/slideshow/render/three'
import { GEOMETRIC_ELEMENTS } from '../src/slideshow/render/three/geometric'
import { PARTICLE_ROWS } from '../src/slideshow/render/three/particles'
import { DEFORM_ROWS } from '../src/slideshow/render/three/deform'
import { dispatchTransition } from '../src/slideshow/render/three/framework'
import type { AllocationOwner, Compositor, Lease, TransitionSpec } from '../src/slideshow/types'

function makeOwner(log: Array<{ tag: string; bytes: number }> = []): AllocationOwner & { used(): number } {
  let used = 0
  return {
    used: () => used,
    lease(bytes, tag): Lease | undefined {
      used += bytes
      log.push({ tag, bytes })
      let live = true
      const lease: Lease = {
        tag, get bytes() { return live ? bytes : 0 },
        grow: () => false, shrink: () => undefined, transfer: () => lease,
        release() { if (live) { live = false; used -= bytes } },
      }
      return lease
    },
  } as AllocationOwner & { used(): number }
}

function fakeGL() {
  const counts = new Map<string, number>()
  const constants = new Map<string, number>()
  let serial = 0
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
      if (/^[A-Z][A-Z0-9_]*$/.test(prop)) {
        let v = constants.get(prop)
        if (v === undefined) { v = 0x1000 + constants.size; constants.set(prop, v) }
        return v
      }
      return () => {
        counts.set(prop, (counts.get(prop) ?? 0) + 1)
        return prop.startsWith('create') ? { kind: prop.slice(6), serial: ++serial } : undefined
      }
    },
  }) as WebGL2RenderingContext
  return { gl, count: (name: string) => counts.get(name) ?? 0 }
}

const idOf = (t: Pick<Transition3D, 'element' | 'preset'>) =>
  CATALOGUE.find(r => r.kind === 'transition' && r.element === t.element && (r.id.split(':')[3] ?? undefined) === t.preset && r.stage === 3)?.id

/** Every attribute variant that changes a row's grid (glitter's and shred's patterns). */
const VARIANTS: Readonly<Record<string, ReadonlyArray<Record<string, string>>>> = {
  glitter: [{}, { pattern: 'hexagon' }],
  shred: [{}, { pattern: 'rectangle' }],
}

beforeEach(() => { resetTransitions3D() })
afterEach(() => { resetTransitions3D() })

describe('S3-7 Stage 3 close: 3-D wiring', () => {
  it('registerTransitions3D registers a shader for every 3-D catalogue row, idempotently', () => {
    expect(registeredTransitions3D()).toEqual([])
    registerTransitions3D()
    registerTransitions3D()
    const ids = registeredTransitions3D().map(idOf)
    expect(ids.every(id => id !== undefined)).toBe(true)
    expect([...ids].sort()).toEqual([...THREE_D_IDS].sort())
    expect(ids).toHaveLength(GEOMETRIC_ELEMENTS.length + PARTICLE_ROWS.length + DEFORM_ROWS.length)
    expect(THREE_D_IDS).toHaveLength(26)
  })

  it('with the groups wired, every 3-D row dispatches to its shader on WebGL2 and to its look-alike on Canvas 2D', () => {
    registerTransitions3D()
    const gl = { backend: 'webgl2', prepareTransition3D: () => true } as unknown as Compositor
    const c2d = { backend: 'canvas2d' } as unknown as Compositor
    for (const row of CATALOGUE.filter(r => r.kind === 'transition' && r.stage === 3)) {
      const [, ns, element, preset] = row.id.split(':')
      const spec: TransitionSpec = { ns: ns as TransitionSpec['ns'], element, attrs: {}, durationMs: 500, advClick: true, ...(preset !== undefined ? { preset } : {}) }
      if (ns === 'p159') {
        expect(dispatchTransition(spec, gl), row.id).toBe('morph')
        expect(dispatchTransition(spec, c2d), row.id).toBe('morph')
        continue
      }
      expect(transition3DFor(spec), row.id).toBeDefined()
      expect(dispatchTransition(spec, gl), row.id).toBe('shader')
      expect(dispatchTransition(spec, c2d), row.id).toBe('lookalike')
    }
  })

  it('the show backend registers the 3-D rows before it creates its WebGL2 compositor', async () => {
    const { createBackendManager } = await import('../src/slideshow/render/backend')
    const { createSlideShowWorkspace } = await import('../src/slideshow/workspace')
    const { createDiagnosticSink } = await import('../src/slideshow/diagnostics')
    expect(registeredTransitions3D()).toEqual([])
    const host = document.createElement('div')
    const output = document.createElement('canvas')
    output.width = 320; output.height = 180
    document.body.append(output, host)
    const workspace = createSlideShowWorkspace('desktop')
    const terminal = workspace.reserveTerminal(320 * 180 * 4)!
    // jsdom has no WebGL2, so the manager starts on Canvas 2D; the rows are registered all the same.
    const m = createBackendManager(host, {
      now: () => 0, schedule: () => 0, paintLive: () => {}, rebuildLayers: async () => {},
      workspace, terminal, output, sink: createDiagnosticSink(),
    })
    expect(registeredTransitions3D().map(idOf).sort()).toEqual([...THREE_D_IDS].sort())
    m.dispose()
    terminal.release()
    document.body.replaceChildren()
  })
})

describe('S3-7 Stage 3 close: the grid cache holds every group (carry item 1)', () => {
  /** Plays every registered transition at a few progresses (each attribute variant), as a slide change would. */
  function sweep(pool: ReturnType<typeof createMeshPool>, gl: WebGL2RenderingContext, only?: ReadonlyArray<Transition3D>) {
    const tex = { tex: 1 } as unknown as WebGLTexture
    for (const t of only ?? registeredTransitions3D()) {
      for (const attrs of VARIANTS[t.preset ?? t.element] ?? [{}]) {
        t.prepare?.(pool.pool)
        for (const p of [0.25, 0.5, 0.75]) t.draw(gl, tex, tex, p, attrs, createRng(7), pool.pool)
      }
    }
  }

  it('one transition from each group back to back rebuilds no cached grid', () => {
    registerTransitions3D()
    const fake = fakeGL()
    const owner = makeOwner()
    const pool = createMeshPool(fake.gl, owner)
    pool.setAspect(16 / 9)
    const pick = (element: string, preset?: string) => transition3DFor({ ns: preset ? 'p15' : 'p14', element, ...(preset ? { preset } : {}) })!
    const three = [pick('doors'), pick('ripple'), pick('prstTrans', 'curtains')]
    sweep(pool, fake.gl, three)
    const built = fake.count('createVertexArray')
    sweep(pool, fake.gl, three)
    sweep(pool, fake.gl, [...three].reverse())
    expect(fake.count('createVertexArray')).toBe(built)
    expect(fake.count('deleteVertexArray')).toBe(0)
    pool.dispose(true)
    expect(owner.used()).toBe(0)
  })

  it('every grid of every registered row fits the cache: a second sweep of all rows builds nothing', () => {
    registerTransitions3D()
    const fake = fakeGL()
    const log: Array<{ tag: string; bytes: number }> = []
    const owner = makeOwner(log)
    const pool = createMeshPool(fake.gl, owner)
    pool.setAspect(16 / 9)
    sweep(pool, fake.gl)
    const built = fake.count('createVertexArray')
    // Each distinct grid is built once, and none was evicted on the way.
    expect(fake.count('deleteVertexArray')).toBe(0)
    expect(built).toBeLessThanOrEqual(MAX_CACHED_GRIDS)
    sweep(pool, fake.gl)
    sweep(pool, fake.gl, [...registeredTransitions3D()].reverse())
    expect(fake.count('createVertexArray')).toBe(built)
    expect(fake.count('deleteVertexArray')).toBe(0)
    // Every grid is leased against the workspace (webgl2:mesh), once. The whole cache is what the
    // validation doc counts: 9 grids today, 291,308 B.
    const mesh = log.filter(l => l.tag === 'webgl2:mesh')
    expect(mesh).toHaveLength(built)
    expect(pool.bytes()).toBe(mesh.reduce((s, l) => s + l.bytes, 0))
    expect(built).toBe(9)
    expect(pool.bytes()).toBe([[1, 1], [32, 18], [8, 18], [64, 36], [24, 1], [24, 14], [12, 7], [16, 9], [20, 11]].reduce((s, [c, r]) => s + gridBytes(c, r), 0))
    expect(pool.bytes()).toBe(291_308)
    pool.dispose(true)
    expect(owner.used()).toBe(0)
  })
})
