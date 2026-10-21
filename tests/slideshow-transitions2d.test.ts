/**
 * S1-15: the 2-D slide transitions (slideshow spec 6.3 transitions, 4.5
 * reduced motion, 10 seeded randomness).
 *
 * `transitionDraw` turns a `TransitionSpec` into compositor params at a
 * progress; both compositors draw them. Canvas 2D pixels come from
 * node-canvas; the WebGL2 compositor runs on a recording fake GL (Node has no
 * WebGL2), so it is checked for draw calls and diagnostics here and for
 * pixels (parity with Canvas 2D at the row's tolerance) in Chromium by
 * tests/browser/slideshow-gpu.spec.ts.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { CATALOGUE } from '../src/slideshow/catalogue'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createRng } from '../src/slideshow/rng'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { createWebGL2Compositor } from '../src/slideshow/render/webgl2'
import {
  REDUCED_TRANSITION_MS, TRANSITION_2D_IDS, randomPick, reducedTransition, transitionDraw,
} from '../src/slideshow/render/transitions2d'
import type { CanvasFactory, LayerHandle, Lease, Rng, SlideShowWorkspace, TransitionSpec } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace, factory and a recording fake WebGL2 context
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
function makeWorkspace(limit = 64 * 1024 * 1024): SlideShowWorkspace {
  const pool: Pool = { used: 0, limit }
  return { limit, lease: (b, t) => poolLease(pool, b, t), used: () => pool.used, reserveTerminal: () => undefined }
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
    getParameter: (p: number) => (p === constant('MAX_TEXTURE_SIZE') ? 4096 : 0),
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

let currentGl: ReturnType<typeof fakeGL> | null = null
let realGetContext: typeof HTMLCanvasElement.prototype.getContext
beforeEach(() => {
  realGetContext = HTMLCanvasElement.prototype.getContext
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement, type: string, ...rest: unknown[]) {
    if (type === 'webgl2') return (currentGl?.gl ?? null) as never
    return (realGetContext as (...a: unknown[]) => unknown).call(this, type, ...rest) as never
  } as never)
})
afterEach(() => { vi.restoreAllMocks(); currentGl = null })

// ---------------------------------------------------------------------------
// Rendering helpers
// ---------------------------------------------------------------------------
const W = 64, H = 48
type Px = [number, number, number, number]

function spec(element: string, attrs: Record<string, string> = {}, ns: TransitionSpec['ns'] = 'p'): TransitionSpec {
  return { element, ns, attrs, durationMs: 500, advClick: true }
}
/** Every 2-D catalogue row as a spec. */
const ROW_SPECS: Array<{ id: string; spec: TransitionSpec }> = CATALOGUE
  .filter(r => r.kind === 'transition' && r.stage === 1)
  .map(r => ({ id: r.id, spec: spec(r.element!, {}, r.ns as TransitionSpec['ns']) }))
/** The rows plus the attribute variants the brief names. */
const ALL_SPECS = [
  ...ROW_SPECS,
  { id: 'cut thruBlk', spec: spec('cut', { thruBlk: '1' }) },
  { id: 'fade thruBlk', spec: spec('fade', { thruBlk: 'true' }) },
  { id: 'zoom in', spec: spec('zoom', { dir: 'in' }) },
  { id: 'push d', spec: spec('push', { dir: 'd' }) },
  { id: 'cover ru', spec: spec('cover', { dir: 'ru' }) },
  { id: 'pull ld', spec: spec('pull', { dir: 'ld' }) },
  { id: 'flythrough out', spec: spec('flythrough', { dir: 'out' }, 'p14') },
  { id: 'reveal thruBlk r', spec: spec('reveal', { thruBlk: '1', dir: 'r' }, 'p14') },
]

/** Red outgoing, blue incoming full-slide frames on the Canvas 2D compositor. */
function canvasFrame(s: TransitionSpec, progress: number, rng: Rng = createRng(7)) {
  const ws = makeWorkspace()
  const out = createCanvas(W, H)
  const sink = createDiagnosticSink()
  const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws, sink)
  const layer = (color: string): LayerHandle => {
    const c = createCanvas(W, H)
    const x = c.getContext('2d'); x.fillStyle = color; x.fillRect(0, 0, W, H)
    const r = comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: W, h: H }, scale: 1, bytes: W * H * 4 }, ws.lease(W * H * 4, 'layer')!)
    if ('refused' in r) throw new Error(r.refused)
    return r
  }
  const from = layer('#ff0000'), to = layer('#0000ff')
  const params = transitionDraw(s, progress, rng, W / H)
  comp.frame([], { from, to, effect: s.element, progress, params, rng })
  const data = out.getContext('2d').getImageData(0, 0, W, H).data
  const px = (x: number, y: number): Px => [...data.slice((y * W + x) * 4, (y * W + x) * 4 + 4)] as Px
  comp.dispose()
  return { px, data, sink, params, leaked: ws.used() }
}

function webglFrame(s: TransitionSpec, progress: number) {
  currentGl = fakeGL()
  const ws = makeWorkspace()
  const sink = createDiagnosticSink()
  const canvas = document.createElement('canvas')
  canvas.width = W; canvas.height = H
  const comp = createWebGL2Compositor(canvas, ws, { sink, factory })!
  const layer = (): LayerHandle => {
    const c = document.createElement('canvas'); c.width = W; c.height = H
    const r = comp.upload({ canvas: c, bounds: { x: 0, y: 0, w: W, h: H }, scale: 1, bytes: W * H * 4 }, ws.lease(W * H * 4, 'layer')!)
    if ('refused' in r) throw new Error(r.refused)
    return r
  }
  const from = layer(), to = layer()
  const rng = createRng(7)
  const before = currentGl.calls.length
  comp.frame([], { from, to, effect: s.element, progress, params: transitionDraw(s, progress, rng, W / H), rng })
  const calls = currentGl.calls.slice(before)
  comp.dispose()
  return { calls, sink, gl: currentGl, leaked: ws.used() }
}

const near = (a: Px, b: Px, tol = 2) => a.every((v, i) => Math.abs(v - b[i]) <= tol)
const RED: Px = [255, 0, 0, 255], BLUE: Px = [0, 0, 255, 255], BLACK: Px = [0, 0, 0, 255]
function countWhere(data: Uint8ClampedArray, want: Px) {
  let n = 0
  for (let i = 0; i < data.length; i += 4) if (near([data[i], data[i + 1], data[i + 2], data[i + 3]], want)) n++
  return n
}

// ---------------------------------------------------------------------------
// The brief's tests
// ---------------------------------------------------------------------------
describe('transitionDraw', () => {
  it('every 2-D catalogue row has a transitionDraw and both compositors render it', () => {
    // The 2-D set is the catalogue's stage-1 transition rows, not a list kept beside it.
    expect([...TRANSITION_2D_IDS].sort()).toEqual(ROW_SPECS.map(r => r.id).sort())
    expect(ROW_SPECS.length).toBe(26)
    for (const { id, spec: s } of ALL_SPECS) {
      for (const p of [0.25, 0.5, 0.75]) {
        const c2 = canvasFrame(s, p)
        expect(c2.sink.list(), `${id} @${p} canvas2d diagnostics`).toEqual([])
        expect(c2.leaked, `${id} canvas2d leak`).toBe(0)
        // The slide is painted edge to edge: no transition leaves a hole mid-way.
        let opaque = 0
        for (let i = 3; i < c2.data.length; i += 4) if (c2.data[i] >= 254) opaque++
        expect(opaque, `${id} @${p} opaque`).toBe(W * H)
        const gl = webglFrame(s, p)
        expect(gl.sink.list(), `${id} @${p} webgl2 diagnostics`).toEqual([])
        expect(gl.leaked, `${id} webgl2 leak`).toBe(0)
        // One quad per visible frame, and a backdrop clear beyond the frame's own clear when asked for.
        const params = transitionDraw(s, p, createRng(7), W / H)
        const visible = ((params.fromOpacity ?? 1) > 0 ? 1 : 0) + ((params.toOpacity ?? 1) > 0 ? 1 : 0)
        expect(gl.calls.filter(c => c.name === 'drawArrays').length, `${id} @${p} webgl2 quads`).toBe(visible)
        expect(gl.calls.filter(c => c.name === 'clear').length, `${id} @${p} webgl2 clears`).toBe(params.backdrop ? 2 : 1)
        expect(visible + (params.backdrop ? 1 : 0), `${id} @${p} paints something`).toBeGreaterThan(0)
      }
    }
  })

  it('progress 0 shows the outgoing and 1 the incoming slide for every row', () => {
    for (const { id, spec: s } of ALL_SPECS) {
      const start = canvasFrame(s, 0)
      expect(countWhere(start.data, RED), `${id} @0 is the outgoing slide`).toBe(W * H)
      const end = canvasFrame(s, 1)
      expect(countWhere(end.data, BLUE), `${id} @1 is the incoming slide`).toBe(W * H)
    }
  })

  it('direction attributes change the geometry: push l vs r, split horz in vs vert out, wipe d', () => {
    // push dir is the direction of motion: "l" moves left, so the incoming slide enters on the right.
    const pushL = canvasFrame(spec('push', { dir: 'l' }), 0.5)
    expect(near(pushL.px(4, 24), RED)).toBe(true)
    expect(near(pushL.px(60, 24), BLUE)).toBe(true)
    const pushR = canvasFrame(spec('push', { dir: 'r' }), 0.5)
    expect(near(pushR.px(4, 24), BLUE)).toBe(true)
    expect(near(pushR.px(60, 24), RED)).toBe(true)
    expect(pushL.params.from).not.toEqual(pushR.params.from)
    // The default push is dir="l".
    expect(transitionDraw(spec('push'), 0.5, createRng(1))).toEqual(transitionDraw(spec('push', { dir: 'l' }), 0.5, createRng(1)))

    // split horz in: closing from the top and bottom edges toward a horizontal seam.
    const horzIn = canvasFrame(spec('split', { orient: 'horz', dir: 'in' }), 0.5)
    expect(near(horzIn.px(32, 2), BLUE)).toBe(true)
    expect(near(horzIn.px(32, 45), BLUE)).toBe(true)
    expect(near(horzIn.px(32, 24), RED)).toBe(true)
    expect(near(horzIn.px(2, 24), RED)).toBe(true)
    // split vert out: opening from a vertical centre seam outward.
    const vertOut = canvasFrame(spec('split', { orient: 'vert', dir: 'out' }), 0.5)
    expect(near(vertOut.px(32, 2), BLUE)).toBe(true)
    expect(near(vertOut.px(32, 45), BLUE)).toBe(true)
    expect(near(vertOut.px(2, 24), RED)).toBe(true)
    expect(near(vertOut.px(61, 24), RED)).toBe(true)

    // wipe d: moving down, so the reveal starts at the top edge.
    const wipeD = canvasFrame(spec('wipe', { dir: 'd' }), 0.5)
    expect(near(wipeD.px(32, 4), BLUE)).toBe(true)
    expect(near(wipeD.px(32, 44), RED)).toBe(true)
    // The default wipe is dir="l": the reveal starts at the right edge.
    const wipe = canvasFrame(spec('wipe'), 0.5)
    expect(near(wipe.px(60, 24), BLUE)).toBe(true)
    expect(near(wipe.px(4, 24), RED)).toBe(true)
  })

  it('random picks deterministically per seed', () => {
    const picks = new Set<string>()
    for (let seed = 1; seed <= 40; seed++) {
      const a = randomPick(createRng(seed)), b = randomPick(createRng(seed))
      expect(a).toEqual(b)
      expect(['random', 'cut', 'none']).not.toContain(a.element)
      expect(a.ns).toBe('p')
      picks.add(a.element)
      // One rng picks once: every frame of the effect draws the same transition.
      const rng = createRng(seed)
      const first = randomPick(rng)
      for (const p of [0, 0.3, 0.7, 1]) {
        expect(randomPick(rng)).toEqual(first)
        expect(transitionDraw(spec('random'), p, rng, W / H)).toEqual(transitionDraw(first, p, createRng(seed), W / H))
      }
    }
    expect(picks.size).toBeGreaterThan(5)
  })

  it('reduced motion replaces any transition with a fade of at most 200 ms', () => {
    expect(REDUCED_TRANSITION_MS).toBe(200)
    const inputs: TransitionSpec[] = [
      ...ALL_SPECS.map(s => s.spec),
      { ...spec('vortex', {}, 'p14'), durationMs: 3000 },
      { ...spec('prstTrans', { prst: 'curtains' }, 'p15'), preset: 'curtains', durationMs: 2000 },
      { element: 'push', ns: 'p', attrs: { dir: 'u' }, durationMs: 150, advClick: false, advTmMs: 4000, sound: { end: true } },
    ]
    // cut and none are instant already and stay so (fix round 1); every real transition becomes the fade.
    for (const s of inputs.filter(x => x.element !== 'cut' && x.element !== 'none')) {
      const r = reducedTransition(s)
      expect(r.element).toBe('fade')
      expect(r.ns).toBe('p')
      expect(r.preset).toBeUndefined()
      expect(r.attrs).toEqual({})
      expect(r.durationMs).toBeLessThanOrEqual(200)
      expect(r.durationMs).toBe(Math.min(s.durationMs, 200))
      // The advance is not motion: it is kept.
      expect(r.advClick).toBe(s.advClick)
      expect(r.advTmMs).toBe(s.advTmMs)
      // A plain crossfade: no mask, no movement.
      const d = transitionDraw(r, 0.5, createRng(1))
      expect(d.mask).toBeUndefined()
      expect(d.from).toBeUndefined()
      expect(d.to).toBeUndefined()
      expect(d.toOpacity).toBeCloseTo(0.5)
    }
  })

  it('cut thruBlk paints black at 0.5', () => {
    const c2 = canvasFrame(spec('cut', { thruBlk: '1' }), 0.5)
    expect(countWhere(c2.data, BLACK)).toBe(W * H)
    // A plain cut has no black frame: it is the incoming slide from the midpoint.
    expect(countWhere(canvasFrame(spec('cut'), 0.5).data, BLUE)).toBe(W * H)
    expect(countWhere(canvasFrame(spec('cut'), 0.49).data, RED)).toBe(W * H)
    // WebGL2 clears the slide's device rect to opaque black and draws no frame.
    const gl = webglFrame(spec('cut', { thruBlk: '1' }), 0.5)
    const clear = gl.calls.findIndex(c => c.name === 'clearColor' && c.args.join() === '0,0,0,1')
    expect(clear).toBeGreaterThanOrEqual(0)
    expect(gl.calls.find(c => c.name === 'scissor')?.args).toEqual([0, 0, W, H])
    expect(gl.calls.filter(c => c.name === 'drawArrays')).toHaveLength(0)
  })
})

describe('mask mapping details', () => {
  it('wheel sweeps clockwise and p14 wheelReverse counter-clockwise from 12 o\'clock', () => {
    const cw = canvasFrame(spec('wheel', { spokes: '1' }), 0.25)
    expect(near(cw.px(44, 12), BLUE)).toBe(true)
    expect(near(cw.px(20, 12), RED)).toBe(true)
    const ccw = canvasFrame(spec('wheelReverse', { spokes: '1' }, 'p14'), 0.25)
    expect(near(ccw.px(20, 12), BLUE)).toBe(true)
    expect(near(ccw.px(44, 12), RED)).toBe(true)
  })

  it('strips dir is the direction of motion: lu starts at the lower-right corner', () => {
    const s = canvasFrame(spec('strips', { dir: 'lu' }), 0.3)
    expect(near(s.px(62, 46), BLUE)).toBe(true)
    expect(near(s.px(1, 1), RED)).toBe(true)
  })

  it('a catalogued stand-in draws its 2-D look-alike and an unknown element a fade', () => {
    const vortex = transitionDraw(spec('vortex', {}, 'p14'), 0.4, createRng(1))
    expect(vortex.mask?.kind).toBe('dissolve')
    expect(transitionDraw(spec('nonsense'), 0.4, createRng(1))).toEqual({ toOpacity: 0.4 })
    expect(transitionDraw({ ...spec('prstTrans', {}, 'p15'), preset: 'curtains' }, 0.4, createRng(1)).mask?.kind).toBe('split')
  })
})

describe('S1-15 fix round 1', () => {
  it('p14:pan without dir is PowerPoint\'s Pan Left (moves left, the incoming slide enters on the right)', () => {
    const pan = spec('pan', {}, 'p14')
    expect(transitionDraw(pan, 0.5, createRng(1))).toEqual(transitionDraw(spec('pan', { dir: 'l' }, 'p14'), 0.5, createRng(1)))
    const f = canvasFrame(pan, 0.5)
    expect(near(f.px(4, 24), RED)).toBe(true)
    expect(near(f.px(60, 24), BLUE)).toBe(true)
  })

  it('reduced motion keeps cut and none instant and never adds motion', () => {
    for (const s of [spec('cut'), spec('cut', { thruBlk: '1' }), spec('none')]) {
      const r = reducedTransition(s)
      expect(r, `${s.element} ${JSON.stringify(s.attrs)}`).toEqual(s)
      // No intermediate crossfade frame: the effect only switches.
      for (const p of [0.25, 0.4, 0.6, 0.75]) {
        const d = transitionDraw(r, p, createRng(1))
        expect([0, 1]).toContain(d.toOpacity)
      }
    }
    // A real transition still becomes the short fade.
    expect(reducedTransition(spec('push')).element).toBe('fade')
  })

  it('the p14:flash catalogue row says how it is drawn: a transform (fade through white), not a mask', () => {
    const row = CATALOGUE.find(r => r.id === 'transition:p14:flash')!
    expect(row.implementation).toBe('transform')
    expect(transitionDraw(spec('flash', {}, 'p14'), 0.5, createRng(1)).mask).toBeUndefined()
  })
})

describe('compositor transition params (additive for S1-15)', () => {
  it('fromOnTop draws the outgoing frame over the incoming one on both backends', () => {
    // pull (uncover) moves the outgoing slide away over the incoming one.
    const c2 = canvasFrame(spec('pull', { dir: 'l' }), 0.25)
    expect(near(c2.px(4, 24), RED)).toBe(true)
    expect(near(c2.px(60, 24), BLUE)).toBe(true)
    const gl = webglFrame(spec('pull', { dir: 'l' }), 0.25)
    const draws = gl.calls.filter(c => c.name === 'drawArrays' || c.name === 'uniformMatrix3fv')
    // Two quads; the second (top) one is the translated outgoing frame.
    const matrices = draws.filter(c => c.name === 'uniformMatrix3fv').map(c => [...(c.args[2] as Float32Array)])
    expect(matrices).toHaveLength(2)
    // The incoming frame (identity) is drawn first; the outgoing one, moved left by a quarter, on top.
    expect(matrices[0][6]).toBeCloseTo(0)
    expect(matrices[1][6]).toBeCloseTo(-W * 0.25)
  })

  it('a transition is clipped to the slide rect and never spills into the letterbox', () => {
    // A 64 × 48 slide letterboxed by 6 px top and bottom in a 64 × 60 output.
    const view = { matrix: [1, 0, 0, 1, 0, 6] as const, deviceWidth: W, deviceHeight: H + 12 }
    for (const s of [spec('push', { dir: 'u' }), spec('zoom'), spec('flythrough', {}, 'p14'), spec('cover', { dir: 'd' })]) {
      const ws = makeWorkspace()
      const out = createCanvas(W, H + 12)
      const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws)
      comp.setViewport(view)
      const up = (color: string) => {
        const c = createCanvas(W, H); const x = c.getContext('2d'); x.fillStyle = color; x.fillRect(0, 0, W, H)
        const r = comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: W, h: H }, scale: 1, bytes: W * H * 4 }, ws.lease(W * H * 4, 'layer')!)
        if ('refused' in r) throw new Error(r.refused)
        return r
      }
      const from = up('#ff0000'), to = up('#0000ff')
      for (const p of [0, 0.3, 0.6]) {
        const rng = createRng(1)
        comp.frame([], { from, to, effect: s.element, progress: p, params: transitionDraw(s, p, rng, W / H), rng })
        const ctx = out.getContext('2d')
        const top = ctx.getImageData(0, 0, W, 6).data, bottom = ctx.getImageData(0, H + 6, W, 6).data
        expect([...top, ...bottom].every(v => v === 0), `${s.element} @${p} letterbox`).toBe(true)
      }
      comp.dispose()
    }
    // WebGL2 scissors the transition to the slide's device rect (GL origin bottom-left) and turns it off after.
    currentGl = fakeGL()
    const ws = makeWorkspace()
    const canvas = document.createElement('canvas')
    const comp = createWebGL2Compositor(canvas, ws, { factory })!
    comp.setViewport(view)
    const c = document.createElement('canvas'); c.width = W; c.height = H
    const h1 = comp.upload({ canvas: c, bounds: { x: 0, y: 0, w: W, h: H }, scale: 1, bytes: 0 }, ws.lease(0, 'a')!)
    const h2 = comp.upload({ canvas: c, bounds: { x: 0, y: 0, w: W, h: H }, scale: 1, bytes: 0 }, ws.lease(0, 'b')!)
    if ('refused' in h1 || 'refused' in h2) throw new Error('refused')
    const before = currentGl.calls.length
    const rng = createRng(1)
    comp.frame([], { from: h1, to: h2, effect: 'push', progress: 0.3, params: transitionDraw(spec('push', { dir: 'u' }), 0.3, rng), rng })
    const calls = currentGl.calls.slice(before)
    expect(calls.find(x => x.name === 'scissor')?.args).toEqual([0, 6, W, H])
    const names = calls.map(x => x.name)
    expect(names.lastIndexOf('disable')).toBeGreaterThan(names.lastIndexOf('drawArrays'))
    comp.dispose()
  })

  it('a backdrop is drawn under both frames and fade thruBlk passes through it', () => {
    const mid = canvasFrame(spec('fade', { thruBlk: '1' }), 0.5)
    expect(countWhere(mid.data, BLACK)).toBe(W * H)
    const quarter = canvasFrame(spec('fade', { thruBlk: '1' }), 0.25)
    const [r, g, b, a] = quarter.px(32, 24)
    expect(a).toBe(255)
    expect(Math.abs(r - 128)).toBeLessThanOrEqual(2)
    expect(g + b).toBeLessThanOrEqual(2)
    const flash = canvasFrame(spec('flash', {}, 'p14'), 0.5)
    expect(countWhere(flash.data, [255, 255, 255, 255])).toBe(W * H)
  })
})
