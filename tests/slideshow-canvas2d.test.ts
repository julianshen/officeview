/**
 * S1-13: the Canvas 2D compositor and the mask library (slideshow spec 6.3,
 * 7). Rendered through node-canvas, which is injected through the
 * `CanvasFactory`; every oracle is computed by hand in this file.
 *
 * The workspace is a test ledger typed by the S1-2 contracts (S1-16 owns the
 * real one), the same shape S1-12's tests use.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { createRng } from '../src/slideshow/rng'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { MASK_KINDS, isMaskRaster, maskPath, maskRasterBytes, tracePath, type MaskRaster } from '../src/slideshow/render/masks'
import { CANVAS2D_MAX_TEXTURE_SIZE, applyColorMatrix, createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import type {
  Affine, CanvasFactory, Compositor, LayerBitmap, LayerHandle, Lease, MaskKind, MaskState, SlideShowWorkspace, TargetState,
} from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Workspace and factory
// ---------------------------------------------------------------------------
interface Pool { used: number; limit: number }
function poolLease(pool: Pool, bytes: number, tag: string, log?: string[]): Lease | undefined {
  if (bytes < 0 || pool.used + bytes > pool.limit) return undefined
  pool.used += bytes
  log?.push(`lease:${tag}`)
  let held = bytes, live = true
  const lease: Lease = {
    get tag() { return tag },
    get bytes() { return live ? held : 0 },
    grow(extra) {
      if (!live || pool.used + extra > pool.limit) { log?.push(`grow-refused:${tag}`); return false }
      pool.used += extra; held += extra; log?.push(`grow:${tag}`); return true
    },
    shrink(less) { const n = Math.min(less, held); held -= n; pool.used -= n },
    transfer() { return lease },
    release() { if (!live) return; live = false; pool.used -= held },
  }
  return lease
}
function makeWorkspace(limit: number, log?: string[]): SlideShowWorkspace & { pool: Pool } {
  const pool: Pool = { used: 0, limit }
  return {
    pool, limit,
    lease: (bytes, tag) => poolLease(pool, bytes, tag, log),
    used: () => pool.used,
    reserveTerminal: () => undefined,
  }
}

type Ctx = CanvasRenderingContext2D
let filterHits = 0
/** Every context the tests hand out throws if anyone reads or writes `filter`. */
function guardFilter(c: ReturnType<typeof createCanvas>) {
  const ctx = c.getContext('2d')
  Object.defineProperty(ctx, 'filter', {
    configurable: true,
    get() { filterHits++; throw new Error('ctx.filter read') },
    set() { filterHits++; throw new Error('ctx.filter written') },
  })
  return c
}
function makeFactory(log?: string[]) {
  const calls: Array<[number, number]> = []
  const factory: CanvasFactory = (w, h) => {
    calls.push([w, h]); log?.push(`canvas:${w}x${h}`)
    return guardFilter(createCanvas(w, h)) as unknown as HTMLCanvasElement
  }
  return { factory, calls }
}

const W = 64, H = 48
const RED: Px = [255, 0, 0, 255], BLUE: Px = [0, 0, 255, 255], GREEN: Px = [0, 255, 0, 255], CLEAR: Px = [0, 0, 0, 0]
type Px = [number, number, number, number]

function setup(limit = 64 * 1024 * 1024, w = W, h = H) {
  const log: string[] = []
  const ws = makeWorkspace(limit, log)
  const { factory, calls } = makeFactory(log)
  const out = guardFilter(createCanvas(w, h))
  const sink = createDiagnosticSink()
  const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws, sink)
  const ctx = out.getContext('2d') as unknown as Ctx
  /** A solid layer in a test-owned canvas, charged to its own lease, then uploaded. */
  const layer = (color: string, bounds: LayerBitmap['bounds'], scale = 1): LayerHandle => {
    const cw = Math.ceil(bounds.w * scale), ch = Math.ceil(bounds.h * scale)
    const bytes = cw * ch * 4
    const lease = ws.lease(bytes, 'layer')!
    const c = createCanvas(cw, ch)
    const x = c.getContext('2d'); x.fillStyle = color; x.fillRect(0, 0, cw, ch)
    const r = comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds, scale, bytes }, lease)
    if ('refused' in r) throw new Error(`refused ${r.refused}`)
    return r
  }
  const px = (x: number, y: number): Px => [...ctx.getImageData(x, y, 1, 1).data] as Px
  const all = () => ctx.getImageData(0, 0, w, h).data
  return { ws, log, factory, calls, out, ctx, comp, sink, layer, px, all }
}

const I: Affine = [1, 0, 0, 1, 0, 0]
const state = (over: Partial<TargetState> = {}): TargetState => ({ visible: true, opacity: 1, transform: I, ...over })
const near = (a: Px, b: Px, tol = 1) => a.every((v, i) => Math.abs(v - b[i]) <= tol)
function expectPx(actual: Px, expected: Px, tol = 1) {
  if (!near(actual, expected, tol)) expect(actual).toEqual(expected)
}

/** Full-slide outgoing (red) and incoming (blue) frames, composited through a mask. */
function maskedTransition(kind: MaskKind, progress: number, params: MaskState['params'] = {}, seed = 7) {
  const s = setup()
  const from = s.layer('#ff0000', { x: 0, y: 0, w: W, h: H })
  const to = s.layer('#0000ff', { x: 0, y: 0, w: W, h: H })
  s.comp.frame([], { from, to, effect: kind, progress, params: { mask: { kind, progress, params } }, rng: createRng(seed) })
  return s
}

// ---------------------------------------------------------------------------
// Geometric oracles at progress 0.5 (default params, 64 × 48 box, 1 px pad)
// ---------------------------------------------------------------------------
const cx = W / 2, cy = H / 2
const ORACLE: Partial<Record<MaskKind, (x: number, y: number) => boolean>> = {
  wipe: (x) => x < 32,
  split: (_x, y) => Math.abs(y - cy) < 12,
  barn: (_x, y) => Math.abs(y - cy) < 12,
  box: (x, y) => Math.abs(x - cx) < 16.5 && Math.abs(y - cy) < 12.5,
  circle: (x, y) => Math.hypot(x - cx, y - cy) < 0.5 * (Math.hypot(W, H) / 2 + 1),
  diamond: (x, y) => Math.abs(x - cx) / 33 + Math.abs(y - cy) / 25 < 1,
  plus: (x, y) => {
    const dx = Math.abs(x - cx), dy = Math.abs(y - cy)
    return (dx < 16.5 && dy < 25) || (dx < 33 && dy < 12.5)
  },
  wedge: (_x, y) => y < cy,
  wheel: (x) => x > cx,
  checkerboard: (x, y) => {
    const r = Math.floor(y / 8), o = (r % 2) * 4
    return ((((x - o) % 8) + 8) % 8) < 4
  },
  blinds: (_x, y) => (y % 8) < 4,
  comb: (x, y) => (Math.floor(y / 6) % 2 === 0 ? x < 32 : x >= 32),
  strips: (x, y) => x / W + y / H < 1,
}
/** Raster kinds: the incoming layer shows where the stretched raster cell is opaque. */
function rasterOracle(r: MaskRaster) {
  return (x: number, y: number) => {
    const col = Math.min(r.width - 1, Math.floor((x * r.width) / W)), row = Math.min(r.height - 1, Math.floor((y * r.height) / H))
    return r.alpha[row * r.width + col] === 255
  }
}

describe('mask library', () => {
  it('MASK_KINDS lists every MaskKind once', () => {
    const all: Record<MaskKind, true> = {
      wipe: true, split: true, circle: true, diamond: true, plus: true, wedge: true, wheel: true, barn: true,
      checkerboard: true, dissolve: true, blinds: true, box: true, comb: true, randomBar: true, strips: true,
    }
    expect([...MASK_KINDS].sort()).toEqual(Object.keys(all).sort())
  })

  it('noise kinds are rasters and every other kind is a path; rasters fit maskRasterBytes', () => {
    for (const kind of MASK_KINDS) {
      const m = maskPath(kind, 0.5, {}, W, H, createRng(1))
      if (kind === 'dissolve' || kind === 'randomBar') {
        expect(isMaskRaster(m)).toBe(true)
        const r = m as MaskRaster
        expect(r.alpha.length).toBe(r.width * r.height)
        expect(maskRasterBytes(kind, {}, W, H)).toBeGreaterThanOrEqual(r.alpha.byteLength)
      } else {
        expect(isMaskRaster(m)).toBe(false)
        expect(maskRasterBytes(kind, {}, W, H)).toBe(0)
      }
    }
  })

  it('a path traces onto any CanvasPath-shaped target', () => {
    const calls: string[] = []
    const target = {
      moveTo: (x: number, y: number) => calls.push(`M${x},${y}`),
      lineTo: (x: number, y: number) => calls.push(`L${x},${y}`),
      arc: () => calls.push('A'),
      rect: (x: number, y: number, w: number, h: number) => calls.push(`R${x},${y},${w},${h}`),
      closePath: () => calls.push('Z'),
    }
    const m = maskPath('wipe', 0.25, { from: 'l' }, 100, 50, createRng(1))
    if (isMaskRaster(m)) throw new Error('expected a path')
    tracePath(target, m, 10, 20)
    expect(calls).toEqual(['R10,20,25,50'])
  })

  it('direction params change the geometry: wipe from r, split vert in, wheel 4 spokes', () => {
    const r = maskedTransition('wipe', 0.25, { from: 'r' })
    expectPx(r.px(47, 10), RED); expectPx(r.px(48, 10), BLUE)
    const t = maskedTransition('wipe', 0.5, { from: 't' })
    expectPx(t.px(10, 23), BLUE); expectPx(t.px(10, 24), RED)
    const s = maskedTransition('split', 0.5, { orient: 'vert', dir: 'in' })
    expectPx(s.px(15, 5), BLUE); expectPx(s.px(16, 5), RED); expectPx(s.px(47, 5), RED); expectPx(s.px(48, 5), BLUE)
    const w = maskedTransition('wheel', 0.5, { spokes: 4 })
    // Each quarter is half swept, clockwise from its own start.
    expectPx(w.px(cx + 10, cy - 20), BLUE); expectPx(w.px(cx + 20, cy - 10), RED)
    const c = maskedTransition('circle', 0.5, { dir: 'in' })
    expectPx(c.px(cx, cy), RED); expectPx(c.px(1, 1), BLUE)
  })
})

describe('Canvas 2D compositor', () => {
  it('opacity, transform and z-order of three layers match a hand-composited reference', () => {
    const s = setup()
    const a = s.layer('#ff0000', { x: 0, y: 0, w: 32, h: 32 })
    // Rasterised at half scale (a reduced-scale layer): composited at its full bounds.
    const b = s.layer('#0000ff', { x: 0, y: 0, w: 32, h: 32 }, 0.5)
    const c = s.layer('#00ff00', { x: 0, y: 0, w: 10, h: 20 })
    s.comp.frame([
      { handle: a, state: state() },
      { handle: b, state: state({ opacity: 0.5, transform: [1, 0, 0, 1, 16, 16] }) },
      // 90° clockwise about the origin, then moved: the 10×20 layer covers x 30–50, y 4–14.
      { handle: c, state: state({ transform: [0, 1, -1, 0, 50, 4] }) },
    ])
    expectPx(s.px(5, 5), RED)
    expectPx(s.px(20, 20), [128, 0, 127, 255], 1) // blue at 0.5 over red
    expectPx(s.px(40, 40), [0, 0, 255, 128], 2) // blue at 0.5 over nothing (getImageData un-premultiplies)
    expectPx(s.px(40, 8), GREEN) // green on top of blue
    expectPx(s.px(31, 10), GREEN) // green on top of red
    expectPx(s.px(55, 8), CLEAR)
    expectPx(s.px(10, 40), CLEAR)

    // Reversed z-order: blue at 0.5 now under opaque red.
    s.comp.frame([{ handle: b, state: state({ opacity: 0.5, transform: [1, 0, 0, 1, 16, 16] }) }, { handle: a, state: state() }])
    expectPx(s.px(20, 20), RED)
    // Invisible and zero-opacity draws paint nothing; motion is a translation in output px after the transform.
    s.comp.frame([
      { handle: a, state: state({ visible: false }) },
      { handle: c, state: state({ opacity: 0 }) },
      { handle: c, state: state({ motion: { dx: 40, dy: 10 }, transform: [1, 0, 0, 1, 2, 2] }) },
    ])
    expectPx(s.px(5, 5), CLEAR)
    expectPx(s.px(42, 12), GREEN); expectPx(s.px(41, 12), CLEAR); expectPx(s.px(51, 31), GREEN); expectPx(s.px(52, 31), CLEAR)
  })

  it('a layer mask is applied in the layer box under its transform', () => {
    const s = setup()
    const a = s.layer('#ff0000', { x: 4, y: 4, w: 20, h: 10 })
    s.comp.frame([{ handle: a, state: state({ transform: [1, 0, 0, 1, 10, 6], mask: { kind: 'wipe', progress: 0.5, params: { from: 'l' } } }) }])
    // Layer box x 14–34, y 10–20; the left half x 14–24 shows.
    expectPx(s.px(14, 12), RED); expectPx(s.px(23, 19), RED); expectPx(s.px(24, 12), CLEAR); expectPx(s.px(13, 12), CLEAR)
    // A raster mask on a layer, with a seed from params, is stable frame to frame.
    const d = { handle: a, state: state({ mask: { kind: 'dissolve' as const, progress: 0.5, params: { seed: 3, cell: 2 } } }) }
    s.comp.frame([d]); const one = [...s.all()]
    s.comp.frame([d]); expect([...s.all()]).toEqual(one)
    const shown = one.filter((_, i) => i % 4 === 3 && one[i] === 255).length
    expect(shown).toBe(100) // exactly half of the 10 × 5 cells of 2 × 2 px
  })

  it('transition affines are in unit slide space and fit the layer to the output', () => {
    const s = setup()
    const from = s.layer('#ff0000', { x: 0, y: 0, w: 32, h: 24 }, 2) // a half-size slide at scale 2
    const to = s.layer('#0000ff', { x: 0, y: 0, w: 32, h: 24 }, 2)
    // push from the right at 0.25: from moves left by a quarter, to follows it in.
    s.comp.frame([], { from, to, effect: 'push', progress: 0.25, params: { from: [1, 0, 0, 1, -0.25, 0], to: [1, 0, 0, 1, 0.75, 0] }, rng: createRng(1) })
    expectPx(s.px(47, 20), RED); expectPx(s.px(48, 20), BLUE); expectPx(s.px(63, 47), BLUE)
    s.comp.frame([], { from, to, effect: 'fade', progress: 0.5, params: { toOpacity: 0.5 }, rng: createRng(1) })
    expectPx(s.px(10, 10), [128, 0, 127, 255])
    // draws composite after the transition (the watermark stays on top).
    const mark = s.layer('#00ff00', { x: 0, y: 0, w: 4, h: 4 })
    s.comp.frame([{ handle: mark, state: state() }], { from, to, effect: 'cut', progress: 1, params: {}, rng: createRng(1) })
    expectPx(s.px(1, 1), GREEN); expectPx(s.px(10, 10), BLUE)
  })

  it('each mask kind at progress 0 hides and 1 shows the incoming layer fully', () => {
    for (const kind of MASK_KINDS) {
      for (const [progress, color] of [[0, RED], [1, BLUE]] as const) {
        const s = maskedTransition(kind, progress)
        const data = s.all()
        for (let i = 0; i < data.length; i += 4) {
          const p: Px = [data[i], data[i + 1], data[i + 2], data[i + 3]]
          if (!near(p, color, 0)) throw new Error(`${kind} at ${progress}: pixel ${(i / 4) % W},${Math.floor(i / 4 / W)} is ${p}`)
        }
      }
    }
  })

  it.each(MASK_KINDS)('%s at progress 0.5 matches its oracle away from the edges', (kind) => {
    // The compositor draws one seed from the transition's rng and re-seeds every frame from it.
    const seed = Math.floor(createRng(7)() * 0x100000000)
    const raster = maskPath(kind, 0.5, {}, W, H, createRng(seed))
    const inside = isMaskRaster(raster) ? rasterOracle(raster) : ORACLE[kind]!
    expect(inside).toBeDefined()
    const s = maskedTransition(kind, 0.5)
    let checked = 0, blue = 0
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const at = (dx: number, dy: number) => inside(x + 0.5 + dx, y + 0.5 + dy)
      const want = at(0, 0)
      let stable = true
      for (const dx of [-1.5, 0, 1.5]) for (const dy of [-1.5, 0, 1.5]) if (at(dx, dy) !== want) stable = false
      if (!stable) continue
      checked++
      const p = s.px(x, y)
      if (!near(p, want ? BLUE : RED, 0)) throw new Error(`${kind}: pixel ${x},${y} is ${p}, oracle says ${want ? 'incoming' : 'outgoing'}`)
      if (want) blue++
    }
    expect(checked).toBeGreaterThan((W * H) / 10)
    expect(blue).toBeGreaterThan(0)
    expect(blue).toBeLessThan(checked)
  })

  it('wipe left at 0.5 reveals exactly the left half', () => {
    const s = maskedTransition('wipe', 0.5, { from: 'l' })
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const p = s.px(x, y)
      if (!near(p, x < W / 2 ? BLUE : RED, 0)) throw new Error(`pixel ${x},${y} is ${p}`)
    }
  })

  it('dissolve with a fixed seed is identical across runs', () => {
    const one = maskedTransition('dissolve', 0.3, {}, 42)
    const two = maskedTransition('dissolve', 0.3, {}, 42)
    const a = [...one.all()]
    expect([...two.all()]).toEqual(a)
    expect([...maskedTransition('dissolve', 0.3, {}, 43).all()]).not.toEqual(a)

    // The transition's rng is stateful; later frames of one transition keep the same pattern.
    const s = setup()
    const from = s.layer('#ff0000', { x: 0, y: 0, w: W, h: H }), to = s.layer('#0000ff', { x: 0, y: 0, w: W, h: H })
    const rng = createRng(42)
    const draw = (progress: number) => ({ from, to, effect: 'dissolve', progress, params: { mask: { kind: 'dissolve' as const, progress, params: {} } }, rng })
    s.comp.frame([], draw(0.3)); expect([...s.all()]).toEqual(a)
    s.comp.frame([], draw(0.6)); s.comp.frame([], draw(0.3)); expect([...s.all()]).toEqual(a)
    // Monotone: every cell shown at 0.3 is still shown at 0.6.
    const r3 = maskPath('dissolve', 0.3, {}, W, H, createRng(9)) as MaskRaster
    const r6 = maskPath('dissolve', 0.6, {}, W, H, createRng(9)) as MaskRaster
    r3.alpha.forEach((v, i) => { if (v === 255) expect(r6.alpha[i]).toBe(255) })
    // Exactly half the cells at 0.5.
    const r5 = maskPath('dissolve', 0.5, {}, W, H, createRng(9)) as MaskRaster
    expect(r5.alpha.filter(v => v === 255).length).toBe(r5.alpha.length / 2)
  })

  it('colour matrix on a premultiplied semi-transparent pixel returns the expected sRGB', () => {
    const c = createCanvas(1, 1)
    const x = c.getContext('2d')
    x.fillStyle = 'rgba(255, 0, 0, 0.5)'; x.fillRect(0, 0, 1, 1)
    const src = x.getImageData(0, 0, 1, 1) as unknown as ImageData
    // R' = 0.5 R, G' = 0.2 (an offset, normalised), B' = B, A' = A.
    const m = [0.5, 0, 0, 0, 0, 0, 0, 0, 0, 0.2, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0]
    const out = applyColorMatrix(src, m)
    expect(out).not.toBe(src)
    expect([...src.data]).toEqual([255, 0, 0, 127]) // input untouched
    // Un-premultiplied: half of 255, not half of the stored 127.
    expect([...out.data]).toEqual([128, 51, 0, 127])
    x.putImageData(out as never, 0, 0)
    expectPx([...x.getImageData(0, 0, 1, 1).data] as Px, [128, 51, 0, 127], 2)
    // Clamps to [0, 255].
    expect([...applyColorMatrix(src, [2, 0, 0, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2]).data]).toEqual([255, 0, 0, 255])
  })

  it('a part layer gets its part colour matrix; a whole layer ignores colour', () => {
    const s = setup()
    const a = s.layer('rgba(255, 0, 0, 0.5)', { x: 0, y: 0, w: 8, h: 8 })
    const gray = [0.2126, 0.7152, 0.0722, 0, 0, 0.2126, 0.7152, 0.0722, 0, 0, 0.2126, 0.7152, 0.0722, 0, 0, 0, 0, 0, 1, 0]
    s.comp.frame([{ handle: a, part: 'fill', state: state({ color: { fill: gray, text: [] } }) }])
    expectPx(s.px(2, 2), [54, 54, 54, 127], 2)
    s.comp.frame([{ handle: a, part: 'line', state: state({ color: { fill: gray } }) }])
    expectPx(s.px(2, 2), [255, 0, 0, 127], 2)
    s.comp.frame([{ handle: a, state: state({ color: { fill: gray } }) }])
    expectPx(s.px(2, 2), [255, 0, 0, 127], 2)
  })

  it('the compositor never touches ctx.filter', () => {
    filterHits = 0
    const gray = [0.3, 0.6, 0.1, 0, 0, 0.3, 0.6, 0.1, 0, 0, 0.3, 0.6, 0.1, 0, 0, 0, 0, 0, 1, 0]
    for (const kind of MASK_KINDS) {
      const s = maskedTransition(kind, 0.4)
      const a = s.layer('#ff0000', { x: 0, y: 0, w: 16, h: 16 })
      s.comp.frame([{ handle: a, part: 'fill', state: state({ opacity: 0.5, color: { fill: gray }, mask: { kind, progress: 0.4, params: {} } }) }])
      s.comp.resize(32, 32)
      s.comp.frame([{ handle: a, state: state() }])
      s.comp.dispose()
    }
    expect(filterHits).toBe(0)
  })

  it('upload with an exhausted lease returns refused and creates no canvas', () => {
    const layerBytes = 16 * 16 * 4
    const ws = makeWorkspace(layerBytes)
    const { factory, calls } = makeFactory()
    const out = createCanvas(W, H)
    const comp = createCanvas2DCompositor(out as unknown as HTMLCanvasElement, factory, ws)
    const lease = ws.lease(layerBytes, 'layer')!
    const before = calls.length
    const bitmap: LayerBitmap = { canvas: createCanvas(16, 16) as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: 16, h: 16 }, scale: 1, bytes: layerBytes }
    expect(comp.upload(bitmap, lease)).toEqual({ refused: 'workspace' })
    expect(calls.length).toBe(before)
    expect(lease.bytes).toBe(layerBytes) // untouched: still the caller's
    expect(ws.used()).toBe(layerBytes)
  })

  it('upload leases before creating, charges the texture after upload, and release/dispose return every byte', () => {
    const s = setup(1024 * 1024)
    const bytes = 16 * 8 * 4
    const lease = s.ws.lease(bytes, 'layer')!
    s.log.length = 0
    const c = createCanvas(16, 8)
    const r = s.comp.upload({ canvas: c as unknown as HTMLCanvasElement, bounds: { x: 0, y: 0, w: 16, h: 8 }, scale: 1, bytes }, lease)
    expect('refused' in r).toBe(false)
    expect(s.log).toEqual(['grow:layer', 'canvas:16x8'])
    expect(lease.bytes).toBe(bytes) // the CPU copy's bytes left; the texture's stayed
    s.comp.release(r as LayerHandle)
    expect(lease.bytes).toBe(0)
    expect(s.ws.used()).toBe(0)

    // Too big for a texture: refused before any charge.
    const big = CANVAS2D_MAX_TEXTURE_SIZE + 1
    const l2 = s.ws.lease(4, 'big')!
    const fake = { width: big, height: 1 } as unknown as HTMLCanvasElement
    expect(s.comp.upload({ canvas: fake, bounds: { x: 0, y: 0, w: big, h: 1 }, scale: 1, bytes: 4 }, l2)).toEqual({ refused: 'texture-size' })
    expect(l2.bytes).toBe(4)
    l2.release()

    // Scratch and mask surfaces are leased before they are created, then all released on dispose.
    const a = s.layer('#ff0000', { x: 0, y: 0, w: 16, h: 16 })
    s.log.length = 0
    s.comp.frame([{ handle: a, part: 'fill', state: state({ color: { fill: [0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0] }, mask: { kind: 'dissolve', progress: 0.5, params: {} } }) }])
    for (let i = 0; i < s.log.length; i++) {
      if (s.log[i].startsWith('canvas:')) expect(s.log[i - 1]).toMatch(/^(lease|grow):/)
    }
    expect(s.log.some(e => e.startsWith('canvas:'))).toBe(true)
    expect(s.ws.used()).toBeGreaterThan(16 * 16 * 4)
    s.comp.dispose()
    expect(s.ws.used()).toBe(0)
    s.comp.dispose() // idempotent
    expect(s.ws.used()).toBe(0)
  })

  it('a refused scratch lease degrades a raster mask to a diagnosed cut', () => {
    // Half-scale frames of L/4 bytes each: the second upload peaks at 3L/4 (both copies held), then L/2 stays
    // charged, leaving L/4, too little for the output-sized scratch surface (L).
    const L = W * H * 4
    const s = setup((3 * L) / 4)
    const from = s.layer('#ff0000', { x: 0, y: 0, w: W, h: H }, 0.5), to = s.layer('#0000ff', { x: 0, y: 0, w: W, h: H }, 0.5)
    const before = s.calls.length
    const draw = (progress: number) => ({ from, to, effect: 'dissolve', progress, params: { mask: { kind: 'dissolve' as const, progress, params: {} } }, rng: createRng(1) })
    s.comp.frame([], draw(0.4))
    expectPx(s.px(5, 5), RED)
    s.comp.frame([], draw(0.6))
    expectPx(s.px(5, 5), BLUE)
    expect(s.calls.length).toBe(before)
    const diags = s.sink.list()
    expect(diags.length).toBe(1)
    expect(diags[0].code).toBe('compositor-workspace')
  })

  it('an unknown mask kind shows the layer unmasked with a diagnostic', () => {
    const s = setup()
    const a = s.layer('#ff0000', { x: 0, y: 0, w: 8, h: 8 })
    s.comp.frame([{ handle: a, state: state({ mask: { kind: 'zigzag' as MaskKind, progress: 0.1, params: {} } }) }])
    expectPx(s.px(4, 4), RED)
    expect(s.sink.list().map(d => d.code)).toEqual(['mask-unknown'])
  })

  it('reports its backend and maximum texture size', () => {
    const s = setup()
    const c: Compositor = s.comp
    expect(c.backend).toBe('canvas2d')
    expect(c.maxTextureSize).toBe(CANVAS2D_MAX_TEXTURE_SIZE)
    s.comp.resize(20, 10)
    expect(s.out.width).toBe(20)
    expect(s.out.height).toBe(10)
  })
})
