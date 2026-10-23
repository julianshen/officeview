/**
 * Morph rendering (slideshow plan S3-6, spec 6.3): `morphFrame` composites
 * the matched-object interpolation, fades and statics over the compositor,
 * and the dispatch routes Morph through it on both backends (Canvas 2D:
 * transforms plus crossfade), always diagnosed as an approximation. The
 * session/stage path is covered by tests/slideshow-morph-stage.test.tsx.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { mapTransform, morphFrame, type MorphFrameSpec } from '../src/slideshow/morph/render'
import { dispatchTransition } from '../src/slideshow/render/three/framework'
import { createCanvas2DCompositor } from '../src/slideshow/render/canvas2d'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { openShowPackage, readShowIndex } from '../src/slideshow/read/deck'
import type { BaseGeometry } from '../src/slideshow/engine/behaviours'
import type { Compositor, LayerDraw, LayerHandle, Lease, TargetState, TransitionSpec } from '../src/slideshow/types'

// A 160 × 90 miniature output, viewer-viewport-free (scale 1, no letterbox).
const DEV_W = 160, DEV_H = 90

const geometry = (x: number, y: number, w: number, h: number, rot = 0): BaseGeometry => ({ x, y, w, h, rot })
const plain = (): TargetState => ({ visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0] })

/** A borrowable fake workspace lease, as the frame goldens' harness uses. */
function lease(n = 0): Lease {
  let live = true
  const l: Lease = {
    tag: 'test', get bytes() { return live ? n : 0 },
    grow() { return true }, shrink() {}, transfer: () => l,
    release() { live = false },
  }
  return l
}

/** A compositor stub that records every frame's layer draw list. */
function trackingCompositor(): { draws: LayerDraw[][] } & Compositor {
  const draws: LayerDraw[][] = []
  const stub = {
    backend: 'canvas2d' as const,
    maxTextureSize: 4096,
    upload() { return { refused: 'workspace' as const } },
    release() {},
    frame(list: readonly LayerDraw[]) { draws.push([...list]) },
    resize() {},
    setViewport() {},
    setSlideIndex() {},
    dispose() {},
    prepareTransition3D() { return false },
  }
  return { draws, ...stub } as never as Compositor & { draws: LayerDraw[][] }
}

describe('morphFrame', () => {
  it('a matched rectangle interpolates position, size, rotation and colour linearly at 0.5', () => {
    const comp = trackingCompositor()
    // Position, size and rotation all differ between the two bases.
    const g0 = geometry(10, 20, 100, 50, 10)
    const g1 = geometry(210, 40, 200, 100, 90)
    const spec: MorphFrameSpec = {
      pairs: [{ from: { handle: { id: 1 }, geometry: g0, state: plain() }, to: { handle: { id: 2 }, geometry: g1, state: plain() }, under: true }],
    }
    morphFrame(spec, 0.5, comp)
    const list = comp.draws[0]!
    expect(list).toHaveLength(2)
    expect(list.map(d => d.handle.id)).toEqual([1, 2])
    // The interpolated box at 0.5: linear position and size, shortest-path rotation.
    const mid = geometry(110, 30, 150, 75, 50)
    // Each side's content maps onto the interpolated box: its centre lands on
    // the box centre, its own axes rotate to the box rotation and scale to
    // the box size.
    const expectMap = (d: LayerDraw, base: BaseGeometry): void => {
      const m = d.state.transform
      const at = (x: number, y: number) => [m[0]! * x + m[2]! * y + m[4]!, m[1]! * x + m[3]! * y + m[5]!]
      const bx = base.x + base.w / 2, by = base.y + base.h / 2
      const [cx, cy] = at(bx, by)
      expect(cx).toBeCloseTo(mid.x + mid.w / 2, 9)
      expect(cy).toBeCloseTo(mid.y + mid.h / 2, 9)
      // The base's own half-width axis (rotated by base.rot) maps to the box's.
      const r0 = base.rot * Math.PI / 180, r1 = mid.rot * Math.PI / 180
      const [ex, ey] = at(bx + Math.cos(r0) * base.w / 2, by + Math.sin(r0) * base.w / 2)
      expect(ex - cx).toBeCloseTo(Math.cos(r1) * mid.w / 2, 9)
      expect(ey - cy).toBeCloseTo(Math.sin(r1) * mid.w / 2, 9)
      const [fx, fy] = at(bx - Math.sin(r0) * base.h / 2, by + Math.cos(r0) * base.h / 2)
      expect(fx - cx).toBeCloseTo(-Math.sin(r1) * mid.h / 2, 9)
      expect(fy - cy).toBeCloseTo(Math.cos(r1) * mid.h / 2, 9)
    }
    expectMap(list[0]!, g0)
    expectMap(list[1]!, g1)
    // Same outline: the outgoing side holds at full weight under the incoming at 0.5.
    expect(list[0]!.state.opacity).toBeCloseTo(1, 9)
    expect(list[1]!.state.opacity).toBeCloseTo(0.5, 9)

    // Colour, on real pixels (Canvas 2D): a blue square at g0 matched to a red
    // one at g1 paints the 50/50 colour at the interpolated box's centre,
    // fully opaque (no dip to the background).
    const ws = { limit: 1 << 30, used: () => 0, reserveTerminal: () => undefined }
    const out = createCanvas(DEV_W * 3, DEV_H * 3)
    const c2 = createCanvas2DCompositor(out as never, createCanvas as never, ws as never, createDiagnosticSink())
    c2.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: DEV_W * 3, deviceHeight: DEV_H * 3 })
    const solid = (g: BaseGeometry, rgb: string): LayerHandle => {
      const c = createCanvas(g.w, g.h)
      const x = c.getContext('2d')
      x.fillStyle = rgb
      x.fillRect(0, 0, g.w, g.h)
      const r = c2.upload({ canvas: c as never, bounds: { x: g.x, y: g.y, w: g.w, h: g.h }, scale: 1, bytes: g.w * g.h * 4 }, lease(g.w * g.h * 4))
      if ('refused' in r) throw new Error(r.refused)
      return r
    }
    // Unrotated boxes, so each bitmap is exactly its base box.
    const a0 = geometry(20, 20, 60, 40), a1 = geometry(300, 160, 120, 80)
    const fromH = solid(a0, 'rgb(0,0,255)'), toH = solid(a1, 'rgb(255,0,0)')
    morphFrame({ pairs: [{ from: { handle: fromH, geometry: a0, state: plain() }, to: { handle: toH, geometry: a1, state: plain() }, under: true }] }, 0.5, c2)
    const px = (out as unknown as HTMLCanvasElement).getContext('2d')!.getImageData(Math.round(160 + 45), Math.round(90 + 30), 1, 1).data
    // Box at 0.5: x 160..250, y 90..150 → centre (205, 120).
    expect(Math.abs(px[0]! - 128)).toBeLessThanOrEqual(2)
    expect(px[1]).toBeLessThanOrEqual(2)
    expect(Math.abs(px[2]! - 128)).toBeLessThanOrEqual(2)
    expect(px[3]).toBe(255)
  })

  it('a rotated shape that changes aspect ratio stays a rotated rectangle (no shear)', () => {
    // A square rotated 45° morphs to 200 × 50 at 45°: every frame maps the
    // base's axes onto perpendicular axes.
    const base = geometry(0, 0, 100, 100, 45)
    for (const target of [geometry(0, 0, 200, 50, 45), geometry(50, 25, 150, 75, 45), geometry(10, 10, 160, 40, 20)]) {
      const m = mapTransform(base, target)
      // The base's own (rotated) unit axes.
      const r = base.rot * Math.PI / 180
      const u = [Math.cos(r), Math.sin(r)], v = [-Math.sin(r), Math.cos(r)]
      const mu = [m[0] * u[0]! + m[2] * u[1]!, m[1] * u[0]! + m[3] * u[1]!]
      const mv = [m[0] * v[0]! + m[2] * v[1]!, m[1] * v[0]! + m[3] * v[1]!]
      expect(mu[0]! * mv[0]! + mu[1]! * mv[1]!).toBeCloseTo(0, 9)
      expect(Math.hypot(mu[0]!, mu[1]!)).toBeCloseTo(target.w / base.w, 9)
      expect(Math.hypot(mv[0]!, mv[1]!)).toBeCloseTo(target.h / base.h, 9)
      // And the mapped width axis points along the target rotation.
      expect(Math.atan2(mu[1]!, mu[0]!) * 180 / Math.PI).toBeCloseTo(target.rot, 9)
    }
  })

  it('geometry change crossfades the rendered content', () => {
    const comp = trackingCompositor()
    const g: BaseGeometry = geometry(0, 0, 100, 50)
    const item = {
      from: { handle: { id: 1 }, geometry: g, state: plain() },
      to: { handle: { id: 2 }, geometry: geometry(0, 0, 100, 50, 45), state: plain() },
    }
    morphFrame({ pairs: [item] }, 0.5, comp)
    const list = comp.draws[0]!
    expect(list.map(d => d.handle.id)).toEqual([1, 2])
    expect(list[0]!.state.opacity).toBeCloseTo(0.5, 9)
    expect(list[1]!.state.opacity).toBeCloseTo(0.5, 9)
    // And the frame is empty again at the ends (weights reach 0).
    const comp2 = trackingCompositor()
    morphFrame({ pairs: [item] }, 0, comp2)
    expect(comp2.draws[0]!.map(d => d.handle.id)).toEqual([1])
    const comp3 = trackingCompositor()
    morphFrame({ pairs: [item] }, 1, comp3)
    expect(comp3.draws[0]!.map(d => d.handle.id)).toEqual([2])
  })

  it('tdf149588 morph plays as approximate with a diagnostic', async () => {
    const deck = join(__dirname, '..', 'corpus', 'lo-tdf149588_transparentSolidFill.pptx')
    if (!existsSync(deck)) throw new Error('run bun scripts/fetch-corpus.ts')
    const bytes = readFileSync(deck)
    const pptx = await parsePptx(await OfficePackage.load(bytes))
    const handle = await openShowPackage(bytes)
    const sink = createDiagnosticSink()
    const index = await readShowIndex(pptx, handle, sink)
    // The real metadata reader gives the slide its transition spec.
    const ws = {
      limit: 1 << 26, used: () => 0, reserveTerminal: () => undefined,
      lease: (n: number, tag: string) => {
        let live = true
        const l: Lease = { tag, get bytes() { return live ? n : 0 }, grow() { return true }, shrink() {}, transfer: () => l, release() { live = false } }
        return l
      },
    }
    const { readSlideTiming } = await import('../src/slideshow/read/timing')
    const { shapeResolver } = await import('../src/slideshow/read/resolve')
    const { createSlideMetadataStore } = await import('../src/slideshow/read/store')
    const store = createSlideMetadataStore({ doc: pptx, handle, index, workspace: ws, sink, timing: readSlideTiming, resolver: (slide) => shapeResolver(slide, []) })
    const meta = (await store.acquire(0)).slide
    const spec = meta.transition as TransitionSpec
    expect(meta.transition).toMatchObject({ ns: 'p159', element: 'morph', attrs: { option: 'byObject' } })
    // WebGL2: the morph route, still diagnosed as an approximation.
    const webglSink = createDiagnosticSink()
    const webgl = dispatchTransition(spec, { backend: 'webgl2' } as Compositor, webglSink, 0)
    expect(webgl).toBe('morph')
    expect(webglSink.list().map(d => d.code)).toContain('transition-approximate')
    // Canvas 2D (no GPU): the morph route too (transforms plus crossfade), diagnosed.
    const cpuSink = createDiagnosticSink()
    const cpu = dispatchTransition(spec, { backend: 'canvas2d' } as Compositor, cpuSink, 0)
    expect(cpu).toBe('morph')
    expect(cpuSink.list().map(d => d.code)).toContain('transition-approximate')
    store.close()
  })

  it('canvas2d morph renders without a GPU', () => {
    const ws = { limit: 1 << 30, used: () => 0, reserveTerminal: () => undefined }
    const leaseOf = (n: number): Lease => {
      let live = true
      return {
        tag: 'test', get bytes() { return live ? n : 0 },
        grow() { return true }, shrink() {}, transfer: () => leaseOf(n),
        release() { live = false },
      }
    }
    const out = createCanvas(DEV_W, DEV_H)
    const comp = createCanvas2DCompositor(out as never, createCanvas as never, ws as never, createDiagnosticSink())
    comp.setViewport({ matrix: [1, 0, 0, 1, 0, 0], deviceWidth: DEV_W, deviceHeight: DEV_H })
    const fill = (rgb: [number, number, number]): HTMLCanvasElement => {
      const c = createCanvas(60, 30) as unknown as HTMLCanvasElement
      const x = c.getContext('2d') as unknown as { fillStyle: string | CanvasGradient | CanvasPattern; fillRect: (x: number, y: number, w: number, h: number) => void }
      x.fillStyle = `rgb(${rgb[0] ?? 0},${rgb[1] ?? 0},${rgb[2] ?? 0})`
      x.fillRect(0, 0, 60, 30)
      return c
    }
    const g = geometry(50, 30, 60, 30)
    const upload = (c: HTMLCanvasElement): LayerHandle => {
      const x = createCanvas(1, 1).getContext('2d')
      void x
      const r = comp.upload({ canvas: c, bounds: g, scale: 1, bytes: c.width * c.height * 4 }, lease(c.width * c.height * 4))
      if ('refused' in r) throw new Error(r.refused)
      return r
    }
    const fh = upload(fill([64, 160, 208])), th = upload(fill([208, 96, 64]))
    const render = (progress: number): Uint8ClampedArray => {
      morphFrame({ pairs: [{ from: { handle: fh, geometry: g, state: plain() }, to: { handle: th, geometry: g, state: plain() }, under: true }] }, progress, comp)
      return (out as unknown as HTMLCanvasElement).getContext('2d')!.getImageData(0, 0, DEV_W, DEV_H).data
    }
    const at = (p: number): [number, number, number] => {
      const v = render(p)
      const i = ((DEV_H / 2) * DEV_W + (DEV_W / 2)) * 4
      return [(v[i] ?? 0) / 255, (v[i + 1] ?? 0) / 255, (v[i + 2] ?? 0) / 255]
    }
    // The colour interpolates linearly over an opaque shape (no background bleed).
    for (const p of [0.25, 0.5, 0.75]) {
      const got = at(p)
      const want = [64, 160, 208].map((f, i) => ((1 - p) * f + p * [208, 96, 64][i]!) / 255)
      for (let i = 0; i < 3; i++) expect(Math.abs(got[i]! - want[i]!)).toBeLessThan(2 / 255)
    }
  })
})