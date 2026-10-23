/**
 * Morph rendering (slideshow plan S3-6, spec 6.3): `morphFrame` composites
 * the matched-object interpolation, fades and statics over the compositor,
 * and the dispatch routes Morph through it on WebGL2 while Canvas 2D keeps
 * the catalogued fade look-alike with a diagnostic.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { morphFrame, type MorphFrameSpec } from '../src/slideshow/morph/render'
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

/** mul, for test-side expectation matrices (same affine convention as the module: x' = a x + c y + e). */
function mul(m: readonly number[], n: readonly number[]): number[] {
  return [
    m[0]! * n[0]! + m[2]! * n[2]!, m[1]! * n[0]! + m[3]! * n[2]!,
    m[0]! * n[1]! + m[2]! * n[3]!, m[1]! * n[1]! + m[3]! * n[3]!,
    m[0]! * n[4]! + m[2]! * n[5]! + m[4]!, m[1]! * n[4]! + m[3]! * n[5]! + m[5]!,
  ]
}

describe('morphFrame', () => {
  it('a matched rectangle interpolates position, size, rotation and colour linearly at 0.5', () => {
    const comp = trackingCompositor()
    // Position, size and rotation all differ between the two bases.
    const g0 = geometry(10, 20, 100, 50, 10)
    const g1 = geometry(210, 40, 200, 100, 90)
    const spec: MorphFrameSpec = {
      pairs: [{ from: { handle: { id: 1 }, geometry: g0, state: plain() }, to: { handle: { id: 2 }, geometry: g1, state: plain() } }],
    }
    morphFrame(spec, 0.5, comp)
    const list = comp.draws[0]!
    expect(list).toHaveLength(2)
    expect(list.map(d => d.handle.id)).toEqual([1, 2])
    // The interpolated box at 0.5: linear position and size, shortest-path rotation.
    const mid = geometry(110, 30, 150, 75, 50)
    // Mapping a layer's content (base box + its own rotation baked in) onto
    // the interpolated box: centre-preserving scale, then the remaining
    // rotation about the interpolated centre.
    const expectMap = (d: LayerDraw, base: BaseGeometry, opacity: number): void => {
      const sx = mid.w / base.w, sy = mid.h / base.h
      const cx = mid.x + mid.w / 2, cy = mid.y + mid.h / 2
      const bx = base.x + base.w / 2, by = base.y + base.h / 2
      const S = [sx, 0, 0, sy, cx - sx * bx, cy - sy * by]
      const delta = (mid.rot - base.rot) * Math.PI / 180
      const c = Math.cos(delta), s = Math.sin(delta)
      const R = [c, s, -s, c, cx - c * cx + s * cy, cy - s * cx - c * cy]
      expect(d.state.opacity).toBeCloseTo(opacity, 9)
      for (let i = 0; i < 6; i++) expect(d.state.transform[i]).toBeCloseTo(mul(R, S)[i]!, 9)
    }
    // Colour rides the cross-fade weights: both sides at 0.5. The outgoing
    // layer draws first; the incoming layer on top.
    expect(list[0]!.state.opacity).toBeCloseTo(0.5, 9)
    expect(list[1]!.state.opacity).toBeCloseTo(0.5, 9)
    expectMap(list[0]!, g0, 0.5)
    expectMap(list[1]!, g1, 0.5)
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
    console.log('[tdf149588] index slides', JSON.stringify(index.slides), 'diags', sink.list().map(d => d.code))
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
    // Canvas 2D (no GPU): the catalogued fade look-alike, diagnosed too.
    const cpuSink = createDiagnosticSink()
    const cpu = dispatchTransition(spec, { backend: 'canvas2d' } as Compositor, cpuSink, 0)
    expect(cpu).toBe('lookalike')
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
      morphFrame({ pairs: [{ from: { handle: fh, geometry: g, state: plain() }, to: { handle: th, geometry: g, state: plain() } }] }, progress, comp)
      return (out as unknown as HTMLCanvasElement).getContext('2d')!.getImageData(0, 0, DEV_W, DEV_H).data
    }
    const at = (p: number): [number, number, number] => {
      const v = render(p)
      const i = ((DEV_H / 2) * DEV_W + (DEV_W / 2)) * 4
      return [(v[i] ?? 0) / 255, (v[i + 1] ?? 0) / 255, (v[i + 2] ?? 0) / 255]
    }
    // The over-composited mix straightened: a half out, then half in over it.
    const expectMix = (f: [number, number, number], t: [number, number, number]) =>
      [(0.25 * f[0]! + 0.5 * t[0]!) / 0.75, (0.25 * f[1]! + 0.5 * t[1]!) / 0.75, (0.25 * f[2]! + 0.5 * t[2]!) / 0.75]
    const mid = at(0.5)
    const want = expectMix([64, 160, 208], [208, 96, 64]).map(v => v / 255)
    for (let i = 0; i < 3; i++) expect(mid[i]).toBeCloseTo(want[i]!, 1)
  })
})