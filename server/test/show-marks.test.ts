/** Piece 4 Task 6: both marks on show-plan layer assets (§7, D7, D9), and the validation that gates them. */
import { describe, expect, it } from 'bun:test'
import { canvasBackend, installCanvasBackend } from '../../src/core/canvas-backend'
import { compareFrames } from '../../src/slideshow/compare'
import type { LayerBitmap } from '../../src/slideshow/types'
import { flyInDeck } from '../../src/testdata/show-decks'
import { basis } from '../src/mark/basis'
import { extendedCodeword } from '../src/mark/bch'
import { detect } from '../src/mark/detect'
import { markLayerRaster } from '../src/mark/mark-raster'
import { messageFor } from '../src/mark/message'
import { DEFAULT_PARAMS } from '../src/mark/params'
import { decodePng, encodePng } from '../src/mark/png'
import { paintMark } from '../src/mark/visible'
import { markOptions } from '../src/render'
import type { BuiltShowPlan, ShowPlanJob } from '../src/show/build'
import { burnVisibleMark, MARK_FIDELITY } from '../src/show/marks'
import { serverBackend } from '../src/backend'
import { appearMoveDeck, openDeck, tinyBoxDeck } from './show-fixtures'

installCanvasBackend(serverBackend())
/** Canvases from the server backend: the visible mark's font resolves as it does in the build. */
const createCanvas = (w: number, h: number) => canvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement

const job = (over: Partial<ShowPlanJob> = {}): ShowPlanJob => ({ unitIndex: 0, scale: 1, motion: 'full', mark: { kind: 'fixed' }, strictMarks: false, ...over })
const built = (r: unknown): BuiltShowPlan => { if ((r as { refused?: string }).refused) throw new Error(`refused: ${(r as { refused: string }).refused}`); return r as BuiltShowPlan }
const key = Uint8Array.from({ length: 32 }, (_, i) => i + 1)
const codeFor = (token: number) => extendedCodeword(messageFor(key, 1, token, 17))
const TAU = { abs: 7.93, margin: 3.45 }

/**
 * The layers of `b` composed at the slide's last boundary, where every layer of these decks is drawn at its
 * rest: at its bounds translated by `offset(key)` slide px, opacity 1, in z order.
 */
function composeAtRest(b: BuiltShowPlan, scale: number, pngOf: (assetId: number) => Uint8Array, offset: (key: string | undefined) => number = () => 0) {
  const W = Math.ceil(b.plan.slide.w * scale), H = Math.ceil(b.plan.slide.h * scale)
  const c = createCanvas(W, H), ctx = c.getContext('2d')!
  for (const l of [...b.plan.layers].sort((p, q) => p.z - q.z)) {
    const d = decodePng(pngOf(l.assetId))
    const src = createCanvas(d.w, d.h), sx = src.getContext('2d')!
    const img = sx.createImageData(d.w, d.h); img.data.set(d.rgba); sx.putImageData(img, 0, 0)
    ctx.drawImage(src, (l.bounds.x + offset(l.key)) * scale, l.bounds.y * scale, l.bounds.w * scale, l.bounds.h * scale)
  }
  return { canvas: c, ctx, image: ctx.getImageData(0, 0, W, H) }
}
const asset = (b: BuiltShowPlan) => (id: number) => b.assets.find(a => a.assetId === id)!.png

describe('show-plan marks', () => {
  it('layer asset carries the visible mark source-atop: alpha unchanged, mark ink only where alpha > 0', () => {
    const w = 800, h = 600
    const canvas = createCanvas(w, h)
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#336699'
    ctx.beginPath(); ctx.ellipse(w / 2, h / 2, 380, 280, 0, 0, Math.PI * 2); ctx.fill()
    const layer: LayerBitmap = { canvas: canvas as never, bounds: { x: 80, y: 60, w, h }, scale: 1, bytes: w * h * 4 }
    const before = ctx.getImageData(0, 0, w, h).data
    const out = burnVisibleMark(layer, { x: 80, y: 60, w, h }, { kind: 'fixed' }, { w: 960, h: 720 })
    expect(out.canvas).not.toBe(layer.canvas)
    expect(out.bounds).toEqual(layer.bounds)
    const after = (out.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, w, h).data
    let inked = 0
    for (let i = 0; i < before.length; i += 4) {
      expect(after[i + 3]).toBe(before[i + 3]!)
      if (before[i + 3] === 0) expect([after[i], after[i + 1], after[i + 2]]).toEqual([0, 0, 0])
      else if (after[i] !== before[i] || after[i + 1] !== before[i + 1] || after[i + 2] !== before[i + 2]) inked++
    }
    expect(inked).toBeGreaterThan(100)
    // The input layer is not touched.
    expect(ctx.getImageData(0, 0, w, h).data).toEqual(before)
    // Off: the layer is returned as it is.
    expect(burnVisibleMark(layer, layer.bounds, { kind: 'off' }, { w: 960, h: 720 })).toBe(layer)
  })

  it("mover's mark is drawn at rest: composite at the step end equals the reference within MARK_FIDELITY", async () => {
    // The mover appears and moves 240 px right: its one visible rest is not where it is painted.
    const deck = await appearMoveDeck(0.25)
    const { host, docKey } = await openDeck(deck)
    const marked = built(await host.showPlan(docKey, job()))
    const plain = built(await host.showPlan(docKey, job({ mark: { kind: 'off' } })))
    expect(marked.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
    const mover = marked.plan.layers.find(l => l.key === 'sp:3')!
    expect(marked.assets.find(a => a.assetId === mover.assetId)!.rest.x).toBe(mover.bounds.x + 240)
    const offset = (k: string | undefined) => (k === 'sp:3' ? 240 : 0)
    const composite = composeAtRest(marked, 1, asset(marked), offset).image
    // The reference: the unmarked layers composed, then the page mark painted once.
    const ref = composeAtRest(plain, 1, asset(plain), offset)
    paintMark(ref.ctx as never, { widthPx: 960, heightPx: 720 }, markOptions({ kind: 'fixed' })!)
    const reference = ref.ctx.getImageData(0, 0, 960, 720)
    expect(compareFrames(composite as never, reference as never, MARK_FIDELITY).pass).toBe(true)
    // Per pixel over the mover's rest box (inset past its antialiased edge): its mark lines up with the page's.
    const region = { x: mover.bounds.x + 240 + 12, y: mover.bounds.y + 12, w: mover.bounds.w - 24, h: mover.bounds.h - 24 }
    const worst = (img: { data: Uint8ClampedArray }) => {
      let m = 0
      for (let y = region.y; y < region.y + region.h; y++) for (let x = region.x; x < region.x + region.w; x++) {
        const i = (y * 960 + x) * 4
        for (let c = 0; c < 3; c++) m = Math.max(m, Math.abs(img.data[i + c]! - reference.data[i + c]!))
      }
      return m
    }
    expect(worst(composite)).toBeLessThanOrEqual(2)
    // Control: burned at its painted position instead, the mover's mark is off the page's by whole glyphs.
    const wrongPng = (id: number) => {
      if (id !== mover.assetId) return asset(marked)(id)
      const d = decodePng(asset(plain)(id))
      const c = createCanvas(d.w, d.h), x = c.getContext('2d')!
      const img = x.createImageData(d.w, d.h); img.data.set(d.rgba); x.putImageData(img, 0, 0)
      const w = burnVisibleMark({ canvas: c as never, bounds: mover.bounds, scale: 1, bytes: 0 }, mover.bounds, { kind: 'fixed' }, { w: 960, h: 720 })
      return encodePng(d.w, d.h, (w.canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, d.w, d.h).data)
    }
    expect(worst(composeAtRest(marked, 1, wrongPng, offset).image)).toBeGreaterThan(8)
  }, 60_000)

  it('Layer-1 page basis cropped to rest bounds: composite at rest attributes the code', async () => {
    const { host, docKey } = await openDeck(await flyInDeck())
    const b = built(await host.showPlan(docKey, job()))
    const kBasis = Uint8Array.from({ length: 32 }, (_, i) => 200 - i)
    const code = codeFor(5)
    const sent = new Map<number, Uint8Array>()
    let carried = 0
    for (const a of b.assets) {
      if (a.tag !== 2) continue
      const r = markLayerRaster({ png: a.png, w: a.w, h: a.h, kBasis, code, crop: { x: a.rest.x, y: a.rest.y, pageW: 960, pageH: 720 } })
      if ('carriers' in r) carried += r.carriers
      sent.set(a.assetId, r.png)
    }
    expect(carried).toBeGreaterThan(400)
    const capture = composeAtRest(b, 1, id => sent.get(id)!).image
    const reference = composeAtRest(b, 1, asset(b)).image
    const chips = basis(kBasis, 960, 720, DEFAULT_PARAMS)
    const bmp = (i: { width: number; height: number; data: Uint8ClampedArray }) => ({ width: i.width, height: i.height, data: new Uint8ClampedArray(i.data) })
    const d = detect(bmp(capture), bmp(reference), chips, [codeFor(1), code, codeFor(9), codeFor(13)], DEFAULT_PARAMS, TAU)
    expect(d.kind).toBe('attributed')
    expect(d.kind === 'attributed' && d.index).toBe(1)
    // Unmarked, the same composite attributes nobody.
    expect(detect(bmp(reference), bmp(reference), chips, [codeFor(1), code], DEFAULT_PARAMS, TAU).kind).not.toBe('attributed')
  }, 60_000)

  it('zero-carrier layer is sent and audited unmarked/layer-capacity', async () => {
    const { host, docKey } = await openDeck(await tinyBoxDeck())
    const b = built(await host.showPlan(docKey, job()))
    const dot = b.plan.layers.find(l => l.key === 'sp:2')!
    const a = b.assets.find(x => x.assetId === dot.assetId)!
    const r = markLayerRaster({ png: a.png, w: a.w, h: a.h, kBasis: key, code: codeFor(3), crop: { x: a.rest.x, y: a.rest.y, pageW: 960, pageH: 720 } })
    expect('unmarked' in r && r.unmarked).toBe('layer-capacity')
    expect(r.png).toBe(a.png)
    // A layer whose carrier cells are all transparent keeps none either.
    const clear = encodePng(64, 64, new Uint8ClampedArray(64 * 64 * 4))
    const all = markLayerRaster({ png: clear, w: 64, h: 64, kBasis: key, code: codeFor(3), crop: { x: 0, y: 0, pageW: 960, pageH: 720 } })
    expect('unmarked' in all && all.unmarked).toBe('layer-capacity')
  })

  it('validation failure → every step composited', async () => {
    const { host, docKey } = await openDeck(await flyInDeck())
    // Seam: the layer composite at every boundary is replaced by a frame that cannot match its reference.
    host.showOptions = { compose: (_step, img) => { const d = new Uint8ClampedArray(img.data); for (let i = 0; i < d.length; i += 4) d[i] = 255 - d[i]!; return { width: img.width, height: img.height, data: d } as ImageData } }
    const b = built(await host.showPlan(docKey, job()))
    expect(b.plan.steps.every(s => s.eligibility === 'composited' && !!s.track)).toBe(true)
    expect(b.plan.layers).toEqual([])
    expect(b.assets.map(a => a.tag)).toEqual([3])
    expect(b.plan.diagnostics.map(d => d.code)).toContain('step-composited')
    host.showOptions = {}
    const ok = built(await host.showPlan(docKey, job({ mark: { kind: 'viewer', text: 'fresh' } })))
    expect(ok.plan.steps.map(s => s.eligibility)).toEqual(['layers'])
  }, 60_000)
})
