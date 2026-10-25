/**
 * S1-12: the z-order render plan, the layer rasteriser and the direct painter
 * (slideshow spec 6.3, 7). Pixels come from node-canvas; layering and order
 * are asserted against `renderSlide`, the static viewer's painter.
 *
 * Fixtures are hand-built `PptxShape` trees plus builder decks parsed by
 * `parsePptx` ("Office-shaped synthetic"; never reported as Office evidence).
 */
import { describe, expect, it } from 'vitest'
import { createCanvas, type Canvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'
import type { PptxShape, PptxSlide } from '../src/pptx/types'
import { buildPptx, type PptxShapeSpec } from '../src/testdata/ooxml-builders'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { buildRenderPlan, type PlanLayer, type RenderPlan } from '../src/slideshow/render/plan'
import { rasterLayer, type RasterContext } from '../src/slideshow/render/raster'
import { paintDirect } from '../src/slideshow/render/direct'
import { textFilter } from '../src/slideshow/render/scene'
import type { BuildOwnership, ColorMatrix, Lease, SlideShowWorkspace, SubOwner, TargetState, TerminalReserve } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const PX = 9525
const W = 200, H = 120

function src(id: number) { return { partPath: 'ppt/slides/slide1.xml', treePath: `spTree/sp[${id}]`, element: 'sp', id: String(id), representation: 'native' as const } }
function rect(id: number, x: number, y: number, w: number, h: number, fill?: string, more: Partial<PptxShape> = {}): PptxShape {
  return { xEmu: x * PX, yEmu: y * PX, widthEmu: w * PX, heightEmu: h * PX, geometry: 'rect', fill, source: src(id), ...more }
}
/** A group placed at (x, y, w, h) whose child space is (0, 0)–(cw, ch). */
function group(id: number, x: number, y: number, w: number, h: number, cw: number, ch: number, children: PptxShape[]): PptxShape {
  return {
    ...rect(id, x, y, w, h), geometry: 'rect', children,
    group: { off: { x: x * PX, y: y * PX }, ext: { width: w * PX, height: h * PX }, chOff: { x: 0, y: 0 }, chExt: { width: cw * PX, height: ch * PX } },
  }
}
function slideOf(shapes: PptxShape[], more: Partial<PptxSlide> = {}): PptxSlide {
  return { index: 0, widthEmu: W * PX, heightEmu: H * PX, shapes, background: '#f0f0f0', ...more }
}
async function parsedSlide(shapes: PptxShapeSpec[]): Promise<PptxSlide> {
  const doc = await parsePptx(await OfficePackage.load(await buildPptx(shapes)))
  const slide = doc.slides[0]
  // buildPptx writes every p:sp with cNvPr id 2; give each shape its spec's id.
  slide.shapes.forEach((shape, i) => { if (shapes[i].id !== undefined) shape.source = { ...shape.source!, id: String(shapes[i].id) } })
  return { ...slide, widthEmu: W * PX, heightEmu: H * PX }
}
const emu = (px: number) => String(px * PX)

// ---------------------------------------------------------------------------
// Workspace (test ledger typed by the S1-2 contracts; S1-16 owns the real one)
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
    grow(extra) { if (!live || pool.used + extra > pool.limit) return false; pool.used += extra; held += extra; return true },
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
    reserveTerminal(outputBytes) {
      const scratchBytes = 64 * 1024
      const whole = poolLease(pool, outputBytes + scratchBytes, 'terminal')
      if (!whole) return undefined
      const sub: Pool = { used: 0, limit: scratchBytes }
      const scratch: SubOwner = { lease: (bytes, tag) => poolLease(sub, bytes, tag) }
      const reserve: TerminalReserve & { sub: Pool } = { outputBytes, scratch, sub, resize: () => false, release: () => whole.release() }
      return reserve
    },
  }
}

function canvasFactory(log?: string[]) {
  const calls: Array<[number, number]> = []
  const make = (w: number, h: number) => { calls.push([w, h]); log?.push('canvas'); return createCanvas(w, h) as unknown as HTMLCanvasElement }
  return { make, calls }
}
function rasterCtx(more: Partial<RasterContext> = {}): RasterContext {
  return { scale: 1, images: [], resolveFont: f => f, canvas: canvasFactory().make, workspace: makeWorkspace(1 << 30), maxTextureSize: 4096, ...more }
}

// ---------------------------------------------------------------------------
// Pixels
// ---------------------------------------------------------------------------
type C2D = CanvasRenderingContext2D
function reference(slide: PptxSlide, images?: Array<CanvasImageSource | undefined>): Canvas {
  const canvas = createCanvas(W, H)
  renderSlide(slide, canvas.getContext('2d') as unknown as C2D, undefined, images)
  return canvas
}
/** Composite every layer in plan order (sub-layers in place), as the compositor will. */
function composite(slide: PptxSlide, plan: RenderPlan, ctx: RasterContext, matrix?: (layer: PlanLayer) => ColorMatrix | undefined): Canvas {
  const out = createCanvas(W, H)
  const o = out.getContext('2d')
  const draw = (layer: PlanLayer): void => {
    const r = rasterLayer(layer, slide, ctx)
    if (!('refused' in r)) {
      let img = r.canvas as unknown as Canvas
      const m = matrix?.(layer)
      if (m) img = applyMatrix(img, m)
      o.drawImage(img, r.bounds.x, r.bounds.y, r.bounds.w, r.bounds.h)
      r.lease.release()
    }
    if (layer.kind === 'target') for (const sub of layer.sub ?? []) draw(sub)
  }
  for (const layer of plan.layers) draw(layer)
  return out
}
function applyMatrix(src: Canvas, m: ColorMatrix): Canvas {
  const c = createCanvas(src.width, src.height), x = c.getContext('2d')
  x.drawImage(src, 0, 0)
  const img = x.getImageData(0, 0, c.width, c.height), d = img.data
  for (let i = 0; i < d.length; i += 4) {
    const v = [d[i] / 255, d[i + 1] / 255, d[i + 2] / 255, d[i + 3] / 255]
    for (let row = 0; row < 4; row++) {
      const k = row * 5
      d[i + row] = Math.round(255 * Math.min(1, Math.max(0, m[k] * v[0] + m[k + 1] * v[1] + m[k + 2] * v[2] + m[k + 3] * v[3] + m[k + 4])))
    }
  }
  x.putImageData(img, 0, 0)
  return c
}
function pixels(c: Canvas): Uint8ClampedArray { return c.getContext('2d').getImageData(0, 0, c.width, c.height).data }
function diff(a: Canvas, b: Canvas): number {
  const pa = pixels(a), pb = pixels(b)
  let n = 0
  for (let i = 0; i < pa.length; i++) if (pa[i] !== pb[i]) n++
  return n
}
function maxDelta(a: Canvas, b: Canvas): number {
  const pa = pixels(a), pb = pixels(b)
  let max = 0
  for (let i = 0; i < pa.length; i++) max = Math.max(max, Math.abs(pa[i] - pb[i]))
  return max
}
function at(c: Canvas, x: number, y: number): number[] { return [...c.getContext('2d').getImageData(x, y, 1, 1).data] }
function regionDiff(a: Canvas, b: Canvas, x0: number, y0: number, x1: number, y1: number): number {
  let n = 0
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) if (at(a, x, y).some((v, i) => v !== at(b, x, y)[i])) n++
  return n
}

/** Every leaf (non-group) node reachable from a layer, sub-layers included. */
function leaves(layer: PlanLayer): PptxShape[] {
  if (layer.kind === 'watermark') return []
  const walk = (nodes: readonly PptxShape[]): PptxShape[] => nodes.flatMap(n => n.group ? walk(n.children ?? []) : [n])
  return [...walk(layer.nodes), ...(layer.kind === 'target' ? (layer.sub ?? []).flatMap(leaves) : [])]
}
const shown = (over: Partial<TargetState> = {}): TargetState => ({ visible: true, opacity: 1, transform: [1, 0, 0, 1, 0, 0], ...over })
const NONE = new Map<string, ReadonlySet<'fill' | 'line' | 'text'>>()

// ---------------------------------------------------------------------------

describe('buildRenderPlan', () => {
  it('static A, animated B, static C becomes three layers in authored order', () => {
    const a = rect(2, 10, 10, 50, 40, '#ff0000'), b = rect(3, 40, 30, 50, 40, '#00ff00'), c = rect(4, 70, 50, 50, 40, '#0000ff')
    const plan = buildRenderPlan(slideOf([a, b, c]), new Set(['sp:3']), NONE, [])
    expect(plan.layers.map(l => l.kind)).toEqual(['static', 'target', 'static', 'watermark'])
    expect(plan.layers[0]).toMatchObject({ kind: 'static', nodes: [a], background: true })
    expect(plan.layers[1]).toMatchObject({ kind: 'target', key: 'sp:3', nodes: [b] })
    expect(plan.layers[2]).toMatchObject({ kind: 'static', nodes: [c] })
    expect((plan.layers[2] as { background?: true }).background).toBeUndefined()
  })

  it('underlay nodes join the first static run, under the slide shapes, and are never targets', () => {
    const art = rect(3, 0, 0, 20, 20, '#123456')
    const a = rect(3, 30, 30, 20, 20, '#ff0000')
    const slide = slideOf([a], { underlay: [art] } as Partial<PptxSlide>)
    // `sp:3` names the slide shape; the master shape with the same id stays static.
    const plan = buildRenderPlan(slide, new Set(['sp:3']), NONE, [])
    expect(plan.layers[0]).toMatchObject({ kind: 'static', nodes: [art], background: true })
    expect(plan.layers[1]).toMatchObject({ kind: 'target', key: 'sp:3', nodes: [a] })
    const unanimated = buildRenderPlan(slide, new Set(), NONE, [])
    expect(unanimated.layers[0]).toMatchObject({ kind: 'static', nodes: [art, a] })
  })

  it('an animated first shape still gets a background layer underneath it', () => {
    const a = rect(2, 10, 10, 50, 40, '#ff0000')
    const plan = buildRenderPlan(slideOf([a]), new Set(['sp:2']), NONE, [])
    expect(plan.layers.map(l => l.kind)).toEqual(['static', 'target', 'watermark'])
    expect(plan.layers[0]).toMatchObject({ nodes: [], background: true })
  })

  it('the watermark layer is last and no render state moves it', () => {
    const slide = slideOf([rect(2, 0, 0, W, H, '#336699'), group(5, 20, 20, 100, 60, 100, 60, [rect(6, 0, 0, 40, 40, '#ff0000')])])
    for (const animated of [new Set<string>(), new Set(['sp:2']), new Set(['sp:5', 'sp:6']), new Set(['sp:6'])]) {
      const plan = buildRenderPlan(slide, animated, NONE, [])
      expect(plan.layers[plan.layers.length - 1]).toEqual({ kind: 'watermark' })
      expect(plan.layers.filter(l => l.kind === 'watermark')).toHaveLength(1)
    }
    // The direct painter keeps it on top and outside every target's transform.
    const mark = { text: 'WM', placement: 'center' as const, rotate: 0, opacity: 0.8, color: '#ff00ff', fontSizePt: 40 }
    const slide2 = slideOf([rect(2, 0, 0, W, H, '#336699')])
    const paint = (state: Map<string, TargetState>) => {
      const c = createCanvas(W, H)
      paintDirect(c.getContext('2d') as unknown as C2D, slide2, state, { images: [], resolveFont: f => f, scale: 1, scratch: makeWorkspace(1 << 20), watermark: mark })
      return c
    }
    const hidden = paint(new Map([['sp:2', shown({ visible: false })]]))
    const movedAway = paint(new Map([['sp:2', shown({ transform: [1, 0, 0, 1, 5000, 0], opacity: 0.3, motion: { dx: 3000, dy: 0 } })]]))
    expect(diff(hidden, movedAway)).toBe(0)
    // ...and the mark is really there.
    expect(diff(hidden, paint(new Map([['sp:2', shown({ visible: false })]])))).toBe(0)
    const noMark = createCanvas(W, H)
    paintDirect(noMark.getContext('2d') as unknown as C2D, slide2, new Map([['sp:2', shown({ visible: false })]]), { images: [], resolveFont: f => f, scale: 1, scratch: makeWorkspace(1 << 20) })
    expect(diff(hidden, noMark)).toBeGreaterThan(0)
  })

  it('bldP build=p bldLvl=2 expands into one TargetKey per paragraph, level-3 paragraphs joining their level-2 parent; animBg animates the shape body', async () => {
    // Paragraph levels 1, 2, 3, 2: S1-8's expandBuilds yields ranges 0-0, 1-2 (the level-3 paragraph joins its level-2 parent) and 3-3.
    const slide = await parsedSlide([{ id: 5, off: [emu(10), emu(10)], ext: [emu(180), emu(100)], fill: 'CCCCCC', paragraphs: ['one', 'two', 'three', 'four'].map(text => ({ runs: [{ text }] })) }])
    const ranges = [{ key: 'sp:5/p:0-0', st: 0, end: 0 }, { key: 'sp:5/p:1-2', st: 1, end: 2 }, { key: 'sp:5/p:3-3', st: 3, end: 3 }]
    const animated = new Set(ranges.map(r => r.key))
    const withBg: BuildOwnership = { spid: '5', mode: 'paragraph', ranges, animBg: true, body: 'sp:5' }
    const plan = buildRenderPlan(slide, new Set([...animated, 'sp:5']), NONE, [withBg])
    const targets = plan.layers.filter((l): l is Extract<PlanLayer, { kind: 'target' }> => l.kind === 'target')
    expect(targets.map(t => [t.key, t.part])).toEqual([['sp:5', 'body'], ['sp:5/p:0-0', 'range'], ['sp:5/p:1-2', 'range'], ['sp:5/p:3-3', 'range']])
    const keeps = (layer: PlanLayer) => [0, 1, 2, 3].filter(i => textFilter.get(leaves(layer)[0])!(i))
    expect(targets.map(keeps)).toEqual([[], [0], [1, 2], [3]])
    // Range layers carry text only: no geometry, fill or line.
    for (const t of targets.slice(1)) expect(leaves(t)[0]).toMatchObject({ drawingGeometry: undefined, drawingStyle: undefined })

    // animBg off: the body is static and joins the surrounding static run.
    const noBg = buildRenderPlan(slide, animated, NONE, [{ ...withBg, animBg: false }])
    expect(noBg.layers.map(l => l.kind === 'target' ? l.key : l.kind)).toEqual(['static', 'sp:5/p:0-0', 'sp:5/p:1-2', 'sp:5/p:3-3', 'watermark'])
    expect(keeps(noBg.layers[0])).toEqual([])
  })
})

describe('rasterLayer', () => {
  it('translucent overlap composites identically to renderSlide at the static state', () => {
    // Every translucent overlap crosses a layer boundary or lies over an opaque shape (see the next test for overlaps inside one layer).
    const slide = slideOf([
      rect(2, 5, 5, 60, 50, '#ff0000', { line: { color: '#000000', widthEmu: 25400 } }),
      rect(3, 40, 25, 60, 50, 'rgba(0,0,255,0.5)'),
      rect(4, 80, 45, 60, 50, 'rgba(255,200,0,0.6)', { rotationDeg: 20 }),
      group(9, 155, 72, 40, 40, 120, 120, [rect(10, 0, 0, 80, 80, '#00aa00'), rect(11, 40, 40, 80, 80, 'rgba(0,255,255,0.5)')]),
    ])
    const ref = reference(slide)
    for (const animated of [new Set<string>(), new Set(['sp:3']), new Set(['sp:3', 'sp:4']), new Set(['sp:2', 'sp:4']), new Set(['sp:9', 'sp:11']), new Set(['sp:11'])]) {
      expect(diff(composite(slide, buildRenderPlan(slide, animated, NONE, []), rasterCtx()), ref)).toBe(0)
    }
  })

  it('translucent overlaps inside one non-base layer differ from renderSlide by at most 1 (8-bit premultiplied rounding)', async () => {
    const parsed = await parsedSlide([
      { id: 7, off: [emu(100), emu(10)], ext: [emu(80), emu(60)], fill: '00AA00', lineW: '38100', paragraphs: [{ runs: [{ text: 'Hello', color: '000000' }] }] },
    ])
    const outline = parsed.shapes[0]
    outline.drawingStyle!.fill = { kind: 'solid', color: { r: 0, g: 170, b: 0, a: 0.5 } }
    const slide = slideOf([
      rect(2, 10, 10, 80, 60, '#ff0000'),
      rect(3, 50, 30, 80, 60, 'rgba(0,0,255,0.5)', { line: { color: '#000000', widthEmu: 25400 } }),
      rect(4, 90, 50, 80, 60, 'rgba(255,200,0,0.6)', { rotationDeg: 20 }),
      outline,
      group(9, 20, 70, 60, 40, 120, 80, [rect(10, 0, 0, 80, 60, 'rgba(255,0,255,0.5)'), rect(11, 40, 20, 80, 60, 'rgba(0,255,255,0.5)')]),
    ])
    const ref = reference(slide)
    for (const animated of [new Set<string>(), new Set(['sp:3']), new Set(['sp:3', 'sp:7']), new Set(['sp:2', 'sp:4', 'sp:11']), new Set(['sp:9', 'sp:11'])]) {
      expect(maxDelta(composite(slide, buildRenderPlan(slide, animated, NONE, []), rasterCtx()), ref)).toBeLessThanOrEqual(1)
    }
    // A single base layer is exact.
    expect(diff(composite(slide, buildRenderPlan(slide, new Set(), NONE, []), rasterCtx()), ref)).toBe(0)
  })

  it('grouped child animation keeps the group transform and is painted once', () => {
    const c1 = rect(11, 0, 0, 60, 60, 'rgba(255,0,0,0.5)')
    const c2 = rect(12, 30, 30, 60, 60, 'rgba(0,0,255,0.5)')
    const c3 = rect(13, 60, 0, 60, 60, 'rgba(0,160,0,0.5)')
    // Child space 120×90 shown in 60×45 px at (40, 20): every child is halved.
    const g = group(10, 40, 20, 60, 45, 120, 90, [c1, c2, c3])
    const slide = slideOf([rect(2, 0, 0, 30, 30, '#888888'), g])
    const plan = buildRenderPlan(slide, new Set(['sp:10', 'sp:12']), NONE, [])
    const target = plan.layers[1] as Extract<PlanLayer, { kind: 'target' }>
    expect(target).toMatchObject({ kind: 'target', key: 'sp:10', nodes: [] })
    expect(target.sub!.map(l => l.kind === 'target' ? l.key : l.kind)).toEqual(['static', 'sp:12', 'static'])
    // Each sub-layer keeps the group as its pruned ancestor.
    for (const sub of target.sub!) expect((sub as { nodes: PptxShape[] }).nodes[0].group).toEqual(g.group)
    // Painted once: every leaf appears in exactly one layer.
    const all = plan.layers.flatMap(leaves)
    for (const leaf of [c1, c2, c3]) expect(all.filter(n => n === leaf)).toHaveLength(1)
    // The sub-target sits where the group transform puts it: (40 + 15, 20 + 15), 30×30.
    const r = rasterLayer(target.sub![1], slide, rasterCtx())
    if ('refused' in r) throw new Error('refused')
    expect(r.bounds.x).toBeGreaterThan(50); expect(r.bounds.x).toBeLessThanOrEqual(55)
    expect(r.bounds.x + r.bounds.w).toBeGreaterThanOrEqual(85); expect(r.bounds.x + r.bounds.w).toBeLessThan(90)
    expect(diff(composite(slide, plan, rasterCtx()), reference(slide))).toBe(0)
    // An unanimated group with an animated child splits into runs around it.
    const plain = buildRenderPlan(slide, new Set(['sp:12']), NONE, [])
    expect(plain.layers.map(l => l.kind === 'target' ? l.key : l.kind)).toEqual(['static', 'sp:12', 'static', 'watermark'])
    expect(diff(composite(slide, plain, rasterCtx()), reference(slide))).toBe(0)
  })

  it('shape background plus paragraph build paints the background once', async () => {
    const slide = await parsedSlide([{ id: 5, off: [emu(10), emu(10)], ext: [emu(180), emu(100)], fill: '3366CC', paragraphs: ['Alpha', 'Beta', 'Gamma'].map(text => ({ runs: [{ text, sz: '2000' }] })) }])
    slide.shapes[0].drawingStyle!.fill = { kind: 'solid', color: { r: 51, g: 102, b: 204, a: 0.5 } }
    const ranges = [{ key: 'sp:5/p:0-0', st: 0, end: 0 }, { key: 'sp:5/p:1-1', st: 1, end: 1 }, { key: 'sp:5/p:2-2', st: 2, end: 2 }]
    const ref = reference(slide)
    for (const animBg of [false, true]) {
      const plan = buildRenderPlan(slide, new Set(ranges.map(r => r.key)), NONE, [{ spid: '5', mode: 'paragraph', ranges, animBg, body: 'sp:5' }])
      const out = composite(slide, plan, rasterCtx())
      // Painting the half-transparent fill twice would shift every fill pixel by ~60; glyph edges composited
      // from their own layer may round by 1 (8-bit premultiplied), and the fill away from the text is exact.
      expect(maxDelta(out, ref)).toBeLessThanOrEqual(1)
      expect(regionDiff(out, ref, 150, 15, 185, 105)).toBe(0)
      // The range layers hold glyphs only: their corners (inside the shape, away from text) are transparent.
      for (const layer of plan.layers.filter(l => l.kind === 'target' && l.part === 'range')) {
        const r = rasterLayer(layer, slide, rasterCtx())
        if ('refused' in r) throw new Error('refused')
        const c = r.canvas as unknown as Canvas
        expect(at(c, c.width - 2, 1)[3]).toBe(0)
      }
    }
  })

  it('a useBgFill shape split for a colour animation shows the background only in its fill layer', async () => {
    const slide = await parsedSlide([{ id: 6, off: [emu(20), emu(20)], ext: [emu(160), emu(80)], fill: 'FF0000', lineW: '38100', paragraphs: [{ runs: [{ text: 'MMMM', sz: '4000', color: 'FF0000' }] }] }])
    slide.shapes[0].useBgFill = true
    const plan = buildRenderPlan(slide, new Set(['sp:6']), new Map([['sp:6', new Set(['fill' as const])]]), [])
    const part = (name: string) => plan.layers.find(l => l.kind === 'target' && l.part === name) as { nodes: PptxShape[] }
    expect(part('fill').nodes.map(n => n.useBgFill)).toEqual([true])
    // The line layer paints over the fill layer: a backdrop there would cover the fill's colour animation.
    expect(part('line').nodes.map(n => n.useBgFill)).toEqual([undefined])
  })

  it('same-colour fill and text, fill-only colour change with partial transparency, changes only the fill', async () => {
    const slide = await parsedSlide([{ id: 6, off: [emu(20), emu(20)], ext: [emu(160), emu(80)], fill: 'FF0000', lineW: '38100', paragraphs: [{ runs: [{ text: 'MMMM', sz: '4000', b: true, color: 'FF0000' }] }] }])
    slide.shapes[0].drawingStyle!.fill = { kind: 'solid', color: { r: 255, g: 0, b: 0, a: 0.5 } }
    const plan = buildRenderPlan(slide, new Set(['sp:6']), new Map([['sp:6', new Set(['fill' as const])]]), [])
    expect(plan.layers.filter(l => l.kind === 'target').map(l => (l as { part?: string }).part)).toEqual(['fill', 'line', 'text'])
    const ref = reference(slide)
    // Glyph edges composited from their own layer may round by 1; the fill and outline are exact.
    const same = composite(slide, plan, rasterCtx())
    expect(maxDelta(same, ref)).toBeLessThanOrEqual(1)
    expect(regionDiff(same, ref, 15, 15, 185, 30)).toBe(0)
    // Swap red and blue on the fill layer only.
    const swap: ColorMatrix = [0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0]
    const out = composite(slide, plan, rasterCtx(), l => l.kind === 'target' && l.part === 'fill' ? swap : undefined)
    // Fill area away from the text: was half-transparent red over the background, now half-transparent blue.
    const [fr, , fb] = at(ref, 30, 95), [or, , ob] = at(out, 30, 95)
    expect(fr).toBeGreaterThan(fb + 100)
    expect(ob).toBeGreaterThan(or + 100)
    // Opaque text pixels are untouched; so is the outline.
    const refPx = pixels(ref), outPx = pixels(out)
    let textPixels = 0
    for (let i = 0; i < refPx.length; i += 4) {
      if (refPx[i] === 255 && refPx[i + 1] === 0 && refPx[i + 2] === 0 && refPx[i + 3] === 255) {
        textPixels++
        // Fully covered glyph pixels; the swapped fill may bleed in by rounding only.
        expect(outPx[i]).toBeGreaterThanOrEqual(253)
        expect(Math.max(outPx[i + 1], outPx[i + 2])).toBeLessThanOrEqual(2)
      }
    }
    expect(textPixels).toBeGreaterThan(50)
    expect(at(out, 20, 60)).toEqual(at(ref, 20, 60))
  })

  it('a layer larger than maxTextureSize rasterises at reduced scale with a diagnostic', () => {
    const shape = rect(2, 10, 10, 150, 100, '#ff0000')
    const slide = slideOf([shape])
    const plan = buildRenderPlan(slide, new Set(['sp:2']), NONE, [])
    const sink = createDiagnosticSink()
    const r = rasterLayer(plan.layers[1], slide, rasterCtx({ scale: 2, maxTextureSize: 128, sink }))
    if ('refused' in r) throw new Error('refused')
    expect(r.scale).toBeLessThan(2)
    expect(r.canvas.width).toBeLessThanOrEqual(128)
    expect(r.canvas.height).toBeLessThanOrEqual(128)
    expect(r.bytes).toBe(r.canvas.width * r.canvas.height * 4)
    // Still covers the shape in slide px.
    expect(r.bounds.x).toBeLessThanOrEqual(10); expect(r.bounds.x + r.bounds.w).toBeGreaterThanOrEqual(160)
    expect(sink.list()).toEqual([expect.objectContaining({ code: 'layer-scale', slideIndex: 0 })])
  })

  it('layers are cropped to padded bounds, not full slide', () => {
    const slide = slideOf([rect(2, 50, 40, 20, 10, '#ff0000', { line: { color: '#000000', widthEmu: 12700 } })])
    const plan = buildRenderPlan(slide, new Set(['sp:2']), NONE, [])
    const r = rasterLayer(plan.layers[1], slide, rasterCtx({ scale: 2 }))
    if ('refused' in r) throw new Error('refused')
    expect(r.bounds.x).toBeLessThan(50); expect(r.bounds.y).toBeLessThan(40)
    expect(r.bounds.x + r.bounds.w).toBeGreaterThan(70); expect(r.bounds.y + r.bounds.h).toBeGreaterThan(50)
    expect(r.bounds.w).toBeLessThan(50); expect(r.bounds.h).toBeLessThan(40)
    expect(r.canvas.width).toBe(r.bounds.w * 2)
    expect(r.bytes).toBe(r.canvas.width * r.canvas.height * 4)
    expect(r.bytes).toBeLessThan(W * 2 * H * 2 * 4 / 10)
    // Integer device-pixel origin, so compositing never resamples.
    expect(Number.isInteger(r.bounds.x * 2) && Number.isInteger(r.bounds.y * 2)).toBe(true)
    // The background run alone is full-slide.
    const bg = rasterLayer(plan.layers[0], slide, rasterCtx())
    if ('refused' in bg) throw new Error('refused')
    expect(bg.bounds).toEqual({ x: 0, y: 0, w: W, h: H })
  })

  it('rasterLayer leases before creating its canvas and refuses without allocating', async () => {
    const text = await parsedSlide([{ id: 8, off: [emu(10), emu(10)], ext: [emu(100), emu(40)], fill: 'FF0000', paragraphs: [{ runs: [{ text: 'Hi' }] }] }])
    const slide = slideOf([rect(2, 10, 10, 40, 40, '#ff0000'), text.shapes[0]])
    const plan = buildRenderPlan(slide, new Set(['sp:2', 'sp:8']), NONE, [])
    // Fully reserved: nothing is allocated, for a plain layer, a text layer, the background or the watermark.
    const factory = canvasFactory()
    const full = makeWorkspace(0)
    for (const layer of plan.layers) {
      expect(rasterLayer(layer, slide, rasterCtx({ canvas: factory.make, workspace: full, watermark: { text: 'W' } }))).toEqual({ refused: 'workspace' })
    }
    expect(factory.calls).toEqual([])
    expect(full.used()).toBe(0)
    // With room: the lease always precedes the canvas it pays for.
    const log: string[] = []
    const ws = makeWorkspace(1 << 24, log)
    const spy = canvasFactory(log)
    const results = plan.layers.map(layer => rasterLayer(layer, slide, rasterCtx({ canvas: spy.make, workspace: ws })))
    expect(log.length).toBeGreaterThan(0)
    for (let i = 0; i < log.length; i++) if (log[i] === 'canvas') expect(log[i - 1]).toMatch(/^lease:/)
    const held = results.reduce((n, r) => n + ('refused' in r ? 0 : r.bytes), 0)
    expect(ws.used()).toBe(held)
    for (const r of results) if (!('refused' in r)) r.lease.release()
    expect(ws.used()).toBe(0)
  })

  it('with a watermark, an opaque shape over the watermark area still shows the watermark on top', () => {
    const mark = { text: 'CONFIDENTIAL', placement: 'center' as const, rotate: 0, opacity: 0.6, color: '#ff00ff', fontSizePt: 18 }
    const slide = slideOf([rect(2, 0, 0, W, H, '#003366')])
    const plain = reference(slide)
    // The static viewer paints the mark under the shapes, so the opaque shape hides it.
    const staticView = createCanvas(W, H)
    renderSlide(slide, staticView.getContext('2d') as unknown as C2D, undefined, undefined, mark)
    expect(diff(staticView, plain)).toBe(0)
    // The slideshow puts it on top: layered and direct.
    for (const animated of [new Set<string>(), new Set(['sp:2'])]) {
      const layered = composite(slide, buildRenderPlan(slide, animated, NONE, []), rasterCtx({ watermark: mark }))
      expect(regionDiff(layered, plain, 60, 50, 140, 70)).toBeGreaterThan(20)
    }
    const direct = createCanvas(W, H)
    paintDirect(direct.getContext('2d') as unknown as C2D, slide, new Map(), { images: [], resolveFont: f => f, scale: 1, scratch: makeWorkspace(1 << 20), watermark: mark })
    expect(regionDiff(direct, plain, 60, 50, 140, 70)).toBeGreaterThan(20)
    expect(diff(direct, composite(slide, buildRenderPlan(slide, new Set(), NONE, []), rasterCtx({ watermark: mark })))).toBe(0)
  })
})

describe('paintDirect', () => {
  it('paintDirect at the static state matches renderSlide pixel-for-pixel', async () => {
    const parsed = await parsedSlide([
      { id: 5, off: [emu(100), emu(10)], ext: [emu(90), emu(70)], fill: '3366CC', lineW: '25400', paragraphs: ['Alpha', 'Beta', 'Gamma'].map(text => ({ runs: [{ text }] })) },
    ])
    parsed.shapes[0].drawingStyle!.fill = { kind: 'solid', color: { r: 51, g: 102, b: 204, a: 0.5 } }
    const slide = slideOf([
      rect(2, 10, 10, 80, 60, '#ff0000'),
      rect(3, 50, 30, 80, 60, 'rgba(0,0,255,0.5)', { rotationDeg: 15, flipH: true }),
      parsed.shapes[0],
      group(9, 20, 70, 60, 40, 120, 80, [rect(10, 0, 0, 80, 60, 'rgba(255,0,255,0.5)'), rect(11, 40, 20, 80, 60, 'rgba(0,255,255,0.5)')]),
    ])
    const ref = reference(slide)
    const scratch = makeWorkspace(1 << 20)
    const paint = (state: Map<string, TargetState>, builds?: BuildOwnership[]) => {
      const c = createCanvas(W, H)
      paintDirect(c.getContext('2d') as unknown as C2D, slide, state, { images: [], resolveFont: f => f, scale: 1, scratch, builds })
      return c
    }
    expect(diff(paint(new Map()), ref)).toBe(0)
    // Every kind of target at its static state: shapes, a group with a sub-target, paragraph ranges (keys read back from the state).
    const state = new Map(['sp:2', 'sp:3', 'sp:9', 'sp:11', 'sp:5', 'sp:5/p:0-0', 'sp:5/p:1-2'].map(k => [k, shown()]))
    expect(diff(paint(state), ref)).toBe(0)
    // A hidden target is skipped; a translated one moves.
    expect(diff(paint(new Map([['sp:2', shown({ visible: false })]])), ref)).toBeGreaterThan(0)
    expect(at(paint(new Map([['sp:2', shown({ visible: false })]])), 15, 15)).toEqual([240, 240, 240, 255])
    expect(at(paint(new Map([['sp:2', shown({ motion: { dx: -100, dy: 0 } })]])), 15, 15)).toEqual([240, 240, 240, 255])
    // At device scale 2 the output is the slide at twice the size.
    const big = createCanvas(W * 2, H * 2)
    paintDirect(big.getContext('2d') as unknown as C2D, slide, new Map(), { images: [], resolveFont: f => f, scale: 2, scratch })
    expect(at(big, 30, 30)).toEqual(at(ref, 15, 15))
    expect(scratch.used()).toBe(0)
  })

  it('tables paint through the delegated content path identically to renderSlide, layered and direct', async () => {
    const parsed = await parsedSlide([
      { id: 20, off: [emu(30), emu(20)], ext: [emu(140), emu(80)], table: {
        colWidths: [emu(70), emu(70)],
        rows: [{ h: emu(40), cells: [{ fill: 'FFCC00', paragraphs: [{ runs: [{ text: 'A1' }] }] }, { paragraphs: [{ runs: [{ text: 'B1' }] }] }] },
          { h: emu(40), cells: [{ paragraphs: [{ runs: [{ text: 'A2' }] }] }, { fill: '3366CC', paragraphs: [{ runs: [{ text: 'B2' }] }] }] }],
      } },
    ])
    const table = parsed.shapes[0]
    expect(table.table).toBeDefined()
    table.source = { ...table.source!, id: '20' }
    const slide = slideOf([rect(2, 0, 0, 60, 60, '#ff0000'), table], { theme: parsed.theme })
    const ref = reference(slide)
    expect(at(ref, 40, 25)).toEqual([255, 204, 0, 255])
    expect(diff(composite(slide, buildRenderPlan(slide, new Set(), NONE, []), rasterCtx()), ref)).toBe(0)
    // As its own layer over an opaque base: glyph edges may round by 1, the cell fills are exact.
    const layered = composite(slide, buildRenderPlan(slide, new Set(['sp:20']), NONE, []), rasterCtx())
    expect(maxDelta(layered, ref)).toBeLessThanOrEqual(1)
    expect(regionDiff(layered, ref, 32, 22, 45, 30)).toBe(0)
    const direct = createCanvas(W, H)
    paintDirect(direct.getContext('2d') as unknown as C2D, slide, new Map([['sp:20', shown()]]), { images: [], resolveFont: f => f, scale: 1, scratch: makeWorkspace(1 << 20) })
    expect(diff(direct, ref)).toBe(0)
  })

  it('paintDirect with terminal.scratch succeeds while the parent workspace is fully reserved; a scratch overflow paints placeholders; parent used() does not grow', () => {
    const small = createCanvas(10, 10), smallCtx = small.getContext('2d')
    smallCtx.fillStyle = '#00ff00'; smallCtx.fillRect(0, 0, 10, 10)
    // 200×100×4 = 80,000 B: more than the 64 KiB scratch allowance on its own.
    const huge = createCanvas(200, 100), hugeCtx = huge.getContext('2d')
    hugeCtx.fillStyle = '#0000ff'; hugeCtx.fillRect(0, 0, 200, 100)
    const images = [small, huge] as unknown as CanvasImageSource[]
    const slide = slideOf([rect(2, 10, 10, 40, 40, undefined, { imageIndex: 0 }), rect(3, 100, 10, 60, 40, undefined, { imageIndex: 1 })])

    const ws = makeWorkspace(W * H * 4 + 64 * 1024 + 1000)
    const terminal = ws.reserveTerminal(W * H * 4) as TerminalReserve & { sub: Pool }
    expect(terminal).toBeDefined()
    // Every remaining parent byte goes to someone else.
    const rest = ws.lease(ws.limit - ws.used(), 'other')!
    expect(ws.lease(1, 'more')).toBeUndefined()
    const parentBefore = ws.used()

    const sink = createDiagnosticSink()
    const out = createCanvas(W, H)
    paintDirect(out.getContext('2d') as unknown as C2D, slide, new Map(), { images, resolveFont: f => f, scale: 1, scratch: terminal.scratch, sink })
    expect(at(out, 30, 30)).toEqual([0, 255, 0, 255])
    // The over-allowance image is the static viewer's missing-image placeholder: nothing drawn, diagnosed.
    expect(at(out, 130, 30)).toEqual([240, 240, 240, 255])
    expect(sink.list()).toEqual([expect.objectContaining({ code: 'scratch-image', slideIndex: 0 })])
    expect(slide.shapes[1].diagnostics).toBeUndefined()
    expect(ws.used()).toBe(parentBefore)
    expect(terminal.sub.used).toBe(0)

    // The static viewer with the image withheld paints the same pixels.
    expect(diff(out, reference(slide, [small as unknown as CanvasImageSource, undefined]))).toBe(0)
    rest.release()
  })
})
