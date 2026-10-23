/**
 * S2-8: text builds by word and by letter (slideshow spec 6.5, 7).
 *
 * `p:iterate type="wd"/"lt"` splits the laid-out lines into word or glyph
 * boxes (from the text-layout runs), each its own layer, staggered by
 * `tmAbs`/`tmPct` and optionally `backwards`. Fixtures are hand-built
 * ("Office-shaped synthetic"; never reported as Office evidence).
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { layoutTextBody, type MeasureText } from '../src/drawing/text-layout'
import { createTextBodyMeasurer } from '../src/drawing/text-paint'
import { emuToPx } from '../src/core/geometry'
import type { PptxShape, PptxSlide } from '../src/pptx/types'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createScheduler } from '../src/slideshow/engine/scheduler'
import { expandIterations, iterateDelays, renderStateAt, type BaseGeometry } from '../src/slideshow/engine/behaviours'
import { buildRenderPlan, type PlanLayer } from '../src/slideshow/render/plan'
import { rasterLayer, type RasterContext } from '../src/slideshow/render/raster'
import { MAX_TEXT_BOXES, splitTextBoxes, textBoxBases } from '../src/slideshow/render/text-split'
import type { BehaviourNode, CTn, ContainerNode, Iterate, Lease, Rng, SlideShowWorkspace, TargetRef, TimeNode, TimingTree } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const PX = 9525
const SLIDE = { w: 960, h: 720 }
const rng: Rng = () => 0.5

/** 10 px per UTF-16 unit, whatever the style: boxes are exact multiples of 10. */
function fixedMeasure(calls?: { n: number }): MeasureText {
  return (text) => { if (calls) calls.n++; return { width: text.length * 10, ascent: 8, descent: 2, normalHeight: 10 } }
}

type RunSpec = string | { text: string; bold?: boolean }
function textShape(id: number, paragraphs: RunSpec[][], w = 400, h = 200, x = 0, y = 0): PptxShape {
  return {
    xEmu: x * PX, yEmu: y * PX, widthEmu: w * PX, heightEmu: h * PX, geometry: 'rect', fill: '#ffffff',
    source: { partPath: 'ppt/slides/slide1.xml', treePath: `spTree/sp[${id}]`, element: 'sp', id: String(id), representation: 'native' },
    textBody: {
      anchor: 't', wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
      paragraphs: paragraphs.map(runs => ({
        align: 'left' as const, level: 0,
        runs: runs.map(r => typeof r === 'string' ? { text: r, fontSizePt: 12 } : { text: r.text, bold: r.bold, fontSizePt: 12 }),
      })),
    },
  }
}
const layoutOf = (shape: PptxShape, measure: MeasureText) =>
  layoutTextBody(shape.textBody!, emuToPx(shape.widthEmu), emuToPx(shape.heightEmu), measure)
function slideOf(shapes: PptxShape[]): PptxSlide {
  return { index: 0, widthEmu: SLIDE.w * PX, heightEmu: SLIDE.h * PX, shapes, background: '#f0f0f0' }
}

function ctn(id: number, over: Partial<Omit<CTn, 'id'>> = {}): CTn {
  return { id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over }
}
function par(id: number, children: TimeNode[], over: Partial<Omit<CTn, 'id'>> = {}): ContainerNode {
  return { kind: 'par', ctn: ctn(id, over), children }
}
function tree(...children: TimeNode[]): TimingTree {
  return { root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot' }), build: [] }
}
function appear(id: number, target: TargetRef): BehaviourNode {
  return { kind: 'set', ctn: ctn(id, { durMs: 1, fill: 'hold' }), target, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' }
}
function fadeIn(id: number, target: TargetRef, durMs: number): BehaviourNode {
  return {
    kind: 'anim', ctn: ctn(id, { durMs, fill: 'hold' }), target, attrNames: ['style.opacity'], additive: 'base', accumulate: 'none',
    calcmode: 'lin', keys: [{ tm: 0, value: 0 }, { tm: 1, value: 1 }],
  }
}
/** An entrance effect container carrying `iterate`, the way PowerPoint writes "animate text by word/letter". */
function effect(id: number, iterate: Iterate, children: TimeNode[]): ContainerNode {
  return par(id, children, { presetClass: 'entr', presetId: 10, fill: 'hold', iterate })
}

function poolWorkspace(limit = 1 << 30): SlideShowWorkspace {
  let used = 0
  const lease = (bytes: number, tag: string): Lease | undefined => {
    if (used + bytes > limit) return undefined
    used += bytes
    let held = bytes, live = true
    const l: Lease = {
      get tag() { return tag }, get bytes() { return live ? held : 0 },
      grow(extra) { if (used + extra > limit) return false; used += extra; held += extra; return true },
      shrink(less) { const n = Math.min(less, held); held -= n; used -= n },
      transfer() { return l },
      release() { if (!live) return; live = false; used -= held },
    }
    return l
  }
  return { limit, lease, used: () => used, reserveTerminal: () => undefined }
}
function rasterCtx(): RasterContext {
  return { scale: 1, images: [], resolveFont: f => f, canvas: (w, h) => createCanvas(w, h) as unknown as HTMLCanvasElement, workspace: poolWorkspace(), maxTextureSize: 4096 }
}
const targets = (layers: readonly PlanLayer[]) => layers.flatMap(l => l.kind === 'target' ? [`${l.key}${l.part ? `:${l.part}` : ''}`] : [l.kind])

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('S2-8 text builds by word and by letter', () => {
  it('by word: each word is its own layer, staggered by tmPct', () => {
    const shape = textShape(5, [['Alpha beta gamma']])
    const boxes = splitTextBoxes(shape, layoutOf(shape, fixedMeasure()), 'wd', { measure: fixedMeasure() })
    expect(boxes.map(b => b.key)).toEqual(['sp:5/c:0-4', 'sp:5/c:6-9', 'sp:5/c:11-15'])

    // Each word is its own range layer over one body layer (geometry and fill painted once).
    const keys = boxes.map(b => b.key)
    const plan = buildRenderPlan(slideOf([shape]), new Set(['sp:5', ...keys]), new Map(), [], new Map([['sp:5', boxes]]))
    expect(targets(plan.layers)).toEqual(['static', 'sp:5:body', 'sp:5/c:0-4:range', 'sp:5/c:6-9:range', 'sp:5/c:11-15:range', 'watermark'])

    // tmPct 10 % of the 1000 ms effect: words start 0, 100 and 200 ms in.
    const T: TargetRef = { spid: '5' }
    const source = tree(effect(10, { type: 'wd', backwards: false, tmPct: 0.1 }, [appear(11, T), fadeIn(12, T, 1000)]))
    const sink = createDiagnosticSink()
    const x = expandIterations(source, new Map([['sp:5', keys]]), { sink, slideIndex: 0 })
    expect(x.fan.size).toBe(2)
    expect(sink.list().map(d => d.code)).toContain('text-build-approximate')
    // The source tree is never mutated.
    expect((source.root as ContainerNode).children[0].ctn.durMs).toBeUndefined()
    expect(((source.root as ContainerNode).children[0] as ContainerNode).children[1].ctn.durMs).toBe(1000)

    const base = new Map<string, BaseGeometry>([['sp:5', { x: 0, y: 0, w: 400, h: 200, rot: 0 }]])
    const sched = createScheduler(x.tree, { orderOf: x.orderOf })
    const at = (t: number) => {
      const snap = sched.evaluate([], t)
      return { snap, state: renderStateAt(snap, base, rng, sink, SLIDE, x.fan) }
    }
    const early = at(150)
    expect(early.state.get('sp:5/c:0-4')?.opacity).toBeCloseTo(0.15)
    expect(early.state.get('sp:5/c:6-9')?.opacity).toBeCloseTo(0.05)
    expect(early.state.get('sp:5/c:11-15')?.visible).toBe(false)
    // The body follows the first item.
    expect(early.state.get('sp:5')?.opacity).toBeCloseTo(0.15)
    // `order` stays the source pre-order index (root 0, effect 1, set 2, anim 3).
    expect([...early.snap.active, ...early.snap.frozen].map(e => e.order).sort()).toEqual([2, 3])

    // The effect lasts 1000 + 2 × 100 ms: the last word is still fading at 1100.
    const late = at(1100)
    expect(late.state.get('sp:5/c:0-4')?.opacity).toBeCloseTo(1)
    expect(late.state.get('sp:5/c:11-15')?.opacity).toBeCloseTo(0.9)
    expect(late.snap.ended.has(10)).toBe(false)
    const done = at(1250)
    expect(done.snap.ended.has(10)).toBe(true)
    for (const k of keys) expect(done.state.get(k)?.opacity).toBeCloseTo(1)
  })

  it('by letter with tmAbs and backwards', () => {
    expect(iterateDelays({ type: 'lt', backwards: true, tmAbsMs: 50 }, 3, 200)).toEqual([100, 50, 0])
    expect(iterateDelays({ type: 'lt', backwards: false, tmAbsMs: 50 }, 3, 200)).toEqual([0, 50, 100])
    expect(iterateDelays({ type: 'wd', backwards: false, tmPct: 0.25 }, 3, 200)).toEqual([0, 50, 100])

    const shape = textShape(7, [['ab c']])
    const boxes = splitTextBoxes(shape, layoutOf(shape, fixedMeasure()), 'lt', { measure: fixedMeasure() })
    // Whitespace is not a letter; boxes tile the line (the gap is split at its midpoint).
    expect(boxes.map(b => b.key)).toEqual(['sp:7/c:0-0', 'sp:7/c:1-1', 'sp:7/c:3-3'])
    expect(boxes.map(b => [b.box.x, b.box.x + b.box.w])).toEqual([[0, 10], [10, 25], [25, 400]])

    const T: TargetRef = { spid: '7' }
    const source = tree(effect(20, { type: 'lt', backwards: true, tmAbsMs: 50 }, [appear(21, T), fadeIn(22, T, 200)]))
    const sink = createDiagnosticSink()
    const keys = boxes.map(b => b.key)
    const x = expandIterations(source, new Map([['sp:7', keys]]), { sink })
    const sched = createScheduler(x.tree, { orderOf: x.orderOf })
    const state = (t: number) => renderStateAt(sched.evaluate([], t), new Map(), rng, sink, SLIDE, x.fan)
    const s60 = state(60)
    // Backwards: the last letter starts first.
    expect(s60.get('sp:7/c:3-3')?.opacity).toBeCloseTo(0.3)
    expect(s60.get('sp:7/c:1-1')?.opacity).toBeCloseTo(0.05)
    expect(s60.get('sp:7/c:0-0')?.visible).toBe(false)
    const s250 = state(250)
    expect(s250.get('sp:7/c:0-0')?.opacity).toBeCloseTo(0.75)
    expect(s250.get('sp:7/c:3-3')?.opacity).toBeCloseTo(1)
    for (const k of keys) expect(state(300).get(k)?.opacity).toBeCloseTo(1)
  })

  it('word boxes come from the text layout runs and cover the line', () => {
    // "Hel" + "lo" are two runs (two segments) forming one word; the line wraps after "big".
    const shape = textShape(9, [[{ text: 'Hel', bold: true }, 'lo big', ' world']], 100, 100)
    const layout = layoutOf(shape, fixedMeasure())
    expect(layout.lines).toHaveLength(2)
    expect(layout.lines[0].segments.length).toBeGreaterThanOrEqual(2)
    const boxes = splitTextBoxes(shape, layout, 'wd', { measure: fixedMeasure() })
    expect(boxes.map(b => b.key)).toEqual(['sp:9/c:0-4', 'sp:9/c:6-8', 'sp:9/c:10-14'])
    const [l1, l2] = layout.lines
    expect(boxes.map(b => b.box)).toEqual([
      { x: 0, y: l1.y, w: 55, h: l1.height },
      { x: 55, y: l1.y, w: 45, h: l1.height },
      { x: 0, y: l2.y, w: 100, h: l2.height },
    ])
    expect(boxes.map(b => b.paragraph)).toEqual([0, 0, 0])
    // Without a measurer the run widths are apportioned per UTF-16 unit: the same boxes here.
    expect(splitTextBoxes(shape, layout, 'wd').map(b => b.box)).toEqual(boxes.map(b => b.box))

    // Real glyphs: the layer for one word is cropped to its box and paints inside it.
    const c2d = createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
    const measure = createTextBodyMeasurer(c2d, f => f)
    const real = textShape(4, [['Alpha beta gamma']], 400, 60, 100, 50)
    const realBoxes = splitTextBoxes(real, layoutOf(real, measure), 'wd', { measure })
    expect(realBoxes).toHaveLength(3)
    const slide = slideOf([real])
    const plan = buildRenderPlan(slide, new Set(['sp:4', ...realBoxes.map(b => b.key)]), new Map(), [], new Map([['sp:4', realBoxes]]))
    const beta = plan.layers.find((l): l is Extract<PlanLayer, { kind: 'target' }> => l.kind === 'target' && l.key === realBoxes[1].key)!
    const r = rasterLayer(beta, slide, rasterCtx())
    if ('refused' in r) throw new Error(`refused ${r.refused}`)
    const box = realBoxes[1].box
    expect(r.bounds.x).toBeGreaterThanOrEqual(100 + box.x - 3)
    expect(r.bounds.x + r.bounds.w).toBeLessThanOrEqual(100 + box.x + box.w + 3)
    const data = (r.canvas as unknown as ReturnType<typeof createCanvas>).getContext('2d').getImageData(0, 0, r.canvas.width, r.canvas.height).data
    let ink = 0
    for (let i = 3; i < data.length; i += 4) if (data[i] > 0) ink++
    expect(ink).toBeGreaterThan(0)
    r.lease.release()
    // Base geometry per box, in slide px (shape origin plus the box).
    expect(textBoxBases(real, realBoxes).get(realBoxes[1].key)).toEqual({ x: 100 + box.x, y: 50 + box.y, w: box.w, h: box.h, rot: 0 })
  })

  it('more than 128 boxes coarsen to paragraph layers with a diagnostic', () => {
    expect(MAX_TEXT_BOXES).toBe(128)
    const words = (n: number, w: string) => Array.from({ length: n }, () => w).join(' ')
    const shape = textShape(3, [[words(100, 'aa')], [words(100, 'bb')]], 400, 4000)
    const layout = layoutOf(shape, fixedMeasure())
    const sink = createDiagnosticSink()
    const calls = { n: 0 }
    const boxes = splitTextBoxes(shape, layout, 'wd', { measure: fixedMeasure(calls), sink, slideIndex: 2 })
    expect(boxes.map(b => b.key)).toEqual(['sp:3/p:0-0', 'sp:3/p:1-1'])
    expect(boxes.map(b => b.paragraph)).toEqual([0, 1])
    // Counting stops at the bound: nothing was measured.
    expect(calls.n).toBe(0)
    expect(sink.list()).toEqual([expect.objectContaining({ code: 'text-build-coarsened', slideIndex: 2, effect: 'sp:3' })])
    const plan = buildRenderPlan(slideOf([shape]), new Set(['sp:3', 'sp:3/p:0-0', 'sp:3/p:1-1']), new Map(), [], new Map([['sp:3', boxes]]))
    expect(targets(plan.layers)).toEqual(['static', 'sp:3:body', 'sp:3/p:0-0:range', 'sp:3/p:1-1:range', 'watermark'])

    // More paragraphs than the bound as well: refused (the effect plays on the whole shape), diagnosed.
    const many = textShape(6, Array.from({ length: 130 }, () => ['x']), 400, 4000)
    const refusedSink = createDiagnosticSink()
    expect(splitTextBoxes(many, layoutOf(many, fixedMeasure()), 'lt', { sink: refusedSink })).toEqual([])
    expect(refusedSink.list().map(d => d.code)).toEqual(['text-build-refused'])

    // An allocation refused by the charge never happens: no boxes, diagnosed.
    const small = textShape(8, [['one two']])
    const chargeSink = createDiagnosticSink()
    const charged: number[] = []
    expect(splitTextBoxes(small, layoutOf(small, fixedMeasure()), 'wd', { sink: chargeSink, charge: (b) => { charged.push(b); return false } })).toEqual([])
    expect(charged).toHaveLength(1)
    expect(chargeSink.list().map(d => d.code)).toEqual(['text-build-refused'])

    // The fan-out is charged before the working copy is built; a refusal leaves the tree as it is.
    const T: TargetRef = { spid: '5' }
    const source = tree(effect(10, { type: 'wd', backwards: false, tmPct: 0.1 }, [fadeIn(12, T, 1000)]))
    const fanSink = createDiagnosticSink()
    const x = expandIterations(source, new Map([['sp:5', ['sp:5/c:0-1', 'sp:5/c:3-4']]]), { sink: fanSink, charge: () => false })
    expect(x.tree).toBe(source)
    expect(x.fan.size).toBe(0)
    expect(fanSink.list().map(d => d.code)).toContain('text-build-refused')
  })

  it('a paragraph-range target splits only its range; a target without boxes plays whole, diagnosed', () => {
    const shape = textShape(5, [['Intro'], ['one two']])
    const layout = layoutOf(shape, fixedMeasure())
    const boxes = splitTextBoxes(shape, layout, 'wd', { measure: fixedMeasure(), paragraphs: { st: 1, end: 1 } })
    // Paragraph 0 is 5 characters plus its mark: paragraph 1 starts at 6.
    expect(boxes.map(b => b.key)).toEqual(['sp:5/c:6-8', 'sp:5/c:10-12'])
    const builds = [{ spid: '5', mode: 'paragraph' as const, ranges: [{ key: 'sp:5/p:0-0', st: 0, end: 0 }, { key: 'sp:5/p:1-1', st: 1, end: 1 }], animBg: false, body: 'sp:5' }]
    const plan = buildRenderPlan(slideOf([shape]), new Set(['sp:5/p:0-0', 'sp:5/p:1-1']), new Map(), builds, new Map([['sp:5/p:1-1', boxes]]))
    expect(targets(plan.layers)).toEqual(['static', 'sp:5/p:0-0:range', 'sp:5/c:6-8:range', 'sp:5/c:10-12:range', 'watermark'])

    const R: TargetRef = { spid: '5', text: { kind: 'p', st: 1, end: 1 } }
    const other: TargetRef = { spid: '6' }
    const sink = createDiagnosticSink()
    const x = expandIterations(tree(effect(30, { type: 'wd', backwards: false, tmAbsMs: 100 }, [fadeIn(31, R, 100), fadeIn(32, other, 100)])),
      new Map([['sp:5/p:1-1', boxes.map(b => b.key)]]), { sink })
    expect(x.fan.size).toBe(1)
    expect([...x.fan.values()][0].items.map(i => i.key)).toEqual(['sp:5/c:6-8', 'sp:5/c:10-12'])
    expect(x.items.get('sp:5/p:1-1')).toEqual(['sp:5/c:6-8', 'sp:5/c:10-12'])
    expect(sink.list().map(d => d.code).sort()).toEqual(['text-build-approximate', 'text-build-unsplit'])
    const state = renderStateAt(createScheduler(x.tree, { orderOf: x.orderOf }).evaluate([], 150), new Map(), rng, sink, SLIDE, x.fan)
    expect(state.get('sp:5/c:6-8')?.opacity).toBeCloseTo(1)
    expect(state.get('sp:5/c:10-12')?.opacity).toBeCloseTo(0.5)
    expect(state.has('sp:5/p:1-1')).toBe(false)
    // The unsplit target played whole and ended at 100 ms.
    expect(state.get('sp:6')?.opacity).toBeCloseTo(1)
  })
})

describe('S2-8 fix round 1', () => {
  it('I1: unspaced CJK splits into words with Intl.Segmenter; punctuation stays with its word', () => {
    const cjk = textShape(5, [['我们是学生']])
    const boxes = splitTextBoxes(cjk, layoutOf(cjk, fixedMeasure()), 'wd', { measure: fixedMeasure() })
    expect(boxes.length).toBeGreaterThan(1)
    // The boxes partition the five characters in order.
    const ranges = boxes.map(b => b.key.split('/c:')[1].split('-').map(Number))
    expect(ranges[0][0]).toBe(0)
    expect(ranges[ranges.length - 1][1]).toBe(4)
    for (let i = 1; i < ranges.length; i++) expect(ranges[i][0]).toBe(ranges[i - 1][1] + 1)
    // Western text keeps its punctuation: "(Hello," and "world)" are two words.
    const west = textShape(6, [['(Hello, world)']])
    expect(splitTextBoxes(west, layoutOf(west, fixedMeasure()), 'wd', { measure: fixedMeasure() }).map(b => b.key)).toEqual(['sp:6/c:0-6', 'sp:6/c:8-13'])
  })

  it('I1: without Intl.Segmenter, CJK falls back to one box per character, diagnosed', () => {
    const real = (Intl as unknown as { Segmenter?: unknown }).Segmenter
    ;(Intl as unknown as { Segmenter?: unknown }).Segmenter = undefined
    try {
      const cjk = textShape(5, [['我们 ok']])
      const sink = createDiagnosticSink()
      const boxes = splitTextBoxes(cjk, layoutOf(cjk, fixedMeasure()), 'wd', { measure: fixedMeasure(), sink })
      expect(boxes.map(b => b.key)).toEqual(['sp:5/c:0-0', 'sp:5/c:1-1', 'sp:5/c:3-4'])
      expect(sink.list().map(d => d.code)).toEqual(['text-build-cjk-chars'])
    } finally {
      ;(Intl as unknown as { Segmenter?: unknown }).Segmenter = real
    }
  })

  it('I2: textBoxBases rotates each box about the shape pivot', () => {
    const shape: PptxShape = { ...textShape(5, [['x']], 100, 50), rotationDeg: 90 }
    const bases = textBoxBases(shape, [{ key: 'k', box: { x: 0, y: 0, w: 10, h: 10 }, paragraph: 0 }])
    const g = bases.get('k')!
    // Box centre (5, 5) is (−45, −20) from the pivot (50, 25); rotated 90° it is (20, −45).
    expect(g.x).toBeCloseTo(65)
    expect(g.y).toBeCloseTo(-25)
    expect([g.w, g.h, g.rot]).toEqual([10, 10, 90])
  })

  it('I3: under an exit effect the shape body follows the last item', () => {
    const T: TargetRef = { spid: '5' }
    const exit = par(10, [fadeIn(12, T, 1000)], { presetClass: 'exit', presetId: 10, fill: 'hold', iterate: { type: 'wd', backwards: false, tmAbsMs: 100 } })
    const x = expandIterations(tree(exit), new Map([['sp:5', ['sp:5/c:0-1', 'sp:5/c:3-4', 'sp:5/c:6-7']]]))
    expect([...x.fan.values()][0].items).toEqual([
      { key: 'sp:5', delayMs: 200 }, { key: 'sp:5/c:0-1', delayMs: 0 }, { key: 'sp:5/c:3-4', delayMs: 100 }, { key: 'sp:5/c:6-7', delayMs: 200 },
    ])
    const entr = expandIterations(tree(effect(10, { type: 'wd', backwards: false, tmAbsMs: 100 }, [fadeIn(12, T, 1000)])), new Map([['sp:5', ['sp:5/c:0-1', 'sp:5/c:3-4']]]))
    expect([...entr.fan.values()][0].items[0]).toEqual({ key: 'sp:5', delayMs: 0 })
  })

  it('M1: RTL lines give their boxes in reading order', () => {
    // A bidi layout places the first Hebrew word on the right.
    const shape = textShape(5, [['שלום עולם']], 200, 50)
    const seg = (text: string, x: number, st: number) => ({ text, x, width: text.length * 10, style: { fontSizePt: 12 }, sourceStart: st, sourceEnd: st + text.length, runIndex: 0, sourceRuns: [], graphemeBoundaries: Array.from({ length: text.length }, (_, i) => st + i) })
    const layout = { height: 20, lines: [{ paragraphIndex: 0, x: 0, y: 0, baseline: 16, height: 20, segments: [seg('עולם', 0, 5), seg(' ', 40, 4), seg('שלום', 50, 0)] }] }
    const boxes = splitTextBoxes(shape, layout as never, 'wd')
    expect(boxes.map(b => b.key)).toEqual(['sp:5/c:0-3', 'sp:5/c:5-8'])
    expect(boxes.map(b => [b.box.x, b.box.x + b.box.w])).toEqual([[45, 200], [0, 45]])
  })
})

describe('PR #19 review 6 follow-up (colour)', () => {
  it('a whole-shape colour effect composes on top of each split item\'s own colour', async () => {
    const { renderStateAt } = await import('../src/slideshow/engine/behaviours')
    const { createRng } = await import('../src/slideshow/rng')
    const { createDiagnosticSink } = await import('../src/slideshow/diagnostics')
    const ctn = (id: number) => ({ id, durMs: 1000, accel: 0, decel: 0, autoRev: false, start: [], end: [], fill: 'hold' as const })
    const set = { kind: 'set', ctn: ctn(5), target: { spid: '2' }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' } as const
    const tint = { kind: 'animClr', ctn: ctn(6), target: { spid: '2' }, attrNames: ['style.color'], additive: 'base', accumulate: 'none', clrSpc: 'rgb', to: '#ff0000' } as unknown
    const entry = (node: unknown, order: number) => ({ node, progress: 1, iteration: 0, order, phase: 'frozen' })
    const snapshot = { active: [], frozen: [entry(set, 0), entry(tint, 1)], ended: new Set<number>() } as never
    const fan = new Map([[set as never, { items: [{ key: 'sp:2', delayMs: 0 }, { key: 'sp:2/c:0-4', delayMs: 0 }], spanMs: 1, totalMs: 1, accel: 0, decel: 0, autoRev: false, entrance: true }]])
    const base = new Map([['sp:2', { x: 0, y: 0, w: 100, h: 50, rot: 0 }]])
    const states = renderStateAt(snapshot, base, createRng(1), createDiagnosticSink(), { w: 960, h: 540 }, fan as never)
    const body = states.get('sp:2')!
    expect(body.color?.text).toBeDefined()
    expect(states.get('sp:2/c:0-4')!.color?.text).toEqual(body.color?.text)
  })
})
