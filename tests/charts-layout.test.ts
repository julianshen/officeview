// @vitest-environment node
import { describe, it, expect } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import type { ThemeContext } from '../src/drawing/style'

const tx = (size: number, body = '') => `<txPr><bodyPr ${body}/><p><pPr><defRPr sz="${size}"/></pPr></p></txPr>`
export const chartXml = (extra = '', axisExtra = '', rootExtra = '') => `<chartSpace>${rootExtra}<chart><autoTitleDeleted val="1"/><plotArea>${extra}<barChart><ser><cat><strLit><pt idx="0"><v>Alpha long label</v></pt><pt idx="1"><v>Beta long label</v></pt></strLit></cat><val><numLit><pt idx="0"><v>3</v></pt><pt idx="1"><v>7</v></pt></numLit></val></ser><axId val="1"/><axId val="2"/></barChart><catAx><axId val="1"/>${axisExtra}<crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx></plotArea></chart></chartSpace>`
const parse = (xml: string, theme?: ThemeContext) => parseXlsxChart(parseXmlOrdered(xml), theme)!

describe('chart placement input', () => {
  it('uses the theme minor font and clamps txPr to 6–72 pt', () => {
    const theme = { palette: {}, fonts: { minor: { latin: 'Test Minor' } } } as ThemeContext
    expect(parse(chartXml('', '', tx(9000)), theme).fontFamily).toBe('Test Minor')
    expect(parse(chartXml('', '', tx(9000)), theme).fontSizePt).toBe(72)
    expect(parse(chartXml('', '', tx(200))).fontSizePt).toBe(6)
  })
  it('preserves axis rotation, vertical text and per-axis size', () => {
    const model = parse(chartXml('', tx(1800, 'rot="-2700000" vert="vert270"')))
    expect(model.axes[0]).toMatchObject({ text: { sizePt: 18, rotation: -45, vertical: 'vert270' } })
  })
  it('preserves manual modes and leaves omitted values automatic', () => {
    const model = parse(chartXml('<layout><manualLayout><x val="0.1"/><yMode val="edge"/><y val="0.2"/><wMode val="edge"/><w val="0.8"/></manualLayout></layout>'))
    expect(model).toMatchObject({ plotLayout: { x: .1, y: .2, w: .8, yMode: 'edge', wMode: 'edge' } })
  })
})

import * as placement from '../src/charts/layout'
import { documentChartContext } from '../src/charts/context'
const measure = { width: (s: string) => s.length * 7, lineHeight: () => 14 }
describe('charged chart placement', () => {
  it('caches by size, dpr and font generation, keeps charges until invalidation/dispose', () => {
    const model = parse(chartXml()), context = documentChartContext('mobile')
    let calls = 0
    const metrics = { ...measure, width: (s: string) => { calls++; return measure.width(s) } }
    const first = placement.layoutChart(model, 400, 260, 1, 0, metrics, context)
    const initialCalls = calls, held = context.workspace.used
    expect(held).toBeGreaterThan(0)
    expect(placement.layoutChart(model, 400, 260, 1, 0, metrics, context)).toBe(first)
    expect(calls).toBe(initialCalls)
    for (const [w, dpr, gen] of [[500, 1, 0], [500, 2, 0], [500, 2, 1]]) {
      expect(placement.layoutChart(model, w, 260, dpr, gen, metrics, context)).not.toBe(first)
      expect(context.workspace.used).toBe(held)
    }
    placement.releaseChartLayouts(model)
    expect(context.workspace.used).toBe(0)
  })
  it('refuses before measuring when the live parent has no headroom', () => {
    const context = documentChartContext('mobile', undefined, { reserve: () => false, release: () => {} })
    let calls = 0
    expect(() => placement.layoutChart(parse(chartXml()), 400, 260, 1, 0, { ...measure, width: () => { calls++; return 10 } }, context)).toThrow()
    expect(calls).toBe(0)
    expect(context.workspace.used).toBe(0)
  })
  it('resolves manual edge/factor modes in order and falls back on inverted boxes', () => {
    const auto = { x: 20, y: 30, w: 200, h: 100 }, diagnostics: any[] = []
    expect(placement.resolveManualLayout({ x: .1, y: .2, w: .8, h: .7, wMode: 'edge', hMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual({ x: 60, y: 90, w: 260, h: 120 })
    expect(placement.resolveManualLayout({ x: .1, y: .2, xMode: 'edge', yMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual({ x: 40, y: 60, w: 200, h: 100 })
    expect(placement.resolveManualLayout({ x: .9, w: .2, xMode: 'edge', wMode: 'edge' }, auto, 400, 300, diagnostics)).toEqual(auto)
    expect(diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-layout-invalid' }))
  })
  it('thins rotated labels in screen order without overlapping extents', () => {
    const model = parse(chartXml('', tx(1800, 'rot="-2700000"'))), context = documentChartContext('mobile')
    const layout = placement.layoutChart(model, 130, 200, 1, 0, measure, context)
    const labels = layout.labels.get('1')!
    expect(labels.length).toBeGreaterThan(0)
    expect(labels[0].angle).toBe(-Math.PI / 4)
    for (let i = 1; i < labels.length; i++) expect(labels[i].position - labels[i].extent / 2).toBeGreaterThanOrEqual(labels[i-1].position + labels[i-1].extent / 2 + 4)
    placement.releaseChartLayouts(model)
  })
})

import { paintXlsxChart, acquireChartContext } from '../src/charts/paint'
import { createCanvas } from 'canvas'
describe('layout painter lifetime', () => {
  it('retains layouts but refunds paint transients, avoids warm measurements, and disposes with receiver', () => {
    const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
    const ctx = createCanvas(400, 260).getContext('2d')
    let calls = 0; const original = ctx.measureText.bind(ctx)
    ctx.measureText = s => { calls++; return original(s) }
    paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
    const held = receiver.context.workspace.used, measured = calls
    expect(held).toBeGreaterThan(0)
    paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
    expect(calls).toBe(measured)
    expect(receiver.context.workspace.used).toBe(held)
    receiver.release()
    expect(receiver.context.workspace.used).toBe(0)
  })
})

import { fillChartRanges } from '../src/charts/fill'
describe('size-independent preparation', () => {
  it('transfers domains, candidate ticks and formatted labels; placement does not rescan values', () => {
    const model = structuredClone(parse(chartXml()))
    expect(model).toHaveProperty('preparedAxes')
    const numeric = model.groups[0].series[0].values
    Object.defineProperty(numeric, 'values', { get: () => { throw new Error('Stage 2 rescanned values') } })
    const context = documentChartContext('mobile')
    const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
    expect(layout.graph.axes.get('2')!.max).toBeGreaterThanOrEqual(7)
    expect(layout.labels.get('2')!.some(l => l.text === '0')).toBe(true)
    placement.releaseChartLayouts(model)
  })
})

import { releaseChartModel, chartModelOwner } from '../src/charts/own'
it('model disposal releases its placement sidecar as well as its prepared data', () => {
  const context = documentChartContext('mobile')
  const model = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!, owner = chartModelOwner(model)!
  placement.layoutChart(model, 400, 260, 1, 0, measure)
  expect(owner.workspace.used).toBeGreaterThan(0)
  releaseChartModel(model)
  expect(owner.workspace.used).toBe(0)
})
it('uses the resolved legend box size and inner plot box literally', () => {
  const xml = chartXml('<layout><manualLayout><layoutTarget val="inner"/><xMode val="edge"/><yMode val="edge"/><x val="0.2"/><y val="0.2"/><w val="0.5"/><h val="0.5"/></manualLayout></layout>').replace('</chart>', '<legend><legendPos val="r"/><layout><manualLayout><xMode val="edge"/><yMode val="edge"/><x val="0.1"/><y val="0.1"/><w val="0.2"/><h val="0.4"/></manualLayout></layout></legend></chart>')
  const model = parse(xml), layout = placement.layoutChart(model, 400, 300, 1, 0, measure)
  expect(layout.plot).toEqual({ x: 80, y: 60, w: 200, h: 150 })
  expect(layout.legend).toEqual({ x: 40, y: 30, w: 80, h: 120 })
  const ctx = createCanvas(400, 300).getContext('2d'), swatches: number[][] = []
  const fill = ctx.fillRect.bind(ctx)
  ctx.fillRect = (x, y, w, h) => { swatches.push([x,y,w,h]); fill(x,y,w,h) }
  paintXlsxChart(model, ctx as any, 400, 300)
  expect(swatches).toContainEqual([46, 30 + (model.fontSizePt * 96 / 72 + 6 - model.fontSizePt * 96 / 72 * .7) / 2, model.fontSizePt * 96 / 72 * .7, model.fontSizePt * 96 / 72 * .7])
  placement.releaseChartLayouts(model)
})
it('explicit receiving contexts retain layouts until the last preparation releases them', () => {
  const model = structuredClone(parse(chartXml())), context = documentChartContext('mobile')
  const a = acquireChartContext({ owner: {}, context }), b = acquireChartContext({ owner: {}, context })
  placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  a.release(); expect(context.workspace.used).toBeGreaterThan(0)
  b.release(); expect(context.workspace.used).toBe(0)
})
it('honours tickMarkSkip even when category labels are hidden', () => {
  const paint = (skip: number) => {
    const model = parse(chartXml('', `<tickLblPos val="none"/><majorTickMark val="out"/><tickMarkSkip val="${skip}"/>`))
    const ctx = createCanvas(400, 260).getContext('2d'), marks: number[][] = []
    const line = ctx.lineTo.bind(ctx); ctx.lineTo = (x, y) => { marks.push([x,y]); line(x,y) }
    paintXlsxChart(model, ctx as any, 400, 260)
    return marks.length
  }
  expect(paint(1) - paint(2)).toBe(1)
})
import { PointWriter } from '../src/charts/points'
it('drops later area series at the path work bound with a diagnostic', () => {
  const model = parse(chartXml()), points = new PointWriter(4096)
  for (let i = 0; i < 4096; i++) points.value(i, i % 7)
  delete model.preparedAxes
  model.groups[0].family = 'area'
  model.groups[0].series = Array.from({ length: 64 }, () => ({ name: 's', color: '112233', values: points.done() }))
  model.categories = []
  const context = documentChartContext('desktop'); let segments = 0
  const ctx = new Proxy({ font: '', measureText: () => ({ width: 10 }) }, { get: (t: any, k: any) => k in t ? t[k] : (..._args: any[]) => { if (k === 'lineTo' || k === 'moveTo' || k === 'arc') segments++ }, set: (t:any,k:any,v:any) => {t[k]=v;return true} })
  paintXlsxChart(model, ctx as any, 400, 260, undefined, context)
  const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  expect(segments).toBeGreaterThan(4096)
  expect(segments).toBeLessThanOrEqual(262144)
  expect(layout.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-path-limit' }))
  placement.releaseChartLayouts(model)
})
it('keeps an invalidated layout charged until its active paint returns', () => {
  const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
  const ctx = createCanvas(400, 260).getContext('2d'), original = ctx.fillRect.bind(ctx)
  let once = false, during = 0
  ctx.fillRect = (x,y,w,h) => { if (!once) { once = true; receiver.release(); during = receiver.context.workspace.used }; original(x,y,w,h) }
  paintXlsxChart(model, ctx as any, 400, 260, undefined, receiver.context)
  expect(during).toBeGreaterThanOrEqual(placement.layoutBound(model))
  expect(receiver.context.workspace.used).toBe(0)
})
it('invalidates charged widths on font publication and resolves the theme font through the font manager', () => {
  const model = structuredClone(parse(chartXml())), receiver = acquireChartContext({ owner: {}, profile: 'mobile' })
  const ctx = createCanvas(400, 260).getContext('2d'), fonts = new EventTarget()
  Object.defineProperty(ctx.canvas, 'ownerDocument', { value: { fonts } })
  const seen: string[] = []; const original = ctx.measureText.bind(ctx)
  ctx.measureText = text => { seen.push(ctx.font); return original(text) }
  paintXlsxChart(model, ctx as any, 400, 260, () => '"Registered Minor"', receiver.context)
  expect(seen.some(font => font.includes('Registered Minor'))).toBe(true)
  const count = seen.length
  fonts.dispatchEvent(new Event('loadingdone'))
  expect(receiver.context.workspace.used).toBe(0)
  paintXlsxChart(model, ctx as any, 400, 260, () => '"Registered Minor"', receiver.context)
  expect(seen.length).toBeGreaterThan(count)
  receiver.release()
})
it('replaces stage-1 ownership on a refill instead of accumulating obsolete snapshots', () => {
  const context = documentChartContext('mobile')
  const first = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!
  const before = context.workspace.used
  const second = fillChartRanges(first, 'S', () => undefined)
  expect(second.preparedAxes).not.toBe(first.preparedAxes)
  expect(context.workspace.used).toBe(before)
  releaseChartModel(second)
  expect(context.workspace.used).toBe(0)
})
it('keeps authored chart text within 72 pt when painting a title', () => {
  const model = parse(chartXml('', '', tx(7200)).replace('<autoTitleDeleted val="1"/>', '<title><tx><rich><p><r><t>Title</t></r></p></rich></tx></title>'))
  const ctx = createCanvas(400, 260).getContext('2d'), sizes: number[] = [], fill = ctx.fillText.bind(ctx)
  ctx.fillText = ((text: string, x: number, y: number, width?: number) => { if (text === 'Title') sizes.push(parseFloat(ctx.font)); if(width === undefined)fill(text,x,y);else fill(text,x,y,width) }) as typeof ctx.fillText
  paintXlsxChart(model, ctx as any, 400, 260)
  expect(sizes).toEqual([96])
})
it('keeps existing value labels when their boxes fit without overlap', () => {
  const model = parse(chartXml()), context = documentChartContext('mobile')
  const layout = placement.layoutChart(model, 100, 100, 1, 0, measure, context)
  expect(layout.labels.get('2')).toHaveLength(layout.graph.axes.get('2')!.ticks.length)
  placement.releaseChartLayouts(model)
})
it('reuses prepared value-label strings during warm paint', () => {
  const model = parse(chartXml()), context = documentChartContext('mobile')
  const ctx = createCanvas(400, 260).getContext('2d')
  paintXlsxChart(model, ctx as any, 400, 260, undefined, context)
  const layout = placement.layoutChart(model, 400, 260, 1, 0, measure, context)
  for (const axis of layout.graph.axes.values()) axis.labelValue = () => { throw new Error('Reformatted during paint') }
  expect(() => paintXlsxChart(model, ctx as any, 400, 260, undefined, context)).not.toThrow()
  placement.releaseChartLayouts(model)
})
import { stampWorkbookEpoch } from '../src/charts/fill'
it('contains an epoch preparation refusal to its chart', () => {
  const context = documentChartContext('mobile')
  const model = parseXlsxChart(parseXmlOrdered(chartXml()), undefined, { ...context, partPath: 'c.xml' } as any)!
  context.workspace.reserve(context.workspace.limit - context.workspace.used, 'pressure')
  const drawings = [{ content: { kind: 'xlsxChart', chart: model } }]
  expect(() => stampWorkbookEpoch(drawings as any, '1904')).not.toThrow()
  expect(drawings[0].content.chart.refused).toBe(true)
  context.workspace.release('pressure')
  releaseChartModel(drawings[0].content.chart)
  expect(context.workspace.used).toBe(0)
})
