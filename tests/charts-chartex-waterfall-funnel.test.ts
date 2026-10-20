/**
 * Task B2: the ChartEx waterfall and funnel families (spec 4.3; docs/charts/chartex-map.md §§6, 9).
 *
 * Native evidence: WF1, WF2 (Excel waterfalls), FN1 (Excel funnel), FN2 (PowerPoint funnel) in tests/fixtures/chartex
 * (PROVENANCE.md). Inline parts are Office-shaped synthetic (never Office evidence), used for the cases no native
 * part exhibits: a subtotal whose own value differs from the running total, data-point colour overrides, connector
 * visibility off, several visible series. Numeric oracles are computed by hand in the comments.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContextFor } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { isChartExGroup } from '../src/charts/chartex/parse'
import { forEachWaterfallBar, waterfallLevels, WATERFALL_CONNECTOR, type WaterfallKind } from '../src/charts/chartex/waterfall'
import { CHARTEX_SEVERAL_VISIBLE, funnelGeometry } from '../src/charts/chartex/funnel'
import { prepareChartAxes } from '../src/charts/prepare'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart } from '../src/charts/layout'
import { displayView } from '../src/charts/display'
import type { ChartExGroup, ChartModel, ChartPoint, ChartResult } from '../src/charts/model'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { describeFamilyConformance, type ConformanceFamily } from './helpers/chart-family-conformance'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const FIXTURES = join(__dirname, 'fixtures', 'chartex')
const fixture = (name: string) => new Uint8Array(readFileSync(join(FIXTURES, name)))
const theme = parseThemeContext()
const W = 480, H = 320
const ACCENT = ['#4472c4', '#ed7d31', '#a5a5a5']

/** A part read through the production bounded reader and the production adapter. */
async function adapt(xml: string, path = 'ppt/charts/chartEx1.xml'): Promise<ChartResult> {
  const zip = new JSZip()
  zip.file(path, xml)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const context = chartParseContextFor(pkg, path)
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error(`${path}: ${read ? read.reason : 'missing'}`)
  try { return await parseChartPart(read.root, theme, context) } finally { read.release() }
}
async function native(xml: string, path?: string): Promise<ChartModel> {
  const result = await adapt(xml, path)
  expect(result.kind, result.kind === 'native' ? '' : result.reason).toBe('native')
  if (result.kind !== 'native') throw new Error('not native')
  return prepareChartAxes(result.model)
}
const space = (body: string) => `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="${A}">${body}</cx:chartSpace>`
const level = (values: Array<number | string | undefined>) => `<cx:lvl ptCount="${values.length}">${values.map((v, i) => v === undefined ? '' : `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`
/** An Office-shaped synthetic waterfall: one cached value dimension, the given layoutPr and series extras. */
const waterfall = (values: Array<number | string | undefined>, layoutPr = '<cx:subtotals />', extra = '') => space(
  `<cx:chartData><cx:data id="0"><cx:numDim type="val">${level(values)}</cx:numDim></cx:data></cx:chartData>` +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall">${extra}<cx:dataId val="0" /><cx:layoutPr>${layoutPr}</cx:layoutPr></cx:series></cx:plotAreaRegion>` +
  `<cx:axis id="0"><cx:catScaling gapWidth="0.5" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis></cx:plotArea>` +
  `<cx:legend pos="t" align="ctr" overlay="0" /></cx:chart>`)
/** An Office-shaped synthetic funnel: per series its values and hidden flag, shared categories. */
const funnel = (series: Array<{ values: Array<number | string | undefined>; hidden?: boolean }>, categories?: string[], gap = '0.5') => space(
  `<cx:chartData>${series.map((s, i) => `<cx:data id="${i}">${categories ? `<cx:strDim type="cat">${level(categories)}</cx:strDim>` : ''}<cx:numDim type="val">${level(s.values)}</cx:numDim></cx:data>`).join('')}</cx:chartData>` +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion>${series.map((s, i) => `<cx:series layoutId="funnel"${s.hidden ? ' hidden="1"' : ''} formatIdx="${i}"><cx:tx><cx:txData><cx:v>S${i + 1}</cx:v></cx:txData></cx:tx><cx:dataId val="${i}" /></cx:series>`).join('')}</cx:plotAreaRegion>` +
  `<cx:axis id="1"><cx:catScaling gapWidth="${gap}" /><cx:tickLabels /></cx:axis></cx:plotArea></cx:chart>`)

const measureCtx = createCanvas(10, 10).getContext('2d')
const plotOf = (model: ChartModel) => layoutChart(model, W, H, 1, 0, { width: (t, f) => { measureCtx.font = f; return measureCtx.measureText(t).width }, lineHeight: () => model.fontSizePt * 96 / 72 }, undefined, JSON.stringify(model.fontFamily)).plot
interface Rect { x: number; y: number; w: number; h: number; color: string }
/** Paints `model` on a recording canvas: filled rectangles, connector segments and drawn text. */
function record(model: ChartModel) {
  const ctx = createCanvas(W, H).getContext('2d') as unknown as CanvasRenderingContext2D & Record<string, unknown>
  const rects: Rect[] = [], connectors: Array<[number, number, number, number]> = [], texts: Array<{ text: string; x: number; y: number; align: string }> = []
  let path: Array<[number, number]> = []
  const fillRect = ctx.fillRect.bind(ctx), moveTo = ctx.moveTo.bind(ctx), lineTo = ctx.lineTo.bind(ctx), beginPath = ctx.beginPath.bind(ctx), stroke = ctx.stroke.bind(ctx), fillText = ctx.fillText.bind(ctx)
  ctx.fillRect = (x: number, y: number, w: number, h: number) => { rects.push({ x, y, w, h, color: String(ctx.fillStyle).toLowerCase() }); fillRect(x, y, w, h) }
  ctx.beginPath = () => { path = []; beginPath() }
  ctx.moveTo = (x: number, y: number) => { path.push([x, y]); moveTo(x, y) }
  ctx.lineTo = (x: number, y: number) => { path.push([x, y]); lineTo(x, y) }
  ctx.stroke = () => {
    if (String(ctx.strokeStyle).toLowerCase() === WATERFALL_CONNECTOR.toLowerCase()) for (let i = 0; i + 1 < path.length; i += 2) connectors.push([...path[i], ...path[i + 1]])
    stroke()
  }
  ctx.fillText = (text: string, x: number, y: number, max?: number) => { texts.push({ text, x, y, align: ctx.textAlign }); fillText(text, x, y, max) }
  paintXlsxChart(model, ctx, W, H)
  const plot = plotOf(model)
  // The marks: rectangles inside the plot (the chart and plot backgrounds and legend swatches excluded).
  const marks = rects.filter(r => r.x >= plot.x - .5 && r.x + r.w <= plot.x + plot.w + .5 && r.y >= plot.y - .5 && r.y + r.h <= plot.y + plot.h + .5 && !(r.w >= plot.w - .5 && r.h >= plot.h - .5))
  return { rects, marks, connectors, texts, plot }
}
const close = (a: number, b: number, digits = 6) => expect(a).toBeCloseTo(b, digits)

describe('ChartEx waterfall and funnel (chartex-map §§6, 9)', () => {
  it('subtotals show the running total anchored at zero; the accumulator continues', async () => {
    // Oracle (synthetic; the subtotal's own value 99 differs from the running total):
    //   values 10, -3, 5, [99 as subtotal], 4
    //   acc: 0 → 10 (increase [0,10]) → 7 (decrease [10,7]) → 12 (increase [7,12]) → total [0,12] (acc stays 12)
    //        → 16 (increase [12,16]).
    const bars: Array<[WaterfallKind, number, number]> = []
    const points: ChartPoint[] = [10, -3, 5, 99, 4].map(v => ({ state: 'value', v }))
    forEachWaterfallBar(5, slot => slot, slot => points[slot], [3], (_slot, _i, kind, from, to) => bars.push([kind, from, to]))
    expect(bars).toEqual([['increase', 0, 10], ['decrease', 10, 7], ['increase', 7, 12], ['total', 0, 12], ['increase', 12, 16]])
    // Blank, error and hidden points draw no bar and contribute 0; a subtotal still shows the running total.
    const mixed: Array<ChartPoint | undefined> = [{ state: 'value', v: 2 }, { state: 'blank' }, { state: 'error', code: '#DIV/0!' }, undefined, { state: 'unavailable' }, { state: 'value', v: 3 }, { state: 'value', v: 0 }]
    const got: Array<[number, WaterfallKind, number, number]> = [], skipped: Array<[number, string | undefined]> = []
    forEachWaterfallBar(7, slot => slot, slot => mixed[slot], [6], (slot, _i, kind, from, to) => got.push([slot, kind, from, to]), (slot, state) => skipped.push([slot, state]))
    expect(got).toEqual([[0, 'increase', 0, 2], [5, 'increase', 2, 5], [6, 'total', 0, 5]])
    expect(skipped).toEqual([[1, 'blank'], [2, 'error'], [3, undefined], [4, 'unavailable']])

    // The same oracle through the adapter, stage 1, layout and paint (Office-shaped synthetic part).
    const model = await native(waterfall([10, -3, 5, 99, 4], '<cx:subtotals><cx:idx val="3" /></cx:subtotals>'))
    const group = model.groups[0] as ChartExGroup
    expect(group.kind).toBe('chartEx')
    expect(group.family).toBe('waterfall')
    const levels = waterfallLevels(group, displayView(model), 0)
    expect([...levels.from]).toEqual([0, 10, 7, 0, 12])
    expect([...levels.to]).toEqual([10, 7, 12, 12, 16])
    // The value axis spans the running totals (0..16), not the subtotal's own 99.
    const value = model.preparedAxes!.axes.find(a => a.role === 'value')!
    expect(value.min).toBe(0)
    expect(value.max).toBeGreaterThanOrEqual(16)
    expect(value.max).toBeLessThan(99)
    const { marks } = record(model)
    expect(marks).toHaveLength(5)
    const [b0, b1, b2, total, b4] = [...marks].sort((p, q) => p.x - q.x)
    const bottom = (r: Rect) => r.y + r.h
    // Anchored at zero: the total's base is the first bar's base (y(0)); its top is the running total 12 (b2's top).
    close(bottom(total), bottom(b0))
    close(total.y, b2.y)
    // Heights are proportional to the drawn intervals: |[0,12]| / |[0,10]| = 1.2; |[10,7]| / |[0,10]| = 0.3.
    close(total.h / b0.h, 1.2)
    close(b1.h / b0.h, 0.3)
    // The accumulator continues from the total: the next bar floats from 12 (the total's top) to 16.
    close(bottom(b4), total.y)
    close(b4.h / b0.h, 0.4)

    // Native WF1 (Excel; no subtotals): values 2,5,4,6,-5,-6,-4,7,8,6,7,5 read through its _xlchart defined name.
    // Running totals 2,7,11,17,12,6,2,9,17,23,30,35.
    const xlsx = await parseXlsx(await OfficePackage.load(fixture('waterfall/waterfall.xlsx')))
    const drawing = xlsx.sheets[0].drawings![0]
    expect(drawing.source.representation, 'the mc:Choice is selected, not the fallback').toBe('choice')
    expect((xlsx.drawingCoverage ?? []).filter(e => e.scope === 'original')[0]).toMatchObject({ status: 'native', selectedRepresentation: 'chart', representation: 'choice' })
    expect(drawing.content?.kind).toBe('xlsxChart')
    if (drawing.content?.kind !== 'xlsxChart') return
    const wf1 = prepareChartAxes(drawing.content.chart)
    expect(wf1.unsupported).toBeUndefined()
    expect(isChartExGroup(wf1.groups[0]) && wf1.groups[0].family).toBe('waterfall')
    const wf1Levels = waterfallLevels(wf1.groups[0] as ChartExGroup, displayView(wf1), 0)
    expect([...wf1Levels.to]).toEqual([2, 7, 11, 17, 12, 6, 2, 9, 17, 23, 30, 35])
    expect([...wf1Levels.from]).toEqual([0, 2, 7, 11, 17, 12, 6, 2, 9, 17, 23, 30])
    expect(wf1.preparedAxes!.axes.find(a => a.role === 'value')!.max).toBeGreaterThanOrEqual(35)
    // Ordinal category labels (no cat dimension; map §2.3 [P]).
    expect(wf1.preparedAxes!.axes.find(a => a.role === 'category')!.labels).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'])
    const wf1Marks = record(wf1).marks
    expect(wf1Marks).toHaveLength(12)
    // The last bar spans 30..35 and its top is the highest mark.
    const last = [...wf1Marks].sort((p, q) => p.x - q.x)[11]
    expect(Math.min(...wf1Marks.map(r => r.y))).toBeCloseTo(last.y, 6)
  })

  it('increase, decrease and total colours; connector visibility', async () => {
    // Synthetic: 10, -3, 5, [0 as subtotal], 4; point 4 has a red dataPt override.
    const override = '<cx:dataPt idx="4"><cx:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></cx:spPr></cx:dataPt>'
    const model = await native(waterfall([10, -3, 5, 0, 4], '<cx:visibility connectorLines="1" /><cx:subtotals><cx:idx val="3" /></cx:subtotals>', override))
    const drawn = record(model)
    const bars = [...drawn.marks].sort((p, q) => p.x - q.x)
    // Increase accent 1, decrease accent 2, total accent 3 (the style palette, map §6 [U]); dataPt overrides one bar.
    expect(bars.map(b => b.color)).toEqual([ACCENT[0], ACCENT[1], ACCENT[0], ACCENT[2], '#ff0000'])
    // The legend lists the three kinds in their colours.
    for (const [i, name] of ['Increase', 'Decrease', 'Total'].entries()) {
      const text = drawn.texts.find(t => t.text === name)
      expect(text, name).toBeDefined()
      expect(drawn.rects.some(r => r.color === ACCENT[i] && Math.abs(r.y + r.h / 2 - text!.y) < 2 && r.x < text!.x)).toBe(true)
    }
    // Connectors: one per adjacent pair, horizontal at the level after the left bar, from its right edge to the next
    // bar's left edge. Levels after each bar: 10, 7, 12, 12 (the total keeps the running total).
    expect(drawn.connectors).toHaveLength(4)
    const tops = { b0: bars[0].y, b1bottom: bars[1].y + bars[1].h, b2: bars[2].y, total: bars[3].y }
    const levels = [tops.b0, tops.b1bottom, tops.b2, tops.total]
    drawn.connectors.forEach(([x0, y0, x1, y1], i) => {
      close(y0, y1)
      close(y0, levels[i])
      close(x0, bars[i].x + bars[i].w)
      close(x1, bars[i + 1].x)
    })
    // connectorLines="0" hides them; an absent flag shows them (map §4.1 default 1 [U]); WF1 omits it: 11 connectors.
    expect(record(await native(waterfall([10, -3, 5], '<cx:visibility connectorLines="0" />'))).connectors).toHaveLength(0)
    expect(record(await native(waterfall([10, -3, 5], ''))).connectors).toHaveLength(2)
    const wf1 = await parseXlsx(await OfficePackage.load(fixture('waterfall/waterfall.xlsx')))
    const content = wf1.sheets[0].drawings![0].content
    if (content?.kind !== 'xlsxChart') throw new Error('WF1 chart')
    const wf1Drawn = record(prepareChartAxes(content.chart))
    expect(wf1Drawn.connectors).toHaveLength(11)
    // WF1: values 2,5,4,6 increase; -5,-6,-4 decrease; 7,8,6,7,5 increase.
    expect([...wf1Drawn.marks].sort((p, q) => p.x - q.x).map(b => b.color)).toEqual([0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0].map(k => ACCENT[k]))
    // WF2 (Excel; connectorLines="1" authored, 16 values): 15 connectors, native.
    const wf2 = await parseXlsx(await OfficePackage.load(fixture('waterfall2/waterfall2.xlsx')))
    const content2 = wf2.sheets[0].drawings![0].content
    if (content2?.kind !== 'xlsxChart') throw new Error('WF2 chart')
    expect(content2.chart.unsupported).toBeUndefined()
    const wf2Drawn = record(prepareChartAxes(content2.chart))
    expect(wf2Drawn.marks).toHaveLength(16)
    expect(wf2Drawn.connectors).toHaveLength(15)
  })

  it('funnel centred bars proportional to value in source order', async () => {
    // Oracle (pure geometry): values 4, 2, 0, -1, blank in a 100 x 50 plot at (0,0) with gapWidth 100 %:
    //   band 50/5 = 10, bar thickness 10/(1+1) = 5, centred in its band (y = 10i + 2.5); max 4;
    //   widths 100·4/4 = 100, 100·2/4 = 50, 0, 0 (negative), none; each centred on x = 50.
    const points: ChartPoint[] = [{ state: 'value', v: 4 }, { state: 'value', v: 2 }, { state: 'value', v: 0 }, { state: 'value', v: -1 }, { state: 'blank' }]
    const geometry = funnelGeometry(5, slot => points[slot], { x: 0, y: 0, w: 100, h: 50 }, 100)
    expect(geometry.max).toBe(4)
    expect([0, 1, 2, 3, 4].map(i => geometry.bar(i))).toEqual([
      { x: 0, y: 2.5, w: 100, h: 5 }, { x: 25, y: 12.5, w: 50, h: 5 }, { x: 50, y: 22.5, w: 0, h: 5 }, { x: 50, y: 32.5, w: 0, h: 5 }, undefined,
    ])

    // Native FN2 (PowerPoint): series 1 visible (4.3, 2.5, 3.5, 4.5), series 2 and 3 hidden. End to end: native
    // model, coverage chart, distinct centred bars of widths 4.3/4.5, 2.5/4.5, 3.5/4.5, 1 of the plot, top to bottom.
    const pptx = await parsePptx(await OfficePackage.load(fixture('funnel-pp1/funnel-pp1.pptx')))
    const coverage = (pptx.drawingCoverage ?? []).filter(e => e.scope === 'original')
    expect(coverage[0]).toMatchObject({ status: 'native', selectedRepresentation: 'chart', representation: 'choice' })
    const shape = pptx.slides[0].shapes.find(s => s.content?.kind === 'xlsxChart')
    expect(shape, 'the chart, not the fallback picture').toBeDefined()
    expect(pptx.slides[0].shapes.filter(s => s.imageIndex !== undefined)).toHaveLength(0)
    if (shape?.content?.kind !== 'xlsxChart') return
    const fn2 = prepareChartAxes(shape.content.chart)
    expect(fn2.unsupported).toBeUndefined()
    expect(fn2.groups[0].family).toBe('funnel')
    const drawn = record(fn2)
    const bars = drawn.marks
    expect(bars, 'only the visible series is drawn').toHaveLength(4)
    expect(new Set(bars.map(b => b.color))).toEqual(new Set([ACCENT[0]]))
    // Source order from top to bottom.
    expect(bars.map(b => b.y)).toEqual([...bars.map(b => b.y)].sort((p, q) => p - q))
    const widest = Math.max(...bars.map(b => b.w))
    close(widest, drawn.plot.w)
    ;[4.3, 2.5, 3.5, 4.5].forEach((v, i) => {
      close(bars[i].w / widest, v / 4.5)
      close(bars[i].x + bars[i].w / 2, drawn.plot.x + drawn.plot.w / 2)
    })
    // Thickness from catScaling@gapWidth 2.19: band / (1 + 2.19).
    close(bars[0].h, drawn.plot.h / 4 / (1 + 2.19), 2)
    // The category labels sit left of the plot, right-aligned, one per band.
    for (const [i, name] of ['Thing 1', 'Thing 2', 'Thing 3', 'Thing 4'].entries()) {
      const label = drawn.texts.find(t => t.text === name && t.align === 'right')
      expect(label, name).toBeDefined()
      expect(label!.x).toBeLessThan(drawn.plot.x)
      close(label!.y, bars[i].y + bars[i].h / 2, 4)
    }

    // Native FN1 (Excel; values 6,3,4,8,9,6,4,6,4,5,3,2 via _xlchart names): 12 bars, widths v/9 of the widest.
    const xlsx = await parseXlsx(await OfficePackage.load(fixture('funnel-label-options/funnel-label-options.xlsx')))
    expect((xlsx.drawingCoverage ?? []).filter(e => e.scope === 'original')[0]).toMatchObject({ status: 'native', selectedRepresentation: 'chart', representation: 'choice' })
    const content = xlsx.sheets[0].drawings![0].content
    if (content?.kind !== 'xlsxChart') throw new Error('FN1 chart')
    const fn1 = prepareChartAxes(content.chart)
    expect(fn1.unsupported).toBeUndefined()
    const fn1Bars = record(fn1).marks
    expect(fn1Bars).toHaveLength(12)
    const max = Math.max(...fn1Bars.map(b => b.w))
    ;[6, 3, 4, 8, 9, 6, 4, 6, 4, 5, 3, 2].forEach((v, i) => close(fn1Bars[i].w / max, v / 9))
    // FN1's value labels are placed on the bars (label 2 "gamma, 4" at the centre of bar 2: pos ctr).
    const fn1Texts = record(fn1).texts
    const label = fn1Texts.find(t => t.text.startsWith('gamma'))
    expect(label, 'gamma label').toBeDefined()
    close(label!.y, fn1Bars[2].y + fn1Bars[2].h / 2, 4)
    // Its category labels come from the series' filled categories (not ordinals), left of the plot.
    expect(fn1Texts.filter(t => t.align === 'right').map(t => t.text)).toEqual(['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta', 'iota', 'kappa', 'lambda', 'mu'])

    // Several visible series are unsupported (map §9): the fallback picture is kept.
    expect(await adapt(funnel([{ values: [1, 2] }, { values: [3, 4] }]))).toMatchObject({ kind: 'unsupported', reason: CHARTEX_SEVERAL_VISIBLE, feature: 'ChartEx' })
    // A hidden second series is fine.
    expect((await adapt(funnel([{ values: [1, 2] }, { values: [3, 4], hidden: true }]))).kind).toBe('native')
  })
})

/** The conformance cases (spec 7) as Office-shaped synthetic ChartEx parts: the harness's numeric cache becomes a level. */
const cacheLevel = (cache: string): string => {
  const count = Number(/<c:ptCount val="(\d+)"/.exec(cache)?.[1] ?? 0)
  const pts = [...cache.matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v><\/c:pt>/g)].map(m => `<cx:pt idx="${m[1]}">${m[2]}</cx:pt>`).join('')
  return `<cx:lvl ptCount="${count}">${pts}</cx:lvl>`
}
const chartExFamily = (layout: 'waterfall' | 'funnel'): ConformanceFamily => ({
  name: `chartex ${layout}`,
  chart: cache => space(`<cx:chartData>${cache === null ? '' : `<cx:data id="0"><cx:numDim type="val">${cacheLevel(cache)}</cx:numDim></cx:data>`}</cx:chartData>` +
    `<cx:chart><cx:plotArea><cx:plotAreaRegion>${cache === null ? '' : `<cx:series layoutId="${layout}"><cx:dataLabels pos="ctr"><cx:visibility value="1" /></cx:dataLabels><cx:dataId val="0" /></cx:series>`}</cx:plotAreaRegion>` +
    `${layout === 'waterfall' ? '<cx:axis id="0"><cx:catScaling /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:tickLabels /></cx:axis>' : '<cx:axis id="1"><cx:catScaling /><cx:tickLabels /></cx:axis>'}</cx:plotArea><cx:legend pos="r" /></cx:chart>`),
  parse: (xml, partPath) => adapt(xml, partPath.replace(/^\//, '').replace(/chart1\.xml$/, 'chartEx1.xml')),
  expect: { 'no series': { kind: 'card' }, 'empty series': { kind: 'native' }, 'all blank': { kind: 'native' }, 'NaN and infinity': { kind: 'native' }, 'negative values': { kind: 'native' }, 'single point': { kind: 'native' }, 'unavailable point': { kind: 'native' } },
})
describeFamilyConformance(chartExFamily('waterfall'))
describeFamilyConformance(chartExFamily('funnel'))
