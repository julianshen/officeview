/**
 * Task B3: histogram and Pareto ChartEx families (chartex-map §5; spec 4.3).
 *
 * Native evidence: HI1 (SimpleHistogram), HI2 (tdf163727 underflow/overflow), PA1 (paretoLine) and PA2
 * (pareto-gapwidth), Office-authored parts in tests/fixtures/chartex (PROVENANCE.md), read end to end through
 * `parseXlsx`. Every other part here is Office-shaped synthetic (never Office evidence). Expected bins, counts and
 * cumulative fractions are hand-computed in the comments beside each assertion.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContext, chartParseContextFor, documentChartContext } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { CHARTEX_PAINTED, parseChartEx } from '../src/charts/chartex/parse'
import { fillWorkbookCharts } from '../src/charts/fill'
import { releaseChartModel } from '../src/charts/own'
import { paintXlsxChart } from '../src/charts/paint'
import { pointList } from '../src/charts/points'
import type { ChartModel, ChartResult } from '../src/charts/model'
import { parseXlsx } from '../src/xlsx/parse'
import { buildChartPptx, cell, embeddedWorkbook, row } from './helpers/chart-fixtures'
import { describeFamilyConformance } from './helpers/chart-family-conformance'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const theme = parseThemeContext()
const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'chartex', name)))
const space = (body: string) => `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</cx:chartSpace>`

/** The column and value axes Office writes for a histogram (HI1, HI2), plus PA1's percent axis for a Pareto. */
const AXES = '<cx:axis id="0"><cx:catScaling gapWidth="0" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis>'
const PERCENT_AXIS = '<cx:axis id="2"><cx:valScaling max="1" min="0" /><cx:units unit="percentage" /><cx:tickLabels /></cx:axis>'
interface Spec {
  values: Array<number | string | undefined>
  /** `cx:layoutPr` content (default: automatic binning, closure r). */
  layoutPr?: string
  categories?: Array<string | undefined>
  pareto?: boolean
  axes?: string
  /** Extra column series (same layoutPr), each with its own values. */
  more?: Array<Array<number | string | undefined>>
}
const numLevel = (values: Spec['values']) => `<cx:lvl ptCount="${values.length}">${values.map((v, i) => v === undefined ? '' : `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`
/** An Office-shaped synthetic histogram (or Pareto) part with cached data. */
function part(spec: Spec): string {
  const layoutPr = spec.layoutPr ?? '<cx:binning intervalClosed="r" />'
  const cat = spec.categories ? `<cx:strDim type="cat"><cx:lvl ptCount="${spec.categories.length}">${spec.categories.map((c, i) => c === undefined ? '' : `<cx:pt idx="${i}">${c}</cx:pt>`).join('')}</cx:lvl></cx:strDim>` : ''
  const sets = [spec.values, ...spec.more ?? []]
  const data = sets.map((values, i) => `<cx:data id="${i}">${cat}<cx:numDim type="val">${numLevel(values)}</cx:numDim></cx:data>`).join('')
  const columns = sets.map((_, i) => `<cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:v>S${i + 1}</cx:v></cx:txData></cx:tx><cx:dataId val="${i}" /><cx:layoutPr>${layoutPr}</cx:layoutPr><cx:axisId val="1" /></cx:series>`).join('')
  const line = spec.pareto ? '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2" /></cx:series>' : ''
  return space(`<cx:chartData>${data}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${columns}${line}</cx:plotAreaRegion>${spec.axes ?? AXES + (spec.pareto ? PERCENT_AXIS : '')}</cx:plotArea></cx:chart>`)
}
async function readPart(pkg: OfficePackage, path: string, context = chartParseContextFor(pkg, path)) {
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error(`${path}: ${read ? read.reason : 'missing'}`)
  return { root: read.root, context, release: read.release }
}
/** The part through the production adapter (default painted families). */
async function adapt(xml: string, path = 'ppt/charts/chartEx1.xml'): Promise<ChartResult> {
  const zip = new JSZip()
  zip.file(path, xml)
  const p = await readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), path)
  try { return await parseChartPart(p.root, theme, p.context) } finally { p.release() }
}
async function native(spec: Spec | string): Promise<ChartModel> {
  const result = await adapt(typeof spec === 'string' ? spec : part(spec))
  if (result.kind !== 'native') throw new Error(`not native: ${result.kind} ${result.reason}`)
  return result.model
}
const nums = (points: ChartModel['groups'][number]['series'][number]['values']) => pointList(points).map(p => p.state === 'value' ? p.v : p.state)
/** The drawn bars: the derived column group's labels and per-series heights. */
const bars = (model: ChartModel, series = 0) => ({ labels: model.categories, counts: nums(model.groups[0].series[series].values) })

/** A recording canvas: filled rectangles and stroked polylines with their colours. */
function record(model: ChartModel, w = 480, h = 320) {
  const ctx = createCanvas(w, h).getContext('2d')
  const rects: Array<{ x: number; y: number; w: number; h: number; fill: string }> = []
  const paths: Array<{ stroke: string; points: Array<[number, number]> }> = []
  let current: Array<[number, number]> = []
  const wrap = <K extends 'fillRect' | 'beginPath' | 'moveTo' | 'lineTo' | 'stroke'>(name: K, before: (...args: number[]) => void) => {
    const original = (ctx[name] as (...args: number[]) => void).bind(ctx)
    ;(ctx as unknown as Record<string, unknown>)[name] = (...args: number[]) => { before(...args); original(...args) }
  }
  wrap('fillRect', (x, y, rw, rh) => rects.push({ x, y, w: rw, h: rh, fill: String(ctx.fillStyle).toLowerCase() }))
  wrap('beginPath', () => { current = [] })
  wrap('moveTo', (x, y) => current.push([x, y]))
  wrap('lineTo', (x, y) => current.push([x, y]))
  wrap('stroke', () => paths.push({ stroke: String(ctx.strokeStyle).toLowerCase(), points: current.slice() }))
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  return { rects, paths }
}
/** The painted bars of a series colour, left to right (zero-height bars included). */
const paintedBars = (model: ChartModel, color: string) => record(model).rects.filter(r => r.fill === `#${color.toLowerCase()}` && r.w > 0 && r.w < 400).sort((a, b) => a.x - b.x)
/** Native success (spec 6.3): an XLSX drawing whose chart is native, selected over the fallback, covered as `chart`. */
async function nativeXlsx(file: string): Promise<ChartModel> {
  const xlsx = await parseXlsx(await OfficePackage.load(fixture(file)))
  const sheet = xlsx.sheets[0]
  const content = sheet.drawings?.[0].content
  expect(content?.kind, `${file}: a native chart, not the fallback`).toBe('xlsxChart')
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  expect(content.chart.unsupported, file).toBeUndefined()
  expect(sheet.drawings?.[0].source.representation, file).toBe('choice')
  expect(sheet.drawingCoverage?.find(entry => entry.element === 'graphicFrame'), file).toMatchObject({ selectedRepresentation: 'chart' })
  expect(content.chart.preparedAxes, `${file}: stage 1 ran`).toBeDefined()
  return content.chart
}
/** General (spec 5.2): 15 significant digits. */
const general = (n: number) => String(Number(n.toPrecision(15)))

describe('histogram and Pareto (chartex-map §5)', () => {
  it('binSize, binCount and automatic Scott bins', async () => {
    expect([...CHARTEX_PAINTED]).toEqual(expect.arrayContaining(['histogram', 'pareto']))
    // HI2 (native): 12, 10, 11, 13, 15; binSize 2, underflow 11, overflow 14, closure r (map §5.1 worked example).
    // Underflow <=11 = {10, 11} = 2; overflow >14 = {15} = 1; inner {12, 13}, lo = 11, hi = 14, k = ceil(3/2) = 2:
    // (11,13] = {12, 13} = 2, (13,15] = {} = 0.
    const hi2 = await nativeXlsx('tdf163727_histogram_underflow_overflow/tdf163727_histogram_underflow_overflow.xlsx')
    expect(bars(hi2)).toEqual({ labels: ['≤11', '(11, 13]', '(13, 15]', '>14'], counts: [2, 2, 0, 1] })
    // Distinct native geometry: four contiguous columns (gapWidth 0) with heights in the ratio 2:2:0:1.
    const drawn = paintedBars(hi2, hi2.groups[0].series[0].color)
    expect(drawn).toHaveLength(4)
    expect(drawn[0].h / drawn[3].h).toBeCloseTo(2, 6)
    expect(drawn[1].h).toBeCloseTo(drawn[0].h, 6)
    expect(drawn[2].h).toBe(0)
    for (let i = 1; i < 4; i++) expect(drawn[i].x).toBeCloseTo(drawn[i - 1].x + drawn[i - 1].w, 6)

    // HI1 (native): 12, 10, 11, 13, 15, automatic. Scott: mean 12.2, sum of squares 14.8, sample variance 3.7,
    // h = 3.49 * sqrt(3.7) * 5^(-1/3) = 3.92587..., k = ceil((15 - 10) / h) = 2:
    // [10, 13.9259] = {12, 10, 11, 13} = 4, (13.9259, 17.8517] = {15} = 1.
    const h1 = 3.49 * Math.sqrt(3.7) * 5 ** (-1 / 3)
    expect(h1).toBeCloseTo(3.925873, 5)
    const hi1 = await nativeXlsx('SimpleHistogram/SimpleHistogram.xlsx')
    expect(bars(hi1)).toEqual({ labels: [`[10, ${general(10 + h1)}]`, `(${general(10 + h1)}, ${general(10 + 2 * h1)}]`], counts: [4, 1] })
    expect(paintedBars(hi1, hi1.groups[0].series[0].color)).toHaveLength(2)

    // binCount 4 over 1..10 (no outer bins): h = 9/4 = 2.25, edges 1, 3.25, 5.5, 7.75, 10.
    const ten = Array.from({ length: 10 }, (_, i) => i + 1)
    expect(bars(await native({ values: ten, layoutPr: '<cx:binning intervalClosed="r"><cx:binCount val="4" /></cx:binning>' })))
      .toEqual({ labels: ['[1, 3.25]', '(3.25, 5.5]', '(5.5, 7.75]', '(7.75, 10]'], counts: [3, 2, 2, 3] })
    // binCount includes the outer bins [D-hist]: 4 - 2 = 2 inner bins over (2, 8], h = 3.
    expect(bars(await native({ values: ten, layoutPr: '<cx:binning intervalClosed="r" underflow="2" overflow="8"><cx:binCount val="4" /></cx:binning>' })))
      .toEqual({ labels: ['≤2', '(2, 5]', '(5, 8]', '>8'], counts: [2, 3, 3, 2] })

    // Exactly 1,000 bins is drawn; 1,001 (outer bins included) and a binSize-derived 2,000 are refused.
    const thousand = await native({ values: [0, 1000], layoutPr: '<cx:binning><cx:binSize val="1" /></cx:binning>' })
    expect(thousand.categories).toHaveLength(1000)
    expect(await adapt(part({ values: [0, 1000], layoutPr: '<cx:binning underflow="0"><cx:binSize val="1" /></cx:binning>' }))).toMatchObject({ kind: 'refused', reason: 'invalid binning', feature: 'ChartEx' })
    expect(await adapt(part({ values: [0, 2000], layoutPr: '<cx:binning><cx:binSize val="1" /></cx:binning>' }))).toMatchObject({ kind: 'refused', reason: 'invalid binning' })
    // Scott's h is undefined (one inner value, n - 1 = 0) while hi > lo (underflow 2 < 5): refused.
    expect(await adapt(part({ values: [1, 5], layoutPr: '<cx:binning underflow="2" />' }))).toMatchObject({ kind: 'refused', reason: 'invalid binning' })
    // hi = lo: one inner bin [lo, lo], whatever the mode.
    expect(bars(await native({ values: [7, 7, 7] }))).toEqual({ labels: ['[7, 7]'], counts: [3] })
    expect(bars(await native({ values: [7, 7], layoutPr: '<cx:binning><cx:binCount val="3" /></cx:binning>' }))).toEqual({ labels: ['[7, 7]'], counts: [2] })
  })

  it('intervalClosed right and left', async () => {
    // 0..4, binSize 2: r gives [0,2] {0,1,2} and (2,4] {3,4}; l gives [0,2) {0,1} and [2,4] {2,3,4} (the last bin
    // closed when no overflow is set).
    const five = [0, 1, 2, 3, 4]
    expect(bars(await native({ values: five, layoutPr: '<cx:binning intervalClosed="r"><cx:binSize val="2" /></cx:binning>' }))).toEqual({ labels: ['[0, 2]', '(2, 4]'], counts: [3, 2] })
    expect(bars(await native({ values: five, layoutPr: '<cx:binning intervalClosed="l"><cx:binSize val="2" /></cx:binning>' }))).toEqual({ labels: ['[0, 2)', '[2, 4]'], counts: [2, 3] })
    // The default closure is r.
    expect(bars(await native({ values: five, layoutPr: '<cx:binning><cx:binSize val="2" /></cx:binning>' }))).toEqual({ labels: ['[0, 2]', '(2, 4]'], counts: [3, 2] })
    // Outer bins mirror the closure so the bins stay disjoint: 1..6, underflow 2, overflow 5, binSize 1.
    // r: <=2 {1,2}; (2,3] {3}; (3,4] {4}; (4,5] {5}; >5 {6}.  l: <2 {1}; [2,3) {2}; [3,4) {3}; [4,5) {4}; >=5 {5,6}.
    const six = [1, 2, 3, 4, 5, 6]
    expect(bars(await native({ values: six, layoutPr: '<cx:binning intervalClosed="r" underflow="2" overflow="5"><cx:binSize val="1" /></cx:binning>' })))
      .toEqual({ labels: ['≤2', '(2, 3]', '(3, 4]', '(4, 5]', '>5'], counts: [2, 1, 1, 1, 1] })
    expect(bars(await native({ values: six, layoutPr: '<cx:binning intervalClosed="l" underflow="2" overflow="5"><cx:binSize val="1" /></cx:binning>' })))
      .toEqual({ labels: ['<2', '[2, 3)', '[3, 4)', '[4, 5)', '≥5'], counts: [1, 1, 1, 1, 2] })
    // Membership is decided against the same edges the labels show: 0.1, 0.2, 0.3 with binSize 0.1
    // (0.3 - 0.1 = 0.19999999999999998, k = 2; edges 0.1, 0.2, 0.30000000000000004): [0.1, 0.2] {0.1, 0.2}, (0.2, 0.3] {0.3}.
    expect(bars(await native({ values: [0.1, 0.2, 0.3], layoutPr: '<cx:binning><cx:binSize val="0.1" /></cx:binning>' }))).toEqual({ labels: ['[0.1, 0.2]', '(0.2, 0.3]'], counts: [2, 1] })
  })

  it('underflow and overflow, explicit and auto', async () => {
    // Underflow only: 1..6, underflow 3, binSize 1.5: <=3 {1,2,3}; inner lo = 3, hi = 6, k = 2: (3,4.5] {4}, (4.5,6] {5,6}.
    const six = [1, 2, 3, 4, 5, 6]
    expect(bars(await native({ values: six, layoutPr: '<cx:binning underflow="3"><cx:binSize val="1.5" /></cx:binning>' })))
      .toEqual({ labels: ['≤3', '(3, 4.5]', '(4.5, 6]'], counts: [3, 1, 2] })
    // Overflow only: overflow 4: >4 {5,6}; inner lo = 1, hi = 4, k = 2: [1,2.5] {1,2}, (2.5,4] {3,4}.
    expect(bars(await native({ values: six, layoutPr: '<cx:binning overflow="4"><cx:binSize val="1.5" /></cx:binning>' })))
      .toEqual({ labels: ['[1, 2.5]', '(2.5, 4]', '>4'], counts: [2, 2, 2] })
    // An empty inner set draws no inner bins; the outer bins stay.
    expect(bars(await native({ values: [1, 9], layoutPr: '<cx:binning underflow="2" overflow="8"><cx:binSize val="1" /></cx:binning>' })))
      .toEqual({ labels: ['≤2', '>8'], counts: [1, 1] })
    // Blank and error points are excluded; every value point is counted once.
    expect(bars(await native({ values: [1, undefined, '#N/A', 2, 3], layoutPr: '<cx:binning><cx:binSize val="1" /></cx:binning>' })))
      .toEqual({ labels: ['[1, 2]', '(2, 3]'], counts: [2, 1] })
    // underflow >= overflow (closure r) is invalid binning, refused; automatic outer bins are unsupported.
    expect(await adapt(part({ values: six, layoutPr: '<cx:binning underflow="4" overflow="4" />' }))).toMatchObject({ kind: 'refused', reason: 'invalid binning' })
    expect(await adapt(part({ values: six, layoutPr: '<cx:binning underflow="auto" />' }))).toMatchObject({ kind: 'unsupported', reason: 'Automatic histogram overflow and underflow bins are not supported', feature: 'ChartEx' })
    expect(await adapt(part({ values: six, layoutPr: '<cx:binning overflow="auto" />' }))).toMatchObject({ kind: 'unsupported', feature: 'ChartEx' })
  })

  it('interval label notation', async () => {
    // Office interval notation: inner (a, b] / [a, b) and the closed first (r) or last (l) bin; outer <=U / >O (r) and
    // <U / >=O (l). Edges are formatted with the category axis number format when one is authored, else General.
    const axes = '<cx:axis id="0"><cx:catScaling gapWidth="0" /><cx:numFmt formatCode="0.00" sourceLinked="0" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:tickLabels /></cx:axis>'
    expect((await native({ values: [1, 2, 3, 4], axes, layoutPr: '<cx:binning underflow="1.5" overflow="3.5"><cx:binSize val="1" /></cx:binning>' })).categories)
      .toEqual(['≤1.50', '(1.50, 2.50]', '(2.50, 3.50]', '>3.50'])
    expect((await native({ values: [1, 2, 3, 4], layoutPr: '<cx:binning intervalClosed="l" underflow="1.5" overflow="3.5"><cx:binSize val="1" /></cx:binning>' })).categories)
      .toEqual(['<1.5', '[1.5, 2.5)', '[2.5, 3.5)', '≥3.5'])
    // The labels are what the category axis draws (stage 1).
    const model = await native({ values: [1, 2, 3, 4], layoutPr: '<cx:binning><cx:binSize val="1.5" /></cx:binning>' })
    expect(model.preparedAxes?.axes.find(a => a.id === '0')?.labels).toEqual(['[1, 2.5]', '(2.5, 4]'])
  })

  it('pre-aggregated data drawn as given', async () => {
    // By Category (aggregation): already-unique categories are drawn as given, in source order.
    const given = await native({ values: [5, 3, 8], categories: ['A', 'B', 'C'], layoutPr: '<cx:aggregation />' })
    expect(bars(given)).toEqual({ labels: ['A', 'B', 'C'], counts: [5, 3, 8] })
    const drawn = paintedBars(given, given.groups[0].series[0].color)
    expect(drawn.map(r => r.h / drawn[1].h).map(v => Number(v.toFixed(6)))).toEqual([5 / 3, 1, 8 / 3].map(v => Number(v.toFixed(6))))
    // Equal categories group by first occurrence and sum (map §5.1, controller ruling): (A,2), (A,3), (B,1) -> A 5, B 1.
    expect(bars(await native({ values: [2, 3, 1], categories: ['A', 'A', 'B'], layoutPr: '<cx:aggregation />' }))).toEqual({ labels: ['A', 'B'], counts: [5, 1] })
    // Blank and error values contribute nothing; their category still appears (at 0 when it has no value).
    expect(bars(await native({ values: [2, undefined, '#DIV/0!', 4], categories: ['A', 'C', 'A', 'B'], layoutPr: '<cx:aggregation />' }))).toEqual({ labels: ['A', 'C', 'B'], counts: [2, 0, 4] })
    // A By Category histogram without a category dimension is unsupported (the fallback is kept).
    expect(await adapt(part({ values: [1, 2], layoutPr: '<cx:aggregation />' }))).toMatchObject({ kind: 'unsupported', feature: 'ChartEx' })
  })

  it('pareto descending with stable ties; cumulative percent on its own axis', async () => {
    // By category with ties: (A,2), (B,5), (C,2), (D,1) -> B 5, A 2, C 2 (A before C: source order), D 1; total 10:
    // cumulative 0.5, 0.7, 0.9, 1.
    const model = await native({ values: [2, 5, 2, 1], categories: ['A', 'B', 'C', 'D'], layoutPr: '<cx:aggregation />', pareto: true })
    expect(bars(model)).toEqual({ labels: ['B', 'A', 'C', 'D'], counts: [5, 2, 2, 1] })
    const line = model.groups[1]
    expect(line.family).toBe('line')
    expect(nums(line.series[0].values)).toEqual([0.5, 0.7, 0.9, 1])
    // Its own axis: the paretoLine's axisId (2), 0..1, percent ticks, on the right.
    expect(line.binding).toMatchObject({ kind: 'cartesian', value: '2' })
    const percent = model.preparedAxes?.axes.find(a => a.id === '2')
    expect(percent).toMatchObject({ min: 0, max: 1, side: 'r', role: 'value' })
    expect(percent?.labels).toEqual(['0%', '20%', '40%', '60%', '80%', '100%'])
    expect(model.preparedAxes?.axes.find(a => a.id === '1')).toMatchObject({ side: 'l', min: 0 })
    // Geometry: four bars 5:2:2:1 and a polyline through the band centres at the cumulative fractions.
    const { rects, paths } = record(model)
    const columns = rects.filter(r => r.fill === `#${model.groups[0].series[0].color.toLowerCase()}` && r.w < 400).sort((a, b) => a.x - b.x)
    expect(columns.map(r => Number((r.h / columns[3].h).toFixed(6)))).toEqual([5, 2, 2, 1])
    const polyline = paths.find(p => p.stroke === `#${line.series[0].color.toLowerCase()}` && p.points.length === 4)
    expect(polyline, 'the cumulative line').toBeDefined()
    const ys = polyline!.points.map(p => p[1]), xs = polyline!.points.map(p => p[0])
    // y = top + (1 - f) * height, so (y_i - y_3) / (y_0 - y_3) = (1 - f_i) / (1 - f_0).
    expect((ys[1] - ys[3]) / (ys[0] - ys[3])).toBeCloseTo(0.3 / 0.5, 6)
    expect((ys[2] - ys[3]) / (ys[0] - ys[3])).toBeCloseTo(0.1 / 0.5, 6)
    columns.forEach((c, i) => expect(xs[i]).toBeCloseTo(c.x + c.w / 2, 6))
    // An all-zero owner draws a zero line (never NaN).
    expect(nums((await native({ values: [0, 0], categories: ['A', 'B'], layoutPr: '<cx:aggregation />', pareto: true })).groups[1].series[0].values)).toEqual([0, 0])

    // PA1 (native): 3, 4, 5, 3, 4, 3, 4, 1, 2, 1, automatic. Scott: mean 3, sum of squares 16, variance 16/9,
    // h = 3.49 * (4/3) * 10^(-1/3) = 2.15989, k = ceil(4 / h) = 2: [1, 3.16] = 6, (3.16, 5.32] = 4 (already
    // descending); cumulative 0.6, 1.
    const h = 3.49 * (4 / 3) * 10 ** (-1 / 3)
    const pa1 = await nativeXlsx('paretoLine/paretoLine.xlsx')
    expect(bars(pa1)).toEqual({ labels: [`[1, ${general(1 + h)}]`, `(${general(1 + h)}, ${general(1 + 2 * h)}]`], counts: [6, 4] })
    expect(nums(pa1.groups[1].series[0].values)).toEqual([0.6, 1])
    expect(pa1.preparedAxes?.axes.find(a => a.id === '2')).toMatchObject({ side: 'r', min: 0, max: 1 })

    // PA2 (native): 3,5,7,6,9,6,8,4,5,3,1,4, automatic, axis kinds swapped (valScaling on 0, catScaling on 1).
    // Scott: mean 61/12, sum of squares 56.9167, h = 3.49 * sqrt(56.9167 / 11) * 12^(-1/3) = 3.46754, k = ceil(8 / h) = 3:
    // [1, 4.47] = {3,4,3,1,4} 5, (4.47, 7.94] = {5,7,6,6,5} 5, (7.94, 11.4] = {9,8} 2. The 5/5 tie keeps bin order;
    // cumulative 5/12, 10/12, 1. Its labels draw counts (dataLabels value=1, pos inEnd).
    const h2 = 3.49 * Math.sqrt((56 + 11 / 12) / 11) * 12 ** (-1 / 3)
    const pa2 = await nativeXlsx('pareto-gapwidth/pareto-gapwidth.xlsx')
    expect(bars(pa2)).toEqual({ labels: [`[1, ${general(1 + h2)}]`, `(${general(1 + h2)}, ${general(1 + 2 * h2)}]`, `(${general(1 + 2 * h2)}, ${general(1 + 3 * h2)}]`], counts: [5, 5, 2] })
    nums(pa2.groups[1].series[0].values).forEach((v, i) => expect(v).toBeCloseTo([5 / 12, 10 / 12, 1][i] as number, 12))
    expect(pa2.groups[0].gapWidth).toBe(255)
    expect(pa2.preparedAxes?.dataLabels.map(l => l.text)).toEqual(['5', '5', '2'])
  })

  it('a 4097-sample histogram uses every sample', async () => {
    // 1..4097 with binSize 1000 (r): [1,1001] 1001, (1001,2001] 1000, (2001,3001] 1000, (3001,4001] 1000,
    // (4001,5001] {4002..4097} 96; the counts sum to 4097, so every sample was used (the display cap is 4,096).
    const values = Array.from({ length: 4097 }, (_, i) => i + 1)
    const expected = { labels: ['[1, 1001]', '(1001, 2001]', '(2001, 3001]', '(3001, 4001]', '(4001, 5001]'], counts: [1001, 1000, 1000, 1000, 96] }
    const layoutPr = '<cx:binning><cx:binSize val="1000" /></cx:binning>'
    // The cache path.
    const cached = await native({ values, layoutPr })
    expect(bars(cached)).toEqual(expected)
    expect(cached.truncated).toBeUndefined()
    // The XLSX defined-name path (the workbook fill).
    const referenced = (formula: string, external = false) => space(`<cx:chartData>${external ? '<cx:externalData r:id="rId1" />' : ''}<cx:data id="0"><cx:numDim type="val"><cx:f>${formula}</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr>${layoutPr}</cx:layoutPr></cx:series></cx:plotAreaRegion>${AXES}</cx:plotArea></cx:chart>`)
    const zip = new JSZip()
    zip.file('xl/charts/chartEx1.xml', referenced('_xlchart.v1.0'))
    const p = await readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), 'xl/charts/chartEx1.xml')
    const result = await parseChartPart(p.root, theme, p.context)
    expect(result.kind).toBe('native')
    if (result.kind !== 'native') return
    const rows = values.map((v, i) => ({ index: i, cells: [{ row: i, col: 0, value: v }] }))
    const sheet = { name: 'Sheet1', workbookIndex: 0, rows, drawings: [{ content: { kind: 'xlsxChart', chart: result.model } }] }
    fillWorkbookCharts([sheet], { names: [{ name: '_xlchart.v1.0', source: 'Sheet1!$A$1:$A$4097' }], workspace: p.context.workspace })
    const filled = (sheet.drawings[0].content as { chart: ChartModel }).chart
    expect(bars(filled)).toEqual(expected)
    releaseChartModel(filled)
    p.release()
    // The DOCX/PPTX embedded-workbook path.
    const data = values.map((v, i) => row(i + 1, cell(`A${i + 1}`, v))).join('')
    const built = buildChartPptx([{ chartXml: referenced('Sheet1!$A$1:$A$4097', true), workbook: embeddedWorkbook({ sheets: [{ name: 'Sheet1', data }] }) }])
    const pkg = await OfficePackage.load(built.bytes)
    const doc = documentChartContext('desktop')
    const e = await readPart(pkg, built.chartPaths[0], chartParseContext(doc, pkg, built.chartPaths[0]))
    const embedded = await parseChartPart(e.root, theme, e.context)
    expect(embedded.kind).toBe('native')
    if (embedded.kind === 'native') { expect(bars(embedded.model)).toEqual(expected); releaseChartModel(embedded.model) }
    e.release()
    expect(doc.workspace.used, 'every derived byte was owned and refunded').toBe(0)
  }, 60_000)

  it('65537 samples refuse', async () => {
    // The analytic bound is 65,536 per series: one more refuses the chart (never a prefix), on the cache path and
    // through the workbook fill; 65,536 is drawn from every sample.
    const at = (n: number) => Array.from({ length: n }, (_, i) => i % 10)
    expect(await adapt(part({ values: at(65537) }))).toMatchObject({ kind: 'refused', reason: 'This chart has too many data points to summarize', feature: 'ChartEx' })
    const full = await native({ values: at(65536), layoutPr: '<cx:binning><cx:binCount val="10" /></cx:binning>' })
    expect(nums(full.groups[0].series[0].values).reduce((a, b) => (a as number) + (b as number), 0)).toBe(65536)
    const zip = new JSZip()
    zip.file('xl/charts/chartEx1.xml', space(`<cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr><cx:binning /></cx:layoutPr></cx:series></cx:plotAreaRegion>${AXES}</cx:plotArea></cx:chart>`))
    const p = await readPart(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), 'xl/charts/chartEx1.xml')
    const model = parseChartEx(p.root, theme, p.context)
    const rows = at(65537).map((v, i) => ({ index: i, cells: [{ row: i, col: 0, value: v }] }))
    const sheet = { name: 'Sheet1', workbookIndex: 0, rows, drawings: [{ content: { kind: 'xlsxChart', chart: model } }] }
    fillWorkbookCharts([sheet], { names: [{ name: '_xlchart.v1.0', source: 'Sheet1!$A$1:$A$65537' }], workspace: p.context.workspace })
    expect((sheet.drawings[0].content as { chart: ChartModel }).chart).toMatchObject({ unsupported: 'This chart has too many data points to summarize', refused: true })
    p.release()
  }, 60_000)

  it('several column series share one set of bins; a Pareto binds its owner through ownerIdx', async () => {
    // Pooled bins [P]: S1 {1, 2}, S2 {3, 4}, binSize 1 over 1..4: [1,2] (2,3] (3,4]; S1 2,0,0; S2 0,1,1.
    const model = await native({ values: [1, 2], more: [[3, 4]], layoutPr: '<cx:binning><cx:binSize val="1" /></cx:binning>' })
    expect(bars(model, 0)).toEqual({ labels: ['[1, 2]', '(2, 3]', '(3, 4]'], counts: [2, 0, 0] })
    expect(bars(model, 1).counts).toEqual([0, 1, 1])
    // A Pareto with more than one column series has no single sort order: unsupported.
    expect(await adapt(part({ values: [1, 2], more: [[3]], pareto: true }))).toMatchObject({ kind: 'unsupported', feature: 'ChartEx' })
  })
})

/** Conformance (spec 7): histogram and Pareto parts around each case's numeric level (Office-shaped synthetic). */
const cxPart = (level: string | null, pareto: boolean) => space(`<cx:chartData>${level === null ? '' : `<cx:data id="0"><cx:numDim type="val">${level}</cx:numDim></cx:data>`}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${level === null ? '' : `<cx:series layoutId="clusteredColumn"><cx:dataId val="0" /><cx:layoutPr><cx:binning intervalClosed="r" /></cx:layoutPr><cx:axisId val="1" /></cx:series>${pareto ? '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2" /></cx:series>' : ''}`}</cx:plotAreaRegion>${AXES}${pareto ? PERCENT_AXIS : ''}</cx:plotArea></cx:chart>`)
for (const pareto of [false, true]) describeFamilyConformance({
  name: pareto ? 'pareto' : 'histogram',
  chartEx: true,
  chart: level => cxPart(level, pareto),
  // No series: the part has nothing to draw (unsupported, card); an unavailable point refuses an analytic chart.
  expect: { 'no series': { kind: 'card' }, 'empty series': { kind: 'native' }, 'all blank': { kind: 'native' }, 'NaN and infinity': { kind: 'native' }, 'negative values': { kind: 'native' }, 'single point': { kind: 'native' }, 'unavailable point': { kind: 'card' } },
})
