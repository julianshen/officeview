// @vitest-environment node
/**
 * Task A6b (spec 5.7, 5.8, 7): error bars. Numeric and geometry oracles are hand-computed (closed forms noted beside
 * each case); fixtures are Office-shaped synthetic chart parts, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXmlOrdered } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart } from '../src/charts/layout'
import { pointList, PointWriter } from '../src/charts/points'
import { pendingLines } from '../src/charts/fill'
import { errorBarStatistic } from '../src/charts/errorBars'
import { chartParseContext, documentChartContext } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { applyDataPrecedence } from '../src/charts/dataPrecedence'
import { releaseChartModel } from '../src/charts/own'
import { createChartWorkspace } from '../src/charts/workspace'
import { preparedAxesBytes } from '../src/charts/prepare'
import type { ChartModel, ChartPoints } from '../src/charts/model'
import { parseXlsx } from '../src/xlsx/parse'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { buildChartDocx, cell, embeddedWorkbook, row } from './helpers/chart-fixtures'

const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const lit = (values: Array<number | string | null>) => `<numLit><ptCount val="${values.length}"/>${values.map((v, i) => v === null ? '' : `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numLit>`
const RED = '<spPr><ln w="9525"><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr>'
const BLUE = '<spPr><ln w="9525"><solidFill><srgbClr val="0000FF"/></solidFill></ln></spPr>'
/** One `c:errBars` (schema order: errDir, errBarType, errValType, noEndCap, plus, minus, val, spPr). */
const bars = (o: { dir?: string; type?: string; val?: string; valType?: string; noEndCap?: boolean; plus?: string; minus?: string; spPr?: string }) =>
  `<errBars>${o.dir ? `<errDir val="${o.dir}"/>` : ''}<errBarType val="${o.type ?? 'both'}"/><errValType val="${o.valType ?? 'fixedVal'}"/><noEndCap val="${o.noEndCap ? 1 : 0}"/>`
  + `${o.plus !== undefined ? `<plus>${o.plus}</plus>` : ''}${o.minus !== undefined ? `<minus>${o.minus}</minus>` : ''}${o.val !== undefined ? `<val val="${o.val}"/>` : ''}${o.spPr ?? RED}</errBars>`
const AX = '<axId val="1"/><axId val="2"/>'
const scale = (min: number, max: number) => `<scaling><orientation val="minMax"/><max val="${max}"/><min val="${min}"/></scaling>`
const space = (plot: string) => `<chartSpace><chart><autoTitleDeleted val="1"/><plotArea>${plot}</plotArea></chart></chartSpace>`
const scatter = (xs: number[], ys: number[], errBars: string) => parse(space(
  `<scatterChart><scatterStyle val="marker"/><ser><idx val="0"/><order val="0"/>${errBars}<xVal>${lit(xs)}</xVal><yVal>${lit(ys)}</yVal></ser>${AX}</scatterChart>`
  + `<valAx><axId val="1"/>${scale(0, 10)}<axPos val="b"/><majorUnit val="2"/><crossAx val="2"/></valAx><valAx><axId val="2"/>${scale(0, 20)}<axPos val="l"/><majorUnit val="5"/><crossAx val="1"/></valAx>`))
/** A category chart; `valueScale` authors the value axis bounds (absent: automatic). */
const category = (family: 'lineChart' | 'barChart', series: Array<{ values: Array<number | string | null>; errBars: string }>, o: { grouping?: string; barDir?: string; valueScale?: [number, number]; extra?: string } = {}) => parse(space(
  `<${family}>${family === 'barChart' ? `<barDir val="${o.barDir ?? 'col'}"/>` : ''}<grouping val="${o.grouping ?? (family === 'barChart' ? 'clustered' : 'standard')}"/>`
  + series.map((s, i) => `<ser><idx val="${i}"/><order val="${i}"/>${s.errBars}<val>${lit(s.values)}</val></ser>`).join('')
  + `${family === 'barChart' ? `<gapWidth val="100"/>${o.grouping === 'stacked' ? '<overlap val="100"/>' : ''}` : ''}${AX}</${family}>`
  + `<catAx><axId val="1"/><axPos val="${o.barDir === 'bar' ? 'l' : 'b'}"/><crossAx val="2"/></catAx><valAx><axId val="2"/>${o.valueScale ? scale(...o.valueScale) : ''}<axPos val="${o.barDir === 'bar' ? 'b' : 'l'}"/>${o.valueScale ? `<majorUnit val="${(o.valueScale[1] - o.valueScale[0]) / 5}"/>` : ''}<crossAx val="1"/></valAx>${o.extra ?? ''}`))
const prepared = (model: ChartModel) => model.preparedAxes?.errorBars ?? []
const codes = (model: ChartModel) => (model.preparedAxes?.diagnostics ?? []).map(d => d.code)
const states = (points: ChartPoints | undefined) => points ? pointList(points).map(p => p.state === 'value' ? p.v : p.state === 'error' ? p.code : p.state) : undefined

type Segment = { x1: number; y1: number; x2: number; y2: number }
const W = 400, H = 260
function paint(model: ChartModel) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const strokes: Array<{ style: string; segments: Segment[] }> = []
  let path: Segment[] = [], pen: [number, number] = [0, 0]
  const o = { moveTo: ctx.moveTo.bind(ctx), lineTo: ctx.lineTo.bind(ctx), beginPath: ctx.beginPath.bind(ctx), stroke: ctx.stroke.bind(ctx) }
  ctx.beginPath = () => { path = []; o.beginPath() }
  ctx.moveTo = (x, y) => { pen = [x, y]; o.moveTo(x, y) }
  ctx.lineTo = (x, y) => { path.push({ x1: pen[0], y1: pen[1], x2: x, y2: y }); pen = [x, y]; o.lineTo(x, y) }
  ctx.stroke = (...a: []) => { strokes.push({ style: String(ctx.strokeStyle).toUpperCase(), segments: path }); o.stroke(...a) }
  paintXlsxChart(model, ctx, W, H)
  const layout = layoutChart(model, W, H, 1, 0, { width: (text: string, css: string) => { ctx.font = css; return ctx.measureText(text).width }, lineHeight: () => model.fontSizePt * 96 / 72 })
  const of = (style: string) => strokes.filter(s => s.style === style).flatMap(s => s.segments)
  return { strokes, of, plot: layout.plot }
}
/** Segments compared as an unordered set, to 6 decimals. */
const key = (s: Segment) => [s.x1, s.y1, s.x2, s.y2].map(v => v.toFixed(6)).join(',')
const sameSegments = (actual: Segment[], expected: Segment[]) => expect(actual.map(key).sort()).toEqual(expected.map(key).sort())
const seg = (x1: number, y1: number, x2: number, y2: number): Segment => ({ x1, y1, x2, y2 })
/** The engine's end cap: 3 px either side of the bar's end, across the bar. */
const CAP = 3
/** A vertical bar at x from value a to b (shaft), with its cap at b. */
const vbar = (x: number, a: number, b: number, cap = true) => [seg(x, a, x, b), ...(cap ? [seg(x - CAP, b, x + CAP, b)] : [])]
const hbar = (y: number, a: number, b: number, cap = true) => [seg(a, y, b, y), ...(cap ? [seg(b, y - CAP, b, y + CAP)] : [])]

describe('charts-error-bars (Task A6b)', () => {
  it('x and y directions; plus, minus and both; noEndCap segment geometry', () => {
    // Scatter, x axis 0..10 and y axis 0..20 (authored): px(x) = plot.x + plot.w·x/10, py(y) = plot.y + plot.h·(1 - y/20).
    // Points (2,5), (6,10). y bars: fixed 2, both, capped (red). x bars: fixed 1, plus only, no end cap (blue).
    const xy = scatter([2, 6], [5, 10], bars({ dir: 'y', val: '2' }) + bars({ dir: 'x', type: 'plus', val: '1', noEndCap: true, spPr: BLUE }))
    expect(xy.groups[0].series[0].errorBars?.map(b => [b.dir, b.type, b.valType, b.noEndCap])).toEqual([['y', 'both', 'fixedVal', false], ['x', 'plus', 'fixedVal', true]])
    const p = paint(xy), { plot } = p
    const px = (x: number) => plot.x + plot.w * x / 10, py = (y: number) => plot.y + plot.h * (1 - y / 20)
    // Both: per point a plus shaft (centre -> v+2) and a minus shaft (centre -> v-2), each with its cap: 4 segments.
    sameSegments(p.of('#FF0000'), [
      ...vbar(px(2), py(5), py(7)), ...vbar(px(2), py(5), py(3)),
      ...vbar(px(6), py(10), py(12)), ...vbar(px(6), py(10), py(8)),
    ])
    // Plus only, no end cap: one horizontal shaft per point (x -> x+1), nothing else.
    sameSegments(p.of('#0000FF'), [...hbar(py(5), px(2), px(3), false), ...hbar(py(10), px(6), px(7), false)])

    // Line on a category axis (4 slots between ticks): x_i = plot.x + plot.w·(i + 0.5)/4; value axis 0..10.
    // Minus only, fixed 1: per point one shaft down to v - 1 and its cap.
    const line = category('lineChart', [{ values: [2, 4, 6, 8], errBars: bars({ type: 'minus', val: '1' }) }], { valueScale: [0, 10] })
    const l = paint(line), lp = l.plot
    const lx = (i: number) => lp.x + lp.w * (i + 0.5) / 4, ly = (v: number) => lp.y + lp.h * (1 - v / 10)
    sameSegments(l.of('#FF0000'), [2, 4, 6, 8].flatMap((v, i) => vbar(lx(i), ly(v), ly(v - 1))))

    // Columns (one series, gap 100): the bar's centre is the slot centre; the bar sits on its end (the value).
    const columns = category('barChart', [{ values: [3, 5], errBars: bars({ type: 'plus', val: '2' }) }], { valueScale: [0, 10] })
    const c = paint(columns), cp = c.plot
    const cx = (i: number) => cp.x + cp.w * (i + 0.5) / 2, cy = (v: number) => cp.y + cp.h * (1 - v / 10)
    sameSegments(c.of('#FF0000'), [...vbar(cx(0), cy(3), cy(5)), ...vbar(cx(1), cy(5), cy(7))])

    // Horizontal bars: "y" is the value direction, horizontal on screen; categories run bottom to top.
    const horizontal = category('barChart', [{ values: [3, 5], errBars: bars({ type: 'plus', val: '2' }) }], { barDir: 'bar', valueScale: [0, 10] })
    const h = paint(horizontal), hp = h.plot
    const hy = (i: number) => hp.y + hp.h - hp.h * (i + 0.5) / 2, hx = (v: number) => hp.x + hp.w * v / 10
    sameSegments(h.of('#FF0000'), [...hbar(hy(0), hx(3), hx(5)), ...hbar(hy(1), hx(5), hx(7))])

    // Stacked columns: a series' bar sits on its segment end (the running stack): 2+1 = 3, 3+1 = 4.
    const stacked = category('barChart', [
      { values: [2, 3], errBars: '' },
      { values: [1, 1], errBars: bars({ type: 'plus', val: '0.5' }) },
    ], { grouping: 'stacked', valueScale: [0, 10] })
    const s = paint(stacked), sp = s.plot
    const sx = (i: number) => sp.x + sp.w * (i + 0.5) / 2, sy = (v: number) => sp.y + sp.h * (1 - v / 10)
    sameSegments(s.of('#FF0000'), [...vbar(sx(0), sy(3), sy(3.5)), ...vbar(sx(1), sy(4), sy(4.5))])

    // x on a non-XY chart is never drawn (spec 5.7: y only for non-XY), diagnosed.
    const sideways = category('lineChart', [{ values: [2, 4], errBars: bars({ dir: 'x', val: '1' }) }], { valueScale: [0, 10] })
    expect(prepared(sideways)).toHaveLength(0)
    expect(codes(sideways)).toContain('chart-error-bars-dropped')
    expect(paint(sideways).of('#FF0000')).toHaveLength(0)

    // Adding error bars moves nothing else: the series' own strokes are identical with and without them.
    const plain = category('lineChart', [{ values: [2, 4, 6, 8], errBars: '' }], { valueScale: [0, 10] })
    const series = (r: ReturnType<typeof paint>) => r.strokes.filter(x => x.style !== '#FF0000').map(x => x.segments.map(key))
    expect(series(l)).toEqual(series(paint(plain)))
  })

  it('fixed, percentage, stdDev (n-1, centred on the series mean), stdErr', () => {
    // y = 2,4,4,4,5,5,7,9: n 8, mean 5, Σ(y-5)² = 9+1+1+1+0+0+4+16 = 32; sample sd = sqrt(32/7).
    const ys = [2, 4, 4, 4, 5, 5, 7, 9]
    const sd = Math.sqrt(32 / 7)
    // stdDev × 2: every bar is centred on the mean 5 (not the point), half-length 2·sd.
    const dev = category('lineChart', [{ values: ys, errBars: bars({ valType: 'stdDev', val: '2' }) }], { valueScale: [-10, 20] })
    expect(prepared(dev)).toHaveLength(1)
    expect(prepared(dev)[0].center).toBeCloseTo(5, 12)
    expect(prepared(dev)[0].amount).toBeCloseTo(2 * sd, 12)
    const d = paint(dev), dp = d.plot
    const x = (i: number, n = 8) => dp.x + dp.w * (i + 0.5) / n, y = (v: number) => dp.y + dp.h * (1 - (v + 10) / 30)
    sameSegments(d.of('#FF0000'), ys.flatMap((_, i) => [...vbar(x(i), y(5), y(5 + 2 * sd)), ...vbar(x(i), y(5), y(5 - 2 * sd))]))
    // Blank and error points leave the statistic (spec 5.8): the same set with a gap and a #N/A gives the same bars.
    const gapped = category('lineChart', [{ values: [2, 4, null, 4, 4, '#N/A', 5, 5, 7, 9], errBars: bars({ valType: 'stdDev', val: '2' }) }], { valueScale: [-10, 20] })
    expect(prepared(gapped)[0].center).toBeCloseTo(5, 12)
    expect(prepared(gapped)[0].amount).toBeCloseTo(2 * sd, 12)
    // stdErr = sd/√n = sqrt(32/7/8) = sqrt(4/7), centred on each point; the authored val has no role.
    const err = category('lineChart', [{ values: ys, errBars: bars({ valType: 'stdErr', val: '3' }) }], { valueScale: [-10, 20] })
    const se = Math.sqrt(4 / 7)
    expect(prepared(err)[0].amount).toBeCloseTo(se, 12)
    expect(prepared(err)[0].center).toBeUndefined()
    const e = paint(err)
    sameSegments(e.of('#FF0000'), ys.flatMap((v, i) => [...vbar(x(i), y(v), y(v + se)), ...vbar(x(i), y(v), y(v - se))]))
    // Fixed 1.5 and percentage 50% (of |v|: -4 -> 2, 2 -> 1), each from the point.
    const fixed = paint(category('lineChart', [{ values: [1, 3], errBars: bars({ type: 'plus', val: '1.5' }) }], { valueScale: [-10, 20] }))
    sameSegments(fixed.of('#FF0000'), [...vbar(x(0, 2), y(1), y(2.5)), ...vbar(x(1, 2), y(3), y(4.5))])
    const percent = paint(category('lineChart', [{ values: [-4, 2], errBars: bars({ valType: 'percentage', val: '50' }) }], { valueScale: [-10, 20] }))
    sameSegments(percent.of('#FF0000'), [
      ...vbar(x(0, 2), y(-4), y(-2)), ...vbar(x(0, 2), y(-4), y(-6)),
      ...vbar(x(1, 2), y(2), y(3)), ...vbar(x(1, 2), y(2), y(1)),
    ])
    // The statistic itself: an unavailable point drops the bars; fewer than 2 values have no sample deviation.
    const points = (list: Array<number | 'blank' | 'unavailable'>) => {
      const out = new PointWriter(list.length)
      list.forEach((v, i) => { if (v === 'unavailable') out.unavailable(i); else if (v !== 'blank') out.value(i, v) })
      return out.done()
    }
    expect(errorBarStatistic('stdDev', points([1, 2, 3]), 1)).toMatchObject({ kind: 'stat', center: 2, amount: 1, n: 3 })
    expect(errorBarStatistic('stdErr', points([1, 2, 3]), 1)).toMatchObject({ kind: 'stat', amount: 1 / Math.sqrt(3) })
    expect(errorBarStatistic('stdDev', points([1, 'unavailable', 3]), 1)).toMatchObject({ kind: 'none', reason: expect.stringMatching(/unavailable/) })
    expect(errorBarStatistic('stdDev', points([4, 'blank']), 1)).toMatchObject({ kind: 'none' })
    const single = category('lineChart', [{ values: [3], errBars: bars({ valType: 'stdDev', val: '1' }) }])
    expect(prepared(single)).toHaveLength(0)
    expect(codes(single)).toContain('chart-error-bars-dropped')
    // Automatic value axes include the bar ends: 1..3 with fixed 10 reach -9 and 13.
    const wide = category('lineChart', [{ values: [1, 2, 3], errBars: bars({ val: '10' }) }])
    const value = wide.preparedAxes!.axes.find(a => a.id === '2')!
    expect(value.max).toBeGreaterThanOrEqual(13)
    expect(value.min).toBeLessThanOrEqual(-9)
    // A malformed or negative authored amount draws nothing (never a default), diagnosed.
    for (const val of ['-1', 'NaN', 'abc']) {
      const bad = category('lineChart', [{ values: [1, 2], errBars: bars({ val }) }])
      expect(prepared(bad)).toHaveLength(0)
      expect(codes(bad)).toContain('chart-error-bars-parameter')
    }
  })

  it('custom plus and minus from caches and the embedded workbook; a missing custom point draws no bar', async () => {
    // Cached: plus [1, -, 2, -], minus [0.5, -, 0.5, 1]. Point 1 has neither: no bar. Point 3 has no plus: no plus side.
    const cached = category('lineChart', [{ values: [2, 4, 6, 8], errBars: bars({ valType: 'cust', plus: lit([1, null, 2, null]), minus: lit([0.5, null, 0.5, 1]) }) }], { valueScale: [0, 10] })
    expect(states(cached.groups[0].series[0].errorBars![0].plus)).toEqual([1, 'blank', 2, 'blank'])
    const c = paint(cached), cp = c.plot
    const x = (i: number, n = 4) => cp.x + cp.w * (i + 0.5) / n, y = (v: number) => cp.y + cp.h * (1 - v / 10)
    sameSegments(c.of('#FF0000'), [
      ...vbar(x(0), y(2), y(3)), ...vbar(x(0), y(2), y(1.5)),
      ...vbar(x(2), y(6), y(8)), ...vbar(x(2), y(6), y(5.5)),
      ...vbar(x(3), y(8), y(7)),
    ])

    // DOCX: no cache, so the custom amounts come from the embedded workbook (spec 6.4 precedence 2). C3 is empty and
    // C4 holds #N/A: those points draw no plus side. A plus reference to a missing sheet leaves that side unavailable,
    // diagnosed, while the chart and its other side still draw.
    const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    const xml = (plus: string) => `<c:chartSpace ${NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/>`
      + `<c:errBars><c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="0"/><c:plus><c:numRef><c:f>${plus}</c:f></c:numRef></c:plus><c:minus><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f></c:numRef></c:minus><c:spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:errBars>`
      + `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>`
      + `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="10"/><c:min val="0"/></c:scaling><c:axPos val="l"/><c:majorUnit val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:externalData r:id="rId1"/></c:chartSpace>`
    const data = [
      row(2, cell('B2', 2), cell('C2', 1), cell('D2', 0.5)),
      row(3, cell('B3', 4), cell('D3', 0.5)),
      row(4, cell('B4', 6), cell('C4', '#N/A', { t: 'e' }), cell('D4', 0.5)),
    ].join('')
    const workbook = embeddedWorkbook({ sheets: [{ name: 'Sheet1', data }] })
    const fixture = buildChartDocx([{ chartXml: xml('Sheet1!$C$2:$C$4'), workbook }, { chartXml: xml('Missing!$C$2:$C$4'), workbook }])
    const pkg = await OfficePackage.load(fixture.bytes)
    const doc = documentChartContext('desktop')
    const charts: ChartModel[] = []
    for (const path of fixture.chartPaths) {
      const context = chartParseContext(doc, pkg, path)
      const read = await readChartPart(pkg, path, context)
      if (read?.kind !== 'tree') throw new Error('chart part unreadable')
      const model = parseXlsxChart(read.root, undefined, context)!
      read.release()
      charts.push(await applyDataPrecedence(model, read.root, context))
    }
    const [filled, missing] = charts
    expect(filled.unsupported).toBeUndefined()
    const errorBars = filled.groups[0].series[0].errorBars![0]
    expect(states(errorBars.plus)).toEqual([1, 'blank', '#N/A'])
    expect(states(errorBars.minus)).toEqual([0.5, 0.5, 0.5])
    const f = paint(filled), fp = f.plot
    const fx = (i: number) => fp.x + fp.w * (i + 0.5) / 3, fy = (v: number) => fp.y + fp.h * (1 - v / 10)
    sameSegments(f.of('#FF0000'), [
      ...vbar(fx(0), fy(2), fy(3)), ...vbar(fx(0), fy(2), fy(1.5)),
      ...vbar(fx(1), fy(4), fy(3.5)),
      ...vbar(fx(2), fy(6), fy(5.5)),
    ])
    expect(missing.unsupported).toBeUndefined()
    expect(missing.diagnostics?.map(d => d.code)).toContain('chart-data-sheet-missing')
    expect(missing.groups[0].series[0].errorBars![0].plus).toBeUndefined()
    const m = paint(missing)
    sameSegments(m.of('#FF0000'), [0, 1, 2].flatMap(i => vbar(fx(i), fy([2, 4, 6][i]), fy([2, 4, 6][i] - 0.5))))
    for (const chart of charts) releaseChartModel(chart)
    expect(doc.workspace.used).toBe(0)

    // XLSX: an uncached plus reference is filled from the worksheet (resident), and the paged reader asks for its line.
    const sheetXml = space(`<lineChart><grouping val="standard"/><ser><idx val="0"/><order val="0"/>${bars({ valType: 'cust', type: 'plus', plus: '<numRef><f>Styled!$B$1:$B$3</f></numRef>' })}<val>${lit([2, 4, 6])}</val></ser>${AX}</lineChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>`)
    const parsed = parse(sheetXml)
    expect(parsed.groups[0].series[0].pending).toEqual({ errorPlus: { r0: 0, c0: 1, r1: 2, c1: 1, sheet: 'Styled' } })
    const lines = pendingLines([{ content: { kind: 'xlsxChart', chart: parsed } as { kind: string }  }], 'Styled')
    expect(lines).toContainEqual(expect.objectContaining({ sheet: 'Styled', vertical: true, fixed: 1, start: 0, count: 3 }))
    const book = await parseXlsx(await chartWorkbook(sheetXml, { cells: [{ ref: 'B1', v: 0.5 }, { ref: 'B3', v: 1.5 }] }))
    const content = book.sheets[0].drawings?.[0].content as { kind: string; chart: ChartModel }
    expect(book.drawingCoverage?.find(entry => entry.feature === 'chart')).toMatchObject({ status: 'native', selectedRepresentation: 'chart' })
    expect(content.chart.unsupported).toBeUndefined()
    expect(states(content.chart.groups[0].series[0].errorBars![0].plus)).toEqual([0.5, 'blank', 1.5])
    expect(prepared(content.chart)).toHaveLength(1)
  })

  it('records, custom amounts and prepared sets are charged before they are built', () => {
    // The same chart with and without one custom set (4-point plus and minus literals, a red a:ln): the model tag
    // gains exactly each record (128), each red line (112, plus 96 for its solid fill) and two 4-point roles
    // (4 × 9 bytes each, no formats).
    const run = (errBars: string) => {
      const workspace = createChartWorkspace('mobile')
      const calls: Array<{ tag: string; bytes: number }> = []
      const reserve = workspace.reserve.bind(workspace)
      workspace.reserve = (bytes, tag) => { calls.push({ tag, bytes }); return reserve(bytes, tag) }
      const xml = space(`<lineChart><grouping val="standard"/><ser><idx val="0"/><order val="0"/>${errBars}<val>${lit([2, 4, 6, 8])}</val></ser>${AX}</lineChart><catAx><axId val="1"/><crossAx val="2"/></catAx><valAx><axId val="2"/><crossAx val="1"/></valAx>`)
      const model = parseXlsxChart(parseXmlOrdered(xml), undefined, { workspace, profile: 'mobile', partPath: 'xl/charts/chart1.xml', pkg: undefined as never })!
      const modelBytes = calls.filter(c => /^chart-model#\d+$/.test(c.tag)).reduce((n, c) => n + c.bytes, 0)
      return { model, workspace, calls, modelBytes }
    }
    const plain = run(''), withBars = run(bars({ valType: 'cust', plus: lit([1, 1, 1, 1]), minus: lit([1, 1, 1, 1]) }) + bars({ dir: 'x', val: '1' }))
    // The second (x) set on a line chart is still a parsed record (dropped at stage 1, never drawn).
    expect(withBars.modelBytes - plain.modelBytes).toBe(2 * 128 + 2 * (112 + 96) + 2 * 36)
    // One up-front reservation on the preparation tag covers everything stage 1 retains.
    const prep = withBars.calls.filter(c => /^chart-prepare#\d+$/.test(c.tag))
    expect(prep).toHaveLength(1)
    expect(prepared(withBars.model)).toHaveLength(1)
    expect(preparedAxesBytes(withBars.model.preparedAxes)).toBeLessThanOrEqual(prep[0].bytes)
    releaseChartModel(withBars.model); releaseChartModel(plain.model)
    expect(withBars.workspace.used).toBe(0)
  })

  it('analytic stdDev over 65536 drops the bars', () => {
    const range = (n: number) => Array.from({ length: n }, (_, i) => i + 1)
    const over = category('lineChart', [{ values: range(65537), errBars: bars({ valType: 'stdDev', val: '1' }) }])
    expect(over.groups[0].series[0].analytic).toBe('refused')
    expect(prepared(over)).toHaveLength(0)
    expect(over.preparedAxes!.diagnostics.find(d => d.code === 'chart-error-bars-dropped')?.reason).toMatch(/65,536/)
    expect(over.unsupported).toBeUndefined()
    // A fixed amount is not analytic: the same series keeps its (display-capped) bars.
    const fixed = category('lineChart', [{ values: range(65537), errBars: bars({ val: '1' }) }])
    expect(fixed.groups[0].series[0].analytic).toBeUndefined()
    expect(prepared(fixed)).toHaveLength(1)
    // 4,097 points: the statistic uses every acquired point, never the 4,096-point display prefix.
    // Values 1..4096 then 1e6: mean = (4096·4097/2 + 1e6)/4097 (the prefix mean would be 2048.5).
    const n = 4097, ys = range(n).map(v => v === n ? 1e6 : v)
    const all = category('lineChart', [{ values: ys, errBars: bars({ valType: 'stdDev', val: '1' }) }])
    const mean = (4096 * 4097 / 2 + 1e6) / n
    let ss = 0
    for (const v of ys) ss += (v - mean) ** 2
    expect(all.groups[0].series[0].values.length).toBe(n)
    expect(prepared(all)[0].center).toBeCloseTo(mean, 9)
    expect(prepared(all)[0].amount! / Math.sqrt(ss / (n - 1))).toBeCloseTo(1, 12)
  })
})
