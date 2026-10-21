/**
 * A5a: bubble and radar charts (spec 4.1), the radar cells of the point matrix
 * (spec 5.8), and the family conformance harness (spec 7) over every family
 * drawn natively. Fixtures are Office-shaped synthetic, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { documentChartContext, chartParseContext } from '../src/charts/context'
import { parseChartPart, CHART_DATA_UNAVAILABLE } from '../src/charts/dataPrecedence'
import { parseXlsxChart } from '../src/charts/parse'
import { fillChartRanges } from '../src/charts/fill'
import { paintXlsxChart } from '../src/charts/paint'
import { plotPoint } from '../src/charts/plotPoint'
import { pointList, PointWriter } from '../src/charts/points'
import { prepareChartAxes } from '../src/charts/prepare'
import { createChartWorkspace } from '../src/charts/workspace'
import type { ChartModel, ChartPoint } from '../src/charts/model'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'
import { chartDescription } from '../src/charts/describe'
import { buildChartDocx, buildChartPptx, cell, embeddedWorkbook, row } from './helpers/chart-fixtures'
import { CLASSIC_FAMILIES, CONFORMANCE_NS, describeFamilyConformance, type ConformanceCase } from './helpers/chart-family-conformance'

const lit = (values: Array<number | string | undefined>) => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const strLit = (labels: string[]) => `<c:strLit><c:ptCount val="${labels.length}"/>${labels.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strLit>`
const space = (plot: string, options = '') => `<c:chartSpace ${CONFORMANCE_NS}><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:plotVisOnly val="1"/>${options}</c:chart></c:chartSpace>`
const VAL_VAL = '<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'
const RADAR_AXES = '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorGridlines/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>'
const RED_FILL_BLUE_LINE = '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>'

interface BubbleSpec { sizes: Array<number | string | undefined>; scale?: number; represents?: 'area' | 'w'; showNeg?: '0' | '1'; sizeNode?: string }
function bubbleXml({ sizes, scale, represents, showNeg, sizeNode }: BubbleSpec): string {
  const xs = sizes.map((_, i) => i + 1), ys = sizes.map((_, i) => (i % 2) + 2)
  return space(`<c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>B</c:v></c:tx>${RED_FILL_BLUE_LINE}<c:invertIfNegative val="0"/><c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal>${sizeNode ?? `<c:bubbleSize>${lit(sizes)}</c:bubbleSize>`}<c:bubble3D val="0"/></c:ser>${scale !== undefined ? `<c:bubbleScale val="${scale}"/>` : ''}${showNeg !== undefined ? `<c:showNegBubbles val="${showNeg}"/>` : ''}${represents ? `<c:sizeRepresents val="${represents}"/>` : ''}<c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${VAL_VAL}`)
}
const parse = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!

interface Shape { x: number; y: number; r: number; op: 'fill' | 'stroke'; style: string }
/** Every circle a paint fills or strokes, with the style it was filled or stroked with. */
function circles(model: ChartModel, w = 400, h = 300): Shape[] {
  const ctx = createCanvas(w, h).getContext('2d'), out: Shape[] = []
  let current: Array<{ x: number; y: number; r: number }> = []
  const begin = ctx.beginPath.bind(ctx), arc = ctx.arc.bind(ctx), fill = ctx.fill.bind(ctx), stroke = ctx.stroke.bind(ctx)
  ctx.beginPath = () => { current = []; begin() }
  ctx.arc = (x: number, y: number, r: number, a0: number, a1: number, ccw?: boolean) => { current.push({ x, y, r }); arc(x, y, r, a0, a1, ccw) }
  ctx.fill = ((...args: Parameters<typeof ctx.fill>) => { for (const c of current) out.push({ ...c, op: 'fill', style: String(ctx.fillStyle) }); fill(...args) }) as typeof ctx.fill
  ctx.stroke = ((...args: Parameters<typeof ctx.stroke>) => { for (const c of current) out.push({ ...c, op: 'stroke', style: String(ctx.strokeStyle) }); stroke(...args) }) as typeof ctx.stroke
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  return out
}
const filled = (shapes: Shape[]) => shapes.filter(s => s.op === 'fill' && s.style === '#ff0000').sort((a, b) => a.x - b.x)

describe('bubble (spec 4.1)', () => {
  it('bubble: bubbleScale, sizeRepresents area and width, zero not drawn, showNegBubbles 0 and 1', () => {
    // Sizes 1, 4, 0, -4 at x = 1..4: zero is never drawn; the negative only when showNegBubbles=1.
    const model = parse(bubbleXml({ sizes: [1, 4, 0, -4] }))
    expect(model.unsupported).toBeUndefined()
    expect(model.groups[0]).toMatchObject({ family: 'bubble', binding: { kind: 'xy', x: '1', y: '2' } })
    expect(pointList(model.groups[0].series[0].bubbleSizes!).map(p => p.state === 'value' ? p.v : p.state)).toEqual([1, 4, 0, -4])
    // Defaults: bubbleScale 100, size represents area, negative bubbles hidden.
    const area = filled(circles(model))
    expect(area).toHaveLength(2)
    expect(area[1].r / area[0].r).toBeCloseTo(2, 6) // area: radius ∝ √size
    const all = circles(model)
    expect(new Set(all.filter(s => s.r > 0).map(s => Math.round(s.x))).size).toBe(2) // zero and the hidden negative draw nothing at all
    // sizeRepresents=w: radius ∝ size.
    const width = filled(circles(parse(bubbleXml({ sizes: [1, 4, 0, -4], represents: 'w' }))))
    expect(width).toHaveLength(2)
    expect(width[1].r / width[0].r).toBeCloseTo(4, 6)
    expect(width[1].r).toBeCloseTo(area[1].r, 6) // the largest bubble is the same size either way
    // bubbleScale scales every bubble (50% halves, 200% doubles).
    const half = filled(circles(parse(bubbleXml({ sizes: [1, 4, 0, -4], scale: 50 }))))
    const double = filled(circles(parse(bubbleXml({ sizes: [1, 4, 0, -4], scale: 200 }))))
    expect(half.map(s => s.r / 2)).toEqual(area.map(s => expect.closeTo(s.r / 4, 6)))
    expect(double[1].r).toBeCloseTo(area[1].r * 2, 6)
    // The largest bubble's diameter is a quarter of the plot's smaller side at 100% (Office's default).
    expect(area[1].r * 2).toBeGreaterThan(40); expect(area[1].r * 2).toBeLessThan(75)
    // showNegBubbles=0 explicit: same as the default.
    expect(filled(circles(parse(bubbleXml({ sizes: [1, 4, 0, -4], showNeg: '0' }))))).toEqual(area)
    // showNegBubbles=1: the negative bubble is drawn at |size|, unfilled, in the series line colour.
    const shown = circles(parse(bubbleXml({ sizes: [1, 4, 0, -4], showNeg: '1' })))
    const negative = shown.filter(s => Math.abs(s.x - Math.max(...area.map(a => a.x)) - (area[1].x - area[0].x) * 2) < 1)
    expect(filled(shown)).toHaveLength(2)
    expect(negative.filter(s => s.op === 'fill')).toHaveLength(0)
    expect(negative.filter(s => s.op === 'stroke').map(s => s.style)).toEqual(['#0000ff'])
    expect(negative[0].r).toBeCloseTo(area[1].r, 6)
    expect(shown.filter(s => s.r > 0 && Math.abs(s.x - (area[0].x + (area[1].x - area[0].x) * 2)) < 1)).toHaveLength(0) // zero: still none
    // A size that is not a value (blank, error) draws no bubble.
    expect(filled(circles(parse(bubbleXml({ sizes: [1, undefined, '#N/A', 4] }))))).toHaveLength(2)
  })

  it('bubble x and y take independent value axes, and the sizes are a required data role', async () => {
    const model = parse(bubbleXml({ sizes: [1, 4] }))
    const prepared = model.preparedAxes!
    const x = prepared.axes.find(a => a.id === '1')!, y = prepared.axes.find(a => a.id === '2')!
    expect([x.min, x.max]).toEqual([0, 2])
    expect([y.min, y.max]).toEqual([0, 3])
    // DOCX/PPTX: a bubble series with no bubbleSize at all (no saved data, no reference) is unavailable.
    const pkg = await OfficePackage.load(buildChartDocx([]).bytes)
    for (const partPath of ['/word/charts/chart1.xml', '/ppt/charts/chart1.xml']) {
      const context = chartParseContext(documentChartContext('desktop'), pkg, partPath)
      const missing = await parseChartPart(parseXmlOrdered(bubbleXml({ sizes: [1, 4], sizeNode: '' })), undefined, context)
      expect(missing).toMatchObject({ kind: 'unsupported', reason: CHART_DATA_UNAVAILABLE })
      // A cacheless reference with no embedded workbook is unavailable too.
      const dangling = await parseChartPart(parseXmlOrdered(bubbleXml({ sizes: [1, 4], sizeNode: '<c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef></c:bubbleSize>' })), undefined, context)
      expect(dangling).toMatchObject({ kind: 'unsupported', reason: CHART_DATA_UNAVAILABLE })
      const ok = await parseChartPart(parseXmlOrdered(bubbleXml({ sizes: [1, 4] })), undefined, context)
      expect(ok.kind).toBe('native')
    }
    // DOCX/PPTX: cacheless x, y and bubbleSize references are read from the embedded workbook.
    const ref = (f: string) => `<c:numRef><c:f>${f}</c:f></c:numRef>`
    const embeddedXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${bubbleXml({ sizes: [1, 4], sizeNode: `<c:bubbleSize>${ref('Sheet1!$C$2:$C$3')}</c:bubbleSize>` })
      .replace(/<c:xVal>.*?<\/c:xVal>/, `<c:xVal>${ref('Sheet1!$A$2:$A$3')}</c:xVal>`).replace(/<c:yVal>.*?<\/c:yVal>/, `<c:yVal>${ref('Sheet1!$B$2:$B$3')}</c:yVal>`)
      .replace('</c:chart>', '</c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>')}`
    const workbook = embeddedWorkbook({ sheets: [{ name: 'Sheet1', data: row(2, cell('A2', 1), cell('B2', 2), cell('C2', 9)) + row(3, cell('A3', 2), cell('B3', 3), cell('C3', 16)) }] })
    for (const format of ['docx', 'pptx'] as const) {
      const fixture = format === 'docx' ? buildChartDocx([{ chartXml: embeddedXml, workbook }]) : buildChartPptx([{ chartXml: embeddedXml, workbook }])
      const doc = documentChartContext('desktop'), opened = await OfficePackage.load(fixture.bytes)
      const charts = format === 'docx'
        ? (await parseDocx(opened, { chartContext: doc })).sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart] : [])
        : (await parsePptx(opened, { chartContext: doc })).slides.flatMap(s => s.shapes).flatMap(s => s.content?.kind === 'xlsxChart' ? [s.content.chart] : [])
      expect(charts[0].unsupported, format).toBeUndefined()
      const embedded = charts[0].groups[0].series[0]
      expect([embedded.xValues, embedded.values, embedded.bubbleSizes].map(p => pointList(p!).map(q => q.state === 'value' ? q.v : q.state)), format).toEqual([[1, 2], [2, 3], [9, 16]])
      const bubbles = filled(circles(charts[0]))
      expect(bubbles, format).toHaveLength(2)
      expect(bubbles[1].r / bubbles[0].r, format).toBeCloseTo(4 / 3, 6)
    }
    // XLSX: a cacheless bubbleSize reference is filled from the worksheet cells like any other numeric role.
    const workspace = createChartWorkspace('desktop')
    const pending = parseXlsxChart(parseXmlOrdered(bubbleXml({ sizes: [1, 4], sizeNode: '<c:bubbleSize><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef></c:bubbleSize>' })), undefined)!
    expect(pending.groups[0].series[0].pending?.bubbleSizes).toMatchObject({ sheet: 'Sheet1', r0: 1, r1: 2, c0: 2, c1: 2 })
    const cells: Record<string, number> = { '1,2': 9, '2,2': 16 }
    const filledModel = fillChartRanges(pending, 'Sheet1', (_sheet, row, col) => cells[`${row},${col}`] !== undefined ? { value: cells[`${row},${col}`] } : undefined, { workspace })
    expect(pointList(filledModel.groups[0].series[0].bubbleSizes!).map(p => p.state === 'value' ? p.v : p.state)).toEqual([9, 16])
    const drawn = filled(circles(prepareChartAxes(filledModel)))
    expect(drawn).toHaveLength(2)
    expect(drawn[1].r / drawn[0].r).toBeCloseTo(4 / 3, 6)
  })
})

interface Op { op: 'M' | 'L' | 'Z'; x: number; y: number }
interface Painted { paths: Array<{ kind: 'fill' | 'stroke'; ops: Op[] }>; texts: Array<{ text: string; x: number; y: number }>; markers: number }
/** The red series paths (stroked or filled), the text drawn, and the red marker circles. */
function radarPaint(model: ChartModel, w = 400, h = 400): Painted {
  const ctx = createCanvas(w, h).getContext('2d'), out: Painted = { paths: [], texts: [], markers: 0 }
  let ops: Op[] = [], arcs = 0
  const begin = ctx.beginPath.bind(ctx), move = ctx.moveTo.bind(ctx), line = ctx.lineTo.bind(ctx), close = ctx.closePath.bind(ctx), arc = ctx.arc.bind(ctx)
  const stroke = ctx.stroke.bind(ctx), fill = ctx.fill.bind(ctx), text = ctx.fillText.bind(ctx)
  ctx.beginPath = () => { ops = []; arcs = 0; begin() }
  ctx.moveTo = (x: number, y: number) => { ops.push({ op: 'M', x, y }); move(x, y) }
  ctx.lineTo = (x: number, y: number) => { ops.push({ op: 'L', x, y }); line(x, y) }
  ctx.closePath = () => { ops.push({ op: 'Z', x: NaN, y: NaN }); close() }
  ctx.arc = (x: number, y: number, r: number, a0: number, a1: number, ccw?: boolean) => { arcs++; arc(x, y, r, a0, a1, ccw) }
  ctx.stroke = ((...args: Parameters<typeof ctx.stroke>) => { if (String(ctx.strokeStyle) === '#ff0000' && ops.length) out.paths.push({ kind: 'stroke', ops: [...ops] }); stroke(...args) }) as typeof ctx.stroke
  ctx.fill = ((...args: Parameters<typeof ctx.fill>) => { if (String(ctx.fillStyle) === '#ff0000') { if (arcs) out.markers += arcs; else if (ops.length) out.paths.push({ kind: 'fill', ops: [...ops] }) } fill(...args) }) as typeof ctx.fill
  ctx.fillText = (t: string, x: number, y: number, max?: number) => { out.texts.push({ text: t, x, y }); if (max === undefined) text(t, x, y); else text(t, x, y, max) }
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  return out
}
function radarXml(values: Array<number | string | undefined>, style: 'standard' | 'marker' | 'filled' | undefined, options = '', marker = ''): string {
  const cats = values.map((_, i) => String.fromCharCode(65 + i))
  return space(`<c:radarChart>${style ? `<c:radarStyle val="${style}"/>` : ''}<c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>R</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>${marker}<c:cat>${strLit(cats)}</c:cat><c:val>${lit(values)}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:radarChart>${RADAR_AXES}`, options)
}
const shape = (ops: Op[]) => ops.map(o => o.op)

describe('radar (spec 4.1)', () => {
  it('radar: clockwise from 12, radial axis, standard, marker and filled styles', () => {
    const model = parse(radarXml([4, 2, 4, 2], 'standard'))
    expect(model.unsupported).toBeUndefined()
    expect(model.groups[0]).toMatchObject({ family: 'radar', radarStyle: 'standard', binding: { kind: 'radial', value: '2' } })
    // The radial (value) axis: auto scale 0..4 every 1.
    const radial = model.preparedAxes!.axes.find(a => a.id === '2')!
    expect([radial.min, radial.max, radial.major]).toEqual([0, 4, 1])
    const standard = radarPaint(model)
    expect(standard.markers).toBe(0)
    expect(standard.paths.filter(p => p.kind === 'fill')).toHaveLength(0)
    const [series] = standard.paths.filter(p => p.kind === 'stroke')
    const points = series.ops.filter(o => o.op !== 'Z')
    // Category 0 at 12 o'clock, then clockwise: 3 o'clock, 6 o'clock, 9 o'clock; then the polygon closes.
    const [top, right, bottom, left] = points
    const cx = top.x, cy = right.y
    expect(right.x).toBeGreaterThan(cx); expect(top.y).toBeLessThan(cy)
    expect(bottom.x).toBeCloseTo(cx, 6); expect(left.y).toBeCloseTo(cy, 6); expect(left.x).toBeLessThan(cx)
    const R = cy - top.y
    expect(bottom.y - cy).toBeCloseTo(R, 6)
    expect(right.x - cx).toBeCloseTo(R / 2, 6) // value 2 on a 0..4 radial axis is half the radius
    expect(cx - left.x).toBeCloseTo(R / 2, 6)
    expect(points[4]).toMatchObject({ x: expect.closeTo(top.x, 6), y: expect.closeTo(top.y, 6) }) // closed back to category 0
    // Category labels sit outside each spoke; radial tick labels run up the 12 o'clock spoke.
    const label = (t: string) => standard.texts.find(x => x.text === t)!
    expect(label('A').y).toBeLessThan(top.y); expect(label('B').x).toBeGreaterThan(right.x + R / 2)
    expect(label('C').y).toBeGreaterThan(bottom.y); expect(label('D').x).toBeLessThan(cx - R)
    for (const [t, f] of [['0', 0], ['2', 0.5], ['4', 1]] as const) expect(label(t).y).toBeCloseTo(cy - f * R, 0)
    // Marker style: the same polygon plus a marker at each point; filled: a filled polygon, no markers.
    const marker = radarPaint(parse(radarXml([4, 2, 4, 2], 'marker')))
    expect(marker.markers).toBe(4)
    expect(marker.paths.filter(p => p.kind === 'stroke')[0].ops.filter(o => o.op !== 'Z').slice(0, 4)).toEqual(points.slice(0, 4))
    // Excel's plain radar is radarStyle=marker with symbol none: no markers.
    expect(radarPaint(parse(radarXml([4, 2, 4, 2], 'marker', '', '<c:marker><c:symbol val="none"/></c:marker>'))).markers).toBe(0)
    const filledRadar = radarPaint(parse(radarXml([4, 2, 4, 2], 'filled')))
    expect(filledRadar.markers).toBe(0)
    const area = filledRadar.paths.filter(p => p.kind === 'fill')
    expect(area).toHaveLength(1)
    expect(area[0].ops.filter(o => o.op !== 'Z').slice(0, 4)).toEqual(points.slice(0, 4).map(p => ({ ...p, x: expect.closeTo(p.x, 6), y: expect.closeTo(p.y, 6) })))
    // The default radarStyle is standard.
    expect(parse(radarXml([1, 2, 3], undefined)).groups[0].radarStyle).toBe('standard')
  })

  it('radar is native through the DOCX and PPTX adapters, with chart coverage', async () => {
    for (const format of ['docx', 'pptx'] as const) {
      const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${radarXml([4, 2, 4, 2], 'marker')}`
      const bubble = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${bubbleXml({ sizes: [1, 4] })}`
      const fixture = format === 'docx' ? buildChartDocx([{ chartXml }, { chartXml: bubble }]) : buildChartPptx([{ chartXml }, { chartXml: bubble }])
      const pkg = await OfficePackage.load(fixture.bytes)
      const doc = documentChartContext('desktop')
      let charts: ChartModel[], coverage: DrawingCoverageEntry[]
      if (format === 'docx') {
        const parsed = await parseDocx(pkg, { chartContext: doc })
        charts = parsed.sections.flatMap(s => s.paragraphs).flatMap(p => p.images ?? []).flatMap(i => i.drawing?.kind === 'xlsxChart' ? [i.drawing.chart] : [])
        coverage = (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original')
      } else {
        const parsed = await parsePptx(pkg, { chartContext: doc })
        charts = parsed.slides.flatMap(s => s.shapes).flatMap(s => s.content?.kind === 'xlsxChart' ? [s.content.chart] : [])
        coverage = (parsed.drawingCoverage ?? []).filter(e => e.scope === 'original')
      }
      expect(charts.map(c => [c.unsupported, c.groups[0]?.family]), format).toEqual([[undefined, 'radar'], [undefined, 'bubble']])
      expect(coverage.map(c => c.selectedRepresentation), format).toEqual(['chart', 'chart'])
      expect(radarPaint(charts[0]).markers, format).toBe(4)
      expect(filled(circles(charts[1])), format).toHaveLength(2)
      expect([chartDescription(charts[0]), chartDescription(charts[1])].map(d => d.split(',')[0]), format).toEqual(['Chart. Radar chart', 'Chart. Bubble chart'])
    }
    // A radar combined with a planar group is a card (never drawn without its other group).
    const combo = parse(space(`<c:radarChart><c:radarStyle val="marker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val>${lit([1, 2, 3])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:radarChart><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="1"/><c:order val="1"/><c:val>${lit([1, 2, 3])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>${RADAR_AXES}`))
    expect(combo.unsupported).toBe('Combination radar charts are not supported')
  })
})

describe('radar cells of the point matrix (spec 5.8)', () => {
  const states: ChartPoint[] = [{ state: 'blank' }, { state: 'error', code: '#N/A' }, { state: 'error', code: '#DIV/0!' }, { state: 'unavailable' }]
  for (const mode of ['gap', 'zero', 'span'] as const) it(`radar ${mode}: blank, #N/A, other error and unavailable follow the line row`, () => {
    const blank = mode === 'zero' ? 0 : mode
    expect(states.map(p => plotPoint(p, 'radar', 'standard', mode, false))).toEqual([blank, 'span', 'span', 'gap'])
    expect(plotPoint(states[1], 'radar', 'standard', mode, true)).toBe(blank) // dispNaAsBlank: #N/A as blank
    // Native stroke oracle: [2, X, 4, 3] around the polygon.
    const drawn = (middle: string | undefined) => radarPaint(parse(radarXml([2, middle, 4, 3], 'standard', `<c:dispBlanksAs val="${mode}"/>`))).paths.filter(p => p.kind === 'stroke').map(p => shape(p.ops))
    expect(drawn(undefined)).toEqual([mode === 'gap' ? ['M', 'M', 'L', 'L'] : mode === 'zero' ? ['M', 'L', 'L', 'L', 'L'] : ['M', 'L', 'L', 'L']])
    expect(drawn('#N/A')).toEqual([['M', 'L', 'L', 'L']]) // #N/A connects across, whatever dispBlanksAs says
    expect(drawn('#DIV/0!')).toEqual([['M', 'L', 'L', 'L']])
  })
  it('radar: zero puts the point at the centre; an unavailable point is an unconditional gap with its diagnostic', () => {
    const zero = radarPaint(parse(radarXml([2, undefined, 4, 3], 'standard', '<c:dispBlanksAs val="zero"/>'))).paths[0].ops
    const centre = radarPaint(parse(radarXml([4, 4, 4, 4], 'standard'))).paths[0].ops
    expect(zero[1].x).toBeCloseTo(centre[0].x, 6) // the centre's x is category 0's x
    const model = parse(radarXml([2, 9, 4, 3], 'standard', '<c:dispBlanksAs val="span"/>'))
    const p = new PointWriter(4); p.value(0, 2); p.unavailable(1); p.value(2, 4); p.value(3, 3)
    model.groups[0].series[0].values = p.done()
    const prepared = prepareChartAxes({ ...model, preparedAxes: undefined })
    expect(prepared.preparedAxes!.diagnostics).toContainEqual(expect.objectContaining({ code: 'chart-point-unavailable' }))
    expect(radarPaint(prepared).paths.map(x => shape(x.ops))).toEqual([['M', 'M', 'L', 'L']])
  })
})

/** Every family's cases: no series is a card; every other case is drawn natively. */
const NATIVE: Partial<Record<ConformanceCase, { kind: 'native' | 'card' }>> = {
  'no series': { kind: 'card' }, 'empty series': { kind: 'native' }, 'all blank': { kind: 'native' }, 'NaN and infinity': { kind: 'native' },
  'negative values': { kind: 'native' }, 'single point': { kind: 'native' }, 'unavailable point': { kind: 'native' },
}
for (const family of CLASSIC_FAMILIES) describeFamilyConformance({ ...family, expect: NATIVE })
