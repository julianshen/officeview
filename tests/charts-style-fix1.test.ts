// @vitest-environment node
/**
 * A4a fix round 1: colour precedence (spPr, chart style part, colour part, c:style, accent cycle), area drop lines,
 * dPt invertIfNegative, source-series alignment, plot-area order, unchanged defaults, path admission and charging.
 * Packages here are Office-shaped synthetic fixtures (tests/helpers builders), never Office evidence.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import type JSZip from 'jszip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsxChart } from '../src/charts/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { layoutChart, releaseChartLayouts } from '../src/charts/layout'
import { documentChartContext } from '../src/charts/context'
import { createChartWorkspace } from '../src/charts/workspace'
import { PointWriter } from '../src/charts/points'
import { resolveDrawingColor, type ColorDefinition } from '../src/drawing/style'
import { parseXlsx } from '../src/xlsx/parse'
import type { ChartModel } from '../src/charts/model'
import { chartWorkbook, chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { OfficePackage } from '../src/core/zip'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { paintMarker, type MarkerSymbol } from '../src/charts/markers'
import { resolveSeriesColor } from '../src/charts/style'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle'
const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const cats = (n: number) => `<c:cat><c:strLit><c:ptCount val="${n}"/>${Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>C${i}</c:v></c:pt>`).join('')}</c:strLit></c:cat>`
const ser = (i: number, values: number[], extra = '', name = `S${i}`, tag = 'val') => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${name}</c:v></c:tx>${extra}${tag === 'val' ? cats(values.length) : ''}<c:${tag}>${lit(values)}</c:${tag}></c:ser>`
const AX = '<c:axId val="1"/><c:axId val="2"/>'
const AXES = '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/><c:majorGridlines/></c:valAx>'
const space = (plot: string, opts: { root?: string; legend?: boolean; title?: string } = {}) =>
  `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}">${opts.root ?? ''}<c:chart>${opts.title ?? '<c:autoTitleDeleted val="1"/>'}<c:plotArea>${plot}</c:plotArea>${opts.legend === false ? '' : '<c:legend><c:legendPos val="r"/></c:legend>'}</c:chart></c:chartSpace>`
const bars = (series: string, extra = '') => `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${extra}${series}${AX}</c:barChart>${AXES}`
const solid = (hex: string) => `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>`
const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const srgb = (value: string, ...transforms: Array<[string, number]>): string => {
  const c = resolveDrawingColor({ kind: 'srgb', value, transforms: transforms.map(([name, v]) => ({ name, value: String(v) })) } as ColorDefinition)!
  return [c.r, c.g, c.b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()
}
function pixels(model: ChartModel, w = 400, h = 260) {
  const canvas = createCanvas(w, h), ctx = canvas.getContext('2d')
  paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, w, h)
  const data = ctx.getImageData(0, 0, w, h).data
  const at = (x: number, y: number) => { const i = (y * w + x) * 4; return [data[i], data[i + 1], data[i + 2], data[i + 3]] }
  const count = (hex: string, tolerance = 8) => {
    const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16))
    let n = 0
    for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i] - r) <= tolerance && Math.abs(data[i + 1] - g) <= tolerance && Math.abs(data[i + 2] - b) <= tolerance && data[i + 3] > 200) n++
    return n
  }
  /** Columns holding at least `min` pixels of the colour. */
  const columns = (hex: string, min: number) => {
    const [r, g, b] = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16)), out: number[] = []
    for (let x = 0; x < w; x++) { let n = 0; for (let y = 0; y < h; y++) { const [pr, pg, pb, pa] = at(x, y); if (Math.abs(pr - r) <= 8 && Math.abs(pg - g) <= 8 && Math.abs(pb - b) <= 8 && pa > 200) n++ } if (n >= min) out.push(x) }
    return out
  }
  return { at, count, columns }
}
/** A real canvas whose fills and strokes are recorded with the state in force. */
function recorder(w = 400, h = 260) {
  const canvas = createCanvas(w, h), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const style = (s: unknown) => typeof s === 'string' ? s.toUpperCase() : (s as object).constructor.name
  const rects: Array<{ x: number; y: number; w: number; h: number; style: string }> = []
  const strokes: Array<{ style: string; width: number; cap: string; join: string; dash: number[] }> = []
  const arcs: Array<{ r: number; sweep: number }> = []
  let rectCalls = 0
  const fillRect = ctx.fillRect.bind(ctx), stroke = ctx.stroke.bind(ctx), arc = ctx.arc.bind(ctx), rect = ctx.rect.bind(ctx)
  ctx.fillRect = (x, y, rw, rh) => { rects.push({ x, y, w: rw, h: rh, style: style(ctx.fillStyle) }); fillRect(x, y, rw, rh) }
  ctx.stroke = (...a: []) => { strokes.push({ style: style(ctx.strokeStyle), width: ctx.lineWidth, cap: ctx.lineCap, join: ctx.lineJoin, dash: ctx.getLineDash() }); stroke(...a) }
  ctx.arc = (x, y, r, s, e, ccw) => { arcs.push({ r, sweep: Math.abs(e - s) }); arc(x, y, r, s, e, ccw) }
  ctx.rect = (x, y, rw, rh) => { rectCalls++; rect(x, y, rw, rh) }
  return { ctx, rects, strokes, arcs, get rectCalls() { return rectCalls } }
}

const RELS = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/></Relationships>'
const COLORS = `<cs:colorStyle xmlns:cs="${CS}" xmlns:a="${A}" meth="cycle" id="10"><a:srgbClr val="FF0000"/><a:srgbClr val="00FF00"/><cs:variation/><cs:variation><a:lumMod val="50000"/></cs:variation></cs:colorStyle>`
const style = (dataPoint: string, dataPointLine = '') => `<cs:chartStyle xmlns:cs="${CS}" xmlns:a="${A}" id="201">${dataPoint ? `<cs:dataPoint>${dataPoint}</cs:dataPoint>` : ''}${dataPointLine ? `<cs:dataPointLine>${dataPointLine}</cs:dataPointLine>` : ''}</cs:chartStyle>`
const AUTO_PH = '<cs:lnRef idx="0"/><cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:effectRef idx="0"/><cs:fontRef idx="minor"/><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>'
/** The chart's series colours read through the production XLSX path (rels -> style and colour parts -> parse). */
async function packagedColors(chartXml: string, styleXml: string | undefined, colorsXml: string | null = COLORS): Promise<string[]> {
  const pkg = await chartWorkbook(chartXml, {
    patch: (zip: JSZip) => {
      zip.file('xl/charts/_rels/chart1.xml.rels', RELS)
      if (styleXml) zip.file('xl/charts/style1.xml', styleXml)
      if (colorsXml) zip.file('xl/charts/colors1.xml', colorsXml)
    },
  })
  const content = (await parseXlsx(pkg)).sheets[0].drawings?.[0].content
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  return content.chart.groups.flatMap(g => g.series.map(s => s.color))
}
const opened: Array<{ release(): Promise<void> }> = []
afterEach(async () => { for (const source of opened.splice(0)) await source.release() })
/** The same workbook read resident (parseXlsx) and paged (the import session's drawing read): the chart models. */
async function residentAndPaged(bytes: Uint8Array): Promise<[ChartModel, ChartModel]> {
  const resident = (await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })).sheets[0].drawings?.[0].content
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'A4a', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  opened.push(source)
  const store = source.store('1')
  const read = await store.readDrawings!({ stamp: store.stamp, signal: new AbortController().signal }, 'D')
  const paged = read.drawings[0].content
  if (resident?.kind !== 'xlsxChart' || paged?.kind !== 'xlsxChart') throw new Error('no chart')
  return [resident.chart, paged.chart]
}
const four = (extra = (_i: number) => '') => [0, 1, 2, 3].map(i => ser(i, [3, 5], extra(i))).join('')

describe('A4a fix round 1: colour precedence (spec 5.11)', () => {
  it('I1: the colour part cycles its colours, each later cycle taking its next variation', async () => {
    expect(await packagedColors(space(bars(four())), style(AUTO_PH))).toEqual(['FF0000', '00FF00', srgb('FF0000', ['lumMod', 50000]), srgb('00FF00', ['lumMod', 50000])])
    // Without a style part the colour part still decides; without either the accent cycle does.
    expect(await packagedColors(space(bars(four())), undefined)).toEqual(['FF0000', '00FF00', '800000', '008000'])
    expect(await packagedColors(space(bars(four())), undefined, null)).toEqual(['4472C4', 'ED7D31', 'A5A5A5', 'FFC000'])
  })

  it('I1: the paged XLSX read bakes the same colours from the parts as the resident parse (D0 parity)', async () => {
    const bytes = await chartWorkbookBytes(space(bars(four())), { patch: (zip: JSZip) => {
      zip.file('xl/charts/_rels/chart1.xml.rels', RELS)
      zip.file('xl/charts/style1.xml', style(AUTO_PH))
      zip.file('xl/charts/colors1.xml', COLORS)
    } })
    const [resident, paged] = await residentAndPaged(bytes)
    const colors = (model: ChartModel) => model.groups.flatMap(g => g.series.map(s => s.color))
    expect(colors(resident)).toEqual(['FF0000', '00FF00', '800000', '008000'])
    expect(colors(paged)).toEqual(colors(resident))
  })

  it('I1: the chart style part wins over the colour part: an explicit colour, a styleClr index, phClr transforms; spPr wins over both', async () => {
    const xml = space(bars(four()))
    expect(await packagedColors(xml, style('<cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cs:spPr>'))).toEqual(['123456', '123456', '123456', '123456'])
    expect(await packagedColors(xml, style('<cs:fillRef idx="1"><cs:styleClr val="1"/></cs:fillRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>'))).toEqual(['00FF00', '00FF00', '00FF00', '00FF00'])
    expect(await packagedColors(xml, style('<cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></cs:spPr>')))
      .toEqual(['800000', '008000', srgb('800000', ['lumMod', 50000]), srgb('008000', ['lumMod', 50000])])
    expect(await packagedColors(space(bars(four(i => i === 1 ? solid('0000FF') : ''))), style('<cs:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cs:spPr>'))).toEqual(['123456', '0000FF', '123456', '123456'])
    // Line series read dataPointLine (its line colour), bars dataPoint.
    const combo = space(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [3, 5])}${AX}</c:barChart><c:lineChart><c:grouping val="standard"/>${ser(1, [2, 4])}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`)
    expect(await packagedColors(combo, style('<cs:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></cs:spPr>', '<cs:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></cs:spPr>'))).toEqual(['123456', 'ABCDEF'])
  })

  it('I1/M2: c:style picks the gallery palette, read directly or through mc:AlternateContent (c14:style = 100 + style)', () => {
    const three = [0, 1, 2].map(i => ser(i, [3, 5])).join('')
    const colors = (root: string) => parse(space(bars(three), { root })).groups[0].series.map(s => s.color)
    // Style 3: accent1 only, the leading series shaded, the trailing tinted, over (-70%, 70%) (LibreOffice oox).
    const mono = ['4472C4', '4472C4', '4472C4'].map((c, i) => [srgb(c, ['shade', 65000]), c, srgb(c, ['tint', 65000])][i])
    expect(colors('<c:style val="3"/>')).toEqual(mono)
    expect(colors(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="103"/></mc:Choice></mc:AlternateContent>`)).toEqual(mono)
    expect(colors(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="3"/></mc:Fallback></mc:AlternateContent>`)).toEqual(mono)
    // Style 2 (and its rows: 10, 18 ...) is the six-accent cycle.
    expect(colors('<c:style val="2"/>')).toEqual(['4472C4', 'ED7D31', 'A5A5A5'])
    expect(colors('<c:style val="10"/>')).toEqual(['4472C4', 'ED7D31', 'A5A5A5'])
  })

  it('I1: the accent cycle takes Office\'s default variations after six series (exact values)', () => {
    const many = Array.from({ length: 13 }, (_, i) => ser(i, [3, 5])).join('')
    const colors = parse(space(bars(many))).groups[0].series.map(s => s.color)
    expect(colors.slice(0, 6)).toEqual(['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'])
    expect(colors[6]).toBe(srgb('4472C4', ['lumMod', 60000]))
    expect(colors[7]).toBe(srgb('ED7D31', ['lumMod', 60000]))
    expect(colors[12]).toBe(srgb('4472C4', ['lumMod', 80000], ['lumOff', 20000]))
    // varyColors points follow the same cycle.
    const pie = parse(space(`<c:pieChart><c:varyColors val="1"/>${ser(0, [1, 2, 3, 4, 5, 6, 7, 8])}</c:pieChart>`)).groups[0].series[0].pointColors!
    expect(pie[6]).toBe(srgb('4472C4', ['lumMod', 60000]))
  })

  it('chart area: noFill is transparent; DOCX/PPTX take the c:style default (dark row: dk1); XLSX stays white', () => {
    const xml = (root: string) => space(bars(ser(0, [3, 5])), { root })
    const at = (partPath: string, root: string) => {
      const model = parseXlsxChart(parseXmlOrdered(xml(root)), undefined, { pkg: {} as never, partPath, profile: 'desktop', workspace: createChartWorkspace('desktop') })!
      return pixels(model).at(3, 3)
    }
    expect(at('word/charts/chart1.xml', '<c:style val="42"/>')).toEqual([0, 0, 0, 255])
    expect(at('ppt/charts/chart1.xml', '<c:style val="42"/>')).toEqual([0, 0, 0, 255])
    expect(at('ppt/charts/chart1.xml', '<c:style val="2"/>')).toEqual([255, 255, 255, 255])
    expect(at('xl/charts/chart1.xml', '<c:style val="42"/>')).toEqual([255, 255, 255, 255])
    expect(at('word/charts/chart1.xml', '<c:style val="42"/><c:spPr><a:noFill/></c:spPr>')[3]).toBe(0)
  })
})

describe('A4a fix round 1: series and group options (spec 5.4, 5.8)', () => {
  it('I2: area groups draw their drop lines, once per category from the plotted points to the axis', () => {
    const drop = '<c:dropLines><c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:dropLines>'
    const area = (lines: string) => parse(space(`<c:areaChart><c:grouping val="standard"/>${ser(0, [3, 5, 4, 6], solid('0000FF'))}${lines}${AX}</c:areaChart>${AXES}`, { legend: false }))
    expect(pixels(area('')).count('FF0000')).toBe(0)
    const columns = pixels(area(drop)).columns('FF0000', 20)
    // Four categories: four vertical drop lines (a line may straddle two pixel columns).
    const groups = columns.filter((x, i) => i === 0 || x - columns[i - 1] > 2)
    expect(groups.length).toBe(4)
  })

  it('I3: a dPt invertIfNegative overrides its series; an inverted negative is white, or the c14 invertSolidFillFmt colour', () => {
    const fills = (serExtra: string, ptInvert?: string) => {
      const dPt = ptInvert === undefined ? '' : `<c:dPt><c:idx val="1"/><c:invertIfNegative val="${ptInvert}"/></c:dPt>`
      const model = parse(space(bars(ser(0, [3, -2], `${solid('0000FF')}${serExtra}${dPt}`)), { legend: false }))
      const r = recorder()
      paintXlsxChart(model, r.ctx, 400, 260)
      return r.rects.filter(rect => !(rect.x === 0 && rect.y === 0)).map(rect => rect.style)
    }
    expect(fills('<c:invertIfNegative val="0"/>', '1')).toEqual(['#0000FF', '#FFFFFF'])
    expect(fills('<c:invertIfNegative val="1"/>', '0')).toEqual(['#0000FF', '#0000FF'])
    expect(fills('<c:invertIfNegative val="1"/>')).toEqual(['#0000FF', '#FFFFFF'])
    const c14 = '<c:extLst><c:ext uri="{6F2FDCE9-48DA-4B69-8628-5D25D57E0C9C}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:invertSolidFillFmt><c14:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c14:spPr></c14:invertSolidFillFmt></c:ext></c:extLst>'
    expect(fills(`<c:invertIfNegative val="1"/>${c14}`)).toEqual(['#0000FF', '#00FF00'])
  })

  it('I4: formatting follows the drawn series, not its view position or its name', () => {
    // The first series is hidden (plotVisOnly): the second keeps its own fill.
    const model = parse(space(bars(ser(0, [3, 5, 4], solid('FF0000')) + ser(1, [2, 4, 3], solid('0000FF'))), { legend: false }))
    model.groups[0].series[0].values.hidden = Uint32Array.from([0, 1, 2])
    const p = pixels(model)
    expect(p.count('FF0000')).toBe(0)
    expect(p.count('0000FF')).toBeGreaterThan(100)
    // Two line series with the same name: each strokes with its own line.
    const ln = (attrs: string, hex: string) => `<c:spPr><a:ln ${attrs}><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`
    const lines = parse(space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 2, 3], ln('w="76200"', 'FF0000'), 'Same')}${ser(1, [3, 2, 1], ln('', '0000FF'), 'Same')}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`, { legend: false }))
    const r = recorder()
    paintXlsxChart(lines, r.ctx, 400, 260)
    const series = r.strokes.filter(s => s.style === '#FF0000' || s.style === '#0000FF')
    expect(series).toEqual([expect.objectContaining({ style: '#FF0000', width: 8 }), expect.objectContaining({ style: '#0000FF', width: 2.25 })])
  })

  it('I5: a plot-area fill sits under the gridlines and axis lines', () => {
    const chart = (plotSpPr: string) => parse(space(bars(ser(0, [3, 5, 4])) + plotSpPr, { legend: false }))
    const plain = pixels(chart('')), filled = pixels(chart('<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr>'))
    expect(plain.count('D9D9D9', 2)).toBeGreaterThan(500)
    expect(filled.count('D9D9D9', 2)).toBe(plain.count('D9D9D9', 2))
    expect(filled.count('BFBFBF', 2)).toBe(plain.count('BFBFBF', 2))
  })

  it('M1: exploded slices shrink the pie to stay in its box; an exploded doughnut is drawn as ring segments, not over a painted hole', () => {
    const pie = (family: string, extra = '') => parse(space(`<c:${family}><c:varyColors val="1"/>${ser(0, [1, 1, 1, 1], '<c:explosion val="100"/>')}${extra}</c:${family}>`, { legend: false }))
    const w = 300, h = 300, p = pixels(pie('pieChart'), w, h)
    let edge = 0
    for (let k = 0; k < w; k++) for (const [x, y] of [[k, 2], [k, h - 3], [2, k], [w - 3, k]]) { const [r, g, b] = p.at(x, y); if (!(r > 240 && g > 240 && b > 240) && !(r === 217 && g === 217 && b === 217)) edge++ }
    expect(edge).toBe(0)
    const r = recorder(w, h)
    paintXlsxChart(pie('doughnutChart', '<c:holeSize val="50"/>'), r.ctx, w, h)
    expect(r.arcs.filter(a => Math.abs(a.sweep - Math.PI * 2) < 1e-9)).toEqual([])
    // Exploded slices are clamped (0-400 %).
    expect(parse(space(`<c:pieChart>${ser(0, [1, 1], '<c:explosion val="9000"/>')}</c:pieChart>`)).groups[0].series[0].explosion).toBe(400)
  })

  it('M3: an authored marker size or spPr does not switch markers on when the group or scatter style has them off', () => {
    const model = parse(space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 2, 3], '<c:marker><c:size val="7"/></c:marker>')}<c:marker val="0"/>${AX}</c:lineChart>${AXES}`))
    expect(model.groups[0].series[0].markers).toBe(false)
    const scatter = parse(space(`<c:scatterChart><c:scatterStyle val="line"/>${ser(0, [1, 2, 3], '<c:marker><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:marker>', 'S', 'yVal')}${AX}</c:scatterChart><c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`))
    expect(scatter.groups[0].series[0].markers).toBe(false)
  })

  it('M4: an authored bar border is stroked; a noFill border is not', () => {
    const border = (ln: string) => pixels(parse(space(bars(ser(0, [3, 5], `<c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>${ln}</c:spPr>`)), { legend: false }))).count('00FF00', 40)
    expect(border('<a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>')).toBeGreaterThan(100)
    expect(border('<a:ln><a:noFill/></a:ln>')).toBe(0)
  })
})

describe('A4a fix round 1: unauthored defaults are unchanged (I6)', () => {
  const markerChart = (family: 'line' | 'scatter', marker = '') => parse(space(family === 'line'
    ? `<c:lineChart><c:grouping val="standard"/>${ser(0, [5, 5, 5, 5], `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${marker}`)}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`
    : `<c:scatterChart><c:scatterStyle val="lineMarker"/>${ser(0, [5, 5, 5, 5], marker, 'S', 'yVal')}${AX}</c:scatterChart><c:valAx><c:axId val="1"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`, { legend: false }))
  it('a series without c:marker keeps the filled circle (r 2.5 line, 3 scatter); an explicit auto symbol cycles per series', () => {
    for (const [family, radius] of [['line', 2.5], ['scatter', 3]] as const) {
      const r = recorder()
      paintXlsxChart(markerChart(family), r.ctx, 400, 260)
      expect(r.arcs.filter(a => a.r === radius).length).toBe(4)
      // The chart and plot clips are the only rects: no square markers.
      expect(r.rectCalls).toBe(2)
    }
    const auto = parse(space(`<c:lineChart><c:grouping val="standard"/>${[0, 1].map(i => ser(i, [5, 5], '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="auto"/></c:marker>')).join('')}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`, { legend: false }))
    const r = recorder()
    paintXlsxChart(auto, r.ctx, 400, 260)
    // Series 0 diamonds (paths), series 1 squares (rects): no circles.
    expect(r.arcs.filter(a => a.sweep > 6)).toEqual([])
    expect(r.rectCalls).toBe(2 + 2)
  })

  it('varyColors applies to a lone single-series bar group only', () => {
    const fills = (n: number) => {
      const model = parse(space(bars([0, 1].slice(0, n).map(i => ser(i, [3, 5, 4])).join(''), '<c:varyColors val="1"/>'), { legend: false }))
      const r = recorder()
      paintXlsxChart(model, r.ctx, 400, 260)
      return [...new Set(r.rects.filter(rect => !(rect.x === 0 && rect.y === 0)).map(rect => rect.style))]
    }
    expect(fills(2)).toEqual(['#4472C4', '#ED7D31'])
    expect(fills(1)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5'])
  })

  it('an a:ln without a colour strokes in the series colour; series strokes keep round joins; Office caps and joins are honoured', () => {
    const stroke = (spPr: string) => {
      const model = parse(space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 3, 2], `<c:spPr>${spPr}</c:spPr><c:marker><c:symbol val="none"/></c:marker>`)}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`, { legend: false }))
      const r = recorder()
      paintXlsxChart(model, r.ctx, 400, 260)
      return r.strokes.filter(s => s.style !== '#D9D9D9' && s.style !== '#BFBFBF').at(-1)!
    }
    expect(stroke('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="28575"/>')).toMatchObject({ style: '#FF0000', width: 3, join: 'round' })
    expect(stroke('<a:ln><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln>')).toMatchObject({ style: '#00AA00', width: 2.25, join: 'round', cap: 'butt' })
    expect(stroke('<a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:round/></a:ln>')).toMatchObject({ style: '#4472C4', width: 3, join: 'round', cap: 'round' })
    expect(stroke('<a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:bevel/></a:ln>')).toMatchObject({ width: 12700 / 12700 * 96 / 72, join: 'bevel' })
  })
})

describe('A4a fix round 1: admission and charging (I8)', () => {
  const segments = (model: ChartModel) => {
    const context = documentChartContext('desktop'); let n = 0
    const ctx = new Proxy({ font: '', measureText: () => ({ width: 10 }), getLineDash: () => [] }, { get: (t: any, k: any) => k in t ? t[k] : (..._args: any[]) => { if (k === 'lineTo' || k === 'moveTo') n++ }, set: (t: any, k: any, v: any) => { t[k] = v; return true } })
    paintXlsxChart(model, ctx as unknown as CanvasRenderingContext2D, 400, 260, undefined, context)
    const layout = layoutChart(model, 400, 260, 1, 0, { width: () => 10, lineHeight: () => 14 }, context)
    const limited = layout.diagnostics.some(d => d.code === 'chart-path-limit')
    releaseChartLayouts(model)
    return { n, limited }
  }
  const big = (xml: string, family: 'line' | 'bar', count: number) => {
    const model = parse(xml), points = new PointWriter(4096)
    for (let i = 0; i < 4096; i++) points.value(i, 1 + (i % 7))
    delete model.preparedAxes
    model.groups[0].family = family
    model.groups[0].series = Array.from({ length: count }, (_, i) => ({ name: `s${i}`, color: '112233', values: points.done() }))
    model.categories = []
    return model
  }
  it('drop lines and series lines are admitted against the path work bound', () => {
    const drops = segments(big(space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 2])}<c:dropLines/>${AX}</c:lineChart>${AXES}`), 'line', 64))
    expect(drops.n).toBeLessThanOrEqual(262144)
    expect(drops.limited).toBe(true)
    const serLines = segments(big(space(`<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:overlap val="100"/>${ser(0, [1, 2])}<c:serLines/>${AX}</c:barChart>${AXES}`), 'bar', 40))
    expect(serLines.n).toBeLessThanOrEqual(262144)
    expect(serLines.limited).toBe(true)
  })

  it('retained formatting is charged before it is built: a gradient keeps 32 stops, each charged', () => {
    const used = (stops: number) => {
      const workspace = createChartWorkspace('desktop')
      const gs = Array.from({ length: stops }, (_, i) => `<a:gs pos="${Math.round(i / Math.max(1, stops - 1) * 100000)}"><a:srgbClr val="00${(i % 256).toString(16).padStart(2, '0')}00"/></a:gs>`).join('')
      const model = parseXlsxChart(parseXmlOrdered(space(bars(ser(0, [3, 5])) + `<c:spPr><a:gradFill><a:gsLst>${gs}</a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr>`)), undefined, { pkg: {} as never, partPath: 'xl/charts/chart1.xml', profile: 'desktop', workspace })!
      return { used: workspace.used, stops: model.plotAreaFill?.kind === 'gradient' ? model.plotAreaFill.stops.length : 0 }
    }
    const two = used(2), three = used(3), many = used(32), huge = used(1000)
    expect(three.used - two.used).toBeGreaterThan(0)
    expect(many.used - two.used).toBe(30 * (three.used - two.used))
    expect(huge).toEqual(many)
  })
})

describe('A4a fix round 2', () => {
  const gradientLine = (stops: number) => `<a:ln w="25400"><a:gradFill><a:gsLst>${Array.from({ length: stops }, (_, i) => `<a:gs pos="${Math.round(i / Math.max(1, stops - 1) * 100000)}"><a:srgbClr val="0000${(i % 256).toString(16).padStart(2, '0')}"/></a:gs>`).join('')}</a:gsLst><a:lin ang="0"/></a:gradFill></a:ln>`
  const strokeChart = (stops: number) => space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 2, 3], `<c:spPr>${gradientLine(stops)}</c:spPr>`)}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`)
  const parseWith = (xml: string, room = Infinity) => {
    let held = 0, peak = 0
    const workspace = createChartWorkspace('desktop', { reserve: (b: number) => { if (held + b > room) return false; held += b; peak = Math.max(peak, held); return true }, release: (_t: string, b: number) => { held -= b } })
    const model = parseXlsxChart(parseXmlOrdered(xml), undefined, { pkg: {} as never, partPath: 'xl/charts/chart1.xml', profile: 'desktop', workspace })!
    return { model, used: workspace.used, peak }
  }

  it('R1: a gradient stroke charges its fill and every kept stop before it is built; one byte short refuses the chart', () => {
    const two = parseWith(strokeChart(2)), three = parseWith(strokeChart(3)), many = parseWith(strokeChart(32)), huge = parseWith(strokeChart(1000))
    expect(three.used - two.used).toBeGreaterThan(0)
    expect(many.used - two.used).toBe(30 * (three.used - two.used))
    expect(huge.used).toBe(many.used)
    const line = huge.model.groups[0].series[0].lineStyle!
    expect(line.fill?.kind === 'gradient' && line.fill.stops.length).toBe(32)
    // The refusal boundary is the parse's admitted peak: exactly that admits, one byte less refuses.
    expect(many.peak - two.peak).toBe(30 * (three.used - two.used))
    expect(parseWith(strokeChart(32), many.peak).model.refused).toBeUndefined()
    expect(parseWith(strokeChart(32), many.peak - 1).model.refused).toBe(true)
  })

  it('R2: solid fill and line alpha reach the canvas (zero and partial), unauthored colours stay opaque', () => {
    const alpha = (a: number) => `<a:srgbClr val="FF0000"><a:alpha val="${a}"/></a:srgbClr>`
    expect(pixels(parse(space(bars(ser(0, [3, 5])), { root: `<c:spPr><a:solidFill>${alpha(0)}</a:solidFill></c:spPr>` }))).at(3, 3)[3]).toBe(0)
    const half = pixels(parse(space(bars(ser(0, [3, 5])), { root: `<c:spPr><a:solidFill>${alpha(50000)}</a:solidFill></c:spPr>` }))).at(3, 3)
    expect(half[0]).toBe(255); expect(half[3]).toBeGreaterThan(120); expect(half[3]).toBeLessThan(135)
    // Bars and series strokes: the fill and stroke styles in force carry the alpha.
    const record = (serExtra: string, family: 'bar' | 'line') => {
      const model = parse(space(family === 'bar' ? bars(ser(0, [3, 5], serExtra)) : `<c:lineChart><c:grouping val="standard"/>${ser(0, [1, 3, 2], serExtra + '<c:marker><c:symbol val="none"/></c:marker>')}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`, { legend: false }))
      const canvas = createCanvas(400, 260), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
      const fills: string[] = [], strokes: string[] = []
      const fr = ctx.fillRect.bind(ctx), st = ctx.stroke.bind(ctx)
      ctx.fillRect = (x, y, w, h) => { if (!(x === 0 && y === 0)) fills.push(String(ctx.fillStyle)); fr(x, y, w, h) }
      ctx.stroke = (...a: []) => { strokes.push(String(ctx.strokeStyle)); st(...a) }
      paintXlsxChart(model, ctx, 400, 260)
      return { fills, strokes }
    }
    expect(record(`<c:spPr><a:solidFill>${alpha(0)}</a:solidFill></c:spPr>`, 'bar').fills).toEqual(['rgba(255, 0, 0, 0.00)', 'rgba(255, 0, 0, 0.00)'])
    expect(record(`<c:spPr><a:solidFill>${alpha(40000)}</a:solidFill></c:spPr>`, 'bar').fills).toEqual(['rgba(255, 0, 0, 0.40)', 'rgba(255, 0, 0, 0.40)'])
    expect(record(`<c:spPr><a:ln w="25400"><a:solidFill>${alpha(0)}</a:solidFill></a:ln></c:spPr>`, 'line').strokes.at(-1)).toBe('rgba(255, 0, 0, 0.00)')
    expect(record(`<c:spPr><a:ln w="25400"><a:solidFill>${alpha(40000)}</a:solidFill></a:ln></c:spPr>`, 'line').strokes.at(-1)).toBe('rgba(255, 0, 0, 0.40)')
    expect(record(solid('00FF00'), 'bar').fills).toEqual(['#00ff00', '#00ff00'])
  })

  it('R3: every marker symbol honours fill and stroke suppression; line symbols draw border-only', () => {
    const symbols: MarkerSymbol[] = ['circle', 'square', 'diamond', 'triangle', 'x', 'star', 'dash', 'dot', 'plus']
    for (const symbol of symbols) {
      const ctx = createCanvas(40, 40).getContext('2d') as unknown as CanvasRenderingContext2D
      let fills = 0, strokes = 0
      const f = ctx.fill.bind(ctx), st = ctx.stroke.bind(ctx)
      ctx.fill = (...a: []) => { fills++; f(...a) }
      ctx.stroke = (...a: []) => { strokes++; st(...a) }
      paintMarker(ctx, 20, 20, symbol, 10, false, false)
      expect([symbol, fills, strokes]).toEqual([symbol, 0, 0])
      paintMarker(ctx, 20, 20, symbol, 10, false, true)
      expect([symbol, fills, strokes]).toEqual([symbol, 0, 1])
    }
    // Through the painter: fill and line both noFill paint nothing; a border-only dot shows its border colour.
    const chart = (symbol: string, spPr: string) => parse(space(`<c:lineChart><c:grouping val="standard"/>${ser(0, [5, 5, 5], `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="${symbol}"/><c:size val="20"/><c:spPr>${spPr}</c:spPr></c:marker>`)}<c:marker val="1"/>${AX}</c:lineChart>${AXES}`, { legend: false }))
    for (const symbol of ['x', 'plus', 'dash', 'dot']) {
      expect([symbol, pixels(chart(symbol, '<a:noFill/><a:ln><a:noFill/></a:ln>')).count('4472C4', 40)]).toEqual([symbol, 0])
      expect([symbol, pixels(chart(symbol, '<a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>')).count('00FF00', 40) > 0]).toEqual([symbol, true])
    }
  })

  describe('R4: an authored colour or style part that cannot be admitted refuses the chart', () => {
    const files = async () => {
      const bytes = await chartWorkbookBytes(space(bars(four())), { patch: (zip: JSZip) => {
        zip.file('xl/charts/_rels/chart1.xml.rels', RELS)
        zip.file('xl/charts/style1.xml', style(AUTO_PH))
        zip.file('xl/charts/colors1.xml', COLORS)
      } })
      return OfficePackage.load(bytes)
    }
    const run = async (refuse: (tag: string) => boolean) => {
      const pkg = await files()
      let held = 0
      const workspace = createChartWorkspace('desktop', { reserve: (b: number, tag: string) => { if (refuse(tag)) return false; held += b; return true }, release: (_t: string, b: number) => { held -= b } })
      const root = parseXmlOrdered((await pkg.text('xl/charts/chart1.xml'))!)
      const result = await parseChartPart(root, undefined, { pkg, partPath: 'xl/charts/chart1.xml', profile: 'desktop', workspace })
      return { result, held, used: workspace.used }
    }
    it('admitted: native, the part colours', async () => {
      const { result } = await run(() => false)
      expect(result.kind === 'native' && result.model.groups[0].series.map(s => s.color)).toEqual(['FF0000', '00FF00', '800000', '008000'])
    })
    it('read-time refusal of the colour part', async () => {
      const { result, used } = await run(tag => tag.startsWith('chart-read:xl/charts/colors1.xml'))
      expect(result.kind).toBe('refused')
      expect(used).toBe(0)
    })
    it('extraction-time refusal (the kept palette)', async () => {
      const { result, used } = await run(tag => tag.startsWith('chart-style:'))
      expect(result.kind).toBe('refused')
      expect(used).toBe(0)
    })
  })
})

describe('A4a fix round 3: linear colour methods (MS-ODRAWXML 2.8.4.2)', () => {
  const part = (method: string) => ({ color: { method, colors: [{ kind: 'srgb' as const, value: 'FF0000', transforms: [] }, { kind: 'srgb' as const, value: '00FF00', transforms: [] }], variations: [[], [{ name: 'lumMod', value: '50000' }]] } })
  const rgb = (hex: string) => [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16))
  const lum = (hex: string) => { const [r, g, b] = rgb(hex); return Math.max(r, g, b) + Math.min(r, g, b) }
  const colors = (method: string, count: number) => Array.from({ length: count }, (_, index) => resolveSeriesColor({ index, count, parts: part(method) }))
  it('withinLinear: every object takes the first colour, brightness rising with index/count', () => {
    const within = colors('withinLinear', 3)
    // Index 1 of 3 is red (a shade or tint of FF0000), never green.
    for (const hex of within) { const [r, g, b] = rgb(hex); expect(r).toBeGreaterThan(g); expect(g).toBe(b) }
    expect(lum(within[0])).toBeLessThan(lum(within[1]))
    expect(lum(within[1])).toBeLessThan(lum(within[2]))
  })
  it('acrossLinear: the colour is chosen modulo the list, brightness rising with index/count', () => {
    const across = colors('acrossLinear', 4)
    const hue = (hex: string) => { const [r, g] = rgb(hex); return r > g ? 'red' : 'green' }
    expect(across.map(hue)).toEqual(['red', 'green', 'red', 'green'])
    expect(lum(across[0])).toBeLessThan(lum(across[2]))
    expect(lum(across[1])).toBeLessThan(lum(across[3]))
  })
  it('reversed variants reverse the progression', () => {
    expect(colors('withinLinearReversed', 3)).toEqual(colors('withinLinear', 3).reverse())
    const across = colors('acrossLinear', 4), reversed = colors('acrossLinearReversed', 4)
    expect(reversed.map(h => rgb(h)[0] > rgb(h)[1])).toEqual(across.map(h => rgb(h)[0] > rgb(h)[1]))
    expect(lum(reversed[0])).toBeGreaterThan(lum(reversed[2]))
    expect(lum(reversed[1])).toBeGreaterThan(lum(reversed[3]))
  })
  it('an unknown method cycles with the part variations', () => {
    expect(colors('somethingNew', 3)).toEqual(['FF0000', '00FF00', srgb('FF0000', ['lumMod', 50000])])
    expect(colors('cycle', 3)).toEqual(colors('somethingNew', 3))
  })
})

describe('A4a fix round 4: linear methods select from the variation-expanded total set (MS-ODRAWXML 2.8.3.2)', () => {
  const c = (value: string) => ({ kind: 'srgb' as const, value, transforms: [] as Array<{ name: string; value?: string }> })
  const v = (...pairs: Array<[string, number]>) => pairs.map(([name, n]) => ({ name, value: String(n) }))
  // Compact: two colours, a non-identity first variation and a second one (four logical colours).
  const variations = [v(['lumMod', 20000]), v(['lumMod', 80000])]
  const compact = (method: string) => ({ color: { method, colors: [c('FF0000'), c('00FF00')], variations } })
  // Expanded by hand, in the total-set order (colour fastest, then variation), with no variations.
  const expanded = (method: string) => ({ color: { method, colors: [0, 1].flatMap(k => ['FF0000', '00FF00'].map(hex => ({ ...c(hex), transforms: variations[k] }))), variations: [] } })
  const run = (parts: ReturnType<typeof compact>, count: number) => Array.from({ length: count }, (_, index) => resolveSeriesColor({ index, count, parts }))
  for (const method of ['withinLinear', 'acrossLinear', 'withinLinearReversed', 'acrossLinearReversed', 'cycle']) {
    it(`${method}: the compact part equals its expanded total set (multiple cycles)`, () => {
      expect(run(compact(method), 9)).toEqual(run(expanded(method), 9))
    })
  }
  it('withinLinear uses logical slot 0 with its first variation (a dark red), never the raw first colour alone', () => {
    const single = resolveSeriesColor({ index: 0, count: 1, parts: compact('withinLinear') })
    expect(single).toBe(srgb('FF0000', ['lumMod', 20000]))
  })
})
