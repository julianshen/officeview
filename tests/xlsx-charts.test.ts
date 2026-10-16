import { describe, expect, test } from 'vitest'
import { parseXlsx } from '../src/xlsx/parse'
import { CACHED_COLUMN, chartWorkbook, emu, openpyxlFrame } from './helpers/xlsx-chart-fixture'

describe('worksheet chart frames', () => {
  test('an anchored chart frame with an empty xfrm takes its anchor box (openpyxl)', async () => {
    const sheet = (await parseXlsx(await chartWorkbook(CACHED_COLUMN))).sheets[0]
    const frame = sheet.drawings?.[0]
    expect(frame?.content?.kind).toBe('xlsxChart')
    expect(frame?.transformValid).toBe(true)
    expect(frame?.widthEmu).toBe(emu(320))
    expect(sheet.drawingCoverage?.find(entry => entry.element === 'graphicFrame')?.reason).not.toBe('invalid-transform')
  })

  test('an anchored chart frame whose xfrm is present but malformed stays invalid', async () => {
    const frame = openpyxlFrame().replace('<xdr:xfrm/>', '<xdr:xfrm><a:off x="NaN" y="0"/><a:ext cx="10" cy="10"/></xdr:xfrm>')
    const sheet = (await parseXlsx(await chartWorkbook(CACHED_COLUMN, { frame }))).sheets[0]
    expect(sheet.drawings?.[0]?.transformValid).toBe(false)
  })
})

/** A 2D context that records calls (and the fill style at each fill). */
function recorder() {
  const calls: Array<{ op: string; args: unknown[]; fillStyle?: unknown }> = []
  const state: Record<string, unknown> = { fillStyle: '#000000', strokeStyle: '#000000', font: '10px sans-serif', lineWidth: 1, globalAlpha: 1, textAlign: 'start', textBaseline: 'alphabetic', lineJoin: 'miter' }
  const ctx = new Proxy(state, {
    get(target, prop: string) {
      if (prop in target) return target[prop]
      if (prop === 'measureText') return (text: string) => ({ width: String(text).length * 6 })
      return (...args: unknown[]) => { calls.push({ op: prop, args, fillStyle: target.fillStyle }) }
    },
    set(target, prop: string, value) { target[prop] = value; return true },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, calls }
}
const ser = (i: number, values: number[], extra = '') => `<ser><idx val="${i}"/><order val="${i}"/><tx><v>S${i + 1}</v></tx>${extra}<cat><strLit><ptCount val="${values.length}"/>${values.map((_, k) => `<pt idx="${k}"><v>C${k + 1}</v></pt>`).join('')}</strLit></cat><val><numLit><ptCount val="${values.length}"/>${values.map((v, k) => `<pt idx="${k}"><v>${v}</v></pt>`).join('')}</numLit></val></ser>`
const space = (plot: string, extra = '') => `<chartSpace><chart><title><tx><rich><p><r><t>T</t></r></p></rich></tx></title><plotArea>${plot}<catAx/><valAx><majorGridlines/></valAx></plotArea><legend><legendPos val="b"/></legend>${extra}</chart></chartSpace>`

describe('worksheet chart data', () => {
  test('series without caches (openpyxl) read their values and categories from the cells', async () => {
    const xml = `<chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><chart><title><tx><rich><a:p><a:r><a:t>Bar chart</a:t></a:r></a:p></rich></tx></title><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><order val="0"/><cat><numRef><f>'Styled'!$K$20:$K$24</f></numRef></cat><val><numRef><f>'Styled'!$L$20:$L$24</f></numRef></val></ser></barChart><catAx/><valAx/></plotArea><legend><legendPos val="r"/></legend></chart></chartSpace>`
    const cells = [0, 1, 2, 3, 4].flatMap(i => [{ ref: `K${20 + i}`, v: `Q${i + 1}` }, { ref: `L${20 + i}`, v: (i + 1) * 10 }])
    const sheet = (await parseXlsx(await chartWorkbook(xml, { cells }))).sheets[0]
    const content = sheet.drawings?.[0].content
    expect(content?.kind).toBe('xlsxChart')
    if (content?.kind !== 'xlsxChart') return
    expect(content.chart.title).toBe('Bar chart')
    expect(content.chart.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4', 'Q5'])
    expect(content.chart.groups[0].series[0].values).toEqual([10, 20, 30, 40, 50])
    expect(content.chart.groups[0].series[0].pending).toBeUndefined()
    expect(content.chart.categoryPending).toBeUndefined()
    expect(sheet.drawingCoverage?.find(entry => entry.element === 'graphicFrame')).toMatchObject({ feature: 'chart', selectedRepresentation: 'chart', status: 'native' })
  })

  test('an unsupported chart with no fallback keeps a titled placeholder card', async () => {
    const xml = '<chartSpace><chart><title><tx><rich><p><r><t>Radar</t></r></p></rich></tx></title><plotArea><radarChart><ser><val><numLit><pt idx="0"><v>2</v></pt></numLit></val></ser></radarChart></plotArea></chart></chartSpace>'
    const sheet = (await parseXlsx(await chartWorkbook(xml))).sheets[0]
    const content = sheet.drawings?.[0].content
    expect(content?.kind).toBe('xlsxChart')
    if (content?.kind !== 'xlsxChart') return
    expect(content.chart).toMatchObject({ title: 'Radar', unsupported: 'Radar chart is not supported' })
    expect(sheet.drawingCoverage?.find(entry => entry.element === 'graphicFrame')?.selectedRepresentation).toBe('chart-placeholder')
    const { chartDescription } = await import('../src/xlsx/chart')
    expect(chartDescription(content.chart)).toBe('Chart: Radar. Radar chart is not supported; it is not drawn.')
  })
})

describe('chart parser', async () => {
  const { parseXlsxChart, parseChartRange, chartDescription, fillChartRanges } = await import('../src/xlsx/chart')
  const { parseXmlOrdered } = await import('../src/core/xml')
  const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!

  test('single-area references', () => {
    expect(parseChartRange("'My ''x'' sheet'!$B$2:$B$6")).toEqual({ sheet: "My 'x' sheet", r0: 1, c0: 1, r1: 5, c1: 1 })
    expect(parseChartRange('Sheet1!A1')).toEqual({ sheet: 'Sheet1', r0: 0, c0: 0, r1: 0, c1: 0 })
    expect(parseChartRange('$C$9:$A$1')).toEqual({ r0: 0, c0: 0, r1: 8, c1: 2 })
    expect(parseChartRange('(Sheet1!$A$1,Sheet1!$A$3)')).toBeUndefined()
    expect(parseChartRange('MyName')).toBeUndefined()
  })

  test('stacked horizontal bars, 100% stacked area, line markers, pie colours, doughnut hole and scatter x values', () => {
    const bar = parse(space(`<barChart><barDir val="bar"/><grouping val="stacked"/><overlap val="100"/>${ser(0, [1, 2])}${ser(1, [3, 4])}</barChart>`))
    expect(bar.groups[0]).toMatchObject({ type: 'bar', barDir: 'bar', grouping: 'stacked', overlap: 100 })
    expect(bar.groups[0].series.map(s => s.color)).toEqual(['4472C4', 'ED7D31'])
    expect(bar.categories).toEqual(['C1', 'C2'])
    expect(bar.legend).toBe('b')
    const area = parse(space(`<areaChart><grouping val="percentStacked"/>${ser(0, [1, 2])}</areaChart>`))
    expect(area.groups[0]).toMatchObject({ type: 'area', grouping: 'percentStacked' })
    const line = parse(space(`<lineChart><grouping val="standard"/>${ser(0, [1, 2], '<marker><symbol val="none"/></marker>')}${ser(1, [3, 1], '<spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></spPr>')}<marker val="1"/></lineChart>`))
    expect(line.groups[0].series.map(s => [s.markers, s.color])).toEqual([[false, '4472C4'], [true, 'FF0000']])
    const pie = parse(space(`<pieChart><varyColors val="1"/>${ser(0, [1, 2, 3])}<firstSliceAng val="90"/></pieChart>`))
    expect(pie.groups[0]).toMatchObject({ type: 'pie', firstSliceAngle: 90 })
    expect(pie.groups[0].series[0].pointColors).toEqual(['4472C4', 'ED7D31', 'A5A5A5'])
    const doughnut = parse(space(`<doughnutChart>${ser(0, [1, 2])}<holeSize val="60"/></doughnutChart>`))
    expect(doughnut.groups[0]).toMatchObject({ type: 'doughnut', holeSize: 60 })
    const scatter = parse(space('<scatterChart><scatterStyle val="lineMarker"/><ser><idx val="0"/><xVal><numLit><pt idx="0"><v>1</v></pt><pt idx="1"><v>4</v></pt></numLit></xVal><yVal><numLit><pt idx="0"><v>2</v></pt><pt idx="1"><v>3</v></pt></numLit></yVal></ser></scatterChart>'))
    expect(scatter.groups[0].series[0]).toMatchObject({ xValues: [1, 4], values: [2, 3], line: true, markers: true })
  })

  test('bar and line together share the axes; a pie combination is a placeholder', () => {
    const combo = parse(space(`<barChart><barDir val="col"/>${ser(0, [1, 2])}</barChart><lineChart>${ser(1, [2, 1])}</lineChart>`))
    expect(combo.groups.map(g => g.type)).toEqual(['bar', 'line'])
    expect(combo.unsupported).toBeUndefined()
    expect(parse(space(`<barChart>${ser(0, [1])}</barChart><pieChart>${ser(1, [1])}</pieChart>`)).unsupported).toBe('Combination pie charts are not supported')
  })

  test('the title falls back to the single series name, and an auto-deleted title is omitted', () => {
    const untitled = (deleted: string) => `<chartSpace><chart>${deleted}<title/><plotArea><barChart>${ser(0, [1])}</barChart></plotArea></chart></chartSpace>`
    expect(parse(untitled('')).title).toBe('S1')
    expect(parse(untitled('<autoTitleDeleted val="1"/>')).title).toBeUndefined()
  })

  test('uncached ranges fill from a lookup, bounded per series', () => {
    const model = parse('<chartSpace><chart><plotArea><lineChart><ser><tx><strRef><f>B1</f></strRef></tx><cat><strRef><f>A2:A4</f></strRef></cat><val><numRef><f>Data!B2:B4</f></numRef></val></ser></lineChart></plotArea></chart></chartSpace>')
    const cells: Record<string, unknown> = { 'Owner|0|1': 'Revenue', 'Owner|1|0': 'Jan', 'Owner|2|0': 'Feb', 'Owner|3|0': 'Mar', 'Data|1|1': 5, 'Data|2|1': 'x', 'Data|3|1': 7 }
    const filled = fillChartRanges(model, 'Owner', (sheet, r, c) => cells[`${sheet}|${r}|${c}`] as string | number | undefined)
    expect(filled.categories).toEqual(['Jan', 'Feb', 'Mar'])
    expect(filled.groups[0].series[0]).toMatchObject({ name: 'Revenue', values: [5, undefined, 7] })
    expect(chartDescription(filled)).toBe('Chart. Line chart, 1 series; 3 categories (Jan, Feb, Mar); Revenue: 5, 7.')
    const long = parse('<chartSpace><chart><plotArea><barChart><ser><val><numRef><f>A1:A10000</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>')
    const bounded = fillChartRanges(long, 'S', () => 1)
    expect(bounded.groups[0].series[0].values).toHaveLength(4096)
    expect(bounded.truncated).toBe(true)
  })
})

describe('chart painting', async () => {
  const { parseXlsxChart, paintXlsxChart } = await import('../src/xlsx/chart')
  const { paintDrawingContent } = await import('../src/drawing/content-paint')
  const { parseXmlOrdered } = await import('../src/core/xml')
  const parse = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), undefined)!

  test('a clustered column chart draws one bar per point in the series colour, rising with the value, plus title and legend', () => {
    const { ctx, calls } = recorder()
    paintXlsxChart(parse(space(`<barChart><barDir val="col"/><grouping val="clustered"/>${ser(0, [10, 20, 30, 40, 50])}</barChart>`)), ctx, 320, 200)
    const bars = calls.filter(call => call.op === 'fillRect' && call.fillStyle === '#4472C4' && (call.args[3] as number) > 10)
    expect(bars).toHaveLength(5)
    const heights = bars.map(call => call.args[3] as number)
    expect(heights).toEqual([...heights].sort((a, b) => a - b))
    expect(heights[4] / heights[0]).toBeCloseTo(5, 0)
    const texts = calls.filter(call => call.op === 'fillText').map(call => call.args[0])
    expect(texts).toEqual(expect.arrayContaining(['T', 'S1', 'C1', 'C5', '0', '50']))
  })

  test('horizontal stacked bars stack along x', () => {
    const { ctx, calls } = recorder()
    paintXlsxChart(parse(space(`<barChart><barDir val="bar"/><grouping val="stacked"/><overlap val="100"/>${ser(0, [1, 1])}${ser(1, [1, 1])}</barChart>`)), ctx, 320, 200)
    const first = calls.filter(call => call.op === 'fillRect' && call.fillStyle === '#4472C4' && (call.args[3] as number) > 10)
    const second = calls.filter(call => call.op === 'fillRect' && call.fillStyle === '#ED7D31' && (call.args[3] as number) > 10)
    expect(first).toHaveLength(2)
    expect(second).toHaveLength(2)
    // The second series starts where the first ends, on the same row.
    expect(second[0].args[0]).toBeCloseTo((first[0].args[0] as number) + (first[0].args[2] as number), 5)
    expect(second[0].args[1]).toBeCloseTo(first[0].args[1] as number, 5)
  })

  test('pie slices, lines and scatter markers are drawn', () => {
    const pie = recorder()
    paintXlsxChart(parse(space(`<pieChart><varyColors val="1"/>${ser(0, [1, 1, 2])}</pieChart>`)), pie.ctx, 300, 200)
    expect(pie.calls.filter(call => call.op === 'arc')).toHaveLength(3)
    expect(pie.calls.filter(call => call.op === 'fill').map(call => call.fillStyle)).toEqual(['#4472C4', '#ED7D31', '#A5A5A5'])
    const line = recorder()
    paintXlsxChart(parse(space(`<lineChart>${ser(0, [1, 3, 2])}</lineChart>`)), line.ctx, 300, 200)
    expect(line.calls.filter(call => call.op === 'lineTo').length).toBeGreaterThanOrEqual(2)
    expect(line.calls.some(call => call.op === 'stroke')).toBe(true)
  })

  test('an unsupported chart paints its placeholder card with title and reason through the shared painter', () => {
    const { ctx, calls } = recorder()
    const chart = parse('<chartSpace><chart><title><tx><rich><p><r><t>Radar</t></r></p></rich></tx></title><plotArea><radarChart/></plotArea></chart></chartSpace>')
    paintDrawingContent({ kind: 'xlsxChart', chart }, ctx, 200, 120)
    expect(calls.filter(call => call.op === 'fillText').map(call => call.args[0])).toEqual(['Radar', 'Radar chart is not supported'])
  })
})
