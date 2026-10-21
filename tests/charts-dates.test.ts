// @vitest-environment node
/**
 * A1a: chart date systems, epoch conversion and date axes (spec 5.1, 6.5).
 * Pure model and number tests; fixtures are Office-shaped synthetic, not Office evidence.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { parseXmlOrdered } from '../src/core/xml'
import { chartEpoch, convertSerial, dateTicks, dateUnitIndex, isDateFormat } from '../src/charts/dates'
import { resolveAxes } from '../src/charts/axes'
import { fillChartRanges, paintXlsxChart, parseXlsxChart, type ChartCell } from '../src/charts'
import { createCanvas } from 'canvas'
import { valueAt, pointAt } from '../src/charts/points'
import { encodeCivil } from '../src/xlsx/formula/serial'
import type { ChartModel } from '../src/charts/model'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const space = (plot: string, root = '') => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}">${root}<c:chart><c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`
const num = (values: Array<number | string>, code = '') => `<c:numLit>${code ? `<c:formatCode>${code}</c:formatCode>` : ''}<c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const str = (values: string[]) => `<c:strLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strLit>`
const parse = (xml: string): ChartModel => parseXlsxChart(parseXmlOrdered(xml), undefined)!
const serial = (y: number, m: number, d: number, epoch: '1900' | '1904' = '1900') => encodeCivil({ year: y, month: m, day: d }, epoch) as number
/** A line chart over a date axis (axId 1) and a value axis (axId 2). */
const dateChart = (cat: string, dateAx = '', root = '', val = num([1, 2, 3])) => space(
  `<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${cat}</c:cat><c:val>${val}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
  + `<c:dateAx><c:axId val="1"/><c:axPos val="b"/>${dateAx}<c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/><c:crossBetween val="between"/></c:valAx>`, root)

const opened: Array<{ release(): Promise<void> }> = []
afterEach(async () => { for (const source of opened.splice(0)) await source.release() })

describe('chart dates (spec 5.1)', () => {
  it('chart epoch precedence and val-omitted date1904 is true', () => {
    expect(chartEpoch('1904', '1900')).toBe('1904')
    expect(chartEpoch('1900', '1904')).toBe('1900')
    expect(chartEpoch(undefined, '1904')).toBe('1904')
    expect(chartEpoch(undefined, undefined)).toBe('1900')
    const cat = num([45000, 45001, 45002])
    expect(parse(dateChart(cat, '', '<c:date1904/>')).epoch).toBe('1904')
    expect(parse(dateChart(cat, '', '<c:date1904 val="1"/>')).epoch).toBe('1904')
    expect(parse(dateChart(cat, '', '<c:date1904 val="0"/>')).epoch).toBe('1900')
    // Absent stays undefined on the model; the workbook (or 1900) decides through chartEpoch.
    expect(parse(dateChart(cat)).epoch).toBeUndefined()
  })

  it('same-epoch serial 60 is preserved; cross-epoch conversion uses the civil model', () => {
    // Same epoch: identity, including the fictitious 1900-02-29 and fractions.
    for (const v of [0, 1, 59, 60, 60.5, 61]) expect(convertSerial(v, '1900', '1900')).toBe(v)
    for (const v of [0, 1, 59, 60, 60.5, 61]) expect(convertSerial(v, '1904', '1904')).toBe(v)
    // 1904 -> 1900: every 1904 serial is representable (1904-01-01 is 1900 serial 1462).
    expect(convertSerial(0, '1904', '1900')).toBe(1462)
    expect(convertSerial(1, '1904', '1900')).toBe(1463)
    expect(convertSerial(59, '1904', '1900')).toBe(1521)
    expect(convertSerial(60, '1904', '1900')).toBe(1522)
    expect(convertSerial(60.5, '1904', '1900')).toBe(1522.5)
    expect(convertSerial(61, '1904', '1900')).toBe(1523)
    // 1900 -> 1904: 1, 59, 60, 60.5 and 61 are 1900 dates before 1904-01-01 (60 is the fictitious day): no counterpart.
    for (const v of [0, 0.25, 1, 59, 60, 60.5, 61]) expect(convertSerial(v, '1900', '1904')).toBeUndefined()
    expect(convertSerial(1461, '1900', '1904')).toBeUndefined() // 1903-12-31
    expect(convertSerial(1462, '1900', '1904')).toBe(0)
    expect(convertSerial(1463, '1900', '1904')).toBe(1)
    expect(convertSerial(1521.25, '1900', '1904')).toBe(59.25)
    // Time fractions survive the round trip; it is never a blanket offset applied to the fictitious day.
    expect(convertSerial(convertSerial(45000.75, '1900', '1904')!, '1904', '1900')).toBe(45000.75)
    expect(convertSerial(serial(2024, 2, 29), '1900', '1904')).toBe(serial(2024, 2, 29, '1904'))
  })

  it('same-epoch serial 60 is preserved; cross-epoch conversion uses the civil model: filled roles carry the workbook epoch, cached roles are already in the chart epoch', () => {
    // Chart epoch 1904 (c:date1904), workbook epoch 1900. Categories (date axis) are filled from the workbook;
    // series 0 values are filled on a date-formatted value axis; series 1 values are cached; series 2 is ordinary.
    const xml = space(
      `<c:lineChart><c:grouping val="standard"/>`
      + `<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>S!$A$1:$A$5</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>S!$B$1:$B$5</c:f></c:numRef></c:val></c:ser>`
      + `<c:ser><c:idx val="1"/><c:order val="1"/><c:val>${num([1462, 60, 1, 2, 3])}</c:val></c:ser>`
      + `<c:axId val="1"/><c:axId val="2"/></c:lineChart>`
      + `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="2"/><c:order val="2"/><c:val><c:numRef><c:f>S!$B$1:$B$5</c:f></c:numRef></c:val></c:ser><c:axId val="3"/><c:axId val="4"/></c:barChart>`
      + `<c:dateAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:numFmt formatCode="m/d/yyyy" sourceLinked="0"/><c:crossAx val="1"/></c:valAx>`
      + `<c:catAx><c:axId val="3"/><c:axPos val="b"/><c:delete val="1"/><c:crossAx val="4"/></c:catAx><c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crosses val="max"/><c:crossAx val="3"/></c:valAx>`,
      '<c:date1904/>')
    const model = parse(xml)
    expect(model.unsupported).toBeUndefined()
    const cells: Record<string, number> = { A0: 1463.5, A1: 60, A2: 0, A3: 45000, A4: 1462, B0: 1462, B1: 60.5, B2: 45000.25, B3: 7, B4: 1461 }
    const lookup = (_sheet: string, row: number, col: number): ChartCell | undefined => {
      const v = cells[`${'AB'[col]}${row}`]
      return v === undefined ? undefined : { value: v, number: v, text: String(v) }
    }
    const filled = fillChartRanges(model, 'S', lookup, { epoch: '1900' })
    // Date-axis categories: converted through the civil model; no counterpart -> unavailable.
    const cats = filled.categoryValues!
    expect([0, 1, 2, 3, 4].map(i => pointAt(cats, i))).toEqual([
      { state: 'value', v: 1.5 }, { state: 'unavailable' }, { state: 'unavailable' }, { state: 'value', v: 45000 - 1462 }, { state: 'value', v: 0 },
    ])
    // A date-formatted value axis converts its filled values too (fraction kept).
    const s0 = filled.groups[0].series[0].values
    expect([0, 1, 2, 3, 4].map(i => pointAt(s0, i))).toEqual([
      { state: 'value', v: 0 }, { state: 'unavailable' }, { state: 'value', v: 45000.25 - 1462 }, { state: 'unavailable' }, { state: 'unavailable' },
    ])
    // Cached values are already in the chart epoch: never converted (60 stays 60).
    expect([0, 1, 2].map(i => valueAt(filled.groups[0].series[1].values, i))).toEqual([1462, 60, 1])
    // Ordinary numbers (a value axis without a date format) are never converted.
    expect([0, 1, 2, 3, 4].map(i => valueAt(filled.groups[1].series[0].values, i))).toEqual([1462, 60.5, 45000.25, 7, 1461])
    expect(filled.diagnostics?.filter(d => d.code === 'chart-date-unrepresentable').length).toBeGreaterThan(0)
    // Same epoch: the fictitious day 60 is preserved.
    const same = fillChartRanges(parse(xml.replace('<c:date1904/>', '<c:date1904 val="0"/>')), 'S', lookup, { epoch: '1900' })
    expect(valueAt(same.categoryValues!, 1)).toBe(60)
    expect(valueAt(same.groups[0].series[0].values, 1)).toBe(60.5)
    expect(same.diagnostics?.some(d => d.code === 'chart-date-unrepresentable')).toBeFalsy()
  })

  it('same-epoch serial 60 is preserved; cross-epoch conversion uses the civil model: the resident and paged fills read the workbook epoch', async () => {
    // A 1904 workbook (workbookPr@date1904) holding a chart that says 1900 (c:date1904 val=0): the date-axis
    // categories filled from cells are converted 1904 -> 1900; the ordinary values are not.
    const xml = space(`<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>Styled!$A$1:$A$3</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>Styled!$B$1:$B$3</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
      + `<c:dateAx><c:axId val="1"/><c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`, '<c:date1904 val="0"/>')
    const bytes = await chartWorkbookBytes(xml, {
      cells: [{ ref: 'A1', v: 0 }, { ref: 'B1', v: 5 }, { ref: 'A2', v: 60.5 }, { ref: 'B2', v: 6 }, { ref: 'A3', v: 43000 }, { ref: 'B3', v: 7 }],
      patch: async zip => { const wb = await zip.file('xl/workbook.xml')!.async('string'); zip.file('xl/workbook.xml', wb.replace('<sheets>', '<workbookPr date1904="1"/><sheets>')) },
    })
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const content = doc.sheets[0].drawings![0].content
    if (content?.kind !== 'xlsxChart') throw new Error('no chart')
    const resident = content.chart
    expect([0, 1, 2].map(i => valueAt(resident.categoryValues!, i))).toEqual([1462, 1522.5, 44462])
    expect([0, 1, 2].map(i => valueAt(resident.groups[0].series[0].values, i))).toEqual([5, 6, 7])
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'A1a', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    opened.push(source)
    const store = source.store('1')
    const read = await store.readDrawings!({ stamp: store.stamp, signal: new AbortController().signal }, 'D')
    const pagedContent = read.drawings[0].content
    if (pagedContent?.kind !== 'xlsxChart') throw new Error('no chart')
    expect([0, 1, 2].map(i => valueAt(pagedContent.chart.categoryValues!, i))).toEqual([1462, 1522.5, 44462])
  })

  it('date axis steps by calendar months and years and keeps unequal spacing', () => {
    const jan31 = serial(2024, 1, 31)
    // Months step from the start date (Jan 31 -> Feb 29 -> Mar 31 -> Apr 30), never a constant day count.
    expect(dateTicks(jan31, serial(2024, 5, 1), 1, 'months', '1900').ticks).toEqual([jan31, serial(2024, 2, 29), serial(2024, 3, 31), serial(2024, 4, 30)])
    expect(dateTicks(jan31, serial(2024, 8, 1), 3, 'months', '1900').ticks).toEqual([jan31, serial(2024, 4, 30), serial(2024, 7, 31)])
    expect(dateTicks(serial(2020, 2, 29), serial(2024, 3, 1), 2, 'years', '1900').ticks).toEqual([serial(2020, 2, 29), serial(2022, 2, 28), serial(2024, 2, 29)])
    expect(dateTicks(10, 20, 5, 'days', '1900').ticks).toEqual([10, 15, 20])
    expect(dateUnitIndex(serial(2024, 3, 15), 'months', '1900') - dateUnitIndex(serial(2023, 12, 1), 'months', '1900')).toBe(3)
    expect(dateUnitIndex(serial(2024, 3, 15), 'years', '1900')).toBe(2024)

    // Unequal date spacing is preserved: points at day offsets 0, 1 and 10, not at slots 0, 1, 2.
    const d0 = serial(2024, 1, 1)
    const graph = resolveAxes(parse(dateChart(num([d0, d0 + 1, d0 + 10], 'm/d/yyyy'), '<c:baseTimeUnit val="days"/>')))
    const axis = graph.axes.get('1')!
    expect(axis.role).toBe('date')
    expect(axis.slots).toBe(11)
    expect([0, 1, 2].map(i => graph.categoryIndex(axis, i))).toEqual([0, 1, 10])
    // Monthly base unit: one slot per calendar month; month majors over a month base.
    const monthly = parse(dateChart(num([serial(2024, 1, 15), serial(2024, 2, 15), serial(2024, 6, 15)]), '<c:baseTimeUnit val="months"/><c:majorUnit val="2"/><c:majorTimeUnit val="months"/>'))
    const m = resolveAxes(monthly)
    const mAxis = m.axes.get('1')!
    expect(mAxis.slots).toBe(6)
    expect([0, 1, 2].map(i => m.categoryIndex(mAxis, i))).toEqual([0, 1, 5])
    expect(mAxis.ticks).toEqual([serial(2024, 1, 15), serial(2024, 3, 15), serial(2024, 5, 15)])
    // The auto base unit follows the data: same day of every month -> months; same day and month -> years.
    expect(resolveAxes(parse(dateChart(num([serial(2024, 1, 15), serial(2024, 2, 15), serial(2024, 4, 15)])))).axes.get('1')!.dateBase).toBe('months')
    expect(resolveAxes(parse(dateChart(num([serial(2021, 3, 1), serial(2022, 3, 1), serial(2024, 3, 1)])))).axes.get('1')!.dateBase).toBe('years')
    expect(resolveAxes(parse(dateChart(num([d0, d0 + 1, d0 + 10])))).axes.get('1')!.dateBase).toBe('days')
  })

  it('text categories on a date axis fall back to categorical with a diagnostic', () => {
    const model = parse(dateChart(str(['North', 'South', 'East'])))
    expect(model.unsupported).toBeUndefined()
    const graph = resolveAxes(model)
    const axis = graph.axes.get('1')!
    expect(axis.role).toBe('category')
    expect(axis.slots).toBe(3)
    expect([0, 1, 2].map(i => graph.categoryIndex(axis, i))).toEqual([0, 1, 2])
    expect(graph.diagnostics.some(d => d.code === 'chart-date-axis-categorical')).toBe(true)
    // Never rebuilt from formatted labels: a numeric-looking text category still falls back.
    const mixed = resolveAxes(parse(dateChart(str(['45000', '1/2/2023', '45002']))))
    expect(mixed.axes.get('1')!.role).toBe('category')
  })

  it('classifies date number formats', () => {
    for (const code of ['m/d/yyyy', 'd-mmm-yy', 'mmm-yy', 'yyyy', 'h:mm AM/PM', '[$-409]mmmm d, yyyy', 'dd/mm/yyyy hh:mm']) expect(isDateFormat(code)).toBe(true)
    for (const code of [undefined, 'General', '0.00', '#,##0', '0%', '"days"0', '0.0E+00', '[Red]0.00', '\\d0']) expect(isDateFormat(code)).toBe(false)
  })
})

/** A context that records every call (and still draws). */
function recorder(w = 800, h = 400) {
  const canvas = createCanvas(w, h), target = canvas.getContext('2d')
  const calls: Array<[string, unknown[]]> = []
  const ctx = new Proxy(target, {
    get(t, p) { const v = Reflect.get(t, p); return typeof v === 'function' ? (...args: unknown[]) => { calls.push([String(p), args]); return v.apply(t, args) } : v },
    set(t, p, v) { return Reflect.set(t, p, v) },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, calls, texts: () => calls.filter(([name]) => name === 'fillText').map(([, args]) => String(args[0])) }
}
async function bothPaths(bytes: Uint8Array): Promise<{ resident: ChartModel; paged: ChartModel }> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const content = doc.sheets[0].drawings![0].content
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'A1a-r1', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  opened.push(source)
  const store = source.store('1')
  const read = await store.readDrawings!({ stamp: store.stamp, signal: new AbortController().signal }, 'D')
  const pagedContent = read.drawings[0].content
  if (pagedContent?.kind !== 'xlsxChart') throw new Error('no chart')
  return { resident: content.chart, paged: pagedContent.chart }
}
const date1904Workbook = (patch = true) => async (zip: import('jszip')) => {
  if (!patch) return
  const wb = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', wb.replace('<sheets>', '<workbookPr date1904="1"/><sheets>'))
}

describe('A1a fix round 1: epochs and category source types (spec 5.1)', () => {
  it('I1: a 1904 workbook without a chart date1904 paints its cached date axis in 1904, in every realm', async () => {
    const xml = space(`<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>Styled!$A$1:$A$3</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:cat><c:val>${num([1, 2, 3])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
      + `<c:dateAx><c:axId val="1"/><c:baseTimeUnit val="days"/><c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`)
    const bytes = await chartWorkbookBytes(xml, { cells: [{ ref: 'A1', v: 0 }, { ref: 'A2', v: 1 }, { ref: 'A3', v: 2 }], patch: date1904Workbook() })
    const { resident, paged } = await bothPaths(bytes)
    for (const model of [resident, paged, structuredClone(resident), structuredClone(paged)]) {
      // The authored chart epoch stays absent; the resolved one travels with the model.
      expect(model.epoch).toBeUndefined()
      expect(model.workbookEpoch).toBe('1904')
      expect(resolveAxes(model).epoch).toBe('1904')
      const { ctx, texts } = recorder()
      paintXlsxChart(model, ctx, 800, 400)
      expect(texts()).toEqual(expect.arrayContaining(['1/1/1904', '1/2/1904', '1/3/1904']))
      expect(texts().some(t => t.endsWith('/1900'))).toBe(false)
    }
    // A 1900 workbook carries nothing (absent means the 1900 system).
    const plain = await bothPaths(await chartWorkbookBytes(xml, { cells: [{ ref: 'A1', v: 1 }] }))
    expect([plain.resident.workbookEpoch, plain.paged.workbookEpoch]).toEqual([undefined, undefined])
  })

  it('I2: numeric-looking text categories stay text: cached and filled, resident and paged, fall back to categorical', async () => {
    const chart = (cat: string) => space(`<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${cat}</c:cat><c:val>${num([1, 2, 3])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>`
      + `<c:dateAx><c:axId val="1"/><c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`)
    const roleOf = (model: ChartModel) => { const graph = resolveAxes(model); return { role: graph.axes.get('1')!.role, fallback: graph.diagnostics.some(d => d.code === 'chart-date-axis-categorical'), index: [0, 1, 2].map(i => graph.categoryIndex(graph.axes.get('1')!, i)) } }
    const ref = '<c:numRef><c:f>Styled!$A$1:$A$3</c:f></c:numRef>'
    // Text cells that look like numbers.
    const text = await bothPaths(await chartWorkbookBytes(chart(ref), { cells: [{ ref: 'A1', v: '45000' }, { ref: 'A2', v: '45001' }, { ref: 'A3', v: '45010' }] }))
    const cachedText = parse(chart(str(['45000', '45001', '45010'])))
    for (const model of [text.resident, text.paged, cachedText]) expect(roleOf(model)).toEqual({ role: 'category', fallback: true, index: [0, 1, 2] })
    // Number cells and a numeric cache are dates, with the same positions.
    const numbers = await bothPaths(await chartWorkbookBytes(chart(ref), { cells: [{ ref: 'A1', v: 45000 }, { ref: 'A2', v: 45001 }, { ref: 'A3', v: 45010 }] }))
    const cachedNumbers = parse(chart(num([45000, 45001, 45010])))
    for (const model of [numbers.resident, numbers.paged, cachedNumbers]) expect(roleOf(model)).toEqual({ role: 'date', fallback: false, index: [0, 1, 10] })
    // A paged-style cell that carries both a scanned number and its text value is still text.
    const filled = fillChartRanges(parse(chart(ref)), 'Styled', (_s, row) => ({ value: String(45000 + row), number: 45000 + row, text: String(45000 + row) }))
    expect(roleOf(filled).role).toBe('category')
  })
})
