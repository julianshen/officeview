// @vitest-environment node
import { preparedAxesBytes } from '../src/charts/prepare'
/**
 * D0: the resident and paged XLSX data contract (spec 5.8 point states, 6.4
 * per-role precedence, 7 analytic vs display limits, 5.8 plotVisOnly, 6.4/6.5
 * ownership). Fixtures are Office-shaped synthetic workbooks, not Office evidence.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { MAIN_THREAD_BUDGETS, MemoryPool } from '../src/xlsx/view/tileCache'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { createViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook, openRemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { drawingFootprint } from '../src/xlsx/view/drawingPayload'
import type { ChartModel, ChartPoint } from '../src/charts/model'
import { CHART_WORKSPACE_LIMIT, createChartWorkspace, type ChartParentLedger } from '../src/charts/workspace'
import { documentChartContext } from '../src/charts/context'
import * as pages from '../src/xlsx/storage/sourcePages'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { list, pts } from './helpers/chart-points'
import v8 from 'node:v8'
import vm from 'node:vm'
v8.setFlagsFromString('--expose_gc')
const gc = vm.runInNewContext('gc') as () => void

// Every chart workspace the engine creates is recorded with its parent's charges (D0 ownership).
const workspaces = vi.hoisted(() => [] as Array<{ workspace: { used: number; reserve(bytes: number, tag: string): boolean }; parent: Array<{ op: 'reserve' | 'release'; tag: string; bytes: number; ok: boolean }> | undefined }>)
const onReserve = vi.hoisted(() => ({ hook: undefined as undefined | (() => void) }))
// Paged reads: each scope's share charges (the read's construction share), and a pressure pre-reserved in new
// parentless workspaces (document facets) to exercise the aggregate chart sublimit.
const scopes = vi.hoisted(() => [] as Array<{ share: Array<{ op: 'reserve' | 'release'; tag: string; bytes: number; ok: boolean }> }>)
const pressure = vi.hoisted(() => ({ bytes: 0 }))
vi.mock('../src/charts/workspace', async original => {
  const actual = await original<typeof import('../src/charts/workspace')>()
  return {
    ...actual,
    createChartWorkspace: (profile: Parameters<typeof actual.createChartWorkspace>[0], parent?: ChartParentLedger) => {
      const log = parent ? [] as NonNullable<(typeof workspaces)[number]['parent']> : undefined
      const wrapped: ChartParentLedger | undefined = parent && {
        reserve: (bytes, tag) => { const ok = parent.reserve(bytes, tag); log!.push({ op: 'reserve', tag, bytes, ok }); return ok },
        release: (tag, bytes) => { parent.release(tag, bytes); log!.push({ op: 'release', tag, bytes, ok: true }) },
      }
      const workspace = actual.createChartWorkspace(profile, wrapped)
      if (!parent && pressure.bytes) workspace.reserve(pressure.bytes, 'test-pressure')
      const reserve = workspace.reserve.bind(workspace)
      workspace.reserve = (bytes, tag) => { onReserve.hook?.(); return reserve(bytes, tag) }
      workspaces.push({ workspace, parent: log })
      return workspace
    },
    scopeChartWorkspace: (document: Parameters<typeof actual.scopeChartWorkspace>[0], share: ChartParentLedger) => {
      const log: (typeof scopes)[number]['share'] = []
      scopes.push({ share: log })
      return actual.scopeChartWorkspace(document, {
        reserve: (bytes, tag) => { const ok = share.reserve(bytes, tag); log.push({ op: 'reserve', tag, bytes, ok }); return ok },
        release: (tag, bytes) => { share.release(tag, bytes); log.push({ op: 'release', tag, bytes, ok: true }) },
      })
    },
  }
})
// largeSource's evaluator-cap reopen, armed per test.
const cap = vi.hoisted(() => ({ armed: false }))
vi.mock('../src/xlsx/storage/admission', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/admission')>()
  return { ...actual, evaluationHitCap: (doc: Parameters<typeof actual.evaluationHitCap>[0]) => cap.armed ? (cap.armed = false, 'evaluation-cap') : actual.evaluationHitCap(doc) }
})
const charts = await import('../src/charts')

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const space = (plot: string, chart = '') => `<c:chartSpace xmlns:c="${C}"><c:chart>${chart}<c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`
const ref = (f: string) => `<c:numRef><c:f>${f}</c:f></c:numRef>`
const strRef = (f: string) => `<c:strRef><c:f>${f}</c:f></c:strRef>`
const lit = (values: Array<number | string | undefined>, count = values.length) => `<c:numLit><c:ptCount val="${count}"/>${values.map((v, i) => v === undefined ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const TRENDLINE = '<c:trendline><c:trendlineType val="linear"/></c:trendline>'
/** A workbook whose sheet "Styled" has exactly `sheetData` (and `cols`), with the chart. */
function workbook(chartXml: string, sheetData: string, cols = ''): Promise<Uint8Array> {
  return chartWorkbookBytes(chartXml, { patch: async zip => {
    const path = 'xl/worksheets/sheet1.xml', sheet = await zip.file(path)!.async('string')
    zip.file(path, sheet.replace(/<cols>[\s\S]*?<\/cols>/, '').replace(/<sheetData>[\s\S]*<\/sheetData>/, `${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${sheetData}</sheetData>`))
  } })
}
const cell = (r: string, inner: string, attrs = '') => `<c r="${r}"${attrs}>${inner}</c>`
const num = (r: string, v: number) => cell(r, `<v>${v}</v>`)
const str = (r: string, v: string) => cell(r, `<is><t>${v}</t></is>`, ' t="inlineStr"')
const row = (r: number, cells: string, attrs = '') => `<row r="${r}"${attrs}>${cells}</row>`
const column = (count: number, value: (i: number) => number) => Array.from({ length: count }, (_, i) => row(i + 1, num(`A${i + 1}`, value(i)))).join('')

const opened: Array<{ release(): Promise<void> }> = []
afterEach(async () => { for (const source of opened.splice(0)) await source.release(); workspaces.length = 0; scopes.length = 0; pressure.bytes = 0; onReserve.hook = undefined })
async function resident(bytes: Uint8Array): Promise<ChartModel> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const content = doc.sheets[0].drawings![0].content
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  return content.chart
}
async function pagedSource(bytes: Uint8Array) {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'D0', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  opened.push(source)
  const store = source.store('1')
  const read = (memory?: { set(bytes: number): boolean }, signal = new AbortController().signal) => store.readDrawings!({ stamp: store.stamp, signal }, 'D', memory)
  return { source, store, read }
}
async function paged(bytes: Uint8Array): Promise<ChartModel> {
  const result = await (await pagedSource(bytes)).read()
  expect(result.completeness).toBe('complete')
  const content = result.drawings[0].content
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  return content.chart
}
const mean = (values: number[]) => values.reduce((a, b) => a + b, 0) / values.length

describe('D0 data contract: point states (spec 5.8, 6.4)', () => {
  it('resident and paged fills produce identical ChartPoint arrays for the same workbook', async () => {
    // Saved data (a sparse cache) beside references: blanks, #N/A vs #DIV/0!, a formula without a saved value,
    // a cached formula error and text.
    const cache = `<c:numRef><c:f>Styled!$Z$1:$Z$8</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>#N/A</c:v></c:pt><c:pt idx="3"><c:v>#DIV/0!</c:v></c:pt><c:pt idx="7" formatCode="0.0"><c:v>3</c:v></c:pt></c:numCache></c:numRef>`
    const xml = space(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>`
      + `<c:ser><c:idx val="0"/><c:cat>${strRef('Styled!$B$1:$B$8')}</c:cat><c:val>${ref('Styled!$A$1:$A$8')}</c:val></c:ser>`
      + `<c:ser><c:idx val="1"/><c:val>${cache}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>`
      + `<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>`)
    const sheet = [
      row(1, num('A1', 1) + str('B1', 'one')),
      row(2, str('B2', 'two')),
      row(3, cell('A3', '<v>#N/A</v>', ' t="e"') + str('B3', 'three')),
      row(4, cell('A4', '<v>#DIV/0!</v>', ' t="e"') + str('B4', 'four')),
      row(5, cell('A5', '<f>1+1</f>') + str('B5', 'five')),
      row(6, num('A6', 6) + str('B6', 'six')),
      row(7, str('A7', 'x') + str('B7', 'seven')),
      row(8, cell('A8', '<f>NA()</f><v>#N/A</v>', ' t="e"') + str('B8', 'eight')),
    ].join('')
    const bytes = await workbook(xml, sheet)
    const fromCells = await resident(bytes)
    expect(list(fromCells.groups[0].series[0].values)).toEqual(pts(1, undefined, '#N/A', '#DIV/0!', 'unavailable', 6, undefined, '#N/A'))
    // Sparse cache indices are kept: missing pt@idx are blank at their index, never compacted or zero.
    expect(list(fromCells.groups[0].series[1].values)).toEqual([{ state: 'value', v: 1, fmt: 'General' }, { state: 'blank' }, { state: 'error', code: '#N/A' }, { state: 'error', code: '#DIV/0!' }, { state: 'blank' }, { state: 'blank' }, { state: 'blank' }, { state: 'value', v: 3, fmt: '0.0' }])
    // Stored compactly: one state byte and one float per point, no object per point.
    expect(fromCells.groups[0].series[0].values.states).toBeInstanceOf(Uint8Array)
    expect(fromCells.groups[0].series[0].values.values).toBeInstanceOf(Float64Array)
    expect(fromCells.categories).toEqual(['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight'])
    // The group's binding is recorded structurally from its axIds (A1a validates it).
    expect(fromCells.groups[0].binding).toEqual({ kind: 'cartesian', category: '1', value: '2' })
    expect(fromCells.axes.map(axis => [axis.id, axis.kind, axis.crossAx])).toEqual([['1', 'category', '2'], ['2', 'value', '1']])
    // The paged fill yields the same model, point for point.
    expect(await paged(bytes)).toEqual(fromCells)
    // #N/A and #DIV/0! stay distinct through the fill; the description still lists only values.
    expect(charts.chartDescription(fromCells)).toContain('Series 1: 1, 6')
  })

  it('an empty ptCount=0 cache is present and empty; a formula without a saved value under the engine policy is evaluated', async () => {
    const empty = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:val><c:numRef><c:f>A1:A3</c:f><c:numCache><c:ptCount val="0"/></c:numCache></c:numRef></c:val></c:ser></c:lineChart>`)), undefined)!
    expect(empty.groups[0].series[0].values).toHaveLength(0)
    expect(empty.groups[0].series[0].pending).toBeUndefined()
    const bytes = await workbook(space(`<c:lineChart><c:ser><c:val>${ref('Styled!$A$1:$A$2')}</c:val></c:ser></c:lineChart>`), row(1, num('A1', 3)) + row(2, cell('A2', '<f>A1*2</f>')))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const content = doc.sheets[0].drawings![0].content
    expect(list(content?.kind === 'xlsxChart' && content.chart.groups[0].series[0].values)).toEqual(pts(3, 6))
  })
})

describe('D0 analytic acquisition before display caps (spec 7)', () => {
  it('analytic roles keep all points up to 65536: a 5000-point series is not truncated to 4096 for analysis', async () => {
    // 4,097 points: 4,096 zeros, then 4,097. The full mean is 1; a 4,096-point prefix would give 0.
    const data = [...Array.from({ length: 4096 }, () => 0), 4097]
    const cached = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser>${TRENDLINE}<c:val>${lit(data)}</c:val></c:ser></c:lineChart>`)), undefined)!
    const series = cached.groups[0].series[0]
    expect(series.analytic).toBe(true)
    expect(series.values).toHaveLength(4097)
    const input = charts.analyticInput(series)
    expect(input.kind === 'input' && mean(input.values)).toBe(1)
    // Display still draws 4,096 and says so.
    expect(cached.truncated).toBe(true)
    const view = charts.displayView(cached)
    expect(view.slots).toBe(4096)
    expect(view.series[0].value(4095)).toBe(0)
    // Without an analytic element the series is a display series: 4,096 points.
    const plain = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:val>${lit(data)}</c:val></c:ser></c:lineChart>`)), undefined)!
    expect(plain.groups[0].series[0].values).toHaveLength(4096)
    expect(plain.groups[0].series[0].analytic).toBeUndefined()

    // A 5,000-row reference (stdDev error bars are analytic too): resident and paged acquire every point.
    const errBars = '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="stdDev"/><c:val val="1"/></c:errBars>'
    const bytes = await workbook(space(`<c:lineChart><c:ser>${errBars}<c:val>${ref('Styled!$A$1:$A$5000')}</c:val></c:ser></c:lineChart>`), column(5000, i => i === 4999 ? 5000 : 0))
    const filled = await resident(bytes)
    expect(filled.groups[0].series[0].values).toHaveLength(5000)
    const fromCells = charts.analyticInput(filled.groups[0].series[0])
    expect(fromCells.kind === 'input' && mean(fromCells.values)).toBe(1)
    expect(filled.truncated).toBe(true)
    expect(await paged(bytes)).toEqual(filled)
  }, 60_000)

  it('65537 analytic points refuse; display still truncates at 4096 for non-analytic families', () => {
    // Cache: 65,537 points with a trendline refuse the analytic element (never a prefix); the series is drawn capped.
    const big = Array.from({ length: 65537 }, (_, i) => i)
    const refused = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser>${TRENDLINE}<c:val>${lit(big)}</c:val></c:ser></c:lineChart>`)), undefined)!
    const series = refused.groups[0].series[0]
    expect(series.analytic).toBe('refused')
    expect(series.values).toHaveLength(4096)
    expect(charts.analyticInput(series)).toEqual({ kind: 'refused', reason: 'analytic-point-limit' })
    expect(refused.truncated).toBe(true)
    // Exactly 65,536 is acquired in full.
    const edge = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser>${TRENDLINE}<c:val>${lit(big.slice(0, 65536))}</c:val></c:ser></c:lineChart>`)), undefined)!
    expect(edge.groups[0].series[0].analytic).toBe(true)
    expect(edge.groups[0].series[0].values).toHaveLength(65536)
    // The description reads every acquired value without spreading them into one call (JSC caps arguments).
    expect(charts.chartDescription(edge)).toContain('0 to 65,535')
    // An index past the bound (no declared count) refuses too: never a prefix.
    const sparse = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser>${TRENDLINE}<c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="70000"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>`)), undefined)!
    expect(sparse.groups[0].series[0].analytic).toBe('refused')

    // References: the paged reader acquires the same lines the fill reads (65,536 for analytic, 4,096 when refused or display).
    const model = (rows: number, analytic: boolean) => charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser>${analytic ? TRENDLINE : ''}<c:val>${ref(`S!$A$1:$A$${rows}`)}</c:val></c:ser></c:lineChart>`)), undefined)!
    const drawing = (chart: ChartModel) => [{ content: { kind: 'xlsxChart', chart } }]
    expect(charts.pendingLines(drawing(model(65536, true)), 'S').map(line => line.count)).toEqual([65536])
    expect(charts.pendingLines(drawing(model(65537, true)), 'S').map(line => line.count)).toEqual([4096])
    expect(charts.pendingLines(drawing(model(5000, false)), 'S').map(line => line.count)).toEqual([4096])
    const filledRefused = charts.fillChartRanges(model(65537, true), 'S', () => ({ value: 1 }))
    expect(filledRefused.groups[0].series[0]).toMatchObject({ analytic: 'refused' })
    expect(filledRefused.groups[0].series[0].values).toHaveLength(4096)
    expect(filledRefused.truncated).toBe(true)
    const display = charts.fillChartRanges(model(5000, false), 'S', () => ({ value: 1 }))
    expect(display.groups[0].series[0].values).toHaveLength(4096)
    expect(display.truncated).toBe(true)
  }, 60_000)
})

describe('D0 plotVisOnly (spec 5.8)', () => {
  const RED = 'FF0000', BLUE = '0000FF'
  const ln = (color: string) => `<c:spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="38100"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr>`
  const xml = (plotVisOnly: string) => space(`<c:lineChart><c:grouping val="standard"/>`
    + `<c:ser><c:idx val="0"/>${ln(RED)}<c:cat>${strRef('Styled!$B$1:$B$4')}</c:cat><c:val>${ref('Styled!$A$1:$A$4')}</c:val></c:ser>`
    + `<c:ser><c:idx val="1"/>${ln(BLUE)}<c:val>${ref('Styled!$C$1:$C$4')}</c:val></c:ser></c:lineChart>`, plotVisOnly)
  // Row 2 and column C are hidden.
  const sheet = [1, 2, 3, 4].map(r => row(r, num(`A${r}`, r * 10) + str(`B${r}`, `c${r}`) + num(`C${r}`, r * 100), r === 2 ? ' hidden="1"' : '')).join('')
  const cols = '<col min="3" max="3" width="9" hidden="1" customWidth="1"/>'
  const blue = (chart: ChartModel) => {
    const canvas = createCanvas(320, 200), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    charts.paintXlsxChart(chart, ctx, 320, 200)
    const data = (ctx as unknown as { getImageData(x: number, y: number, w: number, h: number): ImageData }).getImageData(0, 0, 320, 200).data
    let count = 0
    for (let i = 0; i < data.length; i += 4) if (data[i] < 60 && data[i + 1] < 60 && data[i + 2] > 200) count++
    return count
  }

  it('plotVisOnly skips hidden rows and columns; plotVisOnly=0 plots them', async () => {
    for (const [option, visibleOnly] of [['', true], ['<c:plotVisOnly val="1"/>', true], ['<c:plotVisOnly val="0"/>', false]] as const) {
      const bytes = await workbook(xml(option), sheet, cols)
      const chart = await resident(bytes)
      // The fill keeps every point and records each role's own hidden provenance; both paths agree.
      expect(list(chart.groups[0].series[0].values)).toEqual(pts(10, 20, 30, 40))
      expect([...chart.groups[0].series[0].values.hidden ?? []]).toEqual([1])
      expect([...chart.groups[0].series[1].values.hidden ?? []]).toEqual([0, 1, 2, 3])
      expect([...chart.categoryHidden ?? []]).toEqual([1])
      expect(await paged(bytes)).toEqual(chart)
      const view = charts.displayView(chart)
      if (visibleOnly) {
        // Hidden row 2 leaves the categories (and the points paired with it); the all-hidden column C series is not drawn.
        expect(Array.from({ length: view.slots }, (_, i) => view.category(i))).toEqual(['c1', 'c3', 'c4'])
        expect(view.series.map(s => Array.from({ length: view.slots }, (_, i) => s.value(i)))).toEqual([[10, 30, 40]])
        expect(blue(chart)).toBe(0)
      } else {
        expect(view.slots).toBe(4)
        expect(view.series.map(s => Array.from({ length: view.slots }, (_, i) => s.value(i)))).toEqual([[10, 20, 30, 40], [100, 200, 300, 400]])
        expect(blue(chart)).toBeGreaterThan(50)
      }
    }
  })

  const drawn = (view: ReturnType<typeof charts.displayView>) => view.series.map(s => Array.from({ length: view.slots }, (_, i) => s.value(i)))
  it('applies plotVisOnly per series and role by its own provenance: independent and offset ranges, scatter x/y pairing (I2)', async () => {
    // Only worksheet row 2 is hidden. Series A reads A1:A3 (its index 1 is hidden); series B reads C10:C12 (nothing hidden).
    const rows = [1, 2, 3, 10, 11, 12].map(r => row(r, (r <= 3 ? num(`A${r}`, r) : '') + (r >= 10 ? num(`C${r}`, (r - 9) * 10) : ''), r === 2 ? ' hidden="1"' : '')).join('')
    const bars = space(`<c:barChart><c:ser><c:idx val="0"/><c:val>${ref('Styled!$A$1:$A$3')}</c:val></c:ser><c:ser><c:idx val="1"/><c:val>${ref('Styled!$C$10:$C$12')}</c:val></c:ser></c:barChart>`)
    const bytes = await workbook(bars, rows)
    const chart = await resident(bytes)
    expect([...chart.groups[0].series[0].values.hidden ?? []]).toEqual([1])
    expect(chart.groups[0].series[1].values.hidden).toBeUndefined()
    expect(await paged(bytes)).toEqual(chart)
    // A's hidden A2 is a gap in A only; B's visible C11 = 20 is drawn (never a union of indices across series).
    expect(drawn(charts.displayView(chart))).toEqual([[1, undefined, 3], [10, 20, 30]])
    // Scatter: x from A1:A3 (row 2 hidden), y from C10:C12: the pair at index 1 is not drawn; the rest are.
    const scatter = space(`<c:scatterChart><c:scatterStyle val="marker"/><c:ser><c:idx val="0"/><c:xVal>${ref('Styled!$A$1:$A$3')}</c:xVal><c:yVal>${ref('Styled!$C$10:$C$12')}</c:yVal></c:ser></c:scatterChart>`)
    const xy = await resident(await workbook(scatter, rows))
    expect([...xy.groups[0].series[0].xValues!.hidden ?? []]).toEqual([1])
    expect(xy.groups[0].series[0].values.hidden).toBeUndefined()
    const view = charts.displayView(xy)
    expect(Array.from({ length: view.slots }, (_, i) => [view.series[0].x(i), view.series[0].value(i)])).toEqual([[1, 10], [undefined, undefined], [3, 30]])
    expect(await paged(await workbook(scatter, rows))).toEqual(xy)
  })

  it('a cached role keeps its reference for provenance: cached values win, hidden rows and columns still apply (I3)', async () => {
    // The caches say [1,2,3] and [7,8,9]; the worksheet says otherwise. Row 2 and column C are hidden.
    const cached = (f: string, values: number[]) => `<c:numRef><c:f>${f}</c:f><c:numCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
    const rows = [1, 2, 3].map(r => row(r, num(`A${r}`, r * 100) + num(`C${r}`, r * 1000), r === 2 ? ' hidden="1"' : '')).join('')
    for (const [option, visibleOnly] of [['', true], ['<c:plotVisOnly val="0"/>', false]] as const) {
      const xml2 = space(`<c:lineChart><c:ser><c:idx val="0"/><c:val>${cached('Styled!$A$1:$A$3', [1, 2, 3])}</c:val></c:ser><c:ser><c:idx val="1"/><c:val>${cached('Styled!$C$1:$C$3', [7, 8, 9])}</c:val></c:ser></c:lineChart>`, option)
      const bytes = await workbook(xml2, rows, '<col min="3" max="3" width="9" hidden="1" customWidth="1"/>')
      const chart = await resident(bytes)
      expect(list(chart.groups[0].series[0].values)).toEqual(pts(1, 2, 3))
      expect(list(chart.groups[0].series[1].values)).toEqual(pts(7, 8, 9))
      expect([...chart.groups[0].series[0].values.hidden ?? []]).toEqual([1])
      expect([...chart.groups[0].series[1].values.hidden ?? []]).toEqual([0, 1, 2])
      expect(await paged(bytes)).toEqual(chart)
      expect(drawn(charts.displayView(chart))).toEqual(visibleOnly ? [[1, undefined, 3]] : [[1, 2, 3], [7, 8, 9]])
    }
  })
})

describe('D0 ownership: one document chart facet; the paged read share is its parent while the read builds (spec 6.4, 6.5)', () => {
  const xml = space(`<c:barChart><c:ser><c:cat>${strRef('Styled!$B$1:$B$3')}</c:cat><c:val>${ref('Styled!$A$1:$A$3')}</c:val></c:ser></c:barChart>`)
  const sheet = [1, 2, 3].map(r => row(r, num(`A${r}`, r) + str(`B${r}`, `c${r}`))).join('')
  const transient = (tag: string) => tag.endsWith(':crossing') || tag.endsWith(':chunk')
  type Log = (typeof scopes)[number]['share']
  /** The share bytes a scope's charges hold now (charged minus refunded; transient decode work is never in the share). */
  const shareHeld = (log: Log) => log.reduce((sum, e) => sum + (e.ok && !transient(e.tag) ? (e.op === 'reserve' ? e.bytes : -e.bytes) : 0), 0)
  const facet = () => workspaces.find(entry => !entry.parent)!.workspace

  it('a resident model keeps its storage charged in the document workspace while it lives, and returns it on dispose', async () => {
    // The review's probe: an uncached analytic reference A1:A4097, parsed with a real document context and a logging parent.
    const parentLog: Array<[string, number]> = []
    let parentHeld = 0
    const context = documentChartContext('desktop', undefined, { reserve: (bytes, tag) => { parentLog.push([tag, bytes]); parentHeld += bytes; return true }, release: (_tag, bytes) => { parentHeld -= bytes } })
    const bytes = await workbook(space(`<c:lineChart><c:ser>${TRENDLINE}<c:val>${ref('Styled!$A$1:$A$4097')}</c:val></c:ser></c:lineChart>`), row(1, num('A1', 5)))
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached', chartContext: context })
    const content = doc.sheets[0].drawings![0].content
    if (content?.kind !== 'xlsxChart') throw new Error('no chart')
    expect(content.chart.groups[0].series[0].values).toHaveLength(4097)
    expect(content.chart.groups[0].series[0].values.values[0]).toBe(5)
    // The live model's storage (9 B per point) is owned by the document facet, and through it by the parent.
    expect(context.workspace.used).toBeGreaterThanOrEqual(4097 * 9)
    expect(parentHeld).toBe(context.workspace.used)
    expect(charts.chartModelOwner(content.chart)?.workspace).toBe(context.workspace)
    charts.releaseChartModel(content.chart)
    expect(context.workspace.used).toBe(0)
    expect(parentHeld).toBe(0)
  })

  it('a fill whose storage the chart workspace refuses is refused before any cell is read or point built', () => {
    const model = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:val>${ref('S!$A$1:$A$100')}</c:val></c:ser></c:lineChart>`)), undefined)!
    // Room for less than 100 points (900 B) in the parent ledger.
    let parentHeld = 0
    const workspace = createChartWorkspace('desktop', { reserve: bytes => { if (parentHeld + bytes > 500) return false; parentHeld += bytes; return true }, release: (_tag, bytes) => { parentHeld -= bytes } })
    const lookup = vi.fn(() => ({ value: 1 }))
    const filled = charts.fillChartRanges(model, 'S', lookup, { workspace })
    expect(lookup).not.toHaveBeenCalled()
    expect(filled).toMatchObject({ refused: true, unsupported: 'This chart needs more memory than is available for charts' })
    expect(workspace.used).toBe(0)
    expect(parentHeld).toBe(0)
    // With room, the same fill is charged before it is built and stays charged with the model.
    const roomy = createChartWorkspace('desktop')
    const ok = charts.fillChartRanges(model, 'S', lookup, { workspace: roomy })
    expect(ok.groups[0].series[0].values).toHaveLength(100)
    expect(roomy.used - preparedAxesBytes(ok.preparedAxes)).toBe(100 * 9)
    charts.releaseChartModel(ok)
    expect(roomy.used).toBe(0)
  })

  it('retained format and error-code tables are admitted before they are retained (R1)', () => {
    // Exercise retained-format admission at A2's maximum supported code length.
    const fmt = '0'.repeat(255)
    const xmlFmt = space(`<c:lineChart><c:ser><c:tx><c:v>S</c:v></c:tx><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0" formatCode="${fmt}"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>`)
    const workspace = createChartWorkspace('desktop')
    const model = charts.parseXlsxChart(parseXmlOrdered(xmlFmt), undefined, { profile: 'desktop', workspace } as never)!
    expect(list(model.groups[0].series[0].values)).toEqual([{ state: 'value', v: 5, fmt }])
    // The retained charge is the role's whole storage, its format table included (plus the model's axis records, A1a).
    expect(workspace.used - preparedAxesBytes(model.preparedAxes)).toBe(charts.storageBytes(model.groups[0].series[0].values) + model.axes.reduce((sum, axis) => sum + charts.axisBytes(axis), 0))
    expect(workspace.used).toBeGreaterThan(510)
    // Room for the typed storage but not the format string: refused before the format is retained, nothing held.
    let held = 0
    const tight = createChartWorkspace('desktop', { reserve: bytes => { if (held + bytes > 500) return false; held += bytes; return true }, release: (_t, bytes) => { held -= bytes } })
    const refused = charts.parseXlsxChart(parseXmlOrdered(xmlFmt), undefined, { profile: 'desktop', workspace: tight } as never)!
    expect(refused.refused).toBe(true)
    expect(tight.used).toBe(0)
    expect(held).toBe(0)
    // An unknown error code from a cell goes to the role's code table: admitted too.
    const pending = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:val>${ref('S!$A$1:$A$2')}</c:val></c:ser></c:lineChart>`)), undefined)!
    const codes = createChartWorkspace('desktop')
    const filled = charts.fillChartRanges(pending, 'S', (_s, r) => ({ value: '#WEIRD!', error: r ? '#OTHER!' : '#WEIRD!' }), { workspace: codes })
    expect(filled.groups[0].series[0].values.codes).toEqual(['#WEIRD!', '#OTHER!'])
    expect(codes.used - preparedAxesBytes(filled.preparedAxes)).toBe(charts.storageBytes(filled.groups[0].series[0].values))
  })

  it('category labels and series names are admitted by an input bound before any text is built, then trimmed (R3)', () => {
    const cached = '<c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit>'
    const withCats = charts.parseXlsxChart(parseXmlOrdered(space(`<c:barChart><c:ser><c:cat>${strRef('S!$B$1:$B$2')}</c:cat><c:val>${cached}</c:val></c:ser></c:barChart>`)), undefined)!
    const withName = charts.parseXlsxChart(parseXmlOrdered(space(`<c:barChart><c:ser><c:tx>${strRef('S!$C$1')}</c:tx><c:val>${cached}</c:val></c:ser></c:barChart>`)), undefined)!
    for (const model of [withCats, withName]) {
      const refusing = createChartWorkspace('desktop', { reserve: () => false, release: () => {} })
      const texts = vi.fn((_s: string, _r: number, _c: number, text: boolean) => ({ value: 'label', ...(text ? { text: 'label' } : {}) }))
      const refused = charts.fillChartRanges(model, 'S', texts, { workspace: refusing })
      expect(refused.refused).toBe(true)
      // No label or name was formatted or materialized.
      expect(texts.mock.calls.filter(call => call[3]).length).toBe(0)
      expect(refusing.used).toBe(0)
      // With room, the charge is trimmed to exactly the retained text.
      const roomy = createChartWorkspace('desktop')
      const ok = charts.fillChartRanges(model, 'S', texts, { workspace: roomy })
      const retained = model === withCats ? ok.categories.reduce((sum, c) => sum + charts.stringBytes(c), 0) : charts.stringBytes(ok.groups[0].series[0].name)
      expect(model === withCats ? ok.categories : [ok.groups[0].series[0].name]).toEqual(model === withCats ? ['label', 'label'] : ['label'])
      expect(roomy.used - preparedAxesBytes(ok.preparedAxes)).toBe(retained)
    }
  })

  it('a label formatter\'s working buffers are admitted to the chart workspace: a tight budget refuses before the request proceeds (R4)', async () => {
    const { cellDisplayText } = await import('../src/xlsx/render')
    const scientific = { numFmtId: 164, formatCode: '0.00E+00' }
    const cached = '<c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit>'
    const models = [
      charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:cat>${strRef('S!$A$1')}</c:cat><c:val>${cached}</c:val></c:ser></c:lineChart>`)), undefined)!,
      charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:tx>${strRef('S!$A$1')}</c:tx><c:val>${cached}</c:val></c:ser></c:lineChart>`)), undefined)!,
    ]
    for (const model of models) {
      // The formatter reports each working buffer through the observer the fill hands the lookup.
      const requests: number[] = []
      const lookup: Parameters<typeof charts.fillChartRanges>[2] = (_s, _r, _c, text, allocation) => ({ value: 12345, ...(text ? { text: cellDisplayText({} as never, { value: 12345, style: scientific } as never, { maxOutput: 256, allocate: bytes => { requests.push(bytes); allocation?.allocate(bytes) }, release: bytes => allocation?.release(bytes) }) } : {}) })
      let held = 0, peak = 0
      const tight = createChartWorkspace('desktop', { reserve: bytes => { if (held + bytes > 600) return false; held += bytes; peak = Math.max(peak, held); return true }, release: (_t, bytes) => { held -= bytes } })
      const refused = charts.fillChartRanges(model, 'S', lookup, { workspace: tight })
      expect(refused.refused).toBe(true)
      // The formatter asked for its 2,048 B workspace and was refused there: it went no further, nothing is held.
      expect(requests.at(-1)).toBe(128 * 8 + 1024)
      expect(peak).toBeLessThanOrEqual(600)
      expect(held).toBe(0)
      expect(tight.used).toBe(0)
      // With room, every working buffer was admitted at its peak and the charge is trimmed to the retained text.
      requests.length = 0
      let roomyPeak = 0
      const roomy = createChartWorkspace('desktop', { reserve: bytes => { held += bytes; roomyPeak = Math.max(roomyPeak, held); return true }, release: (_t, bytes) => { held -= bytes } })
      const ok = charts.fillChartRanges(model, 'S', lookup, { workspace: roomy })
      expect(model === models[0] ? ok.categories : [ok.groups[0].series[0].name]).toEqual(['1.23E+04'])
      expect(roomyPeak).toBeGreaterThanOrEqual(Math.max(...requests))
      expect(roomy.used - preparedAxesBytes(ok.preparedAxes)).toBe(charts.stringBytes('1.23E+04'))
      held = 0
    }
  })

  it('a diagnostic the label formatter publishes keeps its charge with the sheet, beyond the label and a later refusal (R5)', async () => {
    const { cellDisplayText } = await import('../src/xlsx/render')
    const cached = '<c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit>'
    const model = charts.parseXlsxChart(parseXmlOrdered(space(`<c:lineChart><c:ser><c:cat>${strRef('S!$A$1:$A$2')}</c:cat><c:val>${cached}</c:val></c:ser></c:lineChart>`)), undefined)!
    const diagnosticBytes = (sheet: { diagnostics?: Array<{ message: string }> }) => (sheet.diagnostics ?? []).reduce((n, d) => n + d.message.length * 2 + 72, 32)
    // Row 0 uses an unsupported format (a diagnostic is published), row 1 a scientific one.
    const lookupOn = (sheet: object): Parameters<typeof charts.fillChartRanges>[2] => (_s, r, _c, text, allocation) => ({ value: 12345, ...(text ? { text: cellDisplayText(sheet as never, { value: 12345, style: { numFmtId: 164, formatCode: r ? '0.00E+00' : '[FOO]0' } } as never, { maxOutput: 256, allocate: bytes => allocation?.allocate(bytes), release: bytes => allocation?.release(bytes), retain: (targets, bytes, commit, kind) => allocation?.retain?.(targets, bytes, commit, kind) }) } : {}) })
    // 1. A successful fill: the labels are charged with the model, the diagnostic with the sheet.
    const sheet: { diagnostics?: Array<{ message: string }> } = {}
    const workspace = createChartWorkspace('desktop')
    const ok = charts.fillChartRanges(model, 'S', lookupOn(sheet), { workspace })
    expect(ok.categories).toEqual(['12345', '1.23E+04'])
    expect(sheet.diagnostics).toHaveLength(1)
    const labels = ok.categories.reduce((n, c) => n + charts.stringBytes(c), 0)
    expect(workspace.used - preparedAxesBytes(ok.preparedAxes)).toBe(labels + diagnosticBytes(sheet))
    // The model's release returns only the labels; the diagnostic stays charged while the sheet holds it.
    charts.releaseChartModel(ok)
    expect(workspace.used).toBe(diagnosticBytes(sheet))
    // Released once, with the sheet (the document's release).
    charts.releaseChartRetained(sheet)
    expect(workspace.used).toBe(0)
    charts.releaseChartRetained(sheet)
    expect(workspace.used).toBe(0)

    // 2. A later refusal (the second label's formatter workspace) refunds the labels and keeps the diagnostic.
    const other: { diagnostics?: Array<{ message: string }> } = {}
    let held = 0
    const tight = createChartWorkspace('desktop', { reserve: bytes => { if (held + bytes > 2600) return false; held += bytes; return true }, release: (_t, bytes) => { held -= bytes } })
    const refused = charts.fillChartRanges(model, 'S', lookupOn(other), { workspace: tight })
    expect(refused.refused).toBe(true)
    expect(other.diagnostics).toHaveLength(1)
    expect(tight.used).toBe(diagnosticBytes(other))
    expect(held).toBe(diagnosticBytes(other))
    charts.releaseChartRetained(other)
    expect(tight.used).toBe(0)
    expect(held).toBe(0)
  })

  it('the resident fill wires the formatter\'s observers into the chart workspace and propagates a refusal (R4)', async () => {
    // A category formatted with 0.00E+00, parsed with a 600 B chart parent: the chart is refused, never a silent fallback.
    const bytes = await workbook(space(`<c:lineChart><c:ser><c:cat>${strRef('Styled!$B$1')}</c:cat><c:val>${ref('Styled!$A$1')}</c:val></c:ser></c:lineChart>`), row(1, num('A1', 1) + cell('B1', '<v>12345</v>', ' s="3"')), '')
    const styled = await import('jszip').then(async ({ default: JSZip }) => {
      const zip = await JSZip.loadAsync(bytes)
      const styles = await zip.file('xl/styles.xml')!.async('string')
      zip.file('xl/styles.xml', styles.replace('<cellXfs count="3">', '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00E+00"/></numFmts><cellXfs count="4">').replace('</cellXfs>', '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'))
      return zip.generateAsync({ type: 'uint8array' })
    })
    // The parent leaves the chart read and the numeric storage alone and gives label formatting 600 B.
    let held = 0
    const label = (tag: string) => tag.endsWith(':label')
    const context = documentChartContext('desktop', undefined, { reserve: (n, tag) => { if (!label(tag)) return true; if (held + n > 600) return false; held += n; return true }, release: (tag, n) => { if (label(tag)) held -= n } })
    const doc = await parseXlsx(await OfficePackage.load(styled), { valuePolicy: 'cached', chartContext: context })
    const content = doc.sheets[0].drawings![0].content
    expect(content?.kind === 'xlsxChart' && content.chart.refused).toBe(true)
    expect(held).toBe(0)
    // With room, the same category is formatted.
    const roomy = await parseXlsx(await OfficePackage.load(styled), { valuePolicy: 'cached' })
    const fine = roomy.sheets[0].drawings![0].content
    expect(fine?.kind === 'xlsxChart' && fine.chart.categories).toEqual(['1.23E+04'])
  })

  it('an unretained paged reply keeps its chart-facet charge while it lives and returns it once when released (R2)', async () => {
    const bytes = await workbook(xml, sheet)
    const h = await pagedSource(bytes)
    // No pool bound: the reply is the parsed payload itself (never retained).
    let reply: Awaited<ReturnType<typeof h.read>> | undefined = await h.read({ set: () => true })
    const document = facet()
    expect(reply.drawings[0].content?.kind === 'xlsxChart' && (reply.drawings[0].content as { chart: ChartModel }).chart.unsupported).toBeFalsy()
    expect(document.used).toBeGreaterThan(0)
    const owned = document.used
    // Released when the reply is: collected (here) or transferred away and collected.
    reply = undefined
    for (let i = 0; i < 20 && document.used; i++) { gc(); await new Promise(resolve => setTimeout(resolve, 10)) }
    expect(document.used).toBe(0)
    expect(owned).toBeGreaterThan(0)
  })

  it('the chart sublimit holds across the charts of a paged document: a chart past it is refused, an evicted one returns its room', async () => {
    // Two analytic charts on one sheet, 65,536 points each (576 KiB of storage each, more than a chart read's
    // transient work); the facet has room for one.
    const big = (s: number) => space(`<c:lineChart><c:ser><c:idx val="${s}"/>${TRENDLINE}<c:val>${ref('Styled!$A$1:$A$65536')}</c:val></c:ser></c:lineChart>`)
    const second = `<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3048000" cy="1905000"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 2"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`
    const bytes = await chartWorkbookBytes(big(0), { patch: async zip => {
      const path = 'xl/worksheets/sheet1.xml', s0 = await zip.file(path)!.async('string')
      zip.file(path, s0.replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${column(8, i => i)}</sheetData>`))
      const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string')
      zip.file('xl/drawings/drawing1.xml', drawing.replace('</xdr:wsDr>', `${second}</xdr:wsDr>`))
      zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>')
      zip.file('xl/charts/chart2.xml', big(1))
    } })
    const limit = CHART_WORKSPACE_LIMIT.desktop
    // One analytic chart plus stage-1 graph construction fits; retaining it leaves too little room to construct the second.
    pressure.bytes = limit - Math.round(65536 * 9 * 1.5)
    const h = await pagedSource(bytes)
    const pool = new MemoryPool({ limitBytes: 256 * 1024 * 1024 })
    h.source.bindMemory!(pool)
    const share = pool.charge()
    const result = await h.read({ set: n => share.set(n) })
    share.release()
    const kinds = result.drawings.map(d => d.content?.kind === 'xlsxChart' ? (d.content.chart.refused ? 'refused' : d.content.chart.unsupported ?? 'native') : d.content?.kind)
    expect(kinds).toEqual(['native', 'refused'])
    // The retained native chart holds its storage in the document facet beyond the read.
    const document = facet()
    expect(document.used - pressure.bytes).toBeGreaterThanOrEqual(65536 * 9)
    // A repeated read is a cache hit: its clone is a transport copy in the share, never a second facet charge.
    const held = document.used
    const again = pool.charge()
    await h.read({ set: n => again.set(n) })
    again.release()
    expect(document.used).toBe(held)
    // Eviction returns the room.
    pool.trim(0)
    expect(document.used).toBe(pressure.bytes)
  }, 60_000)

  it('paged charts reserve through the parent document pool: the retained model plus the outgoing clone peak is charged as two live allocations; cancellation, eviction and last release return the parent to its baseline', async () => {
    const bytes = await workbook(xml, sheet)
    const h = await pagedSource(bytes)
    const pool = new MemoryPool({ limitBytes: 64 * 1024 * 1024 })
    h.source.bindMemory!(pool)
    const baseline = pool.stats().totalBytes

    // 1. A read charged to a construction share of the document pool.
    const share = pool.charge(), sets: number[] = []
    const result = await h.read({ set: n => { sets.push(n); return share.set(n) } })
    expect(result.completeness).toBe('complete')
    const chart = result.drawings[0].content
    expect(list(chart?.kind === 'xlsxChart' && chart.chart.groups[0].series[0].values)).toEqual(pts(1, 2, 3))
    // The read worked in a scope of the one document facet, whose parent was the read's share: tree and model
    // charges reached the share (each once) while the read built, and the share was handed off when the payload got
    // its owner (the retained entry) -- the document facet keeps the retained model's charge.
    const document = facet(), log = scopes.at(-1)!.share
    const charged = log.filter(e => e.op === 'reserve' && e.ok && !transient(e.tag))
    expect(charged.some(e => e.tag.startsWith('chart-read:'))).toBe(true)
    expect(charged.some(e => e.tag.startsWith('chart-model#'))).toBe(true)
    // Trees were refunded to the share; the model storage was not refunded (it moved to the pool entry with the payload).
    const modelBytes = shareHeld(log)
    expect(modelBytes).toBeGreaterThan(0)
    // (Label bounds are admitted first and trimmed to the retained text, so the net held is at most what was charged.)
    expect(charged.filter(e => (e.tag.startsWith('chart-model#') || e.tag.startsWith('chart-prepare#'))).reduce((sum, e) => sum + e.bytes, 0)).toBeGreaterThanOrEqual(modelBytes)
    expect(document.used).toBe(modelBytes)
    const peak = Math.max(...sets)
    expect(sets.some((n, i) => i > 0 && n < sets[i - 1])).toBe(true)
    expect(peak).toBeGreaterThanOrEqual(charged.filter(e => e.tag.startsWith('chart-read:')).reduce((sum, e) => sum + e.bytes, 0))
    // Two live allocations: the retained payload (its pool entry, model storage included) and the outgoing clone
    // (the share). The parsed payload, model included, left the share when its pool entry took it.
    const retained = pool.stats().byTier.rebuildable
    expect(retained).toBe(drawingFootprint(result))
    const payloadSet = sets.at(-2)!, cloneSet = sets.at(-1)!
    expect(cloneSet - payloadSet).toBe(drawingFootprint(result) - drawingFootprint(result.drawings))
    // (The pool also caches the decoded cell pages the fill read: their own retained tier, not chart bytes.)
    expect(pool.stats().transientBytes).toBe(cloneSet)
    share.release()
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().byTier.rebuildable).toBe(retained)

    // 2. Eviction returns the pool and the document facet to their baselines.
    pool.trim(0)
    expect(pool.stats().byTier.rebuildable).toBe(0)
    expect(pool.stats().totalBytes).toBe(baseline)
    expect(document.used).toBe(0)

    // 3. A read cancelled during its chart read: the share and the document facet return to baseline.
    const controller = new AbortController(), cancelShare = pool.charge()
    onReserve.hook = () => controller.abort()
    const cancelled = await h.read({ set: n => cancelShare.set(n) }, controller.signal)
    onReserve.hook = undefined
    expect(cancelled.completeness).toBe('incomplete')
    await new Promise(resolve => setTimeout(resolve, 20))
    expect(document.used).toBe(0)
    expect(shareHeld(scopes.at(-1)!.share)).toBe(0)
    cancelShare.release()
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().byTier.rebuildable).toBe(0)

    // 4. Last release of the document: a retained payload and its facet charge go with the source.
    const again = pool.charge()
    await h.read({ set: n => again.set(n) })
    again.release()
    expect(pool.stats().byTier.rebuildable).toBeGreaterThan(0)
    expect(document.used).toBeGreaterThan(0)
    await h.source.release()
    opened.length = 0
    expect(pool.stats().byTier.rebuildable).toBe(0)
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().totalBytes).toBeLessThanOrEqual(baseline)
    expect(document.used).toBe(0)
  })

  it('largeSource reopens with saved values in the same chart workspace: one workspace per document, no orphaned charge', async () => {
    const { parseSpreadsheet } = await import('../src/components/largeSource')
    const bytes = await workbook(xml, sheet)
    const pkg = await OfficePackage.load(bytes)
    const decided = { mode: 'object', decision: { evaluation: { maxArrayCells: 1000, maxArraySide: 1000, maxMaterializedCells: 1000 }, valuePolicy: 'engine', counters: { loadEvaluatedFormulas: 0, alternateContentData: [], highRatioParts: [] } } }
    const ctx = { profile: 'D', options: {}, limits: { object: { maxRatio: 100 } } }
    workspaces.length = 0
    cap.armed = true
    const doc = await parseSpreadsheet(pkg, bytes, ctx as never, decided as never)
    expect(cap.armed).toBe(false)
    expect(doc.diagnostics?.some(d => d.kind === 'cached-values' || /saved/i.test(d.message))).toBe(true)
    const content = doc.sheets[0].drawings![0].content
    expect(list(content?.kind === 'xlsxChart' && content.chart.groups[0].series[0].values)).toEqual(pts(1, 2, 3))
    // Both parses read their charts in the one document workspace. The discarded first parse returned its storage:
    // the workspace holds exactly the live model's, and nothing once that model is released.
    expect(workspaces).toHaveLength(1)
    expect(workspaces[0].workspace.used).toBeGreaterThan(0)
    if (content?.kind === 'xlsxChart') charts.releaseChartModel(content.chart)
    expect(workspaces[0].workspace.used).toBe(0)
  })

  it('the display maximum (64 series x 4,096 points) still crosses from the worker to main under the default budgets', async () => {
    // Point objects make a chart reply larger (about 25.7 MB measured for this chart, 4.2 MB before D0):
    // it must still arrive, through the drawing grant's measured retry, within the main-thread budget.
    const name = (n: number) => { let out = ''; for (n++; n; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out }
    const sers = Array.from({ length: 64 }, (_, s) => `<c:ser><c:idx val="${s}"/><c:val>${ref(`Styled!$${name(s)}$1:$${name(s)}$4096`)}</c:val></c:ser>`).join('')
    const rows = Array.from({ length: 4096 }, (_, r) => row(r + 1, Array.from({ length: 64 }, (_, c) => num(`${name(c)}${r + 1}`, r + c)).join(''))).join('')
    const bytes = await workbook(space(`<c:lineChart>${sers}</c:lineChart>`), rows)
    const blob = new Blob([new Uint8Array(bytes)])
    const session = await openImportSession({ kind: 'blob', blob }, { docId: 'max', profile: 'D', store: new SourcePageStore({ docId: 'max', backend: memoryPageBackend() }) } as never)
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    const kernel = new ViewerKernel({ parse: async () => { throw new Error('the paged source must serve this file') }, openSource: async () => source }, { yield: async () => {} })
    const client = createViewerClient(connectInProcess(kernel)), open = await openRemoteWorkbook(client, 'max', blob, 'D')
    if (!('layouts' in open)) throw new Error('open failed')
    const remote = new RemoteWorkbook(client, open)
    try {
      remote.merges('1')
      for (let i = 0; i < 2000 && remote.loadState('1').completeness !== 'complete'; i++) await new Promise(resolve => setTimeout(resolve, 5))
      let state = remote.drawings('1')
      for (let i = 0; i < 4000 && state.state !== 'ready'; i++) { await new Promise(resolve => setTimeout(resolve, 5)); state = remote.drawings('1') }
      expect(state.state).toBe('ready')
      const content = (state as unknown as { drawings: Array<{ content?: { kind: string; chart?: ChartModel } }> }).drawings[0].content
      expect(content?.chart?.groups[0].series).toHaveLength(64)
      expect(content?.chart?.groups[0].series[63].values).toHaveLength(4096)
      expect(content?.chart?.groups[0].series[63].values.values[4095]).toBe(4095 + 63)
      expect(remote.mainLedger.peakBytes).toBeLessThanOrEqual(MAIN_THREAD_BUDGETS.D)
    } finally { await remote.close() }
  }, 120_000)

  it('existing paged streaming and lease protections still hold', async () => {
    // Errors, unavailable formulas and labels in one chart: the cell workspace is admitted before any page is read,
    // every error code is admitted before it is built, and at most one page is owned at a time.
    const xml2 = space(`<c:barChart><c:ser><c:cat>${strRef('Styled!$B$1:$B$40')}</c:cat><c:val>${ref('Styled!$A$1:$A$40')}</c:val></c:ser></c:barChart>`)
    const rows = Array.from({ length: 40 }, (_, i) => row(i + 1, (i % 3 === 0 ? cell(`A${i + 1}`, '<v>#N/A</v>', ' t="e"') : i % 3 === 1 ? cell(`A${i + 1}`, '<f>1+1</f>') : num(`A${i + 1}`, i)) + str(`B${i + 1}`, `label ${i}`))).join('')
    const bytes = await workbook(xml2, rows)
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'D0s', profile: 'D', pageRows: 4 } as never)
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    opened.push(source)
    const store = source.store('1') as unknown as { readDrawings: (o: object, p: 'D', m?: { set(n: number): boolean }) => Promise<{ completeness: string; drawings: Array<{ content?: { kind: string; chart?: ChartModel } }> }>; stamp: object; valueOf: (...args: unknown[]) => Promise<unknown> }
    const events: string[] = []
    let owned = 0, maxOwned = 0
    const pageStore = (session as unknown as { workbook: { store: { pinnedPage: (...args: unknown[]) => Promise<{ release(): void }> } } }).workbook.store
    const rawPinned = pageStore.pinnedPage.bind(pageStore)
    pageStore.pinnedPage = async (...args: unknown[]) => {
      const held = await rawPinned(...args)
      owned++; maxOwned = Math.max(maxOwned, owned); events.push('page')
      const release = held.release.bind(held)
      held.release = () => { owned--; release() }
      return held
    }
    const rawPrefix = pages.textPrefix
    const prefix = vi.spyOn(pages, 'textPrefix').mockImplementation((...args: Parameters<typeof rawPrefix>) => { const out = rawPrefix(...args); events.push(`build ${args[2]} ${out.text.length}`); return out })
    try {
      const charged: number[] = []
      const result = await store.readDrawings({ stamp: store.stamp, signal: new AbortController().signal }, 'D', { set: n => { charged.push(n); events.push(`set ${n}`); return true } })
      expect(result.completeness).toBe('complete')
      const values = result.drawings[0].content!.chart!.groups[0].series[0].values
      expect((list(values) as ChartPoint[]).slice(0, 6)).toEqual(pts('#N/A', 'unavailable', 2, '#N/A', 'unavailable', 5))
      // The cell workspace is admitted before the first page is owned.
      const firstPage = events.indexOf('page')
      expect(firstPage).toBeGreaterThan(0)
      expect(events.slice(0, firstPage).some(e => e.startsWith('set '))).toBe(true)
      // At most one decoded page is owned at any moment.
      expect(maxOwned).toBe(1)
      // Every prefix (an error code: at most 16 units; a label) is built right after its own reservation: the event
      // immediately before each build is an accepted `set`.
      const buildAt = events.flatMap((e, k) => e.startsWith('build') ? [k] : [])
      expect(buildAt.length).toBeGreaterThan(40)
      for (const k of buildAt) expect(events[k - 1]).toMatch(/^set /)
      expect(events.some(e => /^build 4 4$/.test(e))).toBe(true) // '#N/A': the code prefix, at most CHART_ERROR_CODE_LIMIT units
      // And that reservation is the one admitting it: refusing exactly that `set` means the prefix is never built
      // (checked for the first error-code prefix and the first label prefix).
      const targets = [buildAt.find(k => events[k].startsWith('build 4 ')), buildAt.find(k => events[k].startsWith('build 256 '))]
      for (const target of targets) {
        expect(target).toBeDefined()
        const refuseAt = events.slice(0, target!).filter(e => e.startsWith('set ')).length - 1
        const builtBefore = buildAt.filter(k => k < target!).length
        events.length = 0
        let calls = 0
        const refused = await store.readDrawings({ stamp: store.stamp, signal: new AbortController().signal }, 'D', { set: n => { events.push(`set ${n}`); return calls++ !== refuseAt } })
        expect(refused.completeness).toBe('incomplete')
        expect(events.filter(e => e.startsWith('build')).length).toBe(builtBefore)
      }
    } finally { prefix.mockRestore() }
  })
})
it('A2 resident and paged uncached values retain number formats for sourceLinked labels',async()=>{
 const {addNumberFormat}=await import('./helpers/xlsx-chart-fixture')
 const chart=space('<c:barChart><c:dLbls><c:showVal val="1"/><c:numFmt sourceLinked="1"/></c:dLbls><c:ser><c:val><c:numRef><c:f>Styled!$A$1:$A$2</c:f></c:numRef></c:val></c:ser></c:barChart>')
 const bytes=await chartWorkbookBytes(chart,{cells:[{ref:'A1',v:.25,style:3},{ref:'A2',v:12}],patch:zip=>addNumberFormat(zip,10)})
 const a=await resident(bytes),b=await paged(bytes)
 expect(a.unsupported).toBeUndefined();expect(a.preparedAxes!.dataLabels.map(l=>l.text)).toEqual(['25.00%','12']);expect(b).toEqual(a)
})
