// @vitest-environment node
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
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { drawingFootprint } from '../src/xlsx/view/drawingPayload'
import type { ChartModel } from '../src/charts/model'
import type { ChartParentLedger } from '../src/charts/workspace'
import * as pages from '../src/xlsx/storage/sourcePages'
import { chartWorkbookBytes } from './helpers/xlsx-chart-fixture'
import { pts } from './helpers/chart-points'

// Every chart workspace the engine creates is recorded with its parent's charges (D0 ownership).
const workspaces = vi.hoisted(() => [] as Array<{ workspace: { used: number; reserve(bytes: number, tag: string): boolean }; parent: Array<{ op: 'reserve' | 'release'; tag: string; bytes: number; ok: boolean }> | undefined }>)
const onReserve = vi.hoisted(() => ({ hook: undefined as undefined | (() => void) }))
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
      const reserve = workspace.reserve.bind(workspace)
      workspace.reserve = (bytes, tag) => { onReserve.hook?.(); return reserve(bytes, tag) }
      workspaces.push({ workspace, parent: log })
      return workspace
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
afterEach(async () => { for (const source of opened.splice(0)) await source.release(); workspaces.length = 0; onReserve.hook = undefined })
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
    expect(fromCells.groups[0].series[0].values).toEqual(pts(1, undefined, '#N/A', '#DIV/0!', 'unavailable', 6, undefined, '#N/A'))
    // Sparse cache indices are kept: missing pt@idx are blank at their index, never compacted or zero.
    expect(fromCells.groups[0].series[1].values).toEqual([{ state: 'value', v: 1 }, { state: 'blank' }, { state: 'error', code: '#N/A' }, { state: 'error', code: '#DIV/0!' }, { state: 'blank' }, { state: 'blank' }, { state: 'blank' }, { state: 'value', v: 3, fmt: '0.0' }])
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
    expect(empty.groups[0].series[0].values).toEqual([])
    expect(empty.groups[0].series[0].pending).toBeUndefined()
    const bytes = await workbook(space(`<c:lineChart><c:ser><c:val>${ref('Styled!$A$1:$A$2')}</c:val></c:ser></c:lineChart>`), row(1, num('A1', 3)) + row(2, cell('A2', '<f>A1*2</f>')))
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const content = doc.sheets[0].drawings![0].content
    expect(content?.kind === 'xlsxChart' && content.chart.groups[0].series[0].values).toEqual(pts(3, 6))
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
    expect(charts.displayModel(cached).groups[0].series[0].values).toHaveLength(4096)
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
      // The fill keeps every point and records the hidden provenance; both paths agree.
      expect(chart.groups[0].series[0].values).toEqual(pts(10, 20, 30, 40))
      expect(chart.groups[0].series[0].hidden).toEqual([1])
      expect(chart.groups[0].series[1].hidden).toEqual([0, 1, 2, 3])
      expect(chart.categoryHidden).toEqual([1])
      expect(await paged(bytes)).toEqual(chart)
      const shown = charts.displayModel(chart)
      if (visibleOnly) {
        // Hidden row 2 leaves the categories and every series; the all-hidden column C series is not drawn.
        expect(shown.categories).toEqual(['c1', 'c3', 'c4'])
        expect(shown.groups[0].series.map(s => s.values)).toEqual([pts(10, 30, 40)])
        expect(blue(chart)).toBe(0)
      } else {
        expect(shown).toBe(chart)
        expect(blue(chart)).toBeGreaterThan(50)
      }
    }
  })
})

describe('D0 ownership: the paged document pool is the chart workspace parent (spec 6.4, 6.5)', () => {
  const xml = space(`<c:barChart><c:ser><c:cat>${strRef('Styled!$B$1:$B$3')}</c:cat><c:val>${ref('Styled!$A$1:$A$3')}</c:val></c:ser></c:barChart>`)
  const sheet = [1, 2, 3].map(r => row(r, num(`A${r}`, r) + str(`B${r}`, `c${r}`))).join('')
  const transient = (tag: string) => tag.endsWith(':crossing') || tag.endsWith(':chunk')
  /** The pool bytes the chart workspace's parent holds now (charged minus refunded, tree charges only). */
  const parentHeld = (log: NonNullable<(typeof workspaces)[number]['parent']>) => log.reduce((sum, e) => sum + (e.ok ? (e.op === 'reserve' ? e.bytes : -e.bytes) : 0), 0)

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
    expect(chart?.kind === 'xlsxChart' && chart.chart.groups[0].series[0].values).toEqual(pts(1, 2, 3))
    // The chart workspace was created with the share as its parent: its tree charges reached the share (once),
    // and were refunded to it when the tree was released, so the parent holds nothing of the chart read.
    const { workspace, parent } = workspaces.at(-1)!
    expect(parent).toBeDefined()
    const trees = parent!.filter(e => e.op === 'reserve' && !transient(e.tag))
    expect(trees.length).toBeGreaterThan(0)
    expect(parent!.some(e => e.op === 'reserve' && transient(e.tag))).toBe(true)
    expect(parentHeld(parent!.filter(e => !transient(e.tag)))).toBe(0)
    expect(workspace.used).toBe(0)
    // The tree was in the share at its peak, and the share went down when it was refunded.
    const peak = Math.max(...sets), treeBytes = trees.reduce((sum, e) => sum + e.bytes, 0)
    expect(sets.some((n, i) => i > 0 && n < sets[i - 1])).toBe(true)
    expect(peak).toBeGreaterThanOrEqual(treeBytes)
    // Two live allocations: the retained payload (its pool entry) and the outgoing clone (the share). The parsed
    // payload left the share when its pool entry took it: the share holds the clone, not the payload twice.
    const retained = pool.stats().byTier.rebuildable
    expect(retained).toBe(drawingFootprint(result))
    const payloadSet = sets.at(-2)!, cloneSet = sets.at(-1)!
    expect(cloneSet - payloadSet).toBe(drawingFootprint(result) - drawingFootprint(result.drawings))
    // (The pool also caches the decoded cell pages the fill read: their own retained tier, not chart bytes.)
    expect(pool.stats().transientBytes).toBe(cloneSet)
    // The reply is posted: the share ends (last release of the read).
    share.release()
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().byTier.rebuildable).toBe(retained)

    // 2. Eviction returns the pool to its baseline.
    pool.trim(0)
    expect(pool.stats().byTier.rebuildable).toBe(0)
    expect(pool.stats().totalBytes).toBe(baseline)

    // 3. A read cancelled during its chart read: the share and the workspace's parent return to baseline.
    const controller = new AbortController(), cancelShare = pool.charge()
    onReserve.hook = () => controller.abort()
    const cancelled = await h.read({ set: n => cancelShare.set(n) }, controller.signal)
    onReserve.hook = undefined
    expect(cancelled.completeness).toBe('incomplete')
    await new Promise(resolve => setTimeout(resolve, 20))
    const after = workspaces.at(-1)!
    expect(after.workspace.used).toBe(0)
    expect(parentHeld(after.parent!.filter(e => !transient(e.tag)))).toBe(0)
    cancelShare.release()
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().byTier.rebuildable).toBe(0)

    // 4. Last release of the document: a retained payload is released with the source.
    const again = pool.charge()
    await h.read({ set: n => again.set(n) })
    again.release()
    expect(pool.stats().byTier.rebuildable).toBeGreaterThan(0)
    await h.source.release()
    opened.length = 0
    // Nothing of the charts stays: no retained payload, no share (the session's own bind-time bytes went with it).
    expect(pool.stats().byTier.rebuildable).toBe(0)
    expect(pool.stats().transientBytes).toBe(0)
    expect(pool.stats().totalBytes).toBeLessThanOrEqual(baseline)
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
    expect(content?.kind === 'xlsxChart' && content.chart.groups[0].series[0].values).toEqual(pts(1, 2, 3))
    // Both parses read their charts in the one document workspace, and it holds nothing afterwards.
    expect(workspaces).toHaveLength(1)
    expect(workspaces[0].workspace.used).toBe(0)
  })

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
    const prefix = vi.spyOn(pages, 'textPrefix').mockImplementation((...args: Parameters<typeof rawPrefix>) => { events.push(`build ${args[2]}`); return rawPrefix(...args) })
    try {
      const charged: number[] = []
      const result = await store.readDrawings({ stamp: store.stamp, signal: new AbortController().signal }, 'D', { set: n => { charged.push(n); events.push(`set ${n}`); return true } })
      expect(result.completeness).toBe('complete')
      const values = result.drawings[0].content!.chart!.groups[0].series[0].values
      expect(values.slice(0, 6)).toEqual(pts('#N/A', 'unavailable', 2, '#N/A', 'unavailable', 5))
      // The cell workspace is admitted before the first page is owned.
      const firstPage = events.indexOf('page')
      expect(firstPage).toBeGreaterThan(0)
      expect(events.slice(0, firstPage).some(e => e.startsWith('set '))).toBe(true)
      // At most one decoded page is owned at any moment.
      expect(maxOwned).toBe(1)
      // Every prefix (an error code: at most 16 units; a label) is built only after an accepted reservation grew the share.
      for (let i = 0; i < events.length; i++) if (events[i].startsWith('build')) {
        const lastSet = events.slice(0, i).reverse().find(e => e.startsWith('set '))
        expect(lastSet).toBeDefined()
      }
      expect(events).toContain('build 4') // '#N/A': the code prefix, never more than CHART_ERROR_CODE_LIMIT units
    } finally { prefix.mockRestore() }
  })
})
