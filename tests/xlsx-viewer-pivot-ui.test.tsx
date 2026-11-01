/**
 * Plan Task 28 (spec §6.6, §6.7, §7.1, §7.6, §7.9, §8): pivot buttons as controls and menus in the interactive viewer.
 *
 * - The resident client runs Task 25's session on the main thread and hands the viewer a commit: the live definition
 *   and the sheet's composite value overlay (cumulative over the saved area, so a later change never lets the saved
 *   cells show through rows an earlier one vacated).
 * - The viewer installs a commit (live definition, region styles invalidated, overlay set, valueGen bumped, tiles
 *   dropped); the buttons are DOM hit targets over the canvas-painted ones, named by field and action, or by the
 *   refusal reason when the pivot is not interactive (one status notice per pivot, plus the load-time pivot-*
 *   diagnostics the viewer never showed before).
 * - Menu entries exist only for interaction kinds in the session's verified set.
 */
import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { pivotRuntimeFor } from '../src/xlsx/pivotLoader'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { MainMemoryLedger, estimatePayloadBytes } from '../src/xlsx/view/mainLedger'
import type { XlsxDocument } from '../src/xlsx/types'
import type { PivotDeltaKind } from '../src/xlsx/pivots/session/delta'
import { sheetCharts } from '../src/components/xlsx/sheetCharts'
import { chartDescription } from '../src/charts/describe'
import type { ChartModel, ChartPoints, ChartSeries } from '../src/charts/model'
import { StructuralIndex, fromStoreScan } from '../src/xlsx/view/structure'
import { guardFilter, guardSort } from '../src/xlsx/view/operationGuard'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { openPagedWorkbook } from '../src/components/xlsx/pagedWorkbook'
import { pagedWorkbookOf } from '../src/components/pagedDocument'
import { FORCE_PAGED_LIMITS } from '../src/debug/forcePaged'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'

const FIXTURES = `${__dirname}/fixtures/xlsx-pivots/`
const BASELINE = `${FIXTURES}engine/baseline.xlsx`
const PAGE_SINGLE = `${FIXTURES}engine/quirk-page-single.xlsx`
const CHART_LINKED = `${FIXTURES}interaction/chart-baseline.xlsx`
const chartCase = (name: string) => `${FIXTURES}interaction/${name}.xlsx`
const linkedCharts = (doc: XlsxDocument) => doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings)).filter(chart => chart.pivot)
/** A chart line Task 30's gate does not cover: the pivot is display-only, refused `pivot-chart-mapping`. */
async function lineChartDoc(): Promise<XlsxDocument> {
  const doc = await load(CHART_LINKED)
  for (const chart of linkedCharts(doc)) for (const group of chart.groups) (group as { family: string }).family = 'line'
  return doc
}
const n15 = (v: number) => Number(v.toPrecision(15))
const points = (p: ChartPoints) => Array.from({ length: p.length }, (_, i) => p.states[i] === 1 ? n15(p.values[i]!) : p.states[i])
/** What a pivot chart shows (Task 30's comparison): categories, levels, references, each series' name, colour and points. */
const chartData = (chart: ChartModel) => ({
  categories: chart.categories, categoryRef: chart.categoryRef,
  levels: (chart.categoryLevels ?? []).map(level => ({ starts: [...level.starts], labels: level.labels })),
  series: chart.groups.flatMap(group => group.series.map((s: ChartSeries) => ({ index: s.index, name: s.name, color: s.color, ref: s.refs?.values, values: points(s.values) }))),
})
const signal = () => new AbortController().signal

async function load(file: string): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(file))))
}
const pivotSheet = (doc: XlsxDocument) => doc.sheets.findIndex(sheet => sheet.pivotTables?.length)
const options = (store: { stamp: unknown }) => ({ signal: signal(), stamp: { ...(store.stamp as object) } as never })
async function valueAt(store: any, row: number, col: number): Promise<unknown> {
  const read = await store.readCell({ row, col }, options(store))
  return read.kind === 'cell' ? read.cell.value : read.kind
}
/** The hidden text of the gridcell at (row, col), source coordinates (rows are display slots without a sort). */
function cellText(container: HTMLElement, row: number, col: number): string | undefined {
  const line = container.querySelector(`[role=row][aria-rowindex="${row + 2}"]`)
  return line?.querySelector(`[role=gridcell][aria-colindex="${col + 2}"]`)?.textContent ?? undefined
}
const status = (container: HTMLElement) => [...container.querySelectorAll('[role=status]')].map(node => node.textContent ?? '').join(' ')

afterEach(() => { cleanup(); vi.unstubAllGlobals() })

describe('resident pivot client (headless)', () => {
  it('applies a delta as source data: the live definition and a cumulative overlay, reset gives the saved cells back', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(BASELINE), sheet = pivotSheet(doc), source = createObjectWorkbookSource(doc, 'pivot-client')
    const store: any = source.store(doc.sheets[sheet]!.sheetId ?? String(sheet + 1))
    const installs: any[] = []
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => { installs.push(commit); store.setValueOverlay(commit.overlay) } })
    const opened = await client.open(sheet, 0, signal())
    expect(opened.ok, JSON.stringify(opened)).toBe(true)
    if (!opened.ok) return
    expect(opened.revision).toBe(0)
    expect([...opened.verifiedKinds].sort()).toEqual(['addDataField', 'clearFilters', 'collapse', 'collapseAll', 'collapseField', 'filter', 'items', 'moveDataField', 'moveField', 'page', 'removeDataField', 'sort'])
    const region = opened.model.fields.find(field => field.name === 'Region')!
    expect(region.items.map(item => item.label)).toEqual(['East', 'North', 'South', 'West', 'Central'].slice(0, region.items.length))
    // A3:I24 saved; R24 is the grand total row. Collapsing East (rows 6–8 under it) vacates three rows at the bottom.
    expect(await valueAt(store, 23, 0)).toBe('Grand Total')
    const before = store.stamp.valueGen
    const collapsed = await client.apply(sheet, 0, { kind: 'collapse', field: region.index, item: 0, collapsed: true })
    expect(collapsed.ok, JSON.stringify(collapsed)).toBe(true)
    expect(installs).toHaveLength(1)
    expect(store.stamp.valueGen).toBe(before + 1)
    expect(installs[0].def).not.toBe(doc.sheets[sheet]!.pivotTables![0])
    expect(await valueAt(store, 5, 0)).toBe('East')
    expect(await valueAt(store, 6, 0)).toBe('North')
    expect(await valueAt(store, 23, 0)).toBe('blank')
    // Overlay cells paint with the saved style of a cell in the same role (the browser's oracle repaint caught both):
    // North's label moves up to row 6 and keeps a level-0 label's style, the grand total row keeps its own (no indent).
    const savedStyle = (row: number, col: number) => doc.sheets[sheet]!.rows.find(r => r.index === row)!.cells.find(c => c.col === col)!.style
    expect(installs[0].overlay.cell(6, 0).style).toEqual(savedStyle(9, 0))
    expect(installs[0].overlay.cell(20, 0).style).toEqual(savedStyle(23, 0))
    expect(installs[0].overlay.cell(20, 0).style.vertical).toBe('center')
    // A later change with nothing vacated keeps the earlier vacated rows blank (the overlay covers the saved area).
    const sorted = await client.apply(sheet, 0, { kind: 'sort', field: region.index, order: 'descending' })
    expect(sorted.ok, JSON.stringify(sorted)).toBe(true)
    expect(await valueAt(store, 23, 0)).toBe('blank')
    expect(await valueAt(store, 22, 0)).toBe('blank')
    const reset = await client.apply(sheet, 0, 'reset')
    expect(reset.ok).toBe(true)
    expect(installs.at(-1).def).toBe(opened.def)
    expect(await valueAt(store, 23, 0)).toBe('Grand Total')
    expect(await valueAt(store, 6, 0)).toBe('Apples')
    client.release(sheet, 0)
    expect((await client.apply(sheet, 0, 'reset')).ok).toBe(false)
  })

  it('a resident open is admitted against the main-thread budget, as the worker admits it; refused, nothing stays charged', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const { MAIN_QUOTA_MESSAGE } = await import('../src/xlsx/pivots/session/runtime')
    const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
    const doc = await load(BASELINE), sheet = pivotSheet(doc)
    const ledger = new MainMemoryLedger()
    ledger.set('document', 'everything-else', MAIN_THREAD_BUDGETS.D)   // the main budget is used up
    const full = createResidentPivotClient({ doc, ledger, install: () => {} })
    const refused = await full.open(sheet, 0, signal())
    expect(refused).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-records-limit', message: MAIN_QUOTA_MESSAGE } })
    // The runtime's retained copy is charged by the viewer before any open (S3-F); the store and its decode are not.
    const retained = pivotRuntimeFor(doc)!.retainedBytes
    expect(ledger.bytes).toBe(MAIN_THREAD_BUDGETS.D + retained)
    full.dispose?.()
    // With room the same open is admitted and charges its store.
    ledger.release('document', 'everything-else')
    const roomy = createResidentPivotClient({ doc, ledger, install: () => {} })
    expect((await roomy.open(sheet, 0, signal())).ok).toBe(true)
    expect(ledger.bytes).toBeGreaterThan(retained)
    expect(ledger.bytes).toBeLessThanOrEqual(MAIN_THREAD_BUDGETS.D)
    roomy.dispose?.()
  })

  it('the runtime does not charge or keep a store its caller refuses to admit', async () => {
    const doc = await load(BASELINE), sheet = doc.sheets[pivotSheet(doc)]!
    const runtime = pivotRuntimeFor(doc)!, cache = runtime.cache(sheet.pivotTables![0]!)!
    const ledger = new MainMemoryLedger()
    runtime.chargeRetained(ledger)
    const before = ledger.bytes
    const refused = await runtime.lease(cache, ledger, signal(), undefined, () => false)
    expect(refused).toMatchObject({ ok: false, diagnostic: { feature: 'pivot-records-limit' }, quota: { requiredBytes: expect.any(Number) } })
    expect(ledger.bytes).toBe(before)
    // Not kept: the next lease decodes again and, admitted, holds the store until it is released.
    const admitted = await runtime.lease(cache, ledger, signal(), undefined, () => true)
    if (!admitted.ok) throw new Error(admitted.diagnostic.message)
    expect(ledger.bytes).toBe(before + admitted.lease.store.bytes)
    admitted.lease.release()
    expect(ledger.bytes).toBe(before)
  })

  it('S3-G: after a commit the vacated saved area stays protected from sheet sort and filter', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(BASELINE), sheet = pivotSheet(doc), source = createObjectWorkbookSource(doc, 'pivot-guard')
    const target = doc.sheets[sheet]!, store: any = source.store(target.sheetId ?? String(sheet + 1))
    // The viewer's install: the live definition and region, then the overlay.
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => {
      target.pivotTables = target.pivotTables!.map((def, k) => k === commit.pivot ? commit.def : def)
      target.pivotRegions = target.pivotRegions?.map(region => region.partPath === commit.def.partPath ? { ...region, ref: commit.def.location.ref } : region)
      store.setValueOverlay(commit.overlay ?? undefined)
    } })
    expect((await client.open(sheet, 0, signal())).ok).toBe(true)
    expect((await client.apply(sheet, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })).ok).toBe(true)
    expect(target.pivotTables![0]!.location.ref).toBe('A3:I21')
    // Rows 22–24 (A22:I24) were the saved pivot's and are blank now: still a pivot's for sheet sort and filter.
    const vacated = { r0: 21, r1: 23, c0: 0, c1: 8, headerRows: 0 as const, totalsRows: 0 as const }
    const structure = new StructuralIndex(await store.queryStructural({ r0: 20, r1: 26, c0: 0, c1: 10 }, options(store)))
    expect(await guardSort(vacated, structure, fromStoreScan(store, options(store)))).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    const axes = { rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }
    const before = visibilityFromHidden(createIntervalSet([])), after = visibilityFromHidden(createIntervalSet([{ r0: 22, r1: 22 }]))
    const filter = await guardFilter({ range: vacated, candidateVisibility: after, currentVisibility: before, structure,
      projectionBefore: createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: before }), projectionAfter: createSheetMetrics({ ...axes, order: createIdentityOrder(), visibility: after }) })
    expect(filter).toMatchObject({ ok: false, refusal: { reason: 'pivot' } })
    // A range wholly beside both areas is not refused for a pivot.
    const beside = { r0: 30, r1: 32, c0: 0, c1: 1, headerRows: 0 as const, totalsRows: 0 as const }
    const away = new StructuralIndex(await store.queryStructural({ r0: 30, r1: 32, c0: 0, c1: 1 }, options(store)))
    expect(await guardSort(beside, away, fromStoreScan(store, options(store)))).not.toMatchObject({ refusal: { reason: 'pivot' } })
    await source.release()
  })

  it('the composite overlay over two pivots: the current area answers first, a vacated saved cell reads blank', async () => {
    const { compositeOverlay } = await import('../src/components/xlsx/pivot/pivotClient')
    const layout = (minRow: number, maxRow: number, minCol: number, maxCol: number, cells: Array<[number, number, number]>) =>
      ({ rect: { minRow, maxRow, minCol, maxCol }, merges: [], cells: cells.map(([row, col, value]) => ({ row, col, value: { kind: 'number', value } })) }) as never
    // A shrank from rows 0–5 to 0–2; B grew from rows 6–8 into row 5, a row A vacated.
    const overlay = compositeOverlay([
      { saved: [{ r0: 0, r1: 5, c0: 0, c1: 1 }], current: [{ r0: 0, r1: 2, c0: 0, c1: 1 }], layout: layout(0, 2, 0, 1, [[0, 0, 1], [2, 1, 2]]) },
      { saved: [{ r0: 6, r1: 8, c0: 0, c1: 1 }], current: [{ r0: 5, r1: 8, c0: 0, c1: 1 }], layout: layout(5, 8, 0, 1, [[5, 0, 50], [8, 1, 80]]) },
    ])
    expect(overlay.cell(0, 0)).toEqual({ value: 1 })
    expect(overlay.cell(1, 0)).toBeNull()          // inside A's current area, no layout value
    expect(overlay.cell(4, 1)).toBeNull()          // A vacated it
    expect(overlay.cell(5, 0)).toEqual({ value: 50 })  // B's current area wins over A's vacated saved area
    expect(overlay.cell(8, 1)).toEqual({ value: 80 })
    expect(overlay.cell(9, 0)).toBeUndefined()     // owned by neither
    // Ruling S3-G: protection is every saved area ∪ every current area.
    expect(overlay.protectedRects).toEqual([{ r0: 0, r1: 5, c0: 0, c1: 1 }, { r0: 0, r1: 2, c0: 0, c1: 1 }, { r0: 6, r1: 8, c0: 0, c1: 1 }, { r0: 5, r1: 8, c0: 0, c1: 1 }])
  })

  it('a chart-linked pivot follows Excel: an apply refills its chart as Excel saved it, reset gives the saved chart back', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(CHART_LINKED), sheet = pivotSheet(doc)
    const saved = linkedCharts(doc)
    expect(saved).toHaveLength(1)
    const commits: any[] = []
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => commits.push(commit) })
    const opened = await client.open(sheet, 0, signal())
    expect(opened.ok, JSON.stringify(opened)).toBe(true)
    expect((await client.apply(sheet, 0, { kind: 'items', field: 0, hidden: [3] })).ok).toBe(true)
    expect(commits[0].charts).toHaveLength(1)
    expect(chartData(commits[0].charts[0])).toEqual(chartData(linkedCharts(await load(chartCase('chart-hide-item')))[0]!))
    expect((await client.apply(sheet, 0, 'reset')).ok).toBe(true)
    expect(chartData(commits[1].charts[0])).toEqual(chartData(saved[0]!))
  })

  it('R2-4: a chart links its pivot by sheet and name (pivot names repeat across sheets)', async () => {
    const { pivotSourceRef } = await import('../src/xlsx/pivotChartLink')
    expect(pivotSourceRef('[chart-baseline.xlsx]Pivot!樞紐分析表2')).toEqual({ sheet: 'Pivot', pivot: '樞紐分析表2' })
    expect(pivotSourceRef("'[book.xlsx]My ''Q'' Sheet'!PivotTable1")).toEqual({ sheet: "My 'Q' Sheet", pivot: 'PivotTable1' })
    expect(pivotSourceRef('PivotTable1')).toEqual({ pivot: 'PivotTable1' })
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const { installPivotCharts } = await import('../src/components/xlsx/sheetAssets')
    const doc = await load(CHART_LINKED), sheet = pivotSheet(doc)
    const chart = linkedCharts(doc)[0]!
    // The same pivot name on another sheet: this chart is not the Pivot sheet's pivot's chart.
    chart.pivot = { ...chart.pivot!, name: '[chart-baseline.xlsx]Other!樞紐分析表2' }
    const commits: any[] = []
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: commit => commits.push(commit) })
    expect((await client.open(sheet, 0, signal())).ok).toBe(true)
    expect((await client.apply(sheet, 0, { kind: 'items', field: 0, hidden: [3] })).ok).toBe(true)
    expect(commits[0].charts).toBeUndefined()
    expect(installPivotCharts(doc, '樞紐分析表2', [chart], 'Pivot')).toEqual([])
    expect(linkedCharts(doc)[0]).toBe(chart)
  })

  it('S3-Q: a chart naming a pivot without its sheet links only a unique pivot; two sheets with that name both refuse', async () => {
    const { linkPivotCharts } = await import('../src/xlsx/pivotChartLink')
    const named = (name: string) => ({ pivot: { name } })
    const both = [{ sheet: 'A', name: 'PivotTable1' }, { sheet: 'B', name: 'PivotTable1' }]
    expect(linkPivotCharts([named('PivotTable1')], both[0]!, both)).toEqual({ charts: [], ambiguous: named('PivotTable1') })
    expect(linkPivotCharts([named('[x]B!PivotTable1')], both[0]!, both)).toEqual({ charts: [] })
    expect(linkPivotCharts([named('[x]A!PivotTable1')], both[0]!, both).charts).toHaveLength(1)
    expect(linkPivotCharts([named('PivotTable1')], both[0]!, both.slice(0, 1)).charts).toHaveLength(1)
    // The client: the Pivot sheet's 樞紐分析表2 and a copy of it on the other sheet; its chart names it without a sheet.
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(CHART_LINKED), sheet = pivotSheet(doc), other = doc.sheets.findIndex((_, k) => k !== sheet)
    const chart = linkedCharts(doc)[0]!
    chart.pivot = { ...chart.pivot!, name: '樞紐分析表2' }
    const unique = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: () => {} })
    expect((await unique.open(sheet, 0, signal())).ok).toBe(true)
    const def = doc.sheets[sheet]!.pivotTables![0]!
    doc.sheets[other]!.pivotTables = [{ ...def, sheet: doc.sheets[other]!.name, partPath: 'xl/pivotTables/copy.xml' }]
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: () => {} })
    const opened = await client.open(sheet, 0, signal())
    expect(opened.ok).toBe(false)
    if (!opened.ok) expect(opened.diagnostic.feature).toBe('pivot-chart-mapping')
  })

  it('a pivot whose records cannot be leased, or whose linked chart is not covered, refuses at open with its feature', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await lineChartDoc(), sheet = pivotSheet(doc)
    const client = createResidentPivotClient({ doc, ledger: new MainMemoryLedger(), install: () => {} })
    const opened = await client.open(sheet, 0, signal())
    expect(opened.ok).toBe(false)
    if (!opened.ok) expect(opened.diagnostic.feature).toBe('pivot-chart-mapping')
    // A document whose runtime was never retained (parsed elsewhere and cloned) is display-only, named before any open.
    const cloned = structuredClone(await load(BASELINE)) as XlsxDocument
    const display = createResidentPivotClient({ doc: cloned, ledger: new MainMemoryLedger(), install: () => {} })
    expect(display.preflight?.(pivotSheet(cloned), 0)?.feature).toBe('pivot-runtime-unavailable')
  })
})

describe('the interactive viewer', () => {
  it('S3-F: mounting a resident sheet with pivots charges the runtime’s retained records copy, before any open', async () => {
    const doc = await load(BASELINE)
    const runtime = pivotRuntimeFor(doc)!
    expect(runtime.retainedBytes).toBeGreaterThan(0)
    let ledger: MainMemoryLedger | undefined
    render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
    await waitFor(() => expect(ledger).toBeDefined())
    // The runtime itself is a charged target of this ledger's document pool (other document charges, such as the
    // controls' adopted copy, sit beside it), and charging it again is a no-op: it is charged once per ledger.
    expect(ledger!.hasLifetime(runtime)).toBe(true)
    const pool = ledger!.snapshot().byPool.document
    expect(pool).toBeGreaterThanOrEqual(estimatePayloadBytes(doc) + runtime.retainedBytes)
    runtime.chargeRetained(ledger!)
    expect(ledger!.snapshot().byPool.document).toBe(pool)
  })

  it('each button opens its menu and applies its delta', async () => {
    const doc = await load(BASELINE)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const collapse = await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    expect(collapse.getAttribute('data-officeview-ui')).not.toBeNull()
    expect(collapse.getAttribute('aria-expanded')).toBe('true')
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('Apples'))
    expect(ui.container.querySelector('[role=row][aria-rowindex="7"] [role=gridcell][aria-colindex="2"]')?.getAttribute('aria-expanded')).toBe('true')
    fireEvent.click(collapse)
    const expand = await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    expect(expand.getAttribute('aria-expanded')).toBe('false')
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('North'))
    expect(ui.container.querySelector('[role=row][aria-rowindex="7"] [role=gridcell][aria-colindex="2"]')?.getAttribute('aria-expanded')).toBe('false')

    // The row-labels dropdown opens the field menu; a sort entry applies its delta.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(menu).getByRole('button', { name: 'Sort descending' }))
    await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Region' })).toBeNull())
    await waitFor(() => expect(cellText(ui.container, 5, 0)).toContain('West'))

    // The item checklist: clear West, OK.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const list = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(list).getByRole('checkbox', { name: 'West' }))
    fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(cellText(ui.container, 5, 0)).toContain('South'))
    expect(ui.queryByRole('button', { name: /West \(Region\)/ })).toBeNull()
  })

  it('a page field’s dropdown opens the page menu and selects an item', async () => {
    const doc = await load(PAGE_SINGLE)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const button = await ui.findByRole('button', { name: 'Report filter Product' }, { timeout: 10_000 })
    fireEvent.click(button)
    const menu = await ui.findByRole('dialog', { name: 'Product' })
    fireEvent.click(within(menu).getByRole('radio', { name: 'Cherries' }))
    fireEvent.click(within(menu).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Product' })).toBeNull())
    await waitFor(() => expect(cellText(ui.container, 0, 1)).toContain('Cherries'))
  })

  it('a refused pivot names its reason on every button and shows one notice', async () => {
    const doc = await lineChartDoc()
    const sheet = doc.sheets[pivotSheet(doc)]!
    ;(sheet.diagnostics ??= []).push({ kind: 'unsupported', feature: 'pivot-definition-unsupported', message: 'Load-time pivot notice: an unknown construct.' })
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    await waitFor(() => expect(status(ui.container)).toContain('Load-time pivot notice: an unknown construct.'), { timeout: 10_000 })
    const buttons = await waitFor(() => {
      const found = [...ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')]
      expect(found.length).toBeGreaterThan(1)
      return found
    }, { timeout: 10_000 })
    fireEvent.click(buttons[0]!)
    await waitFor(() => expect(buttons[0]!.getAttribute('aria-label')).toMatch(/pivot chart .* cannot follow its changes/))
    for (const button of ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')) {
      expect(button.getAttribute('aria-label')).toMatch(/pivot chart .* cannot follow its changes/)
      expect(button.getAttribute('title')).toBe(button.getAttribute('aria-label'))
      expect(button.getAttribute('aria-disabled')).toBe('true')
    }
    const text = status(ui.container)
    expect(text.split('cannot follow its changes').length - 1).toBe(1)
    expect(ui.queryByRole('dialog')).toBeNull()
    // Used again, by pointer or from its gridcell, the control says why it does nothing (final review minor 6).
    fireEvent.click(buttons[1]!)
    await waitFor(() => expect(status(ui.container)).toMatch(/Not available: .*cannot follow its changes/))
    expect(ui.queryByRole('dialog')).toBeNull()
  })

  it('in the viewer a pivot chart repaints with its pivot: the committed chart is installed, reset and unmount restore the saved one', async () => {
    const doc = await load(CHART_LINKED)
    const saved = linkedCharts(doc)[0]!
    const expected = linkedCharts(await load(chartCase('chart-collapse-item')))[0]!
    const savedDef = doc.sheets[pivotSheet(doc)]!.pivotTables![0]!
    const printed: XlsxDocument[] = []
    const printLoader = async (printDoc: XlsxDocument) => { printed.push(printDoc); return Object.assign([], { dispose() {}, refusedImages: [] }) as never }
    vi.stubGlobal('print', () => {})
    const ui = render(<InteractiveSheetDoc doc={doc} printLoader={printLoader} />)
    const described = () => ui.container.querySelector('[data-testid="officeview-sheet-charts"] li')?.getAttribute('aria-label')
    await waitFor(() => expect(described()).toBe(chartDescription(saved)), { timeout: 10_000 })
    fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    const live = linkedCharts(doc)[0]!
    expect(live).not.toBe(saved)
    expect(chartData(live)).toEqual(chartData(expected))
    await waitFor(() => expect(described()).toBe(chartDescription(live)))
    expect(described()).not.toBe(chartDescription(saved))
    // Print never runs a session (§7.10): it is handed the saved pivot and its saved chart.
    fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'p', ctrlKey: true })
    await waitFor(() => expect(printed).toHaveLength(1))
    const printedSheet = printed[0]!.sheets[pivotSheet(doc)]!
    expect(printedSheet.pivotTables![0]).toBe(savedDef)
    expect(printed[0]!.sheets.flatMap(sheet => sheetCharts(sheet.drawings)).find(chart => chart.pivot)).toBe(saved)
    expect(linkedCharts(doc)[0]).toBe(live)
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Reset to saved view' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    expect(chartData(linkedCharts(doc)[0]!)).toEqual(chartData(saved))
    fireEvent.click(ui.getByRole('button', { name: 'Collapse East (Region)' }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    ui.unmount()
    expect(linkedCharts(doc)[0]).toBe(saved)
  })

  it('a workbook loaded without its pivot runtime (parsed elsewhere) is honest before any click', async () => {
    const doc = structuredClone(await load(BASELINE)) as XlsxDocument
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const buttons = await waitFor(() => {
      const found = [...ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')]
      expect(found.length).toBeGreaterThan(1)
      return found
    }, { timeout: 10_000 })
    for (const button of buttons) {
      expect(button.getAttribute('aria-label')).toMatch(/not loaded for pivot interaction/)
      expect(button.getAttribute('aria-description')).toMatch(/^(Filter and sort|Collapse|Expand) /)
    }
    await waitFor(() => expect(status(ui.container).split('not loaded for pivot interaction').length - 1).toBe(1))
    fireEvent.click(buttons[0]!)
    await new Promise(resolve => setTimeout(resolve, 50))
    expect(ui.queryByRole('dialog')).toBeNull()
  })

  it('kinds without a passing oracle are absent from the menu', async () => {
    const doc = await load(BASELINE)
    const kinds = new Set<PivotDeltaKind>(['collapse', 'sort'])
    const ui = render(<InteractiveSheetDoc doc={doc} pivotVerifiedKinds={kinds} />)
    fireEvent.click(await ui.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 }))
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    expect(within(menu).getByRole('button', { name: 'Sort ascending' })).toBeTruthy()
    for (const absent of [/Select all/, /Clear filter/, /entire field/, /Expand all|Collapse all/, /Apply label filter/, /Apply value filter/, /Apply top 10/i, /^OK$/]) {
      expect(within(menu).queryByRole('button', { name: absent }), String(absent)).toBeNull()
    }
    expect(within(menu).queryAllByRole('checkbox')).toHaveLength(0)
    expect(within(menu).getByRole('button', { name: 'Reset to saved view' })).toBeTruthy()
    // With every kind verified the entries are there (bottom-N is never offered).
    cleanup()
    const all = render(<InteractiveSheetDoc doc={await load(BASELINE)} />)
    fireEvent.click(await all.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 }))
    const full = await all.findByRole('dialog', { name: 'Region' })
    for (const present of ['Select all', 'Clear', 'Collapse entire field', 'Expand entire field', 'Collapse all', 'Expand all', 'Apply label filter', 'Apply value filter', 'Apply top 10', 'Clear filter from Region', 'OK']) {
      expect(within(full).getByRole('button', { name: present }), present).toBeTruthy()
    }
    expect(within(full).queryByRole('option', { name: /Bottom/ })).toBeNull()
  })

  it('keyboard (Ruling S3-J): one Tab stop; Alt+arrows from the gridcell; described controls; 44 px targets; coarse pointer uses the bottom sheet', async () => {
    const doc = await load(BASELINE)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    // The buttons are pointer targets only: none is a Tab stop, and there is no toolbar.
    const buttons = [...ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')]
    expect(buttons.length).toBeGreaterThan(1)
    for (const button of buttons) expect(button.tabIndex).toBe(-1)
    expect(ui.queryByRole('toolbar', { name: 'Pivot table controls' })).toBeNull()
    // Each gridcell with a control describes it.
    const cell = (row: number, col: number) => ui.container.querySelector(`[role=row][aria-rowindex="${row + 2}"] [role=gridcell][aria-colindex="${col + 2}"]`)!
    await waitFor(() => expect(cell(5, 0).getAttribute('aria-description')).toContain('Expanded, Alt+Left to collapse'))
    expect(cell(4, 0).getAttribute('aria-description')).toContain('Filter: Region, Alt+Down to open')
    // From the grid: go to A6 (East), Alt+Left collapses, Alt+Right expands.
    const grid = ui.getByRole('grid')
    act(() => grid.focus())
    fireEvent.keyDown(grid, { key: 'Home', ctrlKey: true })
    for (let k = 0; k < 5; k++) fireEvent.keyDown(grid, { key: 'ArrowDown' })
    fireEvent.keyDown(grid, { key: 'ArrowLeft', altKey: true })
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(cell(5, 0).getAttribute('aria-description')).toContain('Collapsed, Alt+Right to expand'))
    expect(document.activeElement).toBe(grid)
    fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true })
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    // Alt+Down on the Row Labels cell (A5) opens the field menu; Escape closes it and focus is the grid's again.
    fireEvent.keyDown(grid, { key: 'ArrowUp' })
    fireEvent.keyDown(grid, { key: 'ArrowDown', altKey: true })
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    for (const button of within(menu).getAllByRole('button')) expect(Number.parseFloat(button.style.minHeight), button.textContent ?? '').toBeGreaterThanOrEqual(44)
    expect(menu.querySelector('[data-pivot-menu]')?.getAttribute('data-pivot-menu')).toBe('dropdown')
    fireEvent.keyDown(within(menu).getByRole('button', { name: 'Sort ascending' }), { key: 'Escape' })
    await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Region' })).toBeNull())
    await waitFor(() => expect(document.activeElement).toBe(grid))
    cleanup()

    // Coarse pointer: the menu is a bottom sheet; a target is 44 px wide and stays within its own row band, so the
    // ± targets of consecutive rows never overlap.
    vi.stubGlobal('matchMedia', (query: string) => ({ matches: /coarse|max-width/.test(query), media: query, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} }))
    const touch = render(<InteractiveSheetDoc doc={await load(BASELINE)} />)
    const button = await touch.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 })
    expect(Number.parseFloat(button.style.width)).toBeGreaterThanOrEqual(44)
    fireEvent.click(button)
    const sheetMenu = await touch.findByRole('dialog', { name: 'Region' })
    expect(sheetMenu.querySelector('[data-pivot-menu]')?.getAttribute('data-pivot-menu')).toBe('bottom-sheet')
    cleanup()
    // Every Region item collapsed: ± buttons on consecutive rows (6–10).
    const collapsed = render(<InteractiveSheetDoc doc={await load(`${FIXTURES}interaction/interaction-collapse-all-compact-top.xlsx`)} />)
    await collapsed.findByRole('button', { name: /^Expand East/ }, { timeout: 10_000 })
    const box = (key: string) => { const el = collapsed.container.querySelector<HTMLElement>(`[data-pivot-key="${key}"]`)!; return { top: Number.parseFloat(el.style.top), height: Number.parseFloat(el.style.height), width: Number.parseFloat(el.style.width) } }
    const first = box('0:5:0'), second = box('0:6:0')
    expect(first.width).toBeGreaterThanOrEqual(44)
    expect(first.top + first.height).toBeLessThanOrEqual(second.top)
  })

  it('Task 29: a change autofits the pivot\'s columns (a session size, like a resize) and a user-resized column keeps its width', async () => {
    const doc = await load(BASELINE)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const header = (col: number) => ui.container.querySelector<HTMLElement>(`[role=columnheader][aria-colindex="${col + 2}"]`)
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(header(0)).not.toBeNull())
    const width = () => Number.parseFloat(header(0)!.style.width)
    const saved = width()
    // Collapse the whole Region field: column A no longer holds the indented product labels.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(width()).toBeLessThan(saved))
    const narrowed = width()
    // Expanding it again brings the indented labels and the width back.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Expand entire field' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(width()).toBeGreaterThan(narrowed))
    // The user widens column A by keyboard; the next change leaves it alone.
    const grid = ui.getByRole('grid')
    act(() => grid.focus())
    fireEvent.keyDown(grid, { key: 'Home', ctrlKey: true })
    for (let k = 0; k < 3; k++) fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true, shiftKey: true })
    await waitFor(() => expect(width()).toBeGreaterThan(narrowed + 16))
    const user = width()
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    await new Promise(resolve => setTimeout(resolve, 100))
    expect(width()).toBe(user)
    ui.unmount()
  })

  it('Task 29 r1: an autofit measures in its own admitted, per-fit context and admits its geometry before building; short of memory it is refused', async () => {
    const doc = await load(BASELINE)
    let ledger: MainMemoryLedger | undefined
    const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
    const header = (col: number) => ui.container.querySelector<HTMLElement>(`[role=columnheader][aria-colindex="${col + 2}"]`)
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(header(0)).not.toBeNull())
    const width = () => Number.parseFloat(header(0)!.style.width)
    const saved = width()
    const set = vi.spyOn(ledger!, 'set'), release = vi.spyOn(ledger!, 'release')
    const contexts = vi.spyOn(HTMLCanvasElement.prototype, 'getContext')
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
    await waitFor(() => expect(width()).toBeLessThan(saved))
    const owners = set.mock.calls.map(([pool, owner]) => `${pool}:${owner}`)
    expect(owners).toContain('canvasText:pivot-autofit-measure')
    expect(owners).toContain('geometry:resize-candidate')
    expect(release.mock.calls.map(([pool, owner]) => `${pool}:${owner}`)).toEqual(expect.arrayContaining(['canvasText:pivot-autofit-measure', 'layout:pivot-autofit-measure']))
    const made = contexts.mock.calls.length
    expect(made).toBeGreaterThan(0)
    set.mockRestore(); release.mockRestore()
    // Expanding again with no main memory left: the change shows, its autofit is refused and the width stays.
    const narrowed = width()
    ledger!.set('copy', 'filler', Math.max(0, 64 * 1024 * 1024 - ledger!.bytes))
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Expand entire field' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(ui.container.textContent).toContain('Pivot table changed (view only; the file is unchanged). Its columns were not fitted: not enough memory'))
    expect(width()).toBe(narrowed)
    contexts.mockRestore()
    ledger!.set('copy', 'filler', 0)
    ui.unmount()
  })

  it('Task 29 r1: a hidden column whose fit changed stays hidden and does not refuse the others\' autofit', async () => {
    const doc = await load(BASELINE)
    const sheet = doc.sheets[pivotSheet(doc)]!
    // Column F hidden: keeping one Region moves the year columns, and F's fit changes (so does G's).
    // F and G at the default width (no saved size to keep): a changed fit resizes them.
    sheet.cols = [...sheet.cols.filter(spec => spec.min !== 5 && spec.min !== 6), { min: 5, max: 5, hidden: true }]
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    await ui.findByRole('button', { name: /^Collapse East/ }, { timeout: 10_000 })
    const columnF = () => ui.container.querySelector('[role=columnheader][aria-label^="Column F"]')
    const columnG = () => ui.container.querySelector<HTMLElement>('[role=columnheader][aria-label^="Column G"]')
    await waitFor(() => expect(columnG()).not.toBeNull())
    expect(columnF()).toBeNull()
    const savedG = Number.parseFloat(columnG()!.style.width)
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const list = await ui.findByRole('dialog', { name: 'Region' })
    const boxes = within(list).getAllByRole('checkbox').filter(box => ['East', 'North', 'South', 'West', 'Central'].includes(box.getAttribute('aria-label') ?? ''))
    for (const box of boxes.slice(0, -1)) fireEvent.click(box)
    fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ui.container.textContent).toContain('Pivot table changed (view only; the file is unchanged).'), { timeout: 10_000 })
    expect(ui.container.textContent).not.toMatch(/not fitted/)
    expect(columnF()).toBeNull()
    // G's fit changed and its width was its old fit: it is resized beside the hidden F.
    await waitFor(() => expect(Number.parseFloat(columnG()!.style.width)).not.toBe(savedG))
    ui.unmount()
  }, 20_000)

  it('Task 29 r2: putting the autofit widths back when the pivots go is admitted like any resize; short of memory the widths stay', async () => {
    for (const squeezed of [false, true]) {
      const doc = await load(BASELINE)
      let ledger: MainMemoryLedger | undefined
      const kinds = new Set<PivotDeltaKind>(['collapse', 'collapseField', 'collapseAll', 'items', 'page', 'filter', 'clearFilters', 'sort'])
      const view = (verified: ReadonlySet<PivotDeltaKind>) => <InteractiveSheetDoc doc={doc} pivotVerifiedKinds={verified} onMainLedger={value => { ledger = value }} />
      const ui = render(view(kinds))
      const header = () => ui.container.querySelector<HTMLElement>('[role=columnheader][aria-colindex="2"]')
      await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
      await waitFor(() => expect(header()).not.toBeNull())
      const width = () => Number.parseFloat(header()!.style.width)
      const saved = width()
      fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
      fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
      await waitFor(() => expect(width()).toBeLessThan(saved))
      const narrowed = width()
      const set = vi.spyOn(ledger!, 'set')
      // No room at all (the filler alone is the whole budget): nothing the restore asks for can be admitted.
      if (squeezed) ledger!.set('copy', 'filler', 64 * 1024 * 1024)
      // A new pivot client (the verified kinds changed): the old one's pivots go back to the file, widths too.
      ui.rerender(view(new Set(kinds)))
      await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
      // Reserved before it is built (refused before even that when there is no room).
      expect(set.mock.calls.some(([pool, owner]) => pool === 'geometry' && owner === 'resize-candidate')).toBe(!squeezed)
      if (squeezed) {
        expect(width()).toBe(narrowed)
        expect(ui.container.textContent).toContain('fitted column widths were kept: not enough memory')
        ledger!.set('copy', 'filler', 0)
      } else await waitFor(() => expect(width()).toBe(saved))
      set.mockRestore()
      ui.unmount()
    }
  }, 30_000)

  it('reset returns the saved view', async () => {
    const doc = await load(BASELINE)
    const saved = doc.sheets[pivotSheet(doc)]!.pivotTables![0]
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(menu).getByRole('button', { name: 'Reset to saved view' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('Apples'))
    expect(doc.sheets[pivotSheet(doc)]!.pivotTables![0]).toBe(saved)
    // Unmounting gives the host's document its saved definitions back.
    fireEvent.click(ui.getByRole('button', { name: 'Collapse East (Region)' }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
    expect(doc.sheets[pivotSheet(doc)]!.pivotTables![0]).not.toBe(saved)
    // The pivot's region (no gridlines inside it) follows the live rectangle: A3:I24 saved, A3:I21 collapsed.
    expect(doc.sheets[pivotSheet(doc)]!.pivotRegions![0]!.ref).toBe('A3:I21')
    ui.unmount()
    expect(doc.sheets[pivotSheet(doc)]!.pivotTables![0]).toBe(saved)
    expect(doc.sheets[pivotSheet(doc)]!.pivotRegions![0]!.ref).toBe('A3:I24')
  })
})

/** An in-process worker port over a ViewerKernel (tests/xlsx-viewer-paged-viewer.test.tsx's harness, trimmed). */
function workerHarness(env?: ConstructorParameters<typeof ViewerKernel>[0], intercept?: (command: WorkerCommand) => WorkerReply | undefined) {
  // Admission forced paged (src/debug/forcePaged.ts): these small fixtures take the worker-backed paged path.
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS, ...env }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      const answered = intercept?.(command)
      if (answered) { queueMicrotask(() => { for (const listener of [...listeners]) listener({ data: answered } as MessageEvent) }); return }
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of [...listeners]) listener({ data } as MessageEvent)
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return { port, posted }
}
const pagedFile = (file: string) => new Blob([new Uint8Array(readFileSync(file))])
/** The paged route OfficeDoc takes: the worker-backed shell document and its source. */
async function renderPaged(h: ReturnType<typeof workerHarness>, file: string) {
  const doc = await openPagedWorkbook(pagedFile(file), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
  const handle = doc && pagedWorkbookOf(doc)
  if (!doc || !handle) throw new Error('the workbook did not open paged')
  return render(<InteractiveSheetDoc doc={doc} source={handle.source} />)
}

describe('the paged path (worker sessions through RemotePivotSessions)', () => {
  it('opens, applies, commits and resets in the worker; the menus carry the worker’s model', async () => {
    const h = workerHarness()
    const ui = await renderPaged(h, BASELINE)
    // Before the session opens the worker's labels are unknown: the item's own text names its button.
    const collapse = await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('Apples'), FIRST_PAINT_WAIT)
    fireEvent.click(collapse)
    await ui.findByRole('button', { name: 'Expand East (Region)' }, FIRST_PAINT_WAIT)
    expect(h.posted.map(command => command.kind)).toEqual(expect.arrayContaining(['pivotOpen', 'pivotApply', 'pivotCommit']))
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('North'), FIRST_PAINT_WAIT)
    // The field menu: item labels from the worker's model; a sort applies through the worker.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    expect(within(menu).getAllByRole('checkbox').map(box => box.getAttribute('aria-label'))).toEqual(expect.arrayContaining(['East', 'North', 'South', 'West', 'Central']))
    fireEvent.click(within(menu).getByRole('button', { name: 'Reset to saved view' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('Apples'), FIRST_PAINT_WAIT)
    ui.unmount()
    await waitFor(() => expect(h.posted.some(command => command.kind === 'pivotRelease')).toBe(true))
  }, 60_000)

  it('a reopen the port makes on its own (the worker released the session) ends when the viewer unmounts', async () => {
    const { remotePivotPort } = await import('../src/components/xlsx/pivot/pivotClient')
    let reopenSignal: AbortSignal | undefined
    const sessions = {
      open: (_sheet: number, _pivot: number, signal: AbortSignal) => { reopenSignal = signal; return new Promise<never>(() => {}) },
      apply: async () => ({ kind: 'refused' as const, message: 'Open the pivot before changing it.' }),
      release: () => {},
    }
    const port = remotePivotPort(sessions, { id: sheet => sheet, sheet: () => undefined })
    void port.apply(0, 0, { kind: 'collapse', field: 0, item: 0, collapsed: true })
    await waitFor(() => expect(reopenSignal).toBeDefined())
    expect(reopenSignal!.aborted).toBe(false)
    port.dispose?.()
    expect(reopenSignal!.aborted).toBe(true)
  })

  it('Ruling S3-AK: a change stays shown after switching to another sheet and back, and its menu still says so', async () => {
    const h = workerHarness()
    const ui = await renderPaged(h, BASELINE)
    fireEvent.click(await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('North'), FIRST_PAINT_WAIT)
    const picker = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
    const pivotSheetId = picker.value, other = [...picker.options].find(option => option.value !== pivotSheetId)!.value
    fireEvent.change(picker, { target: { value: other } })
    await waitFor(() => expect(ui.queryByRole('button', { name: 'Expand East (Region)' })).toBeNull(), FIRST_PAINT_WAIT)
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: pivotSheetId } })
    // The worker kept the committed overlay: the collapsed layout shows again, with its controls.
    await ui.findByRole('button', { name: 'Expand East (Region)' }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('North'), FIRST_PAINT_WAIT)
    expect(h.posted.some(command => command.kind === 'pivotRelease')).toBe(false)
    // The menu and the worker agree: the next change applies to the kept session (no reopen), and the reset is real.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Reset to saved view' }))
    await ui.findByRole('button', { name: 'Collapse East (Region)' }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(cellText(ui.container, 6, 0)).toContain('Apples'), FIRST_PAINT_WAIT)
    expect(h.posted.filter(command => command.kind === 'pivotOpen')).toHaveLength(1)
    ui.unmount()
    await waitFor(() => expect(h.posted.some(command => command.kind === 'pivotRelease')).toBe(true))
  }, 60_000)

  it('a chart sheet whose drawings changed while inactive drops its cached tiles when shown again', async () => {
    const { TileCache } = await import('../src/xlsx/view/tileCache')
    const h = workerHarness()
    const doc = await openPagedWorkbook(pagedFile(`${FIXTURES}interaction/chart-hide-item.xlsx`), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
    const handle = doc && pagedWorkbookOf(doc)
    if (!doc || !handle) throw new Error('the workbook did not open paged')
    const frames: string[] = []
    const ui = render(<InteractiveSheetDoc doc={doc} source={handle.source} onFrame={stats => frames.push(stats.activeSheetId)} />)
    const remote = handle.source as unknown as { drawings(sheetId: string): { state: string; gen: number } }
    const chartSheet = doc.sheets[0]!.sheetId ?? '1', otherSheet = doc.sheets[1]!.sheetId ?? '2'
    await waitFor(() => expect(remote.drawings(chartSheet).state).toBe('ready'), FIRST_PAINT_WAIT)
    await waitFor(() => expect(frames.includes(chartSheet)).toBe(true), FIRST_PAINT_WAIT)
    await new Promise(resolve => setTimeout(resolve, 100))
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: otherSheet } })
    await waitFor(() => expect(frames.at(-1)).toBe(otherSheet), FIRST_PAINT_WAIT)
    // Only the chart sheet's drawing generation moves (a chart swapped while it is not shown; no other generation).
    ;(remote.drawings(chartSheet) as { gen: number }).gen++
    const invalidate = vi.spyOn(TileCache.prototype, 'invalidate')
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: chartSheet } })
    await waitFor(() => expect(frames.at(-1)).toBe(chartSheet), FIRST_PAINT_WAIT)
    const key = (sheetId: string) => ({ identity: { docId: '', sheetId, valueGen: 0, fontGen: 0, viewGen: 0 }, col: 0, row: 0, zoom: 1, dpr: 1 }) as never
    expect(invalidate.mock.calls.some(([predicate]) => predicate(key(chartSheet)) && !predicate(key(otherSheet)))).toBe(true)
    invalidate.mockRestore()
  }, 60_000)

  it('Task 29: a paged change autofits the pivot\'s columns too', async () => {
    const h = workerHarness()
    const ui = await renderPaged(h, BASELINE)
    const header = (col: number) => ui.container.querySelector<HTMLElement>(`[role=columnheader][aria-colindex="${col + 2}"]`)
    await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT)
    await waitFor(() => expect(header(0)).not.toBeNull())
    const width = () => Number.parseFloat(header(0)!.style.width)
    const saved = width()
    // The first use opens the worker session (its labels then name the menu).
    fireEvent.click(ui.getByRole('button', { name: /^Collapse East/ }))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, FIRST_PAINT_WAIT)
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
    await waitFor(() => expect(width()).toBeLessThan(saved), FIRST_PAINT_WAIT)
    ui.unmount()
  }, 60_000)

  it('a paged autofit waiting for a pivot display that cannot load is dropped with its charge, not kept until close', async () => {
    let failDisplay = false
    const h = workerHarness(undefined, command => failDisplay && command.kind === 'read' && command.query.op === 'pivotDisplay'
      ? { kind: 'error', commandKind: 'read', requestId: command.requestId, owner: command.owner, code: 'page-error', retryable: false, diagnostics: [{ code: 'test', message: 'The pivot display could not be read.' }] } as unknown as WorkerReply
      : undefined)
    const doc = await openPagedWorkbook(pagedFile(BASELINE), { port: h.port as never, profile: 'D', signal: new AbortController().signal })
    const handle = doc && pagedWorkbookOf(doc)
    if (!doc || !handle) throw new Error('the workbook did not open paged')
    const ui = render(<InteractiveSheetDoc doc={doc} source={handle.source} />)
    const ledger = (handle.source as unknown as { mainLedger: MainMemoryLedger }).mainLedger
    const fitCharges = () => [...((ledger as unknown as { charges: Map<string, unknown> }).charges.keys())].filter(key => key.includes('pivot-fit:'))
    fireEvent.click(await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT))
    await ui.findByRole('button', { name: 'Expand East (Region)' }, FIRST_PAINT_WAIT)
    // The second change waits for the refreshed display, which now fails for good.
    failDisplay = true
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Collapse entire field' }))
    await waitFor(() => expect(h.posted.filter(command => command.kind === 'pivotCommit')).toHaveLength(2), FIRST_PAINT_WAIT)
    const remote = handle.source as unknown as { pivotDisplay(sheetId: string): { state: string } }
    const sheetId = (ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement).value
    await waitFor(() => expect(remote.pivotDisplay(sheetId).state).toBe('failed'), FIRST_PAINT_WAIT)
    await waitFor(() => expect(fitCharges()).toEqual([]), FIRST_PAINT_WAIT)
    ui.unmount()
  }, 60_000)

  it('a pivot the worker refuses names its reason on every button, with one notice', async () => {
    const file = `${FIXTURES}interaction/chart-hide-item.xlsx`
    const h = workerHarness({
      parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)),
      openSource: async (source, docId) => {
        if (source.kind !== 'blob') throw new Error('test expects a blob source')
        const session = await openImportSession(source, { docId, profile: 'D' })
        await session.done
        const paged = createPagedWorkbookSource(session)
        const assets = paged.pivotAssets!.bind(paged)
        // The linked chart as a line chart: Task 30's gate does not cover it, so the worker refuses at open.
        paged.pivotAssets = async (...args) => {
          const found = await assets(...args)
          return found && { ...found, linkedCharts: found.linkedCharts.map(item => ({ ...item, chart: { ...item.chart, groups: item.chart.groups.map(group => ({ ...group, family: 'line' as const })) } })) }
        }
        return paged
      },
    })
    const ui = await renderPaged(h, file)
    const button = await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT)
    fireEvent.click(button)
    await waitFor(() => expect(button.getAttribute('aria-label')).toMatch(/cannot follow its changes/), FIRST_PAINT_WAIT)
    for (const other of ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')) {
      expect(other.getAttribute('aria-label')).toMatch(/cannot follow its changes/)
      expect(other.getAttribute('aria-disabled')).toBe('true')
    }
    expect(status(ui.container).split('cannot follow its changes').length - 1).toBe(1)
    // Used again, the control says why it does nothing (final review minor 6).
    fireEvent.click(button)
    await waitFor(() => expect(status(ui.container)).toMatch(/Not available: .*cannot follow its changes/), FIRST_PAINT_WAIT)
  }, 60_000)
})

describe('a refreshed resident pivot (PR #57 review: labels from the session cache)', () => {
  it('after its first change, its collapse buttons name their own items, and hiding North hides North', async () => {
    const doc = await load(`${FIXTURES}recipes/refresh-seed.xlsx`)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const WORDS = ['East', 'North', 'South', 'West', 'Central', 'Apples', 'Bananas', 'Cherries', 'Elderberries']
    // Each row's label in column A, as the word it shows (the gridcell's text also carries its accessible prefix).
    const columnA = () => Array.from({ length: 30 }, (_, row) => { const text = cellText(ui.container, row, 0) ?? ''; return WORDS.find(word => new RegExp(`\\b${word}\\b`).test(text)) })
    // The first change refreshes the pivot (its rebuilt cache moves North ahead of East).
    fireEvent.click(await ui.findByRole('button', { name: 'Collapse South (Region)' }, { timeout: 10_000 }))
    await ui.findByRole('button', { name: 'Expand South (Region)' }, { timeout: 10_000 })
    await waitFor(() => expect(columnA()).toContain('East'))
    const east = columnA().indexOf('East'), north = columnA().indexOf('North')
    expect(columnA()[east + 1]).toBe('Apples')
    expect(columnA()[north + 1]).toBe('Apples')
    // "Collapse North" collapses North, not East.
    fireEvent.click(ui.getByRole('button', { name: 'Collapse North (Region)' }))
    await waitFor(() => expect(columnA()[north + 1]).toBe('South'), { timeout: 10_000 })
    expect(columnA()[east + 1]).toBe('Apples')
    expect(ui.getByRole('button', { name: 'Collapse East (Region)' })).toBeDefined()
    expect(ui.getByRole('button', { name: 'Expand North (Region)' })).toBeDefined()
    // Hiding North in the checklist hides North.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const list = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(list).getByRole('checkbox', { name: 'North' }))
    fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(columnA()).not.toContain('North'), { timeout: 10_000 })
    expect(columnA()).toContain('East')
    ui.unmount()
  }, 30_000)
})

// PR #57 review (bot round 6): a saved multi-item report filter is `selected: 'all'` plus hidden items
// (interaction-page-multi: Bananas and Dates hidden). The page menu opens on that effective selection, an untouched OK
// changes nothing, and checking (All) gives Excel's interaction-page-all.
const PAGE_MULTI = `${FIXTURES}interaction/interaction-page-multi.xlsx`
async function pageMultiMenu(ui: ReturnType<typeof render>) {
  fireEvent.click(await ui.findByRole('button', { name: /^Report filter/ }, FIRST_PAINT_WAIT))
  const menu = await ui.findByRole('dialog', { name: 'Product' }, FIRST_PAINT_WAIT)
  const checked = (name: string) => (within(menu).getByRole('checkbox', { name }) as HTMLInputElement).checked
  return { menu, checked }
}
async function pageMultiFlow(ui: ReturnType<typeof render>, applies: () => number) {
  await waitFor(() => expect(cellText(ui.container, 0, 1)).toContain('(多重項目)'), FIRST_PAINT_WAIT)
  const first = await pageMultiMenu(ui)
  // Multi mode, the saved selection checked: Apples and Cherries, not Bananas, Dates or (All).
  expect(first.checked('Select multiple items')).toBe(true)
  const items = within(first.menu).getAllByRole('checkbox').filter(box => !['Select multiple items', '(All)'].includes(box.getAttribute('aria-label') ?? ''))
  expect(items.map(box => box.getAttribute('aria-label'))).toEqual(expect.arrayContaining(['Apples', 'Bananas', 'Cherries']))
  expect(items.map(box => (box as HTMLInputElement).checked)).toEqual([true, false, true, false])
  expect(first.checked('(All)')).toBe(false)
  const before = applies()
  fireEvent.click(within(first.menu).getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Product' })).toBeNull())
  await new Promise(resolve => setTimeout(resolve, 200))
  // Untouched: no change was applied, and the pivot is the saved one.
  expect(applies()).toBe(before)
  expect(cellText(ui.container, 0, 1)).toContain('(多重項目)')
  expect(cellText(ui.container, 9, 0)).toContain('Grand Total')
  // (All) shows every product: Excel's interaction-page-all (grand total on row 11).
  const second = await pageMultiMenu(ui)
  fireEvent.click(within(second.menu).getByRole('checkbox', { name: '(All)' }))
  fireEvent.click(within(second.menu).getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(cellText(ui.container, 10, 0)).toContain('Grand Total'), FIRST_PAINT_WAIT)
  expect(cellText(ui.container, 0, 1)).toContain('(全部)')
}

describe('a multi-item report filter saved as (All) with hidden items (PR #57 review)', () => {
  it('resident: the menu opens on the effective selection; an untouched OK is a no-op; (All) gives Excel\'s', async () => {
    const doc = await load(PAGE_MULTI)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    let commits = 0
    const pivots = () => doc.sheets[pivotSheet(doc)]!.pivotTables![0]!
    const saved = pivots()
    await pageMultiFlow(ui, () => { if (pivots() !== saved) commits++; return commits })
  }, 60_000)

  it('worker: the same, through the worker\'s menu model', async () => {
    const h = workerHarness()
    const ui = await renderPaged(h, PAGE_MULTI)
    await pageMultiFlow(ui, () => h.posted.filter(command => command.kind === 'pivotApply').length)
  }, 60_000)
})

describe('a multi-item report filter: adding a hidden item (PR #57 review)', () => {
  it('a listed selection shows every item it names, the saved hidden flags notwithstanding', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const multi = await load(PAGE_MULTI), all = await load(`${FIXTURES}interaction/interaction-page-all.xlsx`)
    const client = createResidentPivotClient({ doc: multi, ledger: new MainMemoryLedger(), install: () => {} })
    const opened = await client.open(pivotSheet(multi), 0, signal())
    if (!opened.ok) throw new Error(JSON.stringify(opened))
    const product = opened.model.fields.find(field => field.axis === 'page')!
    expect(product.page).toEqual([0, 2])
    // Every product, listed: the same layout as Excel's (All) (interaction-page-all, A3:I11).
    const listed = await client.apply(pivotSheet(multi), 0, { kind: 'page', field: product.index, selected: product.items.map(item => item.index) })
    if (!listed.ok) throw new Error(JSON.stringify(listed))
    expect(listed.def.location.ref).toBe(all.sheets[pivotSheet(all)]!.pivotTables![0]!.location.ref)
  })
})
