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
    expect([...opened.verifiedKinds].sort()).toEqual(['clearFilters', 'collapse', 'collapseAll', 'collapseField', 'filter', 'items', 'page', 'sort'])
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
function workerHarness(env?: ConstructorParameters<typeof ViewerKernel>[0]) {
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
  }, 60_000)
})
