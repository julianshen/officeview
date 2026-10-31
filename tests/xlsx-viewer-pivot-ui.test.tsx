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

const FIXTURES = `${__dirname}/fixtures/xlsx-pivots/`
const BASELINE = `${FIXTURES}engine/baseline.xlsx`
const PAGE_SINGLE = `${FIXTURES}engine/quirk-page-single.xlsx`
const CHART_LINKED = `${FIXTURES}interaction/chart-baseline.xlsx`
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

  it('a pivot whose records cannot be leased, or with a linked chart, refuses at open with its feature', async () => {
    const { createResidentPivotClient } = await import('../src/components/xlsx/pivot/pivotClient')
    const doc = await load(CHART_LINKED), sheet = pivotSheet(doc)
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
    const doc = await load(CHART_LINKED)
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
    await waitFor(() => expect(buttons[0]!.getAttribute('aria-label')).toMatch(/linked chart/))
    for (const button of ui.container.querySelectorAll<HTMLButtonElement>('button[data-officeview-ui="pivot-button"]')) {
      expect(button.getAttribute('aria-label')).toMatch(/linked chart/)
      expect(button.getAttribute('title')).toBe(button.getAttribute('aria-label'))
      expect(button.getAttribute('aria-disabled')).toBe('true')
    }
    const text = status(ui.container)
    expect(text.split('linked chart').length - 1).toBe(1)
    expect(ui.queryByRole('dialog')).toBeNull()
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

  it('keyboard: roving focus, Escape, 44 px targets; coarse pointer uses the bottom sheet', async () => {
    const doc = await load(BASELINE)
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    const first = await ui.findByRole('button', { name: 'Filter and sort Year' }, { timeout: 10_000 })
    const layer = first.closest('[role=toolbar]') as HTMLElement
    expect(layer.getAttribute('aria-label')).toBe('Pivot table controls')
    const buttons = within(layer).getAllByRole('button')
    expect(buttons.filter(button => button.tabIndex === 0)).toHaveLength(1)
    act(() => first.focus())
    fireEvent.keyDown(first, { key: 'ArrowRight' })
    expect(document.activeElement).toBe(buttons[1])
    expect(buttons[1]!.tabIndex).toBe(0)
    expect(first.tabIndex).toBe(-1)
    fireEvent.keyDown(buttons[1]!, { key: 'End' })
    expect(document.activeElement).toBe(buttons.at(-1))
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' })
    expect(document.activeElement?.getAttribute('role')).toBe('grid')
    // Menu controls are 44 px targets; Escape closes the menu.
    fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    for (const button of within(menu).getAllByRole('button')) expect(Number.parseFloat(button.style.minHeight), button.textContent ?? '').toBeGreaterThanOrEqual(44)
    expect(menu.querySelector('[data-pivot-menu]')?.getAttribute('data-pivot-menu')).toBe('dropdown')
    fireEvent.keyDown(within(menu).getByRole('button', { name: 'Sort ascending' }), { key: 'Escape' })
    await waitFor(() => expect(ui.queryByRole('dialog', { name: 'Region' })).toBeNull())
    cleanup()

    // Coarse pointer: the menu is a bottom sheet and the buttons are 44 px targets.
    vi.stubGlobal('matchMedia', (query: string) => ({ matches: /coarse|max-width/.test(query), media: query, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} }))
    const touch = render(<InteractiveSheetDoc doc={await load(BASELINE)} />)
    const button = await touch.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 })
    expect(Number.parseFloat(button.style.width)).toBeGreaterThanOrEqual(44)
    expect(Number.parseFloat(button.style.height)).toBeGreaterThanOrEqual(44)
    fireEvent.click(button)
    const sheetMenu = await touch.findByRole('dialog', { name: 'Region' })
    expect(sheetMenu.querySelector('[data-pivot-menu]')?.getAttribute('data-pivot-menu')).toBe('bottom-sheet')
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
    ui.unmount()
    expect(doc.sheets[pivotSheet(doc)]!.pivotTables![0]).toBe(saved)
  })
})
