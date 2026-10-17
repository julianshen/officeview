import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as controlsModule from '../src/components/xlsx/viewerControlsClient'
import { workbookOf } from './helpers/xlsx-find'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import { preparePrintUnits } from '../src/components/xlsx/sheetAssets'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
const docOf = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Name' }, { row: 0, col: 1, value: 'Score' }, { row: 1, col: 0, value: 'beta' }, { row: 1, col: 1, value: 2 }, { row: 2, col: 0, value: 'alpha' }, { row: 2, col: 1, value: 1 }] }])
afterEach(() => { cleanup(); vi.restoreAllMocks(); vi.unstubAllGlobals() })
function clipboard() { const write = vi.fn(async () => {}); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: write } }); return write }
function go(ui: ReturnType<typeof render>, at: string) { fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } }); fireEvent.click(ui.getByRole('button', { name: 'Go To' })); const grid = ui.getByRole('grid'); act(() => grid.focus()); return grid }
it('copies the active source selection on Ctrl/Cmd+C and on the UI button; input copy remains native', async () => {
  const write = clipboard(), ui = render(<InteractiveSheetDoc doc={docOf()} />)
  const grid = go(ui, 'A2'); fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
  await waitFor(() => expect(write).toHaveBeenCalledWith('beta'))
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' })); await waitFor(() => expect(write).toHaveBeenCalledTimes(2))
  fireEvent.keyDown(ui.getByRole('textbox', { name: 'Go To cell' }), { key: 'c', ctrlKey: true }); expect(write).toHaveBeenCalledTimes(2)
})
it('denied keyboard/UI/programmatic paths make zero clipboard or print calls', async () => {
  const write = clipboard(), print = vi.spyOn(window, 'print').mockImplementation(() => {})
  const ui = render(<InteractiveSheetDoc doc={docOf()} allowCopy={false} allowPrint={false} />)
  const grid = go(ui, 'A2'); fireEvent.keyDown(grid, { key: 'c', metaKey: true }); fireEvent.keyDown(grid, { key: 'p', ctrlKey: true })
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' })); fireEvent.click(ui.getByRole('button', { name: 'Print' }))
  expect(write).not.toHaveBeenCalled(); expect(print).not.toHaveBeenCalled()
  const loader = vi.fn(); await expect(preparePrintUnits(docOf(), { allowed: () => false, assetLoader: loader })).rejects.toMatchObject({ code: 'permission-denied' }); expect(loader).not.toHaveBeenCalled()
})
it('header sort cycle publishes view order atomically and preserves source selection for copy', async () => {
  const write = clipboard(), doc = docOf()
  doc.tables = [{ id: 'T', name: 'T', displayName: 'T', sheetId: 'A', partPath: 'table.xml', extent: { sheetId: 'A', firstRow: 0, firstCol: 0, rows: 3, cols: 2 }, headerRowCount: 1, totalsRowCount: 0, columns: [{ id: '1', name: 'Name', index: 0 }, { id: '2', name: 'Score', index: 1 }] }]; const ui = render(<InteractiveSheetDoc doc={doc} />); await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); const grid = go(ui, 'A2'); fireEvent.keyDown(grid, { key: 'ArrowDown', shiftKey: true })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
  await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' })); await waitFor(() => expect(write).toHaveBeenCalledWith('alpha\nbeta'))
  expect(doc.sheets[0].rows[1].cells[0].value).toBe('beta')
})
it('mounts filter and sort drafts; Cancel does not change the view, OK commits visibility/counts', async () => {
  const ui = render(<InteractiveSheetDoc doc={docOf()} />); await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); go(ui, 'A2')
  fireEvent.click(ui.getByRole('button', { name: 'More sort options for A' }))
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy()); fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  expect(ui.getByRole('columnheader', { name: 'Column A' }).getAttribute('aria-sort')).toBeNull()
  fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beta' })).toBeTruthy())
  fireEvent.click(ui.getByRole('checkbox', { name: 'beta' })); await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false)); fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 1 \/ complete body 2/))
  expect(ui.getByRole('button', { name: /Filter column A/ }).getAttribute('aria-pressed')).toBe('true')
})
it('prints an explicit large-export refusal with the approved wording', async () => {
  const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'Tall', rows: [{ r: 20000, cells: [{ ref: 'A20000', v: 'tail' }] }] }])))
  const print = vi.spyOn(window, 'print').mockImplementation(() => {}), ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'p', ctrlKey: true })
  await waitFor(() => expect(print).toHaveBeenCalled())
  expect(ui.getByTestId('officeview-sheet-print-pages').textContent).toContain('This sheet is too large to export as one image. Choose a range or print selected pages.')
  expect(ui.getByRole('status').textContent).toContain('This sheet is too large to export as one image.')
})
it('shows the saved-values notice once in the interactive status', async () => {
  const doc = docOf(); doc.valuePolicy = 'cached'; doc.diagnostics = [{ kind: 'large-source', feature: 'cached-value-policy', message: 'Showing saved values; formulas are not recalculated' }]
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Showing saved values/))
  expect(ui.getByRole('status').textContent?.match(/Showing saved values/g)).toHaveLength(1)
})

it('pending copy can be cancelled and permission withdrawal prevents its eventual clipboard write', async () => {
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const doc = docOf(), source = createObjectWorkbookSource(doc, 'pending'), store = source.store('A'), write = clipboard()
  const read = store.readCell.bind(store)
  let complete!: () => void, loaded = false
  store.readCell = async (at, options) => loaded ? read(at, options) : { kind: 'unknown', pageId: 'p' }
  store.ensureComplete = async range => { await new Promise<void>(resolve => { complete = resolve }); loaded = true; return { range, stamp: store.stamp, completeness: 'complete', missing: [] } }
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />); await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); const grid = go(ui, 'A2')
  fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Copy is waiting/))
  ui.rerender(<InteractiveSheetDoc doc={doc} source={source} allowCopy={false} />)
  await act(async () => complete())
  expect(write).not.toHaveBeenCalled()
})
it('rereads counted saved-value status and open diagnostics after a progressive source completes', async () => {
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const { valueStatus } = await import('../src/xlsx/view/values')
  const doc = docOf(); doc.valuePolicy = 'cached'; doc.diagnostics = [{ kind: 'large-source', feature: 'cached-value-policy', message: 'old duplicate' }]
  const base = createObjectWorkbookSource(doc, 'paged'); let formulas: number | undefined
  const source = { ...base, valueStatus: vi.fn(() => valueStatus('cached', formulas)), get openDiagnostics() { return [{ code: 'cached-value-policy', message: valueStatus('cached', formulas).message! }] } }
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Showing saved values; formulas are not recalculated'))
  formulas = 7
  await waitFor(() => expect(ui.getByRole('status').textContent).toContain('7 formulas not recalculated'))
  expect(ui.getByRole('status').textContent?.match(/Showing saved values/g)).toHaveLength(1)
  expect(ui.getByRole('status').textContent).not.toContain('old duplicate')
})

it('diagnoses a rejected asynchronous print request instead of leaving an unhandled rejection', async () => {
  const print = vi.spyOn(window, 'print').mockImplementation(() => {})
  const ui = render(<InteractiveSheetDoc doc={docOf()} printLoader={async () => { throw new Error('Print unavailable; try again.') }} />)
  fireEvent.click(ui.getByRole('button', { name: 'Print' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Print unavailable; try again.'))
  expect(print).not.toHaveBeenCalled()
})

it('touch headers and the canvas share the full four-control header height', async () => {
  vi.stubGlobal('matchMedia', () => ({ matches: true, addEventListener() {}, removeEventListener() {} }))
  const ui = render(<InteractiveSheetDoc doc={docOf()} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/))
  const canvas = ui.getByTestId('officeview-sheet-base') as HTMLCanvasElement
  const header = ui.container.querySelector('[data-officeview-ui="grid-headers"]') as HTMLElement
  expect(header.style.height).toBe('176px')
  expect(canvas.style.height).toBe('200px') // 400px pane - 176px header - 24px status line
  expect(canvas.parentElement?.style.top).toBe('176px')
})

it('filter commits use authoritative source values rather than a stale document preview', async () => {
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const doc = docOf(), source = createObjectWorkbookSource(doc, 'authority'), store = source.store('A'), query = store.queryColumn.bind(store)
  store.queryColumn = async (request, opts) => {
    const page = await query({ ...request, encoding: 'entries' }, opts)
    return { ...page, items: page.items.map(entry => entry.kind === 'cell' && entry.cell.value === 'beta' ? { ...entry, cell: { ...entry.cell, value: 'gamma' } } : entry) }
  }
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); go(ui, 'A2')
  fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'gamma' })).toBeTruthy())
  fireEvent.click(ui.getByRole('checkbox', { name: 'gamma' }))
  await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 1 \/ complete body 2/))
})

it('replacing a document/source with reused wire ids cannot retain the old view or selection', FIRST_PAINT_TEST, async () => {
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const first = docOf(), source = createObjectWorkbookSource(first, 'reuse')
  const ui = render(<InteractiveSheetDoc doc={first} source={source} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); go(ui, 'A2')
  const second = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'replacement' }] }]), next = createObjectWorkbookSource(second, 'reuse')
  ui.rerender(<InteractiveSheetDoc doc={second} source={next} />)
  await waitFor(() => expect(ui.getByRole('gridcell', { name: /replacement/ })).toBeTruthy(), FIRST_PAINT_WAIT)
  expect((ui.getByRole('button', { name: 'Copy selection' }) as HTMLButtonElement).disabled).toBe(true)
})


it('keeps only the two most recently visited control ports and disposes evictions', async () => {
  const original = controlsModule.viewerControls, ports: { sheet: string; dispose: ReturnType<typeof vi.fn> }[] = []
  vi.spyOn(controlsModule, 'viewerControls').mockImplementation(async (...args) => {
    const port = await original(...args), dispose = vi.fn(() => port.dispose())
    ports.push({ sheet: args[0].sheetId, dispose })
    return { client: port.client, dispose }
  })
  const doc = workbookOf(['A', 'B', 'C'].map(sheetId => ({ sheetId, cells: [{ row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: sheetId }] })))
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  const visit = async (sheet: string, count: number) => {
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: sheet } })
    await waitFor(() => expect(ports).toHaveLength(count))
  }
  await waitFor(() => expect(ports).toHaveLength(1))
  await visit('B', 2); await visit('A', 2); await visit('C', 3)
  expect(ports[1].dispose).toHaveBeenCalledTimes(1) // A was touched; B is oldest.
  expect(ports[0].dispose).not.toHaveBeenCalled()
  await visit('B', 4)
  expect(ports[0].dispose).toHaveBeenCalledTimes(1)
  ui.unmount()
  for (const port of ports) expect(port.dispose).toHaveBeenCalledTimes(1)
})

it('preserves applied value keys when a sheet control port is evicted and recreated', async () => {
  const doc = docOf()
  doc.sheets.push(...['B', 'C'].map(sheetId => ({ ...doc.sheets[0], sheetId, name: `Sheet ${sheetId}` })))
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)); go(ui, 'A2')
  fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beta' })).toBeTruthy())
  fireEvent.click(ui.getByRole('checkbox', { name: 'beta' }))
  await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 1 \/ complete body 2/))
  for (const sheet of ['B', 'C', 'A']) {
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: sheet } })
    await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/))
    go(ui, sheet === 'A' ? 'A3' : 'A2')
    fireEvent.click(ui.getByRole('button', { name: /Filter column A/ }))
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beta' })).toBeTruthy())
    if (sheet === 'A') {
      expect((ui.getByRole('checkbox', { name: 'beta' }) as HTMLInputElement).checked).toBe(false)
      expect((ui.getByRole('checkbox', { name: 'alpha' }) as HTMLInputElement).checked).toBe(true)
      fireEvent.click(ui.getByRole('checkbox', { name: 'alpha' }))
      await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
      fireEvent.click(ui.getByRole('button', { name: 'OK' }))
      await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Visible body 0 \/ complete body 2/))
    } else fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  }
})


it.each(['Escape', 'replace'])('cancels pending copy synchronously on selection %s', async action => {
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const doc = docOf(), source = createObjectWorkbookSource(doc, 'clear'), store = source.store('A'), write = clipboard()
  const read = store.readCell.bind(store)
  let complete!: () => void, loaded = false, requestSignal!: AbortSignal
  store.readCell = async (at, options) => loaded ? read(at, options) : { kind: 'unknown', pageId: 'p' }
  store.ensureComplete = async (range, options) => {
    requestSignal = options!.signal!
    await new Promise<void>(resolve => { complete = resolve })
    loaded = true
    return { range, stamp: store.stamp, completeness: 'complete', missing: [] }
  }
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/))
  const grid = go(ui, 'A2')
  fireEvent.keyDown(grid, { key: 'c', ctrlKey: true })
  await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Copy is waiting/))
  if (action === 'replace') fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A3' } })
  // Observe cancellation within the input event, before passive effects can run.
  let abortedDuringEvent = false
  const target = action === 'Escape' ? grid : ui.getByRole('button', { name: 'Go To' }).closest('form')!
  const event = action === 'Escape' ? new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }) : new Event('submit', { bubbles: true, cancelable: true })
  const observe = () => { abortedDuringEvent = requestSignal.aborted }
  document.addEventListener(event.type, observe, { once: true })
  act(() => { target.dispatchEvent(event); abortedDuringEvent ||= requestSignal.aborted })
  document.removeEventListener(event.type, observe)
  await act(async () => complete())
  expect(abortedDuringEvent).toBe(true)
  expect(write).not.toHaveBeenCalled()
  expect(ui.getByRole('status').textContent).not.toContain('Selection copied.')
})
