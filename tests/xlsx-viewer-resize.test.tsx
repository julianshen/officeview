/**
 * View-only row/column resizing in the interactive viewer (spec §5.6
 * amendment): pointer drag on a header edge, double-click auto-fit, the
 * Alt+Shift+Arrow keyboard route announced in the status line, `####` giving
 * way to the value once a column is wide enough, persistence across sort,
 * scroll and sheet switches, Reset sizes, and the worker-backed path where the
 * main-thread planner is authoritative while a sheet is resized.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, type RenderResult } from '@testing-library/react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel } from '../src/worker/viewerKernel'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { XlsxDocument } from '../src/xlsx/types'
import { xlsxBytes } from './helpers/xlsx-package'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import { columnHeaderHeight } from '../src/components/xlsx/GridHeaders'
import { createRowOrder } from '../src/xlsx/view/order'

// Admission's object path allows one row, so the worker fixture below is paged.
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})
afterEach(() => { cleanup(); vi.restoreAllMocks() })

async function workbook(): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(await buildXlsx([{
    name: 'S1',
    // Column B is 1 character wide (12 CSS px): its number paints as '#'.
    cols: '<col min="2" max="2" width="1.7142857" customWidth="1"/>',
    rows: [
      { r: 1, cells: [{ ref: 'A1', v: 'name' }, { ref: 'B1', v: 123456 }, { ref: 'C1', v: 'a considerably longer piece of text' }] },
      { r: 2, cells: [{ ref: 'A2', v: 'gamma' }, { ref: 'B2', v: 2 }] },
      { r: 3, cells: [{ ref: 'A3', v: 'alpha' }, { ref: 'B3', v: 3 }] },
      { r: 4, cells: [{ ref: 'A4', v: 'beta' }, { ref: 'B4', v: 4 }] },
    ],
  }, { name: 'S2', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'other' }] }] }])))
}
function frames() {
  const all: SheetFrameStats[] = []
  return { all, onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
async function settle(): Promise<void> { await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) }) }
async function mount(doc?: XlsxDocument) {
  const f = frames()
  const ui = render(<InteractiveSheetDoc doc={doc ?? await workbook()} onFrame={f.onFrame} />)
  await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  await settle()
  return { ui, f }
}
const headerWidth = (ui: RenderResult, letter: string) => parseFloat((ui.getAllByRole('columnheader').find(el => new RegExp(`^Column ${letter}\\b`).test(el.getAttribute('aria-label') ?? '')) as HTMLElement).style.width)
const rowHeightOf = (ui: RenderResult, value: string) => parseFloat((ui.getByRole('gridcell', { name: new RegExp(`value ${value}(;|$)`) }).parentElement as HTMLElement).style.height)
function goTo(ui: RenderResult, at: string): HTMLElement {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}
const resizeKey = (target: HTMLElement, key: string) => fireEvent.keyDown(target, { key, altKey: true, shiftKey: true })
const status = (ui: RenderResult) => ui.container.textContent ?? ''
function paintedText() {
  const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d')!) as CanvasRenderingContext2D
  const painted: string[] = []
  const fillText = proto.fillText
  vi.spyOn(proto, 'fillText').mockImplementation(function (this: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth?: number) {
    painted.push(String(text))
    return maxWidth === undefined ? fillText.call(this, text, x, y) : fillText.call(this, text, x, y, maxWidth)
  })
  return painted
}

describe('view-only resize: keyboard route and announcements', () => {
  it('Alt+Shift+Arrow resizes the active column and row, announced in the status line', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    expect(headerWidth(ui, 'A')).toBe(64)
    const grid = goTo(ui, 'A2')
    resizeKey(grid, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(72))
    expect(status(ui)).toMatch(/Column A width 72 pixels\. View only; the file is unchanged\./)
    resizeKey(grid, 'ArrowLeft'); resizeKey(grid, 'ArrowLeft')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(56))
    resizeKey(grid, 'ArrowDown')
    await waitFor(() => expect(rowHeightOf(ui, 'gamma')).toBe(28))
    expect(status(ui)).toMatch(/Row 2 height 28 pixels/)
    // Plain arrows still navigate; Alt+Shift does not move the active cell.
    expect(grid.getAttribute('aria-activedescendant')).toMatch(/cell-/)
  })

  it('control mode: Alt+Shift+Arrow on a focused column header resizes that column', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    const grid = goTo(ui, 'C1')
    fireEvent.keyDown(grid, { key: 'F2' })
    const header = ui.container.querySelector<HTMLElement>('[data-header-column="2"]')!
    expect(document.activeElement).toBe(header)
    resizeKey(header, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'C')).toBe(72))
    resizeKey(header, 'Enter')
    await waitFor(() => expect(headerWidth(ui, 'C')).toBeGreaterThan(120))
    expect(status(ui)).toMatch(/Column C width \d+ pixels, fitted to content/)
  })
})

describe('view-only resize: pointer', () => {
  it('dragging a column header edge shows a guide line and applies once on release; #### becomes the value', FIRST_PAINT_TEST, async () => {
    const painted = paintedText()
    const { ui, f } = await mount()
    expect(painted.some(text => /^#+$/.test(text))).toBe(true)
    expect(painted).not.toContain('123456')
    expect(headerWidth(ui, 'B')).toBe(12)
    const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="1"]')!
    expect(handle.getAttribute('aria-hidden')).toBe('true')
    expect(handle.style.cursor).toBe('col-resize')
    const paints = f.last().basePaints
    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, clientX: 100, clientY: 10 })
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 160, clientY: 10 })
    // Live guide only: the geometry is unchanged until release.
    expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')?.textContent).toBe('Width: 72 px')
    expect(headerWidth(ui, 'B')).toBe(12)
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 200, clientY: 10 })
    painted.length = 0
    fireEvent.pointerUp(window, { pointerId: 1, clientX: 200, clientY: 10 })
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(112))
    expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')).toBeNull()
    await waitFor(() => expect(f.last().basePaints).toBeGreaterThan(paints))
    await waitFor(() => expect(painted).toContain('123456'))
    // Hit-testing follows the new geometry: a click inside the widened B selects B1.
    const grid = ui.getByRole('grid'), bounds = grid.getBoundingClientRect()
    fireEvent.click(grid, { clientX: bounds.left + 44 + 64 + 100, clientY: bounds.top + 44 + 5, detail: 1 })
    await waitFor(() => expect(grid.getAttribute('aria-activedescendant')).toMatch(/cell-.*$/))
    expect(status(ui)).toMatch(/B1, value 123456/)
  })

  it('a click or double-click without movement never resizes; double-click auto-fits to the widest value', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="2"]')!
    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, clientX: 100 })
    fireEvent.pointerUp(window, { pointerId: 1, clientX: 101 })
    expect(headerWidth(ui, 'C')).toBe(64)
    fireEvent.doubleClick(handle)
    await waitFor(() => expect(headerWidth(ui, 'C')).toBeGreaterThan(150))
    expect(status(ui)).toMatch(/Column C width \d+ pixels, fitted to content/)
    // A row edge: dragging it applies a height, double-clicking fits the content.
    const rowHandle = ui.container.querySelector<HTMLElement>('[data-resize-row="1"]')!
    expect(rowHandle.style.cursor).toBe('row-resize')
    fireEvent.pointerDown(rowHandle, { button: 0, pointerId: 2, clientY: 50 })
    fireEvent.pointerMove(window, { pointerId: 2, clientY: 80 })
    fireEvent.pointerUp(window, { pointerId: 2, clientY: 80 })
    await waitFor(() => expect(rowHeightOf(ui, 'gamma')).toBe(50))
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-row="1"]')!)
    await waitFor(() => expect(rowHeightOf(ui, 'gamma')).toBe(20))
  })
})

describe('view-only resize: coarse pointers', () => {
  it('shows one 44px handle for the active column and row only, clear of the sort/filter buttons and of every row header', FIRST_PAINT_TEST, async () => {
    const previous = window.matchMedia
    window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null,
      addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true })) as typeof window.matchMedia
    try {
      const { ui } = await mount()
      expect(ui.container.querySelectorAll('[data-officeview-ui="resize-handle"]')).toHaveLength(0)
      goTo(ui, 'C2')
      await waitFor(() => expect(ui.container.querySelectorAll('[data-officeview-ui="resize-handle"]')).toHaveLength(2))
      const column = ui.container.querySelector<HTMLElement>('[data-resize-column="2"]')!
      const row = ui.container.querySelector<HTMLElement>('[data-resize-row="1"]')!
      for (const handle of [column, row]) {
        expect(parseFloat(handle.style.width)).toBeGreaterThanOrEqual(44)
        expect(parseFloat(handle.style.height)).toBeGreaterThanOrEqual(44)
        expect(handle.style.touchAction).toBe('none')
      }
      // The column handle stays in the top (select) lane; the stacked sort and
      // filter buttons below it keep their whole targets.
      expect(parseFloat(column.style.top) + parseFloat(column.style.height)).toBeLessThanOrEqual(44)
      // The row handle lives in the header corner's free slot, above the first
      // row header: it never takes a tap that belongs to a row (§9.3).
      const rowY = Number(/translate\([^,]+,\s*([-\d.]+)px\)/.exec(row.style.transform)?.[1])
      expect(rowY).toBe(44)
      expect(rowY + parseFloat(row.style.height)).toBe(88)
      expect(rowY + parseFloat(row.style.height)).toBeLessThanOrEqual(columnHeaderHeight(true, true, true))
      // The column target is clear of the worksheet corner (x >= 44).
      expect(Number(/translate\(([-\d.]+)px/.exec(column.style.transform)?.[1])).toBeGreaterThanOrEqual(44)
      expect(row.textContent).toBe('↕2')
      fireEvent.pointerDown(row, { button: 0, pointerId: 4, clientY: 60, pointerType: 'touch' })
      fireEvent.pointerMove(window, { pointerId: 4, clientY: 76, pointerType: 'touch' })
      fireEvent.pointerUp(window, { pointerId: 4, clientY: 76, pointerType: 'touch' })
      await waitFor(() => expect(parseFloat((ui.getByRole('rowheader', { name: 'Row 2' }).parentElement as HTMLElement).style.height)).toBe(36))
      fireEvent.pointerDown(column, { button: 0, pointerId: 3, clientX: 300, pointerType: 'touch' })
      fireEvent.pointerMove(window, { pointerId: 3, clientX: 340, pointerType: 'touch' })
      fireEvent.pointerUp(window, { pointerId: 3, clientX: 340, pointerType: 'touch' })
      await waitFor(() => expect(headerWidth(ui, 'C')).toBe(104))
    } finally { window.matchMedia = previous }
  })
})

describe('view-only resize: session persistence and Reset sizes', () => {
  it('sizes survive sort (rows by source row), scroll and sheet switches; Reset restores the workbook sizes', FIRST_PAINT_TEST, async () => {
    const { ui, f } = await mount()
    const grid = goTo(ui, 'B3')
    for (let i = 0; i < 5; i++) resizeKey(grid, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(52))
    resizeKey(grid, 'ArrowDown'); resizeKey(grid, 'ArrowDown')
    await waitFor(() => expect(rowHeightOf(ui, '3')).toBe(36))
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(false)
    // Sort by column A: 'alpha' (source row 3) moves; its height goes with it.
    goTo(ui, 'A1')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
    await waitFor(() => expect(status(ui)).toMatch(/View-only sort applied/))
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value 3; original B3/ })).toBeTruthy())
    expect(rowHeightOf(ui, '3')).toBe(36)
    expect(rowHeightOf(ui, '2')).toBe(20)
    expect(headerWidth(ui, 'B')).toBe(52)
    // Scroll away and back.
    const scroller = ui.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 4000; fireEvent.scroll(scroller) })
    await settle()
    act(() => { scroller.scrollTop = 0; fireEvent.scroll(scroller) })
    await settle()
    expect(headerWidth(ui, 'B')).toBe(52)
    // Another sheet has its own (workbook) sizes; returning keeps this sheet's.
    const select = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
    const [first, second] = [...select.options].map(option => option.value)
    fireEvent.change(select, { target: { value: second } })
    await waitFor(() => expect(f.last().activeSheetId).toBe(second))
    expect(headerWidth(ui, 'B')).toBe(64)
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(true)
    fireEvent.change(select, { target: { value: first } })
    await waitFor(() => expect(f.last().activeSheetId).toBe(first))
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(52))
    fireEvent.click(ui.getByRole('button', { name: 'Reset sizes' }))
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(12))
    expect(rowHeightOf(ui, '3')).toBe(20)
    expect(status(ui)).toMatch(/restored to the workbook sizes/)
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(true)
  })

  it('charges the session sizes to the main ledger and releases them on Reset', FIRST_PAINT_TEST, async () => {
    const ledgers: Array<{ ownerBytes(pool: 'layout', owner: string): number }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={await workbook()} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const sizes = () => ledgers[ledgers.length - 1].ownerBytes('layout', 'resize-sizes')
    expect(sizes()).toBe(0)
    resizeKey(goTo(ui, 'A1'), 'ArrowRight')
    await waitFor(() => expect(sizes()).toBeGreaterThan(0))
    fireEvent.click(ui.getByRole('button', { name: 'Reset sizes' }))
    await waitFor(() => expect(sizes()).toBe(0))
  })
})

/** A port with dedicated-worker semantics over a real kernel (structured-cloned messages). */
function workerHarness() {
  const parse = vi.fn(async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)))
  const kernel = new ViewerKernel({ parse }, { yield: async () => {} })
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
  return { port, posted, viewports: () => posted.filter(command => command.kind === 'viewport').length }
}
const pagedRows = (rows: number[]) => rows.map(r => `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>v${r}</t></is></c><c r="B${r}"><v>${r * 1000}</v></c></row>`).join('')

describe('view-only resize on a worker-backed sheet', () => {
  it('main-thread plans are authoritative while resized (no worker viewport commands); Reset resumes worker plans', FIRST_PAINT_TEST, async () => {
    const h = workerHarness()
    const f = frames()
    const data = new Blob([xlsxBytes({ sheetData: pagedRows([1, 2, 3, 5000, 5001, 9000, 9001, 20000]) }) as BlobPart])
    const ui = render(<OfficeFile data={data} interactiveSheets viewerWorker={() => h.port} {...({ onFrame: f.onFrame } as object)} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /v1/ })).toBeTruthy(), FIRST_PAINT_WAIT)
    await waitFor(() => expect(h.viewports()).toBeGreaterThan(0))
    const grid = goTo(ui, 'B1')
    resizeKey(grid, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(72))
    await settle()
    // A cold region while resized: rows are read through the row cache, the worker plans nothing.
    const before = h.viewports()
    goTo(ui, 'B5000')
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /B5000, value 5000000/ })).toBeTruthy())
    await settle()
    expect(h.viewports()).toBe(before)
    expect(headerWidth(ui, 'B')).toBe(72)
    // Reset: the workbook geometry is the worker's again, so cold tiles ask it.
    fireEvent.click(ui.getByRole('button', { name: 'Reset sizes' }))
    await waitFor(() => expect(headerWidth(ui, 'B')).toBe(64))
    goTo(ui, 'B20000')
    await waitFor(() => expect(h.viewports()).toBeGreaterThan(before))
  })
})

describe('codex review round 1', () => {
  const sheetIdsOf = (ui: RenderResult) => [...(ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement).options].map(option => option.value)

  it('#1 auto-fit measures in a per-fit scratch that is charged while it runs and released after', FIRST_PAINT_TEST, async () => {
    const ledgers: Array<{ ownerBytes(pool: 'layout', owner: string): number; set(pool: 'layout', owner: string, bytes: number): void }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={await workbook()} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger as never) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const ledger = ledgers[ledgers.length - 1]
    const seen: number[] = []
    const set = ledger.set.bind(ledger)
    vi.spyOn(ledger, 'set').mockImplementation((pool, owner, bytes) => { if (owner === 'resize-measure') seen.push(bytes); set(pool, owner, bytes) })
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-column="2"]')!)
    await waitFor(() => expect(headerWidth(ui, 'C')).toBeGreaterThan(150))
    expect(Math.max(...seen)).toBeGreaterThan(0)
    expect(ledger.ownerBytes('layout', 'resize-measure')).toBe(0)
  })

  it('C2 a resize whose drawing replacements do not fit the main budget is refused, restored and announced', FIRST_PAINT_TEST, async () => {
    const doc = await workbook()
    const E = 9525
    doc.sheets[0].drawings = Array.from({ length: 64 }, (_, i) => ({ xEmu: (100 + i) * E, yEmu: 0, widthEmu: 64 * E, heightEmu: 40 * E, geometry: 'rect', fill: '#ff0000',
      source: { partPath: 'xl/drawings/drawing1.xml', treePath: `twoCellAnchor[${i}]/sp[0]`, element: 'sp', representation: 'native' } })) as never
    const ledgers: Array<{ bytes: number; peakBytes?: number; set(pool: 'copy', owner: string, bytes: number): void; snapshot(): { peakBytes: number } }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger as never) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const ledger = ledgers[ledgers.length - 1]
    const budget = 64 * 1024 * 1024
    const grid = goTo(ui, 'A1')
    await settle()
    const drawingsBefore = (ledger as unknown as { snapshot(): { byPool: Record<string, number> } }).snapshot().byPool.drawings
    // Leave less room than 64 repositioned copies need.
    ledger.set('copy', 'filler', budget - ledger.bytes - 8192)
    resizeKey(grid, 'ArrowRight')
    await waitFor(() => expect(status(ui)).toMatch(/Column A was not resized: .*memory/))
    expect(headerWidth(ui, 'A')).toBe(64)
    expect(ledger.bytes).toBeLessThanOrEqual(budget)
    expect((ledger as unknown as { snapshot(): { byPool: Record<string, number> } }).snapshot().byPool.drawings).toBe(drawingsBefore)
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(true)
    ledger.set('copy', 'filler', 0)
  })

  it('C3 keyboard resizing across the 44px lane threshold keeps focus on an equivalent header control', FIRST_PAINT_TEST, async () => {
    const previous = window.matchMedia
    window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null,
      addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true })) as typeof window.matchMedia
    try {
      const { ui } = await mount()
      const grid = goTo(ui, 'B1')
      fireEvent.keyDown(grid, { key: 'F2' })
      const lane = await waitFor(() => ui.getByRole('button', { name: 'Actions for column B' }))
      act(() => lane.focus())
      for (let i = 0; i < 4; i++) resizeKey(document.activeElement as HTMLElement, 'ArrowRight')
      await waitFor(() => expect(headerWidth(ui, 'B')).toBe(44))
      expect(document.activeElement).not.toBe(document.body)
      expect(grid.contains(document.activeElement)).toBe(true)
      resizeKey(document.activeElement as HTMLElement, 'ArrowRight')
      await waitFor(() => expect(headerWidth(ui, 'B')).toBe(52))
      // And back below the threshold from an ordinary action button.
      act(() => ui.getByRole('button', { name: 'Cycle sort column B' }).focus())
      for (let i = 0; i < 2; i++) resizeKey(document.activeElement as HTMLElement, 'ArrowLeft')
      await waitFor(() => expect(headerWidth(ui, 'B')).toBe(36))
      expect(grid.contains(document.activeElement)).toBe(true)
      resizeKey(document.activeElement as HTMLElement, 'ArrowLeft')
      await waitFor(() => expect(headerWidth(ui, 'B')).toBe(28))
    } finally { window.matchMedia = previous }
  })

  it('C4 row auto-fit of rotated wrapped text fits its rotated extent', FIRST_PAINT_TEST, async () => {
    const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'abcdefghijklmno' }] }] }])))
    const cell = doc.sheets[0].rows.flatMap(row => row.cells).find(c => c.ref === 'A1')!
    cell.style = { ...cell.style, wrapText: true, fontSizePt: 10, textRotation: 90 } as NonNullable<typeof cell.style>
    const { ui } = await mount(doc)
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-row="0"]')!)
    await waitFor(() => expect(parseFloat((ui.getByRole('rowheader', { name: 'Row 1' }).parentElement as HTMLElement).style.height)).toBeGreaterThan(60))
  })

  it('Excel parity: column auto-fit does not widen a wrapped cell to its unwrapped width (longest word only)', FIRST_PAINT_TEST, async () => {
    const sentence = 'Wrapped text in a cell that should break across several lines'
    const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: sentence }] }] }])))
    const cell = doc.sheets[0].rows.flatMap(row => row.cells).find(c => c.ref === 'A1')!
    cell.style = { ...cell.style, wrapText: true } as NonNullable<typeof cell.style>
    const { ui } = await mount(doc)
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!)
    await waitFor(() => expect(status(ui)).toMatch(/Column A width \d+ pixels, fitted to content/))
    // 'several' is the longest word: far narrower than the one-line sentence.
    expect(headerWidth(ui, 'A')).toBeLessThan(100)
  })

  it('R2a a resize whose rebuilt view owners do not fit is refused (32 grouped drawings, 20 KiB free)', FIRST_PAINT_TEST, async () => {
    const doc = await workbook()
    const E = 9525
    doc.sheets[0].drawings = Array.from({ length: 32 }, (_, i) => ({ xEmu: (100 + i) * E, yEmu: 0, widthEmu: 64 * E, heightEmu: 40 * E,
      group: { off: { x: (100 + i) * E, y: 0 }, ext: { width: 64 * E, height: 40 * E }, chOff: { x: 0, y: 0 }, chExt: { width: 64 * E, height: 40 * E } },
      children: [{ xEmu: 0, yEmu: 0, widthEmu: 64 * E, heightEmu: 40 * E, geometry: 'rect', fill: '#00f', source: { partPath: 'p', treePath: `twoCellAnchor[${i}]/grpSp[0]/sp[0]`, element: 'sp', representation: 'native' } }],
      source: { partPath: 'xl/drawings/drawing1.xml', treePath: `twoCellAnchor[${i}]/grpSp[0]`, element: 'grpSp', representation: 'native' } })) as never
    const ledgers: Array<{ bytes: number; set(pool: 'copy', owner: string, bytes: number): void; snapshot(): { peakBytes: number } }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger as never) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const ledger = ledgers[ledgers.length - 1], budget = 64 * 1024 * 1024
    const grid = goTo(ui, 'A1')
    await settle()
    ledger.set('copy', 'filler', budget - ledger.bytes - 20 * 1024)
    const peakBefore = ledger.snapshot().peakBytes
    resizeKey(grid, 'ArrowRight')
    await settle()
    const resized = headerWidth(ui, 'A') === 72
    if (resized) expect(ledger.snapshot().peakBytes).toBeLessThanOrEqual(budget)
    else expect(status(ui)).toMatch(/Column A was not resized: .*memory/)
    expect(ledger.snapshot().peakBytes - Math.max(peakBefore, 0)).toBeLessThanOrEqual(budget - peakBefore + 1)
    expect(ledger.bytes).toBeLessThanOrEqual(budget)
    ledger.set('copy', 'filler', 0)
  })

  it('R2b prospective projection growth is admitted: a sorted uniform sheet that would become dense is refused', { timeout: 60_000 }, async () => {
    const rows = 100_000
    const doc: XlsxDocument = { sheets: [{ name: 'S1', sheetId: '1', rows: Array.from({ length: rows }, (_, index) => ({ index, cells: [{ ref: `A${index + 1}`, row: index, col: 0, value: index, styleIndex: 0, hasCachedValue: true }] })), cols: [], merges: [], mergeRanges: [] }] }
    const order = createRowOrder([{ r0: 0, r1: rows - 1, sourceAtOffset: Uint32Array.from({ length: rows }, (_, i) => rows - 1 - i) }])
    const ledgers: Array<{ bytes: number; set(pool: 'copy', owner: string, bytes: number): void; snapshot(): { peakBytes: number } }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={doc} views={{ '1': { revision: 1, order } }} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger as never) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const ledger = ledgers[ledgers.length - 1], budget = 64 * 1024 * 1024
    const grid = goTo(ui, 'A1')
    await settle()
    // Room for a rebuilt cell index (136 B per row + 8 B per cell) plus 1 MiB:
    // not for the 20 B/row dense band the resized row turns the sorted band into.
    ledger.set('copy', 'filler', budget - ledger.bytes - rows * (136 + 8) - 1024 * 1024)
    const peakBefore = ledger.snapshot().peakBytes
    resizeKey(grid, 'ArrowDown')
    await settle()
    expect(status(ui)).toMatch(/Row 1 was not resized: .*memory/)
    expect(ledger.snapshot().peakBytes).toBe(peakBefore)
    expect(ledger.bytes).toBeLessThanOrEqual(budget)
    ledger.set('copy', 'filler', 0)
  })

  it('R3 a completed resize never steals focus the user later moves outside the grid', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    const grid = goTo(ui, 'A1')
    fireEvent.keyDown(grid, { key: 'F2' })
    const sort = ui.getByRole('button', { name: 'Cycle sort column A' })
    act(() => sort.focus())
    resizeKey(sort, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(72))
    expect(document.activeElement).toBe(sort)
    const outside = ui.getByRole('button', { name: 'Print' })
    act(() => outside.focus())
    await settle()
    act(() => { fireEvent.scroll(ui.getByTestId('officeview-sheet-scroller')) })
    await settle()
    expect(document.activeElement).toBe(outside)
  })

  it('R4 oblique rotated text fits the canonical measured glyph band', FIRST_PAINT_TEST, async () => {
    const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S1', cols: '<col min="1" max="1" width="57.7142857" customWidth="1"/>', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'ab' }] }] }])))
    const cell = doc.sheets[0].rows.flatMap(row => row.cells).find(c => c.ref === 'A1')!
    cell.style = { ...cell.style, wrapText: true, fontSizePt: 120, textRotation: 1 } as NonNullable<typeof cell.style>
    const { ui } = await mount(doc)
    expect(headerWidth(ui, 'A')).toBe(404)
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-row="0"]')!)
    // 120 pt: canonical ascent >= .85em and descent >= .25em of 160 px, plus padding.
    await waitFor(() => expect(parseFloat((ui.getByRole('rowheader', { name: 'Row 1' }).parentElement as HTMLElement).style.height)).toBeGreaterThanOrEqual(180))
  })

  it('#3 repositioned resident drawings are charged to the drawings pool', FIRST_PAINT_TEST, async () => {
    const doc = await workbook()
    const E = 9525
    doc.sheets[0].drawings = [{ xEmu: 200 * E, yEmu: 0, widthEmu: 64 * E, heightEmu: 40 * E, geometry: 'rect', fill: '#ff0000',
      source: { partPath: 'xl/drawings/drawing1.xml', treePath: 'twoCellAnchor[0]/sp[0]', element: 'sp', representation: 'native' } } as never]
    const ledgers: Array<{ snapshot(): { byPool: Record<string, number> } }> = []
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={f.onFrame} onMainLedger={ledger => { ledgers.push(ledger) }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    const drawings = () => ledgers[ledgers.length - 1].snapshot().byPool.drawings
    const before = drawings()
    resizeKey(goTo(ui, 'A1'), 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(72))
    await waitFor(() => expect(drawings()).toBeGreaterThan(before))
  })

  it('#4 a drag in progress is cancelled by a sheet switch and never resizes the previous sheet', FIRST_PAINT_TEST, async () => {
    const { ui, f } = await mount()
    const [first, second] = sheetIdsOf(ui)
    fireEvent.pointerDown(ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!, { button: 0, pointerId: 1, clientX: 100 })
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 120 })
    const select = ui.getByRole('combobox', { name: 'Worksheet' })
    fireEvent.change(select, { target: { value: second } })
    await waitFor(() => expect(f.last().activeSheetId).toBe(second))
    fireEvent.pointerUp(window, { pointerId: 1, clientX: 140 })
    await settle()
    expect(headerWidth(ui, 'A')).toBe(64)
    fireEvent.change(select, { target: { value: first } })
    await waitFor(() => expect(f.last().activeSheetId).toBe(first))
    expect(headerWidth(ui, 'A')).toBe(64)
  })

  it('#6 a stale external view revision does not block a later resize', FIRST_PAINT_TEST, async () => {
    const doc = await workbook()
    const f = frames()
    const sheetId = doc.sheets[0].sheetId ?? '1'
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={f.onFrame} views={{ [sheetId]: { revision: 3 } }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    ui.rerender(<InteractiveSheetDoc doc={doc} onFrame={f.onFrame} views={{ [sheetId]: { revision: 2 } }} />)
    await settle()
    resizeKey(goTo(ui, 'A1'), 'ArrowRight')
    await waitFor(() => expect(status(ui)).toMatch(/Column A width 72 pixels/))
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(72))
    expect(f.last().identity.viewGen).toBe(3)
  })

  it('#7 row auto-fit uses the painter wrapping: long unbroken words wrap, so the fitted row holds every line', FIRST_PAINT_TEST, async () => {
    const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'x'.repeat(78) }] }] }])))
    const cell = doc.sheets[0].rows.flatMap(row => row.cells).find(c => c.ref === 'A1')!
    cell.style = { ...cell.style, wrapText: true, fontSizePt: 10 } as NonNullable<typeof cell.style>
    const { ui } = await mount(doc)
    fireEvent.doubleClick(ui.container.querySelector<HTMLElement>('[data-resize-row="0"]')!)
    await waitFor(() => expect(parseFloat((ui.getByRole('rowheader', { name: 'Row 1' }).parentElement as HTMLElement).style.height)).toBeGreaterThan(120))
  })

  it('#8 only the pointer that started the drag can move or end it', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    fireEvent.pointerDown(ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!, { button: 0, pointerId: 1, clientX: 100 })
    fireEvent.pointerMove(window, { pointerId: 2, clientX: 200 })
    fireEvent.pointerUp(window, { pointerId: 2, clientX: 200 })
    await settle()
    expect(headerWidth(ui, 'A')).toBe(64)
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 116 })
    fireEvent.pointerUp(window, { pointerId: 1, clientX: 116 })
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(80))
  })

  it('#10 the click slop is in CSS pixels at any zoom', FIRST_PAINT_TEST, async () => {
    const { ui, f } = await mount()
    for (let i = 0; i < 8 && f.last().zoom > 0.5; i++) { fireEvent.click(ui.getByRole('button', { name: 'Zoom out' })); await settle() }
    expect(f.last().zoom).toBeLessThanOrEqual(0.5)
    const zoom = f.last().zoom
    const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!
    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, clientX: 100 })
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 101 })
    fireEvent.pointerUp(window, { pointerId: 1, clientX: 101 })
    await settle()
    expect(headerWidth(ui, 'A')).toBeCloseTo(64 * zoom, 5)
  })

  it('#11 control-mode resize keys also work from a focused sort/filter/more action button', FIRST_PAINT_TEST, async () => {
    const { ui } = await mount()
    const grid = goTo(ui, 'A1')
    fireEvent.keyDown(grid, { key: 'F2' })
    const sort = ui.getByRole('button', { name: 'Cycle sort column A' })
    act(() => sort.focus())
    resizeKey(sort, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(72))
    const filter = ui.getByRole('button', { name: /^Filter column A/ })
    act(() => filter.focus())
    resizeKey(filter, 'ArrowRight')
    await waitFor(() => expect(headerWidth(ui, 'A')).toBe(80))
  })
})
