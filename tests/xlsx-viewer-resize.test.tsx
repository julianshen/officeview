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
    cols: '<col min="2" max="2" width="1" customWidth="1"/>',
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
