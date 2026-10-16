/**
 * View-only resize admission (spec §5.6, §12.4), codex confirmation r3: a
 * resize builds its candidate geometry off to the side, measures it with the
 * entry's own accounting, admits that against the live ledger while the old
 * geometry is still held, and publishes only if it fits. The peak never
 * passes the hard limit. Also: a structure-only worker layout publication
 * refreshes authored-hidden rows even when the view revision is unchanged.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ReactElement } from 'react'
import { act, cleanup, fireEvent, render, waitFor, type RenderResult } from '@testing-library/react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import type { XlsxCell, XlsxDocument, XlsxRow } from '../src/xlsx/types'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import * as viewportModule from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'

afterEach(() => { cleanup(); vi.restoreAllMocks() })
const BUDGET = 64 * 1024 * 1024
const cell = (row: number, col: number, value: string | number): XlsxCell => ({ ref: `R${row}C${col}`, row, col, value, styleIndex: 0, hasCachedValue: true })
const docOf = (rows: XlsxRow[], sheetId = 's'): XlsxDocument => ({ sheets: [{ name: 'S1', sheetId, rows, cols: [], merges: [], mergeRanges: [] }] })
function frames() {
  const all: SheetFrameStats[] = []
  return { onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
async function settle(): Promise<void> { await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) }) }
const status = (ui: RenderResult) => ui.container.textContent ?? ''
const headerWidth = (ui: RenderResult, letter: string) => parseFloat((ui.getAllByRole('columnheader').find(el => new RegExp(`^Column ${letter}\\b`).test(el.getAttribute('aria-label') ?? '')) as HTMLElement).style.width)
const rowHeight = (ui: RenderResult, label: string) => parseFloat((ui.getByRole('rowheader', { name: label }).parentElement as HTMLElement).style.height)
function goTo(ui: RenderResult, at: string): HTMLElement {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}
const resizeKey = (target: HTMLElement, key: string) => fireEvent.keyDown(target, { key, altKey: true, shiftKey: true })
async function mount(element: (onLedger: (ledger: MainMemoryLedger) => void, onFrame: (s: SheetFrameStats) => void) => ReactElement) {
  let ledger: MainMemoryLedger | undefined
  const f = frames()
  const ui = render(element(value => { ledger = value }, f.onFrame))
  await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  await settle()
  return { ui, f, ledger: () => ledger! }
}
/** Leaves `free` bytes under the hard limit and returns the peak/limit check. */
function squeeze(ledger: MainMemoryLedger, free: number) {
  ledger.set('copy', 'filler', Math.max(0, BUDGET - ledger.bytes - free))
  const peakBefore = ledger.snapshot().peakBytes
  return () => {
    expect(ledger.snapshot().peakBytes).toBeLessThanOrEqual(Math.max(BUDGET, peakBefore))
    expect(ledger.bytes).toBeLessThanOrEqual(BUDGET)
  }
}
async function remoteOf(doc: XlsxDocument) {
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  expect(open.kind).toBe('complete')
  const session = kernel.session('doc')!
  return new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
}
const shell = (doc: XlsxDocument): XlsxDocument => ({ sheets: doc.sheets.map(sheet => ({ ...sheet, rows: [] })) })

describe('confirmation r3 C1: a resize admits its measured candidate before publishing it', () => {
  it('default resident view (no views prop), one cell, 100 KiB free: the view mask is reused, never allocated twice', FIRST_PAINT_TEST, async () => {
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf([{ index: 0, cells: [cell(0, 0, 'a')] }])} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const viewBuffers = ledger().snapshot().byPool.viewBuffers
    const check = squeeze(ledger(), 100 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle()
    check()
    expect(ledger().snapshot().byPool.viewBuffers).toBe(viewBuffers)
    expect(headerWidth(ui, 'A') === 72 || /Column A was not resized: .*memory/.test(status(ui))).toBe(true)
    ledger().set('copy', 'filler', 0)
  })

  it('worker-backed sheet with 2,999 authored row heights, 400 KiB free: the real rebuilt row axis is admitted', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 2999 }, (_, index) => ({ index, customHeight: true, heightPt: 15 + (index % 7), cells: [cell(index, 0, index)] }))
    const doc = docOf(rows)
    const remote = await remoteOf(doc)
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 400 * 1024)
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })

  it('auto-fitting a 400 px first row of a 100-row sheet (200 KiB free) admits the semantic DOM the fit reveals', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 100 }, (_, index) => ({ index, ...(index === 0 ? { customHeight: true, heightPt: 300 } : {}), cells: Array.from({ length: 12 }, (_, col) => cell(index, col, `${index}:${col}`)) }))
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    fireEvent.keyDown(grid, { key: 'F2' })
    const rowSelect = ui.getByRole('button', { name: 'Select row 1' })
    act(() => rowSelect.focus())
    const check = squeeze(ledger(), 200 * 1024)
    resizeKey(rowSelect, 'Enter'); await settle()
    check()
    expect(rowHeight(ui, 'Row 1') < 400 || /Row 1 was not (resized|fitted): .*memory/.test(status(ui))).toBe(true)
    ledger().set('copy', 'filler', 0)
  })

  it('an unrepresentable size (3,000 rows, first two 1e-15 px, 600 KiB free) is refused without publishing a fallback geometry', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 3000 }, (_, index) => ({ index, ...(index < 2 ? { customHeight: true, heightPt: 1e-15 * 72 / 96 } : {}), cells: [cell(index, 0, index)] }))
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const geometryBefore = ledger().snapshot().byPool.geometry
    const check = squeeze(ledger(), 600 * 1024)
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    expect(status(ui)).toMatch(/Row 1 was not resized/)
    expect(ledger().snapshot().byPool.geometry).toBe(geometryBefore)
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(true)
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r4: admission happens before construction', () => {
  it('R1 a candidate whose construction cannot fit is refused before the geometry is built (100k sorted rows, 512 KiB free)', { timeout: 60_000 }, async () => {
    const rows = 100_000
    const doc = docOf(Array.from({ length: rows }, (_, index) => ({ index, cells: [cell(index, 0, index)] })))
    const order = createRowOrder([{ r0: 0, r1: rows - 1, sourceAtOffset: Uint32Array.from({ length: rows }, (_, i) => rows - 1 - i) }])
    const views = { s: { revision: 1, order } }
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={views} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 512 * 1024)
    const builds = vi.spyOn(viewportModule, 'buildSheetGeometry')
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    expect(status(ui)).toMatch(/Row 1 was not resized: .*memory/)
    expect(builds).not.toHaveBeenCalled()
    ledger().set('copy', 'filler', 0)
  })

  it('R3 repaint layout growth is admitted even when the visible counts do not change (2,048-character wrapped cell, 64 KiB free)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'word '.repeat(410).slice(0, 2048)), style: { numFmtId: 0, wrapText: true } }] }])
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 64 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r5: repaint records and reservation transfer', () => {
  it('F2 a visible merge whose anchor is above the pane is admitted (A1:A200, 1,024-line wrapped anchor, scrolled to row 101, 200 KiB free)', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 200 }, (_, index) => ({ index, cells: index === 0 ? [{ ...cell(0, 0, 'a\n'.repeat(1024)), style: { numFmtId: 0, wrapText: true } }] : [] }))
    const doc: XlsxDocument = { sheets: [{ ...docOf(rows).sheets[0], merges: ['A1:A200'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 199, maxCol: 0 }] }] }
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    goTo(ui, 'A1'); await settle()
    const scroller = ui.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 100 * 20; fireEvent.scroll(scroller) }); await settle()
    const grid = ui.getByRole('grid'); act(() => grid.focus())
    const check = squeeze(ledger(), 200 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })

  it('F3 a single layout record larger than the cache budget is admitted at its own construction size (18,000-unit wrapped value)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'a\n'.repeat(9000)), style: { numFmtId: 0, wrapText: true } }] }])
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 8_528_608)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })

  it('F4 the reservation transfer never holds the construction and repaint reservations together', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'z'.repeat(128))] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    const remote = await remoteOf(doc)
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    await act(async () => { await remote.filter(remote.identity('s'), null, [], new AbortController().signal) })
    await settle()
    const grid = goTo(ui, 'A1'); await settle()
    let overlapped = false
    const set = ledger().set.bind(ledger())
    vi.spyOn(ledger(), 'set').mockImplementation((pool, owner, bytes) => {
      set(pool, owner, bytes)
      if (ledger().ownerBytes('geometry', 'resize-candidate') > 0 && ledger().ownerBytes('dom', 'resize-reservation') > 0) overlapped = true
    })
    const check = squeeze(ledger(), 400_000)
    resizeKey(grid, 'ArrowRight'); await settle()
    check()
    expect(headerWidth(ui, 'A')).toBe(72)
    expect(overlapped).toBe(false)
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r6: repaint memory is admitted per record at construction', () => {
  const heavy = (row: number, col: number) => ({ ...cell(row, col, 'a\n'.repeat(9000)), style: { numFmtId: 0, wrapText: true } })
  const layoutMemoryNotice = (ui: RenderResult) => /Not enough memory to lay out this text/.test(status(ui)) || ui.container.querySelectorAll('canvas').length > 0

  it('R1 three adjacent 18,000-unit wrapped cells (22 MiB free): the tile batch never holds more than admitted', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [heavy(0, 0), heavy(0, 1), heavy(0, 2)] }])
    let peakOk = true
    const f = frames()
    let live: MainMemoryLedger | undefined
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value; value.set('copy', 'filler', Math.max(0, BUDGET - value.bytes - 22 * 1024 * 1024)) }} onFrame={stats => { f.onFrame(stats); if (live && live.snapshot().peakBytes > BUDGET) peakOk = false }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A1'); await settle()
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    expect(peakOk).toBe(true)
    expect(live!.snapshot().peakBytes).toBeLessThanOrEqual(BUDGET)
    expect(live!.bytes).toBeLessThanOrEqual(BUDGET)
    expect(layoutMemoryNotice(view)).toBe(true)
  })

  it('R2 detail tiles under a coarse aggregate are bounded too (heavy A2 among 400 dense 8 px rows, 1 MiB free, column A 8 → 16)', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 400 }, (_, index) => ({ index, customHeight: true, heightPt: 6, cells: Array.from({ length: 64 }, (_, col) => index === 1 && col === 0 ? heavy(1, 0) : cell(index, col, col)) }))
    const doc: XlsxDocument = { sheets: [{ ...docOf(rows).sheets[0], cols: [{ min: 0, max: 63, widthChars: 3 / 7 }] }] }
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A1'); await settle()
    const check = squeeze(live!, 1024 * 1024)
    for (let i = 0; i < 1; i++) resizeKey(grid, 'ArrowRight')
    await settle(); await settle()
    check()
    live!.set('copy', 'filler', 0)
  })

  it('R3 formatting during painting is admitted through the allocation observer (2,048-unit text, 200-token @ format, 256 KiB free)', FIRST_PAINT_TEST, async () => {
    const style = { numFmtId: 164, formatCode: `0;0;0;${'@'.repeat(200)}` }
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'q'.repeat(2048)), style }] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A2'); await settle()
    const check = squeeze(live!, 256 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    live!.set('copy', 'filler', 0)
  })

  it('scrolling into a heavy region without any resize stays within the hard limit and degrades to the coarse notice', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = [{ index: 0, cells: [cell(0, 0, 'top')] }, { index: 300, cells: [heavy(300, 0), heavy(300, 1)] }]
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const check = squeeze(live!, 4 * 1024 * 1024)
    const scroller = view.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 300 * 20; fireEvent.scroll(scroller) })
    await settle(); await settle()
    check()
    expect(status(view)).toMatch(/Not enough memory to lay out this text/)
    live!.set('copy', 'filler', 0)
  })
})

describe('find highlights are admitted like tiles', () => {
  it('a find hit on a heavy cell at low free memory stays within the hard limit (whole-cell highlight fallback)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, `needle ${'a\n'.repeat(9000)}`), style: { numFmtId: 0, wrapText: true } }] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const check = squeeze(live!, 2 * 1024 * 1024)
    const paintRecords: number[] = []
    const set = live!.set.bind(live!)
    vi.spyOn(live!, 'set').mockImplementation((pool, owner, bytes) => { set(pool, owner, bytes); if (owner === 'paint-records') paintRecords.push(bytes) })
    fireEvent.change(view.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'needle' } })
    await waitFor(() => expect(status(view)).toMatch(/A1/))
    await settle(); await settle()
    check()
    // Not admitted at 2 MiB: no record was built for the overlay (whole-cell highlight).
    expect(paintRecords).toEqual([])
    // With room, the same highlight is built through the admitted batch.
    live!.set('copy', 'filler', Math.max(0, BUDGET - live!.bytes + live!.ownerBytes('copy', 'filler') - 40 * 1024 * 1024))
    fireEvent.change(view.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'needl' } })
    await waitFor(() => expect(paintRecords.length).toBeGreaterThan(0))
    expect(live!.snapshot().peakBytes).toBeLessThanOrEqual(BUDGET)
    live!.set('copy', 'filler', 0)
  })
})

describe('confirmation r3 C2: structure-only worker layout publications refresh authored-hidden rows', () => {
  it('publishLayout hiding row 1 hides it in the viewer even though the view revision is unchanged', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'a')] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    const remote = await remoteOf(doc)
    const { ui } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    expect(rowHeight(ui, 'Row 1')).toBe(20)
    act(() => remote.publishLayout('s', { ...remote.layout('s'), hidden: [{ r0: 0, r1: 0 }] }))
    await settle()
    await waitFor(() => expect(ui.queryByRole('rowheader', { name: 'Row 1' })).toBeNull())
  })

  it('R4 after a committed worker filter view, a layout publication hiding row 1 still hides it (mask AND NOT authored-hidden)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'a')] }, { index: 1, cells: [cell(1, 0, 'b')] }, { index: 2, cells: [cell(2, 0, 'c')] }])
    const remote = await remoteOf(doc)
    const { ui } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    await act(async () => { await remote.filter(remote.identity('s'), null, [], new AbortController().signal) })
    await settle()
    expect(remote.view('s').visibility).toBeTruthy()
    expect(rowHeight(ui, 'Row 1')).toBe(20)
    act(() => remote.publishLayout('s', { ...remote.layout('s'), hidden: [{ r0: 0, r1: 0 }] }))
    await settle()
    await waitFor(() => expect(ui.queryByRole('rowheader', { name: 'Row 1' })).toBeNull())
    expect(rowHeight(ui, 'Row 2')).toBe(20)
  })
})
