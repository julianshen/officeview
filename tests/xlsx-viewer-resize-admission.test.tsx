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
