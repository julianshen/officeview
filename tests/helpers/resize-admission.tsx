/**
 * Shared fixtures of the view-only resize / repaint admission suites
 * (xlsx-viewer-resize-admission and its -heavy companion).
 */
import { expect } from 'vitest'
import type { ReactElement } from 'react'
import { act, fireEvent, render, waitFor, type RenderResult } from '@testing-library/react'
import type { SheetFrameStats } from '../../src/components/xlsx/InteractiveSheetDoc'
import { ViewerClient } from '../../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../../src/worker/viewerKernel'
import { RemoteWorkbook } from '../../src/worker/viewerRemoteSource'
import { createObjectWorkbookSource } from '../../src/worker/viewerObjectSource'
import type { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import type { XlsxCell, XlsxDocument, XlsxRow } from '../../src/xlsx/types'
import { createCellLayoutHost } from '../../src/xlsx/view/renderViewport'
import { estimateLayoutBytes, layoutUpperBytes } from '../../src/xlsx/view/textLayout'
import { FIRST_PAINT_WAIT } from './viewer-timing'

export const BUDGET = 64 * 1024 * 1024
export const cell = (row: number, col: number, value: string | number): XlsxCell => ({ ref: `R${row}C${col}`, row, col, value, styleIndex: 0, hasCachedValue: true })
export const docOf = (rows: XlsxRow[], sheetId = 's'): XlsxDocument => ({ sheets: [{ name: 'S1', sheetId, rows, cols: [], merges: [], mergeRanges: [] }] })
export function frames() {
  const all: SheetFrameStats[] = []
  return { onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}
export async function settle(): Promise<void> { await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) }) }
export const status = (ui: RenderResult) => ui.container.textContent ?? ''
export const headerWidth = (ui: RenderResult, letter: string) => parseFloat((ui.getAllByRole('columnheader').find(el => new RegExp(`^Column ${letter}\\b`).test(el.getAttribute('aria-label') ?? '')) as HTMLElement).style.width)
export const rowHeight = (ui: RenderResult, label: string) => parseFloat((ui.getByRole('rowheader', { name: label }).parentElement as HTMLElement).style.height)
export function goTo(ui: RenderResult, at: string): HTMLElement {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  return grid
}
export const resizeKey = (target: HTMLElement, key: string) => fireEvent.keyDown(target, { key, altKey: true, shiftKey: true })
export async function mount(element: (onLedger: (ledger: MainMemoryLedger) => void, onFrame: (s: SheetFrameStats) => void) => ReactElement) {
  let ledger: MainMemoryLedger | undefined
  const f = frames()
  const ui = render(element(value => { ledger = value }, f.onFrame))
  await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  await settle()
  return { ui, f, ledger: () => ledger! }
}
/** Leaves `free` bytes under the hard limit and returns the peak/limit check. */
export function squeeze(ledger: MainMemoryLedger, free: number) {
  ledger.set('copy', 'filler', Math.max(0, BUDGET - ledger.bytes - free))
  const peakBefore = ledger.snapshot().peakBytes
  return () => {
    expect(ledger.snapshot().peakBytes).toBeLessThanOrEqual(Math.max(BUDGET, peakBefore))
    expect(ledger.bytes).toBeLessThanOrEqual(BUDGET)
  }
}
export async function remoteOf(doc: XlsxDocument) {
  const source = createObjectWorkbookSource(doc, 'doc')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  const open = await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'doc', source: { kind: 'object', document: doc }, profile: 'D' })
  expect(open.kind).toBe('complete')
  const session = kernel.session('doc')!
  return new RemoteWorkbook(client, { docId: 'doc', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(s => [s.sheetId, session.layout(s.sheetId)])) })
}
export const shell = (doc: XlsxDocument): XlsxDocument => ({ sheets: doc.sheets.map(sheet => ({ ...sheet, rows: [] })) })
/** The fewest 'a\n' lines whose wrapped record bound (layoutUpperBytes, two
 * units per line) exceeds `bytes`. Layout time grows with the text, so heavy
 * fixtures are sized from the admission boundary they cross, not a fixed count. */
export const linesOver = (bytes: number) => Math.floor((bytes - layoutUpperBytes(0, 0)) / (layoutUpperBytes(2, 0) - layoutUpperBytes(0, 0))) + 1
/** The fewest 'a\n' lines whose BUILT wrapped record (estimateLayoutBytes in a
 * default 64 x 20 cell) exceeds `bytes`. The estimate is linear in the line
 * count, so it is extrapolated from two small built records; used where a test
 * needs a record the layout cache cannot retain (DEFAULT_LAYOUT_CACHE_BYTES). */
export function linesBuiltOver(bytes: number): number {
  const host = createCellLayoutHost(document.createElement('canvas').getContext('2d')!, { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] })
  const built = (lines: number) => estimateLayoutBytes('', host.layoutBody({ ...cell(0, 0, 'a\n'.repeat(lines)), style: { numFmtId: 0, wrapText: true } }, 64, 20))
  const perLine = (built(200) - built(100)) / 100, base = built(100) - 100 * perLine
  return Math.floor((bytes - base) / perLine) + 1
}
