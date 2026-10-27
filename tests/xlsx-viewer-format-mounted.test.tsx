import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as display from '../src/xlsx/render'
import type { XlsxDocument } from '../src/xlsx/types'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import * as viewport from '../src/xlsx/view/viewport'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MergeQueryCache } from '../src/xlsx/view/viewport'

afterEach(() => { cleanup(); vi.restoreAllMocks() })
const doc = (): XlsxDocument => ({ sheets: [{ name: 'S', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 3, styleIndex: 0, style: { numFmtId: 164, formatCode: '[Red]0.00' } }] }], cols: [], merges: [], mergeRanges: [] }] })

it('mounted resident Find reaches only observed display builders', async () => {
  const calls: boolean[] = [], native = display.cellDisplayText
  vi.spyOn(display, 'cellDisplayText').mockImplementation((...args) => {
    if (new Error().stack?.includes('textOf')) calls.push(!!args[2])
    return native(...args)
  })
  const ui = render(<InteractiveSheetDoc doc={doc()} />)
  fireEvent.change(ui.getByLabelText('Search document'), { target: { value: '3.00' } })
  await waitFor(() => expect(calls.length).toBeGreaterThan(0))
  expect(calls.every(Boolean)).toBe(true)
  await waitFor(() => expect(ui.getByLabelText('Next match').hasAttribute('disabled')).toBe(false))
})

it('does not build sparkline view state for filtered worksheets when the workbook has no sparklines', async () => {
  const document = doc()
  document.sheets = [
    { ...document.sheets[0], name: 'A', sheetId: 'A' },
    { ...document.sheets[0], name: 'B', sheetId: 'B', rows: [{ index: 0, hidden: true, cells: [] }] },
  ]
  const buildHidden = vi.spyOn(viewport, 'sheetBaseHidden')
  const visibility = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: 0 }]))
  const ui = render(<InteractiveSheetDoc doc={document} views={{ B: { revision: 1, visibility } }} />)
  await waitFor(() => expect(ui.getByRole('grid')).toBeTruthy(), FIRST_PAINT_WAIT)
  expect(buildHidden.mock.calls.some(([sheet]) => sheet.name === 'B')).toBe(false)
})

it('uses the latest external source-sheet filter when an inactive sheet owns a sparkline source', async () => {
  const { preloadSparklines, registerSparklinePainter } = await import('../src/xlsx/sparklineLoader')
  const paint = await import('../src/xlsx/sparklines/paint')
  await preloadSparklines()
  const hostCandidates: boolean[] = []
  registerSparklinePainter((_ctx, sheet, _rect, _visible, _scale, candidates) => {
    if (sheet.name === 'Host') {
      hostCandidates.push(candidates?.get(0)?.has(0) ?? false)
    }
    return true
  })
  const sparkline = { host: { row: 0, col: 0 }, dataSheet: 'Source', sourceRows: [0] }
  const document: XlsxDocument = { sheets: [
    { name: 'Host', sheetId: 'A', rows: [{ index: 0, cells: [] }], cols: [], merges: [], mergeRanges: [],
      resolvedSparklines: new Map([[0, new Map([[0, sparkline as never]])]]) },
    { name: 'Source', sheetId: 'B', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 5, styleIndex: 0 }] }], cols: [], merges: [], mergeRanges: [] },
  ] }
  const visible = visibilityFromHidden(createIntervalSet([]))
  const hidden = visibilityFromHidden(createIntervalSet([{ r0: 0, r1: 0 }]))
  const ui = render(<InteractiveSheetDoc doc={document} views={{ B: { revision: 1, visibility: visible } }} />)
  try {
    await waitFor(() => expect(hostCandidates.some(Boolean)).toBe(true), FIRST_PAINT_WAIT)
    const worksheet = ui.getByRole('combobox', { name: 'Worksheet' })
    fireEvent.change(worksheet, { target: { value: 'B' } })
    await waitFor(() => expect((worksheet as HTMLSelectElement).value).toBe('B'), FIRST_PAINT_WAIT)
    fireEvent.change(worksheet, { target: { value: 'A' } })
    await waitFor(() => expect((worksheet as HTMLSelectElement).value).toBe('A'), FIRST_PAINT_WAIT)
    await waitFor(() => expect(hostCandidates.slice(-10).some(Boolean)).toBe(true), FIRST_PAINT_WAIT)
    ui.rerender(<InteractiveSheetDoc doc={document} views={{ B: { revision: 2, visibility: hidden } }} />)
    await waitFor(() => expect(document.sheets[0].diagnostics?.some(item => item.feature === 'sparkline-filter-hidden')).toBe(true), FIRST_PAINT_WAIT)
  } finally {
    cleanup()
    registerSparklinePainter(paint.paintSparklinesIn)
  }
}, 30000)

it('mounted semantic formatting refusal preserves the grid and exposes an explicit state', async () => {
  let ledger: MainMemoryLedger | undefined
  const document = doc(), onLedger = (value: MainMemoryLedger) => { ledger = value }
  const ui = render(<InteractiveSheetDoc doc={document} onMainLedger={onLedger} />)
  await waitFor(() => expect(ledger).toBeDefined())
  ledger!.set('copy', 'pressure-test', MAIN_THREAD_BUDGETS.D)
  document.sheets[0].rows[0].cells[0].value = 4
  ui.rerender(<InteractiveSheetDoc doc={document} onMainLedger={onLedger} showSearch={false} />)
  await waitFor(() => expect(ui.getAllByRole('gridcell').some(cell => cell.textContent?.includes('formatting memory budget'))).toBe(true))
  expect(ui.getAllByRole('gridcell').find(cell => cell.textContent?.includes('formatting memory budget'))?.getAttribute('aria-description')).toContain('error')
  expect(document.sheets[0].rows[0].cells[0].value).toBe(4)
  ledger!.release('copy', 'pressure-test')
  ui.rerender(<InteractiveSheetDoc doc={document} onMainLedger={onLedger} showSearch />)
  await waitFor(() => expect(ui.getAllByRole('gridcell').some(cell => cell.textContent?.includes('value 4.00'))).toBe(true))
  expect(ui.getAllByRole('gridcell').some(cell => cell.textContent?.includes('formatting memory budget'))).toBe(false)
})

it('tight dense/coarse repaint stays inside the hard peak beside a nonempty offscreen merge index', { timeout: 60000 }, async () => {
  const rows = Array.from({ length: 400 }, (_, row) => ({ index: row, customHeight: true, heightPt: 6, cells: Array.from({ length: 64 }, (_, col) => ({ ref: `R${row}C${col}`, row, col, value: row === 1 && col === 0 ? 'a\n'.repeat(9000) : col, styleIndex: 0, hasCachedValue: true, ...(row === 1 && col === 0 ? { style: { numFmtId: 0, wrapText: true } } : {}) })) }))
  const document: XlsxDocument = { sheets: [{ name: 'S', sheetId: 's', rows, cols: [{ min: 0, max: 63, widthChars: 8 / 7 }], merges: ['A1000:B1000'], mergeRanges: [{ minRow: 999, maxRow: 999, minCol: 0, maxCol: 1 }] }] }
  let ledger: MainMemoryLedger | undefined, paints = 0
  const ui = render(<InteractiveSheetDoc doc={document} views={{ s: { revision: 0 } }} onMainLedger={value => { ledger = value }} onFrame={value => { paints = value.basePaints }} />)
  const settle = () => act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
  await waitFor(() => expect(paints).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  await settle()
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A1' } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  await settle()
  const limit = MAIN_THREAD_BUDGETS.D
  ledger!.set('copy', 'filler', Math.max(0, limit - ledger!.bytes - 1024 * 1024))
  const peak = ledger!.snapshot().peakBytes
  fireEvent.keyDown(grid, { key: 'ArrowRight', altKey: true, shiftKey: true })
  await settle(); await settle()
  expect(ledger!.snapshot().peakBytes).toBeLessThanOrEqual(Math.max(limit, peak))
  expect(ledger!.bytes).toBeLessThanOrEqual(limit)
  ledger!.release('copy', 'filler')
})

it('viewer unmount does not refund a merge page still held by a caller', async () => {
  const document = doc()
  document.sheets[0].merges = ['A1:B1']
  document.sheets[0].mergeRanges = [{ minRow: 0, maxRow: 0, minCol: 0, maxCol: 1 }]
  let ledger: MainMemoryLedger | undefined, alias: ReturnType<MergeQueryCache['query']> | undefined
  const native = MergeQueryCache.prototype.query
  vi.spyOn(MergeQueryCache.prototype, 'query').mockImplementation(function(this: MergeQueryCache, ...args) {
    const page = native.apply(this, args)
    if (page.items.length) alias = page
    return page
  })
  const ui = render(<InteractiveSheetDoc doc={document} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ledger?.snapshot().byPool.merges ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
  ui.unmount()
  expect(alias?.items).toHaveLength(1)
  expect(ledger!.snapshot().byPool.merges).toBeGreaterThan(0)
})
