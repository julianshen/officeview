import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { FindController, FindError } from '../src/xlsx/view/find'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { workbookOf } from './helpers/xlsx-find'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
afterEach(() => { cleanup(); vi.restoreAllMocks() })
const doc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'text' }] }])
const status = (ui: ReturnType<typeof render>) => ui.getByTestId('officeview-search-status').textContent

it('initial real formatter refusal is inline resource failure rather than a false no-matches result', FIRST_PAINT_TEST, async () => {
 let ledger: MainMemoryLedger | undefined
 const ui = render(<InteractiveSheetDoc doc={doc()} onMainLedger={value => { ledger = value }} />)
 await waitFor(() => expect(ui.getByRole('gridcell', { name: /value text/ })).toBeTruthy(), FIRST_PAINT_WAIT)
 ledger!.set('copy', 'peer-pressure', MAIN_THREAD_BUDGETS.D)
 fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'text' } })
 await waitFor(() => expect(status(ui)).not.toBe('indexing…'))
 console.log(JSON.stringify({ initialRefusalStatus: status(ui) }))
 expect(status(ui)).toMatch(/formatting.*memory budget/i)
 expect(status(ui)).not.toMatch(/no matches/i)
 expect((ui.getByRole('button', { name: 'Next match' }) as HTMLButtonElement).disabled).toBe(true)
 ledger!.release('copy', 'peer-pressure')
 fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'tex' } })
 await waitFor(() => expect(status(ui)).toBe('1 of 1'))
})

it('a new query clears a previous initial failure during indexing and recovers normally', FIRST_PAINT_TEST, async () => {
 const spy = vi.spyOn(FindController.prototype, 'run').mockRejectedValueOnce(new FindError('resource-limit', 'peer initial memory refusal'))
 const ui = render(<InteractiveSheetDoc doc={doc()} />)
 await waitFor(() => expect(ui.getByRole('gridcell', { name: /value text/ })).toBeTruthy(), FIRST_PAINT_WAIT)
 const box = ui.getByRole('searchbox', { name: 'Search document' })
 fireEvent.change(box, { target: { value: 'text' } })
 await waitFor(() => expect(ui.container.textContent).toContain('peer initial memory refusal'))
 spy.mockRestore()
 let release!: () => void
 const gate = new Promise<void>(resolve => { release = resolve })
 const run = FindController.prototype.run
 vi.spyOn(FindController.prototype, 'run').mockImplementationOnce(async function (this: FindController, input, query, options) { const result = await run.call(this, input, query, options); await gate; return result })
 try {
  fireEvent.change(box, { target: { value: 'tex' } })
  await waitFor(() => expect(status(ui)).toBe('indexing…'))
  expect(status(ui)).not.toMatch(/refusal|search stopped|no matches/)
  release()
  await waitFor(() => expect(status(ui)).toBe('1 of 1'))
  fireEvent.change(box, { target: { value: '' } })
  await waitFor(() => expect(status(ui)).toBe(''))
 } finally { release() }
})
