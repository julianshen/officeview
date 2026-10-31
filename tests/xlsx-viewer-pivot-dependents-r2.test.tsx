// Round 2, mounted: the real viewer wiring after fix round 1.
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { XlsxDocument } from '../src/xlsx/types'
import { patchedSeed } from './xlsx-pivots-dependent-r2-patched'

const value = (doc: XlsxDocument, ref: string) => doc.sheets.find(s => s.name === 'Formulas')!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!.value
const status = (container: HTMLElement) => [...container.querySelectorAll('[role=status]')].map(node => node.textContent ?? '').join(' ')
afterEach(cleanup)

async function hide(doc: XlsxDocument, item: string) {
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  const list = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(list).getByRole('checkbox', { name: item }))
  fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(value(doc, 'B1')).not.toBe(222603.36000000004), { timeout: 10_000 })
  return ui
}

async function showFormulaSheet(ui: ReturnType<typeof render>) {
  const selector = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
  fireEvent.change(selector, { target: { value: '3' } })
  await waitFor(() => expect(selector.value).toBe('3'))
}

it('R2-1 (Critical): a cell on another sheet at coordinates inside the pivot rectangle reads its own sheet, not the pivot overlay', async () => {
  // Formulas!K1 = Pivot!B6 + Data!E6 (Data!E6 is an input cell; (row 5, col 4) is inside the pivot's A3:I24).
  const doc = await patchedSeed([{ row: 1, formula: 'Pivot!B6+Data!E6', cached: 0 }])
  const dataE6 = doc.sheets.find(s => s.name === 'Data')!.rows.flatMap(r => r.cells).find(c => c.ref === 'E6')!.value as number
  await hide(doc, 'East')
  await waitFor(() => expect(value(doc, 'B11')).not.toBe(11204.24), { timeout: 10_000 })
  expect(value(doc, 'K1')).toBeCloseTo((value(doc, 'B11') as number) + dataE6, 6)
}, 30_000)

it('R2-3 (Important): S3-T — a retained (unmeasured) GETPIVOTDATA is marked stale with the notice after a change', async () => {
  const doc = await patchedSeed([{ row: 1, formula: 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Channel","Web")', cached: 5 }])
  const ui = await hide(doc, 'West')
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  expect(value(doc, 'K1')).toBe(5) // retained, as S3-T requires
  await showFormulaSheet(ui)
  await waitFor(() => expect(status(ui.container)).toMatch(/saved values/i)) // ... and flagged stale, as S3-T requires
  fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: '2' } })
  await waitFor(() => expect((ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement).value).toBe('2'))
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  const region = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(region).getByRole('button', { name: 'Reset to saved view' }))
  await waitFor(() => expect(status(ui.container)).not.toMatch(/saved values/i))
}, 30_000)

it('R2-2 mounted (Important): INDIRECT into the pivot neither follows the change nor shows a notice', async () => {
  const doc = await patchedSeed([{ row: 1, formula: 'INDIRECT("Pivot!B6")', cached: 11204.24 }])
  const ui = await hide(doc, 'East')
  await waitFor(() => expect(value(doc, 'B11')).not.toBe(11204.24), { timeout: 10_000 }) // =Pivot!B6 followed
  await showFormulaSheet(ui)
  await waitFor(() => expect(/saved values/i.test(status(ui.container))).toBe(true))
  expect(value(doc, 'K1')).toBe(11204.24)
}, 30_000)
