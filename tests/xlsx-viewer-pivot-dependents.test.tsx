// Mounted: the real viewer wiring (InteractiveSheetDoc installPivot) on the GETPIVOTDATA seed.
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'

const SEED = resolve(__dirname, 'fixtures/xlsx-pivots/recipes/getpivotdata-seed.xlsx')
const load = async () => parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(SEED))))
const value = (doc: XlsxDocument, ref: string) => doc.sheets.find(s => s.name === 'Formulas')!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!.value
afterEach(cleanup)

async function hideWest(doc: XlsxDocument) {
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  const list = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(list).getByRole('checkbox', { name: 'West' }))
  fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.queryByRole('button', { name: /West \(Region\)/ })).toBeNull(), { timeout: 10_000 })
  return ui
}

it('mounted control (passes): GETPIVOTDATA on the committed layout updates in the viewer', async () => {
  const doc = await load()
  await hideWest(doc)
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  expect(value(doc, 'B6')).toBe('#REF!')
}, 30_000)

it('F1 mounted (Critical): =SUM(Pivot!B6:B8) follows a collapse in the real viewer (Excel gpd-collapse-item saves 20019.48)', async () => {
  const doc = await load()
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await waitFor(() => expect(value(doc, 'B3')).toBe('#REF!'), { timeout: 10_000 }) // the recompute ran (GETPIVOTDATA)
  expect(value(doc, 'B12')).toBeCloseTo(20019.48, 6)
  expect(value(doc, 'B11')).toBeCloseTo(11204.24, 6)
}, 30_000)

it('F2 mounted (Critical): a same-sheet GETPIVOTDATA (unqualified anchor) keeps evaluating after a collapse', async () => {
  const doc = await load()
  const pivot = doc.sheets.find(s => s.name === 'Pivot')!
  pivot.rows.find(r => r.index === 0)?.cells.push({ ref: 'K1', col: 10, row: 0, value: 222603.36000000004, formula: 'GETPIVOTDATA("Sum of Sales",$A$3)', styleIndex: 0, hasCachedValue: true } as never)
    ?? pivot.rows.unshift({ index: 0, cells: [{ ref: 'K1', col: 10, row: 0, value: 222603.36000000004, formula: 'GETPIVOTDATA("Sum of Sales",$A$3)', styleIndex: 0, hasCachedValue: true } as never] })
  // This test adds a formula to an already evaluated parse. Rebuild the
  // evaluator's runtime index before mounting, as callers must after edits.
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await waitFor(() => expect(value(doc, 'B3')).toBe('#REF!'), { timeout: 10_000 })
  const k1 = pivot.rows.flatMap(r => r.cells).find(c => c.ref === 'K1')!
  expect(k1.value).toBeCloseTo(222603.36, 6) // Excel: the grand total is still visible
}, 30_000)

it('reset control: after reset the GETPIVOTDATA cells read the saved layout again', async () => {
  const doc = await load()
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await waitFor(() => expect(value(doc, 'B3')).toBe('#REF!'), { timeout: 10_000 })
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  const menu = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(menu).getByRole('button', { name: 'Reset to saved view' }))
  await waitFor(() => expect(value(doc, 'B3')).toBeCloseTo(9585.23, 6), { timeout: 10_000 })
  expect(value(doc, 'B9')).toBe('#REF!')
}, 30_000)

it('F12 (Important): unmounting gives the host document its saved formula values back, like its saved pivot definitions', async () => {
  const doc = await load()
  const ui = await hideWest(doc)
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  ui.unmount()
  expect(value(doc, 'B1')).toBeCloseTo(222603.36, 6)
  expect(value(doc, 'B6')).toBeCloseTo(61499.16, 6)
}, 30_000)
