// Ruling S3-AO in the interactive viewer: an overlapping change spills the pivot (#SPILL! at its top-left, announced
// naming the blocking cell); the dependent formulas recompute (a ref into the old area reads empty, the anchor the
// error) and a later change that fits shows the pivot again.
import { readFileSync } from 'node:fs'
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import JSZip from 'jszip'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'

afterEach(() => cleanup())
const FIXTURE = 'tests/fixtures/xlsx-pivots/engine/layout-collapsed-item.xlsx'
/** layout-collapsed-item with `blocker` in A23 (O1's start), K1 = A3 and K2 = B6 (East 2022 Sales, 11204.24). */
async function o1WithFormulas(): Promise<XlsxDocument> {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE))
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', xml
    .replace('<sheetData>', '<sheetData><row r="1"><c r="K1"><f>A3</f><v>0</v></c></row><row r="2"><c r="K2"><f>B6</f><v>11204.24</v></c></row>')
    .replace('</sheetData>', '<row r="23"><c r="A23" t="inlineStr"><is><t>blocker</t></is></c></row></sheetData>')
    .replace(/<dimension ref="[^"]*"\/>/, ''))
  return parseXlsx(await OfficePackage.load(new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }))))
}
const cell = (doc: XlsxDocument, ref: string) => doc.sheets.find(sheet => sheet.pivotTables?.length)!.rows.flatMap(row => row.cells).find(entry => entry.ref === ref)
const status = (container: HTMLElement) => [...container.querySelectorAll('[role=status]')].map(node => node.textContent ?? '').join(' ')

it('S3-AO: expanding into a value spills, announces the blocking cell, recomputes dependents; Reset from the spilled cell shows the pivot again', async () => {
  const doc = await o1WithFormulas()
  expect(cell(doc, 'K2')?.value).toBe(11204.24)
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.click(await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 }))
  await waitFor(() => expect(status(ui.container)).toMatch(/can't fit: #SPILL!\. Cell A23 on Pivot is in the way\./), { timeout: 10_000 })
  // The pivot is one #SPILL! cell; its other cells are empty; the blocker stays.
  await waitFor(() => expect(ui.container.textContent).toMatch(/A3, value #SPILL!/))
  expect(ui.container.textContent).not.toMatch(/B6, value 11204/)
  // Dependents: the anchor reads the error, a ref into the old area reads empty.
  await waitFor(() => expect(cell(doc, 'K1')?.value).toBe('#SPILL!'), { timeout: 10_000 })
  expect([0, null, '']).toContain(cell(doc, 'K2')?.value)
  // The spilled cell's dropdown opens the field menu: Reset shows the saved view again.
  fireEvent.click(await ui.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 }))
  const menu = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(menu).getByRole('button', { name: 'Reset to saved view' }))
  await waitFor(() => expect(cell(doc, 'K2')?.value).toBe(11204.24), { timeout: 10_000 })
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
}, 30_000)
