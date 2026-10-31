// Mounted: the real viewer wiring (InteractiveSheetDoc installPivot) on the GETPIVOTDATA seed.
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'

const SEED = resolve(__dirname, 'fixtures/xlsx-pivots/recipes/getpivotdata-seed.xlsx')
const load = async () => parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(SEED))))
const value = (doc: XlsxDocument, ref: string) => doc.sheets.find(s => s.name === 'Formulas')!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!.value
afterEach(() => { cleanup(); vi.unstubAllGlobals() })

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

// ---- print (Ruling S3-AL, spec §7.10): the pivot prints as saved, and so does everything that reads it -------------
const CHART_ON_B1 = '<chartSpace><chart><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><val><numRef><f>Formulas!$B$1</f><numCache><ptCount val="1"/><pt idx="0"><v>222603.36</v></pt></numCache></numRef></val></ser></barChart><catAx/><valAx/></plotArea></chart></chartSpace>'
/** The seed with a plain (non-pivot) chart on Formulas!B1, the GETPIVOTDATA grand total: Task 31 refills it. */
async function withDependentChart(doc: XlsxDocument) {
  const charted = await parseXlsx(await chartWorkbook(CHART_ON_B1))
  const node = charted.sheets[0]!.drawings!.find(d => d.content?.kind === 'xlsxChart')!
  const formulas = doc.sheets.find(s => s.name === 'Formulas')!
  formulas.drawings = [...(formulas.drawings ?? []), node]
  if (node.content?.kind !== 'xlsxChart' || !node.content.chart) throw new Error('fixture chart missing')
  return { node, saved: node.content.chart }
}
function printing(doc: XlsxDocument) {
  const printed: XlsxDocument[] = []
  const printLoader = async (printDoc: XlsxDocument) => { printed.push(printDoc); return Object.assign([], { dispose() {}, refusedImages: [] }) as never }
  vi.stubGlobal('print', () => {})
  const ui = render(<InteractiveSheetDoc doc={doc} printLoader={printLoader} />)
  const print = async () => {
    const count = printed.length
    fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'p', ctrlKey: true })
    await waitFor(() => expect(printed).toHaveLength(count + 1))
    return printed.at(-1)!
  }
  return { ui, print, note: () => ui.getByTestId('officeview-sheet-print-pages').querySelector('p')?.textContent ?? '' }
}
const filterRegion = async (ui: ReturnType<typeof render>, item: string) => {
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  const list = await ui.findByRole('dialog', { name: 'Region' })
  fireEvent.click(within(list).getByRole('checkbox', { name: item }))
  fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
}

it('a printout after a pivot change shows the saved pivot and the saved values of the formulas and charts that read it', async () => {
  const doc = await load()
  const savedB1 = value(doc, 'B1') as number   // =GETPIVOTDATA grand total, saved 222603.36
  const savedDef = doc.sheets.find(s => s.name === 'Pivot')!.pivotTables![0]!
  const { node, saved: savedChart } = await withDependentChart(doc)
  const { ui, print, note } = printing(doc)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  await filterRegion(ui, 'West')
  // The recompute ran (the live grand total without West) and refilled the chart that reads it.
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  await waitFor(() => expect(node.content?.kind === 'xlsxChart' && node.content.chart).not.toBe(savedChart), { timeout: 10_000 })
  const printed = await print()
  expect(printed.sheets.find(s => s.name === 'Pivot')!.pivotTables![0]).toBe(savedDef)
  expect(value(printed, 'B1')).toBeCloseTo(savedB1, 6)
  expect(value(printed, 'B6')).toBeCloseTo(61499.16, 6)   // West's own GETPIVOTDATA: saved, not #REF!
  const printedChart = printed.sheets.find(s => s.name === 'Formulas')!.drawings!.find(d => d.content?.kind === 'xlsxChart')!
  expect(printedChart.content?.kind === 'xlsxChart' && printedChart.content.chart).toBe(savedChart)
  await waitFor(() => expect(note()).toBe('Printing shows the pivot, and the values that depend on it, as saved.'))
  // The viewer keeps showing the live values: print reads copies.
  expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6)
  expect(node.content?.kind === 'xlsxChart' && node.content.chart).not.toBe(savedChart)
}, 30_000)

it('a printout after every pivot is back to its saved view prints the document as it is, with no note', async () => {
  const doc = await load()
  const { ui, print, note } = printing(doc)
  fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await waitFor(() => expect(value(doc, 'B3')).toBe('#REF!'), { timeout: 10_000 })
  expect(await print()).not.toBe(doc)
  await waitFor(() => expect(note()).toBe('Printing shows the pivot, and the values that depend on it, as saved.'))
  fireEvent.click(ui.getByRole('button', { name: 'Filter and sort Region' }))
  fireEvent.click(within(await ui.findByRole('dialog', { name: 'Region' })).getByRole('button', { name: 'Reset to saved view' }))
  await waitFor(() => expect(value(doc, 'B3')).toBeCloseTo(9585.23, 6), { timeout: 10_000 })
  expect(await print()).toBe(doc)
  expect(note()).toBe('')
}, 30_000)

// ---- the reader sheets' values move with the recompute: value-keyed consumers on them see the live values --------
const pickSheet = (ui: ReturnType<typeof render>, name: string) => {
  const picker = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
  fireEvent.change(picker, { target: { value: [...picker.options].find(option => option.textContent === name)!.value } })
}
const goTo = (ui: ReturnType<typeof render>, at: string) => {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
}

it('a filter list on a reader sheet built before a pivot change lists the recomputed values', async () => {
  const doc = await load()
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  // The Formulas sheet's controls (its value lanes) are built once, before the change.
  pickSheet(ui, 'Formulas'); goTo(ui, 'B1')
  fireEvent.click(await ui.findByRole('button', { name: /Filter column B/ }, { timeout: 10_000 }))
  await ui.findByRole('checkbox', { name: /222,?603\.36/ }, { timeout: 10_000 })
  fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  pickSheet(ui, 'Pivot')
  await filterRegion(await ui.findByRole('button', { name: 'Filter and sort Region' }, { timeout: 10_000 }).then(() => ui), 'West')
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  pickSheet(ui, 'Formulas'); goTo(ui, 'B1')
  fireEvent.click(await ui.findByRole('button', { name: /Filter column B/ }, { timeout: 10_000 }))
  await ui.findByRole('checkbox', { name: /161,?104\.2/ }, { timeout: 10_000 })
  expect(ui.queryByRole('checkbox', { name: /222,?603\.36/ })).toBeNull()
}, 60_000)

it('a search running when a pivot changes finds the recomputed values, as a new search does', async () => {
  const doc = await load()
  const ui = render(<InteractiveSheetDoc doc={doc} showSearch />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  const search = ui.getByRole('searchbox', { name: 'Search document' }) as HTMLInputElement
  const count = () => ui.getByTestId('officeview-search-status').textContent ?? ''
  fireEvent.change(search, { target: { value: '161104.2' } })
  await waitFor(() => expect(count()).toMatch(/no matches/), { timeout: 10_000 })
  await filterRegion(ui, 'West')
  await waitFor(() => expect(value(doc, 'B1')).toBeCloseTo(161104.2, 6), { timeout: 10_000 })
  await waitFor(() => expect(count()).toMatch(/ of /), { timeout: 10_000 })
  const running = count()
  fireEvent.change(search, { target: { value: '' } })
  fireEvent.change(search, { target: { value: '161104.2' } })
  await waitFor(() => expect(count()).toMatch(/ of /), { timeout: 10_000 })
  expect(running).toBe(count())
}, 60_000)
