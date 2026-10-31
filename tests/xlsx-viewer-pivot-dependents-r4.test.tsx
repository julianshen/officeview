// Round 4, mounted (R4-1, Ruling S3-AF): resetting one pivot recomputes only its own dependents; the other pivots'
// live dependent values stay, and the document-wide restore runs only when every pivot is back to its saved view.
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { XlsxDocument } from '../src/xlsx/types'
import { twoPivotSeed } from './helpers/xlsx-pivots-two-pivot-seed'
import { chartWorkbook } from './helpers/xlsx-chart-fixture'
import { parseXlsx } from '../src/xlsx/parse'
import { evaluateWorkbookFormulas, MAX_DEPENDENT_CELL_VISITS } from '../src/xlsx/formula/workbook'
import { pointList } from '../src/charts/points'
import type { ChartModel } from '../src/charts/model'

const value = (doc: XlsxDocument, ref: string) => doc.sheets.find(s => s.name === 'Formulas')!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!.value
afterEach(cleanup)
const pause = (ms: number) => new Promise(resolve => setTimeout(resolve, ms))

const status = (container: HTMLElement) => [...container.querySelectorAll('[role=status]')].map(node => node.textContent ?? '').join(' ')
const showSheet = async (ui: ReturnType<typeof render>, index: string) => {
  const selector = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
  fireEvent.change(selector, { target: { value: index } })
  await waitFor(() => expect(selector.value).toBe(index))
}
/** A cached column chart (what Excel writes) over `ref`, as a drawing node to put on the Formulas sheet. */
async function chartNode(ref: string, saved: readonly number[]) {
  const xml = `<chartSpace><chart><plotArea><barChart><barDir val="col"/><grouping val="clustered"/><ser><idx val="0"/><val><numRef><f>${ref}</f><numCache><ptCount val="${saved.length}"/>${saved.map((v, i) => `<pt idx="${i}"><v>${v}</v></pt>`).join('')}</numCache></numRef></val></ser></barChart><catAx/><valAx/></plotArea></chart></chartSpace>`
  return (await parseXlsx(await chartWorkbook(xml))).sheets[0]!.drawings!.find(node => node.content?.kind === 'xlsxChart')!
}
const points = (node: Awaited<ReturnType<typeof chartNode>>) => pointList((node.content as { chart: ChartModel }).chart.groups[0]!.series[0]!.values).map(point => (point as { v?: number }).v)

// K1 = Pivot!L6 (pivot 2's East 2022 Sales, saved 11204.24), K2 = Pivot!B7 (pivot 1's East/Apples 2022, saved 1234.5;
// collapsing East in pivot 1 moves North's subtotal there).
async function mount(extra: Parameters<typeof twoPivotSeed>[0] = [], prepare?: (doc: XlsxDocument) => void) {
  const doc = await twoPivotSeed([{ ref: 'K1', row: 1, formula: 'Pivot!L6', cached: 11204.24 }, { ref: 'K2', row: 2, formula: 'Pivot!B7', cached: 1234.5 }, ...extra])
  prepare?.(doc)
  expect(doc.sheets.find(s => s.name === 'Pivot')!.pivotTables!.length).toBe(2)
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  const goTo = async (ref: string) => {
    fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: ref } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    await pause(300)
  }
  /** Pivot 2 (K3:S24) is off-screen in jsdom: scroll to it, then hide East. Pivot!L6 becomes North 2022 (6497.84). */
  const hideEastInPivot2 = async () => {
    await goTo('K3')
    await waitFor(() => expect(ui.getAllByRole('button', { name: 'Filter and sort Region' }).length).toBeGreaterThan(0), { timeout: 10_000 })
    const buttons = ui.getAllByRole('button', { name: 'Filter and sort Region' })
    fireEvent.click(buttons[buttons.length - 1]!)
    const list = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(list).getByRole('checkbox', { name: 'East' }))
    fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(value(doc, 'K1')).toBeCloseTo(6497.84, 6), { timeout: 10_000 })
  }
  const collapseEastInPivot1 = async () => {
    await goTo('A1')
    fireEvent.click((await ui.findAllByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))[0]!)
    await waitFor(() => expect(value(doc, 'B3')).toBe('#REF!'), { timeout: 10_000 })
  }
  const reset = async (which: 'first' | 'last') => {
    const buttons = ui.getAllByRole('button', { name: 'Filter and sort Region' })
    fireEvent.click(which === 'first' ? buttons[0]! : buttons[buttons.length - 1]!)
    const menu = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(menu).getByRole('button', { name: 'Reset to saved view' }))
  }
  return { doc, ui, goTo, hideEastInPivot2, collapseEastInPivot1, reset }
}

it('R4-1: resetting pivot 1 keeps formulas that read pivot 2 (still filtered) at pivot 2\'s live value', async () => {
  const { doc, goTo, hideEastInPivot2, collapseEastInPivot1, reset } = await mount()
  await hideEastInPivot2()
  await collapseEastInPivot1()
  await reset('first')
  await waitFor(() => expect(value(doc, 'B3')).toBeCloseTo(9585.23, 6), { timeout: 10_000 })
  await pause(500)
  expect(value(doc, 'K1')).toBeCloseTo(6497.84, 6)
  expect(value(doc, 'K2')).toBeCloseTo(1234.5, 6)
  // Pivot 2 back to saved too: every pivot is saved, so the document-wide restore puts every saved value back.
  await goTo('K3')
  await reset('last')
  await waitFor(() => expect(value(doc, 'K1')).toBeCloseTo(11204.24, 6), { timeout: 10_000 })
  expect(value(doc, 'B3')).toBeCloseTo(9585.23, 6)
}, 60_000)

it('R4-1: the reset reads the sheet\'s current composite overlay, whichever pivot committed last', async () => {
  // Pivot 1 commits first, then pivot 2: pivot 2's commit carries the composite that still has pivot 1's live cells.
  const { doc, goTo, hideEastInPivot2, collapseEastInPivot1, reset } = await mount()
  await collapseEastInPivot1()
  await hideEastInPivot2()
  await goTo('A1')
  await reset('first')
  await waitFor(() => expect(value(doc, 'B3')).toBeCloseTo(9585.23, 6), { timeout: 10_000 })
  await pause(500)
  expect(value(doc, 'K2')).toBeCloseTo(1234.5, 6)
  expect(value(doc, 'K1')).toBeCloseTo(6497.84, 6)
}, 60_000)

it('R4-1: a reset puts back only the dependent charts no other live pivot feeds', async () => {
  // Chart A plots K1:K2 (both pivots); chart B plots K2 only (pivot 1).
  const both = await chartNode('Formulas!$K$1:$K$2', [11204.24, 1234.5]), only1 = await chartNode('Formulas!$K$2:$K$2', [1234.5])
  const savedOnly1 = (only1.content as { chart: ChartModel }).chart
  const { doc, hideEastInPivot2, collapseEastInPivot1, reset } = await mount([], doc => { doc.sheets.find(s => s.name === 'Formulas')!.drawings = [both, only1] })
  await hideEastInPivot2()
  await waitFor(() => expect(points(both)[0]).toBeCloseTo(6497.84, 6), { timeout: 10_000 })
  await collapseEastInPivot1()
  await waitFor(() => expect(points(only1)[0]).not.toBeCloseTo(1234.5, 6), { timeout: 10_000 })
  await reset('first')
  await waitFor(() => expect(value(doc, 'K2')).toBeCloseTo(1234.5, 6), { timeout: 10_000 })
  await pause(500)
  // Pivot 1's own chart has its saved model back; the shared chart is refilled with pivot 2 still live.
  expect((only1.content as { chart: ChartModel }).chart).toBe(savedOnly1)
  expect(points(both)[0]).toBeCloseTo(6497.84, 6)
  expect(points(both)[1]).toBeCloseTo(1234.5, 6)
}, 60_000)

it('R4-1: a reset clears only the stale marks its own pivot caused', async () => {
  // Unmeasured GETPIVOTDATA shapes (S3-T): K3 reads pivot 1, K4 reads pivot 2; each is marked stale by its pivot's change.
  const { ui, hideEastInPivot2, collapseEastInPivot1, reset, goTo } = await mount([
    { ref: 'K3', row: 3, formula: 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Channel","Web")', cached: 5 },
    { ref: 'K4', row: 4, formula: 'GETPIVOTDATA("Sum of Sales",Pivot!$K$3,"Channel","Web")', cached: 5 },
  ])
  await collapseEastInPivot1()
  await hideEastInPivot2()
  await goTo('A1')
  await reset('first')
  await pause(500)
  await showSheet(ui, '3')
  // Pivot 2's mark on the Formulas sheet stays while pivot 2 is live.
  await waitFor(() => expect(status(ui.container)).toMatch(/still show saved values/i))
  await showSheet(ui, '2')
  await goTo('K3')
  await reset('last')
  await pause(500)
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).not.toMatch(/saved values|out of date/i))
}, 60_000)

it('R4-1: resetting the pivot that caused the only stale mark clears the notice while the other pivot stays live', async () => {
  const { doc, ui, hideEastInPivot2, collapseEastInPivot1, reset, goTo } = await mount([
    { ref: 'K3', row: 3, formula: 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Channel","Web")', cached: 5 },
  ])
  await collapseEastInPivot1()
  await hideEastInPivot2()
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).toMatch(/still show saved values/i))
  await showSheet(ui, '2')
  await goTo('A1')
  await reset('first')
  await pause(500)
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).not.toMatch(/saved values|out of date/i))
  expect(value(doc, 'K1')).toBeCloseTo(6497.84, 6)
}, 60_000)

it('M-4: readers over the work budget are flagged as possibly out of date, not as saved values', async () => {
  const { ui } = await mount([], doc => {
    // 75k stored cells and 40 range readers of pivot 1: the estimate (40 × the stored map) is over the budget (S3-AE).
    const sheet = doc.sheets.find(s => s.name === 'Formulas')!
    for (let row = 100; row < 25_100; row++) sheet.rows.push({ index: row, cells: [0, 1, 2].map(col => ({ ref: '', col: 20 + col, row, value: row, styleIndex: 0, hasCachedValue: true })) })
    expect(sheet.rows.some(r => r.index >= 20 && r.index < 60)).toBe(false)
    for (let row = 20; row < 60; row++) sheet.rows.push({ index: row, cells: [{ ref: `L${row + 1}`, col: 11, row, value: 16655.85, formula: 'SUM(Pivot!B6:B8)', styleIndex: 0, hasCachedValue: true }] })
    sheet.rows.sort((a, b) => a.index - b.index)
    evaluateWorkbookFormulas(doc)
    expect(40 * 75_000).toBeGreaterThan(MAX_DEPENDENT_CELL_VISITS)
  })
  fireEvent.click((await ui.findAllByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))[0]!)
  await ui.findAllByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await pause(500)
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).toMatch(/may be out of date/i), { timeout: 10_000 })
  expect(status(ui.container)).not.toMatch(/still show saved values/i)
}, 60_000)
