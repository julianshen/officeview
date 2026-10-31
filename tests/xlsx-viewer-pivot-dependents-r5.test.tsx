// Round 5, mounted (codex r5): an over-budget pivot's owed rectangles stay bounded across commits, and a later
// recompute within budget clears the "out of date" marks of the readers it recomputed (only those).
import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { twoPivotSeed } from './helpers/xlsx-pivots-two-pivot-seed'
import { evaluateWorkbookFormulas, noteWorkbookCellsChanged } from '../src/xlsx/formula/workbook'

/** Every rectangle list the viewer hands to a dependent recompute. */
const calls = vi.hoisted(() => [] as Array<readonly unknown[]>)
vi.mock('../src/xlsx/formula/workbook', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/workbook')>()
  return { ...actual, recomputeDependents: (...args: Parameters<typeof actual.recomputeDependents>) => { calls.push(args[1]); return actual.recomputeDependents(...args) } }
})

afterEach(() => { cleanup(); calls.length = 0 })
const pause = (ms: number) => new Promise(resolve => setTimeout(resolve, ms))
const status = (container: HTMLElement) => [...container.querySelectorAll('[role=status]')].map(node => node.textContent ?? '').join(' ')
const showSheet = async (ui: ReturnType<typeof render>, index: string) => {
  const selector = ui.getByRole('combobox', { name: 'Worksheet' }) as HTMLSelectElement
  fireEvent.change(selector, { target: { value: index } })
  await waitFor(() => expect(selector.value).toBe(index))
}
const BIG = 10_000
/** 75k stored cells on Data (rows ≥ BIG) and 40 range readers of each pivot on Formulas: every commit is over the
 * work budget (S3-AE) until the block goes. */
async function mount() {
  const doc = await twoPivotSeed([])
  const formulas = doc.sheets.find(s => s.name === 'Formulas')!, data = doc.sheets.find(s => s.name === 'Data')!
  expect(data.rows.some(row => row.index >= BIG)).toBe(false)
  for (let row = BIG; row < BIG + 25_000; row++) data.rows.push({ index: row, cells: [0, 1, 2].map(col => ({ ref: '', col: 20 + col, row, value: row, styleIndex: 0, hasCachedValue: true })) })
  expect(formulas.rows.some(row => row.index >= 20 && row.index < 60)).toBe(false)
  for (let row = 20; row < 60; row++) formulas.rows.push({ index: row, cells: [
    { ref: `L${row + 1}`, col: 11, row, value: 16655.85, formula: 'SUM(Pivot!B6:B8)', styleIndex: 0, hasCachedValue: true },
    { ref: `M${row + 1}`, col: 12, row, value: 16655.85, formula: 'SUM(Pivot!L6:L8)', styleIndex: 0, hasCachedValue: true },
  ] })
  formulas.rows.sort((a, b) => a.index - b.index)
  evaluateWorkbookFormulas(doc)
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 })
  const goTo = async (ref: string) => {
    fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: ref } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    await pause(300)
  }
  const togglePivot1 = async () => {
    const collapse = ui.queryAllByRole('button', { name: 'Collapse East (Region)' })[0]
    fireEvent.click(collapse ?? ui.getAllByRole('button', { name: 'Expand East (Region)' })[0]!)
    await ui.findAllByRole('button', { name: collapse ? 'Expand East (Region)' : 'Collapse East (Region)' }, { timeout: 10_000 })
    await pause(50)
  }
  const toggleEastInPivot2 = async () => {
    await goTo('K3')
    await waitFor(() => expect(ui.getAllByRole('button', { name: 'Filter and sort Region' }).length).toBeGreaterThan(0), { timeout: 10_000 })
    const buttons = ui.getAllByRole('button', { name: 'Filter and sort Region' })
    fireEvent.click(buttons[buttons.length - 1]!)
    const list = await ui.findByRole('dialog', { name: 'Region' })
    fireEvent.click(within(list).getByRole('checkbox', { name: 'East' }))
    fireEvent.click(within(list).getByRole('button', { name: 'OK' }))
    await pause(500)
    await goTo('A1')
  }
  const dropBlock = () => { data.rows = data.rows.filter(row => row.index < BIG); noteWorkbookCellsChanged(doc) }
  return { doc, ui, togglePivot1, toggleEastInPivot2, dropBlock }
}
const rectKey = (entry: unknown) => JSON.stringify((entry as { rect: unknown }).rect)

it('r5-1: 20 consecutive over-budget commits keep the owed rectangles bounded and without duplicates', async () => {
  const { ui, togglePivot1 } = await mount()
  for (let commit = 0; commit < 20; commit++) await togglePivot1()
  await waitFor(() => expect(calls.length).toBeGreaterThanOrEqual(20), { timeout: 10_000 })
  for (const changed of calls) expect(new Set(changed.map(rectKey)).size).toBe(changed.length)
  // Saved, collapsed and expanded areas of one pivot: never more than those, however many commits were owed.
  expect(Math.max(...calls.map(changed => changed.length))).toBeLessThanOrEqual(3)
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).toMatch(/may be out of date/i), { timeout: 10_000 })
}, 120_000)

it('r5-2: a recompute within budget clears the out-of-date marks of the readers it recomputed, keeping the others', async () => {
  const { ui, togglePivot1, toggleEastInPivot2, dropBlock } = await mount()
  await togglePivot1()       // pivot 1's readers: over budget
  await toggleEastInPivot2() // pivot 2's readers: over budget
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).toMatch(/may be out of date/i), { timeout: 10_000 })
  await showSheet(ui, '2')
  dropBlock()                // now every recompute fits the budget
  await togglePivot1()       // recomputes pivot 1's readers; pivot 2's stay out of date
  await pause(300)
  await showSheet(ui, '3')
  await pause(300)
  expect(status(ui.container)).toMatch(/may be out of date/i)
  await showSheet(ui, '2')
  await toggleEastInPivot2() // recomputes pivot 2's readers
  await showSheet(ui, '3')
  await waitFor(() => expect(status(ui.container)).not.toMatch(/out of date/i), { timeout: 10_000 })
}, 120_000)
