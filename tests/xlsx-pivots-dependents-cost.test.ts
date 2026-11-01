// What a pivot commit's dependent recompute costs (Task 31; rulings S3-AE, S3-AJ): its work follows the readers of the
// changed cells, never the workbook. Bounds are counted (cells touched, passes over a sheet's rows), not timed.
import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, noteWorkbookCellsChanged, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../src/xlsx/types'

const pivotSheet = (rows = 1): XlsxSheet => ({ name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [],
  rows: Array.from({ length: rows }, (_, r) => ({ index: r, cells: [{ ref: `A${r + 1}`, col: 0, row: r, value: r + 1, styleIndex: 0, hasCachedValue: true }] })) })
const calcSheet = (rows: XlsxSheet['rows']): XlsxSheet => ({ name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows })
const PIVOT_A1 = [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }]
const pivotA1 = (value: number) => ({ readCellValue: (sheet: string, col: number, row: number) => sheet === 'Pivot' && col === 0 && row === 0 ? value : undefined })

/** `unrelated` rows of an input and a formula that does not read the pivot, and `readers` formulas that do (each the
 * one above it, when chained). The unrelated cells are returned so a test can count what touches them. */
function bigDoc(unrelated: number, readers: number, chain: boolean): { doc: XlsxDocument; unrelatedCells: XlsxCell[] } {
  const rows: XlsxSheet['rows'] = [], unrelatedCells: XlsxCell[] = []
  for (let r = 0; r < unrelated; r++) {
    const input: XlsxCell = { ref: `A${r + 1}`, col: 0, row: r, value: r, styleIndex: 0, hasCachedValue: true }
    const formula: XlsxCell = { ref: `B${r + 1}`, col: 1, row: r, value: r * 2, formula: `A${r + 1}*2+ROUND(A${r + 1}/3,2)`, styleIndex: 0, hasCachedValue: true }
    unrelatedCells.push(input, formula)
    const cells = [input, formula]
    if (r < readers) cells.push({ ref: `C${r + 1}`, col: 2, row: r, value: 2, formula: chain && r > 0 ? `C${r}+1` : 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true })
    rows.push({ index: r, cells })
  }
  return { doc: { sheets: [pivotSheet(), calcSheet(rows)], definedNames: [], tables: [] }, unrelatedCells }
}
/** Counts every read or write of the given cells' properties from now on. */
function countTouches(cells: readonly XlsxCell[]): { count(): number; reset(): void } {
  let touches = 0
  for (const cell of cells) for (const key of Object.keys(cell)) {
    let held = (cell as unknown as Record<string, unknown>)[key]
    Object.defineProperty(cell, key, { enumerable: true, configurable: true,
      get() { touches++; return held }, set(next) { touches++; held = next } })
  }
  return { count: () => touches, reset: () => { touches = 0 } }
}

for (const [readers, chain] of [[1, false], [300, false], [300, true]] as const) {
  it(`${readers} ${chain ? 'chained' : 'direct'} reader${readers > 1 ? 's' : ''} of a changed pivot cell among 20,000 unrelated formulas: a commit touches none of them`, () => {
    const { doc, unrelatedCells } = bigDoc(20_000, readers, chain)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const touches = countTouches(unrelatedCells)
    // The first commit after the load builds the reader index: one pass over the cells, not one per reader.
    expect(recomputeDependents(doc, PIVOT_A1, pivotA1(10)).recomputed).toBe(readers)
    expect(touches.count()).toBeLessThanOrEqual(4 * unrelatedCells.length)
    touches.reset()
    // Every later commit evaluates the readers only.
    const result = recomputeDependents(doc, PIVOT_A1, pivotA1(11))
    expect(result.recomputed).toBe(readers)
    expect(touches.count()).toBe(0)
    const last = doc.sheets[1]!.rows[readers - 1]!.cells[2]!
    expect(last.value).toBe(chain ? 22 + readers - 1 : 22)
  }, 120_000)
}

const TRAVERSALS = new Set<PropertyKey>([Symbol.iterator, 'reduce', 'some', 'every', 'forEach', 'map', 'filter', 'flatMap', 'find', 'findIndex', 'indexOf', 'includes'])
/** Counts every whole-array operation on each sheet's rows (any pass over a sheet's cells starts with one). */
function countRowPasses(doc: XlsxDocument): { count(): number } {
  let passes = 0
  for (const sheet of doc.sheets) {
    sheet.rows = new Proxy(sheet.rows, { get(target, key, receiver) {
      if (TRAVERSALS.has(key)) passes++
      return Reflect.get(target, key, receiver)
    } })
  }
  return { count: () => passes }
}
function pointAndRangeReaders(rowsCount: number): XlsxDocument {
  const rows: XlsxSheet['rows'] = []
  for (let row = 0; row < rowsCount; row++) rows.push({ index: row, cells: [
    { ref: `A${row + 1}`, col: 0, row, value: row, styleIndex: 0, hasCachedValue: true },
    { ref: `B${row + 1}`, col: 1, row, value: row * 2, formula: `A${row + 1}*2`, styleIndex: 0, hasCachedValue: true },
  ] })
  // One point reader and one range reader of the pivot.
  rows[0]!.cells.push({ ref: 'C1', col: 2, row: 0, value: 2, formula: 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true })
  rows[1]!.cells.push({ ref: 'C2', col: 2, row: 1, value: 3, formula: 'SUM(Pivot!A1:A3)', styleIndex: 0, hasCachedValue: true })
  return { sheets: [pivotSheet(3), calcSheet(rows)], definedNames: [], tables: [] }
}
const PIVOT_A1_A3 = [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 2, maxCol: 0 } }]
const calc = (doc: XlsxDocument, row: number, col: number) => doc.sheets[1]!.rows[row]!.cells.find(cell => cell.col === col)!

it('a later commit with no document change makes no pass over any sheet\'s rows', () => {
  const doc = pointAndRangeReaders(20_000)
  evaluateWorkbookFormulas(doc)
  recomputeDependents(doc, PIVOT_A1_A3, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 5 : undefined })
  expect([calc(doc, 0, 2).value, calc(doc, 1, 2).value]).toEqual([10, 15])
  const passes = countRowPasses(doc)
  const result = recomputeDependents(doc, PIVOT_A1_A3, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 7 : undefined })
  expect(result.recomputed).toBe(2)
  expect([calc(doc, 0, 2).value, calc(doc, 1, 2).value]).toEqual([14, 21])
  expect(passes.count()).toBe(0)
})

it('S3-AJ: an announced edit of the cells rebuilds the indexes from them', () => {
  const doc = pointAndRangeReaders(200)
  evaluateWorkbookFormulas(doc)
  recomputeDependents(doc, PIVOT_A1_A3, { readCellValue: () => 5 })
  calc(doc, 2, 1).formula = 'Pivot!A2*10'   // B3 now reads the pivot
  noteWorkbookCellsChanged(doc)
  const passes = countRowPasses(doc)
  const result = recomputeDependents(doc, PIVOT_A1_A3, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 4 : undefined })
  expect(passes.count()).toBeGreaterThan(0)
  expect(result.recomputed).toBe(3)
  expect(calc(doc, 2, 1).value).toBe(40)
})

it('S3-AE: point readers well under the visit budget are all recomputed', () => {
  const rows: XlsxSheet['rows'] = []
  for (let row = 0; row < 80; row++) rows.push({ index: row, cells: [{ ref: `A${row + 1}`, col: 0, row, value: row, styleIndex: 0, hasCachedValue: true }, { ref: `B${row + 1}`, col: 1, row, value: 0, formula: 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true }] })
  const doc: XlsxDocument = { sheets: [pivotSheet(), calcSheet(rows)], definedNames: [], tables: [] }
  evaluateWorkbookFormulas(doc)
  const result = recomputeDependents(doc, PIVOT_A1, { readCellValue: () => 5 })
  expect(result.recomputed).toBe(80)
  expect(result.stale).toEqual([])
  expect(doc.sheets[1]!.rows[0]!.cells[1]!.value).toBe(10)
})

it('S3-AE: SUM readers whose ranges would each enumerate the workbook are refused before any runs, keeping their values', () => {
  const rows: XlsxSheet['rows'] = []
  for (let r = 0; r < 1600; r++) {
    const cells: XlsxSheet['rows'][number]['cells'] = [{ ref: `A${r + 1}`, col: 0, row: r, value: r, styleIndex: 0, hasCachedValue: true }]
    if (r < 1000) cells.push({ ref: `B${r + 1}`, col: 1, row: r, value: 777, formula: 'SUM(Pivot!A1:A3)', styleIndex: 0, hasCachedValue: true })
    rows.push({ index: r, cells })
  }
  const doc: XlsxDocument = { sheets: [pivotSheet(3), calcSheet(rows)], definedNames: [], tables: [] }
  evaluateWorkbookFormulas(doc) // valid cached values remain the displayed values
  const result = recomputeDependents(doc, PIVOT_A1_A3)
  expect(result.recomputed).toBe(0)
  expect(result.overBudget).toBe(true)
  expect(result.stale).toHaveLength(1000)
  expect(doc.sheets[1]!.rows[0]!.cells[1]!.value).toBe(777)
})
