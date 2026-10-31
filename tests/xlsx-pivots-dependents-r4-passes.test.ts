// Round 4, M-1 (and the perf audit's finding 2): a pivot commit with no document mutation walks no worksheet.
import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, noteWorkbookCellsChanged, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

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

function workbook(rowsCount: number): XlsxDocument {
  const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'P', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [0, 1, 2].map(row => ({ index: row, cells: [{ ref: `A${row + 1}`, col: 0, row, value: 1, styleIndex: 0, hasCachedValue: true }] })) }
  const rows: XlsxSheet['rows'] = []
  for (let row = 0; row < rowsCount; row++) rows.push({ index: row, cells: [
    { ref: `A${row + 1}`, col: 0, row, value: row, styleIndex: 0, hasCachedValue: true },
    { ref: `B${row + 1}`, col: 1, row, value: row * 2, formula: `A${row + 1}*2`, styleIndex: 0, hasCachedValue: true },
  ] })
  // One point reader and one range reader of the pivot.
  rows[0]!.cells.push({ ref: 'C1', col: 2, row: 0, value: 2, formula: 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true })
  rows[1]!.cells.push({ ref: 'C2', col: 2, row: 1, value: 3, formula: 'SUM(Pivot!A1:A3)', styleIndex: 0, hasCachedValue: true })
  return { sheets: [pivot, { name: 'Calc', sheetId: 'C', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows }], definedNames: [], tables: [] }
}
const changed = [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 2, maxCol: 0 } }]
const calc = (doc: XlsxDocument, row: number, col: number) => doc.sheets[1]!.rows[row]!.cells.find(cell => cell.col === col)!

it('M-1: a second commit with no document mutation does no whole-workbook pass', () => {
  const doc = workbook(20_000)
  evaluateWorkbookFormulas(doc)
  // The first commit after the load builds the reader index (one pass).
  recomputeDependents(doc, changed, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 5 : undefined })
  expect([calc(doc, 0, 2).value, calc(doc, 1, 2).value]).toEqual([10, 15])
  const passes = countRowPasses(doc)
  const result = recomputeDependents(doc, changed, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 7 : undefined })
  expect(result.recomputed).toBe(2)
  expect([calc(doc, 0, 2).value, calc(doc, 1, 2).value]).toEqual([14, 21])
  expect(passes.count()).toBe(0)
})

it('M-1: an announced edit rebuilds the indexes from the cells', () => {
  const doc = workbook(200)
  evaluateWorkbookFormulas(doc)
  recomputeDependents(doc, changed, { readCellValue: () => 5 })
  calc(doc, 2, 1).formula = 'Pivot!A2*10'   // B3 now reads the pivot
  noteWorkbookCellsChanged(doc)
  const passes = countRowPasses(doc)
  const result = recomputeDependents(doc, changed, { readCellValue: (sheet, _col, row) => sheet === 'Pivot' && row <= 2 ? 4 : undefined })
  expect(passes.count()).toBeGreaterThan(0)
  expect(result.recomputed).toBe(3)
  expect(calc(doc, 2, 1).value).toBe(40)
})
