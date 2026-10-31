import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, noteWorkbookCellsChanged, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument } from '../src/xlsx/types'

const cell = (ref: string, col: number, value: number, formula?: string) => ({ ref, col, row: 0, value, ...(formula ? { formula } : {}), styleIndex: 0, hasCachedValue: true })
function document(formula: string): XlsxDocument {
  return { sheets: [
    { name: 'Pivot', sheetId: 'P', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell('A1', 0, 1)] }] },
    { name: 'Calc', sheetId: 'C', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell('A1', 0, 77, formula), cell('B1', 1, 154, 'A1*2')] }] },
  ], definedNames: [], tables: [] }
}

it('R3-5: a formula reading a retained stale formula is also stale and keeps its own cached value', () => {
  const doc = document('INDIRECT("Pivot!A1")')
  evaluateWorkbookFormulas(doc)
  const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { readCellValue: () => 5 })
  expect(result.stale).toEqual(expect.arrayContaining(['calc!0:0', 'calc!1:0']))
  expect(doc.sheets[1]!.rows[0]!.cells.map(item => item.value)).toEqual([77, 154])
})

it('R3-6: an in-place formula edit rebuilds the syntactic and runtime cell indexes', () => {
  const doc = document('Pivot!A1*2')
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  const formula = doc.sheets[1]!.rows[0]!.cells[0]!
  formula.formula = 'Pivot!B1*2'
  doc.sheets[0]!.rows[0]!.cells.push(cell('B1', 1, 9))
  // T31 r4 M-1: a direct edit of the cell model is announced; the commit path no longer walks every cell to find it.
  noteWorkbookCellsChanged(doc)
  const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 1, maxRow: 0, maxCol: 1 } }])
  expect(result.recomputed).toBe(2)
  expect(formula.value).toBe(18)
})
