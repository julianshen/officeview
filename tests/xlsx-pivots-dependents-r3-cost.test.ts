import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents, MAX_DEPENDENT_CELL_VISITS } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

it('R3-2: admits a bounded set of point readers under the named visit budget', () => {
  const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 1, styleIndex: 0, hasCachedValue: true }] }] }
  const rows: XlsxSheet['rows'] = []
  for (let row = 0; row < 80; row++) rows.push({ index: row, cells: [{ ref: `A${row + 1}`, col: 0, row, value: row, styleIndex: 0, hasCachedValue: true }, { ref: `B${row + 1}`, col: 1, row, value: 0, formula: 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true }] })
  const doc: XlsxDocument = { sheets: [pivot, { name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows }], definedNames: [], tables: [] }
  evaluateWorkbookFormulas(doc)
  const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { readCellValue: () => 5 })
  expect(MAX_DEPENDENT_CELL_VISITS).toBe(2_000_000)
  expect(result.recomputed).toBe(80)
  expect(result.stale).toEqual([])
  expect(doc.sheets[1]!.rows[0]!.cells[1]!.value).toBe(10)
})
