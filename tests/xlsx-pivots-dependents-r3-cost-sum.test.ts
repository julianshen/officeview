import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

it('R3-2: refuses before SUM range readers can enumerate the whole workbook per formula', () => {
  const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [0, 1, 2].map(r => ({ index: r, cells: [{ ref: `A${r + 1}`, col: 0, row: r, value: r + 1, styleIndex: 0, hasCachedValue: true }] })) }
  const rows: XlsxSheet['rows'] = []
  for (let r = 0; r < 1600; r++) {
    const cells: XlsxSheet['rows'][number]['cells'] = [{ ref: `A${r + 1}`, col: 0, row: r, value: r, styleIndex: 0, hasCachedValue: true }]
    if (r < 1000) cells.push({ ref: `B${r + 1}`, col: 1, row: r, value: 777, formula: 'SUM(Pivot!A1:A3)', styleIndex: 0, hasCachedValue: true })
    rows.push({ index: r, cells })
  }
  const doc: XlsxDocument = { sheets: [pivot, { name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows }], definedNames: [], tables: [] }
  evaluateWorkbookFormulas(doc) // valid cached values remain the displayed values
  const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 2, maxCol: 0 } }])
  expect(result.recomputed).toBe(0)
  expect(result.stale).toHaveLength(1000)
  expect(doc.sheets[1]!.rows[0]!.cells[1]!.value).toBe(777)
})
