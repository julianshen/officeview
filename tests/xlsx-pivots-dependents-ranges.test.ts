// Range readers of a committed pivot: aggregates count the cells the live layout adds, and a cell two pivots share
// (one vacated it, the other grew into it) reads as the composite overlay shows it.
import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { viewerReadCellValue } from './helpers/xlsx-pivots-dependent-fixtures'

function sheet(name: string, cells: Array<{ ref: string; col: number; row: number; value: string | number | null; formula?: string }>): XlsxSheet {
  const rows = new Map<number, XlsxSheet['rows'][number]>()
  for (const cell of cells) {
    let row = rows.get(cell.row)
    if (!row) rows.set(cell.row, row = { index: cell.row, cells: [] })
    row.cells.push({ ...cell, styleIndex: 0, hasCachedValue: cell.value !== null })
  }
  return { name, sheetId: name, workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [...rows.values()] }
}

it('SUM over a pivot range counts cells the committed layout grew into', async () => {
  const doc: XlsxDocument = { definedNames: [], tables: [], sheets: [
    sheet('Pivot', [{ ref: 'A1', col: 0, row: 0, value: 1 }]),
    sheet('Calc', [{ ref: 'A1', col: 0, row: 0, value: 1, formula: 'SUM(Pivot!A1:A3)' }, { ref: 'B1', col: 1, row: 0, value: 1, formula: 'COUNTA(Pivot!A1:A3)' }]),
  ] }
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  // The live pivot now fills A1:A3 with 1, 2, 3 (it grew by two rows; the overlay is view-only, not sheet cells).
  recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 2, maxCol: 0 } }], {
    readCellValue: (name, col, row) => name === 'Pivot' && col === 0 && row <= 2 ? row + 1 : undefined,
    overlayEntries: [0, 1, 2].map(row => ({ sheetName: 'Pivot', col: 0, row, value: row + 1 })),
  })
  expect(doc.sheets[1]!.rows[0]!.cells[0]!.value).toBe(6)
  expect(doc.sheets[1]!.rows[0]!.cells[1]!.value).toBe(3)
})

it('two committed pivots: a cell B grew into (A vacated) reads B\'s value, as the overlay shows it', async () => {
  const layout = (minRow: number, maxRow: number, minCol: number, maxCol: number, cells: Array<[number, number, number]>) =>
    ({ rect: { minRow, maxRow, minCol, maxCol }, merges: [], cells: cells.map(([row, col, value]) => ({ row, col, value: { kind: 'number', value } })) }) as never
  // compositeOverlay's own test case (tests/xlsx-viewer-pivot-ui.test.tsx:157): overlay.cell(5, 0) is { value: 50 }.
  const a = { sheet: 'P', def: {} as never, cache: {} as never, date1904: false, layout: layout(0, 2, 0, 1, [[0, 0, 1], [2, 1, 2]]), rects: [{ r0: 0, r1: 5, c0: 0, c1: 1 }, { r0: 0, r1: 2, c0: 0, c1: 1 }] }
  const b = { sheet: 'P', def: {} as never, cache: {} as never, date1904: false, layout: layout(5, 8, 0, 1, [[5, 0, 50], [8, 1, 80]]), rects: [{ r0: 6, r1: 8, c0: 0, c1: 1 }, { r0: 5, r1: 8, c0: 0, c1: 1 }], overlay: { cell: (row: number, col: number) => row === 5 && col === 0 ? { value: 50 } : undefined } as never }
  const read = viewerReadCellValue([a, b]) // pivotFormulaContexts insertion order: A committed first
  expect(read('P', 0, 5)).toBe(50)
})
