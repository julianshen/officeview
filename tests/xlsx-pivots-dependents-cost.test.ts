// F11: recomputeDependents' cost scales with the workbook, not with the readers of the changed rectangle.
import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

function bigDoc(unrelated: number): XlsxDocument {
  const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 1, styleIndex: 0, hasCachedValue: true }] }] }
  const rows: XlsxSheet['rows'] = []
  for (let r = 0; r < unrelated; r++) rows.push({ index: r, cells: [{ ref: `A${r + 1}`, col: 0, row: r, value: r, styleIndex: 0, hasCachedValue: true }, { ref: `B${r + 1}`, col: 1, row: r, value: r * 2, formula: `A${r + 1}*2+ROUND(A${r + 1}/3,2)`, styleIndex: 0, hasCachedValue: true }] })
  rows[0]!.cells.push({ ref: 'C1', col: 2, row: 0, value: 2, formula: 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true })
  const calc: XlsxSheet = { name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows }
  return { sheets: [pivot, calc], definedNames: [], tables: [] }
}

it('F11 (Important): one reader of a changed pivot costs far less than a full workbook evaluation', () => {
  const doc = bigDoc(60_000)
  let t = performance.now()
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  const full = performance.now() - t
  const times: number[] = []
  for (let i = 0; i < 3; i++) {
    t = performance.now()
    const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { readCellValue: (n, c, r) => n === 'Pivot' && c === 0 && r === 0 ? 10 + i : undefined })
    times.push(performance.now() - t)
    expect(result.recomputed).toBe(1)
  }
  const one = Math.min(...times)
  console.log(`full evaluation ${full.toFixed(0)} ms; recomputeDependents with 1 reader ${one.toFixed(0)} ms (${(100 * one / full).toFixed(0)}% of full)`)
  expect(one).toBeLessThan(full * 0.1)
})
