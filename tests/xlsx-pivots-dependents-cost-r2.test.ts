// Round 2: cost of recomputeDependents when the change has many readers (not just one).
import { expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'

function bigDoc(unrelated: number, readers: number, chain: boolean): XlsxDocument {
  const pivot: XlsxSheet = { name: 'Pivot', sheetId: 'Pivot', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [{ ref: 'A1', col: 0, row: 0, value: 1, styleIndex: 0, hasCachedValue: true }] }] }
  const rows: XlsxSheet['rows'] = []
  for (let r = 0; r < unrelated; r++) {
    const cells: XlsxSheet['rows'][number]['cells'] = [{ ref: `A${r + 1}`, col: 0, row: r, value: r, styleIndex: 0, hasCachedValue: true }, { ref: `B${r + 1}`, col: 1, row: r, value: r * 2, formula: `A${r + 1}*2+ROUND(A${r + 1}/3,2)`, styleIndex: 0, hasCachedValue: true }]
    if (r < readers) cells.push({ ref: `C${r + 1}`, col: 2, row: r, value: 2, formula: chain && r > 0 ? `C${r}+1` : 'Pivot!A1*2', styleIndex: 0, hasCachedValue: true })
    rows.push({ index: r, cells })
  }
  return { sheets: [pivot, { name: 'Calc', sheetId: 'Calc', workbookIndex: 1, cols: [], merges: [], mergeRanges: [], rows }], definedNames: [], tables: [] }
}

for (const chain of [false, true]) it(`R2-4 (Critical): 60k unrelated formulas + 3000 ${chain ? 'chained' : 'direct'} readers: recompute stays well under a full evaluation`, () => {
  const doc = bigDoc(60_000, 3000, chain)
  let t = performance.now()
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  const full = performance.now() - t
  t = performance.now()
  const result = recomputeDependents(doc, [{ sheetId: 0, rect: { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 } }], { readCellValue: (n, c, r) => n === 'Pivot' && c === 0 && r === 0 ? 10 : undefined })
  const one = performance.now() - t
  console.log(`${chain ? 'chain' : 'direct'}: full ${full.toFixed(0)} ms; recompute of ${result.recomputed} readers ${one.toFixed(0)} ms`)
  expect(result.recomputed).toBe(3000)
  expect(one).toBeLessThan(full)
}, 120_000)
