// Dependents on a plain file load: the reader index is built from the parsed formulas, so cell references into a
// pivot recompute without a forced full evaluation first.
import { expect, it } from 'vitest'
import { recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerGetPivotData, getPivotData } from '../src/xlsx/pivots/session/getpivotdata'
import { load, cellAt, liveContext, viewerReadCellValue, changedOf } from './helpers/xlsx-pivots-dependent-fixtures'

it('after a plain parseXlsx (as the viewer gets it), =Pivot!B6 and =SUM(Pivot!B6:B8) follow a hide-West commit', async () => {
  registerGetPivotData()
  const doc = await load() // parse.ts:539 file-load evaluation only; no forceRecalc
  const { ctx } = await liveContext('West')
  const contexts = [ctx]
  const before = { b11: cellAt(doc, 'Formulas', 'B11')!.value, b12: cellAt(doc, 'Formulas', 'B12')!.value }
  const live = viewerReadCellValue(contexts)
  const b6 = live('Pivot', 1, 5), b7 = live('Pivot', 1, 6), b8 = live('Pivot', 1, 7)
  // Make the change observable regardless of which rows West occupied: override Pivot!B6 through the overlay.
  const read = (name: string, col: number, row: number) => name === 'Pivot' && col === 1 && row === 5 ? 999 : live(name, col, row)
  const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
    getPivotData: (args, evaluate) => getPivotData(contexts, args, evaluate, 'Formulas'), readCellValue: read,
  })
  expect(before.b11).toBe(11204.24)
  expect(cellAt(doc, 'Formulas', 'B11')!.value).toBe(999)
  expect(cellAt(doc, 'Formulas', 'B12')!.value).toBeCloseTo(999 + Number(b7 ?? 0) + Number(b8 ?? 0), 6)
  void b6; void result
})

it('the same recompute after a forced full evaluation', async () => {
  const { evaluateWorkbookFormulas } = await import('../src/xlsx/formula/workbook')
  registerGetPivotData()
  const doc = await load()
  evaluateWorkbookFormulas(doc, { forceRecalc: true })
  const { ctx } = await liveContext('West')
  const live = viewerReadCellValue([ctx])
  recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
    getPivotData: (args, evaluate) => getPivotData([ctx], args, evaluate, 'Formulas'),
    readCellValue: (n, c, r) => n === 'Pivot' && c === 1 && r === 5 ? 999 : live(n, c, r),
  })
  expect(cellAt(doc, 'Formulas', 'B11')!.value).toBe(999)
})
