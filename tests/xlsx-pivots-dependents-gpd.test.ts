// GETPIVOTDATA counter-examples (Task 31 review).
import { describe, expect, it } from 'vitest'
import { evaluateWorkbookFormulas, recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerGetPivotData, getPivotData } from '../src/xlsx/pivots/session/getpivotdata'
import { load, cellAt, addCell, liveContext, viewerReadCellValue, changedOf, F } from './helpers/xlsx-pivots-dependent-fixtures'

registerGetPivotData()

async function run(formulas: Array<[sheet: string, ref: string, col: number, row: number, formula: string]>, opts: { hide?: string; viewerShape?: boolean; file?: string } = {}) {
  const doc = await load(opts.file)
  for (const [sheet, ref, col, row, formula] of formulas) addCell(doc, sheet, ref, col, row, formula, -123456789)
  evaluateWorkbookFormulas(doc, { forceRecalc: true }) // sidestep F1 so only GETPIVOTDATA semantics are probed
  const { ctx } = await liveContext(opts.hide, opts.file)
  const contexts = [ctx]
  const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
    // viewerShape: InteractiveSheetDoc.tsx passes no current sheet (getPivotData(contexts, args, evaluate)).
    getPivotData: (args, evaluate, currentSheet) => getPivotData(contexts, args, evaluate, opts.viewerShape ? currentSheet : 'Formulas'),
    readCellValue: viewerReadCellValue(contexts),
  })
  return { doc, result }
}

describe('GETPIVOTDATA', () => {
  it('F2 (Critical): a same-sheet, unqualified anchor (what Excel writes) evaluates in the viewer, not #REF!', async () => {
    const { doc } = await run([['Pivot', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",$A$3)']], { hide: 'West', viewerShape: true })
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(161104.2, 6)
  })

  it('F2 control (passes): a sheet-qualified anchor works through the same viewer call shape', async () => {
    // control: the shipped oracle formulas are qualified (Pivot!$A$3) and pass through the viewer shape.
    const { doc } = await run([['Formulas', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3)']], { hide: 'West', viewerShape: true })
    expect(cellAt(doc, 'Formulas', 'K1')!.value).toBeCloseTo(161104.2, 6)
  })

  it('F5 (Important): a formula reading a GETPIVOTDATA cell recomputes with it', async () => {
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'B1*2']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'B1')!.value).toBeCloseTo(161104.2, 6)
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBeCloseTo(322208.4, 6)
  })

  it('F6 (Important, S3-T): an unmeasured field on no axis keeps its saved value', async () => {
    const { doc } = await run([
      ['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Channel","NoSuchItem")'],
      ['Formulas', 'C2', 2, 1, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region","East","Channel","Web")'],
    ], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBe(-123456789)
    expect(cellAt(doc, 'Formulas', 'C2')!.value).toBe(-123456789)
  })

  it('F7 (Important, unmeasured): an inner item unique to one parent ("Elderberries") does not get a value the measured rule says is #REF!', async () => {
    // Measured: Product "Cherries" alone is #REF!. "Elderberries" (only under Central) is the same shape: an inner
    // item with no line of its own. The adapter's "single candidate" rule returns 926 instead.
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Product","Elderberries")']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).not.toBe(926)
  })

  it('F8 (Important): Values-only column axis (two data fields, no column field) resolves Region totals', async () => {
    const file = F + 'engine/layout-values-on-columns-no-column-field.xlsx'
    const doc = await load(file)
    addCell(doc, 'Pivot', 'K1', 10, 0, 'GETPIVOTDATA("Sum of Sales",$A$3,"Region","East")', -1)
    addCell(doc, 'Pivot', 'K2', 10, 1, 'GETPIVOTDATA("Sum of Units",$A$3)', -1)
    evaluateWorkbookFormulas(doc, { forceRecalc: true })
    const { ctx } = await liveContext(undefined, file)
    recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), { getPivotData: (args, evaluate) => getPivotData([ctx], args, evaluate, 'Pivot') })
    expect(cellAt(doc, 'Pivot', 'K1')!.value).toBeCloseTo(48118.96, 6)
    expect(cellAt(doc, 'Pivot', 'K2')!.value).toBe(2417)
  })

  it('Minor: data_field given as the source field name ("Sales", Microsoft\'s documented form) is not #REF!', async () => {
    const { doc } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sales",Pivot!$A$3)']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).not.toBe('#REF!')
  })

  it('R2-8: a non-text field-name argument retains its saved value under S3-T', async () => {
    const { doc, result } = await run([['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,42,"East")']], { hide: 'West' })
    expect(cellAt(doc, 'Formulas', 'C1')!.value).toBe(-123456789)
    expect(result.stale.some(key => key.endsWith('!2:0'))).toBe(true)
  })

  it('S3-T: unmeasured GETPIVOTDATA shapes retain their saved values instead of inventing #REF!', async () => {
    const formulas = [
      ['Formulas', 'C1', 2, 0, 'GETPIVOTDATA("Unknown data",Pivot!$A$3)'],
      ['Formulas', 'C2', 2, 1, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3:$A$3)'],
      ['Formulas', 'C3', 2, 2, 'GETPIVOTDATA("Sum of Sales",Pivot!$A$3,"Region")'],
      ['Formulas', 'C4', 2, 3, 'GETPIVOTDATA("Sum of Units",Pivot!$A$3,"Year","2023")'],
    ] as Array<[string, string, number, number, string]>
    const { doc } = await run(formulas, { hide: 'West' })
    for (const [, ref] of formulas) expect(cellAt(doc, 'Formulas', ref)!.value).toBe(-123456789)
  })
})
