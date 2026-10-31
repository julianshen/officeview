// Round 2: the syntactic reader index (workbook.ts referenceIndexFor) from a plain parse of an Excel-shaped file.
import { describe, expect, it } from 'vitest'
import { recomputeDependents } from '../src/xlsx/formula/workbook'
import { registerGetPivotData, getPivotData } from '../src/xlsx/pivots/session/getpivotdata'
import { cellAt, liveContext, changedOf } from './helpers/xlsx-pivots-dependent-fixtures'
import { patchedSeed } from './helpers/xlsx-pivots-dependent-patched'

registerGetPivotData()

async function run(cells: Array<{ row: number; formula: string; cached: number | string }>, names: Array<[string, string]> = []) {
  const doc = await patchedSeed(cells, names)
  const { ctx } = await liveContext('West')
  // The live Pivot!B6 now reads 999 (stand-in for any committed change at that cell).
  const result = recomputeDependents(doc, changedOf(doc, 'Pivot', ctx.rects), {
    getPivotData: (args, evaluate, sheet) => getPivotData([ctx], args, evaluate, sheet),
    readCellValue: (n, c, r) => n.toLowerCase() === 'pivot' && c === 1 && r === 5 ? 999 : undefined,
  })
  return { doc, result, k: (row: number) => cellAt(doc, 'Formulas', `K${row}`)!.value }
}

describe('syntactic reader index', () => {
  it('control: plain ref, defined name, 3-D ref, whole column, whole row into the pivot recompute', async () => {
    const { k } = await run([
      { row: 1, formula: 'Pivot!B6*1', cached: 11204.24 },
      { row: 2, formula: 'SalesCell*1', cached: 11204.24 },
      { row: 3, formula: 'SUM(Pivot:Pivot!B6)', cached: 11204.24 },
      { row: 4, formula: 'MAX(Pivot!B:B)', cached: -1 },
      { row: 5, formula: 'MAX(Pivot!6:6)', cached: -1 },
      { row: 6, formula: 'PRANGE*1', cached: -1 },
    ], [['SalesCell', 'Pivot!$B$6'], ['PRANGE', 'Pivot!$B$6:$B$6']])
    // K4/K5 were cached -1: any recomputed MAX is the column/row maximum (>= 999).
    expect({ ref: k(1), threeD: k(3), wholeCol: (k(4) as number) >= 999, wholeRow: (k(5) as number) >= 999 }).toEqual({ ref: 999, threeD: 999, wholeCol: true, wholeRow: true })
  })

  it('control: a defined name into the pivot recomputes (Excel-saved <definedName>)', async () => {
    const { k } = await run([
      { row: 2, formula: 'SalesCell*1', cached: 11204.24 },
      { row: 6, formula: 'PRANGE*1', cached: 11204.24 },
    ], [['SalesCell', 'Pivot!$B$6'], ['PRANGE', 'Pivot!$B$6:$B$6']])
    expect([k(2), k(6)]).toEqual([999, 999])
  })

  it('R2-6 Minor: rebuilds the runtime cell index when cells are added after full evaluation', async () => {
    const { doc, k } = await run([{ row: 1, formula: 'Pivot!B6*1', cached: 11204.24 }])
    const formulas = doc.sheets.find(sheet => sheet.name === 'Formulas')!
    formulas.rows[0]!.cells.push({ ref: 'Z1', col: 25, row: 0, value: 1, styleIndex: 0, hasCachedValue: true })
    const result = recomputeDependents(doc, changedOf(doc, 'Pivot', [{ r0: 5, r1: 5, c0: 1, c1: 1 }]), { readCellValue: () => 12345 })
    expect(result.recomputed).toBeGreaterThan(0)
    expect(k(1)).toBe(12345)
  })

  it('R2-2 (Important): INDIRECT into the pivot is neither recomputed nor flagged (no conservative fallback)', async () => {
    const { k, result } = await run([{ row: 1, formula: 'INDIRECT("Pivot!B6")', cached: 11204.24 }])
    expect(k(1)).toBe(11204.24) // unsupported runtime reference keeps the saved value
    expect(result.stale).toHaveLength(1)
  })

  it('R2-2 (Important): OFFSET from a base outside the pivot into it is neither recomputed nor flagged', async () => {
    const { k, result } = await run([{ row: 1, formula: 'OFFSET(Pivot!$A$1,5,1)', cached: 11204.24 }])
    expect(k(1)).toBe(11204.24) // unsupported runtime reference keeps the saved value
    expect(result.stale).toHaveLength(1)
  })

  it('R2-2 (Important): a name defined by OFFSET into the pivot', async () => {
    const { k, result } = await run([{ row: 1, formula: 'DYN*1', cached: 11204.24 }], [['DYN', 'OFFSET(Pivot!$A$1,5,1)']])
    expect(k(1)).toBe(11204.24) // unsupported runtime reference keeps the saved value
    expect(result.stale).toHaveLength(1)
  })
})
