/**
 * S1.1b evaluator budget (design note §4.3 rules 3-5, §7.6): the additive
 * cumulative `maxMaterializedCells` counter and the `parseXlsx` evaluation
 * options. Defaults leave existing callers unchanged.
 */
import { describe, expect, it } from 'vitest'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import type { XlsxCell, XlsxDocument, XlsxRow, XlsxSheet } from '../src/xlsx/types'

function mkCell(ref: string, col: number, row: number, extra: Partial<XlsxCell> = {}): XlsxCell {
  return { ref, col, row, value: null, styleIndex: 0, hasCachedValue: false, ...extra }
}

function docOf(cells: XlsxCell[]): XlsxDocument {
  const byRow = new Map<number, XlsxCell[]>()
  for (const c of cells) byRow.set(c.row, [...(byRow.get(c.row) ?? []), c])
  const rows: XlsxRow[] = [...byRow.keys()].sort((a, b) => a - b).map((index) => ({ index, cells: byRow.get(index)! }))
  const sheet: XlsxSheet = { name: 'S', sheetId: 's', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows }
  return { sheets: [sheet], definedNames: [], tables: [] }
}

const resourceHits = (doc: XlsxDocument) => (doc.diagnostics ?? []).filter((d) => d.feature === 'array-resource-limit')
const rowCount = (doc: XlsxDocument) => doc.sheets[0].rows.length
const DYNAMIC = { forceRecalc: true, arrayMode: 'dynamic' as const }

describe('cumulative maxMaterializedCells', () => {
  it('is charged across formulas and never reset within one evaluation', () => {
    const cells = () => [mkCell('A1', 1, 1, { formula: 'SEQUENCE(3000,1)' }), mkCell('C1', 3, 1, { formula: 'SEQUENCE(3000,1)' })]
    const tight = docOf(cells())
    evaluateWorkbookFormulas(tight, { ...DYNAMIC, maxMaterializedCells: 5000 })
    // the first spill fits (3,000 rows published); the second trips the shared counter
    expect(resourceHits(tight).filter((d) => /cumulative/.test(d.message))).toHaveLength(1)
    expect(tight.sheets[0].rows.find((r) => r.index === 1)!.cells.find((c) => c.ref === 'C1')!.value).toBe('#NAME?')
    expect(rowCount(tight)).toBe(3000)

    const roomy = docOf(cells())
    evaluateWorkbookFormulas(roomy, { ...DYNAMIC, maxMaterializedCells: 6000 })
    expect(resourceHits(roomy)).toHaveLength(0)
    expect(rowCount(roomy)).toBe(3000)
  })

  it('charges reference materialization too (SORT over 5,000 cells)', () => {
    const values = Array.from({ length: 5000 }, (_, i) => mkCell(`A${i + 1}`, 1, i + 1, { value: 5000 - i, hasCachedValue: true }))
    const make = () => docOf([...values.map((c) => ({ ...c })), mkCell('C1', 3, 1, { formula: 'SORT(A1:A5000,1,1)' })])
    const tight = make()
    evaluateWorkbookFormulas(tight, { ...DYNAMIC, maxMaterializedCells: 9_999 })
    expect(resourceHits(tight).length).toBeGreaterThan(0)
    const roomy = make()
    evaluateWorkbookFormulas(roomy, { ...DYNAMIC, maxMaterializedCells: 86_264 })
    expect(resourceHits(roomy)).toHaveLength(0)
  })

  it('the per-array cap stops SEQUENCE(200000,1) before 200,000 rows are built', () => {
    const doc = docOf([mkCell('A1', 1, 1, { formula: 'SEQUENCE(200000,1)' })])
    evaluateWorkbookFormulas(doc, { ...DYNAMIC, maxArrayCells: 16_384, maxArraySide: 16_384, maxMaterializedCells: 50_000 })
    expect(resourceHits(doc).length).toBeGreaterThan(0)
    expect(rowCount(doc)).toBeLessThan(10)
  })

  it('no option means no cumulative limit (default behaviour unchanged)', () => {
    const doc = docOf([mkCell('A1', 1, 1, { formula: 'SEQUENCE(3000,1)' }), mkCell('C1', 3, 1, { formula: 'SEQUENCE(3000,1)' })])
    evaluateWorkbookFormulas(doc, DYNAMIC)
    expect(resourceHits(doc)).toHaveLength(0)
  })
})

describe('parseXlsx evaluation options (additive)', () => {
  it("valuePolicy 'cached' keeps saved values and never evaluates", async () => {
    const bytes = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', formula: '1+1' }, { ref: 'B1', formula: '2+2', v: 9 }] }] }])
    const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
    const [a1, b1] = doc.sheets[0].rows[0].cells
    expect(a1.value).toBeNull()
    expect(a1.hasCachedValue).toBe(false)
    expect(b1.value).toBe(9)
  })

  it('passes the caps to the load-time evaluation; without options the result is unchanged', async () => {
    const bytes = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', formula: 'SUM(SEQUENCE(200000,1))' }] }] }])
    const plain = await parseXlsx(await OfficePackage.load(bytes))
    expect(plain.sheets[0].rows[0].cells[0].value).toBe(20000100000)
    expect(resourceHits(plain)).toHaveLength(0)

    const capped = await parseXlsx(await OfficePackage.load(bytes), { evaluation: { maxArrayCells: 16_384, maxArraySide: 16_384, maxMaterializedCells: 40_000 } })
    expect(resourceHits(capped).length).toBeGreaterThan(0)
  })
})
