/**
 * CORE-REVISION3 gates (RED first):
 * 1. Numeric text coercion: surrounding-whitespace trim before the base-10
 *    grammar (currency/grouping still applied), empty/allspace/hex rejected,
 *    nonfinite parsed text (incl. overflow) → #VALUE! (never #NUM!).
 * 2. Binary combine finishes on a known left TYPED error — the right child is
 *    never scheduled/read (native Excel L→R, no RHS callbacks).
 * 3. Reusable AST/context: locally-created node memo is NOT installed
 *    persistently on the caller context (workbook-provided memos stay).
 */
import { describe, expect, it } from 'vitest'
import { evaluateFormula, evaluateFormulaInternal, evaluateNode } from '../src/xlsx/formula/evaluator'
import { parseFormula } from '../src/xlsx/formula/parser'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

describe('REV3-1: numeric text coercion with surrounding whitespace', () => {
  it('trimmed signed whitespace forms coerce natively', () => {
    expect(evaluateFormula('" 5"+1')).toBe(6)
    expect(evaluateFormula('"5 "+1')).toBe(6)
    expect(evaluateFormula('" 5 "+1')).toBe(6)
    expect(evaluateFormula('" $5 "+1')).toBe(6)
    expect(evaluateFormula('SUM(" 5")')).toBe(5)
    expect(evaluateFormula('COUNT(" 5")')).toBe(1)
  })
  it('nonfinite parsed text is #VALUE! (never #NUM!), incl. overflow percent', () => {
    expect(evaluateFormula('"1e309"+1')).toBe('#VALUE!')
    expect(evaluateFormula('"-1e309"+1')).toBe('#VALUE!')
    expect(evaluateFormula('"1e309%"+1')).toBe(1e307) // native: percent exponent fuses (1e309 × 1e-2)
  })
  it('grouping: malformed single group rejected; double groups accepted', () => {
    expect(evaluateFormula('"1,23"+1')).toBe('#VALUE!')
    expect(evaluateFormula('"1,234,567"+1')).toBe(1234568)
    expect(evaluateFormula('"1234,567"+1')).toBe(1234568)
    expect(evaluateFormula('"+10%"+1')).toBe(1.1)
  })
  it('empty/allspace/hex still rejected', () => {
    expect(evaluateFormula('""+1')).toBe('#VALUE!')
    expect(evaluateFormula('"   "+1')).toBe('#VALUE!')
    expect(evaluateFormula('"0x10"+1')).toBe('#VALUE!')
  })
})

describe('REV3-2: binary finishes on known left typed error; no RHS scheduling', () => {
  it('error + formula-ref: right never read, propagate left error', () => {
    const reads: string[] = []
    const ctx = {
      getCellValue: (_s: string | undefined, col: number) => {
        reads.push(`c${col}`)
        return 1
      },
    } as never
    expect(evaluateFormula('#N/A+B1', ctx)).toBe('#N/A')
    expect(reads, 'no RHS getCellValue callback for a known left typed error').toEqual([])
  })
  it('1/0+formula-ref stays #DIV/0! without reading the RHS', () => {
    const reads: string[] = []
    const ctx = {
      getCellValue: (_s: string | undefined, col: number) => {
        reads.push(`c${col}`)
        return 1
      },
    } as never
    expect(evaluateFormula('1/0+B1', ctx)).toBe('#DIV/0!')
    expect(reads).toEqual([])
  })
  it('workbook A1=#N/A+B1 with conditional child keeps baseline [#N/A, #N/A]', () => {
    const d: XlsxDocument = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: null, formula: '=#N/A+B1' },
      { ref: 'B1', row: 0, col: 1, value: null, formula: '=A1' },
    ] as unknown as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const a = d.sheets[0].rows[0].cells[0] as XlsxCell
    const b = d.sheets[0].rows[0].cells[1] as XlsxCell
    expect(a.value).toBe('#N/A')
    expect(a.valueIsError ?? false).toBe(true)
    expect(b.value).toBe('#N/A')
  })
  it('workbook A1=1/0+B1 stays #DIV/0! (b298 baseline), right child never read', () => {
    const d: XlsxDocument = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: null, formula: '=1/0+B1' },
      { ref: 'B1', row: 0, col: 1, value: null, formula: '=A1' },
    ] as unknown as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const a = d.sheets[0].rows[0].cells[0] as XlsxCell
    const b = d.sheets[0].rows[0].cells[1] as XlsxCell
    expect(a.value).toBe('#DIV/0!')
    expect(b.value).toBe('#DIV/0!')
  })
})

describe('REV3-3: reusable AST/context — no persistent self-installed memo', () => {
  it('parseFormula AST + mutable ctx: 2 then 6 (baseline), not stale 2 then 2', () => {
    const ast = parseFormula('A1+1')
    let calls = 0
    const ctx = {
      getCellValue: () => [1, 5][Math.min(calls++, 1)],
    } as never
    expect(evaluateNode(ast as never, ctx)).toBe(2)
    expect(evaluateNode(ast as never, ctx)).toBe(6)
  })
  it('locally created nodeValues is restored (never installed persistently)', () => {
    const ctx = { getCellValue: () => 1 } as never
    evaluateFormula('A1+1', ctx)
    expect((ctx as { nodeValues?: unknown }).nodeValues).toBeUndefined()
  })
  it('evaluateFormulaInternal same behavior under typedValues', () => {
    const ast = parseFormula('A1+1')
    let calls = 0
    const ctx = {
      typedValues: true,
      getCellValue: () => [1, 5][Math.min(calls++, 1)],
    } as never
    expect(evaluateFormulaInternal(ast as never, ctx)).toBe(2)
    expect(evaluateFormulaInternal(ast as never, ctx)).toBe(6)
  })
})
