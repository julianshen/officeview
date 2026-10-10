/**
 * CORE-REVISION1 gates — four bounded fixes, RED first.
 * 1. Frame-owned unsupportedFeatures survive dependency suspension/resumption.
 * 2. Retained cached genuine ERROR stays tagged for dependents (IFERROR catches);
 *    error-looking cached TEXT (valueIsError false) stays text.
 * 3. Calculated blank normalizes to 0 BEFORE memo (COUNT sees 1); model 0
 *    compares as numeric 0 (B1=FALSE false) while native blank inputs stay
 *    null/TRUE for =FALSE.
 * 4. Public evaluateFormula over a typed blank ref returns native 0; unbound
 *    (no getCellValue) stays null.
 */
import { describe, expect, it } from 'vitest'
import {
  evaluateFormula,
  evaluateFormulaInternal,
} from '../src/xlsx/formula/evaluator'
import type { EvaluationContext } from '../src/xlsx/formula/types'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

const findCell = (d: XlsxDocument, ref: string): XlsxCell => {
  for (const s of d.sheets) for (const r of s.rows) {
    const c = r.cells.find(x => x.ref === ref)
    if (c) return c as XlsxCell
  }
  throw new Error(`missing ${ref}`)
}

const doc = (): XlsxDocument => ({
  sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
    // parent: cached 999, formula recalc on force, references B1
    { ref: 'A1', row: 0, col: 0, value: 999, hasCachedValue: true, formula: '=SUM(UNKNOWN(),B1)' },
    // B1: uncached formula child evaluated during the parent's suspension
    { ref: 'B1', row: 0, col: 1, value: null, formula: '=1' },
  ] as XlsxCell[] }] }],
  images: [], drawingCoverage: [],
} as never as XlsxDocument)

describe('REV1-1: frame-owned unsupportedFeatures survive suspension', () => {
  it('parent SUM(UNKNOWN(),B1) retains cache 999 + diagnostic despite the B1 child clearing', () => {
    const d = doc()
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const a1 = d.sheets[0].rows[0].cells[0] as XlsxCell
    expect(a1.value, 'cached value retained against #NAME? overwrite').toBe(999)
    const diags = (d as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics
    expect(diags?.some(x => x.feature === 'UNKNOWN'), 'frame-owned feature survives the child run').toBe(true)
  })
  it('dead IF branches still record nothing (not statically denied)', () => {
    const d: XlsxDocument = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: undefined, formula: '=IF(FALSE,SQRT(4),"ok")' },
    ] as unknown as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    expect(d.sheets[0].rows[0].cells[0].value).toBe('ok')
    expect((d as unknown as { diagnostics?: unknown[] }).diagnostics ?? []).toHaveLength(0)
  })
})

describe('REV1-2/REV2: retained cached error stays tagged; text control stays text', () => {
  /** In-pass dependent (uncached null IFERROR formula) inside the SAME forced
   * pass so the assertion is NOT vacuous: 'caught' can only come from the run. */
  const errorsDoc = (): XlsxDocument => ({ sheets: [{ name: 'S', sourcePartPath: 's', rows: [
    { index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: '#N/A', hasCachedValue: true, valueIsError: true, formula: '=UNKNOWNF(1)' },
      { ref: 'D1', row: 0, col: 3, value: null, formula: '=IFERROR(A1,"caught")' },
    ] as XlsxCell[] },
    { index: 1, cells: [
      { ref: 'A2', row: 1, col: 0, value: 'kept-as-text', hasCachedValue: true, valueIsError: false, formula: '=UNKNOWNF(2)' },
    ] as XlsxCell[] },
  ] }], images: [], drawingCoverage: [] } as never as XlsxDocument)
  it('dependent IFERROR computes against the retained TYPED error in the same pass', () => {
    const d = errorsDoc()
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const err = findCell(d, 'A1')
    expect(err.value).toBe('#N/A')
    expect(err.valueIsError).toBe(true)
    const dep = findCell(d, 'D1')
    // assert the ACTUAL computed dependent result: primitive JSON + identity flag
    expect(JSON.stringify(dep.value)).toBe(JSON.stringify('caught'))
    expect(dep.valueIsError ?? false).toBe(false)
  })
  it('error-looking cached TEXT (valueIsError false) stays text', () => {
    const d = errorsDoc()
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const text = findCell(d, 'A2')
    expect(text.value).toBe('kept-as-text')
    expect(text.valueIsError).toBe(false)
    expect((d as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics?.some(x => x.feature === 'UNKNOWNF')).toBe(true)
  })
})

describe('REV2-legacy: legacy tri-state retention identity (valueIsError undefined → canonical code tags)', () => {
  const legacyDoc = (): XlsxDocument => ({ sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
    // LEGACY hand-built API: valueIsError ABSENT (undefined) + canonical '#N/A'
    { ref: 'A1', row: 0, col: 0, value: '#N/A', hasCachedValue: true, formula: '=UNKNOWNF(1)' },
    // error-LOOKING TEXT with the explicit FALSE flag in the SAME model
    { ref: 'B1', row: 0, col: 1, value: '#N/A', hasCachedValue: true, valueIsError: false, formula: '=UNKNOWNF(2)' },
    { ref: 'D1', row: 0, col: 3, value: null, formula: '=IFERROR(A1,"caught-legacy")' },
    { ref: 'E1', row: 0, col: 4, value: null, formula: '=IFERROR(B1,"caught-text")' },
  ] as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument)
  it('legacy undefined flag + canonical code retains TAGGED (IFERROR catches); explicit false stays text', () => {
    const d = legacyDoc()
    evaluateWorkbookFormulas(d as never, { forceRecalc: true })
    const legacy = findCell(d, 'A1')
    expect(legacy.value, 'model value identity preserved').toBe('#N/A')
    const legacyDep = findCell(d, 'D1')
    expect(legacyDep.value, 'dependent computed against the TAGGED legacy error').toBe('caught-legacy')
    expect(legacyDep.valueIsError ?? false).toBe(false)
    const text = findCell(d, 'B1')
    expect(text.value, 'explicit-false flag identity stays text').toBe('#N/A')
    const textDep = findCell(d, 'E1')
    // E1: IFERROR over a TEXT '#N/A' returns the text itself, not the alternate
    expect(textDep.value).toBe('#N/A')
  })
})

describe('REV1-3: calculated blank memoizes numeric 0', () => {
  const blankDoc = (): XlsxDocument => ({ sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
    { ref: 'A1', row: 0, col: 0, value: null },
    { ref: 'B1', row: 0, col: 1, value: null, formula: '=A1' },
    { ref: 'C1', row: 0, col: 2, value: null, formula: '=COUNT(B1,B1)' },
    { ref: 'E1', row: 0, col: 4, value: null, formula: '=B1=FALSE' },
  ] as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument)
  it('B1=blank memoizes 0: uncached C1 computes COUNT exactly 2 in the same pass', () => {
    const d = blankDoc()
    evaluateWorkbookFormulas(d as never, { forceRecalc: true }) // same forced pass — non-vacuous
    const b = d.sheets[0].rows[0].cells[1] as XlsxCell
    expect(b.value, 'model writes 0').toBe(0)
    const c = d.sheets[0].rows[0].cells.find(c => c.ref === 'C1') as XlsxCell
    expect(c.value, 'COUNT over the calculated-0 ref counts EXACTLY 2 (numeric ref counted twice)').toBe(2)
    const e = d.sheets[0].rows[0].cells.find(c => c.ref === 'E1') as XlsxCell
    expect(e.value, 'calculated 0 = FALSE is FALSE (0≠FALSE rank)').toBe(false)
  })
  it('native blank INPUT cells still equal FALSE and stay ignorable in COUNT', () => {
    const blankCtx: EvaluationContext = { getCellValue: () => null }
    expect(evaluateFormula('D1=FALSE', blankCtx)).toBe(true)
    expect(evaluateFormula('COUNT(D1)', blankCtx)).toBe(0)
    expect(evaluateFormula('COUNTA(D1)', blankCtx)).toBe(0)
  })
})

describe('REV1-4: public typed-context blank ref is 0; unbound stays null', () => {
  it('evaluateFormula("D1", typed getCell null) = 0', () => {
    const typed: EvaluationContext = { getCellValue: () => null, typedValues: true }
    expect(evaluateFormulaInternalParser('D1', typed)).toBe(0)
  })
  it('no getCellValue (unbound) stays null', () => {
    expect(evaluateFormulaInternalParser('D1', undefined)).toBe(null)
  })
})

function evaluateFormulaInternalParser(text: string, ctx?: EvaluationContext): unknown {
  return evaluateFormulaInternal(text as never, ctx as never)
}
void evaluateFormula

// ---------------------------------------------------------------------------
// ADDENDUM issues 5 + 6 (root-review confirmed, same bounded revision)
// ---------------------------------------------------------------------------

describe('REV1-5: canonical-error shared masters translate/evaluate as real errors', () => {
  for (const master of ['#N/A', '#DIV/0!', '#REF!', '#NAME?'] as const) {
    it(`master ${master} is not treated as an unparsable shared master`, () => {
      const d: XlsxDocument = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
        { ref: 'A1', row: 0, col: 0, value: master, hasCachedValue: true, formula: master, sharedFormula: { si: 0, ref: 'A1:A2' } },
        { ref: 'A2', row: 0, col: 1, value: null, sharedFormula: { si: 0 } },
      ] as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
      evaluateWorkbookFormulas(d as never, {})
      const follower = d.sheets[0].rows[0].cells[1] as XlsxCell
      expect(follower.value, `${master} follower resolves as the same error`).toBe(master)
      expect(follower.valueIsError ?? false, `follower ${master} tagged as an error`).toBe(true)
      const diags = (d as unknown as { diagnostics?: Array<{ feature: string }> }).diagnostics
      expect(diags?.some(x => x.feature === 'shared-formula') ?? false, 'no false shared-formula diagnostic').toBe(false)
    })
  }
})

describe('REV1-6: INT over nested aggregates keeps native 1 (no intermediate pre-rounding)', () => {
  for (const chain of [
    'INT(1.999999999999999/1)',
    'INT(SUM(1.999999999999999))',
    'INT(AVERAGE(1.999999999999999))',
    'INT(ABS(1.999999999999999))',
    'INT(ROUND(1.999999999999999,15))',
    'INT(1.999999999999999+0)',
    'INT(1.999999999999999*1)',
    'INT(1.999999999999999^1)',
  ]) {
    it(`${chain} = 1 (native)`, () => {
      expect(evaluateFormula(chain)).toBe(1)
    })
  }
  it('(1E16+1)-1E16 native float chain = 0', () => {
    expect(evaluateFormula('(1E16+1)-1E16')).toBe(0)
  })
  it('repeated evaluation + ACTUAL JSON round trip; changed root input forces recalculation', () => {
    const d: XlsxDocument = { sheets: [{ name: 'S', sourcePartPath: 's', rows: [{ index: 0, cells: [
      { ref: 'A1', row: 0, col: 0, value: null, formula: '=MOD(6.3,2.1)' },
      { ref: 'B1', row: 0, col: 1, value: null, formula: '=A1' },
    ] as XlsxCell[] }] }], images: [], drawingCoverage: [] } as never as XlsxDocument
    evaluateWorkbookFormulas(d as never, {})
    const first = (d.sheets[0].rows[0].cells[0] as XlsxCell).value
    expect(first).toBe(2.1)
    // ACTUAL JSON serialization round trip (not skipped): the serialized doc
    // must reproduce model values when re-evaluated.
    const roundtrip = JSON.parse(JSON.stringify(d)) as XlsxDocument
    evaluateWorkbookFormulas(roundtrip as never, {})
    expect((roundtrip.sheets[0].rows[0].cells[0] as XlsxCell).value).toBe(first)
    expect((roundtrip.sheets[0].rows[0].cells[1] as XlsxCell).value).toBe(first)
    // Force recalc with a REAL input change: A1 formula → MOD(6.3,2.2) ≈ 1.9 —
    // the forced pass must produce the CHANGED value, not replay caches.
    const changed = JSON.parse(JSON.stringify(d)) as XlsxDocument
    ;(changed.sheets[0].rows[0].cells[0] as XlsxCell).formula = '=MOD(6.3,2.2)'
    evaluateWorkbookFormulas(changed as never, { forceRecalc: true })
    const changedA1 = (changed.sheets[0].rows[0].cells[0] as XlsxCell)
    expect(changedA1.value, 'real root input change forces a changed calculated value').toBeCloseTo(1.9, 14)
    expect(changedA1.valueIsError ?? false).toBe(false)
    expect((changed.sheets[0].rows[0].cells[1] as XlsxCell).value).toBeCloseTo(1.9, 14)
    expect((changed.sheets[0].rows[0].cells[1] as XlsxCell).valueIsError ?? false).toBe(false)
  })
})
