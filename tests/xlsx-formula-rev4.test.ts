/**
 * CORE-REVISION4 gates (RED first) — root NATIVE-COERCION-MORE-FORMULAS.json:
 * percent scaling fuses into the DECIMAL EXPONENT before floating parse, so
 * "1e309%" = 1e307 (finite, native). Plain overflow "1e309"/"-1e309" stay
 * #VALUE!; true SCALED overflow ("1e400%") stays #VALUE!; malformed text
 * ("1,309%x") stays #VALUE!. No limit adjustments to pass bad tests.
 */
import { describe, expect, it } from 'vitest'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'

describe('REV4: percent exponent fused before float parse (native counterexample)', () => {
  it('"1e309%"+1 = 1E+307 (native), NOT #VALUE!', () => {
    expect(evaluateFormula('"1e309%"+1')).toBe(1e307)
  })
  it('plain overflow rejects with #VALUE!', () => {
    expect(evaluateFormula('"1e309"+1')).toBe('#VALUE!')
    expect(evaluateFormula('"-1e309"+1')).toBe('#VALUE!')
  })
  it('true SCALED overflow rejects with #VALUE!', () => {
    expect(evaluateFormula('"1e400%"+1')).toBe('#VALUE!')
  })
  it('malformed text still rejects with #VALUE!', () => {
    expect(evaluateFormula('"1,309%x"+1')).toBe('#VALUE!')
  })
  it('ordinary percent scales unchanged', () => {
    expect(evaluateFormula('"10%"+1')).toBe(1.1)
    expect(evaluateFormula('COUNT("10%")')).toBe(1)
    expect(evaluateFormula('SUM("10%")')).toBe(0.1)
  })
})
