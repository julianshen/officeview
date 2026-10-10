/**
 * Ported accepted C3 pure-helper retention/counterexample suites, imported
 * against the REAL shipping engine modules (no placeholder contracts).
 *
 * These are the QUALITY safety (Q1-Q3), SPEC counterexample (F1-F9) and R3
 * residual controls from the accepted pure prototype freeze
 * (ROOT-C3-PURE-ACCEPTANCE.json, c3-pure-review-r4). Provenance classes
 * [MEASURED]/[DOC-DERIVED]/[ABI-DERIVED]/[MODEL] are preserved verbatim.
 */
import { describe, expect, it } from "vitest"
import type { ResolvedSemantics } from "../src/xlsx/formula/types"
import { formulaError, isEvaluationError } from "../src/xlsx/formula/evaluator"
import { formatCommon } from "../src/xlsx/formula/format"
import {
  dateOf,
  dateTextParse,
  dateValueOf,
  daysOf,
  textOf,
  timeOf,
  timeValueOf,
  yearOf,
} from "../src/xlsx/formula/functions/datetime"

const SY1900 = "1900" as const
const SY1904 = "1904" as const
const tagged = formulaError("#N/A")

function sem(system: "1900" | "1904" = "1900"): ResolvedSemantics {
  return {
    dateSystem: system,
    unicode: { version: 2, source: "standalone-default" },
    locale: "en-US",
    timeZone: "UTC",
    epochNowMs: 1768478400000,
  }
}
function expectTagged(v: unknown, code: string) {
  expect(isEvaluationError(v) && v.code === code).toBe(true)
}
function expectKind(v: unknown, kind: string) {
  expect((v as { kind?: string }).kind).toBe(kind)
}
function txt(v: unknown): string {
  if (isEvaluationError(v)) return JSON.stringify(v)
  if (typeof v === "object" && v !== null && "kind" in v) {
    const r = v as { kind: string; text?: string }
    if (r.kind === "text") return r.text ?? ""
  }
  return JSON.stringify(v)
}

// ---- ported from c3-safety.test.ts ----
describe('Q1 — TIME finite huge arguments never escape NaN/Infinity', () => {
  // The exact Excel result at 1e308 is UNMEASURED (not asserted). The accepted
  // measured rule is the hour wrap (TIME(25,0,0)=1/24); the implementation
  // applies the same whole-day-shift-safe wrap arithmetic, so the contract
  // asserted here is exactly the safety property: the helper returns either a
  // FINITE number or a tagged #NUM! — a non-finite numeric success is
  // impossible.
  function expectFiniteOrTaggedNum(r: unknown) {
    if (typeof r === 'number') {
      expect(Number.isFinite(r)).toBe(true)
      return
    }
    expect(isEvaluationError(r) && r.code === '#NUM!').toBe(true)
  }
  it('timeOf(1e308, 0, 0) → finite number or tagged #NUM!, never NaN/Infinity [QUALITY Q1]', () => {
    expectFiniteOrTaggedNum(timeOf(1e308, 0, 0))
  })
  it('timeOf(0, 1e308, 0) → same contract [QUALITY Q1]', () => {
    expectFiniteOrTaggedNum(timeOf(0, 1e308, 0))
  })
  it('timeOf(0, 0, 1e308) → same contract [QUALITY Q1 symmetric]', () => {
    expectFiniteOrTaggedNum(timeOf(0, 0, 1e308))
  })
  it('wrap control retained: timeOf(25,0,0) ≈ 1/24 [MEASURED families id33]', () => {
    const r = timeOf(25, 0, 0)
    expect(typeof r === 'number' && Math.abs(r - 1 / 24) < 5e-13).toBe(true)
  })
})

describe('Q2 — percent/scaling overflow is a diagnosed numeric limit, not text success', () => {
  it("formatCommon(1e308,'0%') = unsupported numeric-limit (never 'Infinity%') [QUALITY Q2]", () => {
    const r = formatCommon(1e308, '0%', sem())
    expect(r.kind).toBe('unsupported')
    if (r.kind === 'unsupported') {
      expect(r.reason.toLowerCase()).toContain('numeric-limit')
    }
  })
  it("formatCommon(1e308,'0%%') = unsupported [QUALITY Q2]", () => {
    const r = formatCommon(1e308, '0%%', sem())
    expect(r.kind).toBe('unsupported')
  })
  it('1 with 155 percent tokens = unsupported [QUALITY Q2 documented-by-probe case]', () => {
    const r = formatCommon(1, '0' + '%'.repeat(155), sem())
    expect(r.kind).toBe('unsupported')
  })
  it('ordinary controls retained: 0.125+"0.0%" → "12.5%"; 1234567+"0,," → "1" [MEASURED]', () => {
    const p = formatCommon(0.125, '0.0%', sem())
    expect(p.kind === 'text' && (p as { text: string }).text).toBe('12.5%')
    const c = formatCommon(1234567, '0,,', sem())
    expect(c.kind === 'text' && (c as { text: string }).text).toBe('1')
  })
})

describe('Q3 — optional fraction suffix significance is linear (large-input outcome)', () => {
  it('formatCommon(0, "0." + "#".repeat(200000)) = "0." without timeout [OPERATION-COMPLEXITY PROOF: O(n^2) would exceed the 5s test budget by orders of magnitude; outcome asserted, not timing]', () => {
    const r = formatCommon(0, '0.' + '#'.repeat(200000), sem())
    expect(r.kind === 'text' && (r as { text: string }).text).toBe('0.')
  })
  it('mid-size sanity: n=3000 optional fraction zeros → "0." [EXTRA outcome control]', () => {
    const r = formatCommon(0, '0.' + '#'.repeat(3000), sem())
    expect(r.kind === 'text' && (r as { text: string }).text).toBe('0.')
  })
})


// ---- ported from c3-counterexamples.test.ts ----
describe('F1 — DATE backward normalization must borrow the PREVIOUS month', () => {
  it('[PROBE #1] DATE(2024,3,0)=45351 (Feb 29 2024) [DOC-DERIVED]', () => {
    expect(dateOf(2024, 3, 0, SY1900)).toBe(45351)
  })
  it('[PROBE #2] DATE(2024,3,-1)=45350 (Feb 28 2024) [DOC-DERIVED]', () => {
    expect(dateOf(2024, 3, -1, SY1900)).toBe(45350)
  })
  it('[PROBE #3] DATE(2024,3,0)/1904 = 43889 [DOC-DERIVED]', () => {
    expect(dateOf(2024, 3, 0, SY1904)).toBe(43889)
  })
  it('[PROBE #4] DATE(2024,3,-1)/1904 = 43888 [DOC-DERIVED]', () => {
    expect(dateOf(2024, 3, -1, SY1904)).toBe(43888)
  })
  it('[PROBE #5] DATE(1900,3,0)=60 — borrow through the parallel Feb-1900 fictitious slot [ABI-DERIVED from measured DATE(1900,2,29)=60]', () => {
    expect(dateOf(1900, 3, 0, SY1900)).toBe(60)
  })
  it('[EXTRA] adjacent unequal-month borrows: DATE(2024,5,-1)=45411 (Apr 29); DATE(2024,3,-2)=45349 (Feb 27); day-0 controls DATE(2024,3,0)=45351 [DOC-DERIVED]', () => {
    expect(dateOf(2024, 5, -1, SY1900)).toBe(45411)
    expect(dateOf(2024, 3, -2, SY1900)).toBe(45349)
  })
})

describe('F2 — DATEVALUE text years keep literal identity (1900..9999; DATE remap must not apply)', () => {
  it('[PROBE #6] DATEVALUE("1899-12-31","1900") = tagged #VALUE! (not 693962) [DOC-DERIVED primary text-year range]', () => {
    expectTagged(dateValueOf('1899-12-31', SY1900), '#VALUE!')
  })
  it('[EXTRA] DATE(1899,12,31)=693962 retained — the remap belongs to DATE only [MEASURED]', () => {
    expect(dateOf(1899, 12, 31, SY1900)).toBe(693962)
  })
})

describe('F3 — DAYS keeps raw numeric subtraction and applies numeric date guards', () => {
  it('[PROBE #7] DAYS(61.75, 59.25) = 2.5 raw [DOC-DERIVED primary numeric subtraction]', () => {
    expect(daysOf(61.75, 59.25, SY1900)).toBeCloseTo(2.5, 12)
  })
  it('[PROBE #8] DAYS(2958466, 1) = tagged #NUM! (numeric out-of-range) [DOC-DERIVED primary range guard]', () => {
    expectTagged(daysOf(2958466, 1, SY1900), '#NUM!')
  })
  it('[PROBE #9] DAYS(Infinity, 1) = tagged #NUM! (non-finite guard) [DOC-DERIVED]', () => {
    expectTagged(daysOf(Infinity, 1, SY1900), '#NUM!')
  })
  it('[EXTRA] 1904 mirror: DAYS(2957004, 1) = #NUM! [DOC-DERIVED symmetric gate]', () => {
    expectTagged(daysOf(2957004, 1, SY1904), '#NUM!')
  })
})

describe('F4 — tagged EvaluationError inputs pass through datetime helpers unchanged', () => {
  it('[PROBE #18] dateOf(tagged #N/A, 1, 1) = #N/A [ABI ElementValue preservation]', () => {
    const r = dateOf(tagged, 1, 1, SY1900)
    expectKind(r, 'formula-error')
    expectTagged(r, '#N/A')
  })
  it('[PROBE #19] yearOf(tagged #N/A) = #N/A [ABI]', () => {
    expectTagged(yearOf(tagged, SY1900), '#N/A')
  })
  it('[PROBE #20] timeOf(1, 2, tagged) = #N/A [ABI]', () => {
    const r = timeOf(1, 2, tagged)
    expectKind(r, 'formula-error')
    expectTagged(r, '#N/A')
  })
  it('[PROBE #21] dateValueOf(tagged) = #N/A [ABI]', () => {
    expectTagged(dateValueOf(tagged as unknown as string, SY1900), '#N/A')
  })
  it('[PROBE #22] timeValueOf(tagged) = #N/A [ABI]', () => {
    expectTagged(timeValueOf(tagged as unknown as string), '#N/A')
  })
  it('[PROBE #23] textOf(tagged value, format) = #N/A AND textOf(value, tagged format) = #N/A [ABI]', () => {
    const valueTag = textOf(tagged, 'yyyy-mm-dd', sem())
    expectKind(valueTag, 'formula-error')
    expectTagged(valueTag, '#N/A')
    const formatTag = textOf(1, tagged as unknown as string, sem())
    expectKind(formatTag, 'formula-error')
    expectTagged(formatTag, '#N/A')
  })
  it('[EXTRA] quoted "#N/A" text stays TEXT, and tagged format arg + null value in textOf stays tagged [ABI/parser boundary counterexample]', () => {
    const quoted = textOf('#N/A', '@', sem())
    expect(quoted.kind === 'text' && (quoted as { text: string }).text).toBe('#N/A')
    expectTagged(textOf(null as never, tagged as unknown as string, sem()), '#N/A')
  })
})

describe('F5 — sub-unit rounding must use the dropped digit', () => {
  it('[PROBE #13] format(0.6,"0") = "1" [DOC-DERIVED]', () => {
    expect(txt(formatCommon(0.6, '0', sem()))).toBe('1')
  })
  it('[EXTRA] half / negative sub-unit / threshold controls [DOC-DERIVED]', () => {
    expect(txt(formatCommon(0.5, '0', sem()))).toBe('1')
    expect(txt(formatCommon(1.4, '0', sem()))).toBe('1')
    expect(txt(formatCommon(-0.6, '0', sem()))).toBe('-1')
    expect(txt(formatCommon(-0.5, '0', sem()))).toBe('-1')
  })
})

describe('F6 — optional # placeholders must not force zeros', () => {
  it('[PROBE #10] format(8.9,"#.##") = "8.9" [DOC-DERIVED]', () => {
    expect(txt(formatCommon(8.9, '#.##', sem()))).toBe('8.9')
  })
  it('[PROBE #11] format(12,"#.0#") = "12.0" (required 0 holds its digit) [DOC-DERIVED]', () => {
    expect(txt(formatCommon(12, '#.0#', sem()))).toBe('12.0')
  })
  it('[PROBE #12] format(0.47,"#.##") = ".47" (optional integer part, no leading 0) [DOC-DERIVED]', () => {
    expect(txt(formatCommon(0.47, '#.##', sem()))).toBe('.47')
  })
  it('[EXTRA] significant-zero integrity retained [DOC-DERIVED]', () => {
    expect(txt(formatCommon(105, '#,##0.00', sem()))).toBe('105.00')
    expect(txt(formatCommon(100, '0.0#', sem()))).toBe('100.0')
  })
})

describe('F7 — grouping from the right, no leading comma, across overflow', () => {
  it('[PROBE #14] format(12,"#,##0") = "12" [DOC-DERIVED]', () => {
    expect(txt(formatCommon(12, '#,##0', sem()))).toBe('12')
  })
  it('[PROBE #15] format(1234567,"#,##0") = "1,234,567" [DOC-DERIVED]', () => {
    expect(txt(formatCommon(1234567, '#,##0', sem()))).toBe('1,234,567')
  })
  it('[EXTRA] comma boundary: 123 → "123"; 1234 → "1,234"; 12345 → "12,345" [DOC-DERIVED]', () => {
    expect(txt(formatCommon(123, '#,##0', sem()))).toBe('123')
    expect(txt(formatCommon(1234, '#,##0', sem()))).toBe('1,234')
    expect(txt(formatCommon(12345, '#,##0', sem()))).toBe('12,345')
  })
})

describe('F8 — fourth text section renders literals and can hide text', () => {
  it('[PROBE #16] "bob" with 0;0;0;"hello "@ = "hello bob" [DOC-DERIVED]', () => {
    expect(txt(textOf('bob', '0;0;0;"hello "@', sem()))).toBe('hello bob')
  })
  it('[PROBE #17] "bob" with 0;0;0; (empty 4th) = "" (hidden) [DOC-DERIVED]', () => {
    expect(txt(textOf('bob', '0;0;0;', sem()))).toBe('')
  })
  it('[EXTRA] no fourth section (3-section format) → text falls through [MODEL; r1 scope unchanged]', () => {
    expect(txt(textOf('bob', '0.00;(0.00);0', sem()))).toBe('bob')
  })
})

describe('F9 — unsupported text-date locale must be a distinguishable capability gate', () => {
  it('[PROBE #23 of unknownObservations, capability half] 29/02/2024 in en-GB: status unsupported-locale, NOT a bare counterfeit computed #VALUE! [REVIEW F9]', () => {
    const r = dateTextParse('29/02/2024', SY1900, 'en-GB')
    expect(r.status).toBe('unsupported-locale')
    if (r.status === 'unsupported-locale') {
      expect(r.locale).toBe('en-GB')
      expectTagged(r.error, '#VALUE!')
    }
  })
  it('[REVIEW F9 preflight] parsed vs invalid-date statuses are distinguishable [CAPABILITY]', () => {
    const ok = dateTextParse('2024-02-29', SY1900, 'en-US')
    expect(ok.status === 'parsed' && ok.serial === 45351).toBe(true)
    const bad = dateTextParse('1899-12-31', SY1900, 'en-US')
    expect(bad.status).toBe('invalid-date')
    if (bad.status === 'invalid-date') expectTagged(bad.error, '#VALUE!')
  })
  it('[PENDING-GATE] the public dateValueOf still returns #VALUE! for the unsupported locale; the owner adapter uses dateTextParse to distinguish [MODEL boundary, no integration claimed]', () => {
    expectTagged(dateValueOf('29/02/2024', SY1900, 'en-GB'), '#VALUE!')
  })
})


// ---- ported from c3-residual.test.ts ----
describe('R3 residual — integer optional-# preserves significant zeros', () => {
  it('[RESIDUAL #1] format(1000,"#,###") = "1,000" [DOC-DERIVED integer significance]', () => {
    expect(txt(formatCommon(1000, '#,###', sem()))).toBe('1,000')
  })
  it('[RESIDUAL #2] format(12000,"#,###") = "12,000" [DOC-DERIVED — exact documented Microsoft example]', () => {
    expect(txt(formatCommon(12000, '#,###', sem()))).toBe('12,000')
  })
  it('[RESIDUAL #3] format(10,"##") = "10" [DOC-DERIVED same rule]', () => {
    expect(txt(formatCommon(10, '##', sem()))).toBe('10')
  })
})
