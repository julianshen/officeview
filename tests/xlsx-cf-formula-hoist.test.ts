// @vitest-environment node
/**
 * Stage 5 Task 34, Ruling S5-HOIST: absolute-range aggregates are hoisted out of the per-cell call into a chunked scan
 * with engine-ordered reducers, and the COUNTIF frequency map reproduces the engine's plain-equality arm only.
 */
import { describe, expect, it } from 'vitest'
// functions.ts first: importing cond.ts before it evaluates FUNCTIONS before CONDITIONAL_FUNCTIONS exists (an engine
// import cycle, functions.ts <-> evaluator.ts <-> cond.ts).
import '../src/xlsx/formula/functions'
import { CONDITIONAL_MAX_CRITERION_LENGTH } from '../src/xlsx/formula/functions/cond'
import { DEFAULT_MAX_DP_CELLS, matchWildcard } from '../src/xlsx/formula/functions/wildcard'
import { engineEval } from './helpers/cf-formula'

describe('S5-HOIST: the engine limits the hoist mirrors are the engine\'s own named constants', () => {
  it('CONDITIONAL_MAX_CRITERION_LENGTH and DEFAULT_MAX_DP_CELLS drive the engine refusals unchanged', () => {
    expect(CONDITIONAL_MAX_CRITERION_LENGTH).toBe(255) // cond.ts criterionOf, `value.length > 255` before Task 34
    expect(DEFAULT_MAX_DP_CELLS).toBe(4_000_000) // wildcard.ts, unchanged
    const at = { row: 0, col: 2 }
    const cells = { A1: 'x'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH), B1: 'x'.repeat(CONDITIONAL_MAX_CRITERION_LENGTH + 1) }
    expect(engineEval('COUNTIF(A1,A1)', cells, at).features).toEqual([])
    expect(engineEval('COUNTIF(A1,B1)', cells, at).features).toEqual(['conditional-long-criterion-unverified'])
    // (tokens + 1) x (length + 1) at the limit matches; one DP cell over refuses.
    const n = DEFAULT_MAX_DP_CELLS / 2 - 1 // pattern 'a' (1 token): 2 x (n + 1) = DEFAULT_MAX_DP_CELLS
    expect(matchWildcard('a', 'b'.repeat(n)).status).toBe('not-matched')
    expect(matchWildcard('a', 'b'.repeat(n + 1)).status).toBe('resource-limit')
  })
})
