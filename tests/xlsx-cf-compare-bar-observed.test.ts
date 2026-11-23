// @vitest-environment node
/**
 * cf-compare on a measured Stage 3 bar cell compares whether production DRAWS a bar against Excel's observed `barApplied`
 * (the PNG measurement `capture.ts` writes), not against the seed `matched` flag, which marks every cell of the rule's sqref.
 * stage3-bar-participation B4 (TRUE: no bar) and stage3-bar-composition F2 (halted by a matching stopIfTrue ahead) failed
 * "expected matched: true, but no CF formatting applied" although Excel draws no bar there either.
 */
import { describe, expect, it } from 'vitest'
import { barDrawnOf, verifyCellExpectations, type ActualCellOutcome } from '../scripts/cf-compare'
import type { CfOracleCase } from '../scripts/cf-refs/cases'

const caseOf = (status: string, cells: Array<{ ref: string; barApplied?: boolean }>): CfOracleCase => ({
  id: 'synthetic-bar', stage: 3, status,
  markers: [{ capabilityId: 'bar:x', ruleType: 'dataBar', sqref: 'B2:B5', priority: 1 }],
  expected: { cells: cells.map(c => ({ ...c, matched: true, fillSource: 'unverified' })) },
}) as never

const at = (map: Record<string, ActualCellOutcome>) => (row: number, col: number) => map[`${row},${col}`]!

describe('cf-compare: measured Stage 3 bar cells follow Excel\'s observed bar', () => {
  it('barDrawnOf: a zero-length or absent bar is not drawn', () => {
    expect(barDrawnOf(undefined)).toBe(false)
    expect(barDrawnOf(null)).toBe(false)
    expect(barDrawnOf({ startFraction: 0, endFraction: 0 })).toBe(false)
    expect(barDrawnOf({ startFraction: 0, endFraction: 0.25 })).toBe(true)
  })

  it('Excel draws no bar and production draws none: passes (B4 TRUE, F2 halted)', () => {
    const c = caseOf('measured', [{ ref: 'B2', barApplied: false }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: false, barDrawn: false } }))).toBeUndefined()
  })

  it('Excel draws no bar but production draws one: fails (newly caught)', () => {
    const c = caseOf('measured', [{ ref: 'B2', barApplied: false }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: true, barDrawn: true } }))).toMatch(/Excel draws no bar, but production draws a bar/)
  })

  it('Excel draws a bar but production draws none: still fails', () => {
    const c = caseOf('measured', [{ ref: 'B2', barApplied: true }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: false, barDrawn: false } }))).toMatch(/Excel draws a bar, but production draws no bar/)
  })

  it('a zero-length bar at the column minimum (descriptor present, nothing drawn) agrees with Excel\'s no bar', () => {
    const c = caseOf('measured', [{ ref: 'B2', barApplied: false }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: true, barDrawn: false } }))).toBeUndefined()
  })

  it('a withheld cell stays skipped (fail closed), whatever Excel drew', () => {
    const c = caseOf('measured', [{ ref: 'B2', barApplied: true }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: false, unresolved: true, barDrawn: false } }))).toBeUndefined()
  })

  it('a recipe-pending case keeps the seed check (its barApplied is a seed, not an observation)', () => {
    const c = caseOf('recipe-pending', [{ ref: 'B2', barApplied: false }])
    expect(verifyCellExpectations(c, at({ '1,1': { applied: false, barDrawn: false } }))).toMatch(/expected matched: true, but no CF formatting applied/)
  })
})

describe('cf-compare: the two Stage 3 cases that failed on main now pass on the observed bar', async () => {
  const { readFileSync } = await import('node:fs')
  const { resolve } = await import('node:path')
  const { compareCfCase } = await import('../scripts/cf-compare')
  const root = resolve(__dirname, '..')
  const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: CfOracleCase[] }
  for (const id of ['stage3-bar-participation', 'stage3-bar-composition']) {
    for (const mode of ['interactive', 'paged'] as const) {
      it(`${id} [${mode}] passes`, async () => {
        const c = manifest.cases.find(e => e.id === id)!
        // B4 (TRUE) and F2 (halted) are cells Excel draws no bar at; the seed still marks them matched.
        const probe = id === 'stage3-bar-participation' ? 'B4' : 'F2'
        expect(c.expected.cells.find(e => e.ref === probe)).toMatchObject({ matched: true, barApplied: false })
        const res = await compareCfCase(c, mode, { rootDir: root })
        expect(res.outcome, res.reason).toBe('pass')
      }, 300_000)
    }
  }
})
