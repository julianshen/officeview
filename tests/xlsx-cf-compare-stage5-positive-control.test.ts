// @vitest-environment node
/**
 * cf-compare on an unmeasured Stage 5 case compares a declared positive control on the token the case declares for it,
 * not on the `matched: false` placeholder `s5Case` seeds for every cell until a capture replaces it.
 * stage5-cross-sheet-missing B5 (1 under `cellIs greaterThan 0`) failed "Control cell B5 expected matched: false, but CF
 * formatting applied" in both modes although production paints it correctly.
 */
import { describe, expect, it } from 'vitest'
import { stage5SeedPositiveControls, verifyCellExpectations, type ActualCellOutcome } from '../scripts/cf-compare'
import type { CfOracleCase } from '../scripts/cf-refs/cases'

const caseOf = (status: string, positiveControls: Record<string, string>): CfOracleCase => ({
  id: 'synthetic-s5', stage: 5, status,
  markers: [],
  formulaConfigurations: [{ id: 'cfg', capabilityId: 'formula:x', observable: 'fill', refs: ['B2'], hypotheses: {}, positiveControls }],
  expected: { cells: [{ ref: 'B5', controlRef: 'E5', matched: false, fillSource: 'unverified' }, { ref: 'E5', controlRef: 'B5', matched: false, fillSource: 'unverified' }] },
}) as never

const painted = (fillColor: string, strike = false): ActualCellOutcome => ({ applied: true, style: { fillColor, ...(strike ? { strike } : {}) } } as never)
const blank: ActualCellOutcome = { applied: false }
const at = (map: Record<string, ActualCellOutcome>) => (row: number, col: number) => map[`${row},${col}`] ?? blank

describe('cf-compare: unmeasured Stage 5 positive controls follow their declared token', () => {
  it('collects only main-table controls, upper-cased', () => {
    const c = caseOf('recipe-pending', { b5: 'FFFFC7CE', 'post:B6': 'FFFFC7CE' })
    expect([...stage5SeedPositiveControls(c)]).toEqual([['B5', 'FFFFC7CE']])
  })

  it('a painted red control passes (cross-sheet-missing B5)', () => {
    const c = caseOf('recipe-pending', { B5: 'FFFFC7CE' })
    expect(verifyCellExpectations(c, at({ '4,1': painted('ffffc7ce') }), undefined, new Map())).toBeUndefined()
  })

  it('an unpainted control fails', () => {
    const c = caseOf('recipe-pending', { B5: 'FFFFC7CE' })
    expect(verifyCellExpectations(c, at({}), undefined, new Map())).toMatch(/Positive control B5: expected FFFFC7CE, but production shows none/)
  })

  it('a withheld control fails', () => {
    const c = caseOf('recipe-pending', { B5: 'FFFFC7CE' })
    expect(verifyCellExpectations(c, at({ '4,1': { applied: false, unresolved: true } }), undefined, new Map())).toMatch(/shows withheld/)
  })

  it('a control painted the wrong colour fails, and a strike control needs the strike', () => {
    expect(verifyCellExpectations(caseOf('recipe-pending', { B5: 'FFFFC7CE' }), at({ '4,1': painted('FFC6EFCE') }), undefined, new Map())).toMatch(/shows FFC6EFCE/)
    expect(verifyCellExpectations(caseOf('recipe-pending', { B5: 'none+strike' }), at({ '4,1': { applied: true, style: { strike: true } } as never }), undefined, new Map())).toBeUndefined()
    expect(verifyCellExpectations(caseOf('recipe-pending', { B5: 'none+strike' }), at({ '4,1': painted('FFFFC7CE') }), undefined, new Map())).toMatch(/expected none\+strike/)
  })

  it('the same-content unformatted control keeps the seed check', () => {
    const c = caseOf('recipe-pending', { B5: 'FFFFC7CE' })
    expect(verifyCellExpectations(c, at({ '4,1': painted('FFFFC7CE'), '4,4': painted('FFFFC7CE') }), undefined, new Map())).toMatch(/Control cell E5 expected matched: false/)
  })
})

describe('cf-compare: stage5-cross-sheet-missing no longer fails on its positive control', async () => {
  const { readFileSync } = await import('node:fs')
  const { resolve } = await import('node:path')
  const { compareCfCase } = await import('../scripts/cf-compare')
  const root = resolve(__dirname, '..')
  const manifest = JSON.parse(readFileSync(resolve(root, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as { cases: CfOracleCase[] }
  for (const mode of ['interactive', 'paged'] as const) {
    it(`[${mode}] is recipe-pending, not fail`, async () => {
      const c = manifest.cases.find(e => e.id === 'stage5-cross-sheet-missing')!
      expect(c.status).toBe('recipe-pending')
      expect(c.expected.cells.find(e => e.ref === 'B5')).toMatchObject({ matched: false })
      const res = await compareCfCase(c, mode, { rootDir: root })
      expect(res.outcome, res.reason).toBe('recipe-pending')
    }, 300_000)
  }
})
