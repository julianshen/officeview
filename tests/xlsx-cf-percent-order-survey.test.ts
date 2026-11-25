// @vitest-environment node
/**
 * The percent-order follow-up (brief cf-percent-order), Phase 1 task 1: the survey of every captured percent / percentile
 * bound and every captured cell whose outcome depends on the threshold's arithmetic order
 * (scripts/cf-refs/percent-order-survey.ts). It pins the numbers the phase-1 report gives: 287 in-domain bounds, 7 that
 * differ between `lerp-first` and `multiply-first`, and 5 order-dependent cells (E2..E6 of stage2-scale-formula-bounds,
 * withheld today), none of which contradicts `lerp-first` on its own (it matches Excel there with `fifteen-digit`). It
 * checks the harness, not Excel, and registers nothing.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { percentOrderSurveyOf, surveyPercent, surveyPercentile, type SurveyBound, type SurveyCell } from '../scripts/cf-refs/percent-order-survey'

const HEAVY = 300_000
let bounds: SurveyBound[] = []
let excluded: SurveyBound[] = []
let cells: SurveyCell[] = []
beforeAll(async () => { ({ bounds, excluded, cells } = await percentOrderSurveyOf()) }, HEAVY)
const differs = (r: SurveyBound): boolean => r.scopes.all.differs || r.scopes.visible.differs

describe('the percent-order survey', () => {
  it('the two orders, restated', () => {
    expect(surveyPercent(0, 300, 33.333333333333336, 'lerp-first')).toBe(100.00000000000001)
    expect(surveyPercent(0, 300, 33.333333333333336, 'multiply-first')).toBe(100)
    expect(surveyPercent(10, 100, 67, 'lerp-first')).toBe(70.30000000000001)
    expect(surveyPercent(10, 100, 67, 'multiply-first')).toBe(70.3)
    expect(surveyPercentile([5, 15, 30, 45, 60, 75, 90, 100], 20, 'lerp-first')).toBe(21.000000000000007)
    expect(surveyPercentile([5, 15, 30, 45, 60, 75, 90, 100], 20, 'multiply-first')).toBe(21)
  })

  it('287 in-domain bounds over 25 cases; 8 out-of-domain bounds excluded (production refuses the shape)', () => {
    expect(bounds.length).toBe(287)
    expect(new Set(bounds.map(r => r.caseId)).size).toBe(25)
    expect(excluded.map(r => `${r.caseId} ${r.sqref} ${r.type} ${r.text}`).sort()).toHaveLength(8)
    expect(excluded.every(r => r.p < 0 || r.p > 100)).toBe(true)
    const t = (rule: string, type: string): [number, number] => [bounds.filter(r => r.rule === rule && r.type === type).length, bounds.filter(r => r.rule === rule && r.type === type && differs(r)).length]
    expect([t('colorScale', 'percent'), t('colorScale', 'percentile'), t('dataBar', 'percent'), t('dataBar', 'percentile'), t('iconSet', 'percent'), t('iconSet', 'percentile')])
      .toEqual([[13, 1], [16, 2], [4, 0], [4, 0], [247, 4], [3, 0]])
  })

  it('7 bounds differ between the orders', () => {
    expect(bounds.filter(differs).map(r => `${r.caseId} ${r.sqref} ${r.index}`)).toEqual([
      'stage2-scale-formula-bounds E2:E10 0',
      'stage4-icon-scope B2:B4 D2:D4 2',
      'stage4-icon-scope F2:F6 2',
      'stage4-icon-scope G2:G6 2',
      'stage4-icon-rtl B2:B4 1',
      'stage5-cfvo-formulas G2:G9 0',
      'stage5-cfvo-formulas T2:T9 0',
    ])
  })

  it('only E2..E6 of stage2-scale-formula-bounds change outcome; Excel matches multiply-first there, and lerp-first with fifteen-digit, so no cell contradicts lerp-first', () => {
    expect(cells.map(c => `${c.caseId}:${c.ref}`)).toEqual(['E2', 'E3', 'E4', 'E5', 'E6'].map(r => `stage2-scale-formula-bounds:${r}`))
    for (const c of cells) {
      expect(c.excel).toBe(c.multiply)
      expect(c.excel).not.toBe(c.lerp)
      expect(c.grid!['lerp-first fifteen-digit']).toBe(c.excel)
      expect(c.grid!['multiply-first as-literal']).toBe(c.excel)
      expect(c.grid!['multiply-first fifteen-digit']).toBe(c.excel)
      expect(c.grid!['lerp-first as-literal']).not.toBe(c.excel)
    }
  })
}, HEAVY)
