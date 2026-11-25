// @vitest-environment node
/**
 * The formula-bounds follow-up (brief cf-formula-bounds), Phase 1 task 1: the census of every captured colour-scale cell
 * whose segment is bounded by a formula-resolved cfvo (scripts/cf-refs/formula-bounds-census.ts). It pins the numbers the
 * phase-1 report gives: no painted cell disagrees with `as-literal`, every captured formula result equals its 15-digit
 * rounding, so no captured cell tells `as-literal` from `fifteen-digit`. It checks the harness, not Excel, and registers nothing.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import { fifteenDigit, formulaBoundsCensusOf, type FormulaBoundsRow } from '../scripts/cf-refs/formula-bounds-census'

const HEAVY = 300_000
let census: Awaited<ReturnType<typeof formulaBoundsCensusOf>>
let rows: FormulaBoundsRow[] = []
beforeAll(async () => {
  census = await formulaBoundsCensusOf()
  rows = census.rows
}, HEAVY)
const table = (key: (r: FormulaBoundsRow) => string): string[] => {
  const m = new Map<string, number>()
  for (const r of rows) m.set(key(r), (m.get(key(r)) ?? 0) + 1)
  return [...m].sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, n]) => `${k} ${n}`)
}

describe('the formula-bounds census', () => {
  it('167 formula-bounded cells of 1780 captured segment cells, no harness fault', () => {
    expect(census.collection.cells.length).toBe(1780)
    expect(census.collection.faults).toEqual([])
    expect(rows.length).toBe(167)
    expect(table(r => r.caseId)).toEqual(['stage5-cfvo-formulas 56', 'stage5-cfvo-nonnumeric-v2 15', 'stage5-cfvo-relative-a1 20', 'stage5-cfvo-scope 24', 'stage5-cfvo-scope-v2 42', 'stage5-formula-authored 2', 'stage5-x14-scale 8'])
  })

  it('by shape and bounds: only E4 / E5 of stage5-cfvo-nonnumeric-v2 are in a withheld shape (the upper integer class)', () => {
    expect(table(r => r.shape)).toEqual(['clear 87', 'endpoint 78', 'upper-integer 2'])
    expect(table(r => `${r.colours}c seg${r.segment} ${r.bounds}`)).toEqual([
      '2c seg0 min/num:formula 79', '2c seg0 num:formula/max 54', '2c seg0 num:formula/min 12', '2c seg0 num:formula/num:formula 1',
      '2c seg0 percent:formula/max 8', '2c seg0 percentile:formula/max 8', '3c seg0 min/num:formula 2', '3c seg1 num:formula/max 3',
    ])
    expect(rows.filter(r => r.production === undefined).map(r => `${r.caseId}:${r.ref} ${r.shape}`)).toEqual(['stage5-cfvo-nonnumeric-v2:E4 upper-integer', 'stage5-cfvo-nonnumeric-v2:E5 upper-integer'])
  })

  it('every formula result resolves and equals its 15-digit rounding (0, 1, 4, 7, 11, 15, 20, 50)', () => {
    const results = rows.flatMap(r => r.formula)
    expect(results.every(b => b.result !== undefined && b.differs === false)).toBe(true)
    expect([...new Set(results.map(b => b.result))].sort((a, b) => a! - b!)).toEqual([0, 1, 4, 7, 11, 15, 20, 50])
    expect(fifteenDigit(0.1 + 0.2)).toBe(0.3)
    expect(fifteenDigit(100 / 3)).toBe(33.3333333333333)
  })

  it('no painted cell disagrees with as-literal; both readings equal Excel at every cell; none discriminates', () => {
    expect(rows.filter(r => r.production !== undefined && r.production !== r.excel)).toEqual([])
    expect(rows.filter(r => r.asLiteral !== r.excel).map(r => `${r.caseId}:${r.ref}`)).toEqual([])
    expect(rows.filter(r => r.fifteenDigit !== r.excel).map(r => `${r.caseId}:${r.ref}`)).toEqual([])
    expect(rows.filter(r => r.discriminating)).toEqual([])
  })
}, HEAVY)

describe('after the capture of stage2-scale-formula-bounds (d2660d46): the census on every capture (recipe §14.7)', () => {
  let all: FormulaBoundsRow[] = []
  beforeAll(async () => { all = (await formulaBoundsCensusOf(undefined, undefined, { include: 'all' })).rows }, HEAVY)
  it('248 formula-bounded cells (81 of the capture); as-literal equals Excel at 243, fifteen-digit at 210; 43 discriminate', () => {
    expect([all.length, all.filter(r => r.caseId === 'stage2-scale-formula-bounds').length]).toEqual([248, 81])
    expect(all.filter(r => r.asLiteral !== r.excel).map(r => `${r.caseId}:${r.ref}`)).toEqual(['E2', 'E3', 'E4', 'E5', 'E6'].map(x => `stage2-scale-formula-bounds:${x}`))
    expect(all.filter(r => r.fifteenDigit === r.excel).length).toBe(210)
    expect(all.filter(r => r.discriminating).length).toBe(43)
    expect(all.filter(r => r.production !== undefined && r.production !== r.excel)).toEqual([])
  })
}, HEAVY)
