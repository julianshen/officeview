// @vitest-environment node
/**
 * The upper-segment follow-up (brief cf-upper-segment), Phase 1a: the census of every captured 3-colour UPPER-segment
 * interior cell and the candidate search over it (scripts/cf-refs/upper-segment-candidates.ts). Every capture is included,
 * stage2-scale-literal-lower (a3f85309) among them. This pins the numbers the phase-1 report gives; it checks the harness,
 * not Excel, and registers nothing.
 */
import { beforeAll, describe, expect, it } from 'vitest'
import {
  channelTable,
  countBy,
  decisionByDouble,
  evaluateCandidates,
  probeGrid,
  survivorClasses,
  upperSegmentCandidates,
  upperSegmentCensusOf,
  type CandidateResult,
  type CensusRow,
} from '../scripts/cf-refs/upper-segment-candidates'

const HEAVY = 300_000
let census: Awaited<ReturnType<typeof upperSegmentCensusOf>>
let rows: CensusRow[] = []
let results: CandidateResult[] = []
beforeAll(async () => {
  census = await upperSegmentCensusOf()
  rows = census.rows
  results = evaluateCandidates(rows)
}, HEAVY)
const table = (key: (r: CensusRow) => string): string[] => countBy(rows, key).map(t => `${t.key} ${t.cells}`)

describe('the upper-segment census', () => {
  it('1589 captured colour-scale cells, no harness fault, none dropped under a 3-colour scale by a gate; 302 upper-segment interior cells', () => {
    expect(census.collection.cells.length).toBe(1589)
    expect(census.collection.faults).toEqual([])
    expect(census.collection.excluded.filter(e => e.group === 'b' && e.threeColour).length).toBe(0)
    expect(rows.length).toBe(302)
  })

  it('by class, and by class × case (the documented tallies: 19 + T4/V3, 79 moved of stage2-scale-upper-whole, the 16 literal-lower probes 11 / 2 / 3)', () => {
    expect(table(r => r.cls)).toEqual(['minus-one 112', 'raw-floor 43', 'truncation 147'])
    expect(table(r => `${r.cls} ${r.caseId}`)).toEqual([
      'minus-one stage2-scale-literal-lower 11', 'minus-one stage2-scale-num-segment 19', 'minus-one stage2-scale-scope-disjoint 2', 'minus-one stage2-scale-upper-whole 75', 'minus-one stage2-scale-upper-window 5',
      'raw-floor stage2-scale-literal-lower 2', 'raw-floor stage2-scale-raw-floor 7', 'raw-floor stage2-scale-upper-whole 4', 'raw-floor stage2-scale-upper-window 30',
      'truncation stage2-scale-interpolation 3', 'truncation stage2-scale-literal-lower 3', 'truncation stage2-scale-num-segment 47', 'truncation stage2-scale-raw-floor 15', 'truncation stage2-scale-scope-disjoint 7',
      'truncation stage2-scale-thresholds 4', 'truncation stage2-scale-upper-whole 37', 'truncation stage2-scale-upper-window 27', 'truncation stage5-cfvo-nonnumeric-v2 4',
    ])
    expect(rows.filter(r => r.caseId === 'stage2-scale-literal-lower' && r.cls !== 'minus-one').map(r => `${r.ref} ${r.cls}`).sort()).toEqual(['D11 truncation', 'D12 truncation', 'G8 raw-floor', 'H10 truncation', 'H9 raw-floor'])
    expect(rows.filter(r => r.caseId === 'stage2-scale-scope-disjoint' && r.cls === 'minus-one').map(r => r.ref).sort()).toEqual(['T4', 'V3'])
  })

  it('by class × bounds and × palette kind', () => {
    expect(table(r => `${r.cls} ${r.bounds}`)).toEqual([
      'minus-one num/num 104', 'minus-one percent/max 1', 'minus-one percentile/max 7',
      'raw-floor num/max 1', 'raw-floor num/num 21', 'raw-floor percent/max 8', 'raw-floor percentile/max 11', 'raw-floor percentile/num 2',
      'truncation num/max 6', 'truncation num/num 92', 'truncation num:formula/max 2', 'truncation percent/max 8', 'truncation percentile/max 30', 'truncation percentile/num 9',
    ])
    expect(table(r => `${r.cls} ${r.kind}`)).toEqual(['minus-one distinct 27', 'minus-one p-line 85', 'raw-floor distinct 2', 'raw-floor p-line 41', 'truncation distinct 34', 'truncation p-line 113'])
  })

  it('per channel: a clear product never moves; integer, below and above products move one step or not; one decision per f double', () => {
    expect(channelTable(rows).map(t => `${t.key} ${t.channels}`)).toEqual(['above -1 27', 'above 0 50', 'below -1 127', 'below 0 23', 'clear 0 223', 'integer -1 296', 'integer 0 155'])
    const dec = decisionByDouble(rows)
    expect(dec.keys.size).toBe(75)
    expect([...dec.keys.values()].filter(v => v === 'moved').length).toBe(40)
    expect(dec.mixedCells).toEqual([])
    expect(dec.inconsistent).toEqual([])
  })
})

describe('the candidate search', () => {
  const pool = (p: CandidateResult['pool']): CandidateResult[] => results.filter(r => r.pool === p)
  const best = (p: CandidateResult['pool']): number => Math.min(...pool(p).map(r => r.mismatches.length))

  it('the earlier pools still have no survivor with every capture: truncation misses 155, step-plus-one 156, the first search 42 at best, the extended and window searches 34', () => {
    expect(pool('baseline').map(r => r.mismatches.length)).toEqual([155])
    expect(pool('num-segment').length).toBe(5803)
    expect(pool('upper-window').length).toBe(48603)
    expect(pool('literal-lower').map(r => r.mismatches.length)).toEqual([156])
    expect(best('num-segment')).toBe(42)
    expect(best('upper-window')).toBe(34)
  }, HEAVY)

  it('the position family: 210 candidates, 7 survivors, all one class over the probe grid: (1 + f) − 1 then raw floor', () => {
    expect(pool('position').length).toBe(210)
    const survivors = results.filter(r => r.mismatches.length === 0).map(r => r.id)
    expect(survivors).toEqual([
      'position/index-plus-f/floor', 'position/f-plus-index/floor', 'position/half-plus-half-f/floor', 'position/half-plus-half-f-div/floor', 'position/index-plus-f-halved/floor',
      'position/snap-52-nearest-even/floor', 'position/snap-52-nearest-even/position-product',
    ])
    expect(results.length).toBe(54618)
    expect(probeGrid().length).toBe(129024)
    expect(survivorClasses(survivors, upperSegmentCandidates())).toEqual([survivors])
    // The guard, a float32 position or product, the exact product, a segment index of 2, percent positions, a reciprocal or
    // complement fraction and the other snaps to the 2^-52 grid do not survive.
    for (const id of [
      'position/index-plus-f/guarded', 'position/index-plus-f-f32/floor', 'position/index-plus-f/floor-f32-product', 'position/index-plus-f/floor-exact-product', 'position/index-2-plus-f/floor',
      'position/percent-50/floor', 'position/percent-100/floor', 'position/reciprocal-index-plus-f/floor', 'position/complement-index-plus-f/floor', 'position/half-plus-f/floor',
      'position/snap-52-nearest-up/floor', 'position/snap-52-floor/floor', 'position/snap-52-ceil/floor',
    ]) {
      expect(results.find(r => r.id === id)!.mismatches.length).toBeGreaterThan(0)
    }
  }, HEAVY)
})
