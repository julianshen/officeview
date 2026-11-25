// @vitest-environment node
/**
 * The upper-segment whole-product follow-up, Phase 1 task 1: the extended candidate search
 * (scripts/cf-refs/upper-whole-candidates.ts) over every captured colour-scale cell, stage2-scale-num-segment included.
 * Pins the calibration (production's truncation misses exactly the 21 moved cells), the result (no survivor at any scope,
 * so the follow-up case is exploratory), the near misses docs/validation/xlsx-cf-recipes-stage2.md §9.1 reports, and the
 * fail-closed evaluation. Nothing here registers anything. The same search with the exploratory capture of
 * stage2-scale-upper-whole (54fc3388) included is pinned at the end (recipe §9.5): still no survivor.
 */
import { describe, expect, it, beforeAll, afterAll } from 'vitest'
import { asBeforePercentOrder, asRegistered } from './helpers/cf-percent-order-asof'
import { evaluateUpperWhole, NEAR_MISS_ID, roundBits, truncationMisses, upperWholeCandidates, upperWholeCells, type UpperWholeResult } from '../scripts/cf-refs/upper-whole-candidates'
import type { SegmentCell } from '../scripts/cf-refs/num-segment-candidates'
import type { CfScaleSegment } from '../src/xlsx/cf/colorScale'

// As of its capture: before the percent-order registration (§14.z.24), percent thresholds lerp-first and no
// scale:formula-bounds-v2 reading (tests/helpers/cf-percent-order-asof.ts).
beforeAll(asBeforePercentOrder)
afterAll(asRegistered)

const HEAVY = 300_000
let cells: SegmentCell[] | undefined
// The search as run before stage2-scale-upper-whole was captured (recipe §9.1); the capture is pinned separately below.
const cellsOnce = async (): Promise<SegmentCell[]> => (cells ??= await upperWholeCells(undefined, { include: 'pre-capture' }))
let results: UpperWholeResult[] | undefined
const resultsOnce = async (): Promise<UpperWholeResult[]> => (results ??= evaluateUpperWhole(await cellsOnce()))

const MOVED = [
  'stage2-scale-scope-disjoint:T4', 'stage2-scale-scope-disjoint:V3',
  ...['B17', 'B18', 'B21', 'B24', 'B28', 'C17', 'C18', 'C21', 'C24', 'C28', 'D7', 'D8', 'D12', 'D13', 'E17', 'E18', 'E21', 'E24', 'E28'].map(r => `stage2-scale-num-segment:${r}`),
]

describe('the extended upper-segment whole-product search', () => {
  it('runs on all 942 captured colour-scale cells; production\'s truncation misses exactly the 21 moved cells', async () => {
    const cs = await cellsOnce()
    expect(cs.length).toBe(942)
    expect(cs.filter(c => c.caseId === 'stage2-scale-num-segment').length).toBe(217)
    expect(cs.filter(c => c.caseId.startsWith('stage5-')).length).toBe(520)
    expect(truncationMisses(cs)).toEqual(MOVED)
  }, HEAVY)

  it('48345 candidates at three scopes: no survivor anywhere (the case is exploratory)', async () => {
    expect(upperWholeCandidates().length).toBe(48345)
    const rs = await resultsOnce()
    expect(rs.length).toBe(48345 * 3)
    expect(rs.filter(r => r.mismatches.length === 0)).toEqual([])
    const fewest = (family: string, scope: string): number => Math.min(...rs.filter(r => r.family === family && r.scope === scope).map(r => r.mismatches.length))
    expect(['origin', 'f32', 'precision', 'rational'].map(f => fewest(f, 'upper-3-colour'))).toEqual([14, 1, 2, 1])
  }, HEAVY)

  it('the near misses: float32 from the maximum fails only E4 of stage2-scale-thresholds; the 28-bit pair only the 7/10 separators', async () => {
    const rs = await resultsOnce()
    const one = rs.filter(r => r.mismatches.length === 1)
    expect(new Set(one.map(r => r.mismatches[0]))).toEqual(new Set(['stage2-scale-thresholds:E4']))
    expect(one.every(r => r.scope === 'upper-3-colour' && (r.id.startsWith('f32/from-max[') || r.id === 'rational/from-max-step/p24'))).toBe(true)
    expect(one.length).toBe(29)
    const two = rs.filter(r => r.mismatches.length === 2)
    expect(new Set(two.map(r => r.mismatches.join(',')))).toEqual(new Set(['stage2-scale-thresholds:E4,stage2-scale-thresholds:E5', 'stage2-scale-num-segment:H25,stage2-scale-num-segment:I25']))
    // The 28-bit near miss recurs 12 bits on (the moved denominators' binary periods divide 12): not one precision.
    expect(two.filter(r => r.mismatches[0] === 'stage2-scale-num-segment:H25').map(r => r.id).sort()).toEqual([
      'precision/complement-raw-step/bin28/bin30', 'precision/complement-raw-step/bin40/bin42', 'precision/span-minus-max-step/bin28/bin30', 'precision/span-minus-max-step/bin40/bin42',
    ])
  }, HEAVY)

  it('E4 is a whole product in exact arithmetic, drawn unmoved: (3 − 2.2)/(10 − 2.2)·156 = 16; its double product is not whole', () => {
    // Decimal rationals: 0.8/7.8 = 8/78 = 4/39, times 156 = 16 exactly.
    expect((8n * 156n) % 78n).toBe(0n)
    expect(((3 - 2.2) / (10 - 2.2)) * 156).toBe(15.999999999999996)
    // The float32 from-max product lands one float32 ulp below 16, as 1/15's (moved) lands below 14.
    const f32FromMax = (v: number, lo: number, hi: number): number => Math.fround(1 - Math.fround(Math.fround(hi - v) / Math.fround(hi - lo)))
    expect(Math.fround(f32FromMax(3, 2.2, 10) * 156)).toBe(15.999999046325684)
    expect(Math.fround(f32FromMax(245, 217, 637) * 210)).toBe(13.999999046325684)
  })
})

describe('fail-closed evaluation', () => {
  const seg = (o: Partial<CfScaleSegment>): CfScaleSegment => ({ value: 40, lo: 30, hi: 60, fraction: 1 / 3, c0: 0xffffeb84, c1: 0xff63be7b, colours: 3, segment: 1, loKind: 'num', hiKind: 'num', thresholds: [0, 30, 60], ...o })
  const sample = (): ReturnType<typeof upperWholeCandidates> => {
    const all = upperWholeCandidates()
    return [...new Set(['origin', 'f32', 'precision', 'rational'])].flatMap(f => all.filter(c => c.family === f).filter((_, i) => i % 97 === 0))
  }

  it('every candidate is finite on an ordinary segment', () => {
    for (const c of sample()) expect(c.evaluate(seg({})), c.id).toBeTypeOf('number')
  })

  it('a non-finite or overflowing input, or a zero span, is undefined (never packed)', () => {
    const bad = [
      seg({ value: Number.POSITIVE_INFINITY }),
      seg({ lo: -1.7e308, hi: 1.7e308, value: 0, thresholds: [-1.7e308, -1.7e308, 1.7e308] }),
      seg({ lo: 30, hi: 30, value: 30 }),
    ]
    for (const c of sample()) for (const s of bad) expect(c.evaluate(s), c.id).toBeUndefined()
  })

  it('a float32 conversion that overflows is trapped (finite doubles beyond float32 range)', () => {
    const s = seg({ value: 2e39, lo: 1e39, hi: 4e39, thresholds: [0, 1e39, 4e39] })
    const f32 = upperWholeCandidates().filter(c => /^f32\/[^[]+\[1/.test(c.id))
    expect(f32.length).toBeGreaterThan(0)
    for (const c of f32) expect(c.evaluate(s), c.id).toBeUndefined()
  })

  it('roundBits rounds to a p-bit significand', () => {
    expect(roundBits(1 / 3, 24, 'nearest')).toBe(Math.fround(1 / 3))
    expect(roundBits(1 / 3, 53, 'nearest')).toBe(1 / 3)
    expect(roundBits(0.75, 1, 'truncate')).toBe(0.5)
  })
})

describe('after the capture of stage2-scale-upper-whole (54fc3388, exploratory): the search on every capture', () => {
  const ID = 'stage2-scale-upper-whole'
  // The 79 designated cells Excel drew one step smaller on every exactly-whole channel (the rest are at truncation).
  const NEW_MOVED = ('A2 A3 A4 A5 A8 A9 A14 A15 A16 B2 B3 B4 B5 B9 B11 B12 B14 B15 B16 B19 B20 B21 B23 B24 C2 C3 C4 C5 C7 C8 C9 C10 C12 C13 '
    + 'C14 C15 C17 C18 C19 C20 C22 C23 C24 C25 D2 D3 D4 D5 D7 D8 D9 D10 E2 E3 E4 E5 E7 E8 E9 E10 E12 E13 E14 E15 E17 E18 E20 F2 F3 F4 '
    + 'F6 F10 F12 F14 F15 F16 F17 F18 F19').split(' ').map(r => `${ID}:${r}`)
  let all: SegmentCell[] | undefined
  let allResults: UpperWholeResult[] | undefined
  const allOnce = async (): Promise<SegmentCell[]> => (all ??= await upperWholeCells())
  const allResultsOnce = async (): Promise<UpperWholeResult[]> => (allResults ??= evaluateUpperWhole(await allOnce()))

  it('1101 cells (159 from the new capture); production\'s truncation misses the 21 earlier moved cells and the 79 new ones', async () => {
    const cs = await allOnce()
    expect(cs.length).toBe(1101)
    expect(cs.filter(c => c.caseId === ID).length).toBe(159)
    expect(new Set(truncationMisses(cs))).toEqual(new Set([...MOVED, ...NEW_MOVED]))
  }, HEAVY)

  it('no survivor at any scope; the float32 near miss fails 20 cells (E4 and 19 of the new capture)', async () => {
    const rs = await allResultsOnce()
    expect(rs.filter(r => r.mismatches.length === 0)).toEqual([])
    const fewest = (family: string): number => Math.min(...rs.filter(r => r.family === family && r.scope === 'upper-3-colour').map(r => r.mismatches.length))
    expect(['origin', 'f32', 'precision', 'rational'].map(fewest)).toEqual([55, 20, 13, 14])
    const near = rs.find(r => r.id === NEAR_MISS_ID && r.scope === 'upper-3-colour')!
    expect(near.mismatches).toEqual(['stage2-scale-thresholds:E4', ...'A10 A11 A17 B6 B7 B8 B17 B18 B22 B23 B24 D19 F5 F8 F10 F12 F14 F15 F16'.split(' ').map(r => `${ID}:${r}`)])
  }, HEAVY)
})
