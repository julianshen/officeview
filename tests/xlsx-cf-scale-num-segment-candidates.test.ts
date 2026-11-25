// @vitest-environment node
/**
 * The num-segment follow-up, Phase 1 task 1: the candidate search over every captured colour-scale cell
 * (scripts/cf-refs/num-segment-candidates.ts). Pins the harness calibration (production's own truncation reproduces
 * production at every included cell and Excel at every one but T4 and V3 of stage2-scale-scope-disjoint) and the
 * survivor sets per scope that docs/validation/xlsx-cf-recipes-stage2.md §8 reports. Nothing here registers anything.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it, beforeAll, afterAll } from 'vitest'
import { asBeforePercentOrder, asRegistered } from './helpers/cf-percent-order-asof'
import { BASELINE_ID, CANDIDATES, collectSegmentCells, computeNumSegmentClasses, consensusDomain, consensusWithheldShare, evaluateCandidates, SWEEP_N, type SegmentCollection } from '../scripts/cf-refs/num-segment-candidates'
import { CF_NUM_SEGMENT_CLASSES } from '../scripts/cf-refs/num-segment-classes'
import { numSegmentCandidate, numSegmentCandidates } from '../scripts/cf-refs/num-segment-formulas'
import { S2_NUM_SEGMENT_READINGS } from '../scripts/cf-refs/cases'
import { injectCfScaleSegmentProbe, interpolateScaleColor, scaleColorAt, type CfScaleSegment } from '../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import type { CfManifest } from '../scripts/cf-refs/cases'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// As of its capture: before the percent-order registration (§14.z.24), percent thresholds lerp-first and no
// scale:formula-bounds-v2 reading (tests/helpers/cf-percent-order-asof.ts).
beforeAll(asBeforePercentOrder)
afterAll(asRegistered)

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

const REPO_ROOT = resolve(__dirname, '..')
const HEAVY = 300_000
const manifest = JSON.parse(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/manifest.json'), 'utf8')) as CfManifest
let collected: SegmentCollection | undefined
const collection = async (): Promise<SegmentCollection> => (collected ??= await collectSegmentCells(manifest, REPO_ROOT))

afterEach(() => injectCfScaleSegmentProbe(null))

describe('the segment probe seam (test-only)', () => {
  const RED = 0xfff8696b, YELLOW = 0xffffeb84, GREEN = 0xff63be7b
  const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
  const all = { hasEvidence: () => true }
  const P = [num(0, RED), num(30, YELLOW), num(60, GREEN)]

  it('sees each segment evaluation with its thresholds, colours and cfvo kinds; undefined keeps production', () => {
    const seen: CfScaleSegment[] = []
    injectCfScaleSegmentProbe(s => { seen.push(s); return undefined })
    expect(scaleColorAt(15, P, [0, 30, 60], all)).toEqual({ kind: 'ready', value: interpolateScaleColor(RED, YELLOW, 0.5) })
    expect(scaleColorAt(40, P, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' })
    expect(seen[0]).toEqual({ value: 15, lo: 0, hi: 30, fraction: 0.5, c0: RED, c1: YELLOW, colours: 3, segment: 0, loKind: 'num', hiKind: 'num', thresholds: [0, 30, 60] })
    expect(seen[1]!.fraction).toBe(1 / 3)
  })

  it('may replace the outcome; cleared, production is back', () => {
    injectCfScaleSegmentProbe(() => ({ kind: 'ready', value: 0xff123456 }))
    expect(scaleColorAt(40, P, [0, 30, 60], all)).toEqual({ kind: 'ready', value: 0xff123456 })
    injectCfScaleSegmentProbe(null)
    expect(scaleColorAt(40, P, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' })
    // Re-sorted midpoint above the maximum and the 2-colour inversion report the segment as evaluated (lo < hi).
    const seen: CfScaleSegment[] = []
    injectCfScaleSegmentProbe(s => { seen.push(s); return undefined })
    scaleColorAt(150, [num(10, RED), num(200, YELLOW), num(100, GREEN)], [10, 200, 100], all)
    scaleColorAt(20, [num(40, RED), num(10, GREEN)], [40, 10], all)
    expect(seen.map(s => [s.lo, s.hi, s.c0, s.c1, s.segment])).toEqual([[100, 200, YELLOW, GREEN, 1], [10, 40, RED, GREEN, 0]])
  })
})

describe('the candidate search over every captured colour-scale cell', () => {
  it('collects 725 cells (714 confirmed, 11 withheld as the num-segment shape) with no harness fault; the probe is cleared', async () => {
    const { cells, faults, excluded } = await collection()
    expect(faults).toEqual([])
    expect(cells.length).toBe(725)
    expect(cells.filter(c => c.status === 'withheld').map(c => c.ref)).toEqual(['R3', 'R4', 'R5', 'S2', 'S3', 'S4', 'T3', 'T4', 'U3', 'V3', 'W4'])
    expect(cells.filter(c => c.status === 'withheld').every(c => c.caseId === 'stage2-scale-scope-disjoint')).toBe(true)
    expect(excluded.length).toBe(144)
    // Fix round 1 (R3): 112 inapplicable (no fill observation, Excel shows no fill, a non-numeric value, an endpoint or a
    // fill another rule paints) and 32 numeric, Excel-filled cells dropped only by a production gate, every one of them
    // under 2-colour scales only: outside both changed scopes (every proposed change is 3-colour only).
    expect(excluded.filter(e => e.group === 'a').length).toBe(112)
    const gated = excluded.filter(e => e.group === 'b')
    expect(gated.length).toBe(32)
    expect(gated.every(e => e.threeColour === false)).toBe(true)
    expect([...new Set(gated.map(e => e.caseId))].sort()).toEqual(['stage2-scale-composition', 'stage2-scale-scope', 'stage5-composition'])
    expect(scaleColorAt(40, [{ bound: { type: 'num', value: 0 }, argb: 0 }, { bound: { type: 'num', value: 30 }, argb: 0 }, { bound: { type: 'num', value: 60 }, argb: 0 }], [0, 30, 60], { hasEvidence: () => true }).kind).toBe('unresolved')
  }, HEAVY)

  it('calibration: production\'s truncation misses exactly T4 and V3, at every scope', async () => {
    const { cells } = await collection()
    const base = evaluateCandidates(cells, CANDIDATES.filter(c => c.id === BASELINE_ID))
    for (const r of base) expect(r.mismatches.map(m => m.split(' ')[0]), r.scope).toEqual(['stage2-scale-scope-disjoint:T4', 'stage2-scale-scope-disjoint:V3'])
  }, HEAVY)

  it('survivors: none over all segments or all literal-num segments; the percent round trip over 3-colour segments; a family over 3-colour literal-num segments', async () => {
    const { cells } = await collection()
    const results = evaluateCandidates(cells)
    const survivors = (scope: string): string[] => results.filter(r => r.scope === scope && r.mismatches.length === 0).map(r => r.id)
    expect(survivors('all')).toEqual([])
    expect(survivors('literal-num')).toEqual([])
    expect(survivors('3-colour')).toEqual(['percent-roundtrip/step-truncate-unguarded', 'sweep-roundtrip-100'])
    const narrow = survivors('3-colour-literal-num')
    expect(narrow.length).toBe(857)
    const fixedN = narrow.filter(id => /^sweep-fixed-\d+$/.test(id)).map(id => Number(id.slice('sweep-fixed-'.length)))
    expect([fixedN.length, SWEEP_N.length, Math.min(...fixedN)]).toEqual([790, 5027, 232])
    // Every fixed-point survivor keeps the quarter exact (U4: f = 1/4) and does not hold a third exactly (T4, V3).
    expect(fixedN.every(N => N % 4 === 0 && N % 3 !== 0)).toBe(true)
    expect(narrow.filter(id => id.startsWith('sweep-roundtrip'))).toEqual(['sweep-roundtrip-reciprocal-7', 'sweep-roundtrip-after-7', 'sweep-roundtrip-100', 'sweep-roundtrip-after-100'])
  }, HEAVY)

  it('fix round 1 (R1): the committed classes are the survivors\' classes on the case\'s literal-num cells; consensus withholds 0.22% of the test domain (size-weighted)', async () => {
    const { cells } = await collection()
    const classes = computeNumSegmentClasses(cells)
    expect(classes).toEqual(Object.fromEntries(Object.entries(CF_NUM_SEGMENT_CLASSES).map(([k, v]) => [k, [...v]])))
    expect(Object.keys(classes)).toEqual([...S2_NUM_SEGMENT_READINGS])
    expect(Object.keys(classes).length).toBe(126)
    // Every survivor is in exactly one class.
    expect(Object.values(classes).slice(1).flat().length).toBe(857)
    const domain = consensusDomain()
    const shares = Object.entries(classes).slice(1).map(([, m]) => [m.length, consensusWithheldShare(m, domain)] as const)
    const weighted = shares.reduce((a, [n, w]) => a + n * w, 0) / 857
    expect(weighted).toBeLessThan(0.0025)
    expect(Math.max(...shares.map(([, w]) => w))).toBeLessThan(0.05)
    expect(shares.filter(([, w]) => w === 0).length).toBe(113)
  }, HEAVY)
})

describe('Candidate.evaluate (pre-capture fix rounds 2-3): every formulation fails closed on a non-finite intermediate', () => {
  const P = 0xff14f028, Q = 0xffe61efa
  const seg = (value: number, t: readonly number[], c0 = 0xff00ff00, c1 = 0xffff00ff): CfScaleSegment => {
    const upper = value > t[1]!
    const [lo, hi] = upper ? [t[1]!, t[2]!] : [t[0]!, t[1]!]
    return { value, lo, hi, fraction: Math.max(0, Math.min(1, (value - lo) / (hi - lo))), c0, c1, colours: 3, segment: upper ? 1 : 0, loKind: 'num', hiKind: 'num', thresholds: [...t] }
  }

  it('the reviewer\'s counterexample: float32-ops/step-truncate at value 1 of 0 / 1e40 / 2e40 (Math.fround(hi) is Infinity) is undefined, not FF00FF00', () => {
    const c = numSegmentCandidate('float32-ops/step-truncate')!
    const s = seg(1, [0, 1e40, 2e40])
    expect(c.colour(s).toString(16).toUpperCase()).toBe('FF00FF00')
    expect(c.evaluate(s)).toBeUndefined()
  })

  it('every candidate of the search, on overflow and float32-overflow inputs: undefined or a valid colour; on finite-safe inputs exactly colour()', () => {
    const all = numSegmentCandidates()
    expect(all.length).toBe(5803)
    const unsafe = [seg(5e-324, [0, 2e-323, 4e-323], P, Q), seg(1, [0, 1e40, 2e40]), seg(5e39, [0, 1e40, 2e40]), seg(0, [-9e307, 9e307, 9.5e307]), seg(5e299, [0, 1e300, 2e300]), seg(1.5e300, [0, 1e300, 2e300]), seg(-1e39, [-2e39, 0, 1e39])]
    const safe = [seg(40, [0, 30, 60], 0xffffeb84, 0xff63be7b), seg(77, [7, 217, 637], P, Q), seg(357, [7, 217, 637], Q, P), seg(1.5, [1, 2, 3]), seg(3e37, [0, 1e38, 2e38], P, Q)]
    const valid = (v: number | undefined): boolean => v === undefined || (Number.isInteger(v) && v >= 0 && v <= 0xffffffff && v >>> 24 === 0xff)
    for (const c of all) {
      for (const s of unsafe) expect(valid(c.evaluate(s)), `${c.id} ${s.thresholds.join('/')} ${s.value}`).toBe(true)
      for (const s of safe) expect(c.evaluate(s), `${c.id} ${s.thresholds.join('/')} ${s.value}`).toBe(c.colour(s))
    }
    // No formulation that converts its inputs to float32 accepts an input outside float32's range.
    for (const c of all.filter(x => x.id.startsWith('float32-ops/'))) for (const s of [seg(1, [0, 1e40, 2e40]), seg(5e299, [0, 1e300, 2e300])]) expect(c.evaluate(s), `${c.id}`).toBeUndefined()
  })
})
