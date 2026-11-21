// @vitest-environment node
/**
 * The num-segment follow-up, Phase 1 task 1: the candidate search over every captured colour-scale cell
 * (scripts/cf-refs/num-segment-candidates.ts). Pins the harness calibration (production's own truncation reproduces
 * production at every included cell and Excel at every one but T4 and V3 of stage2-scale-scope-disjoint) and the
 * survivor sets per scope that docs/validation/xlsx-cf-recipes-stage2.md §8 reports. Nothing here registers anything.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { BASELINE_ID, CANDIDATES, collectSegmentCells, evaluateCandidates, SWEEP_N, type SegmentCollection } from '../scripts/cf-refs/num-segment-candidates'
import { injectCfScaleSegmentProbe, scaleColorAt, type CfScaleSegment } from '../src/xlsx/cf/colorScale'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import type { CfManifest } from '../scripts/cf-refs/cases'

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
    expect(scaleColorAt(45, P, [0, 30, 60], all)).toEqual({ kind: 'ready', value: 0xffb1d580 })
    expect(scaleColorAt(40, P, [0, 30, 60], all)).toEqual({ kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' })
    expect(seen[0]).toEqual({ value: 45, lo: 30, hi: 60, fraction: 0.5, c0: YELLOW, c1: GREEN, colours: 3, segment: 1, loKind: 'num', hiKind: 'num', thresholds: [0, 30, 60] })
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
})
