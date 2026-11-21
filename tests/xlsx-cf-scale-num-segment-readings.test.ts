// @vitest-environment node
/**
 * The num-segment follow-up, Phase 1 task 2: each reading the candidate search left standing is implemented and selectable,
 * none is registered. `scale:num-segment-fraction` (a 3-colour segment bounded by two literal num cfvo) has no production
 * reading, so the shape stays withheld (`cf-unmeasured-num-segment-fraction`) and every other cell paints as before;
 * `scale:num-segment-fraction-statistical` (every other 3-colour segment) reads `truncate`, the published behaviour.
 * A reading applies only through the test seam (`injectCfNumSegmentReadings`) together with the capability's evidence: the
 * rehearsal of a proposed registration. Each reading's colour is the candidate class representative of
 * scripts/cf-refs/num-segment-candidates.ts, at every cell.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_NUM_SEGMENT_CAP,
  CF_NUM_SEGMENT_PRODUCTION_READINGS,
  CF_NUM_SEGMENT_READINGS,
  CF_NUM_SEGMENT_STATISTICAL_CAP,
  injectCfNumSegmentReadings,
  interpolateScaleColor,
  numSegmentReadingColor,
  numSegmentUnmeasured,
  scaleColorAt,
  type CfNumSegmentReading,
  type CfScaleSegment,
} from '../src/xlsx/cf/colorScale'
import { CANDIDATES, NUM_SEGMENT_READING_CANDIDATE } from '../scripts/cf-refs/num-segment-candidates'
import type { CfScalePoint } from '../src/xlsx/cf/model'

afterEach(() => injectCfNumSegmentReadings(null))

const RED = 0xfff8696b, YELLOW = 0xffffeb84, GREEN = 0xff63be7b
const P = 0xff14f028, Q = 0xffe61efa
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const evidence = (extra: readonly string[] = []) => ({ hasEvidence: (cap: string) => cap.startsWith('scale:') && (!cap.startsWith('scale:num-segment') || extra.includes(cap)) })
const NONE = evidence()
const WITH_CAP = evidence([CF_NUM_SEGMENT_CAP])
const WITH_STAT = evidence([CF_NUM_SEGMENT_STATISTICAL_CAP])
const WITHHELD = { kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' }

/** A grid of 3-colour literal num cells: two palettes, both colour orders, integer and decimal bounds, both segments. */
function literalGrid(): Array<{ points: CfScalePoint[]; thresholds: number[]; value: number }> {
  const out: Array<{ points: CfScalePoint[]; thresholds: number[]; value: number }> = []
  for (const [a, b, c] of [[7, 217, 637], [0, 30, 60], [0.07, 2.17, 6.37], [-40, 15, 100]] as const)
    for (const cs of [[P, Q, P], [Q, P, Q], [RED, YELLOW, GREEN]])
      for (let i = 0; i <= 60; i++) {
        const value = a + ((c - a) * i) / 60
        out.push({ points: [num(a, cs[0]!), num(b, cs[1]!), num(c, cs[2]!)], thresholds: [a, b, c], value })
      }
  return out
}
const segmentOf = (points: readonly CfScalePoint[], t: readonly number[], v: number): CfScaleSegment => {
  const upper = v > t[1]!
  const [lo, hi] = upper ? [t[1]!, t[2]!] : [t[0]!, t[1]!]
  return { value: v, lo, hi, fraction: Math.max(0, Math.min(1, (v - lo) / (hi - lo))), c0: points[upper ? 1 : 0]!.argb, c1: points[upper ? 2 : 1]!.argb, colours: 3, segment: upper ? 1 : 0, loKind: 'num', hiKind: 'num', thresholds: [...t] }
}

describe('the readings, implemented and unregistered', () => {
  it('six readings; the literal-num id has no production reading, the statistical id reads truncate', () => {
    expect([...CF_NUM_SEGMENT_READINGS]).toEqual(['truncate', 'percent-roundtrip', 'fixed-256', 'fixed-1000', 'fixed-308', 'fixed-700'])
    expect(CF_NUM_SEGMENT_PRODUCTION_READINGS).toEqual({ [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'truncate' })
  })

  it('each reading\'s colour is its candidate class representative at every grid cell', () => {
    const cand = new Map(CANDIDATES.map(c => [c.id, c]))
    for (const r of CF_NUM_SEGMENT_READINGS) {
      const rep = cand.get(NUM_SEGMENT_READING_CANDIDATE[r])!
      expect(rep, r).toBeDefined()
      for (const g of literalGrid()) {
        const s = segmentOf(g.points, g.thresholds, g.value)
        expect(numSegmentReadingColor(r, s.c0, s.c1, s.fraction) >>> 0, `${r} ${g.thresholds.join('/')} ${g.value}`).toBe(rep.colour(s) >>> 0)
      }
    }
  })

  it('truncate is production\'s interpolation; every other reading reproduces T4 and V3 of e448d831, truncate does not', () => {
    for (const g of literalGrid()) {
      const s = segmentOf(g.points, g.thresholds, g.value)
      expect(numSegmentReadingColor('truncate', s.c0, s.c1, s.fraction)).toBe(interpolateScaleColor(s.c0, s.c1, s.fraction))
    }
    const at = (r: CfNumSegmentReading, v: number): string => (numSegmentReadingColor(r, YELLOW, GREEN, (v - 30) / 30) >>> 0).toString(16).toUpperCase()
    expect([at('truncate', 40), at('truncate', 50)]).toEqual(['FFCBDC81', 'FF97CD7E'])
    for (const r of CF_NUM_SEGMENT_READINGS.filter(x => x !== 'truncate')) expect([at(r, 40), at(r, 50)], r).toEqual(['FFCCDD82', 'FF98CE7F'])
  })
})

describe('production default: unchanged', () => {
  it('without the seam (with or without the evidence) the literal-num shape is withheld and every other cell paints production\'s truncation', () => {
    for (const ev of [NONE, WITH_CAP, WITH_STAT]) for (const g of literalGrid()) {
      const s = segmentOf(g.points, g.thresholds, g.value)
      const lo = g.points[s.segment]!, hi = g.points[s.segment + 1]!
      const want = numSegmentUnmeasured(lo, hi, g.value, s.lo, s.hi, s.fraction) ? WITHHELD : { kind: 'ready', value: interpolateScaleColor(s.c0, s.c1, s.fraction) }
      expect(scaleColorAt(g.value, g.points, g.thresholds, ev), `${g.thresholds.join('/')} ${g.value}`).toEqual(want)
    }
  })

  it('a reading injected without its capability\'s evidence changes nothing', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: 'fixed-256', [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'percent-roundtrip' })
    expect(scaleColorAt(40, [num(0, RED), num(30, YELLOW), num(60, GREEN)], [0, 30, 60], NONE)).toEqual(WITHHELD)
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], NONE)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
  })
})

describe('a reading selected through the seam, with its capability\'s evidence (the rehearsal of a registration)', () => {
  for (const r of ['truncate', 'percent-roundtrip', 'fixed-256', 'fixed-1000', 'fixed-308', 'fixed-700'] as const) {
    it(`${r}: every 3-colour literal-num cell paints the reading (dyadic, endpoint and beyond-bound cells included); nothing is withheld`, () => {
      injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: r })
      for (const g of literalGrid()) {
        const s = segmentOf(g.points, g.thresholds, g.value)
        expect(scaleColorAt(g.value, g.points, g.thresholds, WITH_CAP), `${g.thresholds.join('/')} ${g.value}`).toEqual({ kind: 'ready', value: numSegmentReadingColor(r, s.c0, s.c1, s.fraction) })
      }
      // The midpoint above the maximum re-sorts first: num 7 / 637 / 217 reads as 7 / 217 / 637.
      expect(scaleColorAt(77, [num(7, P), num(637, Q), num(217, P)], [7, 637, 217], WITH_CAP)).toEqual({ kind: 'ready', value: numSegmentReadingColor(r, P, Q, (77 - 7) / 210) })
    })
  }

  it('the literal-num reading leaves 2-colour, statistical and mixed segments on production\'s truncation', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: 'fixed-256' })
    const both = evidence([CF_NUM_SEGMENT_CAP, CF_NUM_SEGMENT_STATISTICAL_CAP])
    expect(scaleColorAt(40, [num(30, YELLOW), num(60, GREEN)], [30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    const mixed: CfScalePoint[] = [num(0, RED), num(30, YELLOW), { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, mixed, [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    // A formula bound resolved to a number is not a literal num bound (as numSegmentUnmeasured).
    expect(scaleColorAt(40, [num(0, RED), { bound: { type: 'num', value: 30, resolved: 'formula' }, argb: YELLOW }, num(60, GREEN)], [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
  })

  it('the statistical reading applies to every 3-colour segment not bounded by two literal num cfvo, never to 2-colour scales', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'percent-roundtrip' })
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: P }, { bound: { type: 'percentile', value: 50 }, argb: Q }, { bound: { type: 'max' }, argb: P }]
    for (const v of [77, 147, 42, 357, 497]) {
      const s = segmentOf(stat, [7, 217, 637], v)
      expect(scaleColorAt(v, stat, [7, 217, 637], WITH_STAT), String(v)).toEqual({ kind: 'ready', value: numSegmentReadingColor('percent-roundtrip', s.c0, s.c1, s.fraction) })
    }
    expect(scaleColorAt(77, stat, [7, 217, 637], WITH_STAT)).not.toEqual({ kind: 'ready', value: interpolateScaleColor(P, Q, 1 / 3) })
    // The literal-num shape is still withheld (its own id has no reading), 2-colour scales are untouched.
    expect(scaleColorAt(77, [num(7, P), num(217, Q), num(637, P)], [7, 217, 637], WITH_STAT)).toEqual(WITHHELD)
    expect(scaleColorAt(77, [{ bound: { type: 'min' }, argb: P }, { bound: { type: 'max' }, argb: Q }], [7, 217], WITH_STAT)).toEqual({ kind: 'ready', value: interpolateScaleColor(P, Q, 1 / 3) })
  })

  it('an unknown reading name is no reading (the shape stays withheld)', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: 'round-half-up' as CfNumSegmentReading })
    expect(scaleColorAt(40, [num(0, RED), num(30, YELLOW), num(60, GREEN)], [0, 30, 60], WITH_CAP)).toEqual(WITHHELD)
  })
})
