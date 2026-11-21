// @vitest-environment node
/**
 * The num-segment follow-up, Phase 1 task 2 and fix round 1: the readings, implemented and selectable, none registered.
 * `scale:num-segment-fraction` (a 3-colour segment bounded by two literal num cfvo) has no production reading, so the shape
 * stays withheld (`cf-unmeasured-num-segment-fraction`); `scale:num-segment-fraction-statistical` (every other 3-colour
 * segment) reads `truncate`, the published behaviour. A reading applies only through the test seam
 * (`injectCfNumSegmentReadings`) together with its id's evidence: the rehearsal of a proposed registration.
 * - R1 (consensus): a literal-num reading is its whole committed class (`CF_NUM_SEGMENT_CLASSES`); a cell paints only where
 *   every member agrees, else `cf-num-segment-class-disagreement`.
 * - R2: a non-truncate statistical reading withholds the non-dyadic interior statistical cells
 *   (`cf-num-segment-statistical-withheld`), never switching arithmetic.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_NUM_SEGMENT_CAP,
  CF_NUM_SEGMENT_PRODUCTION_READINGS,
  CF_NUM_SEGMENT_READINGS,
  CF_NUM_SEGMENT_STATISTICAL_CAP,
  injectCfNumSegmentReadings,
  interpolateScaleColor,
  numSegmentConsensusColor,
  numSegmentReadingColor,
  numSegmentUnmeasured,
  scaleColorAt,
  type CfScaleSegment,
} from '../src/xlsx/cf/colorScale'
import { CF_NUM_SEGMENT_CLASSES } from '../src/xlsx/cf/numSegmentClasses'
import { numSegmentCandidate } from '../src/xlsx/cf/numSegmentCandidates'
import { S2_NUM_SEGMENT_READINGS } from '../scripts/cf-refs/cases'
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
const DISAGREE = { kind: 'unresolved', reason: 'cf-num-segment-class-disagreement' }
const STAT_WITHHELD = { kind: 'unresolved', reason: 'cf-num-segment-statistical-withheld' }

/** A grid of 3-colour literal num cells: palettes, colour orders, integer and decimal bounds (the review's 0/1000/2000 and 0/8/16 included). */
function literalGrid(): Array<{ points: CfScalePoint[]; thresholds: number[]; value: number }> {
  const out: Array<{ points: CfScalePoint[]; thresholds: number[]; value: number }> = []
  for (const [a, b, c] of [[7, 217, 637], [0, 30, 60], [0.07, 2.17, 6.37], [-40, 15, 100], [0, 1000, 2000], [0, 8, 16]] as const)
    for (const cs of [[P, Q, P], [Q, P, Q], [RED, YELLOW, GREEN]])
      for (let i = 0; i <= 40; i++) {
        const value = a + ((c - a) * i) / 40
        out.push({ points: [num(a, cs[0]!), num(b, cs[1]!), num(c, cs[2]!)], thresholds: [a, b, c], value })
      }
  for (const v of [6, 10, 700]) out.push({ points: [num(0, P), num(1000, Q), num(2000, P)], thresholds: [0, 1000, 2000], value: v })
  out.push({ points: [num(0, P), num(8, Q), num(16, P)], thresholds: [0, 8, 16], value: 1 })
  return out
}
const segmentOf = (points: readonly CfScalePoint[], t: readonly number[], v: number): CfScaleSegment => {
  const upper = v > t[1]!
  const [lo, hi] = upper ? [t[1]!, t[2]!] : [t[0]!, t[1]!]
  return { value: v, lo, hi, fraction: Math.max(0, Math.min(1, (v - lo) / (hi - lo))), c0: points[upper ? 1 : 0]!.argb, c1: points[upper ? 2 : 1]!.argb, colours: 3, segment: upper ? 1 : 0, loKind: 'num', hiKind: 'num', thresholds: [...t] }
}

describe('the readings: one committed class each, unregistered', () => {
  it('truncate plus one reading per class; the catalogue names the same readings; the statistical id reads truncate, the literal-num id none', () => {
    expect([...CF_NUM_SEGMENT_READINGS]).toEqual(Object.keys(CF_NUM_SEGMENT_CLASSES))
    expect([...CF_NUM_SEGMENT_READINGS]).toEqual([...S2_NUM_SEGMENT_READINGS])
    expect(CF_NUM_SEGMENT_READINGS[0]).toBe('truncate')
    expect(CF_NUM_SEGMENT_CLASSES.truncate).toEqual(['exact/step-truncate'])
    expect(CF_NUM_SEGMENT_PRODUCTION_READINGS).toEqual({ [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'truncate' })
    // Every member is a known candidate, and no candidate sits in two classes.
    const all = Object.values(CF_NUM_SEGMENT_CLASSES).flat()
    for (const id of all) expect(numSegmentCandidate(id), id).toBeDefined()
    expect(new Set(all).size).toBe(all.length)
  })

  it('each reading\'s representative is a member of its class, and the consensus is that member\'s colour wherever the members agree', () => {
    for (const r of CF_NUM_SEGMENT_READINGS) {
      for (const g of literalGrid()) {
        const s = segmentOf(g.points, g.thresholds, g.value)
        const rep = numSegmentReadingColor(r, s.c0, s.c1, s.fraction)
        const memberColours = CF_NUM_SEGMENT_CLASSES[r]!.map(id => numSegmentCandidate(id)!.colour(s) >>> 0)
        expect(memberColours, `${r}: the representative is a member`).toContain(rep)
        const c = numSegmentConsensusColor(r, s)
        if (c !== undefined) expect(c, `${r} ${g.thresholds.join('/')} ${g.value}`).toBe(rep)
        else expect(new Set(memberColours).size, `${r}: withheld only where members disagree`).toBeGreaterThan(1)
      }
    }
  })

  it('truncate is production\'s interpolation; every other reading reproduces T4 and V3 of e448d831, truncate does not', () => {
    for (const g of literalGrid()) {
      const s = segmentOf(g.points, g.thresholds, g.value)
      expect(numSegmentReadingColor('truncate', s.c0, s.c1, s.fraction)).toBe(interpolateScaleColor(s.c0, s.c1, s.fraction))
    }
    const at = (r: string, v: number): string | undefined => { const c = numSegmentConsensusColor(r, segmentOf([num(0, RED), num(30, YELLOW), num(60, GREEN)], [0, 30, 60], v)); return c === undefined ? undefined : c.toString(16).toUpperCase() }
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
    // The review's dyadic eighth (0/8/16, value 1): production paints FF2ED642 today.
    expect(scaleColorAt(1, [num(0, P), num(8, Q), num(16, P)], [0, 8, 16], NONE)).toEqual({ kind: 'ready', value: 0xff2ed642 })
  })

  it('a reading injected without its capability\'s evidence changes nothing', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: CF_NUM_SEGMENT_READINGS[2]!, [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'percent-roundtrip' })
    expect(scaleColorAt(40, [num(0, RED), num(30, YELLOW), num(60, GREEN)], [0, 30, 60], NONE)).toEqual(WITHHELD)
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], NONE)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
  })
})

describe('a reading selected through the seam, with its capability\'s evidence (the rehearsal of a registration)', () => {
  it('every reading: each 3-colour literal-num cell paints the class consensus or is withheld as a class disagreement', () => {
    for (const r of CF_NUM_SEGMENT_READINGS) {
      injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: r })
      for (const g of literalGrid()) {
        const c = numSegmentConsensusColor(r, segmentOf(g.points, g.thresholds, g.value))
        expect(scaleColorAt(g.value, g.points, g.thresholds, WITH_CAP), `${r} ${g.thresholds.join('/')} ${g.value}`).toEqual(c === undefined ? DISAGREE : { kind: 'ready', value: c })
      }
      // The midpoint above the maximum re-sorts first: num 7 / 637 / 217 reads as 7 / 217 / 637.
      const c = numSegmentConsensusColor(r, segmentOf([num(7, P), num(217, Q), num(637, P)], [7, 217, 637], 77))
      expect(scaleColorAt(77, [num(7, P), num(637, Q), num(217, P)], [7, 637, 217], WITH_CAP)).toEqual(c === undefined ? DISAGREE : { kind: 'ready', value: c })
    }
  })

  it('the review\'s counterexamples: the formerly indistinguishable members now sit in separate classes, or the class withholds there', () => {
    const cases: Array<[string, string, number[], number]> = [
      ['sweep-fixed-256', 'fixed-floor-512/step-truncate', [0, 1000, 2000], 6],
      ['sweep-fixed-308', 'sweep-fixed-364', [0, 1000, 2000], 6],
      ['sweep-fixed-700', 'sweep-fixed-980', [0, 1000, 2000], 10],
      ['percent-roundtrip/step-truncate-unguarded', 'sweep-roundtrip-reciprocal-7', [0, 1000, 2000], 700],
      ['sweep-fixed-308', 'exact/step-truncate', [0, 8, 16], 1],
    ]
    const classOf = (id: string): string => Object.entries(CF_NUM_SEGMENT_CLASSES).find(([, m]) => m.includes(id))![0]
    for (const [a, b, t, v] of cases) {
      const s = segmentOf([num(t[0]!, P), num(t[1]!, Q), num(t[2]!, P)], t, v)
      const ca = numSegmentCandidate(a)!.colour(s) >>> 0, cb = numSegmentCandidate(b)!.colour(s) >>> 0
      expect(ca, `${a} vs ${b}: the review's input separates them`).not.toBe(cb)
      if (classOf(a) === classOf(b)) expect(numSegmentConsensusColor(classOf(a), s), `${a}/${b} share ${classOf(a)}: withheld`).toBeUndefined()
    }
  })

  it('the literal-num reading leaves 2-colour, statistical and mixed segments on production\'s truncation', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: CF_NUM_SEGMENT_READINGS[2]! })
    const both = evidence([CF_NUM_SEGMENT_CAP, CF_NUM_SEGMENT_STATISTICAL_CAP])
    expect(scaleColorAt(40, [num(30, YELLOW), num(60, GREEN)], [30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    const mixed: CfScalePoint[] = [num(0, RED), num(30, YELLOW), { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, mixed, [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
    expect(scaleColorAt(40, [num(0, RED), { bound: { type: 'num', value: 30, resolved: 'formula' }, argb: YELLOW }, num(60, GREEN)], [0, 30, 60], both)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
  })

  it('the statistical percent-roundtrip outcome withholds every non-dyadic interior 3-colour cell not bounded by two literal num cfvo, and nothing else', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'percent-roundtrip' })
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: P }, { bound: { type: 'percentile', value: 50 }, argb: Q }, { bound: { type: 'max' }, argb: P }]
    for (const v of [77, 147, 42, 357, 497]) expect(scaleColorAt(v, stat, [7, 217, 637], WITH_STAT), String(v)).toEqual(STAT_WITHHELD)
    // Dyadic interior cells, endpoints and beyond-bound cells paint truncation as before.
    for (const v of [112, 427, 59.5, 7, 217, 637, 900]) {
      const s = segmentOf(stat, [7, 217, 637], v)
      expect(scaleColorAt(v, stat, [7, 217, 637], WITH_STAT), String(v)).toEqual({ kind: 'ready', value: interpolateScaleColor(s.c0, s.c1, s.fraction) })
    }
    // Mixed and formula-resolved bounds are statistical too.
    expect(scaleColorAt(40, [num(0, RED), num(30, YELLOW), { bound: { type: 'max' }, argb: GREEN }], [0, 30, 60], WITH_STAT)).toEqual(STAT_WITHHELD)
    expect(scaleColorAt(40, [num(0, RED), { bound: { type: 'num', value: 30, resolved: 'formula' }, argb: YELLOW }, num(60, GREEN)], [0, 30, 60], WITH_STAT)).toEqual(STAT_WITHHELD)
    // The literal-num shape keeps its own withhold; 2-colour scales are untouched.
    expect(scaleColorAt(77, [num(7, P), num(217, Q), num(637, P)], [7, 217, 637], WITH_STAT)).toEqual(WITHHELD)
    expect(scaleColorAt(77, [{ bound: { type: 'min' }, argb: P }, { bound: { type: 'max' }, argb: Q }], [7, 217], WITH_STAT)).toEqual({ kind: 'ready', value: interpolateScaleColor(P, Q, 1 / 3) })
    // Truncate (the other outcome) changes nothing.
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'truncate' })
    expect(scaleColorAt(77, stat, [7, 217, 637], WITH_STAT)).toEqual({ kind: 'ready', value: interpolateScaleColor(P, Q, 70 / 210) })
  })

  it('an unknown reading name is no reading (the shape stays withheld)', () => {
    injectCfNumSegmentReadings({ [CF_NUM_SEGMENT_CAP]: 'round-half-up', [CF_NUM_SEGMENT_STATISTICAL_CAP]: 'fixed-256' })
    expect(scaleColorAt(40, [num(0, RED), num(30, YELLOW), num(60, GREEN)], [0, 30, 60], evidence([CF_NUM_SEGMENT_CAP, CF_NUM_SEGMENT_STATISTICAL_CAP]))).toEqual(WITHHELD)
    const stat: CfScalePoint[] = [{ bound: { type: 'min' }, argb: RED }, { bound: { type: 'percentile', value: 50 }, argb: YELLOW }, { bound: { type: 'max' }, argb: GREEN }]
    expect(scaleColorAt(40, stat, [0, 30, 60], WITH_STAT)).toEqual({ kind: 'ready', value: interpolateScaleColor(YELLOW, GREEN, 1 / 3) })
  })
})
