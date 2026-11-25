// @vitest-environment node
/**
 * The upper-segment follow-up (brief cf-upper-segment), Phase 1b task 2: the readings of an interior cell of a 3-colour UPPER
 * segment that production withholds today (the literal num-segment shape in segment 1; the `integer` and `window` classes of
 * the whole-product shape), implemented and selectable through the test seam `injectCfUpperSegmentReadings` with the id's
 * evidence, and registered nowhere:
 * - `index-position`: l = (1 + f) − 1, then floor(l·|d|) per channel (the only survivor class of
 *   scripts/cf-refs/upper-segment-candidates.ts over all 302 captured upper-segment interior cells);
 * - `truncate`: production's guarded truncation (the control).
 * Unregistered, every such cell stays withheld as before. Since the capture of stage2-scale-upper-segment (2ee3735a) matched
 * index-position for both ids, it is registered and published: the unregistered state is pinned through an invalid proposal
 * (`unregisterUpperSegment`), the registered one as committed. A formula-resolved bound, another colour space or rounding, the
 * lower segment and 2-colour scales are untouched under every reading. The justifications (`upperWholeShapeAt`,
 * `stage2NumSegmentShape`) narrow only with the registration plus its evidence. This checks production, not Excel.
 */
import { afterEach, describe, expect, it, beforeAll, afterAll } from 'vitest'
import { asBeforePercentOrder, asRegistered } from './helpers/cf-percent-order-asof'
import {
  CF_UPPER_NEAR_TOP_CAP,
  CF_UPPER_NEAR_TOP_PRODUCTION_READINGS,
  CF_UPPER_SEGMENT_CAP,
  CF_UPPER_SEGMENT_PRODUCTION_READINGS,
  CF_UPPER_SEGMENT_READINGS,
  CF_UPPER_WINDOW_CAP,
  cfUpperSegmentReading,
  injectCfUpperNearTopReadings,
  injectCfUpperSegmentReadings,
  injectCfUpperWindowReadings,
  interpolateScaleColor,
  scaleColorAt,
  upperSegmentNearTop,
  upperSegmentPositionColour,
  type CfUpperSegmentReading,
} from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { upperNearTopShapeAt, upperWholeShapeAt, type UpperWholeMarker } from '../scripts/cf-refs/upper-whole-justify'
import { stage2NumSegmentShape } from '../scripts/cf-refs/stage2-withhold'
import { unregisterUpperSegment } from './helpers/cf-upper-segment-recorded'

// As of its capture: before the percent-order registration (§14.z.24), percent thresholds lerp-first and no
// scale:formula-bounds-v2 reading (tests/helpers/cf-percent-order-asof.ts).
beforeAll(asBeforePercentOrder)
afterAll(asRegistered)

afterEach(() => { injectCfUpperSegmentReadings(null); injectCfUpperNearTopReadings(null); injectCfUpperWindowReadings(null) })

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const P = 0xff14f028, Q = 0xffe61efa // P + 210·(1, −1, 1)
const Y = 0xffffeb84, G = 0xff63be7b // T4/V3's segment: spans 156, 45, 9, all descending from Y
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const numF = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value, resolved: 'formula' }, argb })
const pct = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'percent', value }, argb })
const mx = (argb: number): CfScalePoint => ({ bound: { type: 'max' }, argb })
const mn = (argb: number): CfScalePoint => ({ bound: { type: 'min' }, argb })
const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
const colour = (o: ReturnType<typeof scaleColorAt>): string => (o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind)
const NUMSEG = 'cf-unmeasured-num-segment-fraction'
const WHOLE = 'cf-unmeasured-upper-segment-whole-product'
const register = (reading: CfUpperSegmentReading): void => injectCfUpperSegmentReadings({ [CF_UPPER_SEGMENT_CAP]: reading })

/** num 0 / 30 / 60, Y / G / Y: 40 and 50 are T4/V3's values and doubles (f = 1/3, 2/3), the upper segment here running G -> Y. */
const TV = [num(0, Y), num(30, G), num(60, Y)]
const TVT = [0, 30, 60]
/** The P line: num 0 / 210 / 420, P / Q / P (upper segment Q -> P). */
const PL = [num(0, P), num(210, Q), num(420, P)]
const PLT = [0, 210, 420]
/** Statistical bounds: min / percent 50 / max over 0 .. 420 (the midpoint 210). */
const ST = [mn(P), pct(50, Q), mx(P)]
const at = (v: number, points: readonly CfScalePoint[], t: readonly number[], opts: Parameters<typeof scaleColorAt>[3] = { hasEvidence: all }): string => colour(scaleColorAt(v, points, t, opts))

describe('the readings: implemented, selectable; index-position registered (2ee3735a)', () => {
  it('index-position and truncate, the only readings; production registers index-position for the question and nothing else', () => {
    expect([...CF_UPPER_SEGMENT_READINGS]).toEqual(['index-position', 'truncate'])
    expect(CF_UPPER_SEGMENT_CAP).toBe('scale:upper-segment')
    expect(CF_UPPER_SEGMENT_PRODUCTION_READINGS).toEqual({ [CF_UPPER_SEGMENT_CAP]: 'index-position' })
    expect(Object.isFrozen(CF_UPPER_SEGMENT_PRODUCTION_READINGS)).toBe(true)
    expect(cfUpperSegmentReading()).toBe('index-position')
    register('truncate')
    expect(cfUpperSegmentReading()).toBe('truncate')
    injectCfUpperSegmentReadings({ [CF_UPPER_SEGMENT_CAP]: 'nonsense' as CfUpperSegmentReading })
    expect(cfUpperSegmentReading()).toBeUndefined()
    injectCfUpperSegmentReadings({ 'scale:other': 'truncate' })
    expect(cfUpperSegmentReading()).toBe('index-position')
  })

  it('index-position reproduces T4/V3 (e448d831) and the 54fc3388 cells F12, F15, F16, F18, F19 from their segments; it keeps 1/6, 1/10 and 127/210', () => {
    expect(hex(upperSegmentPositionColour(Y, G, 10 / 30)!)).toBe('FFCCDD82')
    expect(hex(upperSegmentPositionColour(Y, G, 20 / 30)!)).toBe('FF98CE7F')
    const F: Array<[number, string]> = [[0.33333333333333326, 'FFA163B5'], [0.06666666666666664, 'FFD92BED'], [0.16666666666666663, 'FFC440D8'], [0.1999999999999999, 'FFBD47D1'], [0.06666666666666672, 'FFD92BED']]
    for (const [f, excel] of F) expect(hex(upperSegmentPositionColour(Q, P, f)!), String(f)).toBe(excel)
    for (const f of [35 / 210, 21 / 210, 127 / 210]) expect(upperSegmentPositionColour(Q, P, f), String(f)).toBe(interpolateScaleColor(Q, P, f))
    // On the 2^-52 grid (1/4) l = f: truncation.
    expect(upperSegmentPositionColour(Q, P, 0.25)).toBe(interpolateScaleColor(Q, P, 0.25))
  })

  it('index-position fails closed on a non-finite fraction; f is clamped to [0, 1]', () => {
    for (const f of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) expect(upperSegmentPositionColour(Q, P, f), String(f)).toBeUndefined()
    expect(upperSegmentPositionColour(Q, P, 1.5)).toBe(P >>> 0)
    expect(upperSegmentPositionColour(Q, P, -0.5)).toBe(Q >>> 0)
  })
})

describe('scaleColorAt: default unchanged; a registered reading paints only the upper-segment shapes withheld today', () => {
  it('unregistered (the state before 2ee3735a): the literal shape, the integer class and the window class stay withheld whatever the evidence', () => {
    unregisterUpperSegment()
    for (const ev of [all, without(CF_UPPER_SEGMENT_CAP)]) {
      expect(at(40, TV, TVT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(280, PL, PLT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(267, PL, PLT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(280, ST, PLT, { hasEvidence: ev })).toBe(WHOLE) // integer class, statistical bounds
      expect(at(315, PL, PLT, { hasEvidence: ev })).toBe(WHOLE) // f = 1/2: f·2^20 integer, the integer class
      expect(at(262.5, PL, PLT, { hasEvidence: ev })).toBe('FFB252C6') // f = 1/4: 52.5 steps, clear: painted
    }
  })

  for (const reading of CF_UPPER_SEGMENT_READINGS) {
    it(`registered ${reading} (with evidence): the three withheld shapes paint ${reading}; painted cells, the lower segment, 2-colour scales and every fail-closed condition unchanged`, () => {
      register(reading)
      const want = (a: number, b: number, f: number): string => hex(reading === 'index-position' ? upperSegmentPositionColour(a, b, f)! : interpolateScaleColor(a, b, f))
      // The literal num-segment shape (segment 1): T4/V3's doubles, a P-line whole product and a window product.
      expect(at(40, TV, TVT)).toBe(want(G, Y, 1 / 3))
      expect(at(50, TV, TVT)).toBe(want(G, Y, 2 / 3))
      expect(at(280, PL, PLT)).toBe(want(Q, P, 70 / 210))
      expect(at(267, PL, PLT)).toBe(want(Q, P, 57 / 210))
      // The integer class under statistical bounds, and f = 1/2 (on the grid: both readings truncate).
      expect(at(280, ST, PLT)).toBe(want(Q, P, 70 / 210))
      expect(at(315, PL, PLT)).toBe(hex(interpolateScaleColor(Q, P, 0.5)))
      // The midpoint-above-max re-sort (num 0 / 420 / 210): the upper segment 210 .. 420, colours 2 -> 3.
      expect(at(280, [num(0, P), num(420, Q), num(210, P)], [0, 420, 210])).toBe(want(Q, P, 70 / 210))
      // Unchanged: a clear upper cell, endpoints, the lower segment and a 2-colour scale.
      expect(at(262.5, PL, PLT)).toBe('FFB252C6')
      expect(at(210, PL, PLT)).toBe(hex(Q))
      expect(at(420, PL, PLT)).toBe(hex(P))
      expect(at(70, PL, PLT)).toBe('FF5AAA6E')
      expect(at(70, [num(0, P), num(210, Q)], [0, 210])).toBe('FF5AAA6E')
      // Fail closed: without the id's evidence, under another colour space or rounding, with a formula-resolved bound.
      expect(at(280, PL, PLT, { hasEvidence: without(CF_UPPER_SEGMENT_CAP) })).toBe(NUMSEG)
      expect(at(280, ST, PLT, { hasEvidence: without(CF_UPPER_SEGMENT_CAP) })).toBe(WHOLE)
      expect(at(280, PL, PLT, { hasEvidence: all, colorSpace: 'linear-light' })).toBe(NUMSEG)
      expect(at(280, ST, PLT, { hasEvidence: all, rounding: 'half-up' })).not.toBe(want(Q, P, 70 / 210))
      expect(at(280, [num(0, P), numF(210, Q), num(420, P)], PLT)).toBe(WHOLE)
      expect(at(280, [mn(P), pct(50, Q), { bound: { type: 'max', resolved: 'formula' } as never, argb: P }], PLT)).toBe(WHOLE)
    })
  }

  it('the precedence over the window reading: a registered upper-segment reading decides the window class first', () => {
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'raw-floor' })
    // 0.44 in 0.4 .. 1 under statistical bounds: a window product (f·210 = 13.999999999999995).
    const st = [mn(P), pct(50, Q), mx(P)], t = [0.4 - 0.6, 0.4, 1]
    const f = (0.44 - 0.4) / 0.6
    const rawFloor = at(0.44, st, t)
    register('truncate')
    expect(at(0.44, st, t)).toBe(hex(interpolateScaleColor(Q, P, f)))
    expect(at(0.44, st, t)).not.toBe(rawFloor)
    register('index-position')
    expect(at(0.44, st, t)).toBe(hex(upperSegmentPositionColour(Q, P, f)!))
  })

  it('the registration is per id: the window id never selects it', () => {
    unregisterUpperSegment()
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'guarded', [CF_UPPER_SEGMENT_CAP]: 'guarded' } as never)
    expect(at(280, PL, PLT)).toBe(NUMSEG)
    expect(at(280, ST, PLT)).toBe(WHOLE)
  })
})

describe('the justifications narrow with the registration plus its evidence, and only for the upper segment', () => {
  const pl: UpperWholeMarker = { ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '210' }, { type: 'num', val: '420' }], colors: ['FF14F028', 'FFE61EFA', 'FF14F028'] } }
  const st: UpperWholeMarker = { ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'min' }, { type: 'percent', val: '50' }, { type: 'max' }], colors: ['FF14F028', 'FFE61EFA', 'FF14F028'] } }
  const fm: UpperWholeMarker = { ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'min' }, { type: 'num', val: '210', f: 'B1' }, { type: 'max' }], colors: ['FF14F028', 'FFE61EFA', 'FF14F028'] } }
  const values = new Map([['A2', 0], ['A3', 280], ['A4', 70], ['A9', 420]])
  const caps = [...VERIFIED_CAPABILITIES, CF_UPPER_SEGMENT_CAP]
  const lit = (v: number): boolean => stage2NumSegmentShape([pl], 'A3', v)
  const whole = (m: UpperWholeMarker, v: number): boolean => upperWholeShapeAt(m, 'A3', v, values, () => 210)

  it('unregistered: the upper literal shape and the whole-product shape are justified; the lower literal shape is the literal-lower registration\'s', () => {
    unregisterUpperSegment()
    withExactCfEvidence(caps, () => {
      expect(lit(280)).toBe(true)
      expect(lit(267)).toBe(true)
      expect(whole(st, 280)).toBe(true)
      expect(whole(fm, 280)).toBe(true)
      expect(lit(70)).toBe(false) // lower-arithmetic is registered and published
    })
  })

  for (const reading of CF_UPPER_SEGMENT_READINGS) {
    it(`registered ${reading}: with evidence the upper shapes are justified no longer, but a formula-bounded one; without the evidence nothing narrows`, () => {
      register(reading)
      withExactCfEvidence(caps, () => {
        expect(lit(280)).toBe(false)
        expect(lit(267)).toBe(false)
        expect(whole(st, 280)).toBe(false)
        expect(whole(fm, 280)).toBe(true)
        expect(lit(70)).toBe(false)
      })
      withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== CF_UPPER_SEGMENT_CAP && id !== CF_UPPER_NEAR_TOP_CAP), () => {
        expect(lit(280)).toBe(true)
        expect(whole(st, 280)).toBe(true)
      })
    })
  }
})

describe('the near-top sub-shape (orchestrator ruling before the capture): withheld unless both ids register the same reading', () => {
  const NEARTOP = 'cf-unmeasured-upper-near-top'
  const top = 426.99999999999994 // one ulp below 427 in num 7 / 217 / 427 (upper Q -> P, span 210): f·210 = 209.99999999999994
  const NT = [num(7, P), num(217, Q), num(427, P)], NTT = [7, 217, 427]
  const st = [mn(P), pct(50, Q), mx(P)]
  const allNear = (cap: string): boolean => cap.length > 0
  it('the predicate: within 1e-11 below the full span on some channel, strictly inside; not 1e-9 below, not the endpoint, not the lower segment', () => {
    const f = (v: number): number => (v - 217) / 210
    expect(upperSegmentNearTop(Q, P, top, 217, 427, f(top))).toBe(true)
    expect(upperSegmentNearTop(Q, P, 427 - 1e-9, 217, 427, f(427 - 1e-9))).toBe(false)
    expect(upperSegmentNearTop(Q, P, 427, 217, 427, 1)).toBe(false)
    expect(upperSegmentNearTop(Q, P, 0.9999999999999999, -1, 1, 0.9999999999999999)).toBe(true)
    expect(CF_UPPER_NEAR_TOP_CAP).toBe('scale:upper-segment-near-top')
    expect(CF_UPPER_NEAR_TOP_PRODUCTION_READINGS).toEqual({ [CF_UPPER_NEAR_TOP_CAP]: 'index-position' })
  })
  it('unregistered: withheld (it was painted at the maximum colour before the ruling); with the main id alone or a different near-top reading: still withheld; with both at the same reading: the reading', () => {
    unregisterUpperSegment()
    expect(at(top, NT, NTT)).toBe(NEARTOP)
    expect(at(top, st, [0, 210, 420].map(x => x + 7))).toBe(NEARTOP)
    for (const reading of CF_UPPER_SEGMENT_READINGS) {
      const other = CF_UPPER_SEGMENT_READINGS.find(x => x !== reading)!
      register(reading)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'unregistered' as never })
      expect(at(top, NT, NTT, { hasEvidence: allNear })).toBe(NEARTOP)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: other })
      expect(at(top, NT, NTT, { hasEvidence: allNear })).toBe(NEARTOP)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: reading })
      const f = (top - 217) / 210
      expect(at(top, NT, NTT, { hasEvidence: allNear })).toBe(hex(reading === 'index-position' ? upperSegmentPositionColour(Q, P, f)! : interpolateScaleColor(Q, P, f)))
      expect(at(top, NT, NTT, { hasEvidence: without(CF_UPPER_NEAR_TOP_CAP) })).toBe(NEARTOP)
      expect(at(top, [num(7, P), numF(217, Q), num(427, P)], NTT, { hasEvidence: allNear })).toBe(NEARTOP)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'unregistered' as never })
    }
    // index-position one step below the maximum colour, truncate the maximum colour.
    expect(hex(upperSegmentPositionColour(Q, P, (top - 217) / 210)!)).toBe('FF15EF29')
    expect(hex(interpolateScaleColor(Q, P, (top - 217) / 210))).toBe('FF14F028')
  })
  it('the justification: the shape is justified unless both ids register the same reading with their evidence and no bound is a formula', () => {
    const m: UpperWholeMarker = { ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'num', val: '7' }, { type: 'num', val: '217' }, { type: 'num', val: '427' }], colors: ['FF14F028', 'FFE61EFA', 'FF14F028'] } }
    const vals = new Map([['A3', top]])
    const caps = [...VERIFIED_CAPABILITIES, CF_UPPER_SEGMENT_CAP, CF_UPPER_NEAR_TOP_CAP]
    unregisterUpperSegment()
    withExactCfEvidence(caps, () => {
      expect(upperNearTopShapeAt(m, 'A3', top, vals)).toBe(true)
      expect(upperNearTopShapeAt(m, 'A3', 300, vals)).toBe(false)
      register('index-position')
      expect(upperNearTopShapeAt(m, 'A3', top, vals)).toBe(true)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'truncate' })
      expect(upperNearTopShapeAt(m, 'A3', top, vals)).toBe(true)
      injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'index-position' })
      expect(upperNearTopShapeAt(m, 'A3', top, vals)).toBe(false)
    })
    withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(id => id !== CF_UPPER_NEAR_TOP_CAP), () => expect(upperNearTopShapeAt(m, 'A3', top, vals)).toBe(true))
  })
})

describe('codex pre-capture review: index-position only between opaque endpoint colours (fail closed on alpha)', () => {
  const K = 0xff000000, T = 0x000000ff // FF000000 -> 000000FF: alpha and blue both span 255
  const AL = [num(0, K), num(13, K), num(26, T)], ALT = [0, 13, 26]
  const v = 13 + 10 // 10/13 in the upper segment: f·255 = 196.15..., clear; 10/11-like products elsewhere
  it('unregistered the shape is withheld as before; with index-position registered a non-opaque endpoint stays withheld (the arithmetic would have painted alpha)', () => {
    const f = 10 / 13
    expect(hex(upperSegmentPositionColour(K, T, f)!)).not.toMatch(/^FF/) // the hazard: alpha interpolated
    expect(at(v, AL, ALT)).toBe(NUMSEG)
    register('index-position')
    expect(at(v, AL, ALT)).toBe(NUMSEG)
    const st = [mn(K), pct(50, K), mx(T)]
    expect(at(280, [mn(K), pct(50, K), mx(0x00ffffff)], [0, 210, 420])).toBe(WHOLE)
    expect(at(280, st, [0, 210, 420])).not.toMatch(/^[0-9A-F]{8}$/)
    // Opaque endpoints unchanged.
    expect(at(280, PL, PLT)).toBe(hex(upperSegmentPositionColour(Q, P, 70 / 210)!))
  })
  it('truncate (the control, production\'s own arithmetic) still paints a non-opaque segment when registered', () => {
    register('truncate')
    expect(at(v, AL, ALT)).toBe(hex(interpolateScaleColor(K, T, 10 / 13)))
  })
  it('the near-top path: a non-opaque segment stays withheld under index-position on both ids', () => {
    register('index-position')
    injectCfUpperNearTopReadings({ [CF_UPPER_NEAR_TOP_CAP]: 'index-position' })
    const top = 25.999999999999996
    expect(at(top, AL, ALT, { hasEvidence: all })).toBe('cf-unmeasured-upper-near-top')
  })
  it('the justifications recompute it: under index-position a non-opaque upper segment stays justified, an opaque one is not', () => {
    const m = (to: string): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '210' }, { type: 'num', val: '420' }], colors: ['FF14F028', 'FFE61EFA', to] } })
    const st = (to: string): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A2:A9', scale: { cfvo: [{ type: 'min' }, { type: 'percent', val: '50' }, { type: 'max' }], colors: ['FF14F028', 'FFE61EFA', to] } })
    const vals = new Map([['A2', 0], ['A3', 280], ['A9', 420]])
    register('index-position')
    withExactCfEvidence([...VERIFIED_CAPABILITIES, CF_UPPER_SEGMENT_CAP], () => {
      expect(stage2NumSegmentShape([m('0014F028')], 'A3', 267)).toBe(true)
      expect(stage2NumSegmentShape([m('FF14F028')], 'A3', 267)).toBe(false)
      expect(upperWholeShapeAt(st('0014F028'), 'A3', 280, vals)).toBe(true)
      expect(upperWholeShapeAt(st('FF14F028'), 'A3', 280, vals)).toBe(false)
    })
  })
})

describe('as committed (index-position registered for both ids, published)', () => {
  it('T4/V3\'s doubles, a P-line whole product, a window product, a statistical integer product and a near-top cell paint index-position; without the evidence they are withheld', () => {
    const ev = (cap: string): boolean => VERIFIED_CAPABILITIES.has(cap)
    expect(at(280, PL, PLT, { hasEvidence: ev })).toBe(hex(upperSegmentPositionColour(Q, P, 70 / 210)!))
    expect(at(267, PL, PLT, { hasEvidence: ev })).toBe(hex(upperSegmentPositionColour(Q, P, 57 / 210)!))
    expect(at(280, ST, PLT, { hasEvidence: ev })).toBe(hex(upperSegmentPositionColour(Q, P, 70 / 210)!))
    expect(at(40, TV, TVT, { hasEvidence: ev })).toBe(hex(upperSegmentPositionColour(G, Y, 1 / 3)!))
    expect(at(426.99999999999994, [num(7, P), num(217, Q), num(427, P)], [7, 217, 427], { hasEvidence: ev })).toBe('FF15EF29')
    expect(at(280, PL, PLT, { hasEvidence: c => ev(c) && c !== CF_UPPER_SEGMENT_CAP })).toBe(NUMSEG)
    expect(at(280, ST, PLT, { hasEvidence: c => ev(c) && c !== CF_UPPER_SEGMENT_CAP })).toBe(WHOLE)
    expect(at(426.99999999999994, [num(7, P), num(217, Q), num(427, P)], [7, 217, 427], { hasEvidence: c => ev(c) && c !== CF_UPPER_NEAR_TOP_CAP })).toBe('cf-unmeasured-upper-near-top')
  })
})
