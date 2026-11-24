// @vitest-environment node
/**
 * The literal-lower follow-up (brief cf-literal-lower), Phase 1 task 2: the readings of an interior cell of a 3-colour LOWER
 * segment bounded by two literal num cfvo where f·2^20 is not an integer (the num-segment shape, withheld today), implemented
 * and selectable through the test seam `injectCfLiteralLowerReadings` with the id's evidence, and registered nowhere:
 * - `lower-arithmetic`: the cell is painted exactly as a non-literal lower segment is (`guardedInterpolation`: truncation,
 *   raw floor at a guard-boundary product under its registration, every fail-closed condition kept);
 * - `step-plus-one`: T4/V3's behaviour (e448d831) transplanted: per channel the guarded truncation step minus one.
 * Unregistered, the shape stays withheld (`cf-unmeasured-num-segment-fraction`). The UPPER segment and every other shape are
 * untouched under every reading. The justifications (stage2NumSegmentShape, guardBoundaryShapeAt) narrow only with the
 * registration plus its evidence. This checks production, not Excel. Since the capture of stage2-scale-literal-lower
 * (a3f85309) matched lower-arithmetic, it is registered and published (docs/validation/xlsx-cf.md §14.z.21): the unregistered
 * state is pinned through an invalid proposal (`unregister`), and the registered one as committed.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_GUARD_BOUNDARY_CAP,
  CF_LITERAL_LOWER_CAP,
  CF_LITERAL_LOWER_PRODUCTION_READINGS,
  CF_LITERAL_LOWER_READINGS,
  CF_UPPER_WINDOW_CAP,
  cfLiteralLowerReading,
  injectCfGuardBoundaryReadings,
  injectCfLiteralLowerReadings,
  injectCfUpperWindowReadings,
  interpolateScaleColor,
  literalLowerStepColour,
  scaleColorAt,
  type CfLiteralLowerReading,
} from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { guardBoundaryShapeAt, type UpperWholeMarker } from '../scripts/cf-refs/upper-whole-justify'
import { stage2NumSegmentShape } from '../scripts/cf-refs/stage2-withhold'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

afterEach(() => { injectCfLiteralLowerReadings(null); injectCfGuardBoundaryReadings(null); injectCfUpperWindowReadings(null) })

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const P = 0xff14f028, Q = 0xffe61efa // P + 210·(1, −1, 1)
const Y = 0xffffeb84, G = 0xff63be7b // T4/V3's segment: spans 156, 45, 9, all descending from Y
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const numF = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value, resolved: 'formula' }, argb })
const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
const colour = (o: ReturnType<typeof scaleColorAt>): string => (o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind)
const NUMSEG = 'cf-unmeasured-num-segment-fraction'
const GUARD = 'cf-unmeasured-guard-boundary'
const register = (reading: CfLiteralLowerReading): void => injectCfLiteralLowerReadings({ [CF_LITERAL_LOWER_CAP]: reading })
/** The pre-capture state: no valid reading registered for the id. */
const unregister = (): void => injectCfLiteralLowerReadings({ [CF_LITERAL_LOWER_CAP]: 'unregistered' as CfLiteralLowerReading })

/** T4/V3 transplanted to the lower segment: num 0 / 30 / 60, Y / G / Y. 10 and 20 are f = 1/3 and 2/3 (the doubles of T4/V3). */
const TV = [num(0, Y), num(30, G), num(60, Y)]
const TVT = [0, 30, 60]
/** The P line: num 0 / 210 / 420, P / Q / P. */
const PL = [num(0, P), num(210, Q), num(420, P)]
const PLT = [0, 210, 420]
const at = (v: number, points: readonly CfScalePoint[], t: readonly number[], opts: Parameters<typeof scaleColorAt>[3] = { hasEvidence: all }): string => colour(scaleColorAt(v, points, t, opts))

describe('the readings: implemented, selectable; lower-arithmetic registered (a3f85309)', () => {
  it('lower-arithmetic and step-plus-one, the only readings; production registers lower-arithmetic for the question and nothing else', () => {
    expect([...CF_LITERAL_LOWER_READINGS]).toEqual(['lower-arithmetic', 'step-plus-one'])
    expect(CF_LITERAL_LOWER_CAP).toBe('scale:num-segment-literal-lower')
    expect(CF_LITERAL_LOWER_PRODUCTION_READINGS).toEqual({ [CF_LITERAL_LOWER_CAP]: 'lower-arithmetic' })
    expect(Object.isFrozen(CF_LITERAL_LOWER_PRODUCTION_READINGS)).toBe(true)
    expect(cfLiteralLowerReading()).toBe('lower-arithmetic')
    register('step-plus-one')
    expect(cfLiteralLowerReading()).toBe('step-plus-one')
    injectCfLiteralLowerReadings({ [CF_LITERAL_LOWER_CAP]: 'nonsense' as CfLiteralLowerReading })
    expect(cfLiteralLowerReading()).toBeUndefined()
    injectCfLiteralLowerReadings({ 'scale:other': 'step-plus-one' })
    expect(cfLiteralLowerReading()).toBe('lower-arithmetic')
  })

  it('step-plus-one reproduces T4/V3 (e448d831) and the literal upper-segment cells of 54fc3388 from their segments; guarded − 1, not raw floor − 1', () => {
    // T4, V3: the upper segment of num 0 / 30 / 60 (Y -> G), f = 10/30 and 20/30; Excel FFCCDD82 and FF98CE7F.
    expect(hex(literalLowerStepColour(Y, G, 10 / 30)!)).toBe('FFCCDD82')
    expect(hex(literalLowerStepColour(Y, G, 20 / 30)!)).toBe('FF98CE7F')
    expect(hex(interpolateScaleColor(Y, G, 10 / 30))).toBe('FFCBDC81')
    // 54fc3388 F12, F15, F16, F18 (products just below a whole number: raw floor − 1 would be one step more) and F19: Q -> P.
    const F: Array<[number, string]> = [[0.33333333333333326, 'FFA163B5'], [0.06666666666666664, 'FFD92BED'], [0.16666666666666663, 'FFC440D8'], [0.1999999999999999, 'FFBD47D1'], [0.06666666666666672, 'FFD92BED']]
    for (const [f, excel] of F) expect(hex(literalLowerStepColour(Q, P, f)!), String(f)).toBe(excel)
  })

  it('step-plus-one is undefined (withheld) where a channel\'s truncation step is 0, and on any non-finite fraction', () => {
    expect(literalLowerStepColour(Y, G, 1 / 30)).toBeUndefined() // blue: 9/30 -> step 0
    for (const f of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) expect(literalLowerStepColour(Y, G, f), String(f)).toBeUndefined()
  })
})

describe('scaleColorAt: default unchanged; a registered reading paints only the literal-num lower segment', () => {
  it('unregistered (the state before a3f85309): the lower-segment shape stays withheld whatever the evidence; painted cells and the upper segment as before', () => {
    unregister()
    for (const ev of [all, without(CF_LITERAL_LOWER_CAP)]) {
      expect(at(10, TV, TVT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(70, PL, PLT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(57, PL, PLT, { hasEvidence: ev })).toBe(NUMSEG)
      expect(at(15, TV, TVT, { hasEvidence: ev })).toBe('FFB1D580') // f = 1/2: f·2^20 an integer, painted
      expect(at(40, TV, TVT, { hasEvidence: ev })).toBe(NUMSEG) // T4 itself, upper
    }
  })

  it('registered lower-arithmetic (with evidence): truncation, raw floor at a guard-boundary product; every fail-closed condition kept; the upper segment untouched', () => {
    register('lower-arithmetic')
    expect(at(10, TV, TVT)).toBe('FFCBDC81')
    expect(at(20, TV, TVT)).toBe('FF97CD7E')
    expect(at(70, PL, PLT)).toBe('FF5AAA6E') // f·210 = 70 exactly
    expect(at(57, PL, PLT)).toBe('FF4CB860') // 56.99999999999999: raw floor under scale:guard-boundary-raw-floor
    // The guard-boundary reading's own evidence still gates its shape.
    expect(at(57, PL, PLT, { hasEvidence: without(CF_GUARD_BOUNDARY_CAP) })).toBe(GUARD)
    expect(at(70, PL, PLT, { hasEvidence: without(CF_GUARD_BOUNDARY_CAP) })).toBe('FF5AAA6E')
    // The midpoint-above-max re-sort (num 0 / 420 / 210): the lower segment 0 .. 210, P -> Q, bounded by cfvo 1 and 3.
    expect(at(70, [num(0, P), num(420, Q), num(210, P)], [0, 420, 210])).toBe('FF5AAA6E')
    // Without this id's evidence, under another colour space or rounding: withheld as before.
    expect(at(10, TV, TVT, { hasEvidence: without(CF_LITERAL_LOWER_CAP) })).toBe(NUMSEG)
    expect(at(10, TV, TVT, { hasEvidence: all, colorSpace: 'linear-light' })).toBe(NUMSEG)
    expect(at(10, TV, TVT, { hasEvidence: all, rounding: 'half-up' })).toBe(NUMSEG)
    // The upper segment keeps every withhold: T4, V3, and the P-line upper segment (a whole product, and a clear one).
    expect(at(40, TV, TVT)).toBe(NUMSEG)
    expect(at(50, TV, TVT)).toBe(NUMSEG)
    expect(at(280, PL, PLT)).toBe(NUMSEG)
    expect(at(267, PL, PLT)).toBe(NUMSEG)
    // Unchanged: f·2^20-integer cells, endpoints; a formula-resolved bound is not the literal shape.
    expect(at(15, TV, TVT)).toBe('FFB1D580')
    expect(at(30, TV, TVT)).toBe(hex(G))
    expect(at(0, TV, TVT)).toBe(hex(Y))
    expect(at(10, [numF(0, Y), num(30, G), num(60, Y)], TVT)).toBe('FFCBDC81')
  })

  it('registered step-plus-one (with evidence): the guarded step minus one per channel; a step-0 channel withholds; the upper segment untouched', () => {
    register('step-plus-one')
    expect(at(10, TV, TVT)).toBe('FFCCDD82')
    expect(at(20, TV, TVT)).toBe('FF98CE7F')
    expect(at(70, PL, PLT)).toBe('FF59AB6D')
    expect(at(57, PL, PLT)).toBe('FF4CB860') // guarded 57 − 1 = raw floor 56: the readings coincide at a below-integer product
    expect(at(1, TV, TVT)).toBe(NUMSEG)
    expect(at(10, TV, TVT, { hasEvidence: without(CF_LITERAL_LOWER_CAP) })).toBe(NUMSEG)
    expect(at(10, TV, TVT, { hasEvidence: all, colorSpace: 'hsl' })).toBe(NUMSEG)
    expect(at(40, TV, TVT)).toBe(NUMSEG)
    expect(at(280, PL, PLT)).toBe(NUMSEG)
    expect(at(15, TV, TVT)).toBe('FFB1D580')
  })

  it('as committed (lower-arithmetic registered, with evidence): T4/V3\'s lower-segment doubles paint truncation; without the id\'s evidence withheld', () => {
    expect(at(10, TV, TVT)).toBe('FFCBDC81')
    expect(at(70, PL, PLT)).toBe('FF5AAA6E')
    expect(at(10, TV, TVT, { hasEvidence: without(CF_LITERAL_LOWER_CAP) })).toBe(NUMSEG)
    expect(at(40, TV, TVT)).toBe(NUMSEG)
  })

  it('the registration is per id: the guard-boundary and upper-window ids never select it', () => {
    unregister()
    injectCfGuardBoundaryReadings({ [CF_GUARD_BOUNDARY_CAP]: 'raw-floor', [CF_LITERAL_LOWER_CAP]: 'raw-floor' } as never)
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'guarded' })
    expect(at(10, TV, TVT)).toBe(NUMSEG)
    expect(at(70, PL, PLT)).toBe(NUMSEG)
  })
})

describe('the justifications narrow with the registration plus its evidence, and only for the lower segment', () => {
  const values = new Map([['A2', 0], ['A3', 10], ['A4', 1], ['A5', 15], ['A6', 40], ['A7', 60]])
  const tv: UpperWholeMarker & { ruleType: string } = { ruleType: 'colorScale', sqref: 'A2:A7', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '30' }, { type: 'num', val: '60' }], colors: ['FFFFEB84', 'FF63BE7B', 'FFFFEB84'] } }
  const pl: UpperWholeMarker & { ruleType: string } = { ruleType: 'colorScale', sqref: 'A2:A7', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'num', val: '210' }, { type: 'num', val: '420' }], colors: ['FF14F028', 'FFE61EFA', 'FF14F028'] } }
  const caps = [...VERIFIED_CAPABILITIES, CF_LITERAL_LOWER_CAP]
  const shape = (m: typeof tv, v: number): boolean => stage2NumSegmentShape([m], 'A3', v)

  it('unregistered: every literal-num interior cell with f·2^20 not an integer is justified, lower and upper; f = 1/2 never', () => {
    unregister()
    withExactCfEvidence(caps, () => {
      for (const v of [10, 20, 1, 40, 50]) expect(shape(tv, v), String(v)).toBe(true)
      expect(shape(tv, 15)).toBe(false)
      expect(guardBoundaryShapeAt(pl, 'A3', 57, values)).toBe(false) // the num-segment shape, not this one
    })
  })

  for (const reading of CF_LITERAL_LOWER_READINGS) {
    it(`registered ${reading}: with evidence the lower segment is justified no longer (but where ${reading === 'step-plus-one' ? 'a step is 0' : 'the guard-boundary withhold applies'}); the upper segment always is; without the evidence nothing narrows`, () => {
      register(reading)
      withExactCfEvidence(caps, () => {
        expect(shape(tv, 10)).toBe(false)
        expect(shape(tv, 20)).toBe(false)
        expect(shape(tv, 1)).toBe(reading === 'step-plus-one')
        expect(shape(tv, 40)).toBe(true)
        expect(shape(tv, 50)).toBe(true)
        expect(shape(pl, 280)).toBe(true)
        // 57 on the P line: lower-arithmetic paints it through the guard-boundary reading (registered and published).
        expect(shape(pl, 57)).toBe(false)
        expect(guardBoundaryShapeAt(pl, 'A3', 57, values)).toBe(false)
      })
      // lower-arithmetic without the guard-boundary evidence: the cell is withheld as the guard-boundary shape, and justified so.
      withExactCfEvidence(caps.filter(c => c !== CF_GUARD_BOUNDARY_CAP), () => {
        expect(shape(pl, 57)).toBe(false)
        expect(guardBoundaryShapeAt(pl, 'A3', 57, values)).toBe(reading === 'lower-arithmetic')
      })
      withExactCfEvidence([...VERIFIED_CAPABILITIES].filter(c => c !== CF_LITERAL_LOWER_CAP), () => {
        expect(shape(tv, 10)).toBe(true)
        expect(guardBoundaryShapeAt(pl, 'A3', 57, values)).toBe(false)
      })
    })
  }
})
