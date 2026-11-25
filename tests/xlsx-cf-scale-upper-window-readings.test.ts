// @vitest-environment node
/**
 * The upper-segment window follow-up (brief cf-upperwindow), Phase 1 task 2: the two window readings (`guarded`,
 * `raw-floor`), implemented and selectable through the test seam `injectCfUpperWindowReadings` with the id's evidence, and
 * registered nowhere. Default production is exactly main: every upper-segment cell of a 3-colour scale in the window class
 * (a product within 1e-11 of a whole number, no IEEE integer) is withheld. A registered reading narrows that withhold, and
 * its justification (scripts/cf-refs/upper-whole-justify.ts), to segments not bounded by two literal num cfvo; the integer
 * class and the literal-num shape stay withheld. This checks production, not Excel.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_UPPER_WINDOW_CAP,
  CF_UPPER_WINDOW_PRODUCTION_READINGS,
  CF_UPPER_WINDOW_READINGS,
  cfUpperWindowReading,
  injectCfUpperWindowReadings,
  interpolateScaleColor,
  scaleColorAt,
  upperSegmentProductClass,
  upperSegmentWholeProduct,
  upperWindowReadingColor,
  type CfUpperWindowReading,
} from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { upperWholeShapeAt, type UpperWholeMarker } from '../scripts/cf-refs/upper-whole-justify'
import { recordedBeforeUpperSegmentRegistration } from './helpers/cf-upper-segment-recorded'

// These tests record production as of their own captures, before the upper-segment registration (2ee3735a).
recordedBeforeUpperSegmentRegistration()

afterEach(() => injectCfUpperWindowReadings(null))

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const STOCK = [0xfff8696b, 0xffffeb84, 0xff63be7b] as const
const P = 0xff14f028, P210 = 0xffe61efa
const stat = (argb: readonly number[]): CfScalePoint[] => [{ bound: { type: 'min' }, argb: argb[0]! }, { bound: { type: 'percentile', value: 50 }, argb: argb[1]! }, { bound: { type: 'max' }, argb: argb[2]! }]
const literal = (t: readonly number[], argb: readonly number[]): CfScalePoint[] => t.map((v, i) => ({ bound: { type: 'num', value: v }, argb: argb[i]! }))
const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
type Outcome = ReturnType<typeof scaleColorAt>
const colour = (o: Outcome): string => (o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind)

describe('the window class (upperSegmentProductClass)', () => {
  it('splits the widened shape into the integer class and the window class; the widened predicate is their union', () => {
    // E4 of stage2-scale-thresholds: 15.999999999999996 on R (d = 156), no IEEE integer.
    const e4 = (3 - 2.2) / (10 - 2.2)
    expect(upperSegmentProductClass(STOCK[1], STOCK[2], 3, 2.2, 10, e4)).toBe('window')
    // F19 of stage2-scale-upper-whole: 14.00000000000001 (above an integer).
    const f19 = (1.36 - 1.3) / (2.2 - 1.3)
    expect(upperSegmentProductClass(P210, P, 1.36, 1.3, 2.2, f19)).toBe('window')
    // An integer product (1/3 of 210): the integer class, whatever the other channels say.
    expect(upperSegmentProductClass(P210, P, 357, 217, 637, (357 - 217) / 420)).toBe('integer')
    // Not whole; or an endpoint.
    expect(upperSegmentProductClass(P210, P, 300, 217, 637, (300 - 217) / 420)).toBeUndefined()
    expect(upperSegmentProductClass(P210, P, 217, 217, 637, 0)).toBeUndefined()
    for (const [v, lo, hi, a, b] of [[3, 2.2, 10, STOCK[1], STOCK[2]], [1.36, 1.3, 2.2, P210, P], [357, 217, 637, P210, P], [300, 217, 637, P210, P]] as const) {
      const f = (v - lo) / (hi - lo)
      expect(upperSegmentWholeProduct(a, b, v, lo, hi, f), `${v}`).toBe(upperSegmentProductClass(a, b, v, lo, hi, f) !== undefined)
    }
  })
})

describe('the readings: implemented, selectable, registered nowhere', () => {
  it('guarded and raw-floor, the only readings; production registers none, so the default is unregistered', () => {
    expect([...CF_UPPER_WINDOW_READINGS]).toEqual(['guarded', 'raw-floor'])
    expect(CF_UPPER_WINDOW_CAP).toBe('scale:upper-window-statistical')
    expect(CF_UPPER_WINDOW_PRODUCTION_READINGS).toEqual({})
    expect(cfUpperWindowReading()).toBeUndefined()
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'raw-floor' })
    expect(cfUpperWindowReading()).toBe('raw-floor')
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'nonsense' as CfUpperWindowReading })
    expect(cfUpperWindowReading()).toBeUndefined()
  })

  it('upperWindowReadingColor: guarded is production\'s interpolation; raw-floor is one smaller exactly where the product is just below an integer', () => {
    const e4 = (3 - 2.2) / (10 - 2.2)
    expect(hex(upperWindowReadingColor('guarded', STOCK[1], STOCK[2], e4)!)).toBe('FFEFE784')
    expect(hex(upperWindowReadingColor('raw-floor', STOCK[1], STOCK[2], e4)!)).toBe('FFF0E784')
    const f12 = (0.6 - 0.4) / (1 - 0.4)
    expect(hex(upperWindowReadingColor('guarded', P210, P, f12)!)).toBe('FFA064B4')
    expect(hex(upperWindowReadingColor('raw-floor', P210, P, f12)!)).toBe('FFA163B5')
    const f19 = (1.36 - 1.3) / (2.2 - 1.3)
    expect(upperWindowReadingColor('raw-floor', P210, P, f19)).toBe(upperWindowReadingColor('guarded', P210, P, f19))
    for (let i = 0; i <= 400; i++) {
      const f = i / 400
      for (const [a, b] of [[P, P210], [P210, P], [STOCK[0], STOCK[1]], [STOCK[1], STOCK[2]]] as const) expect(upperWindowReadingColor('guarded', a, b, f), `${f}`).toBe(interpolateScaleColor(a, b, f))
    }
  })

  it('fails closed on a non-finite fraction', () => {
    for (const r of CF_UPPER_WINDOW_READINGS) for (const f of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) expect(upperWindowReadingColor(r, P210, P, f), `${r} ${f}`).toBeUndefined()
  })
})

describe('scaleColorAt: default unchanged; a registered reading narrows only the statistical window class', () => {
  // E4's shape (min / percentile / max resolving to 1 / 2.2 / 10, value 3), F12's shape under statistical bounds (0.1 / 0.4 /
  // 1, value 0.6), an integer product (7 / 217 / 637, value 357) and a non-whole one (value 300).
  const cases = {
    e4: { v: 3, t: [1, 2.2, 10], argb: STOCK },
    f12: { v: 0.6, t: [0.1, 0.4, 1], argb: [P, P210, P] },
    integer: { v: 357, t: [7, 217, 637], argb: [P, P210, P] },
    plain: { v: 300, t: [7, 217, 637], argb: [P, P210, P] },
    // 0.31: f·210 = 146.99999999999997, the guard-boundary shape (owner ruling after ed0fccb6); 0.33: 161, not that shape.
    lower: { v: 0.31, t: [0.1, 0.4, 1], argb: [P, P210, P] },
    lowerPlain: { v: 0.33, t: [0.1, 0.4, 1], argb: [P, P210, P] },
  } as const
  const run = (k: keyof typeof cases, kind: 'stat' | 'literal', hasEvidence = all): string => {
    const c = cases[k]
    return colour(scaleColorAt(c.v, kind === 'stat' ? stat(c.argb) : literal(c.t, c.argb), c.t, { hasEvidence }))
  }

  it('unregistered (main): the window class is withheld for every bound kind; the rest paints as before', () => {
    expect(run('e4', 'stat')).toBe('cf-unmeasured-upper-segment-whole-product')
    expect(run('f12', 'stat')).toBe('cf-unmeasured-upper-segment-whole-product')
    expect(run('e4', 'literal')).toBe('cf-unmeasured-num-segment-fraction')
    expect(run('integer', 'stat')).toBe('cf-unmeasured-upper-segment-whole-product')
    expect(run('plain', 'stat')).toBe(hex(interpolateScaleColor(P210, P, (300 - 217) / 420)))
    expect(run('lower', 'stat', without('scale:guard-boundary-raw-floor'))).toBe('cf-unmeasured-guard-boundary')
    expect(run('lower', 'stat')).toBe('FFA65EBA') // raw floor (146): the guard-boundary registration of ce3f701d, not this id's
    expect(run('lowerPlain', 'stat')).toBe(hex(interpolateScaleColor(P, P210, (0.33 - 0.1) / (0.4 - 0.1))))
  })

  for (const reading of CF_UPPER_WINDOW_READINGS) {
    it(`registered ${reading} (with evidence): the statistical window class paints the reading; integer, literal and no-evidence stay withheld`, () => {
      injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: reading })
      expect(run('e4', 'stat')).toBe(reading === 'guarded' ? 'FFEFE784' : 'FFF0E784')
      expect(run('f12', 'stat')).toBe(reading === 'guarded' ? 'FFA064B4' : 'FFA163B5')
      expect(run('e4', 'literal')).toBe('cf-unmeasured-num-segment-fraction')
      expect(run('integer', 'stat')).toBe('cf-unmeasured-upper-segment-whole-product')
      expect(run('plain', 'stat')).toBe(hex(interpolateScaleColor(P210, P, (300 - 217) / 420)))
      expect(run('lower', 'stat', without('scale:guard-boundary-raw-floor'))).toBe('cf-unmeasured-guard-boundary')
      expect(run('lower', 'stat')).toBe('FFA65EBA') // raw floor (146): the guard-boundary registration of ce3f701d, not this id's
      expect(run('lowerPlain', 'stat')).toBe(hex(interpolateScaleColor(P, P210, (0.33 - 0.1) / (0.4 - 0.1))))
      // Without the id's evidence the reading is not honoured.
      expect(run('e4', 'stat', without(CF_UPPER_WINDOW_CAP))).toBe('cf-unmeasured-upper-segment-whole-product')
      // A non-default variant keeps the withhold.
      expect(colour(scaleColorAt(3, stat(STOCK), [1, 2.2, 10], { hasEvidence: all, rounding: 'half-up' }))).toBe('cf-unmeasured-upper-segment-whole-product')
      // A mixed bound (num lo, max hi) is statistical; a formula-resolved num bound is too.
      expect(colour(scaleColorAt(3, [{ bound: { type: 'min' }, argb: STOCK[0] }, { bound: { type: 'num', value: 2.2 }, argb: STOCK[1] }, { bound: { type: 'max' }, argb: STOCK[2] }], [1, 2.2, 10], { hasEvidence: all }))).toBe(reading === 'guarded' ? 'FFEFE784' : 'FFF0E784')
      expect(colour(scaleColorAt(3, [{ bound: { type: 'num', value: 1 }, argb: STOCK[0] }, { bound: { type: 'num', value: 2.2, resolved: 'formula' } as CfScalePoint['bound'], argb: STOCK[1] }, { bound: { type: 'num', value: 10 }, argb: STOCK[2] }], [1, 2.2, 10], { hasEvidence: all }))).toBe(reading === 'guarded' ? 'FFEFE784' : 'FFF0E784')
    })
  }
})

describe('the justification (upper-whole-justify.ts) narrows with the registration, and only then', () => {
  const values = new Map([['A2', 1], ['A3', 2], ['A4', 3], ['A5', 4], ['A6', 10]])
  const marker = (cfvo: ReadonlyArray<{ type: string; val?: string }>): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A2:A6', scale: { cfvo, colors: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'] } })
  const statistical = marker([{ type: 'min' }, { type: 'percentile', val: '30' }, { type: 'max' }])
  const literalNum = marker([{ type: 'num', val: '1' }, { type: 'num', val: '2.2' }, { type: 'num', val: '10' }])
  const caps = [...VERIFIED_CAPABILITIES, CF_UPPER_WINDOW_CAP]

  it('unregistered: E4 (window) and D20-like 4 (integer) are justified under both bound kinds', () => {
    for (const m of [statistical, literalNum]) {
      expect(upperWholeShapeAt(m, 'A4', 3, values)).toBe(true)
      expect(upperWholeShapeAt(m, 'A5', 4, values)).toBe(true)
    }
  })

  for (const reading of CF_UPPER_WINDOW_READINGS) {
    it(`registered ${reading} with evidence: the statistical window cell is no longer justified; the integer class and the literal bounds still are`, () => {
      injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: reading })
      withExactCfEvidence(caps, () => {
        expect(upperWholeShapeAt(statistical, 'A4', 3, values)).toBe(false)
        expect(upperWholeShapeAt(statistical, 'A5', 4, values)).toBe(true)
        expect(upperWholeShapeAt(literalNum, 'A4', 3, values)).toBe(true)
      })
      // Without the evidence the registration does not count.
      withExactCfEvidence(VERIFIED_CAPABILITIES, () => expect(upperWholeShapeAt(statistical, 'A4', 3, values)).toBe(true))
    })
  }
})
