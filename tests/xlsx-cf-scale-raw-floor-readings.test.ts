// @vitest-environment node
/**
 * The raw-floor follow-up (brief cf-rawfloor), Phase 1 task 2: the two guard-boundary readings (`raw-floor`, W1; `guarded`,
 * the pre-#102 behaviour), implemented and selectable through the test seam `injectCfGuardBoundaryReadings` with the id's
 * evidence, and registered nowhere. Default production is exactly main: every 2-colour or 3-colour lower-segment interior cell
 * whose step product lies within 1e-11 strictly below a whole number is withheld (`cf-unmeasured-guard-boundary`). A
 * registered reading lifts that withhold, and its justification (scripts/cf-refs/upper-whole-justify.ts), for segments with
 * no formula-resolved bound, under the default sRGB / truncate variants; the literal num-segment withhold and every
 * upper-segment rule are untouched. This checks production, not Excel. Since the capture of stage2-scale-raw-floor
 * (ce3f701d) matched raw-floor, it is registered and published (docs/validation/xlsx-cf.md §14.z.18): the default is
 * raw-floor with the id's evidence, and the pre-capture default (withheld) is pinned without that evidence.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_GUARD_BOUNDARY_CAP,
  CF_GUARD_BOUNDARY_PRODUCTION_READINGS,
  CF_GUARD_BOUNDARY_READINGS,
  CF_UPPER_WINDOW_CAP,
  cfGuardBoundaryReading,
  injectCfGuardBoundaryReadings,
  injectCfUpperWindowReadings,
  interpolateScaleColor,
  scaleColorAt,
  upperWindowReadingColor,
  type CfGuardBoundaryReading,
} from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { guardBoundaryShapeAt, type UpperWholeMarker } from '../scripts/cf-refs/upper-whole-justify'

afterEach(() => { injectCfGuardBoundaryReadings(null); injectCfUpperWindowReadings(null) })

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const P = 0xff14f028, Q = 0xffe61efa // P + 210·(1, −1, 1)
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const numF = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value, resolved: 'formula' }, argb })
const min = (argb: number): CfScalePoint => ({ bound: { type: 'min' }, argb })
const max = (argb: number): CfScalePoint => ({ bound: { type: 'max' }, argb })
const pctl = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'percentile', value }, argb })
const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
const colour = (o: ReturnType<typeof scaleColorAt>): string => (o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind)
const GUARD = 'cf-unmeasured-guard-boundary'
/** The published capabilities without the raw-floor registration's: the shape is withheld, as before ce3f701d. */
const UNPUBLISHED = [...VERIFIED_CAPABILITIES].filter(id => id !== CF_GUARD_BOUNDARY_CAP)

/**
 * The shapes, with Excel's fill in ed0fccb6 (raw floor) and production's guarded truncation: E7 (2-colour num, Q -> P), H8,
 * H18, I5 (2-colour num 0 / 210, P -> Q), E12 (2-colour min / max), A3, A4 (min / percentile 50 / max lower segment), D2
 * (num / percentile / num lower segment).
 */
const SHAPES = {
  E7: { v: 0.6, points: [num(0.4, Q), num(1, P)], t: [0.4, 1], excel: 'FFA163B5', guarded: 'FFA064B4' },
  H8: { v: 57, points: [num(0, P), num(210, Q)], t: [0, 210], excel: 'FF4CB860', guarded: 'FF4DB761' },
  H18: { v: 108, points: [num(0, P), num(210, Q)], t: [0, 210], excel: 'FF7F8593', guarded: 'FF808494' },
  I5: { v: 127, points: [num(0, P), num(210, Q)], t: [0, 210], excel: 'FF9272A6', guarded: 'FF9371A7' },
  E12: { v: 1.21, points: [min(Q), max(P)], t: [1.1, 2.2], excel: 'FFD232E6', guarded: 'FFD133E5' },
  A3: { v: 0.13, points: [min(P), pctl(50, Q), max(P)], t: [0.1, 0.4, 1], excel: 'FF28DC3C', guarded: 'FF29DB3D' },
  A4: { v: 0.21, points: [min(P), pctl(50, Q), max(P)], t: [0.1, 0.4, 1], excel: 'FF60A474', guarded: 'FF61A375' },
  D2: { v: 0.11, points: [num(0.1, P), pctl(50, Q), num(0.8, P)], t: [0.1, 0.2, 0.8], excel: 'FF28DC3C', guarded: 'FF29DB3D' },
  // The inverted 2-colour shape (num 210 / num 0, colours positional) and the midpoint-above-max re-sort (min / num 420 / num 210).
  inverted: { v: 57, points: [num(210, P), num(0, Q)], t: [210, 0], excel: 'FF4CB860', guarded: 'FF4DB761' },
  resort: { v: 57, points: [min(P), num(420, Q), num(210, P)], t: [0, 420, 210], excel: 'FF4CB860', guarded: 'FF4DB761' },
} as const
type Shape = keyof typeof SHAPES
const run = (k: Shape, opts: Parameters<typeof scaleColorAt>[3] = { hasEvidence: all }): string => colour(scaleColorAt(SHAPES[k].v, SHAPES[k].points, SHAPES[k].t, opts))
const register = (reading: CfGuardBoundaryReading): void => injectCfGuardBoundaryReadings({ [CF_GUARD_BOUNDARY_CAP]: reading })

describe('the readings: implemented, selectable; raw-floor registered (ce3f701d)', () => {
  it('raw-floor and guarded, the only readings; production registers raw-floor for the question and nothing else', () => {
    expect([...CF_GUARD_BOUNDARY_READINGS]).toEqual(['raw-floor', 'guarded'])
    expect(CF_GUARD_BOUNDARY_CAP).toBe('scale:guard-boundary-raw-floor')
    expect(CF_GUARD_BOUNDARY_PRODUCTION_READINGS).toEqual({ [CF_GUARD_BOUNDARY_CAP]: 'raw-floor' })
    expect(Object.isFrozen(CF_GUARD_BOUNDARY_PRODUCTION_READINGS)).toBe(true)
    expect(cfGuardBoundaryReading()).toBe('raw-floor')
    expect(cfGuardBoundaryReading('scale:guard-boundary-literal-lower')).toBeUndefined()
    register('guarded')
    expect(cfGuardBoundaryReading()).toBe('guarded')
    // The upper-window id is a different registration: selecting one never selects the other.
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'raw-floor' })
    expect(cfGuardBoundaryReading()).toBe('guarded')
    injectCfGuardBoundaryReadings({ [CF_GUARD_BOUNDARY_CAP]: 'nonsense' as CfGuardBoundaryReading })
    expect(cfGuardBoundaryReading()).toBeUndefined()
    // A proposal for another id leaves this id's production reading.
    injectCfGuardBoundaryReadings({ 'scale:other': 'guarded' })
    expect(cfGuardBoundaryReading()).toBe('raw-floor')
  })

  it('the reading colours, per ARGB channel from the start colour: guarded is production\'s interpolation, raw-floor one step smaller where a product is just below an integer; non-finite fails closed', () => {
    // The segment production evaluates (thresholds sorted, colours positional; the 3-colour lower segment after the re-sort).
    const seg: Record<Shape, [number, number, number, number]> = {
      E7: [0.4, 1, Q, P], H8: [0, 210, P, Q], H18: [0, 210, P, Q], I5: [0, 210, P, Q], E12: [1.1, 2.2, Q, P],
      A3: [0.1, 0.4, P, Q], A4: [0.1, 0.4, P, Q], D2: [0.1, 0.2, P, Q], inverted: [0, 210, P, Q], resort: [0, 210, P, Q],
    }
    for (const k of Object.keys(SHAPES) as Shape[]) {
      const [lo, hi, c0, c1] = seg[k]
      const f = (SHAPES[k].v - lo) / (hi - lo)
      expect(hex(upperWindowReadingColor('raw-floor', c0, c1, f)!), k).toBe(SHAPES[k].excel)
      expect(hex(upperWindowReadingColor('guarded', c0, c1, f)!), k).toBe(SHAPES[k].guarded)
      expect(hex(interpolateScaleColor(c0, c1, f)), k).toBe(SHAPES[k].guarded)
    }
    for (const r of CF_GUARD_BOUNDARY_READINGS) for (const f of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) expect(upperWindowReadingColor(r, P, Q, f), `${r} ${f}`).toBeUndefined()
  })
})

describe('scaleColorAt: default unchanged; a registered reading paints only the guard-boundary shape', () => {
  it('without the id\'s evidence (the default before ce3f701d): every shape is withheld; neighbours paint as before', () => {
    for (const k of Object.keys(SHAPES) as Shape[]) expect(run(k, { hasEvidence: without(CF_GUARD_BOUNDARY_CAP) }), k).toBe(GUARD)
    expect(colour(scaleColorAt(115, [num(0, P), num(210, Q)], [0, 210], { hasEvidence: all }))).toBe('FF877D9B') // H19, above
    expect(colour(scaleColorAt(70, [num(0, P), num(210, Q)], [0, 210], { hasEvidence: all }))).toBe('FF5AAA6E') // exact
  })

  it('as committed (raw-floor registered, with evidence): every shape paints Excel\'s fill of ed0fccb6', () => {
    for (const k of Object.keys(SHAPES) as Shape[]) expect(run(k), k).toBe(SHAPES[k].excel)
  })

  for (const reading of CF_GUARD_BOUNDARY_READINGS) {
    it(`registered ${reading} (with evidence): every shape paints the reading; neighbours, other variants, formula bounds, the literal and upper-segment shapes unchanged`, () => {
      register(reading)
      for (const k of Object.keys(SHAPES) as Shape[]) expect(run(k), k).toBe(reading === 'raw-floor' ? SHAPES[k].excel : SHAPES[k].guarded)
      // Outside the shape nothing changes (raw floor and guarded agree there).
      expect(colour(scaleColorAt(115, [num(0, P), num(210, Q)], [0, 210], { hasEvidence: all }))).toBe('FF877D9B')
      expect(colour(scaleColorAt(70, [num(0, P), num(210, Q)], [0, 210], { hasEvidence: all }))).toBe('FF5AAA6E')
      // Without the id's evidence the reading is not honoured.
      expect(run('E7', { hasEvidence: without(CF_GUARD_BOUNDARY_CAP) })).toBe(GUARD)
      // A non-default colour space keeps the withhold; a non-truncating rounding has no guard, so no shape.
      expect(run('E7', { hasEvidence: all, colorSpace: 'linear-light' })).toBe(GUARD)
      expect(run('H8', { hasEvidence: all, rounding: 'half-up' })).toBe(hex(interpolateScaleColor(P, Q, 57 / 210, 'sRGB', 'half-up')))
      // A formula-resolved bound was never measured in this shape: withheld (fail closed).
      expect(colour(scaleColorAt(57, [numF(0, P), num(210, Q)], [0, 210], { hasEvidence: all }))).toBe(GUARD)
      expect(colour(scaleColorAt(0.13, [min(P), { bound: { type: 'percentile', value: 50, resolved: 'formula' }, argb: Q }, max(P)], [0.1, 0.4, 1], { hasEvidence: all }))).toBe(GUARD)
      // The literal num-segment withhold keeps precedence on a 3-colour lower segment.
      expect(colour(scaleColorAt(0.13, [num(0.1, P), num(0.4, Q), num(1, P)], [0.1, 0.4, 1], { hasEvidence: all }))).toBe('cf-unmeasured-num-segment-fraction')
      // Upper-segment cells keep their own rules: the window class (F12's double) and the integer class stay withheld.
      expect(colour(scaleColorAt(0.6, [min(P), pctl(50, Q), max(P)], [0.1, 0.4, 1], { hasEvidence: all }))).toBe('cf-unmeasured-upper-segment-whole-product')
      expect(colour(scaleColorAt(267, [min(P), pctl(50, Q), max(P)], [0, 210, 420], { hasEvidence: all }))).toBe('cf-unmeasured-upper-segment-whole-product')
    })
  }

  it('the registration is per id: only scale:guard-boundary-raw-floor selects it', () => {
    // With this id unregistered (an invalid proposal), a reading under any other id never paints the shape.
    injectCfGuardBoundaryReadings({ [CF_GUARD_BOUNDARY_CAP]: 'nonsense' as CfGuardBoundaryReading })
    expect(run('E7')).toBe(GUARD)
    injectCfUpperWindowReadings({ [CF_UPPER_WINDOW_CAP]: 'guarded' })
    expect(run('E7')).toBe(GUARD)
    // A proposal for another guard-boundary id leaves the production reading (raw-floor) in place.
    injectCfUpperWindowReadings(null)
    injectCfGuardBoundaryReadings({ 'scale:upper-window-statistical': 'guarded' })
    expect(run('E7')).toBe(SHAPES.E7.excel)
  })
})

describe('the justification (upper-whole-justify.ts) narrows with the registration, and only then', () => {
  const values = new Map([['A2', 0], ['A3', 57], ['A4', 115], ['A5', 210]])
  const twoColour = (cfvo: ReadonlyArray<{ type: string; val?: string; f?: string }>): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A2:A5', scale: { cfvo, colors: ['FF14F028', 'FFE61EFA'] } })
  const literal = twoColour([{ type: 'num', val: '0' }, { type: 'num', val: '210' }])
  const minmax = twoColour([{ type: 'min' }, { type: 'max' }])
  const formula = twoColour([{ type: 'num', val: '0', f: '0' }, { type: 'num', val: '210' }])
  const caps = [...VERIFIED_CAPABILITIES, CF_GUARD_BOUNDARY_CAP]
  const formulaBound = (): number => 0

  it('without the evidence (the state before ce3f701d): H8 (57) is justified under literal, statistical and formula bounds; H19 (115) never is', () => {
    withExactCfEvidence(UNPUBLISHED, () => {
      for (const m of [literal, minmax]) {
        expect(guardBoundaryShapeAt(m, 'A3', 57, values)).toBe(true)
        expect(guardBoundaryShapeAt(m, 'A4', 115, values)).toBe(false)
      }
      expect(guardBoundaryShapeAt(formula, 'A3', 57, values, formulaBound)).toBe(true)
    })
  })

  it('as committed (raw-floor registered and published): the shape is justified only under a formula bound', () => {
    expect(VERIFIED_CAPABILITIES.has(CF_GUARD_BOUNDARY_CAP)).toBe(true)
    expect(guardBoundaryShapeAt(literal, 'A3', 57, values)).toBe(false)
    expect(guardBoundaryShapeAt(minmax, 'A3', 57, values)).toBe(false)
    expect(guardBoundaryShapeAt(formula, 'A3', 57, values, formulaBound)).toBe(true)
  })

  for (const reading of CF_GUARD_BOUNDARY_READINGS) {
    it(`registered ${reading} with evidence: the shape is no longer justified, except under a formula bound; without the evidence it still is`, () => {
      register(reading)
      withExactCfEvidence(caps, () => {
        expect(guardBoundaryShapeAt(literal, 'A3', 57, values)).toBe(false)
        expect(guardBoundaryShapeAt(minmax, 'A3', 57, values)).toBe(false)
        expect(guardBoundaryShapeAt(formula, 'A3', 57, values, formulaBound)).toBe(true)
      })
      withExactCfEvidence(UNPUBLISHED, () => expect(guardBoundaryShapeAt(literal, 'A3', 57, values)).toBe(true))
    })
  }
})
