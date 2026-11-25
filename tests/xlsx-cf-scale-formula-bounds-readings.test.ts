// @vitest-environment node
/**
 * The formula-bounds follow-up (brief cf-formula-bounds), Phase 1 task 2: the readings of a colour-scale segment bounded by a
 * formula-resolved cfvo in a shape whose rule a capture settled (the guard-boundary shape, the upper segment's integer and
 * window classes, and under its own id the upper near-top sub-shape), selectable through `injectCfFormulaBoundsReadings` with
 * the id's evidence and registered nowhere:
 * - `as-literal`: the shape paints by its registered reading with the formula's double substituted;
 * - `fifteen-digit`: the service first rounds the formula result to 15 significant digits (`cfFormulaBoundFifteenDigit`).
 * The core property, over a probe grid: unregistered, production is unchanged (those shapes withheld when a bound is
 * formula-resolved); registered with its evidence, a formula-bounded segment paints exactly as the same segment with literal
 * bounds; the near-top sub-shape only with its own id at the same reading; never between non-opaque endpoints in the
 * guard-boundary shape, nor under another colour space or rounding. The justifications narrow only with the registration plus
 * its evidence. This checks production, not Excel.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_FORMULA_BOUNDS_CAP,
  CF_FORMULA_BOUNDS_NEAR_TOP_CAP,
  CF_FORMULA_BOUNDS_PRODUCTION_READINGS,
  CF_FORMULA_BOUNDS_READINGS,
  cfFormulaBoundFifteenDigit,
  cfFormulaBoundsReading,
  fifteenDigitFormulaBounds,
  guardBoundaryProduct,
  injectCfFormulaBoundsReadings,
  scaleColorAt,
  upperSegmentNearTop,
  upperSegmentProductClass,
  type CfFormulaBoundsReading,
} from '../src/xlsx/cf/colorScale'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import type { CfScalePoint } from '../src/xlsx/cf/model'
import { formulaBoundsReadingRegistered, guardBoundaryShapeAt, independentFifteenDigit, upperNearTopShapeAt, upperWholeShapeAt, type UpperWholeMarker } from '../scripts/cf-refs/upper-whole-justify'

afterEach(() => { injectCfFormulaBoundsReadings(null) })

const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const P = 0xff14f028, Q = 0xffe61efa // P + 210·(1, −1, 1)
const DD0 = 0xfff09b32, DD1 = 0xff1e3214 // spans 210 / 105 / 30, all descending from DD0
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const numF = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value, resolved: 'formula' }, argb })
const mn = (argb: number): CfScalePoint => ({ bound: { type: 'min' }, argb })
const mx = (argb: number): CfScalePoint => ({ bound: { type: 'max' }, argb })
const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
const colour = (o: ReturnType<typeof scaleColorAt>): string => (o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind)
const register = (r: CfFormulaBoundsReading, nearTop?: CfFormulaBoundsReading): void =>
  injectCfFormulaBoundsReadings({ [CF_FORMULA_BOUNDS_CAP]: r, ...(nearTop ? { [CF_FORMULA_BOUNDS_NEAR_TOP_CAP]: nearTop } : {}) })
const ulp = (x: number, n: number): number => { const b = new Float64Array([x]); const u = new BigInt64Array(b.buffer); u[0] = u[0]! + BigInt(n * (x >= 0 ? 1 : -1)); return b[0]! }

type Probe = { label: string; literal: CfScalePoint[]; formula: CfScalePoint[]; thresholds: number[]; value: number; shape: 'guard' | 'upper' | 'near-top' | 'clear' }

/** Segments with one or two formula-resolved bounds, values at k/|d| of the span and ±1, ±3 ulps (near-whole products). */
function probes(): Probe[] {
  const out: Probe[] = []
  const spans: Array<[number, number]> = [[0, 0.3], [7, 217], [-210, 210], [0.1, 0.30000000000000004], [2.1, 3.5]]
  const palettes: Array<[number, number]> = [[P, Q], [Q, P], [DD0, DD1]]
  for (const [lo, hi] of spans) for (const [a, b] of palettes) for (const k of [1, 7, 30, 70, 105, 139, 209]) for (const u of [0, -1, 1, -3, 3]) {
    const v = ulp(lo + (k / 210) * (hi - lo), u)
    if (!(v > lo && v < hi)) continue
    const shapeOf = (c0: number, c1: number, l: number, h: number, upper: boolean): Probe['shape'] => {
      const f = (v - l) / (h - l)
      if (upper && upperSegmentNearTop(c0, c1, v, l, h, f)) return 'near-top'
      if (upper) return upperSegmentProductClass(c0, c1, v, l, h, f) ? 'upper' : 'clear'
      return guardBoundaryProduct(c0, c1, v, l, h, f) ? 'guard' : 'clear'
    }
    // 2-colour: formula hi; formula lo and hi.
    out.push({ label: `2c ${lo}/${hi}F ${hex(a)} k${k} u${u}`, literal: [num(lo, a), num(hi, b)], formula: [num(lo, a), numF(hi, b)], thresholds: [lo, hi], value: v, shape: shapeOf(a, b, lo, hi, false) })
    out.push({ label: `2c ${lo}F/${hi}F ${hex(a)} k${k} u${u}`, literal: [num(lo, a), num(hi, b)], formula: [numF(lo, a), numF(hi, b)], thresholds: [lo, hi], value: v, shape: shapeOf(a, b, lo, hi, false) })
    // 3-colour lower segment: formula midpoint (min below).
    const below = lo - (hi - lo)
    out.push({ label: `3c lower min/${hi}F/max ${hex(a)} k${k} u${u}`, literal: [mn(b), num(hi, a), mx(b)], formula: [mn(b), numF(hi, a), mx(b)], thresholds: [lo, hi, hi + (hi - lo)], value: v, shape: shapeOf(b, a, lo, hi, false) })
    // 3-colour upper segment: formula midpoint and maximum.
    out.push({ label: `3c upper ${below}/${lo}F/${hi}F ${hex(a)} k${k} u${u}`, literal: [num(below, b), num(lo, a), num(hi, b)], formula: [num(below, b), numF(lo, a), numF(hi, b)], thresholds: [below, lo, hi], value: v, shape: shapeOf(a, b, lo, hi, true) })
    out.push({ label: `3c upper min/${lo}F/max ${hex(a)} k${k} u${u}`, literal: [mn(b), num(lo, a), mx(b)], formula: [mn(b), numF(lo, a), mx(b)], thresholds: [below, lo, hi], value: v, shape: shapeOf(a, b, lo, hi, true) })
  }
  // Near-top: values 1 and 7 ulps below the maximum of an upper segment.
  for (const [lo, hi] of spans) for (const u of [-1, -7]) {
    const v = ulp(hi, u)
    const below = lo - (hi - lo)
    out.push({ label: `3c near-top ${lo}F/${hi}F u${u}`, literal: [num(below, P), num(lo, Q), num(hi, P)], formula: [num(below, P), numF(lo, Q), numF(hi, P)], thresholds: [below, lo, hi], value: v, shape: upperSegmentNearTop(Q, P, v, lo, hi, (v - lo) / (hi - lo)) ? 'near-top' : 'clear' })
  }
  return out
}

describe('the formula-bounds readings (production, selectable, unregistered)', () => {
  it('two readings; no production registration; an unknown or absent proposal is no reading', () => {
    expect(CF_FORMULA_BOUNDS_READINGS).toEqual(['as-literal', 'fifteen-digit'])
    expect(CF_FORMULA_BOUNDS_PRODUCTION_READINGS).toEqual({})
    expect(cfFormulaBoundsReading()).toBeUndefined()
    injectCfFormulaBoundsReadings({ [CF_FORMULA_BOUNDS_CAP]: 'rounded' as CfFormulaBoundsReading })
    expect(cfFormulaBoundsReading()).toBeUndefined()
    register('fifteen-digit')
    expect(cfFormulaBoundsReading()).toBe('fifteen-digit')
    expect(fifteenDigitFormulaBounds(all)).toBe(true)
    expect(fifteenDigitFormulaBounds(without(CF_FORMULA_BOUNDS_CAP))).toBe(false)
    register('as-literal')
    expect(fifteenDigitFormulaBounds(all)).toBe(false)
  })

  it('the 15-digit rounding: the double nearest the 15-significant-digit decimal; non-finite stays so', () => {
    expect(cfFormulaBoundFifteenDigit(0.1 + 0.2)).toBe(0.3)
    expect(cfFormulaBoundFifteenDigit(100 / 3)).toBe(33.3333333333333)
    expect(cfFormulaBoundFifteenDigit(210)).toBe(210)
    expect(cfFormulaBoundFifteenDigit(0.3)).toBe(0.3)
    expect(cfFormulaBoundFifteenDigit(Number.POSITIVE_INFINITY)).toBe(Number.POSITIVE_INFINITY)
    expect(cfFormulaBoundFifteenDigit(Number.MAX_VALUE)).toBe(Number.POSITIVE_INFINITY)
    for (const x of [0.1 + 0.2, 100 / 3, 1 / 3 * 630, 426.99999999999994, -0.30000000000000004, 1e-300 / 3]) expect(independentFifteenDigit(x)).toBe(cfFormulaBoundFifteenDigit(x))
  })

  const grid = probes()
  it(`the probe grid holds every shape (${grid.length} probes)`, () => {
    const shapes = new Set(grid.map(p => p.shape))
    expect([...shapes].sort()).toEqual(['clear', 'guard', 'near-top', 'upper'])
    expect(grid.filter(p => p.shape === 'guard').length).toBeGreaterThan(20)
    expect(grid.filter(p => p.shape === 'upper').length).toBeGreaterThan(20)
  })

  it('unregistered (or without the evidence): a formula-bounded cell of a settled shape is withheld, every other one paints as its literal twin', () => {
    for (const [ev, proposal] of [[all, undefined], [without(CF_FORMULA_BOUNDS_CAP), 'as-literal'], [without(CF_FORMULA_BOUNDS_CAP), 'fifteen-digit']] as const) {
      if (proposal) register(proposal)
      for (const p of grid) {
        const lit = colour(scaleColorAt(p.value, p.literal, p.thresholds, { hasEvidence: ev }))
        const got = colour(scaleColorAt(p.value, p.formula, p.thresholds, { hasEvidence: ev }))
        if (p.shape === 'clear') expect(got, p.label).toBe(lit)
        else expect(got, p.label).toMatch(/^cf-unmeasured-/)
      }
    }
  })

  for (const r of CF_FORMULA_BOUNDS_READINGS) {
    it(`${r} registered with its evidence: a formula-bounded segment paints as its literal twin (near-top withheld without its own id)`, () => {
      register(r)
      for (const p of grid) {
        const lit = colour(scaleColorAt(p.value, p.literal, p.thresholds, { hasEvidence: all }))
        const got = colour(scaleColorAt(p.value, p.formula, p.thresholds, { hasEvidence: all }))
        if (p.shape === 'near-top') expect(got, p.label).toBe('cf-unmeasured-upper-near-top')
        else { expect(lit, p.label).toMatch(/^FF/); expect(got, p.label).toBe(lit) }
      }
    })

    it(`${r}: the near-top sub-shape paints only with ${CF_FORMULA_BOUNDS_NEAR_TOP_CAP} at the same reading and its evidence`, () => {
      const nt = grid.filter(p => p.shape === 'near-top')
      expect(nt.length).toBeGreaterThan(3)
      const other = CF_FORMULA_BOUNDS_READINGS.find(x => x !== r)!
      for (const [near, ev, paints] of [[r, all, true], [other, all, false], [r, without(CF_FORMULA_BOUNDS_NEAR_TOP_CAP), false]] as const) {
        register(r, near)
        for (const p of nt) {
          const lit = colour(scaleColorAt(p.value, p.literal, p.thresholds, { hasEvidence: ev }))
          const got = colour(scaleColorAt(p.value, p.formula, p.thresholds, { hasEvidence: ev }))
          expect(got, p.label).toBe(paints ? lit : 'cf-unmeasured-upper-near-top')
        }
      }
    })
  }

  it('every other gate stays: non-opaque endpoints in the guard-boundary shape, another colour space or rounding', () => {
    register('as-literal')
    const guard = grid.filter(p => p.shape === 'guard' && p.formula.length === 2)
    expect(guard.length).toBeGreaterThan(5)
    const translucent = (pts: CfScalePoint[]): CfScalePoint[] => pts.map((x, i) => ({ ...x, argb: i === 0 ? (x.argb & 0x00ffffff) | 0x80000000 : x.argb }))
    for (const p of guard) {
      expect(colour(scaleColorAt(p.value, translucent(p.formula), p.thresholds, { hasEvidence: all })), p.label).toBe('cf-unmeasured-guard-boundary')
      expect(colour(scaleColorAt(p.value, p.formula, p.thresholds, { hasEvidence: all, colorSpace: 'linear-light' })), p.label).toBe('cf-unmeasured-guard-boundary')
    }
    for (const p of grid.filter(x => x.shape === 'upper')) {
      expect(colour(scaleColorAt(p.value, p.formula, p.thresholds, { hasEvidence: all, rounding: 'half-up' })), p.label).toMatch(/^cf-unmeasured-/)
    }
  })
})

describe('the justifications narrow only with the registration plus its evidence', () => {
  const withIds = <T>(fn: () => T): T => withExactCfEvidence([...VERIFIED_CAPABILITIES, CF_FORMULA_BOUNDS_CAP, CF_FORMULA_BOUNDS_NEAR_TOP_CAP], fn)
  const marker3 = (cfvo: UpperWholeMarker['scale'] extends infer S ? S extends { cfvo: infer C } ? C : never : never, colors: string[]): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A1:A5', scale: { cfvo, colors } })
  // P-line upper segment 7 .. 217 (span 210) bounded by =$Z$1 (217) as the maximum: the value 7 + 105 = 112 has an integer product.
  const upper = marker3([{ type: 'num', val: '-203' }, { type: 'num', val: '7' }, { type: 'formula', f: '$Z$1' }], ['FFE61EFA', 'FFE61EFA', 'FF14F028'])
  const values = new Map([['A1', -203], ['A2', 112], ['A3', 217]])
  const resolve = (): number => 217
  it('upperWholeShapeAt: justified unregistered or without the evidence; not with as-literal or fifteen-digit and the evidence', () => {
    withIds(() => expect(upperWholeShapeAt(upper, 'A2', 112, values, resolve)).toBe(true))
    register('as-literal')
    expect(formulaBoundsReadingRegistered()).toBeUndefined()
    expect(upperWholeShapeAt(upper, 'A2', 112, values, resolve)).toBe(true)
    withIds(() => {
      expect(formulaBoundsReadingRegistered()).toBe('as-literal')
      expect(upperWholeShapeAt(upper, 'A2', 112, values, resolve)).toBe(false)
      register('fifteen-digit')
      expect(upperWholeShapeAt(upper, 'A2', 112, values, resolve)).toBe(false)
    })
  })
  it('upperNearTopShapeAt: a formula-bounded near-top cell stays justified until the formula near-top id has the same reading', () => {
    const v = 216.99999999999997
    withIds(() => {
      expect(upperNearTopShapeAt(upper, 'A2', v, values, resolve)).toBe(true)
      register('as-literal')
      expect(upperNearTopShapeAt(upper, 'A2', v, values, resolve)).toBe(true)
      register('as-literal', 'fifteen-digit')
      expect(upperNearTopShapeAt(upper, 'A2', v, values, resolve)).toBe(true)
      register('as-literal', 'as-literal')
      expect(upperNearTopShapeAt(upper, 'A2', v, values, resolve)).toBe(false)
    })
    // Without the near-top evidence it stays justified.
    expect(upperNearTopShapeAt(upper, 'A2', v, values, resolve)).toBe(true)
  })
  it('guardBoundaryShapeAt: a formula-bounded 2-colour cell stays justified unregistered and between non-opaque colours', () => {
    // num 0 / =$Z$1 (0.30000000000000004), P -> Q: v = 0.15 gives f·210 = 104.99999999999997, just below 105.
    const m2 = (colors: string[]): UpperWholeMarker => ({ ruleType: 'colorScale', sqref: 'A1:A5', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'formula', f: '$Z$1' }], colors } })
    const z = (): number => 0.30000000000000004
    const vs = new Map([['A1', 0.15]])
    withIds(() => {
      expect(guardBoundaryShapeAt(m2(['FF14F028', 'FFE61EFA']), 'A1', 0.15, vs, z)).toBe(true)
      register('as-literal')
      expect(guardBoundaryShapeAt(m2(['FF14F028', 'FFE61EFA']), 'A1', 0.15, vs, z)).toBe(false)
      expect(guardBoundaryShapeAt(m2(['8014F028', 'FFE61EFA']), 'A1', 0.15, vs, z)).toBe(true)
      // Under fifteen-digit the bound is 0.3: f·210 = 105 exactly, not the shape.
      register('fifteen-digit')
      expect(guardBoundaryShapeAt(m2(['8014F028', 'FFE61EFA']), 'A1', 0.15, vs, z)).toBe(false)
    })
  })
})
