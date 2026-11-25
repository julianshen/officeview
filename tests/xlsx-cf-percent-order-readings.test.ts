// @vitest-environment node
/**
 * The percent-order follow-up (brief cf-percent-order), Phase 1 task 2: the order of a percent / percentile threshold's
 * arithmetic, selectable per rule kind and bound type through `injectCfPercentOrderReadings` (with the id's evidence) and the
 * pre-committed narrowing through `injectCfPercentOrderNarrowings`, registered nowhere; and `scale:formula-bounds-v2`, the
 * formula-bounds lift under each order. Unregistered, production is unchanged (`lerp-first` everywhere); data bars never
 * take an order. This checks production, not Excel.
 */
import { afterEach, describe, expect, it } from 'vitest'
import {
  CF_FORMULA_BOUNDS_V2_CAP,
  CF_FORMULA_BOUNDS_V2_PRODUCTION_READINGS,
  CF_ICON_PERCENT_ORDER_CAP,
  CF_ICON_PERCENTILE_ORDER_CAP,
  CF_PERCENT_ORDER_CAPS,
  CF_PERCENT_ORDER_NARROWINGS,
  CF_PERCENT_ORDER_PRODUCTION_READINGS,
  CF_SCALE_PERCENT_ORDER_CAP,
  CF_SCALE_PERCENTILE_ORDER_CAP,
  cfPercentOrder,
  cfScalePercentOrders,
  fifteenDigitFormulaBounds,
  injectCfFormulaBoundsV2Readings,
  injectCfPercentOrderNarrowings,
  injectCfPercentOrderReadings,
  orderedThreshold,
  percentileOrderStatistic,
  percentThreshold,
  resolveScaleThresholds,
  scaleColorAt,
  type CfFormulaBoundsV2Reading,
} from '../src/xlsx/cf/colorScale'
import { resolveBarThresholds } from '../src/xlsx/cf/dataBar'
import { resolveIconThresholds } from '../src/xlsx/cf/iconSet'
import type { CfDefinition, CfIconThreshold, CfScalePoint } from '../src/xlsx/cf/model'
import type { CfRangeStatistics } from '../src/xlsx/cf/statistics'
import { VERIFIED_CAPABILITIES, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { formulaBoundsReadingRegistered, guardBoundaryShapeAt, justifyPercent, justifyPercentileRanked, justifyPercentOrder, percentOrderShapeAt } from '../scripts/cf-refs/upper-whole-justify'
import { surveyPercent, surveyPercentile } from '../scripts/cf-refs/percent-order-survey'

afterEach(() => { injectCfPercentOrderReadings(null); injectCfPercentOrderNarrowings(null); injectCfFormulaBoundsV2Readings(null) })

const all = (cap: string): boolean => cap.length > 0
const without = (id: string) => (cap: string): boolean => cap !== id
function statsOf(values: readonly number[]): CfRangeStatistics {
  const sorted = Float64Array.from([...values].sort((a, b) => a - b))
  return {
    scopeKey: 't', numericCount: values.length, sum: 0, mean: undefined, min: sorted[0], max: sorted[sorted.length - 1], stddevPop: undefined,
    sortedAsc: sorted, frequencies: new Map(), declaredCells: values.length, spanCells: values.length, storedCells: values.length,
  }
}
const hex = (n: number): string => (n >>> 0).toString(16).toUpperCase().padStart(8, '0')
const P = 0xff14f028, Q = 0xffe61efa
const pct = (value: number, argb = P): CfScalePoint => ({ bound: { type: 'percent', value }, argb })
const pctF = (value: number, argb = P): CfScalePoint => ({ bound: { type: 'percent', value, resolved: 'formula' }, argb })
const pctl = (value: number, argb = P): CfScalePoint => ({ bound: { type: 'percentile', value }, argb })
const pctlF = (value: number, argb = P): CfScalePoint => ({ bound: { type: 'percentile', value, resolved: 'formula' }, argb })
const mx = (argb = Q): CfScalePoint => ({ bound: { type: 'max' }, argb })
const numF = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value, resolved: 'formula' }, argb })
const num = (value: number, argb: number): CfScalePoint => ({ bound: { type: 'num', value }, argb })
const thresholdsOf = (o: ReturnType<typeof resolveScaleThresholds>): readonly number[] | string => (o.kind === 'ready' ? o.value : o.kind === 'unresolved' ? o.reason : o.kind)
const scaleOrders = () => cfScalePercentOrders(all)
/** The justifications read the published evidence: run them with the given ids held too. */
const holding = <T>(ids: readonly string[], fn: () => T): T => withExactCfEvidence([...VERIFIED_CAPABILITIES, ...ids], fn)
const HELD = [CF_FORMULA_BOUNDS_V2_CAP, ...CF_PERCENT_ORDER_CAPS]

/** E2:E10 of stage2-scale-formula-bounds: p = 33.333333333333336 over 0..300; percentile 20 over stage5-cfvo-formulas G. */
const E = [0, 120, 140, 160, 180, 200, 220, 250, 300]
const G = [5, 15, 30, 45, 60, 75, 90, 100]

describe('the two orders and the narrowing', () => {
  it('percentThreshold and percentileOrderStatistic under each order, and the survey\'s independent restatement agrees', () => {
    expect(percentThreshold(0, 300, 33.333333333333336, 'lerp-first')).toBe(100.00000000000001)
    expect(percentThreshold(0, 300, 33.333333333333336, 'multiply-first')).toBe(100)
    expect(percentileOrderStatistic(G, 20, 'inc')).toBe(21.000000000000007)
    expect(percentileOrderStatistic(G, 20, 'inc', 'multiply-first')).toBe(21)
    // PERCENTILE.EXC under multiply-first is unmeasured: none.
    expect(percentileOrderStatistic(G, 20, 'exc', 'multiply-first')).toBeUndefined()
    expect(percentileOrderStatistic(G, 20, 'exc')).toBeDefined()
    for (const [lo, hi] of [[0, 300], [10, 100], [-37, 0.7], [-1e6, 3]] as const) for (const p of [0, 1, 7, 33, 33.333333333333336, 50, 67, 99.9, 100]) {
      for (const r of ['lerp-first', 'multiply-first'] as const) {
        expect(percentThreshold(lo, hi, p, r)).toBe(surveyPercent(lo, hi, p, r))
        expect(percentThreshold(lo, hi, p, r)).toBe(justifyPercent(lo, hi, p, r))
        const s = [lo, (lo + hi) / 3, (lo + hi) / 2, hi].sort((a, b) => a - b)
        expect(percentileOrderStatistic(s, p, 'inc', r)).toBe(surveyPercentile(s, p, r))
        expect(percentileOrderStatistic(s, p, 'inc', r)).toBe(justifyPercentileRanked(s, p, r))
      }
    }
  })

  it('orderedThreshold: the reading\'s value; the narrowing keeps a common double and refuses a differing one', () => {
    expect(orderedThreshold('lerp-first', r => percentThreshold(0, 300, 33.333333333333336, r))).toBe(100.00000000000001)
    expect(orderedThreshold('multiply-first', r => percentThreshold(0, 300, 33.333333333333336, r))).toBe(100)
    expect(orderedThreshold('withhold-differing', r => percentThreshold(0, 300, 33.333333333333336, r))).toBe('differs')
    expect(orderedThreshold('withhold-differing', r => percentThreshold(0, 300, 50, r))).toBe(150)
    expect(orderedThreshold('withhold-differing', r => (r === 'lerp-first' ? 1 : undefined))).toBe('differs')
    expect(orderedThreshold('withhold-differing', () => undefined)).toBeUndefined()
  })

  it('unregistered: every id is lerp-first, the tables are empty, nothing is narrowed', () => {
    expect(CF_PERCENT_ORDER_PRODUCTION_READINGS).toEqual({})
    expect(CF_PERCENT_ORDER_NARROWINGS).toEqual([])
    expect(CF_FORMULA_BOUNDS_V2_PRODUCTION_READINGS).toEqual({})
    for (const id of CF_PERCENT_ORDER_CAPS) expect(cfPercentOrder(id, all)).toBe('lerp-first')
    for (const id of CF_PERCENT_ORDER_CAPS) expect(justifyPercentOrder(id)).toBe('lerp-first')
  })

  it('a reading applies only with its evidence; the narrowing applies whatever the evidence', () => {
    injectCfPercentOrderReadings({ [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(cfPercentOrder(CF_SCALE_PERCENT_ORDER_CAP, all)).toBe('multiply-first')
    expect(cfPercentOrder(CF_SCALE_PERCENT_ORDER_CAP, without(CF_SCALE_PERCENT_ORDER_CAP))).toBe('lerp-first')
    // Each id is its own: percentile and the icon ids stay lerp-first.
    for (const id of [CF_SCALE_PERCENTILE_ORDER_CAP, CF_ICON_PERCENT_ORDER_CAP, CF_ICON_PERCENTILE_ORDER_CAP]) expect(cfPercentOrder(id, all)).toBe('lerp-first')
    injectCfPercentOrderNarrowings([CF_SCALE_PERCENT_ORDER_CAP])
    expect(cfPercentOrder(CF_SCALE_PERCENT_ORDER_CAP, () => false)).toBe('withhold-differing')
  })
})

describe('colour scales', () => {
  it('resolveScaleThresholds: lerp-first by default; each order and the narrowing when passed', () => {
    const s = statsOf(E)
    expect(thresholdsOf(resolveScaleThresholds([pct(33.333333333333336), mx()], s, { hasEvidence: all }))).toEqual([100.00000000000001, 300])
    expect(thresholdsOf(resolveScaleThresholds([pct(33.333333333333336), mx()], s, { hasEvidence: all, percentOrders: { percent: 'multiply-first', percentile: 'lerp-first' } }))).toEqual([100, 300])
    expect(thresholdsOf(resolveScaleThresholds([pct(33.333333333333336), mx()], s, { hasEvidence: all, percentOrders: { percent: 'withhold-differing', percentile: 'lerp-first' } }))).toBe('cf-unmeasured-percent-order')
    expect(thresholdsOf(resolveScaleThresholds([pct(50), mx()], s, { hasEvidence: all, percentOrders: { percent: 'withhold-differing', percentile: 'lerp-first' } }))).toEqual([150, 300])
    const g = statsOf(G)
    expect(thresholdsOf(resolveScaleThresholds([pctl(20), mx()], g, { hasEvidence: all }))).toEqual([21.000000000000007, 100])
    expect(thresholdsOf(resolveScaleThresholds([pctl(20), mx()], g, { hasEvidence: all, percentOrders: { percent: 'lerp-first', percentile: 'multiply-first' } }))).toEqual([21, 100])
    // The percentile order never moves a percent bound, nor the percent order a percentile one.
    expect(thresholdsOf(resolveScaleThresholds([pctl(20), mx()], g, { hasEvidence: all, percentOrders: { percent: 'multiply-first', percentile: 'lerp-first' } }))).toEqual([21.000000000000007, 100])
    expect(thresholdsOf(resolveScaleThresholds([pctl(20), mx()], g, { hasEvidence: all, percentOrders: { percent: 'lerp-first', percentile: 'withhold-differing' } }))).toBe('cf-unmeasured-percent-order')
  })

  it('multiply-first fails closed where the span or the product overflows', () => {
    const s = statsOf([-1e308, 0, 1e308])
    expect(thresholdsOf(resolveScaleThresholds([pct(10), mx()], s, { hasEvidence: all }))).not.toBe('cf-invalid-bound')
    expect(thresholdsOf(resolveScaleThresholds([pct(10), mx()], s, { hasEvidence: all, percentOrders: { percent: 'multiply-first', percentile: 'lerp-first' } }))).toBe('cf-invalid-bound')
    expect(thresholdsOf(resolveScaleThresholds([pct(10), mx()], statsOf([0, 1.7e308]), { hasEvidence: all, percentOrders: { percent: 'multiply-first', percentile: 'lerp-first' } }))).toBe('cf-invalid-bound')
  })

  it('E2 of stage2-scale-formula-bounds, as a literal percent: lerp-first paints FF28DC3C (production today), multiply-first FF29DB3D', () => {
    const s = statsOf(E)
    const at = (orders?: NonNullable<Parameters<typeof resolveScaleThresholds>[2]>['percentOrders']): string => {
      const t = resolveScaleThresholds([pct(33.333333333333336), mx()], s, { hasEvidence: all, percentOrders: orders })
      if (t.kind !== 'ready') return 'unresolved'
      const o = scaleColorAt(120, [pct(33.333333333333336), mx()], t.value)
      return o.kind === 'ready' ? hex(o.value) : o.kind
    }
    expect(at()).toBe('FF28DC3C')
    expect(at(scaleOrders())).toBe('FF28DC3C')
    injectCfPercentOrderReadings({ [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(at(scaleOrders())).toBe('FF29DB3D')
  })
})

describe('icon sets', () => {
  type IconDef = Extract<CfDefinition, { kind: 'iconSet' }>
  const th = (type: 'percent' | 'percentile', value: number): CfIconThreshold => ({ bound: { type, value }, gte: true })
  const def = (type: 'percent' | 'percentile'): IconDef => ({ kind: 'iconSet', dialect: 'base', formulas: [], set: '3Arrows', thresholds: [th(type, 0), th(type, 33), th(type, 67)], reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false })
  const icons = (d: IconDef, s: CfRangeStatistics, check: (cap: string) => boolean = all): readonly number[] | string => { const o = resolveIconThresholds(d, s, { hasEvidence: check }); return o.kind === 'ready' ? o.value : o.kind === 'unresolved' ? o.reason : o.kind }
  // stage4-icon-scope B2:B4 D2:D4: percent 67 over 10..100.
  const scope = statsOf([10, 20, 30, 80, 90, 100])
  it('percent: lerp-first by default; multiply-first with its registration and evidence; the narrowing withholds', () => {
    expect(icons(def('percent'), scope)).toEqual([10, 39.7, 70.30000000000001])
    injectCfPercentOrderReadings({ [CF_ICON_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(icons(def('percent'), scope)).toEqual([10, 39.7, 70.3])
    expect(icons(def('percent'), scope, without(CF_ICON_PERCENT_ORDER_CAP))).toEqual([10, 39.7, 70.30000000000001])
    injectCfPercentOrderReadings(null)
    injectCfPercentOrderNarrowings([CF_ICON_PERCENT_ORDER_CAP])
    expect(icons(def('percent'), scope)).toBe('cf-unmeasured-percent-order')
    expect(icons(def('percentile'), statsOf([10, 20, 30, 40, 50, 60, 70]))).not.toBe('cf-unmeasured-percent-order')
  })
  it('percentile: its own id; a scale registration never moves an icon', () => {
    const g = statsOf(G)
    const d: IconDef = { ...def('percentile'), thresholds: [th('percentile', 0), th('percentile', 20), th('percentile', 67)] }
    const base = icons(d, g)
    expect(base).toEqual([5, 21.000000000000007, percentileOrderStatistic(G, 67, 'inc')])
    injectCfPercentOrderReadings({ [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first', [CF_SCALE_PERCENTILE_ORDER_CAP]: 'multiply-first', [CF_ICON_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(icons(d, g)).toEqual(base)
    injectCfPercentOrderReadings({ [CF_ICON_PERCENTILE_ORDER_CAP]: 'multiply-first' })
    expect(icons(d, g)).toEqual([5, 21, percentileOrderStatistic(G, 67, 'inc', 'multiply-first')])
  })
})

describe('data bars take no order', () => {
  it('every registration and narrowing leaves bar thresholds bit-identical', () => {
    const bar = (lo: number, hi: number): Extract<CfDefinition, { kind: 'dataBar' }> => ({
      kind: 'dataBar', dialect: 'base', formulas: [], bounds: [{ type: 'percent', value: lo }, { type: 'percentile', value: hi }],
      fillArgb: 0xff0000ff, negativeFillArgb: 0xffff0000, axisArgb: 0xff000000, minLength: 0, maxLength: 100, gradient: false, border: false,
      negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true, axisPosition: 'automatic', direction: 'leftToRight', showValue: true, sheetRightToLeft: false,
    })
    const s = statsOf([0, 120, 140, 160, 180, 200, 220, 250, 300, 5, 15, 30])
    const run = (): string => JSON.stringify([bar(33.333333333333336, 80), bar(20, 90), bar(67, 99)].map(b => resolveBarThresholds(b, s, { hasEvidence: all })))
    const before = run()
    expect(before).toContain('100.00000000000001')
    injectCfPercentOrderReadings(Object.fromEntries(CF_PERCENT_ORDER_CAPS.map(id => [id, 'multiply-first'])))
    expect(run()).toBe(before)
    injectCfPercentOrderReadings(null)
    injectCfPercentOrderNarrowings([...CF_PERCENT_ORDER_CAPS])
    expect(run()).toBe(before)
  })
})

describe('scale:formula-bounds-v2', () => {
  // A guard-boundary cell of a 2-colour formula-bounded segment: num 0 / num =0.30000000000000004, P to Q, v = 0.01.
  const seg = [num(0, P), numF(0.30000000000000004, Q)]
  const at = (points: CfScalePoint[], t: number[], v: number): string => { const o = scaleColorAt(v, points, t, { hasEvidence: all }); return o.kind === 'ready' ? hex(o.value) : o.kind === 'unresolved' ? o.reason : o.kind }
  const register = (r: CfFormulaBoundsV2Reading): void => injectCfFormulaBoundsV2Readings({ [CF_FORMULA_BOUNDS_V2_CAP]: r })
  const literal = at([num(0, P), num(0.30000000000000004, Q)], [0, 0.30000000000000004], 0.01)

  it('unregistered, a formula-bounded guard-boundary cell stays withheld; its literal twin paints', () => {
    expect(literal).toMatch(/^FF/)
    expect(at(seg, [0, 0.30000000000000004], 0.01)).toBe('cf-unmeasured-guard-boundary')
  })
  it('a v2 reading lifts it only when its order is the percent order production applies', () => {
    register('as-literal+lerp-first')
    expect(at(seg, [0, 0.30000000000000004], 0.01)).toBe(literal)
    expect(holding(HELD, formulaBoundsReadingRegistered)).toBe('as-literal')
    expect(formulaBoundsReadingRegistered()).toBeUndefined()
    register('as-literal+multiply-first')
    expect(at(seg, [0, 0.30000000000000004], 0.01)).toBe('cf-unmeasured-guard-boundary')
    expect(holding(HELD, formulaBoundsReadingRegistered)).toBeUndefined()
    injectCfPercentOrderReadings({ [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(at(seg, [0, 0.30000000000000004], 0.01)).toBe(literal)
    expect(holding(HELD, formulaBoundsReadingRegistered)).toBe('as-literal')
    // Without the percent-order evidence the order is lerp-first, so a multiply-first v2 reading lifts nothing.
    expect(holding([CF_FORMULA_BOUNDS_V2_CAP], formulaBoundsReadingRegistered)).toBeUndefined()
    // The narrowing matches neither order.
    injectCfPercentOrderNarrowings([CF_SCALE_PERCENT_ORDER_CAP])
    expect(at(seg, [0, 0.30000000000000004], 0.01)).toBe('cf-unmeasured-guard-boundary')
  })
  it('fifteen-digit+… rounds (the service substitution) only at its order; v2 never lifts a formula percentile', () => {
    register('fifteen-digit+lerp-first')
    expect(fifteenDigitFormulaBounds(all)).toBe(true)
    expect(fifteenDigitFormulaBounds(without(CF_FORMULA_BOUNDS_V2_CAP))).toBe(false)
    register('fifteen-digit+multiply-first')
    expect(fifteenDigitFormulaBounds(all)).toBe(false)
    register('as-literal+lerp-first')
    // A 2-colour segment bounded by a formula percentile (p = 0 over 0..0.30000000000000004): withheld under v2.
    expect(at([pctlF(0, P), num(0.30000000000000004, Q)], [0, 0.30000000000000004], 0.01)).toBe('cf-unmeasured-guard-boundary')
    expect(at([pctF(0, P), num(0.30000000000000004, Q)], [0, 0.30000000000000004], 0.01)).toBe(literal)
  })
  it('the justifications follow the same rule', () => {
    const marker = { ruleType: 'colorScale', sqref: 'A1:A3', scale: { cfvo: [{ type: 'num', val: '0' }, { type: 'num', f: '$Z$1' }], colors: ['FF14F028', 'FFE61EFA'] } }
    const values = new Map([['A1', 0], ['A2', 0.01], ['A3', 0.3]])
    const z = (): number => 0.30000000000000004
    const shape = (): boolean => holding(HELD, () => guardBoundaryShapeAt(marker, 'A2', 0.01, values, z))
    expect(shape()).toBe(true)
    register('as-literal+lerp-first')
    expect(shape()).toBe(false)
    expect(guardBoundaryShapeAt(marker, 'A2', 0.01, values, z)).toBe(true)
    register('as-literal+multiply-first')
    expect(shape()).toBe(true)
    injectCfPercentOrderReadings({ [CF_SCALE_PERCENT_ORDER_CAP]: 'multiply-first' })
    expect(shape()).toBe(false)
  })
})

describe('the narrowing\'s justification', () => {
  const marker = { ruleType: 'colorScale', sqref: 'E2:E10', scale: { cfvo: [{ type: 'percent', f: '$AH$5' }, { type: 'max' }], colors: ['FF14F028', 'FFE61EFA'] } }
  const values = new Map(E.map((v, i) => [`E${i + 2}`, v]))
  const p = (): number => 33.333333333333336
  it('justifies every cell of a narrowed rule whose thresholds differ, nothing otherwise', () => {
    expect(percentOrderShapeAt(marker, 'E8', values, p)).toBe(false)
    injectCfPercentOrderNarrowings([CF_SCALE_PERCENTILE_ORDER_CAP])
    expect(percentOrderShapeAt(marker, 'E8', values, p)).toBe(false)
    injectCfPercentOrderNarrowings([CF_SCALE_PERCENT_ORDER_CAP])
    expect(percentOrderShapeAt(marker, 'E8', values, p)).toBe(true)
    expect(percentOrderShapeAt(marker, 'E8', values, () => 50)).toBe(false)
  })
})
