/**
 * The upper-segment whole-product withhold (owner ruling A after capture stage2-scale-num-segment), recomputed
 * independently of production for the refusal justifications of Stage 2 (`stage2UpperSegmentWholeProductShape`) and
 * Stage 5 (`stage5JustifiedRefusals`). Nothing here imports production's `safeFraction` or `upperSegmentWholeProduct`: the
 * thresholds come from the marker and the workbook's own cells (and, for a formula midpoint, the published treatment of its
 * result kind), and the predicate is re-implemented here with the same IEEE operations:
 * - the segment is the UPPER one, between the second and third thresholds after the midpoint-above-max re-sort, and the
 *   thresholds are strictly ascending once re-sorted;
 * - the value is strictly inside it;
 * - f = (v - lo) / (hi - lo), with both differences halved first when either overflows (as production normalises), clamped
 *   to [0, 1];
 * - some RGB channel of the upper segment's colours (cfvo colours 2 and 3, positional) has d = |c2 - c1| != 0 with f * d
 *   within JUSTIFY_TRUNCATION_GUARD of a whole number k, 0 < k < d (owner ruling after capture 54fc3388).
 * The upper-segment window follow-up (stage2-scale-upper-window): once a window reading is registered with its evidence
 * (`cfUpperWindowReading`, the registration table only, never production's arithmetic), production paints the `window`
 * class of a segment not bounded by two literal num cfvo, so that class is justified no longer: only the `integer` class
 * (an IEEE-integer product on some channel) and the literal-num segments stay justified.
 * The guard-boundary withhold (owner ruling after capture ed0fccb6): `guardBoundaryShapeAt`, a 2-colour or 3-colour
 * lower-segment cell whose step product lies within the guard strictly below a whole number. The raw-floor follow-up
 * (stage2-scale-raw-floor): once a guard-boundary reading is registered with its evidence (`cfGuardBoundaryReading`, the
 * registration table only), production paints that shape where neither bounding cfvo is a formula, so it is justified no
 * longer there; formula-bounded segments stay justified.
 * The literal-lower follow-up (stage2-scale-literal-lower): once a literal-lower reading is registered with its evidence
 * (`cfLiteralLowerReading`, the registration table only), production paints the LOWER segment of the num-segment shape, so it
 * is justified no longer there (`literalLowerPaints`): `lower-arithmetic` hands the cell to the guard-boundary shape (justified
 * by `guardBoundaryShapeAt` while its own reading is unregistered or its evidence absent); `step-plus-one` leaves justified
 * only the cells where some channel's truncation step is 0. The upper segment stays justified whatever literal-lower reading is
 * registered.
 * The upper-segment follow-up (stage2-scale-upper-segment): once an upper-segment reading is registered with its evidence
 * (`cfUpperSegmentReading`, the registration table only), production paints every upper-segment cell of the whole-product
 * shape (both classes) and of the literal num-segment shape whose bounding cfvo are not formulas, so those are justified no
 * longer (`upperWholeShapeAt`, `stage2NumSegmentShape`); formula-bounded upper segments stay justified.
 */
import { isCellInSqref } from './verify'
import { CF_FORMULA_BOUNDS_CAP, CF_FORMULA_BOUNDS_NEAR_TOP_CAP, CF_FORMULA_BOUNDS_V2_CAP, CF_GUARD_BOUNDARY_CAP, CF_LITERAL_LOWER_CAP, CF_SCALE_PERCENT_ORDER_CAP, CF_SCALE_PERCENTILE_ORDER_CAP, cfFormulaBoundsReading, cfFormulaBoundsV2Reading, cfPercentOrderNarrowed, cfPercentOrderReading, CF_UPPER_NEAR_TOP_CAP, CF_UPPER_SEGMENT_CAP, CF_UPPER_WINDOW_CAP, cfGuardBoundaryReading, cfLiteralLowerReading, cfUpperNearTopReading, cfUpperSegmentReading, cfUpperWindowReading, type CfFormulaBoundsReading, type CfLiteralLowerReading, type CfUpperSegmentReading } from '../../src/xlsx/cf/colorScale'
import { hasEvidence } from '../../src/xlsx/cf/evidence'

export type UpperWholeCfvo = Readonly<{ type: string; val?: string; f?: string }>
export type UpperWholeMarker = Readonly<{ ruleType: string; sqref: string; scale?: Readonly<{ cfvo: readonly UpperWholeCfvo[]; colors: readonly string[] }> }>

/**
 * PERCENTILE.INC (Ruling S2-L) over ascending values, under a percent order (the percent-order follow-up): `lerp-first` as
 * before, `multiply-first` from the 1-based rank 1 + (p·(n − 1))/100 (production's definition, restated).
 */
function percentileInc(sorted: readonly number[], p: number, order: JustifyPercentOrder = 'lerp-first'): number {
  if (order === 'multiply-first') return justifyPercentileRanked(sorted, p, order)
  const rank = (p / 100) * (sorted.length - 1)
  const lo = Math.floor(rank)
  return lo + 1 < sorted.length ? sorted[lo]! + (rank - lo) * (sorted[lo + 1]! - sorted[lo]!) : sorted[lo]!
}

/** PERCENTILE.INC from production's 1-based rank under either reading, restated (the narrowing compares the two orders on it). */
export function justifyPercentileRanked(sorted: readonly number[], p: number, order: 'lerp-first' | 'multiply-first'): number {
  const n = sorted.length
  if (p <= 0) return sorted[0]!
  if (p >= 100) return sorted[n - 1]!
  const rank = order === 'lerp-first' ? 1 + (p / 100) * (n - 1) : 1 + (p * (n - 1)) / 100, ip = Math.floor(rank), i = ip - 1
  return i >= n - 1 ? sorted[n - 1]! : sorted[i]! + (rank - ip) * (sorted[i + 1]! - sorted[i]!)
}

export type JustifyPercentOrder = 'lerp-first' | 'multiply-first' | 'withhold-differing'
/**
 * The percent-order follow-up: the order production applies to one id, read from the registration tables only (never
 * production's arithmetic): the narrowing, else the registered reading with its evidence, else `lerp-first`.
 */
export function justifyPercentOrder(capability: string): JustifyPercentOrder {
  if (cfPercentOrderNarrowed(capability)) return 'withhold-differing'
  const r = cfPercentOrderReading(capability)
  return r !== undefined && hasEvidence(capability) ? r : 'lerp-first'
}
/** A percent threshold under a reading, restated (`lerp-first`: min + (p/100)·(max − min); `multiply-first`: min + ((max − min)·p)/100). */
export const justifyPercent = (min: number, max: number, p: number, order: 'lerp-first' | 'multiply-first'): number =>
  order === 'lerp-first' ? min + (p / 100) * (max - min) : min + ((max - min) * p) / 100
/** A threshold under an order; `differs` when the narrowing applies and the two readings give different doubles. */
function underOrder(order: JustifyPercentOrder, at: (r: 'lerp-first' | 'multiply-first') => number): number | 'differs' {
  if (order !== 'withhold-differing') return at(order)
  const a = at('lerp-first'), b = at('multiply-first')
  return a === b ? a : 'differs'
}

/** The segment fraction, mirroring production's overflow handling (halved differences) and its clamp. */
export function independentFraction(v: number, lo: number, hi: number): number | undefined {
  let span = hi - lo, off = v - lo
  if (!Number.isFinite(span) || !Number.isFinite(off)) { span = hi / 2 - lo / 2; off = v / 2 - lo / 2 }
  const f = off / span
  return Number.isFinite(f) ? Math.max(0, Math.min(1, f)) : undefined
}

/**
 * Production's truncation guard (`CF_TRUNCATION_GUARD` in src/xlsx/cf/colorScale.ts, `floor(f·d + 1e-11)`), restated here
 * rather than imported so the justification stays independent; a test pins the two equal.
 */
export const JUSTIFY_TRUNCATION_GUARD = 1e-11

/** Whether some RGB channel of `from` -> `to` (ARGB hex) has a non-zero span d with f·d within the guard of a whole k, 0 < k < d. */
export function independentWholeProduct(from: string, to: string, f: number): boolean {
  return independentProductClass(from, to, f) !== undefined
}

/**
 * The class of `independentWholeProduct`'s shape: `integer` when some channel's f·d is an IEEE integer k, 0 < k < d;
 * `window` when none is but some lies within the guard of such a k; undefined otherwise.
 */
export function independentProductClass(from: string, to: string, f: number): 'integer' | 'window' | undefined {
  let window = false
  for (const i of [2, 4, 6]) {
    const d = Math.abs(Number.parseInt(to.slice(i, i + 2), 16) - Number.parseInt(from.slice(i, i + 2), 16))
    if (d === 0) continue
    const p = f * d, k = Math.round(p)
    if (!(k > 0 && k < d)) continue
    if (Number.isInteger(p)) return 'integer'
    if (Math.abs(p - k) <= JUSTIFY_TRUNCATION_GUARD) window = true
  }
  return window ? 'window' : undefined
}

/** Whether a window reading is registered with its evidence (production then paints the window class of statistical segments). */
const windowReadingRegistered = (): boolean => cfUpperWindowReading(CF_UPPER_WINDOW_CAP) !== undefined && hasEvidence(CF_UPPER_WINDOW_CAP)
const literalNumCfvo = (x: UpperWholeCfvo): boolean => x.type === 'num' && x.f === undefined
/** Whether a guard-boundary reading is registered with its evidence (production then paints the shape of non-formula segments). */
const guardReadingRegistered = (): boolean => cfGuardBoundaryReading(CF_GUARD_BOUNDARY_CAP) !== undefined && hasEvidence(CF_GUARD_BOUNDARY_CAP)
const formulaCfvo = (x: UpperWholeCfvo): boolean => x.type === 'formula' || x.f !== undefined
/** The literal-lower reading registered with its evidence (undefined: none; production then withholds the whole num-segment shape). */
export const literalLowerReadingRegistered = (): CfLiteralLowerReading | undefined => {
  const r = cfLiteralLowerReading(CF_LITERAL_LOWER_CAP)
  return r !== undefined && hasEvidence(CF_LITERAL_LOWER_CAP) ? r : undefined
}

/**
 * The formula-bounds follow-up (stage2-scale-formula-bounds): the formula-bounds reading registered with its evidence
 * (`cfFormulaBoundsReading`, the registration table only; undefined: none, so formula-bounded cells of the settled shapes stay
 * justified). Under `fifteen-digit` a formula result is rounded (`independentFifteenDigit`) before its threshold is resolved.
 */
export const formulaBoundsReadingRegistered = (): CfFormulaBoundsReading | undefined => {
  const r = cfFormulaBoundsReading(CF_FORMULA_BOUNDS_CAP)
  if (r !== undefined && hasEvidence(CF_FORMULA_BOUNDS_CAP)) return r
  return formulaBoundsV2Registered()
}
/** The v1 id alone (the near-top sub-shape is decided with it only). */
const formulaBoundsV1Registered = (): CfFormulaBoundsReading | undefined => {
  const r = cfFormulaBoundsReading(CF_FORMULA_BOUNDS_CAP)
  return r !== undefined && hasEvidence(CF_FORMULA_BOUNDS_CAP) ? r : undefined
}
/**
 * The percent-order follow-up: `scale:formula-bounds-v2` registered with its evidence and its order equal to the colour-scale
 * percent order production applies (`justifyPercentOrder`), split to its formula-bounds reading (undefined: none).
 */
export const formulaBoundsV2Registered = (): CfFormulaBoundsReading | undefined => {
  const v = cfFormulaBoundsV2Reading(CF_FORMULA_BOUNDS_V2_CAP)
  if (v === undefined || !hasEvidence(CF_FORMULA_BOUNDS_V2_CAP)) return undefined
  const [reading, order] = v.split('+') as [CfFormulaBoundsReading, string]
  return justifyPercentOrder(CF_SCALE_PERCENT_ORDER_CAP) === order ? reading : undefined
}
/** Whether a formula-bounds reading admits a segment bounded by `bounds`: the v1 id, or the v2 id unless a bound is a formula percentile. */
const formulaBoundsAdmitBounds = (bounds: readonly UpperWholeCfvo[]): boolean =>
  formulaBoundsV1Registered() !== undefined || (formulaBoundsV2Registered() !== undefined && !bounds.some(x => x.type === 'percentile' && formulaCfvo(x)))
/** `fifteen-digit`'s rounding restated (production's `cfFormulaBoundFifteenDigit`; a test pins the two equal). */
export const independentFifteenDigit = (x: number): number => (Number.isFinite(x) ? Number(x.toPrecision(15)) : x)
/** Whether a segment bounded by `lo` / `hi` is painted in a settled shape: no formula bound, or a formula-bounds reading registered. */
const formulaBoundsAdmit = (lo: UpperWholeCfvo, hi: UpperWholeCfvo): boolean => (!formulaCfvo(lo) && !formulaCfvo(hi)) || formulaBoundsAdmitBounds([lo, hi])
const opaqueHex = (argb: string): boolean => argb.slice(0, 2).toUpperCase() === 'FF'

/** Whether production's upper-segment reading `r` paints between `from` and `to` (ARGB hex): `index-position` only between two opaque colours. */
export const upperSegmentReadingPaints = (r: CfUpperSegmentReading | undefined, from: string, to: string): boolean =>
  r !== undefined && (r !== 'index-position' || (from.slice(0, 2).toUpperCase() === 'FF' && to.slice(0, 2).toUpperCase() === 'FF'))

/** The upper-segment reading registered with its evidence (undefined: none; production then withholds the upper shapes). */
export const upperSegmentReadingRegistered = (): CfUpperSegmentReading | undefined => {
  const r = cfUpperSegmentReading(CF_UPPER_SEGMENT_CAP)
  return r !== undefined && hasEvidence(CF_UPPER_SEGMENT_CAP) ? r : undefined
}

/**
 * Whether production paints a literal-num lower-segment cell of the num-segment shape (colours `from` -> `to`, fraction f) under
 * the registered literal-lower reading, restated: `lower-arithmetic` always (its guard-boundary cells are then
 * `guardBoundaryShapeAt`'s); `step-plus-one` where every ARGB channel with d != 0 has a truncation step
 * floor(f·d + JUSTIFY_TRUNCATION_GUARD) >= 1; none registered, never.
 */
export function literalLowerPaints(from: string, to: string, f: number): boolean {
  const r = literalLowerReadingRegistered()
  if (r === undefined || !Number.isFinite(f)) return false
  if (r === 'lower-arithmetic') return true
  for (const i of [0, 2, 4, 6]) {
    const d = Math.abs(Number.parseInt(to.slice(i, i + 2), 16) - Number.parseInt(from.slice(i, i + 2), 16))
    if (d > 0 && Math.floor(Math.max(0, Math.min(1, f)) * d + JUSTIFY_TRUNCATION_GUARD) < 1) return false
  }
  return true
}

/**
 * Whether `ref` (value `value`) is in the shape under `marker`. `values` holds the numeric cells of the sheet (the scope's
 * statistics); `formulaBound` resolves a formula cfvo to its number (undefined: unresolved, so no justification).
 */
/**
 * The thresholds of a colour-scale marker, positional (cfvo order), recomputed from the scope's numbers: min / max, literal
 * num, percent (min + p/100·(max − min)), PERCENTILE.INC and a formula bound through `formulaBound`. Undefined when the ref
 * is outside the sqref, the scope holds no number or any threshold is unresolved or non-finite.
 */
function independentThresholds(marker: UpperWholeMarker, ref: string, values: ReadonlyMap<string, number>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): number[] | undefined {
  if (marker.ruleType !== 'colorScale' || !marker.scale || !isCellInSqref(ref, marker.sqref)) return undefined
  const scope = [...values].filter(([k]) => isCellInSqref(k, marker.sqref)).map(([, v]) => v).sort((x, y) => x - y)
  if (scope.length === 0) return undefined
  const min = scope[0]!, max = scope[scope.length - 1]!
  const at = (x: UpperWholeCfvo): number | undefined => {
    if (x.type === 'min') return min
    if (x.type === 'max') return max
    // A formula result feeds its cfvo type (a `formula` / `num` bound is the number, a percent / percentile one its p), rounded
    // first under a registered fifteen-digit reading.
    const raw = x.type === 'formula' || x.f !== undefined ? formulaBound?.(x)
      : x.val !== undefined && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(x.val.trim()) ? Number(x.val) : undefined
    if (raw === undefined) return undefined
    const n = (x.type === 'formula' || x.f !== undefined) && formulaBoundsReadingRegistered() === 'fifteen-digit' ? independentFifteenDigit(raw) : raw
    if (x.type === 'num' || x.type === 'formula') return n
    // The percent-order follow-up: the registered order of each bound type; under the narrowing a threshold whose orders
    // differ is withheld by its own justification (`percentOrderShapeAt`), never this shape.
    const known = (t: number | 'differs'): number | undefined => (t === 'differs' ? undefined : t)
    if (x.type === 'percent') return known(underOrder(justifyPercentOrder(CF_SCALE_PERCENT_ORDER_CAP), r => justifyPercent(min, max, n, r)))
    if (x.type === 'percentile') {
      const order = justifyPercentOrder(CF_SCALE_PERCENTILE_ORDER_CAP)
      // Under the narrowing the two orders are compared on production's 1-based rank (`justifyPercentileRanked`).
      return known(underOrder(order, r => (order === 'withhold-differing' ? justifyPercentileRanked(scope, n, r) : percentileInc(scope, n, r))))
    }
    return undefined
  }
  const t = marker.scale.cfvo.map(at)
  return t.some(x => x === undefined || !Number.isFinite(x)) ? undefined : (t as number[])
}

/**
 * Whether `ref` (value `value`) is in the shape under `marker`. `values` holds the numeric cells of the sheet (the scope's
 * statistics); `formulaBound` resolves a formula cfvo to its number (undefined: unresolved, so no justification).
 */
export function upperWholeShapeAt(marker: UpperWholeMarker, ref: string, value: unknown, values: ReadonlyMap<string, number>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) return false
  if (marker.ruleType !== 'colorScale' || marker.scale?.cfvo.length !== 3 || marker.scale.colors.length !== 3) return false
  const t = independentThresholds(marker, ref, values, formulaBound)
  if (!t) return false
  const [t0, t1, t2] = t as [number, number, number]
  const [mid, top] = t1 > t2 ? [t2, t1] : [t1, t2]
  if (!(t0 < mid && mid < top) || !(value > mid && value < top)) return false
  const f = independentFraction(value, mid, top)
  if (f === undefined) return false
  const cls = independentProductClass(marker.scale.colors[1]!, marker.scale.colors[2]!, f)
  // The upper segment is bounded by cfvo 2 and 3 in either order (the midpoint-above-max re-sort swaps them).
  if (cls === 'window' && !(literalNumCfvo(marker.scale.cfvo[1]!) && literalNumCfvo(marker.scale.cfvo[2]!)) && windowReadingRegistered()) return false
  // The upper-segment follow-up: a registered upper-segment reading (with its evidence) paints both classes unless a bound is a formula.
  // A formula-bounded segment only with a formula-bounds reading registered with its evidence.
  if (cls !== undefined && upperSegmentReadingPaints(upperSegmentReadingRegistered(), marker.scale.colors[1]!, marker.scale.colors[2]!) && formulaBoundsAdmit(marker.scale.cfvo[1]!, marker.scale.cfvo[2]!)) return false
  return cls !== undefined
}

/**
 * The near-top sub-shape (production's `upperSegmentNearTop`, reason `cf-unmeasured-upper-near-top`), restated: strictly inside
 * the UPPER segment (after the re-sort), some RGB channel of the upper colours (cfvo colours 2 and 3) with d != 0 has
 * d − f·d <= JUSTIFY_TRUNCATION_GUARD. Justified unless the upper-segment and near-top ids are registered with the same
 * reading, both with their evidence, and neither upper bound is a formula.
 */
export function upperNearTopShapeAt(marker: UpperWholeMarker, ref: string, value: unknown, values: ReadonlyMap<string, number>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) return false
  if (marker.ruleType !== 'colorScale' || marker.scale?.cfvo.length !== 3 || marker.scale.colors.length !== 3) return false
  const t = independentThresholds(marker, ref, values, formulaBound)
  if (!t) return false
  const [t0, t1, t2] = t as [number, number, number]
  const [mid, top] = t1 > t2 ? [t2, t1] : [t1, t2]
  if (!(t0 < mid && mid < top) || !(value > mid && value < top)) return false
  const f = independentFraction(value, mid, top)
  if (f === undefined) return false
  const from = marker.scale.colors[1]!, to = marker.scale.colors[2]!
  const near = [2, 4, 6].some(i => { const d = Math.abs(Number.parseInt(to.slice(i, i + 2), 16) - Number.parseInt(from.slice(i, i + 2), 16)); return d !== 0 && d - f * d <= JUSTIFY_TRUNCATION_GUARD })
  if (!near) return false
  const r = upperSegmentReadingRegistered()
  const formula = formulaCfvo(marker.scale.cfvo[1]!) || formulaCfvo(marker.scale.cfvo[2]!)
  const fr = formulaBoundsV1Registered()
  // A formula-bounded near-top cell: also the formula near-top id at the same formula-bounds reading, with its evidence (the
  // v1 id only: scale:formula-bounds-v2 never lifts the near-top sub-shape).
  const formulaPainted = !formula || (fr !== undefined && cfFormulaBoundsReading(CF_FORMULA_BOUNDS_NEAR_TOP_CAP) === fr && hasEvidence(CF_FORMULA_BOUNDS_NEAR_TOP_CAP))
  const painted = upperSegmentReadingPaints(r, from, to) && cfUpperNearTopReading(CF_UPPER_NEAR_TOP_CAP) === r && hasEvidence(CF_UPPER_NEAR_TOP_CAP) && formulaPainted
  return !painted
}

/**
 * The guard-boundary withhold (owner ruling after capture ed0fccb6, stage2-scale-upper-window; production's
 * `guardBoundaryProduct`, reason `cf-unmeasured-guard-boundary`), restated: some ARGB channel of `from` -> `to` has a span
 * d != 0 whose product p = f·d has floor(p + JUSTIFY_TRUNCATION_GUARD) != floor(p), i.e. p lies within the guard strictly
 * below a whole number.
 */
export function independentGuardBoundary(from: string, to: string, f: number): boolean {
  for (const i of [0, 2, 4, 6]) {
    const d = Math.abs(Number.parseInt(to.slice(i, i + 2), 16) - Number.parseInt(from.slice(i, i + 2), 16))
    if (!(d > 0)) continue
    const p = f * d
    if (Math.floor(p + JUSTIFY_TRUNCATION_GUARD) !== Math.floor(p)) return true
  }
  return false
}

/**
 * Whether `ref` (value `value`) is in the guard-boundary shape under `marker`: a cell strictly inside a 2-colour scale (its
 * thresholds sorted, colours positional, as production inverts them) or strictly inside the LOWER segment of a 3-colour
 * scale (thresholds strictly ascending after the midpoint-above-max re-sort; colours 1 -> 2, positional), at
 * `independentGuardBoundary`. A 3-colour segment bounded by two literal num cfvo whose f·2^20 is not an integer is the
 * num-segment shape instead (its precedence), and is not this shape, unless `lower-arithmetic` is registered with its evidence.
 */
export function guardBoundaryShapeAt(marker: UpperWholeMarker, ref: string, value: unknown, values: ReadonlyMap<string, number>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value) || !marker.scale) return false
  const n = marker.scale.cfvo.length
  if ((n !== 2 && n !== 3) || marker.scale.colors.length !== n) return false
  const t = independentThresholds(marker, ref, values, formulaBound)
  if (!t) return false
  let lo: number, hi: number
  let bounds: UpperWholeCfvo[]
  if (n === 2) {
    [lo, hi] = t[0]! <= t[1]! ? [t[0]!, t[1]!] : [t[1]!, t[0]!]
    bounds = [...marker.scale.cfvo]
  } else {
    const [t0, t1, t2] = t as [number, number, number]
    const mid = t1 > t2 ? t2 : t1
    const top = t1 > t2 ? t1 : t2
    if (!(t0 < mid && mid < top)) return false
    const midCfvo = marker.scale.cfvo[t1 > t2 ? 2 : 1]!
    lo = t0; hi = mid
    bounds = [marker.scale.cfvo[0]!, midCfvo]
    // The num-segment shape (its precedence), unless lower-arithmetic is registered with its evidence: production then paints
    // the cell as a non-literal lower segment, through this shape.
    if (literalNumCfvo(marker.scale.cfvo[0]!) && literalNumCfvo(midCfvo) && value > lo && value < hi && !Number.isInteger(((value - lo) / (hi - lo)) * 1048576) && literalLowerReadingRegistered() !== 'lower-arithmetic') return false
  }
  if (!(value > lo && value < hi)) return false
  const f = independentFraction(value, lo, hi)
  if (f === undefined || !independentGuardBoundary(marker.scale.colors[0]!, marker.scale.colors[1]!, f)) return false
  // A formula-bounded segment is painted only with a formula-bounds reading registered with its evidence, between opaque colours.
  const formulaOk = !bounds.some(formulaCfvo) || (formulaBoundsAdmitBounds(bounds) && opaqueHex(marker.scale.colors[0]!) && opaqueHex(marker.scale.colors[1]!))
  return !(guardReadingRegistered() && formulaOk)
}

/**
 * The percent-order follow-up, the pre-committed narrowing (`withhold-differing`, production's `cf-unmeasured-percent-order`),
 * restated: a colour-scale rule with a percent (percentile) cfvo whose id is narrowed and whose two orders give different
 * doubles over the rule's scope withholds every cell of the rule. Thresholds come from the marker and the scope's numbers
 * (`formulaBound` resolves a formula p), never from production.
 */
export function percentOrderShapeAt(marker: UpperWholeMarker, ref: string, values: ReadonlyMap<string, number>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): boolean {
  if (marker.ruleType !== 'colorScale' || !marker.scale || !isCellInSqref(ref, marker.sqref)) return false
  const scope = [...values].filter(([k]) => isCellInSqref(k, marker.sqref)).map(([, v]) => v).sort((x, y) => x - y)
  if (scope.length === 0) return false
  return percentOrderDiffers(marker.scale.cfvo, scope, { percent: CF_SCALE_PERCENT_ORDER_CAP, percentile: CF_SCALE_PERCENTILE_ORDER_CAP }, formulaBound)
}

/** Whether some percent / percentile cfvo of a narrowed id has thresholds that differ between the orders over `scope` (ascending). */
export function percentOrderDiffers(cfvo: readonly UpperWholeCfvo[], scope: readonly number[], caps: Readonly<{ percent: string; percentile: string }>, formulaBound?: (cfvo: UpperWholeCfvo) => number | undefined): boolean {
  const min = scope[0]!, max = scope[scope.length - 1]!
  return cfvo.some(x => {
    if (x.type !== 'percent' && x.type !== 'percentile') return false
    if (!cfPercentOrderNarrowed(x.type === 'percent' ? caps.percent : caps.percentile)) return false
    const raw = x.f !== undefined ? formulaBound?.(x) : x.val !== undefined && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(x.val.trim()) ? Number(x.val) : undefined
    if (raw === undefined || !(raw >= 0 && raw <= 100)) return false
    const [a, b] = (['lerp-first', 'multiply-first'] as const).map(r => (x.type === 'percent' ? justifyPercent(min, max, raw, r) : justifyPercentileRanked(scope, raw, r)))
    return a !== b
  })
}
