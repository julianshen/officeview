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
 * only the cells where some channel's truncation step is 0. The upper segment stays justified whatever is registered.
 */
import { isCellInSqref } from './verify'
import { CF_GUARD_BOUNDARY_CAP, CF_LITERAL_LOWER_CAP, CF_UPPER_WINDOW_CAP, cfGuardBoundaryReading, cfLiteralLowerReading, cfUpperWindowReading, type CfLiteralLowerReading } from '../../src/xlsx/cf/colorScale'
import { hasEvidence } from '../../src/xlsx/cf/evidence'

export type UpperWholeCfvo = Readonly<{ type: string; val?: string; f?: string }>
export type UpperWholeMarker = Readonly<{ ruleType: string; sqref: string; scale?: Readonly<{ cfvo: readonly UpperWholeCfvo[]; colors: readonly string[] }> }>

/** PERCENTILE.INC (Ruling S2-L) over ascending values. */
function percentileInc(sorted: readonly number[], p: number): number {
  const rank = (p / 100) * (sorted.length - 1)
  const lo = Math.floor(rank)
  return lo + 1 < sorted.length ? sorted[lo]! + (rank - lo) * (sorted[lo + 1]! - sorted[lo]!) : sorted[lo]!
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
    if (x.type === 'formula' || x.f !== undefined) return formulaBound?.(x)
    if (x.type === 'min') return min
    if (x.type === 'max') return max
    const n = x.val !== undefined && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(x.val.trim()) ? Number(x.val) : undefined
    if (n === undefined) return undefined
    if (x.type === 'num') return n
    if (x.type === 'percent') return min + (n / 100) * (max - min)
    if (x.type === 'percentile') return percentileInc(scope, n)
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
  return cls !== undefined
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
  return !(guardReadingRegistered() && !bounds.some(formulaCfvo))
}
