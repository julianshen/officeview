/**
 * Color scale evaluation, percentile order statistics, and interpolation (Stage 2, Task 14).
 *
 * Implements pure evaluation for colour-scale rules:
 * - Exact percentile order statistics over sorted numeric arrays (PERCENTILE.INC / EXC)
 * - Threshold resolution against measured range statistics
 * - Closed interpolation-variant set (colour space and rounding ladders per Ruling S2-B)
 * - Degenerate scope and value handling (Ruling S2-E)
 * - Non-numeric participation policy (Ruling S2-G)
 *
 * The evaluator is completely pure: no DOM, no clock, no formula AST, zero global state.
 */

import type {
  CfLiteral,
  CfOutcome,
  CfScaleBound,
  CfScalePoint,
  CfDefinition,
} from './model'
import type { CfRangeStatistics } from './statistics'
import { hasEvidence as defaultHasEvidence } from './evidence'

export type CfPercentileVariant = 'inc' | 'exc'
export type CfRoundingVariant = 'truncate' | 'half-up' | 'half-even'
export type CfColorSpaceVariant = 'sRGB' | 'linear-light' | 'hsl'

/**
 * Computes the exact percentile order statistic over an ascending sorted array.
 * Closed variant set:
 * - 'inc' (PERCENTILE.INC): rank = 1 + (p/100) * (N - 1)
 * - 'exc' (PERCENTILE.EXC): rank = (p/100) * (N + 1)
 */
/** `lo + f * (hi - lo)`, bit-identical to that formula whenever the span is finite; finite endpoints of opposite sign
 * near the double limit overflow the span, so those use the convex form `lo * (1 - f) + hi * f`. */
export function lerp(lo: number, hi: number, f: number): number {
  const span = hi - lo
  return Number.isFinite(span) ? lo + f * span : lo * (1 - f) + hi * f
}

export function percentileOrderStatistic(
  sortedAsc: Float64Array | readonly number[],
  percentile: number,
  variant: CfPercentileVariant,
): number | undefined {
  const n = sortedAsc.length
  if (n === 0) return undefined

  const p = percentile / 100
  if (!Number.isFinite(p)) return undefined

  if (variant === 'inc') {
    if (p < 0 || p > 1) return undefined
    if (p === 0) return sortedAsc[0]!
    if (p === 1) return sortedAsc[n - 1]!

    const rank = 1 + p * (n - 1)
    const intPart = Math.floor(rank)
    const frac = rank - intPart
    const index = intPart - 1 // 0-indexed

    if (index >= n - 1) return sortedAsc[n - 1]!
    const v0 = sortedAsc[index]!
    const v1 = sortedAsc[index + 1]!
    return lerp(v0, v1, frac)
  } else {
    // variant === 'exc'
    if (p <= 0 || p >= 1) return undefined
    const rank = p * (n + 1)
    if (rank < 1 || rank > n) return undefined

    const intPart = Math.floor(rank)
    const frac = rank - intPart
    const index = intPart - 1

    if (index >= n - 1) return sortedAsc[n - 1]!
    const v0 = sortedAsc[index]!
    const v1 = sortedAsc[index + 1]!
    return lerp(v0, v1, frac)
  }
}

/**
 * Evaluates channel interpolation using a closed rounding-variant set.
 * 'truncate' mirrors Excel's measured display-format behavior:
 * floor towards target when increasing, floor towards start when decreasing.
 */
/**
 * The truncating interpolation's guard: a step product within this of the next whole number truncates up to it
 * (`floor(f·d + guard)`). The upper-segment whole-product withhold uses the same window, and
 * scripts/cf-refs/upper-whole-justify.ts mirrors it as `JUSTIFY_TRUNCATION_GUARD` (pinned equal by a test).
 */
export const CF_TRUNCATION_GUARD = 1e-11

export function interpolateChannel(
  c0: number,
  c1: number,
  fraction: number,
  rounding: CfRoundingVariant = 'truncate',
): number {
  const f = Math.max(0, Math.min(1, fraction))
  if (rounding === 'truncate') {
    // CF_TRUNCATION_GUARD (1e-11) absorbs floating-point representation error from threshold division
    if (c1 >= c0) {
      return c0 + Math.floor(f * (c1 - c0) + CF_TRUNCATION_GUARD)
    } else {
      return c0 - Math.floor(f * (c0 - c1) + CF_TRUNCATION_GUARD)
    }
  } else if (rounding === 'half-up') {
    return Math.round(c0 + f * (c1 - c0))
  } else {
    // half-even
    const val = c0 + f * (c1 - c0)
    const floor = Math.floor(val)
    const diff = val - floor
    if (Math.abs(diff - 0.5) < 1e-9) {
      return floor % 2 === 0 ? floor : floor + 1
    }
    return Math.round(val)
  }
}

function srgbToLinear(c: number): number {
  const s = c / 255
  return s <= 0.04045 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4)
}

function linearToSrgb(l: number): number {
  const c = l <= 0.0031308 ? l * 12.92 : 1.055 * Math.pow(Math.max(0, l), 1 / 2.4) - 0.055
  return Math.max(0, Math.min(255, c * 255))
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / 255, gn = g / 255, bn = b / 255
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn)
  let h = 0, s = 0
  const l = (max + min) / 2

  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    switch (max) {
      case rn: h = (gn - bn) / d + (gn < bn ? 6 : 0); break
      case gn: h = (bn - rn) / d + 2; break
      case bn: h = (rn - gn) / d + 4; break
    }
    h /= 6
  }
  return [h, s, l]
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = l * 255
    return [v, v, v]
  }
  const hue2rgb = (p: number, q: number, t: number): number => {
    let tn = t
    if (tn < 0) tn += 1
    if (tn > 1) tn -= 1
    if (tn < 1 / 6) return p + (q - p) * 6 * tn
    if (tn < 1 / 2) return q
    if (tn < 2 / 3) return p + (q - p) * (2 / 3 - tn) * 6
    return p
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const r = hue2rgb(p, q, h + 1 / 3) * 255
  const g = hue2rgb(p, q, h) * 255
  const b = hue2rgb(p, q, h - 1 / 3) * 255
  return [r, g, b]
}

/**
 * Pure colour interpolation over a closed variant set of colour space and rounding ladder.
 */
export function interpolateScaleColor(
  argb0: number,
  argb1: number,
  fraction: number,
  colorSpace: CfColorSpaceVariant = 'sRGB',
  rounding: CfRoundingVariant = 'truncate',
): number {
  const f = Math.max(0, Math.min(1, fraction))
  const a0 = (argb0 >>> 24) & 0xff
  const r0 = (argb0 >> 16) & 0xff
  const g0 = (argb0 >> 8) & 0xff
  const b0 = argb0 & 0xff

  const a1 = (argb1 >>> 24) & 0xff
  const r1 = (argb1 >> 16) & 0xff
  const g1 = (argb1 >> 8) & 0xff
  const b1 = argb1 & 0xff

  const a = interpolateChannel(a0, a1, f, rounding)

  if (colorSpace === 'sRGB') {
    const r = interpolateChannel(r0, r1, f, rounding)
    const g = interpolateChannel(g0, g1, f, rounding)
    const b = interpolateChannel(b0, b1, f, rounding)
    return (((a << 24) | (r << 16) | (g << 8) | b) >>> 0)
  } else if (colorSpace === 'linear-light') {
    const rLin0 = srgbToLinear(r0), rLin1 = srgbToLinear(r1)
    const gLin0 = srgbToLinear(g0), gLin1 = srgbToLinear(g1)
    const bLin0 = srgbToLinear(b0), bLin1 = srgbToLinear(b1)

    const rLin = rLin0 + f * (rLin1 - rLin0)
    const gLin = gLin0 + f * (gLin1 - gLin0)
    const bLin = bLin0 + f * (bLin1 - bLin0)

    const rTarget = linearToSrgb(rLin)
    const gTarget = linearToSrgb(gLin)
    const bTarget = linearToSrgb(bLin)

    const r = interpolateChannel(r0, Math.round(rTarget), 1, rounding)
    const g = interpolateChannel(g0, Math.round(gTarget), 1, rounding)
    const b = interpolateChannel(b0, Math.round(bTarget), 1, rounding)
    return (((a << 24) | (r << 16) | (g << 8) | b) >>> 0)
  } else {
    // HSL
    const [h0, s0, l0] = rgbToHsl(r0, g0, b0)
    const [h1, s1, l1] = rgbToHsl(r1, g1, b1)

    const h = h0 + f * (h1 - h0)
    const s = s0 + f * (s1 - s0)
    const l = l0 + f * (l1 - l0)

    const [rTarget, gTarget, bTarget] = hslToRgb(h, s, l)
    const r = interpolateChannel(r0, Math.round(rTarget), 1, rounding)
    const g = interpolateChannel(g0, Math.round(gTarget), 1, rounding)
    const b = interpolateChannel(b0, Math.round(bTarget), 1, rounding)
    return (((a << 24) | (r << 16) | (g << 8) | b) >>> 0)
  }
}

/**
 * Computes all scale:* capability IDs required for scale evaluation.
 */
export function scaleEvaluationRequiredCapabilities(
  points: readonly CfScalePoint[],
): Set<string> {
  const req = new Set<string>()

  // Ruling T13-A & Ruling S2-B: colour space and rounding ladder are required
  req.add('scale:interpolation-color-space')
  req.add('scale:interpolation-rounding-ladder')

  if (points.length === 3) {
    req.add('scale:interpolation-spec-palette')
  }

  for (const p of points) {
    const bound = p.bound
    if (bound.type === 'min' || bound.type === 'max') {
      req.add('scale:thresholds-min-max')
    } else if (bound.type === 'num') {
      req.add('scale:thresholds-num-bounds')
    } else if (bound.type === 'percent') {
      req.add('scale:thresholds-percent')
      if (bound.value < 0 || bound.value > 100) {
        req.add('scale:thresholds-percent-percentile-out-of-bounds')
      }
    } else if (bound.type === 'percentile') {
      req.add('scale:thresholds-percentile-inc-exc')
      if (bound.value < 0 || bound.value > 100) {
        req.add('scale:thresholds-percent-percentile-out-of-bounds')
      }
    }
  }

  return req
}

/** The threshold-bearing part of a scale point; bars reuse threshold resolution with bounds alone (Task 20). */
export type CfThresholdPoint = Readonly<{ bound: CfScaleBound }>

function extractPoints<P extends CfThresholdPoint>(pointsOrRule: readonly P[] | { points: readonly P[] } | CfDefinition): readonly P[] | undefined {
  if (Array.isArray(pointsOrRule)) return pointsOrRule
  if (typeof pointsOrRule === 'object' && pointsOrRule !== null && 'points' in pointsOrRule) {
    return (pointsOrRule as { points: readonly P[] }).points
  }
  return undefined
}

/**
 * Resolves numeric threshold boundaries for a colour-scale rule.
 * Pure function: fails closed when evidence is absent or boundaries are degenerate/unsupported.
 */
export function resolveScaleThresholds(
  pointsOrRule: readonly CfThresholdPoint[] | { points: readonly CfThresholdPoint[] } | CfDefinition,
  stats: CfRangeStatistics,
  options: { hasEvidence?: (cap: string) => boolean } = {},
): CfOutcome<readonly number[]> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  const points = extractPoints(pointsOrRule)

  if (!points || (points.length !== 2 && points.length !== 3)) {
    return { kind: 'unresolved', reason: 'cf-invalid-rule' }
  }

  // A formula bound is never a threshold: the service substitutes its evaluated number first (Ruling S5-H).
  if (points.some(p => p.bound.type === 'formula')) return { kind: 'unresolved', reason: 'cf-formula-unevaluated' }

  // Feature gates for bounds
  for (const p of points) {
    const b = p.bound
    if (b.type === 'min' || b.type === 'max') {
      if (!checkEvidence('scale:thresholds-min-max')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
    } else if (b.type === 'num') {
      if (!checkEvidence('scale:thresholds-num-bounds')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
    } else if (b.type === 'percent') {
      if (!checkEvidence('scale:thresholds-percent')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
      if ((b.value < 0 || b.value > 100) && !checkEvidence('scale:thresholds-percent-percentile-out-of-bounds')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
    } else if (b.type === 'percentile') {
      if (!checkEvidence('scale:thresholds-percentile-inc-exc')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
      if ((b.value < 0 || b.value > 100) && !checkEvidence('scale:thresholds-percent-percentile-out-of-bounds')) {
        return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
      }
    }
  }

  if (points.length === 3 && !checkEvidence('scale:interpolation-spec-palette')) {
    return { kind: 'unresolved', reason: 'cf-unsupported-rule' }
  }

  // Ruling F-2 (final review I3-B): the only measured out-of-range shape (percent -10 / percentile 150) paints
  // NOTHING in Excel, and no colour was observed for a lone out-of-range percent or percentile. The scale is
  // withheld outright; a computed extrapolated colour is never painted for an unmeasured shape.
  for (const p of points) {
    const b = p.bound
    if ((b.type === 'percent' || b.type === 'percentile') && (b.value < 0 || b.value > 100)) {
      return { kind: 'unresolved', reason: 'cf-invalid-bound' }
    }
  }

  // Degenerate scope checks (Ruling S2-E)
  if (stats.numericCount === 0) {
    return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
  }

  if (stats.numericCount === 1) {
    if (!checkEvidence('scale:degenerate-single-cell')) {
      return { kind: 'unresolved', reason: 'cf-degenerate-scope' }
    }
  } else if (stats.min !== undefined && stats.max !== undefined && stats.min === stats.max) {
    if (!checkEvidence('scale:degenerate-equal-values')) {
      return { kind: 'unresolved', reason: 'cf-degenerate-scope' }
    }
  }

  const thresholds: number[] = []

  for (let i = 0; i < points.length; i++) {
    const bound = points[i]!.bound
    if (bound.type === 'min') {
      if (stats.min === undefined || !Number.isFinite(stats.min)) {
        return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
      }
      thresholds.push(stats.min)
    } else if (bound.type === 'max') {
      if (stats.max === undefined || !Number.isFinite(stats.max)) {
        return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
      }
      thresholds.push(stats.max)
    } else if (bound.type === 'num') {
      if (!Number.isFinite(bound.value)) {
        return { kind: 'unresolved', reason: 'cf-invalid-bound' }
      }
      thresholds.push(bound.value)
    } else if (bound.type === 'percent') {
      if (stats.min === undefined || stats.max === undefined || !Number.isFinite(stats.min) || !Number.isFinite(stats.max)) {
        return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
      }
      const val = lerp(stats.min, stats.max, bound.value / 100)
      if (!Number.isFinite(val)) {
        return { kind: 'unresolved', reason: 'cf-invalid-bound' }
      }
      thresholds.push(val)
    } else if (bound.type === 'percentile') {
      if (stats.sortedAsc.length === 0) {
        return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
      }
      // Excel definition is PERCENTILE.INC per Ruling S2-L
      const val = percentileOrderStatistic(stats.sortedAsc, bound.value, 'inc')
      if (val === undefined || !Number.isFinite(val)) {
        return { kind: 'unresolved', reason: 'cf-invalid-bound' }
      }
      thresholds.push(val)
    }
  }

  // Degenerate threshold order checks (Ruling S2-E)
  if (points.length === 2) {
    const [t0, t1] = thresholds as [number, number]
    if (t0 > t1) {
      if (!checkEvidence('scale:degenerate-inverted-thresholds')) {
        return { kind: 'unresolved', reason: 'cf-inverted-thresholds' }
      }
    } else if (t0 === t1) {
      // Ruling F-2 (I3-A): equal thresholds were measured only over all-equal data (and one numeric cell).
      // Equal thresholds over distinct data are a different, unmeasured shape.
      const allEqual = stats.numericCount > 0 && stats.min === stats.max
      if (stats.numericCount === 1) {
        if (!checkEvidence('scale:degenerate-single-cell')) {
          return { kind: 'unresolved', reason: 'cf-degenerate-scope' }
        }
      } else if (!allEqual || !checkEvidence('scale:degenerate-equal-values')) {
        return { kind: 'unresolved', reason: 'cf-degenerate-scope' }
      }
    }
  } else if (points.length === 3) {
    const [t0, t1, t2] = thresholds as [number, number, number]
    // Ruling F-2 (I3-A): only two 3-point orderings were measured: strictly ascending (t0 < t1 < t2) and the
    // midpoint above the maximum (t0 < t2 < t1). Every inversion or tie is a different shape: withheld.
    const ascending = t0 < t1 && t1 < t2
    const midpointAboveMax = t0 < t2 && t2 < t1
    if (!ascending && !midpointAboveMax) {
      return { kind: 'unresolved', reason: t0 > t1 || t0 > t2 ? 'cf-inverted-thresholds' : 'cf-degenerate-scope' }
    }
    if (midpointAboveMax && !checkEvidence('scale:degenerate-midpoint-above-max')) {
      return { kind: 'unresolved', reason: 'cf-midpoint-above-max' }
    }
  }

  return { kind: 'ready', value: thresholds }
}

/** Clamped position of `v` in [lo, hi], overflow-safe: finite thresholds of opposite sign near the double limit
 * overflow `hi - lo`, so such spans are normalised in halves. Undefined only if the result is still non-finite, in which
 * case the caller withholds instead of publishing a wrong colour. */
export function safeFraction(v: number, lo: number, hi: number): number | undefined {
  let span = hi - lo, offset = v - lo
  if (!Number.isFinite(span) || !Number.isFinite(offset)) { span = hi / 2 - lo / 2; offset = v / 2 - lo / 2 }
  const fraction = offset / span
  return Number.isFinite(fraction) ? Math.max(0, Math.min(1, fraction)) : undefined
}

/**
 * Finding (capture e448d831, stage2-scale-scope-disjoint twins T4 and V3; user ruling "withhold the num shape"): on a
 * 3-colour scale whose segment is bounded by two literal `num` cfvo (a formula bound resolved to a number is not
 * included: no capture shows that shape), a value strictly inside the segment where f·2^20 is not an integer rendered
 * in Excel 16.115 one unit HIGHER per RGB channel than production's truncating interpolation (T4: CC/DD/82 against CB/DC/81; V3: 98/CE/7F against 97/CD/7E), i.e. the descending step one unit smaller in magnitude. Excel's rule is
 * unexplained, so the rounding is unchanged and such a cell is withheld (`cf-unmeasured-num-segment-fraction`) until a
 * follow-up case pre-registers the readings. Endpoints and fractions with f·2^20 an integer paint as before.
 */
export function numSegmentUnmeasured(lo: CfScalePoint, hi: CfScalePoint, value: number, tLo: number, tHi: number, fraction: number): boolean {
  return literalNumBound(lo) && literalNumBound(hi) && value > tLo && value < tHi && !Number.isInteger(fraction * 1048576)
}

/** The kind of the cfvo bounding a segment, as the scale evaluated it: a formula bound resolved to a number is `num:formula`. */
export type CfScaleSegmentBoundKind = 'min' | 'max' | 'num' | 'num:formula' | 'percent' | 'percent:formula' | 'percentile' | 'percentile:formula'

/**
 * One colour-scale segment evaluation, everything Excel's arithmetic could read: the value, the segment's thresholds and
 * colours (after the 2-colour inversion and the midpoint-above-max re-sort, `lo < hi`), production's clamped fraction, the
 * bounding cfvo kinds and every threshold of the scale (sorted ascending).
 */
export interface CfScaleSegment {
  readonly value: number
  readonly lo: number
  readonly hi: number
  /** `safeFraction(value, lo, hi)`, clamped to [0, 1]. */
  readonly fraction: number
  /** ARGB at `lo` and at `hi`. */
  readonly c0: number
  readonly c1: number
  readonly colours: 2 | 3
  /** 0: the scale's lower (or only) segment; 1: the upper segment of a 3-colour scale. */
  readonly segment: 0 | 1
  readonly loKind: CfScaleSegmentBoundKind
  readonly hiKind: CfScaleSegmentBoundKind
  /** Every threshold of the scale, ascending. */
  readonly thresholds: readonly number[]
}

const segmentBoundKind = (p: CfScalePoint): CfScaleSegmentBoundKind => {
  const b = p.bound
  if (b.type === 'min' || b.type === 'max') return b.type
  if (b.type === 'formula') return 'num:formula'
  return ('resolved' in b && b.resolved === 'formula' ? `${b.type}:formula` : b.type) as CfScaleSegmentBoundKind
}

let segmentProbe: ((segment: CfScaleSegment) => CfOutcome<number> | undefined) | null = null
/**
 * Test seam (the num-segment follow-up, `scripts/cf-refs/num-segment-candidates.ts`): a probe that sees every segment
 * evaluation of `scaleColorAt` and may replace its outcome (undefined keeps production's). Null clears. Production never
 * calls it.
 */
export function injectCfScaleSegmentProbe(probe: ((segment: CfScaleSegment) => CfOutcome<number> | undefined) | null): void {
  segmentProbe = probe
}

/** A bound the service resolved from a formula (`num:formula`, `percent:formula`, ...), or an unresolved formula bound. */
const formulaResolvedBound = (p: CfScalePoint): boolean => p.bound.type === 'formula' || ('resolved' in p.bound && p.bound.resolved === 'formula')
const literalNumBound = (p: CfScalePoint): boolean => p.bound.type === 'num' && !('resolved' in p.bound && p.bound.resolved !== undefined)

/**
 * The num-segment follow-up, owner ruling (A) after capture stage2-scale-num-segment: whether a 3-colour cell lies in the
 * withheld UPPER-segment shape. The cell's segment is the upper one (between the second and third thresholds after the
 * midpoint-above-max re-sort), whatever the bound kinds (literal num, min/max/percent/percentile, mixed or formula-resolved);
 * the value is strictly inside it; and the step product is a whole number: for some RGB channel whose endpoints differ by
 * d != 0, `fraction * d` (production's `safeFraction` double times the integer d, one IEEE multiplication) lies within
 * CF_TRUNCATION_GUARD of a whole number k with 0 < k < d. Excel 16.115 drew 19 such cells of that capture (and T4/V3 of
 * e448d831) one unit smaller per step than truncation, by a rule the pre-registered readings did not explain; no other
 * cell moved. The window (owner ruling after capture 54fc3388, stage2-scale-upper-whole) replaced `Number.isInteger`:
 * F12, F15, F16, F18 and F19 of that capture moved with exact whole products whose doubles are not integers (F16:
 * 34.99999999999999, which the guard truncates to 35; Excel drew 34). It also takes stage2-scale-thresholds E4 and its
 * replica D19 (15.999999999999996), which Excel drew at truncation: fail-closed.
 */
export function upperSegmentWholeProduct(c1: number, c2: number, value: number, tLo: number, tHi: number, fraction: number): boolean {
  return upperSegmentProductClass(c1, c2, value, tLo, tHi, fraction) !== undefined
}

/**
 * The upper-segment window follow-up (brief cf-upperwindow): which part of `upperSegmentWholeProduct`'s shape a cell is in.
 * `integer`: some RGB channel's product `fraction * d` is an IEEE integer k, 0 < k < d (the shape captures 377564e0 and
 * 54fc3388 measured, withheld whatever a window reading says). `window`: no channel is an IEEE integer, but some channel's
 * product lies within CF_TRUNCATION_GUARD of a whole k, 0 < k < d (the widening of PR #99). Undefined: neither, or the value
 * is not strictly inside the segment.
 */
export function upperSegmentProductClass(c1: number, c2: number, value: number, tLo: number, tHi: number, fraction: number): 'integer' | 'window' | undefined {
  if (!(value > tLo && value < tHi)) return undefined
  let window = false
  for (const shift of [16, 8, 0]) {
    const d = Math.abs(((c2 >>> shift) & 0xff) - ((c1 >>> shift) & 0xff))
    if (d === 0) continue
    const p = fraction * d, k = Math.round(p)
    if (!(k > 0 && k < d)) continue
    if (Number.isInteger(p)) return 'integer'
    if (Math.abs(p - k) <= CF_TRUNCATION_GUARD) window = true
  }
  return window ? 'window' : undefined
}

/**
 * The upper-segment window follow-up, owner ruling after capture ed0fccb6 (stage2-scale-upper-window): whether an interior
 * cell (tLo < value < tHi) of a 2-colour scale or of a 3-colour LOWER segment is withheld because production's truncation
 * guard changes its step. For some ARGB channel whose endpoints differ by d != 0, p = fraction · |d| (production's
 * `safeFraction` double times the integer |d|, the one IEEE multiplication `interpolateChannel` performs) has
 * floor(p + CF_TRUNCATION_GUARD) != floor(p): p lies within 1e-11 strictly below a whole number. Excel 16.115 drew every
 * captured cell of that shape one step smaller (raw floor: A3, A4, D2, D4, E7..E10, E12..E14 and the twins H8, H18, I5),
 * where production's guarded truncation paints one step larger; no other lower-segment or 2-colour cell moved. Raw floor
 * was never pre-registered, so the shape is withheld (`cf-unmeasured-guard-boundary`), never repainted.
 */
export function guardBoundaryProduct(c0: number, c1: number, value: number, tLo: number, tHi: number, fraction: number): boolean {
  if (!(value > tLo && value < tHi)) return false
  const f = Math.max(0, Math.min(1, fraction))
  for (const shift of [24, 16, 8, 0]) {
    const d = Math.abs(((c1 >>> shift) & 0xff) - ((c0 >>> shift) & 0xff))
    if (d === 0) continue
    const p = f * d
    if (Math.floor(p + CF_TRUNCATION_GUARD) !== Math.floor(p)) return true
  }
  return false
}

/**
 * The raw-floor follow-up (brief cf-rawfloor, Phase 1 task 2): the readings of Excel's step at a cell in the guard-boundary
 * shape (`guardBoundaryProduct`: a 2-colour scale, or a 3-colour LOWER segment not withheld as the literal num-segment shape,
 * where some channel's p = f·|d| lies within 1e-11 strictly below a whole number). Per ARGB channel, the step from the
 * segment's start colour:
 * - `raw-floor` (W1, docs/validation/xlsx-cf.md §14.z.17): floor(f·|d|), no guard; every captured lower-segment and 2-colour
 *   interior cell was drawn so (the 14 of ed0fccb6 included);
 * - `guarded`: floor(f·|d| + 1e-11), production's `interpolateChannel` before the guard-boundary withhold (the control,
 *   contradicted by ed0fccb6 at those 14 cells).
 * Pre-registered by stage2-scale-raw-floor; its capture (Excel 16.115, ce3f701d) matched `raw-floor` at all 67 evidence cells
 * and every control, so `raw-floor` is registered (docs/validation/xlsx-cf.md §14.z.18).
 */
export const CF_GUARD_BOUNDARY_READINGS = ['raw-floor', 'guarded'] as const
export type CfGuardBoundaryReading = (typeof CF_GUARD_BOUNDARY_READINGS)[number]
/** The capability a guard-boundary reading is registered under (stage2-scale-raw-floor). */
export const CF_GUARD_BOUNDARY_CAP = 'scale:guard-boundary-raw-floor'
/**
 * The reading production implements per id: `raw-floor` for scale:guard-boundary-raw-floor, matched by the capture of
 * stage2-scale-raw-floor (ce3f701d). With its evidence, production paints the guard-boundary shape at raw floor where the
 * default sRGB / truncate variants hold and neither bounding cfvo is formula-resolved; a formula-bounded segment, and a
 * 3-colour lower segment bounded by two literal num cfvo (the num-segment withhold, first, unless a literal-lower reading
 * hands the cell here), stay withheld. Without the
 * evidence the shape stays withheld (`cf-unmeasured-guard-boundary`). The audit's CF_STAGE2_PRODUCTION_READINGS spreads this
 * table.
 */
export const CF_GUARD_BOUNDARY_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfGuardBoundaryReading>>> = Object.freeze({ [CF_GUARD_BOUNDARY_CAP]: 'raw-floor' })
let proposedGuardBoundaryReadings: Readonly<Partial<Record<string, CfGuardBoundaryReading>>> | null = null
/** Test seam: rehearse a proposed registration (id to reading) as if it were in the table above. Null clears. Production never calls it. */
export function injectCfGuardBoundaryReadings(readings: Readonly<Partial<Record<string, CfGuardBoundaryReading>>> | null): void {
  proposedGuardBoundaryReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The guard-boundary reading of one id (undefined: unregistered, or not a reading), a proposed one (the test seam) first. */
export function cfGuardBoundaryReading(capability: string = CF_GUARD_BOUNDARY_CAP): CfGuardBoundaryReading | undefined {
  const r = proposedGuardBoundaryReadings && Object.prototype.hasOwnProperty.call(proposedGuardBoundaryReadings, capability) ? proposedGuardBoundaryReadings[capability] : CF_GUARD_BOUNDARY_PRODUCTION_READINGS[capability]
  return r !== undefined && (CF_GUARD_BOUNDARY_READINGS as readonly string[]).includes(r) ? r : undefined
}

/**
 * The outcome of a 2-colour scale or 3-colour lower-segment interpolation. In the guard-boundary shape (truncation only) the
 * cell is withheld, unless a guard-boundary reading is registered with its evidence and the default sRGB / truncate variants
 * hold and neither bounding cfvo is formula-resolved: then it paints the reading (`upperWindowReadingColor`, per ARGB channel
 * from `c0`; undefined on any non-finite intermediate, which withholds).
 */
function guardedInterpolation(c0: number, c1: number, value: number, lo: number, hi: number, fraction: number, colorSpace: CfColorSpaceVariant, rounding: CfRoundingVariant, checkEvidence: (cap: string) => boolean, bounds: readonly CfScalePoint[]): CfOutcome<number> {
  if (rounding === 'truncate' && guardBoundaryProduct(c0, c1, value, lo, hi, fraction)) {
    // A formula-resolved bound was never captured in this shape (Stage 2 holds none; no Stage 5 cell is in it): withheld.
    const r = colorSpace === 'sRGB' && !bounds.some(formulaResolvedBound) ? cfGuardBoundaryReading(CF_GUARD_BOUNDARY_CAP) : undefined
    const argb = r !== undefined && checkEvidence(CF_GUARD_BOUNDARY_CAP) ? upperWindowReadingColor(r, c0, c1, fraction) : undefined
    return argb !== undefined ? { kind: 'ready', value: argb } : { kind: 'unresolved', reason: 'cf-unmeasured-guard-boundary' }
  }
  return { kind: 'ready', value: interpolateScaleColor(c0, c1, fraction, colorSpace, rounding) }
}

/**
 * The upper-segment window follow-up (brief cf-upperwindow, Phase 1 task 2): the readings of Excel's step at a cell in the
 * `window` class of a 3-colour upper segment that is NOT bounded by two literal num cfvo (min, max, percent, percentile, mixed
 * or formula-resolved bounds; the literal shape is withheld as the num-segment shape first). Per RGB channel, the step from
 * the segment's start colour, |d| the channel's span:
 * - `guarded`: floor(f·|d| + 1e-11), production's `interpolateChannel` (what PR #99 withheld; E4 and D19 were drawn so);
 * - `raw-floor`: floor(f·|d|), no guard (one smaller where the product is just below an integer; F12, F15, F16, F18 of the
 *   literal capture 54fc3388 were drawn so).
 * The candidate search (scripts/cf-refs/upper-window-candidates.ts) left no survivor, so these two are the only readings.
 */
export const CF_UPPER_WINDOW_READINGS = ['guarded', 'raw-floor'] as const
export type CfUpperWindowReading = (typeof CF_UPPER_WINDOW_READINGS)[number]
/** The capability a window reading is registered under (stage2-scale-upper-window). */
export const CF_UPPER_WINDOW_CAP = 'scale:upper-window-statistical'
/**
 * The reading production implements per id: none. Until a capture matches a pre-registered reading and it is registered here
 * (with its evidence), every window-class cell stays withheld (`cf-unmeasured-upper-segment-whole-product`), as on main. The
 * audit's CF_STAGE2_PRODUCTION_READINGS spreads this table.
 */
export const CF_UPPER_WINDOW_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfUpperWindowReading>>> = Object.freeze({})
let proposedUpperWindowReadings: Readonly<Partial<Record<string, CfUpperWindowReading>>> | null = null
/** Test seam: rehearse a proposed registration (id to reading) as if it were in the table above. Null clears. Production never calls it. */
export function injectCfUpperWindowReadings(readings: Readonly<Partial<Record<string, CfUpperWindowReading>>> | null): void {
  proposedUpperWindowReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The window reading of one id (undefined: unregistered, or not a reading), a proposed one (the test seam) first. */
export function cfUpperWindowReading(capability: string = CF_UPPER_WINDOW_CAP): CfUpperWindowReading | undefined {
  const r = proposedUpperWindowReadings && Object.prototype.hasOwnProperty.call(proposedUpperWindowReadings, capability) ? proposedUpperWindowReadings[capability] : CF_UPPER_WINDOW_PRODUCTION_READINGS[capability]
  return r !== undefined && (CF_UPPER_WINDOW_READINGS as readonly string[]).includes(r) ? r : undefined
}

/**
 * One window reading's colour between `argb0` (at f = 0) and `argb1`, per channel from the start colour, alpha included (f
 * clamped to [0, 1]). Undefined (the caller withholds) on any non-finite intermediate.
 */
export function upperWindowReadingColor(reading: CfUpperWindowReading | CfGuardBoundaryReading, argb0: number, argb1: number, fraction: number): number | undefined {
  if (!Number.isFinite(fraction)) return undefined
  const f = Math.max(0, Math.min(1, fraction))
  let out = 0
  for (const shift of [24, 16, 8, 0]) {
    const c0 = (argb0 >>> shift) & 0xff, c1 = (argb1 >>> shift) & 0xff
    const p = f * Math.abs(c1 - c0)
    if (!Number.isFinite(p)) return undefined
    const k = reading === 'guarded' ? Math.floor(p + CF_TRUNCATION_GUARD) : Math.floor(p)
    if (!Number.isFinite(k)) return undefined
    out |= (c1 >= c0 ? c0 + k : c0 - k) << shift
  }
  return out >>> 0
}

/**
 * The literal-lower follow-up (brief cf-literal-lower, Phase 1 task 2): the readings of an interior cell of a 3-colour LOWER
 * segment in the num-segment shape (`numSegmentUnmeasured`: bounded by two literal num cfvo, f·2^20 not an integer). Every
 * captured such cell (64, the census scripts/cf-refs/literal-lower-census.ts) was drawn at production's lower-segment
 * arithmetic; the founding anomaly T4/V3 (e448d831) is in the UPPER segment. Readings:
 * - `lower-arithmetic`: the cell is painted exactly as a lower segment not bounded by two literal num cfvo is today
 *   (`guardedInterpolation`: truncation, the guard-boundary shape at its registered reading with its evidence, else withheld);
 * - `step-plus-one`: T4/V3 transplanted: per ARGB channel with d != 0, the step from the segment's start colour is production's
 *   guarded truncation step floor(f·|d| + 1e-11) minus one (`literalLowerStepColour`). It reproduces T4 and V3 and the literal
 *   upper-segment cells F12, F15, F16, F18, F19 of 54fc3388 (where raw floor minus one does not).
 * Pre-registered by stage2-scale-literal-lower; its capture (a3f85309) matched `lower-arithmetic`, which is registered. The UPPER
 * segment is not a reading's: it stays withheld whatever is registered.
 */
export const CF_LITERAL_LOWER_READINGS = ['lower-arithmetic', 'step-plus-one'] as const
export type CfLiteralLowerReading = (typeof CF_LITERAL_LOWER_READINGS)[number]
/** The capability a literal-lower reading is registered under (stage2-scale-literal-lower). */
export const CF_LITERAL_LOWER_CAP = 'scale:num-segment-literal-lower'
/**
 * The reading production implements per id: `lower-arithmetic` for scale:num-segment-literal-lower, matched by the capture of
 * stage2-scale-literal-lower (Excel 16.115, a3f85309) at all 46 decisive cells of its eight groups. With its evidence,
 * production paints the literal-num LOWER segment as a non-literal lower segment (sRGB / truncate); without it the shape stays
 * withheld (`cf-unmeasured-num-segment-fraction`). The upper segment stays withheld whatever is registered. The audit's
 * CF_STAGE2_PRODUCTION_READINGS spreads this table.
 */
export const CF_LITERAL_LOWER_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfLiteralLowerReading>>> = Object.freeze({ [CF_LITERAL_LOWER_CAP]: 'lower-arithmetic' })
let proposedLiteralLowerReadings: Readonly<Partial<Record<string, CfLiteralLowerReading>>> | null = null
/** Test seam: rehearse a proposed registration (id to reading) as if it were in the table above. Null clears. Production never calls it. */
export function injectCfLiteralLowerReadings(readings: Readonly<Partial<Record<string, CfLiteralLowerReading>>> | null): void {
  proposedLiteralLowerReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The literal-lower reading of one id (undefined: unregistered, or not a reading), a proposed one (the test seam) first. */
export function cfLiteralLowerReading(capability: string = CF_LITERAL_LOWER_CAP): CfLiteralLowerReading | undefined {
  const r = proposedLiteralLowerReadings && Object.prototype.hasOwnProperty.call(proposedLiteralLowerReadings, capability) ? proposedLiteralLowerReadings[capability] : CF_LITERAL_LOWER_PRODUCTION_READINGS[capability]
  return r !== undefined && (CF_LITERAL_LOWER_READINGS as readonly string[]).includes(r) ? r : undefined
}

/**
 * `step-plus-one`'s colour between `argb0` (at f = 0) and `argb1`: per ARGB channel with d = |c1 − c0| != 0, the start colour
 * moved toward the end colour by floor(f·d + CF_TRUNCATION_GUARD) − 1 (f clamped to [0, 1]); a channel with d = 0 keeps its
 * value. Undefined (the caller withholds) on any non-finite intermediate, or where some channel's truncation step is 0 (the
 * reading would step past the start colour: T4/V3 say nothing there).
 */
export function literalLowerStepColour(argb0: number, argb1: number, fraction: number): number | undefined {
  if (!Number.isFinite(fraction)) return undefined
  const f = Math.max(0, Math.min(1, fraction))
  let out = 0
  for (const shift of [24, 16, 8, 0]) {
    const c0 = (argb0 >>> shift) & 0xff, c1 = (argb1 >>> shift) & 0xff
    const d = Math.abs(c1 - c0)
    if (d === 0) { out |= c0 << shift; continue }
    const k = Math.floor(f * d + CF_TRUNCATION_GUARD) - 1
    if (!Number.isFinite(k) || k < 0) return undefined
    out |= (c1 >= c0 ? c0 + k : c0 - k) << shift
  }
  return out >>> 0
}

/**
 * The upper-segment follow-up (brief cf-upper-segment, Phase 1b): the readings of an interior cell of a 3-colour UPPER segment
 * that production withholds today: the literal num-segment shape (`numSegmentUnmeasured` in segment 1) and the `integer` and
 * `window` classes of `upperSegmentProductClass`. Readings:
 * - `index-position`: Excel's position on the whole scale as a double, the segment index plus the fraction (t = 1 + f), the
 *   local fraction recovered as l = t − 1 (f rounded to the 2^-52 grid, ties to even), then per ARGB channel the step
 *   floor(l·|d|) from the segment's start colour (`upperSegmentPositionColour`). It is the only survivor class of the search
 *   over all 302 captured upper-segment interior cells (scripts/cf-refs/upper-segment-candidates.ts; with i = 0 it is the
 *   registered lower-segment raw floor);
 * - `truncate`: production's guarded truncation (`interpolateScaleColor`), the control.
 * Pre-registered by stage2-scale-upper-segment. A segment bounded by a formula-resolved cfvo stays withheld under every reading.
 */
export const CF_UPPER_SEGMENT_READINGS = ['index-position', 'truncate'] as const
export type CfUpperSegmentReading = (typeof CF_UPPER_SEGMENT_READINGS)[number]
/** The capability an upper-segment reading is registered under (stage2-scale-upper-segment). */
export const CF_UPPER_SEGMENT_CAP = 'scale:upper-segment'
/**
 * The reading production implements per id: none. Until a capture matches a pre-registered reading and it is registered here
 * (with its evidence), every upper-segment cell withheld today stays withheld, as on main. The audit's
 * CF_STAGE2_PRODUCTION_READINGS spreads this table.
 */
export const CF_UPPER_SEGMENT_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfUpperSegmentReading>>> = Object.freeze({})
let proposedUpperSegmentReadings: Readonly<Partial<Record<string, CfUpperSegmentReading>>> | null = null
/** Test seam: rehearse a proposed registration (id to reading) as if it were in the table above. Null clears. Production never calls it. */
export function injectCfUpperSegmentReadings(readings: Readonly<Partial<Record<string, CfUpperSegmentReading>>> | null): void {
  proposedUpperSegmentReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The upper-segment reading of one id (undefined: unregistered, or not a reading), a proposed one (the test seam) first. */
export function cfUpperSegmentReading(capability: string = CF_UPPER_SEGMENT_CAP): CfUpperSegmentReading | undefined {
  const r = proposedUpperSegmentReadings && Object.prototype.hasOwnProperty.call(proposedUpperSegmentReadings, capability) ? proposedUpperSegmentReadings[capability] : CF_UPPER_SEGMENT_PRODUCTION_READINGS[capability]
  return r !== undefined && (CF_UPPER_SEGMENT_READINGS as readonly string[]).includes(r) ? r : undefined
}

/**
 * `index-position`'s colour between `argb0` (at f = 0) and `argb1`: l = (1 + f) − 1 with f clamped to [0, 1] (two IEEE
 * operations), then per ARGB channel the start colour moved toward the end colour by floor(l·|c1 − c0|). Undefined (the
 * caller withholds) on any non-finite intermediate.
 */
export function upperSegmentPositionColour(argb0: number, argb1: number, fraction: number): number | undefined {
  if (!Number.isFinite(fraction)) return undefined
  const t = 1 + Math.max(0, Math.min(1, fraction))
  const l = t - 1
  if (!Number.isFinite(l)) return undefined
  let out = 0
  for (const shift of [24, 16, 8, 0]) {
    const c0 = (argb0 >>> shift) & 0xff, c1 = (argb1 >>> shift) & 0xff
    const k = Math.floor(l * Math.abs(c1 - c0))
    if (!Number.isFinite(k)) return undefined
    out |= (c1 >= c0 ? c0 + k : c0 - k) << shift
  }
  return out >>> 0
}

/**
 * The near-top sub-shape (orchestrator ruling before the capture of stage2-scale-upper-segment): an interior cell of a 3-colour
 * UPPER segment where some RGB channel with d = |c1 − c0| != 0 has p = f·d within CF_TRUNCATION_GUARD below the full span
 * (d − p <= 1e-11, the value within about 1e-13 of the maximum). Guarded truncation paints d there (the maximum's colour),
 * `index-position` one step less; no capture holds such a cell. Withheld (`cf-unmeasured-upper-near-top`) unless the near-top
 * id is registered with the SAME reading as `scale:upper-segment`, both with their evidence (pre-committed narrowing: the
 * upper-segment reading can be registered with this sub-shape still withheld).
 */
export function upperSegmentNearTop(c0: number, c1: number, value: number, tLo: number, tHi: number, fraction: number): boolean {
  if (!(value > tLo && value < tHi)) return false
  for (const shift of [16, 8, 0]) {
    const d = Math.abs(((c1 >>> shift) & 0xff) - ((c0 >>> shift) & 0xff))
    if (d !== 0 && d - fraction * d <= CF_TRUNCATION_GUARD) return true
  }
  return false
}
/** The capability the near-top sub-shape's reading is registered under (stage2-scale-upper-segment, its own configuration). */
export const CF_UPPER_NEAR_TOP_CAP = 'scale:upper-segment-near-top'
/** The near-top reading production implements: none (withheld) until a capture matches it and it is registered here. */
export const CF_UPPER_NEAR_TOP_PRODUCTION_READINGS: Readonly<Partial<Record<string, CfUpperSegmentReading>>> = Object.freeze({})
let proposedUpperNearTopReadings: Readonly<Partial<Record<string, CfUpperSegmentReading>>> | null = null
/** Test seam: rehearse a proposed near-top registration as if it were in the table above. Null clears. Production never calls it. */
export function injectCfUpperNearTopReadings(readings: Readonly<Partial<Record<string, CfUpperSegmentReading>>> | null): void {
  proposedUpperNearTopReadings = readings ? Object.freeze({ ...readings }) : null
}
/** The near-top reading of one id (undefined: unregistered, or not a reading), a proposed one (the test seam) first. */
export function cfUpperNearTopReading(capability: string = CF_UPPER_NEAR_TOP_CAP): CfUpperSegmentReading | undefined {
  const r = proposedUpperNearTopReadings && Object.prototype.hasOwnProperty.call(proposedUpperNearTopReadings, capability) ? proposedUpperNearTopReadings[capability] : CF_UPPER_NEAR_TOP_PRODUCTION_READINGS[capability]
  return r !== undefined && (CF_UPPER_SEGMENT_READINGS as readonly string[]).includes(r) ? r : undefined
}

/** The colour of the registered upper-segment reading, if any (with its evidence, sRGB / truncate, no formula-resolved bound). */
function upperSegmentReadingColour(lo: CfScalePoint, hi: CfScalePoint, seg: CfScaleSegment, checkEvidence: (cap: string) => boolean, colorSpace: CfColorSpaceVariant, rounding: CfRoundingVariant, nearTop = false): number | undefined {
  if (seg.segment !== 1 || colorSpace !== 'sRGB' || rounding !== 'truncate' || formulaResolvedBound(lo) || formulaResolvedBound(hi)) return undefined
  const r = cfUpperSegmentReading(CF_UPPER_SEGMENT_CAP)
  if (r === undefined || !checkEvidence(CF_UPPER_SEGMENT_CAP)) return undefined
  // The near-top sub-shape: only with its own registration, the same reading, and its evidence.
  if (nearTop && (cfUpperNearTopReading(CF_UPPER_NEAR_TOP_CAP) !== r || !checkEvidence(CF_UPPER_NEAR_TOP_CAP))) return undefined
  return r === 'index-position' ? upperSegmentPositionColour(seg.c0, seg.c1, seg.fraction) : interpolateScaleColor(seg.c0, seg.c1, seg.fraction, colorSpace, rounding)
}

/**
 * A 3-colour segment's outcome: the literal-num shape is withheld where `numSegmentUnmeasured`, except a LOWER-segment cell
 * when a literal-lower reading is registered with its evidence under the default sRGB / truncate variants (`lower-arithmetic`
 * paints it as a non-literal lower segment, `guardedInterpolation`; `step-plus-one` paints `literalLowerStepColour`, else
 * withheld); an upper-segment cell in the `integer` class is withheld; one in the `window` class follows the window reading
 * when it is registered with its evidence, the segment is not bounded by two literal num cfvo and the default sRGB / truncate
 * variants hold, else is withheld; a lower-segment cell in the guard-boundary shape is withheld (`guardBoundaryProduct`)
 * unless its reading is registered; every other cell is production's interpolation. An upper-segment reading registered with
 * its evidence (sRGB / truncate, no formula-resolved bound) paints every upper-segment cell of those withheld shapes first
 * (`upperSegmentReadingColour`).
 */
function threeColourSegment(lo: CfScalePoint, hi: CfScalePoint, seg: CfScaleSegment, checkEvidence: (cap: string) => boolean, colorSpace: CfColorSpaceVariant, rounding: CfRoundingVariant): CfOutcome<number> {
  const literal = literalNumBound(lo) && literalNumBound(hi)
  // The near-top sub-shape (before every other upper-segment rule): withheld unless its own registration matches.
  if (seg.segment === 1 && upperSegmentNearTop(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction)) {
    const upper = upperSegmentReadingColour(lo, hi, seg, checkEvidence, colorSpace, rounding, true)
    return upper !== undefined ? { kind: 'ready', value: upper } : { kind: 'unresolved', reason: 'cf-unmeasured-upper-near-top' }
  }
  if (literal && numSegmentUnmeasured(lo, hi, seg.value, seg.lo, seg.hi, seg.fraction)) {
    const r = seg.segment === 0 && colorSpace === 'sRGB' && rounding === 'truncate' ? cfLiteralLowerReading(CF_LITERAL_LOWER_CAP) : undefined
    if (r !== undefined && checkEvidence(CF_LITERAL_LOWER_CAP)) {
      if (r === 'lower-arithmetic') return guardedInterpolation(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction, colorSpace, rounding, checkEvidence, [lo, hi])
      const argb = literalLowerStepColour(seg.c0, seg.c1, seg.fraction)
      if (argb !== undefined) return { kind: 'ready', value: argb }
    }
    // The upper-segment follow-up: a registered upper-segment reading (with its evidence) paints the upper segment of the shape.
    const upper = upperSegmentReadingColour(lo, hi, seg, checkEvidence, colorSpace, rounding)
    if (upper !== undefined) return { kind: 'ready', value: upper }
    return { kind: 'unresolved', reason: 'cf-unmeasured-num-segment-fraction' }
  }
  const cls = seg.segment === 1 ? upperSegmentProductClass(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction) : undefined
  // The upper-segment follow-up: a registered upper-segment reading (with its evidence) decides the integer and window classes,
  // before the window reading.
  if (cls !== undefined) {
    const upper = upperSegmentReadingColour(lo, hi, seg, checkEvidence, colorSpace, rounding)
    if (upper !== undefined) return { kind: 'ready', value: upper }
  }
  if (cls === 'window' && !literal && colorSpace === 'sRGB' && rounding === 'truncate') {
    const r = cfUpperWindowReading(CF_UPPER_WINDOW_CAP)
    const argb = r !== undefined && checkEvidence(CF_UPPER_WINDOW_CAP) ? upperWindowReadingColor(r, seg.c0, seg.c1, seg.fraction) : undefined
    if (argb !== undefined) return { kind: 'ready', value: argb }
  }
  if (cls !== undefined) return { kind: 'unresolved', reason: 'cf-unmeasured-upper-segment-whole-product' }
  if (seg.segment === 0) return guardedInterpolation(seg.c0, seg.c1, seg.value, seg.lo, seg.hi, seg.fraction, colorSpace, rounding, checkEvidence, [lo, hi])
  return { kind: 'ready', value: interpolateScaleColor(seg.c0, seg.c1, seg.fraction, colorSpace, rounding) }
}

/**
 * Computes the interpolated ARGB colour for a cell value.
 * Pure function: fails closed when evidence is absent or cell value is outside bounds.
 */
export function scaleColorAt(
  cellValue: CfLiteral | number,
  pointsOrRule: readonly CfScalePoint[] | { points: readonly CfScalePoint[] } | CfDefinition,
  thresholds: readonly number[],
  options: {
    hasEvidence?: (cap: string) => boolean
    colorSpace?: CfColorSpaceVariant
    rounding?: CfRoundingVariant
  } = {},
): CfOutcome<number> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  const points = extractPoints(pointsOrRule)

  if (!points || points.length !== thresholds.length || (points.length !== 2 && points.length !== 3)) {
    return { kind: 'unresolved', reason: 'cf-invalid-scale' }
  }

  // Ruling T13-A & Ruling S2-B: computed interpolated fill requires color-space and rounding-ladder
  if (!checkEvidence('scale:interpolation-color-space') || !checkEvidence('scale:interpolation-rounding-ladder')) {
    return { kind: 'unresolved', reason: 'cf-unsupported-interpolation' }
  }

  // Non-numeric value participation (Ruling S2-G)
  let numericVal: number
  if (typeof cellValue === 'number') {
    numericVal = cellValue
  } else if (cellValue.kind === 'number') {
    numericVal = cellValue.value
  } else {
    // string, boolean, error, blank
    if (checkEvidence('scale:participation-types')) {
      return { kind: 'no-match' }
    }
    return { kind: 'unresolved', reason: 'cf-unmeasured-participation' }
  }

  if (!Number.isFinite(numericVal)) {
    return { kind: 'unresolved', reason: 'cf-non-finite-value' }
  }

  const rounding = options.rounding ?? 'truncate'
  const colorSpace = options.colorSpace ?? 'sRGB'

  if (points.length === 2) {
    let t0 = thresholds[0]!
    let t1 = thresholds[1]!
    let c0 = points[0]!.argb
    let c1 = points[1]!.argb

    if (t0 === t1) {
      // Degenerate scope v_min = v_max or t_low = t_high: upper endpoint colour
      return { kind: 'ready', value: c1 }
    }

    if (t0 > t1) {
      // Inverted thresholds (e.g. num 100 -> num 10):
      // Excel sorts thresholds to 10..100 while keeping original endpoint colors c0, c1
      const tLow = t1
      const tHigh = t0
      if (numericVal < tLow || numericVal > tHigh) {
        if (!checkEvidence('scale:degenerate-values-beyond-bounds')) {
          return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }
        }
      }
      const fraction = safeFraction(numericVal, tLow, tHigh)
      if (fraction === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
      const probed = segmentProbe?.({ value: numericVal, lo: tLow, hi: tHigh, fraction, c0, c1, colours: 2, segment: 0, loKind: segmentBoundKind(points[1]!), hiKind: segmentBoundKind(points[0]!), thresholds: [tLow, tHigh] })
      if (probed) return probed
      return guardedInterpolation(c0, c1, numericVal, tLow, tHigh, fraction, colorSpace, rounding, checkEvidence, points)
    }

    // Normal t0 < t1
    if (numericVal < t0 || numericVal > t1) {
      if (!checkEvidence('scale:degenerate-values-beyond-bounds')) {
        return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }
      }
    }

    const fraction = safeFraction(numericVal, t0, t1)
      if (fraction === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
    const probed = segmentProbe?.({ value: numericVal, lo: t0, hi: t1, fraction, c0, c1, colours: 2, segment: 0, loKind: segmentBoundKind(points[0]!), hiKind: segmentBoundKind(points[1]!), thresholds: [t0, t1] })
    if (probed) return probed
    return guardedInterpolation(c0, c1, numericVal, t0, t1, fraction, colorSpace, rounding, checkEvidence, points)
  } else {
    // 3-point scale
    let t0 = thresholds[0]!
    let t1 = thresholds[1]!
    let t2 = thresholds[2]!
    let c0 = points[0]!.argb
    let c1 = points[1]!.argb
    let c2 = points[2]!.argb

    // Ruling F-2 (I3-A): defence in depth, mirrors resolveScaleThresholds: only the measured orderings evaluate.
    if (!((t0 < t1 && t1 < t2) || (t0 < t2 && t2 < t1))) {
      return { kind: 'unresolved', reason: 'cf-inverted-thresholds' }
    }

    // The cfvo whose thresholds bound each segment (re-sorted with the thresholds below).
    let mid = 1, top = 2
    if (t1 > t2) {
      // Midpoint above max (e.g. 10, 200, 100):
      // Excel sorts thresholds to 10, 100, 200 while associating c0 with 10, c1 with 100, c2 with 200
      t1 = thresholds[2]!
      t2 = thresholds[1]!
      mid = 2; top = 1
    }

    const tMin = Math.min(t0, t1, t2)
    const tMax = Math.max(t0, t1, t2)
    if (numericVal < tMin || numericVal > tMax) {
      if (!checkEvidence('scale:degenerate-values-beyond-bounds')) {
        return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }
      }
    }

    if (numericVal <= t1) {
      if (t0 === t1) return { kind: 'ready', value: c1 }
      const fraction = safeFraction(numericVal, t0, t1)
      if (fraction === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
      const seg: CfScaleSegment = { value: numericVal, lo: t0, hi: t1, fraction, c0, c1, colours: 3, segment: 0, loKind: segmentBoundKind(points[0]!), hiKind: segmentBoundKind(points[mid]!), thresholds: [t0, t1, t2] }
      const probed = segmentProbe?.(seg)
      if (probed) return probed
      return threeColourSegment(points[0]!, points[mid]!, seg, checkEvidence, colorSpace, rounding)
    } else {
      if (t1 === t2) return { kind: 'ready', value: c2 }
      const fraction = safeFraction(numericVal, t1, t2)
      if (fraction === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
      const seg: CfScaleSegment = { value: numericVal, lo: t1, hi: t2, fraction, c0: c1, c1: c2, colours: 3, segment: 1, loKind: segmentBoundKind(points[mid]!), hiKind: segmentBoundKind(points[top]!), thresholds: [t0, t1, t2] }
      const probed = segmentProbe?.(seg)
      if (probed) return probed
      return threeColourSegment(points[mid]!, points[top]!, seg, checkEvidence, colorSpace, rounding)
    }
  }
}
