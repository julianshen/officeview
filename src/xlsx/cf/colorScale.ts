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
    return v0 + frac * (v1 - v0)
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
    return v0 + frac * (v1 - v0)
  }
}

/**
 * Evaluates channel interpolation using a closed rounding-variant set.
 * 'truncate' mirrors Excel's measured display-format behavior:
 * floor towards target when increasing, floor towards start when decreasing.
 */
export function interpolateChannel(
  c0: number,
  c1: number,
  fraction: number,
  rounding: CfRoundingVariant = 'truncate',
): number {
  const f = Math.max(0, Math.min(1, fraction))
  if (rounding === 'truncate') {
    // 1e-11 guard absorbs floating-point representation error from threshold division
    if (c1 >= c0) {
      return c0 + Math.floor(f * (c1 - c0) + 1e-11)
    } else {
      return c0 - Math.floor(f * (c0 - c1) + 1e-11)
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

function extractPoints(pointsOrRule: readonly CfScalePoint[] | { points: readonly CfScalePoint[] } | CfDefinition): readonly CfScalePoint[] | undefined {
  if (Array.isArray(pointsOrRule)) return pointsOrRule
  if (typeof pointsOrRule === 'object' && pointsOrRule !== null && 'points' in pointsOrRule) {
    return (pointsOrRule as { points: readonly CfScalePoint[] }).points
  }
  return undefined
}

/**
 * Resolves numeric threshold boundaries for a colour-scale rule.
 * Pure function: fails closed when evidence is absent or boundaries are degenerate/unsupported.
 */
export function resolveScaleThresholds(
  pointsOrRule: readonly CfScalePoint[] | { points: readonly CfScalePoint[] } | CfDefinition,
  stats: CfRangeStatistics,
  options: { hasEvidence?: (cap: string) => boolean } = {},
): CfOutcome<readonly number[]> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  const points = extractPoints(pointsOrRule)

  if (!points || (points.length !== 2 && points.length !== 3)) {
    return { kind: 'unresolved', reason: 'cf-invalid-rule' }
  }

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
      const val = stats.min + (bound.value / 100) * (stats.max - stats.min)
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
      const fraction = Math.max(0, Math.min(1, (numericVal - tLow) / (tHigh - tLow)))
      return { kind: 'ready', value: interpolateScaleColor(c0, c1, fraction, colorSpace, rounding) }
    }

    // Normal t0 < t1
    if (numericVal < t0 || numericVal > t1) {
      if (!checkEvidence('scale:degenerate-values-beyond-bounds')) {
        return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }
      }
    }

    const fraction = Math.max(0, Math.min(1, (numericVal - t0) / (t1 - t0)))
    return { kind: 'ready', value: interpolateScaleColor(c0, c1, fraction, colorSpace, rounding) }
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

    if (t1 > t2) {
      // Midpoint above max (e.g. 10, 200, 100):
      // Excel sorts thresholds to 10, 100, 200 while associating c0 with 10, c1 with 100, c2 with 200
      t1 = thresholds[2]!
      t2 = thresholds[1]!
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
      const fraction = Math.max(0, Math.min(1, (numericVal - t0) / (t1 - t0)))
      return { kind: 'ready', value: interpolateScaleColor(c0, c1, fraction, colorSpace, rounding) }
    } else {
      if (t1 === t2) return { kind: 'ready', value: c2 }
      const fraction = Math.max(0, Math.min(1, (numericVal - t1) / (t2 - t1)))
      return { kind: 'ready', value: interpolateScaleColor(c1, c2, fraction, colorSpace, rounding) }
    }
  }
}
