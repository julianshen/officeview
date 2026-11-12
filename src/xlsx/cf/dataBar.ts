/**
 * Data-bar threshold resolution and evaluation (Stage 3, Task 20).
 *
 * - Thresholds reuse `resolveScaleThresholds` (percentile order statistics, degenerate-scope checks) behind bar:*
 *   evidence, so no bar gate is ever lifted by scale:* evidence; only `autoMin`/`autoMax` are bar-specific.
 * - Lengths follow a closed layout-variant set (Ruling S3-B). With no variant enabled every bar is unresolved, and every
 *   geometry branch is a hypothesis until its `bar:*` capability verifies.
 * - Colours come only from the normalised definition (Ruling S3-F as amended); an absent colour the cell would draw
 *   withholds the bar.
 * - Overflow-safe (Lesson 6): spans that overflow are normalised in halves through `safeFraction`.
 *
 * The evaluator is pure: no DOM, no clock, no formula AST, no global state beyond the evidence lookup.
 */
import type { CfBarBound, CfBarDescriptor, CfDefinition, CfLiteral, CfOutcome, CfScaleBound } from './model'
import type { CfRangeStatistics } from './statistics'
import { resolveScaleThresholds, safeFraction } from './colorScale'
import { hasEvidence as defaultHasEvidence } from './evidence'

type CfDataBarDefinition = Extract<CfDefinition, { kind: 'dataBar' }>

/** Ruling S3-B: where a bar starts (the cell's leading edge, or the axis) and what its min/max lengths are a share of
 * (the whole cell, or the axis side the bar grows into). Exactly one variant is enabled by evidence; none by default. */
export type CfBarLayoutVariant = Readonly<{ origin: 'span' | 'axis'; lengthBasis: 'cell' | 'side' }>

/** Never verified: a length pair Excel was not measured with (only 0/100 and the 10/90 default were captured). */
const BAR_LENGTHS_UNMEASURED = 'bar:lengths-unmeasured'
/** Never verified throughout Stage 3 (S3-K): a context bar on a sheet that is RTL or of unknown order. */
const BAR_CONTEXT_RTL_SHEET = 'bar:direction-context-rtl-sheet'

const unmeasured = (): CfOutcome<never> => ({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })

function boundCapability(bound: CfBarBound): string {
  switch (bound.type) {
    case 'min': case 'max': return 'bar:bounds-min-max'
    case 'num': return 'bar:bounds-num'
    case 'percent': return 'bar:bounds-percent'
    case 'percentile': return 'bar:bounds-percentile'
    case 'autoMin': return 'bar:bounds-automin'
    case 'autoMax': return 'bar:bounds-automax'
  }
}

/**
 * The data-independent evaluation gates of a bar definition: bounds, lengths, fill type, border, a fixed axis
 * setting and direction. Data-dependent gates (the automatic axis on mixed-sign data, negative colours) are checked
 * per cell by `evaluateDataBar`. Never empty.
 */
export function barEvaluationRequiredCapabilities(def: CfDataBarDefinition): Set<string> {
  const req = new Set<string>()
  for (const b of def.bounds) {
    req.add(boundCapability(b))
    if ((b.type === 'percent' || b.type === 'percentile') && (b.value < 0 || b.value > 100)) req.add('bar:bounds-out-of-domain')
  }
  req.add(def.minLength === 0 && def.maxLength === 100 ? 'bar:lengths-explicit-0-100'
    : def.minLength === 10 && def.maxLength === 90 ? 'bar:lengths-default-10-90' : BAR_LENGTHS_UNMEASURED)
  req.add(def.gradient ? 'bar:render-gradient-stops' : 'bar:render-solid')
  if (def.border) req.add('bar:render-border')
  if (def.axisPosition === 'middle') req.add('bar:axis-middle')
  else if (def.axisPosition === 'none') req.add('bar:axis-none')
  if (def.direction === 'leftToRight') req.add('bar:direction-ltr')
  else if (def.direction === 'rightToLeft') req.add('bar:direction-rtl')
  else req.add(def.sheetRightToLeft === false ? 'bar:direction-context-ltr-sheet' : BAR_CONTEXT_RTL_SHEET)
  return req
}

/** The in-domain threshold checks `resolveScaleThresholds` performs; their bar:* gates are checked before delegating.
 * Every degenerate and out-of-domain check is denied, so those shapes stay unresolved (Ruling S3-E). */
const DELEGATED_THRESHOLD_CHECKS: ReadonlySet<string> = new Set([
  'scale:thresholds-min-max', 'scale:thresholds-num-bounds', 'scale:thresholds-percent', 'scale:thresholds-percentile-inc-exc',
])

/**
 * Resolves `[low, high]` for a bar. `autoMin`/`autoMax` are bar-specific (hypothesis: `min(0, v_min)` and
 * `max(0, v_max)`, gated by `bar:bounds-automin`/`-automax`); every other bound is resolved by the shared
 * `resolveScaleThresholds`, including degenerate scopes (one numeric cell, `v_min = v_max`, inverted bounds), which are
 * unresolved until `stage3-bar-degenerate` is implemented as measured.
 */
export function resolveBarThresholds(
  def: CfDataBarDefinition,
  stats: CfRangeStatistics,
  options: { hasEvidence?: (cap: string) => boolean } = {},
): CfOutcome<readonly [number, number]> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  for (const b of def.bounds) if (!checkEvidence(boundCapability(b))) return unmeasured()
  const delegated: { bound: CfScaleBound }[] = []
  for (const b of def.bounds) {
    if (b.type === 'autoMin' || b.type === 'autoMax') {
      const extreme = b.type === 'autoMin' ? stats.min : stats.max
      if (extreme === undefined || !Number.isFinite(extreme)) return { kind: 'unresolved', reason: 'cf-no-numeric-values' }
      // Substituted as a literal; its own gate passed above.
      delegated.push({ bound: { type: 'num', value: b.type === 'autoMin' ? Math.min(0, extreme) : Math.max(0, extreme) } })
    } else {
      delegated.push({ bound: b })
    }
  }
  const out = resolveScaleThresholds(delegated, stats, { hasEvidence: cap => DELEGATED_THRESHOLD_CHECKS.has(cap) })
  if (out.kind !== 'ready') return out
  const [lo, hi] = out.value as readonly [number, number]
  // resolveScaleThresholds already refuses inverted and equal thresholds under the denied degenerate gates.
  return lo < hi ? { kind: 'ready', value: [lo, hi] } : { kind: 'unresolved', reason: 'cf-degenerate-scope' }
}

/** `|v| / (hi - lo)` without overflowing the span (Lesson 6). */
function shareOfSpan(magnitude: number, lo: number, hi: number): number {
  const span = hi - lo
  return Number.isFinite(span) ? magnitude / span : (magnitude / 2) / (hi / 2 - lo / 2)
}

/**
 * Evaluates one cell's bar against resolved thresholds. Fractions are of the cell's inner width in reading order;
 * `startFraction <= endFraction`. Every branch is unresolved unless its evidence is present and `variant` is given
 * (Ruling S3-B); `stats` carries the scope's min/max for the whole-rule beyond-bounds check (Ruling S3-E).
 */
export function evaluateDataBar(
  value: CfLiteral | number,
  def: CfDataBarDefinition,
  thresholds: readonly [number, number],
  stats: Pick<CfRangeStatistics, 'min' | 'max'>,
  options: { hasEvidence?: (cap: string) => boolean; variant?: CfBarLayoutVariant } = {},
): CfOutcome<Omit<CfBarDescriptor, 'ruleId'>> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  const variant = options.variant
  if (!variant) return unmeasured()
  for (const cap of barEvaluationRequiredCapabilities(def)) if (!checkEvidence(cap)) return unmeasured()
  const direction = def.direction === 'context' ? (def.sheetRightToLeft === false ? 'leftToRight' : undefined) : def.direction
  if (direction === undefined) return unmeasured()

  const [lo, hi] = thresholds
  if (!Number.isFinite(lo) || !Number.isFinite(hi) || !(lo < hi)) return { kind: 'unresolved', reason: 'cf-degenerate-scope' }
  // S3-E: a scope with cells beyond the bounds is unresolved for the whole rule, never clamped.
  if (stats.min === undefined || stats.max === undefined || stats.min < lo || stats.max > hi) {
    return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }
  }

  // S3-G: only numbers draw; other kinds draw no bar only with evidence, else they are unresolved.
  const v = typeof value === 'number' ? value : value.kind === 'number' ? value.value : undefined
  if (v === undefined) return checkEvidence('bar:participation-types') ? { kind: 'no-match' } : { kind: 'unresolved', reason: 'cf-unmeasured-participation' }
  if (!Number.isFinite(v)) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
  if (v < lo || v > hi) return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }

  const isNegative = v < 0
  const l0 = def.minLength / 100, l1 = def.maxLength / 100
  const length = (f: number): number => l0 + f * (l1 - l0)

  // Axis: `middle` always draws one; `automatic` only for mixed-sign thresholds; `none` never.
  let axis: number | undefined
  let sideExtent = 0, fullShare = 0
  if (def.axisPosition === 'middle') {
    axis = 0.5
    const ext = Math.max(Math.abs(lo), Math.abs(hi))
    sideExtent = ext
    fullShare = shareOfSpan(Math.abs(v), -ext, ext)
  } else if (def.axisPosition === 'automatic' && lo < 0) {
    // All-negative data has no layout hypothesis yet (bar:lengths-all-negative): withheld.
    if (hi <= 0) return unmeasured()
    if (!checkEvidence('bar:axis-automatic')) return unmeasured()
    // Spec formula -v_min / (v_max - v_min), in halves when the span overflows.
    axis = safeFraction(0, lo, hi)
    sideExtent = isNegative ? -lo : hi
    fullShare = shareOfSpan(Math.abs(v), lo, hi)
  }

  let startFraction: number, endFraction: number, axisFraction = 0
  if (axis === undefined) {
    const f = safeFraction(v, lo, hi)
    if (f === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
    startFraction = 0
    endFraction = length(f)
  } else {
    // The span variant defines no axis layout.
    if (variant.origin !== 'axis' || def.axisArgb === undefined) return unmeasured()
    const sideWidth = isNegative ? axis : 1 - axis
    const extent = variant.lengthBasis === 'side' ? length(Math.abs(v) / sideExtent) * sideWidth : Math.min(length(fullShare), sideWidth)
    axisFraction = axis
    startFraction = isNegative ? axis - extent : axis
    endFraction = isNegative ? axis : axis + extent
  }

  // S3-L negative colours, each under its own evidence; colours never default (S3-F).
  let fillColor: number | undefined = def.fillArgb
  let borderColor: number | undefined = def.border ? def.borderArgb : undefined
  if (isNegative) {
    if (def.negativeBarColorSameAsPositive) {
      if (!checkEvidence('bar:negative-fill-same-as-positive')) return unmeasured()
    } else {
      if (!checkEvidence('bar:negative-fill-color')) return unmeasured()
      fillColor = def.negativeFillArgb
    }
    if (def.border) {
      if (def.negativeBarBorderColorSameAsPositive) {
        if (!checkEvidence('bar:negative-border-same-as-positive')) return unmeasured()
      } else {
        if (!checkEvidence('bar:negative-border-color')) return unmeasured()
        borderColor = def.negativeBorderArgb
      }
    }
  }
  if (fillColor === undefined || (def.border && borderColor === undefined)) return unmeasured()

  for (const n of [axisFraction, startFraction, endFraction]) {
    if (!Number.isFinite(n) || n < 0 || n > 1) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
  }
  return {
    kind: 'ready',
    value: Object.freeze({
      axisFraction, startFraction, endFraction, fillColor,
      ...(borderColor !== undefined ? { borderColor } : {}),
      isGradient: def.gradient, isNegative,
      ...(axis !== undefined ? { axisColor: def.axisArgb! } : {}),
      showValue: def.showValue, direction,
    }),
  }
}
