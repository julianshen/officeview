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
import { barAutoBoundEvidence, barDefaultLengthsRule, barLayoutEvidence, hasEvidence as defaultHasEvidence, type CfBarDefaultLengths, type CfBarLayout } from './evidence'

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
  // Ruling S3-GRAD-BORDER: a bordered gradient (Excel's UI preset) is its own measured shape.
  if (def.border && def.gradient) req.add('bar:render-gradient-border')
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
 * Resolves `[low, high]` for a bar. `autoMin`/`autoMax` are bar-specific: each resolves to the rule its capability
 * verified (`barAutoBoundEvidence`: `min(0, v_min)`/`max(0, v_max)` or `v_min`/`v_max`), else it is unresolved; every other bound is resolved by the shared
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
      // Ruling S3-AUTO-BOUNDS (Task 22c): the rule the oracle verified for this bound, never a default.
      const rule = barAutoBoundEvidence(b.type, checkEvidence)
      if (rule === undefined) return unmeasured()
      const value = rule === 'min-max' ? extreme : b.type === 'autoMin' ? Math.min(0, extreme) : Math.max(0, extreme)
      // Substituted as a literal; its own gate passed above.
      delegated.push({ bound: { type: 'num', value } })
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
 * `startFraction <= endFraction`. Every branch is unresolved unless its evidence is present and a layout variant is
 * enabled for the scope's sign (Ruling S3-B): `variant` (one variant for both signs) or `layout`, else the published
 * selection through `barLayoutEvidence`. `stats` carries the scope's min/max for the whole-rule beyond-bounds check
 * (Ruling S3-E).
 */
export function evaluateDataBar(
  value: CfLiteral | number,
  def: CfDataBarDefinition,
  thresholds: readonly [number, number],
  stats: Pick<CfRangeStatistics, 'min' | 'max'>,
  options: { hasEvidence?: (cap: string) => boolean; variant?: CfBarLayoutVariant; layout?: CfBarLayout; lengthsRule?: CfBarDefaultLengths } = {},
): CfOutcome<Omit<CfBarDescriptor, 'ruleId'>> {
  const checkEvidence = options.hasEvidence ?? defaultHasEvidence
  const layout: CfBarLayout | undefined = options.variant ? { allPositive: options.variant, mixedSign: options.variant }
    : 'layout' in options ? options.layout : barLayoutEvidence(checkEvidence)
  if (!layout || (!layout.allPositive && !layout.mixedSign)) return unmeasured()
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
  // S3-G: a number displayed as a date evaluates as a number only with evidence.
  if (typeof value !== 'number' && value.kind === 'number' && value.dateSerial && !checkEvidence('bar:participation-dateserial')) {
    return { kind: 'unresolved', reason: 'cf-unmeasured-participation' }
  }
  if (!Number.isFinite(v)) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
  if (v < lo || v > hi) return { kind: 'unresolved', reason: 'cf-value-beyond-bounds' }

  const isNegative = v < 0
  const fullLengths = def.minLength === 0 && def.maxLength === 100
  // Task 22b: lengths other than 0/100 follow the rule the oracle selected (the plain linear reading is refuted for
  // Excel's paired solid bars). With no selected rule, or one this evaluator cannot apply, the bar is unresolved.
  const lengthStart = fullLengths ? 0 : defaultLengthsStart(def, 'lengthsRule' in options ? options.lengthsRule : barDefaultLengthsRule(checkEvidence))
  if (lengthStart === undefined) return unmeasured()
  const l0 = lengthStart
  const span = (def.maxLength - def.minLength) / 100
  const length = (f: number): number => l0 + f * span

  // All-negative thresholds under `automatic` (bar:lengths-all-negative, the oracle's `axis` hypothesis): the bar grows
  // from the trailing edge, longest at v_min, as a share of the whole span. Measured at 0/100 lengths only; the capture
  // records no axis line there, so the descriptor carries none.
  if (def.axisPosition === 'automatic' && lo < 0 && hi <= 0) {
    if (!checkEvidence('bar:lengths-all-negative') || !fullLengths) return unmeasured()
    const f = safeFraction(v, lo, hi)
    if (f === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
    return finish(def, checkEvidence, { axisFraction: 1, startFraction: f, endFraction: 1, isNegative: v < 0, axisColor: undefined, direction })
  }

  // Axis: `middle` always draws one; `automatic` only for mixed-sign thresholds; `none` never.
  let axis: number | undefined
  let sideExtent = 0, fullShare = 0
  if (def.axisPosition === 'middle') {
    axis = 0.5
    const ext = Math.max(Math.abs(lo), Math.abs(hi))
    sideExtent = ext
    fullShare = shareOfSpan(Math.abs(v), -ext, ext)
  } else if (def.axisPosition === 'automatic' && lo < 0) {
    if (!checkEvidence('bar:axis-automatic')) return unmeasured()
    // Spec formula -v_min / (v_max - v_min), in halves when the span overflows.
    axis = safeFraction(0, lo, hi)
    sideExtent = isNegative ? -lo : hi
    fullShare = shareOfSpan(Math.abs(v), lo, hi)
  }

  let startFraction: number, endFraction: number, axisFraction = 0
  if (axis === undefined) {
    // No axis drawn: the all-positive layout (every variant starts at the leading edge and spans the whole cell).
    if (!layout.allPositive) return unmeasured()
    const f = safeFraction(v, lo, hi)
    if (f === undefined) return { kind: 'unresolved', reason: 'cf-non-finite-value' }
    startFraction = 0
    endFraction = length(f)
  } else {
    // The span variant defines no axis layout, and axes were measured at 0/100 lengths only.
    const variant = layout.mixedSign
    if (!variant || variant.origin !== 'axis' || !fullLengths || def.axisArgb === undefined) return unmeasured()
    const sideWidth = isNegative ? axis : 1 - axis
    const extent = variant.lengthBasis === 'side' ? length(Math.abs(v) / sideExtent) * sideWidth : Math.min(length(fullShare), sideWidth)
    axisFraction = axis
    startFraction = isNegative ? axis - extent : axis
    endFraction = isNegative ? axis : axis + extent
  }

  return finish(def, checkEvidence, { axisFraction, startFraction, endFraction, isNegative, axisColor: axis !== undefined ? def.axisArgb : undefined, direction })
}

/**
 * Where a non-0/100 bar's length starts (as a share of the span) under the selected rule (`BAR_LENGTH_RULES`,
 * scripts/cf-refs/verify.ts): `linear` at minLength, `half-min` at minLength / 2 (the slope is maxLength - minLength
 * either way). A scoped rule applies to its class only (`solid`: gradient="0"; `x14`: the paired x14 form, Ruling
 * S3-LEN-X14) and every other bar keeps the linear reading the same capture verified on it. Undefined (unresolved) with
 * no rule, an `offset-9pt` rule on a bar it covers (a fixed point offset is not a share of the span), or an `x14`-scoped
 * rule on a bar whose pairing form is unknown, or a bar whose pairing form x fill type the rule's capture never measured.
 */
function defaultLengthsStart(def: CfDataBarDefinition, selected: CfBarDefaultLengths | undefined): number | undefined {
  if (selected === undefined) return undefined
  // Task 22c (review I2): only the pairing form x fill type combinations the capture measured.
  if (def.pairing === undefined || !selected.measured.includes(`${def.pairing}/${def.gradient ? 'gradient' : 'solid'}`)) return undefined
  const [form, cls] = selected.rule.split('@')
  const linear = def.minLength / 100
  if (form === 'linear' && cls === undefined) return linear
  if ((form !== 'half-min' && form !== 'offset-9pt') || (cls !== 'all' && cls !== 'solid' && cls !== 'x14')) return undefined
  if (cls === 'x14' && def.pairing === undefined) return undefined
  const applies = cls === 'all' || (cls === 'solid' && !def.gradient) || (cls === 'x14' && def.pairing === 'x14')
  if (!applies) return linear
  return form === 'half-min' ? def.minLength / 200 : undefined
}

/** Colours (S3-L, S3-F) and the finite-range check shared by every layout branch. */
function finish(
  def: CfDataBarDefinition,
  checkEvidence: (cap: string) => boolean,
  g: { axisFraction: number; startFraction: number; endFraction: number; isNegative: boolean; axisColor: number | undefined; direction: 'leftToRight' | 'rightToLeft' },
): CfOutcome<Omit<CfBarDescriptor, 'ruleId'>> {
  const { axisFraction, startFraction, endFraction, isNegative, axisColor, direction } = g
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
      ...(axisColor !== undefined ? { axisColor } : {}),
      showValue: def.showValue, direction,
    }),
  }
}
