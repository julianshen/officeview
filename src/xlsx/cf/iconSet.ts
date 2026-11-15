/**
 * Icon-set threshold resolution, buckets and evaluation (Stage 4, Task 26).
 *
 * - Thresholds compose the shared primitives only: `lerp` for `percent`, `percentileOrderStatistic` for `percentile`
 *   (over the session's ascending copy), the literal for `num`. This module orders nothing and computes no percentile.
 * - Every reading is a hypothesis (Rulings S4-C/D/E) until its `icon:*` capability is verified. Where a capability names
 *   several readings and nothing pins the viewer's, the reading is taken from `iconSelectionEvidence`, and without a
 *   selection the shape stays unresolved.
 * - Overflow-safe (Lesson 6): finite thresholds near the double limit go through `lerp`; nothing divides.
 *
 * The evaluator is pure: no DOM, no clock, no glyph data, no global state beyond the evidence lookup.
 */
import type { CfDefinition, CfIconDescriptor, CfIconRef, CfIconSetName, CfLiteral, CfOutcome } from './model'
import type { CfRangeStatistics } from './statistics'
import { lerp, percentileOrderStatistic } from './colorScale'
import { hasEvidence as defaultHasEvidence, iconSelectionEvidence } from './evidence'

type CfIconSetDefinition = Extract<CfDefinition, { kind: 'iconSet' }>
type IconOptions = { hasEvidence?: (cap: string) => boolean }

/** The x14-only families (Ruling S4-E: `icon:families-x14`); every other built-in family is base. */
const X14_FAMILIES: ReadonlySet<CfIconSetName> = new Set(['3Triangles', '3Stars', '5Boxes'])
const PERCENTILE = 'icon:thresholds-percentile-inc-exc'

const unmeasured = (): CfOutcome<never> => ({ kind: 'unresolved', reason: 'cf-unmeasured-capability' })
const unresolved = (reason: string): CfOutcome<never> => ({ kind: 'unresolved', reason })

/** The icon count of a family: its leading digit (3, 4 or 5). */
function familySize(set: CfIconSetName): number {
  return Number(set.charAt(0))
}

function typeCapability(type: CfIconSetDefinition['thresholds'][number]['bound']['type']): string {
  switch (type) {
    case 'num': return 'icon:thresholds-num'
    case 'percent': return 'icon:thresholds-percent'
    case 'percentile': return PERCENTILE
    case 'min': case 'max': return 'icon:thresholds-min-max'
  }
}

/**
 * The data-independent evaluation gates of an icon set (S4-D layer split): family order, each threshold type, `gte`
 * false, `reverse` (and its custom reading) and out-of-domain bounds. Data-dependent gates (floor, degenerate data,
 * participation) are checked where the data is known.
 */
export function iconEvaluationRequiredCapabilities(def: CfIconSetDefinition): Set<string> {
  const req = new Set<string>()
  if (!def.custom) req.add(X14_FAMILIES.has(def.set) ? 'icon:families-x14' : 'icon:families-base')
  for (const t of def.thresholds) {
    req.add(typeCapability(t.bound.type))
    if (!t.gte) req.add('icon:thresholds-gte-false')
    if ((t.bound.type === 'percent' || t.bound.type === 'percentile') && (t.bound.value < 0 || t.bound.value > 100)) req.add('icon:bounds-out-of-domain')
  }
  if (def.reverse) {
    req.add('icon:reverse')
    if (def.custom) req.add('icon:reverse-custom')
  }
  return req
}

/**
 * Gates S4-D names that the capability set above leaves implicit: `icon:thresholds-gte-true` for every comparing
 * threshold (cfvo[1..k-1]) left at, or written as, `gte="1"`, and `icon:thresholds-mixed` for a list mixing types.
 */
function implicitCapabilities(def: CfIconSetDefinition): string[] {
  const out: string[] = []
  if (def.thresholds.some((t, i) => i > 0 && t.gte)) out.push('icon:thresholds-gte-true')
  if (new Set(def.thresholds.map(t => t.bound.type)).size > 1) out.push('icon:thresholds-mixed')
  return out
}

function gated(def: CfIconSetDefinition, check: (cap: string) => boolean): boolean {
  for (const cap of iconEvaluationRequiredCapabilities(def)) if (!check(cap)) return false
  for (const cap of implicitCapabilities(def)) if (!check(cap)) return false
  return true
}

/** Whether a degenerate shape's capability selected the as-written reading (`as-model`); any other reading, or none,
 * leaves the shape unresolved (it is not implemented). */
function asModel(capability: string, check: (cap: string) => boolean): boolean {
  return iconSelectionEvidence(capability, check) === 'as-model'
}

/**
 * Resolves t_0..t_{k-1} for an icon set (Ruling S4-D). `stats` may be absent for a `num`-only set, which needs no
 * statistics. Degenerate data (one numeric cell, all values equal), non-monotonic results and out-of-domain bounds are
 * unresolved unless their capability selected the as-written reading.
 */
export function resolveIconThresholds(
  def: CfIconSetDefinition,
  stats: CfRangeStatistics | undefined,
  options: IconOptions = {},
): CfOutcome<readonly number[]> {
  const check = options.hasEvidence ?? defaultHasEvidence
  if (!gated(def, check)) return unmeasured()
  const k = def.custom?.length ?? familySize(def.set)
  if (def.thresholds.length !== k) return unresolved('cf-invalid-rule')
  for (const t of def.thresholds) {
    const b = t.bound
    if ((b.type === 'percent' || b.type === 'percentile') && (b.value < 0 || b.value > 100) && !asModel('icon:bounds-out-of-domain', check)) return unmeasured()
  }
  const statistical = def.thresholds.some(t => t.bound.type !== 'num')
  if (statistical) {
    if (!stats || stats.numericCount === 0 || stats.min === undefined || stats.max === undefined) return unresolved('cf-no-numeric-values')
    if (!Number.isFinite(stats.min) || !Number.isFinite(stats.max)) return unresolved('cf-non-finite-value')
    if (stats.numericCount === 1) {
      if (!asModel('icon:degenerate-single-cell', check)) return unresolved('cf-degenerate-scope')
    } else if (stats.min === stats.max && !asModel('icon:degenerate-equal-values', check)) return unresolved('cf-degenerate-scope')
  }
  let variant: 'inc' | 'exc' | undefined
  if (def.thresholds.some(t => t.bound.type === 'percentile')) {
    const selected = iconSelectionEvidence(PERCENTILE, check)
    if (selected !== 'inc' && selected !== 'exc') return unmeasured()
    variant = selected
  }
  const out: number[] = []
  for (const t of def.thresholds) {
    const b = t.bound
    let value: number | undefined
    if (b.type === 'num') value = b.value
    else if (b.type === 'min') value = stats?.min
    else if (b.type === 'max') value = stats?.max
    else if (b.type === 'percent') value = stats && stats.min !== undefined && stats.max !== undefined ? lerp(stats.min, stats.max, b.value / 100) : undefined
    else value = stats && variant ? percentileOrderStatistic(stats.sortedAsc, b.value, variant) : undefined
    if (value === undefined) return unresolved('cf-no-numeric-values')
    if (!Number.isFinite(value)) return unresolved('cf-non-finite-value')
    out.push(value)
  }
  for (let i = 1; i < out.length; i++) {
    if (out[i]! < out[i - 1]! && !asModel('icon:degenerate-non-monotonic', check)) return unresolved('cf-degenerate-scope')
  }
  return { kind: 'ready', value: Object.freeze(out) }
}

/** Ruling S4-D hypothesis: the largest i in 1..k-1 with v >= t_i (v > t_i when that threshold's gte is false), else 0. */
export function iconBucket(value: number, thresholds: readonly number[], gte: readonly boolean[]): number {
  let bucket = 0
  for (let i = 1; i < thresholds.length; i++) if (gte[i] === false ? value > thresholds[i]! : value >= thresholds[i]!) bucket = i
  return bucket
}

/** The icon a custom bucket shows under `reverse="1"`, by the reading `icon:reverse-custom` selected (S4-E). */
function reversedCustom(custom: readonly CfIconRef[], bucket: number, reading: string | undefined): CfIconRef | undefined {
  const k = custom.length
  if (reading === 'reverse-then-map') return custom[k - 1 - bucket]
  if (reading === 'ignored') return custom[bucket]
  if (reading === 'map-then-mirror') {
    const ref = custom[bucket]
    // The oracle's reading mirrors the id within the custom bucket count; an id it moves outside the family withholds.
    if (!ref || ref.set === 'NoIcons') return ref
    const iconId = k - 1 - ref.iconId
    return iconId >= 0 && iconId < familySize(ref.set) ? { set: ref.set, iconId } : undefined
  }
  return undefined
}

/**
 * Evaluates one cell's icon against resolved thresholds. `bucket` and `buckets` are always the semantic bucket (before
 * any reversal) and k; `(set, iconId)` is what the bucket shows. Text, boolean, error and blank values draw no icon only
 * with `icon:participation-types` (S4-G), a date serial evaluates only with `icon:participation-dateserial`, and a value
 * below a non-trivial floor follows the reading `icon:thresholds-floor` selected (S4-C).
 */
export function evaluateIconSet(
  value: CfLiteral | number,
  def: CfIconSetDefinition,
  thresholds: readonly number[],
  options: IconOptions = {},
): CfOutcome<Omit<CfIconDescriptor, 'ruleId'>> {
  const check = options.hasEvidence ?? defaultHasEvidence
  if (!gated(def, check)) return unmeasured()
  // Ruling S4-M: the viewer lays out no RTL sheet (parse already refuses; kept so the evaluator is closed alone).
  if (def.sheetRightToLeft !== false) return unmeasured()
  const k = def.custom?.length ?? familySize(def.set)
  if (thresholds.length !== k || def.thresholds.length !== k) return unresolved('cf-invalid-rule')
  if (!thresholds.every(Number.isFinite)) return unresolved('cf-non-finite-value')

  const v = typeof value === 'number' ? value : value.kind === 'number' ? value.value : undefined
  if (v === undefined) return check('icon:participation-types') ? { kind: 'no-match' } : unresolved('cf-unmeasured-participation')
  if (typeof value !== 'number' && value.kind === 'number' && value.dateSerial && !check('icon:participation-dateserial')) return unresolved('cf-unmeasured-participation')
  if (!Number.isFinite(v)) return unresolved('cf-non-finite-value')

  const gte = def.thresholds.map(t => t.gte)
  const floorReached = gte[0] ? v >= thresholds[0]! : v > thresholds[0]!
  if (!floorReached) {
    const reading = iconSelectionEvidence('icon:thresholds-floor', check)
    if (reading === 'below-floor-no-icon') return { kind: 'no-match' }
    if (reading !== 'below-floor-bucket0') return unmeasured()
  }
  const bucket = floorReached ? iconBucket(v, thresholds, gte) : 0

  let ref: CfIconRef | undefined
  if (def.custom) ref = def.reverse ? reversedCustom(def.custom, bucket, iconSelectionEvidence('icon:reverse-custom', check)) : def.custom[bucket]
  else ref = { set: def.set, iconId: def.reverse ? k - 1 - bucket : bucket }
  if (!ref) return unmeasured()
  const inRange = ref.set === 'NoIcons' ? ref.iconId === 0 : Number.isInteger(ref.iconId) && ref.iconId >= 0 && ref.iconId < familySize(ref.set)
  if (!inRange) return unresolved('cf-unsupported-icon')
  return { kind: 'ready', value: Object.freeze({ set: ref.set, iconId: ref.iconId, bucket, buckets: k, showValue: def.showValue }) }
}
