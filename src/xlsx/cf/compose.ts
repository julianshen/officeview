import type { CfBarDescriptor, CfCellState, CfIconDescriptor, CfOutcome, CfRule, CfWriteMask } from './model'
import { hasEvidence as defaultHasEvidence, iconSelectionEvidence } from './evidence'

/** Private backing set: Object.freeze(new Set()) alone does not prevent add/delete. */
function readonlyMask(values: Set<string>): CfWriteMask {
  const view: ReadonlySet<string> = Object.freeze({
    get size() { return values.size }, has: (s: string) => values.has(s),
    entries: () => values.entries(), keys: () => values.keys(), values: () => values.values(),
    [Symbol.iterator]: () => values[Symbol.iterator](),
    forEach(callback: (value: string, key: string, set: ReadonlySet<string>) => void, thisArg?: unknown) {
      values.forEach(v => callback.call(thisArg, v, v, view))
    },
  })
  return view
}

const group = (path: string): string => path.startsWith('fill.') ? 'fill'
  : path.startsWith('numFmt.') ? 'numFmt' : path.startsWith('border.') ? path.split('.').slice(0, 2).join('.') : path

const conflicts = (a: string, b: string): boolean =>
  a === 'ALL' || b === 'ALL' || a === b || a.startsWith(b + '.') || b.startsWith(a + '.')

type VisualSlot = 'colorScale' | 'dataBar' | 'iconSet'

/** The rule's visual slot: scales own `colorScale` (and arbitrate the fill slot), bars own the independent `dataBar`
 * slot (Ruling S3-I), icon sets the independent `iconSet` slot (Ruling S4-I). Unresolved stubs keep their slot through
 * the canonical write masks. */
export function visualSlotOf(rule: CfRule): VisualSlot | undefined {
  if (rule.definition.kind === 'colorScale' || rule.writes.has('colorScale')) return 'colorScale'
  if (rule.definition.kind === 'dataBar' || rule.writes.has('dataBar')) return 'dataBar'
  if (rule.definition.kind === 'iconSet' || rule.writes.has('iconSet')) return 'iconSet'
  return undefined
}

/** A ready icon outcome carries its evaluated descriptor (without the rule id composition adds). */
function iconDescriptorOf(value: unknown): Omit<CfIconDescriptor, 'ruleId'> | undefined {
  return typeof value === 'object' && value !== null && 'buckets' in value && 'iconId' in value ? value as Omit<CfIconDescriptor, 'ruleId'> : undefined
}

/** A ready bar outcome carries its evaluated descriptor (without the rule id composition adds). */
function barDescriptorOf(value: unknown): Omit<CfBarDescriptor, 'ruleId'> | undefined {
  return typeof value === 'object' && value !== null && 'endFraction' in value ? value as Omit<CfBarDescriptor, 'ruleId'> : undefined
}

export interface ComposeCellOptions {
  readonly hasEvidence?: (cap: string) => boolean
  readonly isTextCellInScale?: boolean
}

interface MatchingFill {
  readonly rule: CfRule
  readonly index: number
  readonly isScale: boolean
  readonly argb?: number
  readonly fraction?: number
}

/** Inputs are already ordered by (priority, documentOrder) and scoped to one cell.
 * Caller admits bounded matching-rule and output workspace before invoking this pure helper. */
function* steps(
  orderedRules: readonly CfRule[],
  outcomes: readonly (CfOutcome<unknown> | undefined)[],
  options?: ComposeCellOptions,
): Generator<void, CfCellState> {
  const checkEvidence = options?.hasEvidence ?? defaultHasEvidence
  const properties: Record<string, string | number | boolean | null> = Object.create(null)
  const tainted = new Set<string>(), won = new Set<string>()
  let visualColorScale: { ruleId: number; argb: number } | undefined
  let visualDataBar: CfBarDescriptor | undefined
  let visualIconSet: CfIconDescriptor | undefined
  // Ruling S3-J: a bar carrying stopIfTrue is unresolved until its composition is measured (OD-1 analogue).
  const barStopMeasured = checkEvidence('bar:composition-bar-stopIfTrue')
  const scaleStopMeasured = checkEvidence('scale:composition-scale-stopIfTrue')
  // Ruling S4-J: likewise an icon set carrying stopIfTrue.
  const iconStopMeasured = checkEvidence('icon:composition-icon-stopIfTrue')
  /** Whether a rule's stopIfTrue halts composition: visuals only once their flag's effect is measured as ignored. */
  const halts = (rule: CfRule, slot: VisualSlot | undefined): boolean =>
    rule.stopIfTrue && (slot === 'colorScale' ? !scaleStopMeasured : slot === 'dataBar' ? !barStopMeasured : slot === 'iconSet' ? !iconStopMeasured : true)

  if (options?.isTextCellInScale) {
    if (!checkEvidence('scale:composition-text-cell-lower-dxf')) {
      tainted.add('fill')
    }
  }

  // Pre-scan resolved matching fills to arbitrate fill slot under Ruling S2-I & S2-J
  const matchingFills: MatchingFill[] = []
  let fillsWithheld = false
  // Resolved matching bars reachable on this cell (S3-I exclusivity is decided over all of them).
  let matchingBars = 0
  // Resolved matching icons reachable on this cell, in priority order (S4-I exclusivity is decided over all of them).
  const matchingIcons: { index: number; noIcons: boolean }[] = []
  for (let i = 0; i < orderedRules.length; i++) {
    yield
    const rule = orderedRules[i]!
    const outcome = outcomes[i]
    const slot = visualSlotOf(rule)
    const isScale = slot === 'colorScale'
    // Mirror the execution pass: a visual carrying an unevidenced stopIfTrue, or an unresolved stopping rule, halts
    // composition, so nothing below it may take part in fill arbitration (it could suppress an earlier winner).
    const stopsHere = halts(rule, slot)
    if (slot !== undefined && stopsHere) break
    if (!outcome || outcome.kind === 'unresolved') {
      if (stopsHere) break
      continue
    }
    if (outcome.kind === 'no-match' || !outcome.value) continue
    if (slot === 'dataBar') {
      // A bar never takes part in fill arbitration (S3-I); it only counts towards bar exclusivity.
      if (barDescriptorOf(outcome.value)) matchingBars++
      continue
    }
    if (slot === 'iconSet') {
      // Nor does an icon (S4-I); it only counts towards icon exclusivity.
      const icon = iconDescriptorOf(outcome.value)
      if (icon) matchingIcons.push({ index: i, noIcons: icon.set === 'NoIcons' })
      continue
    }

    const hasDxfFill = rule.dxf && ('fill.bg' in rule.dxf || 'fill.pattern' in rule.dxf)

    if (isScale) {
      let argb: number | undefined
      let fraction: number | undefined
      if (typeof outcome.value === 'number') {
        argb = outcome.value
      } else if (typeof outcome.value === 'object' && outcome.value !== null && 'argb' in outcome.value) {
        argb = (outcome.value as { argb: number }).argb
        fraction = (outcome.value as { fraction?: number }).fraction
      }
      if (argb === undefined) {
        // M2: a scale outcome that carries no evaluated colour (a bare `true`) is unresolved, never endpoint 0.
        fillsWithheld = true
        tainted.add('fill')
        tainted.add('colorScale')
        continue
      }
      matchingFills.push({ rule, index: i, isScale: true, argb, fraction })
    } else if (hasDxfFill) {
      matchingFills.push({ rule, index: i, isScale: false })
    }

    // A matching stopIfTrue DXF halts composition: scales after it are
    // unreachable and must not participate in fill arbitration (bot P2).
    if (stopsHere) break
  }
  // S3-I: two matching bars on one cell withhold both and taint the bar slot, until bar:composition-overlapping-bars is
  // verified. Then Ruling S3-I-OVERLAP: the first matching bar in priority order paints, and every later one is
  // suppressed by the won slot (a known loser, so not reported as unresolved).
  const barsWithheld = matchingBars > 1 && !checkEvidence('bar:composition-overlapping-bars')
  if (barsWithheld) tainted.add('dataBar')
  // S4-I: two matching icons withhold both and taint the icon slot until icon:composition-overlapping-icons; then the
  // first in priority order paints. A NoIcons win over a lower matching icon withholds the lower icon (and the slot)
  // until icon:composition-noicons-slot selects a reading: `first-paints` (NoIcons occupies the slot) or
  // `second-paints` (the next matching icon paints).
  let iconsWithheld = matchingIcons.length > 1 && !checkEvidence('icon:composition-overlapping-icons')
  let iconWinnerIndex = matchingIcons[0]?.index
  if (!iconsWithheld && matchingIcons.length > 1 && matchingIcons[0]!.noIcons) {
    const reading = iconSelectionEvidence('icon:composition-noicons-slot', checkEvidence)
    if (reading === 'second-paints') iconWinnerIndex = matchingIcons[1]!.index
    else if (reading !== 'first-paints') iconsWithheld = true
  }
  if (iconsWithheld) tainted.add('iconSet')
  // S3-I coexistence: a bar never changes the fill slot, but paints over a scale or DXF fill only under evidence.
  const scaleFillPresent = matchingFills.some(f => f.isScale)
  const dxfFillPresent = matchingFills.some(f => !f.isScale)

  // Determine fill slot winner
  let fillWinnerIndex: number | undefined = undefined

  if (matchingFills.length === 1) {
    fillWinnerIndex = matchingFills[0]!.index
  } else if (matchingFills.length > 1) {
    const scaleFills = matchingFills.filter(f => f.isScale)
    const dxfFills = matchingFills.filter(f => !f.isScale)

    if (scaleFills.length > 1) {
      // Two overlapping scales withhold the fill (S2-I)
      fillsWithheld = true
      tainted.add('fill')
      tainted.add('colorScale')
    } else if (scaleFills.length === 1 && dxfFills.length >= 1) {
      const scaleFill = scaleFills[0]!
      const dxfFill = dxfFills[0]!

      if (dxfFill.rule.priority < scaleFill.rule.priority) {
        // DXF p1, Scale p2
        if (checkEvidence('scale:composition-dxf-p1-scale-p2')) {
          fillWinnerIndex = dxfFill.index
        } else {
          fillsWithheld = true
          tainted.add('fill')
          tainted.add('colorScale')
        }
      } else if (scaleFill.rule.priority < dxfFill.rule.priority) {
        // Scale p1, DXF p2
        if (checkEvidence('scale:composition-scale-p1-dxf-p2')) {
          fillWinnerIndex = scaleFill.index
        } else {
          fillsWithheld = true
          tainted.add('fill')
          tainted.add('colorScale')
        }
      } else {
        // Equal priority. Only the DXF-first document order was measured (stage2-scale-composition D2:D5);
        // with the scale first the winner is unobserved, so that order withholds (final review I3-C).
        if (checkEvidence('scale:composition-equal-priority') && dxfFill.rule.documentOrder < scaleFill.rule.documentOrder) {
          // In Excel measured stage2-scale-composition Col D, scale won
          fillWinnerIndex = scaleFill.index
        } else {
          fillsWithheld = true
          tainted.add('fill')
          tainted.add('colorScale')
        }
      }
    }
  }

  // Execute rule composition
  for (let i = 0; i < orderedRules.length; i++) {
    yield
    const rule = orderedRules[i]!
    const outcome = outcomes[i]
    const slot = visualSlotOf(rule)
    const isScale = slot === 'colorScale'
    const isBar = slot === 'dataBar'
    const isIcon = slot === 'iconSet'

    // Ruling S2-J (scale) and S3-J (bar): a visual carrying stopIfTrue="1" is unresolved until its composition variant
    // lands, whatever its outcome, so §5.7.1 withholds every later rule (the p2 DXF fill and font). Once measured, the
    // flag is ignored on the visual (OD-1).
    if (slot !== undefined && halts(rule, slot)) {
      tainted.add('ALL')
      break
    }

    if (!outcome || outcome.kind === 'unresolved') {
      if (halts(rule, slot)) {
        tainted.add('ALL')
        break
      }
      for (const path of rule.writes) { yield; tainted.add(path) }
      continue
    }

    if (outcome.kind === 'no-match' || !outcome.value) continue

    // Matching stopIfTrue rule ahead of a scale
    if (rule.stopIfTrue && slot === undefined) {
      const hasLowerScale = orderedRules.slice(i + 1).some(r => r.definition.kind === 'colorScale' || r.writes.has('colorScale'))
      if (hasLowerScale) {
        if (!checkEvidence('scale:composition-matching-stopIfTrue-ahead')) {
          tainted.add('fill')
        }
      }
    }

    if (isBar) {
      // Bar rule (S3-I): its own slot, exclusive, and over a fill only under evidence.
      const value = barDescriptorOf(outcome.value)
      if (won.has('dataBar')) continue
      let blocked = barsWithheld || value === undefined
      for (const mask of tainted) { yield; if (conflicts(mask, 'dataBar')) { blocked = true; break } }
      // An earlier unresolved fill writer leaves the cell's fill unknown: treated as a possible DXF fill.
      let fillUnknown = false
      for (const mask of tainted) { yield; if (conflicts(mask, 'fill')) { fillUnknown = true; break } }
      if (!blocked && scaleFillPresent && !checkEvidence('bar:composition-with-scale')) blocked = true
      if (!blocked && (dxfFillPresent || fillUnknown) && !checkEvidence('bar:composition-dxf-fill')) blocked = true
      if (blocked || !value) tainted.add('dataBar')
      else { visualDataBar = Object.freeze({ ruleId: rule.id, ...value }); won.add('dataBar') }
    } else if (isIcon) {
      // Icon rule (S4-I): its own slot, exclusive, and with a bar or over a fill only under evidence. Coexistence is
      // decided from the pre-scan, so both priority orders behave the same; the bar keeps its own Stage 3 outcome.
      const value = iconDescriptorOf(outcome.value)
      if (won.has('iconSet') || (value && iconWinnerIndex !== undefined && i !== iconWinnerIndex && !iconsWithheld)) continue
      let blocked = iconsWithheld || value === undefined
      for (const mask of tainted) { yield; if (conflicts(mask, 'iconSet')) { blocked = true; break } }
      // An earlier unresolved bar or fill writer leaves that slot unknown: treated as a possible bar or DXF fill.
      let barUnknown = false, fillUnknown = false
      for (const mask of tainted) { yield; if (conflicts(mask, 'dataBar')) barUnknown = true; if (conflicts(mask, 'fill')) fillUnknown = true }
      if (!blocked && (matchingBars > 0 || barUnknown) && !checkEvidence('icon:composition-with-bar')) blocked = true
      if (!blocked && scaleFillPresent && !checkEvidence('icon:composition-with-scale')) blocked = true
      if (!blocked && (dxfFillPresent || fillUnknown) && !checkEvidence('icon:composition-dxf-fill')) blocked = true
      if (blocked || !value) tainted.add('iconSet')
      else { visualIconSet = Object.freeze({ ruleId: rule.id, ...value }); won.add('iconSet') }
    } else if (isScale) {
      // Scale rule
      let blocked = fillsWithheld
      for (const mask of tainted) { yield; if (conflicts(mask, 'fill') || conflicts(mask, 'colorScale')) { blocked = true; break } }
      if (!blocked && fillWinnerIndex === i) {
        const fillInfo = matchingFills.find(f => f.index === i)
        if (fillInfo && fillInfo.argb !== undefined) {
          visualColorScale = { ruleId: rule.id, argb: fillInfo.argb, ...(fillInfo.fraction !== undefined ? { fraction: fillInfo.fraction } : {}) }
          won.add('fill')
          won.add('colorScale')
        }
      }
    } else {
      // DXF rule
      const added = new Set<string>()
      for (const path in rule.dxf) {
        const slot = group(path)
        let blocked = false
        if (slot === 'fill' && (fillsWithheld || (fillWinnerIndex !== undefined && fillWinnerIndex !== i))) {
          blocked = true
        }
        if (!blocked) {
          for (const mask of tainted) { yield; if (conflicts(mask, slot)) { blocked = true; break } }
        }
        if (!blocked) {
          for (const mask of won) { yield; if (conflicts(mask, slot)) { blocked = true; break } }
        }
        if (blocked) continue
        yield
        properties[path] = rule.dxf[path]!
        added.add(slot)
      }
      for (const slot of added) won.add(slot)
    }

    if (halts(rule, slot)) break
  }

  // S4-I: a painted bar and a painted icon that disagree on showValue keep both descriptors, but whether the text is
  // suppressed is unresolved (the text stays) until icon:composition-bar-showvalue.
  if (visualDataBar && visualIconSet && visualDataBar.showValue !== visualIconSet.showValue && !checkEvidence('icon:composition-bar-showvalue')) {
    tainted.add('showValue')
  }
  const visual = visualColorScale || visualDataBar || visualIconSet ? Object.freeze({
    ...(visualColorScale ? { colorScale: Object.freeze(visualColorScale) } : {}),
    ...(visualDataBar ? { dataBar: visualDataBar } : {}),
    ...(visualIconSet ? { iconSet: visualIconSet } : {}),
  }) : undefined
  const result: CfCellState = Object.freeze({
    properties: Object.freeze(properties),
    ...(tainted.size ? { unresolvedWrites: readonlyMask(tainted) } : {}),
    ...(visual ? { visual } : {}),
  })

  return result
}

export function composeCell(
  orderedRules: readonly CfRule[],
  outcomes: readonly (CfOutcome<unknown> | undefined)[],
  options?: ComposeCellOptions,
): CfCellState {
  const run = steps(orderedRules, outcomes, options)
  let next = run.next()
  while (!next.done) next = run.next()
  return next.value
}

/** Same compositor, with cooperative checkpoints inside rule/property/mask iteration. */
export async function composeCellAsync(
  orderedRules: readonly CfRule[],
  outcomes: readonly (CfOutcome<unknown> | undefined)[],
  work: () => Promise<void>,
  options?: ComposeCellOptions,
): Promise<CfCellState> {
  const run = steps(orderedRules, outcomes, options)
  let next = run.next()
  while (!next.done) { await work(); next = run.next() }
  return next.value
}
