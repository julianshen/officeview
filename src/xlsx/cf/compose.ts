import type { CfCellState, CfOutcome, CfRule, CfWriteMask } from './model'
import { hasEvidence as defaultHasEvidence } from './evidence'

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

  if (options?.isTextCellInScale) {
    if (!checkEvidence('scale:composition-text-cell-lower-dxf')) {
      tainted.add('fill')
    }
  }

  // Pre-scan resolved matching fills to arbitrate fill slot under Ruling S2-I & S2-J
  const matchingFills: MatchingFill[] = []
  let fillsWithheld = false
  for (let i = 0; i < orderedRules.length; i++) {
    yield
    const rule = orderedRules[i]!
    const outcome = outcomes[i]
    const isScale = rule.definition.kind === 'colorScale' || rule.writes.has('colorScale')
    // Mirror the execution pass: a scale carrying an unevidenced stopIfTrue, or an unresolved stopping rule, halts
    // composition, so nothing below it may take part in fill arbitration (it could suppress an earlier winner).
    const stopsHere = rule.stopIfTrue && (!isScale || !checkEvidence('scale:composition-scale-stopIfTrue'))
    if (isScale && rule.stopIfTrue && !checkEvidence('scale:composition-scale-stopIfTrue')) break
    if (!outcome || outcome.kind === 'unresolved') {
      if (stopsHere) break
      continue
    }
    if (outcome.kind === 'no-match' || !outcome.value) continue

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
    if (rule.stopIfTrue && !isScale) break
  }

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
    const isScale = rule.definition.kind === 'colorScale' || rule.writes.has('colorScale')

    // Ruling S2-J: scale carrying stopIfTrue="1"
    if (isScale && rule.stopIfTrue) {
      if (!checkEvidence('scale:composition-scale-stopIfTrue')) {
        // Unresolved until composition variant lands (S2-J); withholds p2 DXF fill and font
        tainted.add('ALL')
        break
      }
      // Once measured, flag is ignored on colorScale (OD-1)
    }

    if (!outcome || outcome.kind === 'unresolved') {
      if (rule.stopIfTrue && (!isScale || !checkEvidence('scale:composition-scale-stopIfTrue'))) {
        tainted.add('ALL')
        break
      }
      for (const path of rule.writes) { yield; tainted.add(path) }
      continue
    }

    if (outcome.kind === 'no-match' || !outcome.value) continue

    // Matching stopIfTrue rule ahead of a scale
    if (rule.stopIfTrue && !isScale) {
      const hasLowerScale = orderedRules.slice(i + 1).some(r => r.definition.kind === 'colorScale' || r.writes.has('colorScale'))
      if (hasLowerScale) {
        if (!checkEvidence('scale:composition-matching-stopIfTrue-ahead')) {
          tainted.add('fill')
        }
      }
    }

    if (isScale) {
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

    if (rule.stopIfTrue && (!isScale || !checkEvidence('scale:composition-scale-stopIfTrue'))) {
      break
    }
  }

  const result: CfCellState = Object.freeze({
    properties: Object.freeze(properties),
    ...(tainted.size ? { unresolvedWrites: readonlyMask(tainted) } : {}),
    ...(visualColorScale ? { visual: Object.freeze({ colorScale: Object.freeze(visualColorScale) }) } : {}),
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
