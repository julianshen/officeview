import type { CfCellState, CfOutcome, CfRule, CfWriteMask } from './model'

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
const conflicts = (a: string, b: string): boolean => a === 'ALL' || b === 'ALL' || a === b || a.startsWith(b + '.') || b.startsWith(a + '.')

/** Inputs are already ordered by (priority, documentOrder) and scoped to one cell.
 * Caller admits bounded matching-rule and output workspace before invoking this pure helper. */
function* steps(orderedRules: readonly CfRule[], outcomes: readonly CfOutcome<boolean>[]): Generator<void, CfCellState> {
  const properties: Record<string, string | number | boolean | null> = Object.create(null)
  const tainted = new Set<string>(), won = new Set<string>()
  for (let i = 0; i < orderedRules.length; i++) {
    yield
    const rule = orderedRules[i]!, outcome = outcomes[i]
    if (!outcome || outcome.kind === 'unresolved') {
      if (rule.stopIfTrue) { tainted.add('ALL'); break }
      for (const path of rule.writes) { yield; tainted.add(path) }
      continue
    }
    if (outcome.kind === 'no-match' || !outcome.value) continue
    const added = new Set<string>()
    for (const path in rule.dxf) {
      const slot = group(path)
      let blocked = false
      for (const mask of tainted) { yield; if (conflicts(mask, slot)) { blocked = true; break } }
      if (!blocked) for (const mask of won) { yield; if (conflicts(mask, slot)) { blocked = true; break } }
      if (blocked) continue
      yield
      properties[path] = rule.dxf[path]!
      added.add(slot)
    }
    for (const slot of added) won.add(slot)
    if (rule.stopIfTrue) break
  }
  return Object.freeze({ properties: Object.freeze(properties), ...(tainted.size ? { unresolvedWrites: readonlyMask(tainted) } : {}) })
}

export function composeCell(orderedRules: readonly CfRule[], outcomes: readonly CfOutcome<boolean>[]): CfCellState {
  const run = steps(orderedRules, outcomes)
  let next = run.next()
  while (!next.done) next = run.next()
  return next.value
}
/** Same compositor, with cooperative checkpoints inside rule/property/mask iteration. */
export async function composeCellAsync(orderedRules: readonly CfRule[], outcomes: readonly CfOutcome<boolean>[], work: () => Promise<void>): Promise<CfCellState> {
  const run = steps(orderedRules, outcomes)
  let next = run.next()
  while (!next.done) { await work(); next = run.next() }
  return next.value
}
