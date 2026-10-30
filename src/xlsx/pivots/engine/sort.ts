/** Pivot-item label sorting with explicit refusals for measured Excel-specific orderings. */
import type { PivotDiagnostic, PivotFieldSortType, PivotValue } from '../model'
import type { EngineResult } from './axes'

const fail = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-unmeasured', message } satisfies PivotDiagnostic,
})
const MONTHS = new Set(['jan', 'january', 'feb', 'february', 'mar', 'march', 'apr', 'april', 'may', 'jun', 'june', 'jul', 'july', 'aug', 'august', 'sep', 'sept', 'september', 'oct', 'october', 'nov', 'november', 'dec', 'december'])
const WEEKDAYS = new Set(['sun', 'sunday', 'mon', 'monday', 'tue', 'tues', 'tuesday', 'wed', 'wednesday', 'thu', 'thur', 'thurs', 'thursday', 'fri', 'friday', 'sat', 'saturday'])
const ENGLISH_LABELS = new Intl.Collator('en', { sensitivity: 'accent' })

/** Task 21 has measured only ASCII letters, digits and spaces for English label order. */
export function hasMeasuredEnglishOrder(value: string): boolean {
  return /^[A-Za-z0-9 ]*$/u.test(value)
}

export function compareEnglishLabels(left: string, right: string): number {
  return ENGLISH_LABELS.compare(left, right)
}
/** Detect a measured-order boundary where a space competes with a word character. */
export function hasUnmeasuredSpaceBoundary(leftValue: string, rightValue: string): boolean {
  const left = leftValue.toLowerCase(), right = rightValue.toLowerCase()
  const isAlphaNumeric = (value: string | undefined) => value !== undefined && /^[a-z0-9]$/u.test(value)
  let offset = 0
  while (offset < left.length && offset < right.length && left[offset] === right[offset]) offset++
  const a = left[offset], b = right[offset]
  return (a === ' ' && isAlphaNumeric(b)) || (b === ' ' && isAlphaNumeric(a))
}
/** Detect an unmeasured space-versus-word ordering without comparing every item pair. */
export function hasUnmeasuredSpaceOrder(values: readonly string[]): boolean {
  const labels = [...new Set(values.map(value => value.toLowerCase()))].sort()
  for (let index = 1; index < labels.length; index++) {
    if (hasUnmeasuredSpaceBoundary(labels[index - 1]!, labels[index]!)) return true
  }
  return false
}
function category(value: PivotValue): string {
  if (value.type === 'number') return 'number'
  if (value.type === 'date') return 'date'
  if (value.type === 'string') return 'string'
  if (value.type === 'missing') return 'blank'
  return value.type
}
function compare(a: PivotValue, b: PivotValue): number | undefined {
  if (a.type !== b.type) return undefined
  if (a.type === 'number' && b.type === 'number') return a.value - b.value
  if (a.type === 'date' && b.type === 'date') return a.value < b.value ? -1 : a.value > b.value ? 1 : 0
  if (a.type === 'string' && b.type === 'string') return compareEnglishLabels(a.value, b.value)
  if (a.type === 'missing' && b.type === 'missing') return 0
  if (a.type === 'boolean' && b.type === 'boolean') return Number(a.value) - Number(b.value)
  return undefined
}

/** Stable ascending/descending order for homogeneous labels; item ids remain the pivotField item indexes. */
export function sortItemOrder(current: readonly number[], valuesByItem: readonly PivotValue[], sortType: PivotFieldSortType, field: number): EngineResult<number[]> {
  if (sortType === 'manual') return { ok: true, value: [...current] }
  const values = current.map(item => valuesByItem[item]).filter((value): value is PivotValue => value !== undefined)
  if (new Set(values.map(category)).size > 1)
    return fail(`mixed-type item order for sorted field ${field} is not measured`)
  if (values.some(value => value.type === 'error')) return fail(`error-item order for sorted field ${field} is not measured`)
  if (values.some(value => value.type === 'string' && !hasMeasuredEnglishOrder(value.value)))
    return fail(`text collation outside ASCII letters, digits and spaces for sorted field ${field} is not measured`)
  if (values.some(value => value.type === 'string' && (MONTHS.has(value.value.trim().toLocaleLowerCase()) || WEEKDAYS.has(value.value.trim().toLocaleLowerCase()))))
    return fail(`custom-list order for sorted field ${field} is not measured`)
  const direction = sortType === 'descending' ? -1 : 1
  const sorted = [...current].sort((left, right) => {
    const a = valuesByItem[left], b = valuesByItem[right]
    if (!a || !b) return left - right
    return direction * (compare(a, b) ?? 0) || left - right
  })
  // Reuse the sort order already needed for output: the only risky labels will
  // meet at the space/word boundary after collation, without a second O(n log n) sort.
  for (let index = 1; index < sorted.length; index++) {
    const left = valuesByItem[sorted[index - 1]!], right = valuesByItem[sorted[index]!]
    if (left?.type === 'string' && right?.type === 'string' && hasUnmeasuredSpaceBoundary(left.value, right.value))
      return fail(`space-versus-word collation for sorted field ${field} is not measured`)
  }
  return { ok: true, value: sorted }
}

/** Builds item-indexed labels from the shared/group item list and pivot field's x references. */
export function itemValuesByPivotIndex(source: readonly PivotValue[], items?: readonly { x?: number; type: string }[]): PivotValue[] {
  const values: PivotValue[] = []
  if (!items) return source.map(item => item)
  for (let index = 0; index < items.length; index++) {
    const x = items[index]!.x
    if (items[index]!.type === 'data' && x !== undefined && source[x]) values[index] = source[x]!
  }
  return values
}
