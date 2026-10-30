/** Pivot-item label sorting with explicit refusals for measured Excel-specific orderings. */
import type { PivotDiagnostic, PivotFieldSortType, PivotValue } from '../model'
import type { EngineResult } from './axes'

const fail = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-unmeasured', message } satisfies PivotDiagnostic,
})
const MONTHS = new Set(['jan', 'january', 'feb', 'february', 'mar', 'march', 'apr', 'april', 'may', 'jun', 'june', 'jul', 'july', 'aug', 'august', 'sep', 'sept', 'september', 'oct', 'october', 'nov', 'november', 'dec', 'december'])
const WEEKDAYS = new Set(['sun', 'sunday', 'mon', 'monday', 'tue', 'tues', 'tuesday', 'wed', 'wednesday', 'thu', 'thur', 'thurs', 'thursday', 'fri', 'friday', 'sat', 'saturday'])
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
  if (a.type === 'string' && b.type === 'string') {
    const left = a.value.toLowerCase(), right = b.value.toLowerCase()
    return left < right ? -1 : left > right ? 1 : 0
  }
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
  if (values.some(value => value.type === 'string' && /[^\x00-\x7f]/u.test(value.value)))
    return fail(`non-ASCII text collation for sorted field ${field} is not measured`)
  if (values.some(value => value.type === 'string' && (MONTHS.has(value.value.trim().toLocaleLowerCase()) || WEEKDAYS.has(value.value.trim().toLocaleLowerCase()))))
    return fail(`custom-list order for sorted field ${field} is not measured`)
  const direction = sortType === 'descending' ? -1 : 1
  const sorted = [...current].sort((left, right) => {
    const a = valuesByItem[left], b = valuesByItem[right]
    if (!a || !b) return left - right
    return direction * (compare(a, b) ?? 0) || left - right
  })
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
