/** Pivot-item ordering for the Excel-measured Task 21 label classes. */
import type { PivotDiagnostic, PivotFieldSortType, PivotValue } from '../model'
import type { EngineResult } from './axes'

const fail = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-unmeasured', message } satisfies PivotDiagnostic,
})
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
/** Task 21's Excel oracle covers alphanumeric labels, spaces and the punctuation in its Code corpus. */
export function hasMeasuredEnglishOrder(value: string): boolean {
  return /^[A-Za-z0-9 #$@_~'\-]*$/u.test(value)
}

export function compareEnglishLabels(left: string, right: string): number {
  return compareMeasuredCodeLabels(left, right)
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
  if (value.type === 'error') return 'error'
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

function excelCharacterWeight(character: string, position: number): number {
  const lower = character.toLowerCase()
  if (position === 0) {
    if (lower === '#') return 0
    if (lower === '$') return 1
    if (lower === '@') return 2
    if (lower === '_') return 3
    if (lower === '~') return 4
    if (lower === '-') return 90
  }
  if (lower === ' ') return 0
  if (lower === '_') return 1
  if (lower >= '0' && lower <= '9') return 10 + lower.charCodeAt(0) - 48
  if (lower >= 'a' && lower <= 'z') return 30 + lower.charCodeAt(0) - 97
  if (lower === "'") return 60
  if (lower === '-') return 70
  return 100 + lower.charCodeAt(0)
}
function compareMeasuredCodeLabels(left: string, right: string): number {
  const count = Math.min(left.length, right.length)
  for (let index = 0; index < count; index++) {
    const a = excelCharacterWeight(left[index]!, index), b = excelCharacterWeight(right[index]!, index)
    if (a !== b) return a - b
  }
  return left.length - right.length
}
function customListIndex(value: string): { kind: 'month' | 'weekday'; index: number } | undefined {
  const month = MONTHS.indexOf(value.trim().toLocaleLowerCase())
  if (month >= 0) return { kind: 'month', index: month }
  const shortMonths = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
  const shortMonth = shortMonths.indexOf(value.trim().toLocaleLowerCase())
  if (shortMonth >= 0) return { kind: 'month', index: shortMonth }
  const weekdays = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday']
  const weekday = weekdays.indexOf(value.trim().toLocaleLowerCase())
  if (weekday >= 0) return { kind: 'weekday', index: weekday }
  const shortWeekdays = ['sun','mon','tue','wed','thu','fri','sat']
  const shortWeekday = shortWeekdays.indexOf(value.trim().toLocaleLowerCase())
  return shortWeekday >= 0 ? { kind: 'weekday', index: shortWeekday } : undefined
}
function compareExcelLabels(left: string, right: string, customListSort: boolean, groupLabels: boolean): number {
  if (groupLabels) {
    const ordinal = (value: string): number | undefined => {
      if (value.startsWith('<')) return Number.NEGATIVE_INFINITY
      if (value.startsWith('>')) return Number.POSITIVE_INFINITY
      return /^(\d{4})(?:年)?$/u.exec(value)?.[1] ? Number(/^(\d{4})(?:年)?$/u.exec(value)![1]) : undefined
    }
    const a = ordinal(left), b = ordinal(right)
    if (a !== undefined && b !== undefined) return a - b
    const rangeStart = (value: string): number | undefined => /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*[-–]/u.exec(value)?.[1] !== undefined
      ? Number(/^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*[-–]/u.exec(value)![1]) : undefined
    const rangeA = rangeStart(left), rangeB = rangeStart(right)
    if (rangeA !== undefined && rangeB !== undefined) return rangeA - rangeB
  }
  if (customListSort) {
    const a = customListIndex(left), b = customListIndex(right)
    if (a && b && a.kind === b.kind) return a.index - b.index
  }
  return compareMeasuredCodeLabels(left, right)
}

/** Stable ascending/descending order for homogeneous labels; item ids remain the pivotField item indexes. */
export function sortItemOrder(current: readonly number[], valuesByItem: readonly PivotValue[], sortType: PivotFieldSortType, field: number,
  customListSort = true, groupLabels = false): EngineResult<number[]> {
  if (sortType === 'manual') return { ok: true, value: [...current] }
  const values = current.map(item => valuesByItem[item])
  if (values.some(value => value?.type === 'string' && !groupLabels && !hasMeasuredEnglishOrder(value.value)))
    return fail(`Excel's saved punctuation order for sorted field ${field} contains characters outside the measured alphanumeric, space and Code punctuation set`)
  const rank: Record<string, number> = { number: 0, date: 0, boolean: 1, string: 2, missing: 3, blank: 3, error: 4 }
  const direction = sortType === 'descending' ? -1 : 1
  const sorted = [...current].sort((left, right) => {
    const a = valuesByItem[left], b = valuesByItem[right]
    if (!a || !b) return left - right
    const categoryOrder = (rank[category(a)] ?? 2) - (rank[category(b)] ?? 2)
    const compared = categoryOrder || (a.type === 'string' && b.type === 'string'
      ? compareExcelLabels(a.value, b.value, customListSort, groupLabels)
      : compare(a, b) ?? 0)
    return direction * compared || left - right
  })
  return { ok: true, value: sorted }
}

/** Builds item-indexed labels from the shared/group item list and pivot field's x references. */
export function itemValuesByPivotIndex(source: readonly PivotValue[], items?: readonly { x?: number; type: string; name?: string }[]): PivotValue[] {
  const values: PivotValue[] = []
  if (!items) return source.map(item => item)
  for (let index = 0; index < items.length; index++) {
    const x = items[index]!.x
    if (items[index]!.type === 'data' && items[index]!.name !== undefined) values[index] = { type: 'string', value: items[index]!.name! }
    else if (items[index]!.type === 'data' && x !== undefined && source[x]) values[index] = source[x]!
  }
  return values
}
