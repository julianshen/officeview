/** Pivot-item ordering for the Excel-measured Task 21 label classes. */
import type { PivotDiagnostic, PivotFieldSortType, PivotValue } from '../model'
import type { EngineResult } from './axes'

const fail = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-unmeasured', message } satisfies PivotDiagnostic,
})
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
/** Task 21 measured ASCII letters/digits/spaces, apostrophes and hyphens in words, and these leading code symbols. */
export function hasMeasuredEnglishOrder(value: string): boolean {
  return [...value].every((character, index) => /^[A-Za-z0-9 '_-]$/u.test(character) || (index === 0 && /^[#$@~]$/u.test(character)))
}

export function compareEnglishLabels(left: string, right: string): number {
  return compareMeasuredCodeLabels(left, right)
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

function excelCharacterWeight(character: string): number {
  const lower = character.toLowerCase()
  if (lower === '#') return 0
  if (lower === '$') return 1
  if (lower === '@') return 2
  if (lower === '_') return 3
  if (lower === '~') return 4
  if (lower === ' ') return 0
  if (lower >= '0' && lower <= '9') return 10 + lower.charCodeAt(0) - 48
  if (lower >= 'a' && lower <= 'z') return 30 + lower.charCodeAt(0) - 97
  return 100 + lower.charCodeAt(0)
}
function compareMeasuredCodeLabels(left: string, right: string): number {
  const primary = (value: string) => [...value.toLowerCase()].filter(character => character !== "'" && character !== '-')
  const aText = primary(left), bText = primary(right)
  const count = Math.min(aText.length, bText.length)
  for (let index = 0; index < count; index++) {
    const a = excelCharacterWeight(aText[index]!), b = excelCharacterWeight(bText[index]!)
    if (a !== b) return a - b
  }
  if (aText.length !== bText.length) return aText.length - bText.length
  // Excel puts a spelling with inserted punctuation after the otherwise equal plain spelling;
  // when both have punctuation, apostrophe precedes hyphen (Task 15 Code oracle).
  const punctuation = (value: string) => [...value.toLowerCase()].filter(character => character === "'" || character === '-').map(character => character === "'" ? 1 : 2)
  const aPunctuation = punctuation(left), bPunctuation = punctuation(right)
  for (let index = 0; index < Math.min(aPunctuation.length, bPunctuation.length); index++) {
    if (aPunctuation[index] !== bPunctuation[index]) return aPunctuation[index]! - bPunctuation[index]!
  }
  return aPunctuation.length - bPunctuation.length
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
function compareExcelLabels(left: string, right: string, customListSort: boolean): number {
  if (customListSort) {
    const a = customListIndex(left), b = customListIndex(right)
    if (a && b && a.kind === b.kind) return a.index - b.index
  }
  return compareMeasuredCodeLabels(left, right)
}

function hasUnmeasuredEqualWeightPair(values: readonly (PivotValue | undefined)[]): boolean {
  const labels = values.flatMap(value => value?.type === 'string' ? [value.value.toLowerCase()] : [])
  const labelSet = new Set(labels)
  const hyphensByText = new Map<string, string>()
  for (const label of labels) {
    if ((label.startsWith(' ') && labelSet.has(`#${label.slice(1)}`)) ||
      (label.startsWith('#') && labelSet.has(` ${label.slice(1)}`))) return true
    const positions = [...label].flatMap((character, index) => character === '-' ? [index] : []).join(',')
    if (!positions) continue
    const withoutHyphens = label.replace(/-/gu, '')
    const previous = hyphensByText.get(withoutHyphens)
    if (previous !== undefined && previous !== positions) return true
    hyphensByText.set(withoutHyphens, positions)
  }
  return false
}

/** Stable ascending/descending order for homogeneous labels; item ids remain the pivotField item indexes. */
export function sortItemOrder(current: readonly number[], valuesByItem: readonly PivotValue[], sortType: PivotFieldSortType, field: number,
  customListSort = true): EngineResult<number[]> {
  if (sortType === 'manual') return { ok: true, value: [...current] }
  const values = current.map(item => valuesByItem[item])
  if (values.some(value => value?.type === 'boolean')) return fail(`boolean items in sorted field ${field} have no measured Excel order`)
  const presentTypes = new Set(values.filter((value): value is PivotValue => value !== undefined).map(value => value.type))
  if (presentTypes.has('date') && presentTypes.size > 1) return fail(`mixed date items in sorted field ${field} have no measured Excel order`)
  if (values.some(value => value?.type === 'string' && !hasMeasuredEnglishOrder(value.value)))
    return fail(`Excel's saved punctuation order for sorted field ${field} contains characters outside the measured alphanumeric, space and Code punctuation set`)
  if (hasUnmeasuredEqualWeightPair(values))
    return fail(`equal-weight space/hash or hyphen-position labels in sorted field ${field} have no measured Excel order`)
  if (customListSort) {
    const lists = values.map(value => value?.type === 'string' ? customListIndex(value.value) : undefined)
    const kinds = new Set(lists.flatMap(value => value ? [value.kind] : []))
    if (kinds.size > 1 || (kinds.size > 0 && lists.some(value => value === undefined)))
      return fail(`sorted field ${field} mixes a custom-list item with labels outside one measured custom list`)
  }
  const rank: Record<string, number> = { number: 0, date: 0, boolean: 1, string: 2, missing: 3, blank: 3, error: 4 }
  const direction = sortType === 'descending' ? -1 : 1
  const sorted = [...current].sort((left, right) => {
    const a = valuesByItem[left], b = valuesByItem[right]
    if (!a || !b) return left - right
    const categoryOrder = (rank[category(a)] ?? 2) - (rank[category(b)] ?? 2)
    const compared = categoryOrder || (a.type === 'string' && b.type === 'string'
      ? compareExcelLabels(a.value, b.value, customListSort)
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
