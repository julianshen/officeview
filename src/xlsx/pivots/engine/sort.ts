/** Pivot-item ordering for the Excel-measured Task 21 label classes. */
import type { PivotDiagnostic, PivotFieldSortType, PivotValue } from '../model'
import type { EngineResult } from './axes'

const fail = (message: string): EngineResult<never> => ({
  ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-engine-unmeasured', message } satisfies PivotDiagnostic,
})
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
/** Task 21 measured ASCII letters/digits/spaces, apostrophes and hyphens in words, and these leading code symbols. */
export function hasMeasuredEnglishOrder(value: string): boolean {
  // One pass over the code units (a surrogate fails both classes, as the code point would).
  for (let index = 0; index < value.length; index++) {
    const c = value.charCodeAt(index)
    const word = (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || c === 32 || c === 39 || c === 95 || c === 45
    if (word) continue
    if (index === 0 && (c === 35 || c === 36 || c === 64 || c === 126)) continue
    return false
  }
  return true
}

/** A label's collation key, computed once: primary weights (apostrophes and hyphens left out) and the punctuation
 * sequence that breaks ties. */
export interface EnglishLabelKey { weights: Int32Array; punctuation: Uint8Array }
export function englishLabelKey(value: string): EnglishLabelKey {
  const weights: number[] = [], punctuation: number[] = []
  for (const character of value.toLowerCase()) {
    if (character === "'") { punctuation.push(1); continue }
    if (character === '-') { punctuation.push(2); continue }
    weights.push(excelCharacterWeight(character))
  }
  return { weights: Int32Array.from(weights), punctuation: Uint8Array.from(punctuation) }
}
export function compareEnglishLabelKeys(left: EnglishLabelKey, right: EnglishLabelKey): number {
  const a = left.weights, b = right.weights, count = Math.min(a.length, b.length)
  for (let index = 0; index < count; index++) if (a[index] !== b[index]) return a[index]! - b[index]!
  if (a.length !== b.length) return a.length - b.length
  // Excel puts a spelling with inserted punctuation after the otherwise equal plain spelling;
  // when both have punctuation, apostrophe precedes hyphen (Task 15 Code oracle).
  const p = left.punctuation, q = right.punctuation
  for (let index = 0; index < Math.min(p.length, q.length); index++) if (p[index] !== q[index]) return p[index]! - q[index]!
  return p.length - q.length
}

export function compareEnglishLabels(left: string, right: string): number {
  return compareEnglishLabelKeys(englishLabelKey(left), englishLabelKey(right))
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

function hasUnmeasuredEqualWeightPair(values: readonly (PivotValue | undefined)[]): boolean {
  const labels = values.flatMap(value => value?.type === 'string' ? [value.value.toLowerCase()] : [])
  let hasLeadingSpace = false, hasLeadingHash = false
  const punctuationShapes = new Map<string, Set<string>>()
  for (const label of labels) {
    if (label.startsWith(' ')) hasLeadingSpace = true
    if (label.startsWith('#')) hasLeadingHash = true
    const shape = label.replace(/['-]/gu, '')
    const spellings = punctuationShapes.get(shape) ?? new Set<string>()
    spellings.add(label)
    punctuationShapes.set(shape, spellings)
  }
  if (hasLeadingSpace && hasLeadingHash) return true
  const measuredPunctuationShapes = new Map([
    ['ab', new Set(['ab', "a'b", 'a-b'])],
    // The Task 21 saved collation fixture also measures the leading-hyphen pair -x / x.
    ['x', new Set(['x', '-x'])],
  ])
  return [...punctuationShapes].some(([shape, spellings]) => spellings.size > 1 &&
    ![...spellings].every(spelling => measuredPunctuationShapes.get(shape)?.has(spelling)))
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
    return fail(`equal-weight leading-space, apostrophe or hyphen-position labels in sorted field ${field} have no measured Excel order`)
  if (customListSort) {
    const lists = values.map(value => value?.type === 'string' ? customListIndex(value.value) : undefined)
    const kinds = new Set(lists.flatMap(value => value ? [value.kind] : []))
    if (kinds.size > 1 || (kinds.size > 0 && lists.some(value => value === undefined)))
      return fail(`sorted field ${field} mixes a custom-list item with labels outside one measured custom list`)
  }
  const rank: Record<string, number> = { number: 0, date: 0, boolean: 1, string: 2, missing: 3, blank: 3, error: 4 }
  const direction = sortType === 'descending' ? -1 : 1
  // Collation keys and custom-list positions once per item, not per comparison (perf audit low note 5).
  const keys = new Map<number, EnglishLabelKey>(), lists = new Map<number, ReturnType<typeof customListIndex>>()
  for (const item of current) {
    const value = valuesByItem[item]
    if (value?.type !== 'string') continue
    keys.set(item, englishLabelKey(value.value))
    if (customListSort) lists.set(item, customListIndex(value.value))
  }
  const sorted = [...current].sort((left, right) => {
    const a = valuesByItem[left], b = valuesByItem[right]
    if (!a || !b) return left - right
    const categoryOrder = (rank[category(a)] ?? 2) - (rank[category(b)] ?? 2)
    let compared = categoryOrder
    if (!compared && a.type === 'string' && b.type === 'string') {
      const listA = lists.get(left), listB = lists.get(right)
      compared = listA && listB && listA.kind === listB.kind ? listA.index - listB.index : compareEnglishLabelKeys(keys.get(left)!, keys.get(right)!)
    } else if (!compared) compared = compare(a, b) ?? 0
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
