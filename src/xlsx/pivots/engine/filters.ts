/** Pure filter predicates shared by computeAxes. */
import type { PivotFilter, PivotFilterColumn, PivotFilterType, PivotValue } from '../model'
import { compareEnglishLabels, hasMeasuredEnglishOrder } from './sort'

const CAPTION_TYPES = new Set<PivotFilterType>([
  'captionEqual', 'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith', 'captionEndsWith', 'captionNotEndsWith',
  'captionContains', 'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan',
  'captionLessThanOrEqual', 'captionBetween', 'captionNotBetween',
])
const VALUE_TYPES = new Set<PivotFilterType>([
  'valueEqual', 'valueNotEqual', 'valueGreaterThan', 'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual', 'valueBetween', 'valueNotBetween',
])
const DATE_TYPES = new Set<PivotFilterType>([
  'dateEqual', 'dateNotEqual', 'dateOlderThan', 'dateOlderThanOrEqual', 'dateNewerThan', 'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween',
])
const DYNAMIC_TYPES = new Set<PivotFilterType>([
  'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth',
  'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate',
])
export const isLabelFilter = (filter: PivotFilter) => CAPTION_TYPES.has(filter.type)
export const isValueFilter = (filter: PivotFilter) => VALUE_TYPES.has(filter.type)
export const isDateFilter = (filter: PivotFilter) => DATE_TYPES.has(filter.type)
export const isDynamicDateFilter = (filter: PivotFilter) => DYNAMIC_TYPES.has(filter.type) || filter.autoFilter.columns.some(column => column.dynamicFilter !== undefined)
export const isTop10Filter = (filter: PivotFilter) => filter.type === 'count' || filter.type === 'percent' || filter.type === 'sum'

function criterion(filter: PivotFilter, column: PivotFilterColumn | undefined, which: number): string | undefined {
  const direct = which === 1 ? filter.stringValue1 : filter.stringValue2
  if (direct !== undefined) return direct
  const values = column?.customFilters?.filters ?? []
  return values[which - 1]?.val
}
function wildcardRegex(pattern: string, mode: 'exact' | 'prefix' | 'contains' | 'suffix'): RegExp {
  const key = `${mode}\u0000${pattern}`
  const cached = WILDCARD_CACHE.get(key)
  if (cached) return cached
  let source = mode === 'contains' || mode === 'suffix' ? '.*' : ''
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!
    if (ch === '~' && i + 1 < pattern.length) source += pattern[++i]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    else if (ch === '*') source += '.*'
    else if (ch === '?') source += '.'
    else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
  if (mode === 'contains' || mode === 'prefix') source += '.*'
  const regex = new RegExp(`^${source}$`, 'isu')
  if (WILDCARD_CACHE.size >= 128) WILDCARD_CACHE.clear()
  WILDCARD_CACHE.set(key, regex)
  return regex
}
const WILDCARD_CACHE = new Map<string, RegExp>()
const printableAscii = (text: string) => /^[\x20-\x7e]*$/u.test(text)
function captionCompare(label: string, op: PivotFilterType, a: string, b?: string): boolean | undefined {
  const usesOrder = op === 'captionGreaterThan' || op === 'captionGreaterThanOrEqual' || op === 'captionLessThan' ||
    op === 'captionLessThanOrEqual' || op === 'captionBetween' || op === 'captionNotBetween'
  if (usesOrder && (!hasMeasuredEnglishOrder(label) || !hasMeasuredEnglishOrder(a) || (b !== undefined && !hasMeasuredEnglishOrder(b)))) return undefined
  const firstVsLabel = compareEnglishLabels(label, a), secondVsLabel = b === undefined ? undefined : compareEnglishLabels(label, b)
  const glob = wildcardRegex(a, op === 'captionBeginsWith' || op === 'captionNotBeginsWith' ? 'prefix'
    : op === 'captionEndsWith' || op === 'captionNotEndsWith' ? 'suffix'
      : op === 'captionContains' || op === 'captionNotContains' ? 'contains' : 'exact')
  switch (op) {
    case 'captionEqual': return glob.test(label)
    case 'captionNotEqual': return !glob.test(label)
    case 'captionBeginsWith': return glob.test(label)
    case 'captionNotBeginsWith': return !glob.test(label)
    case 'captionEndsWith': return glob.test(label)
    case 'captionNotEndsWith': return !glob.test(label)
    case 'captionContains': return glob.test(label)
    case 'captionNotContains': return !glob.test(label)
    case 'captionGreaterThan': return firstVsLabel > 0
    case 'captionGreaterThanOrEqual': return firstVsLabel >= 0
    case 'captionLessThan': return firstVsLabel < 0
    case 'captionLessThanOrEqual': return firstVsLabel <= 0
    case 'captionBetween': return secondVsLabel !== undefined && firstVsLabel >= 0 && secondVsLabel <= 0
    case 'captionNotBetween': return secondVsLabel !== undefined && (firstVsLabel < 0 || secondVsLabel > 0)
    default: return false
  }
}
export function matchesLabelFilter(filter: PivotFilter, value: PivotValue | undefined): boolean | undefined {
  if (!isLabelFilter(filter)) return undefined
  const column = filter.autoFilter.columns[0]
  if (!value || value.type !== 'string' || !printableAscii(value.value)) return undefined
  const label = value.value
  if (column?.filters) {
    if (column.filters.values.some(candidate => !printableAscii(candidate))) return undefined
    return column.filters.values.some(candidate => candidate.toLowerCase() === label.toLowerCase())
  }
  const a = criterion(filter, column, 1)
  const b = criterion(filter, column, 2)
  const custom = column?.customFilters?.filters ?? []
  if ((!custom.length && a === undefined) || (a !== undefined && !printableAscii(a)) || (b !== undefined && !printableAscii(b))) return undefined
  if (custom.length) {
    const decisions = custom.map(rule => {
      if (rule.val === undefined || !printableAscii(rule.val)) return undefined
      const op: Record<string, PivotFilterType> = {
        equal: 'captionEqual', notEqual: 'captionNotEqual', greaterThan: 'captionGreaterThan', greaterThanOrEqual: 'captionGreaterThanOrEqual',
        lessThan: 'captionLessThan', lessThanOrEqual: 'captionLessThanOrEqual',
      }
      return captionCompare(label, op[rule.operator] ?? 'captionEqual', rule.val)
    })
    if (decisions.some(decision => decision === undefined)) return undefined
    return column!.customFilters!.and ? decisions.every(Boolean) : decisions.some(Boolean)
  }
  return captionCompare(label, filter.type, a!, b)
}

function asNumber(value: PivotValue | undefined): number | undefined {
  if (value?.type === 'number') return value.value
  if (value?.type === 'date') {
    const time = Date.parse(value.value.endsWith('Z') ? value.value : `${value.value}Z`)
    return Number.isFinite(time) ? time : undefined
  }
  return undefined
}
export function matchesValueFilter(filter: PivotFilter, value: number | undefined): boolean | undefined {
  if (!isValueFilter(filter)) return undefined
  const a = Number(criterion(filter, filter.autoFilter.columns[0], 1)), bRaw = criterion(filter, filter.autoFilter.columns[0], 2), b = bRaw === undefined ? undefined : Number(bRaw)
  if (!Number.isFinite(a)) return undefined
  switch (filter.type) {
    case 'valueEqual': return value === a
    case 'valueNotEqual': return value !== a
    case 'valueGreaterThan': return value !== undefined && value > a
    case 'valueGreaterThanOrEqual': return value !== undefined && value >= a
    case 'valueLessThan': return value !== undefined && value < a
    case 'valueLessThanOrEqual': return value !== undefined && value <= a
    case 'valueBetween': return value !== undefined && b !== undefined && Number.isFinite(b) && value >= a && value <= b
    case 'valueNotBetween': return value !== undefined && b !== undefined && Number.isFinite(b) && (value < a || value > b)
  }
}
export function matchesDateFilter(filter: PivotFilter, value: PivotValue | undefined): boolean | undefined {
  if (!isDateFilter(filter)) return undefined
  // Excel's date-filter caption and day-boundary semantics need an oracle for blanks,
  // serial numbers and time-bearing values. Only ISO midnight cache dates are settled.
  if (value?.type !== 'date') return undefined
  const midnight = (iso: string) => /^\d{4}-\d\d-\d\d(?:T00:00:00(?:\.0+)?(?:Z|[+-]00:00)?)?$/u.test(iso)
  if (!midnight(value.value)) return undefined
  const v = asNumber(value), aValue = criterion(filter, filter.autoFilter.columns[0], 1), bValue = criterion(filter, filter.autoFilter.columns[0], 2)
  if (v === undefined || aValue === undefined || !midnight(aValue) || (bValue !== undefined && !midnight(bValue))) return undefined
  const a = Date.parse(aValue.endsWith('Z') ? aValue : `${aValue}Z`)
  const b = bValue === undefined ? undefined : Date.parse(bValue.endsWith('Z') ? bValue : `${bValue}Z`)
  if (!Number.isFinite(a)) return undefined
  switch (filter.type) {
    case 'dateEqual': return v === a
    case 'dateNotEqual': return v !== a
    case 'dateOlderThan': return v < a
    case 'dateOlderThanOrEqual': return v <= a
    case 'dateNewerThan': return v > a
    case 'dateNewerThanOrEqual': return v >= a
    case 'dateBetween': return b !== undefined && Number.isFinite(b) && v >= a && v <= b
    case 'dateNotBetween': return b !== undefined && Number.isFinite(b) && (v < a || v > b)
  }
}
