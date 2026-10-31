/** Pure filter predicates shared by computeAxes. */
import type { PivotFilter, PivotFilterColumn, PivotFilterType, PivotValue } from '../model'
import { compareEnglishLabelKeys, englishLabelKey, hasMeasuredEnglishOrder, type EnglishLabelKey } from './sort'
import { tokenizeWildcardPattern } from '../../formula/functions/wildcard'

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
/** The whole filter's wildcard DP work (Σ over items of (tokens + 1) × (label + 1)); past it the filter refuses as
 * `pivot-values-limit` rather than computing a partial result (perf audit item 1). */
export const LABEL_FILTER_DP_BUDGET = 50_000_000
/** The shared matcher's per-call limits (formula/functions/wildcard.ts), kept per item. */
const MAX_PATTERN_LENGTH = 8192, MAX_DP_CELLS = 4_000_000

const printableAscii = (text: string) => /^[\x20-\x7e]*$/u.test(text)
const isAscii = (text: string) => { for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) >= 0x80) return false; return true }
const foldCode = (code: number) => code >= 65 && code <= 90 ? code + 32 : code

/** One label filter compiled once per filter: its criteria validated, its `<filters>` list as a lower-cased Set,
 * its wildcard patterns tokenized and its ordered operands keyed. `match` returns `undefined` for an undecided
 * item (unmeasured) and `'limit'` once the filter's DP budget is spent. */
export interface LabelFilterMatcher { match(value: PivotValue | undefined): boolean | undefined | 'limit' }

type Mode = 'exact' | 'prefix' | 'contains' | 'suffix'
const ORDERED = new Set<PivotFilterType>(['captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan',
  'captionLessThanOrEqual', 'captionBetween', 'captionNotBetween'])
const RULE_OPS: Record<string, PivotFilterType> = {
  equal: 'captionEqual', notEqual: 'captionNotEqual', greaterThan: 'captionGreaterThan', greaterThanOrEqual: 'captionGreaterThanOrEqual',
  lessThan: 'captionLessThan', lessThanOrEqual: 'captionLessThanOrEqual',
}

interface Budget { left: number }
/**
 * Excel wildcard match (`?`, `*`, `~` escapes), the shared matcher's semantics with the pattern tokenized once: no
 * RegExp is built from workbook text, and the DP reuses two rows. The pattern is validated alone first, so a
 * dangling `~` refuses instead of escaping the star a mode appends. `undefined` = not decided (undocumented tilde
 * form, non-ASCII, or the per-item work limit).
 */
function compileWildcard(pattern: string, mode: Mode): (label: string, budget: Budget) => boolean | undefined | 'limit' {
  const full = mode === 'prefix' ? `${pattern}*` : mode === 'suffix' ? `*${pattern}` : mode === 'contains' ? `*${pattern}*` : pattern
  const alone = tokenizeWildcardPattern(pattern), tokens = tokenizeWildcardPattern(full)
  if (!Array.isArray(alone) || !Array.isArray(tokens) || full.length > MAX_PATTERN_LENGTH || !isAscii(full)) return () => undefined
  // Token kinds: -1 any run, -2 any one, else the folded literal code.
  const kinds = Int32Array.from(tokens, token => token.t === 'anyRun' ? -1 : token.t === 'anyOne' ? -2 : foldCode(token.ch.charCodeAt(0)))
  let row = new Uint8Array(0), next = new Uint8Array(0)
  return (label, budget) => {
    if (!isAscii(label)) return undefined
    const n = label.length, cells = (kinds.length + 1) * (n + 1)
    if (cells > MAX_DP_CELLS) return undefined
    if ((budget.left -= cells) < 0) return 'limit'
    if (row.length < n + 1) { row = new Uint8Array(n + 1); next = new Uint8Array(n + 1) }
    row.fill(0, 0, n + 1); row[0] = 1
    for (let k = 0; k < kinds.length; k++) {
      const kind = kinds[k]!
      next.fill(0, 0, n + 1)
      if (kind === -1) {
        next[0] = row[0]!
        for (let j = 0; j < n; j++) next[j + 1] = (row[j + 1] || next[j]) ? 1 : 0
      } else for (let j = 0; j < n; j++) {
        if (!row[j]) continue
        if (kind === -2 || kind === foldCode(label.charCodeAt(j))) next[j + 1] = 1
      }
      const swap = row; row = next; next = swap
    }
    return row[n] === 1
  }
}

/** A caption comparison with its operands prepared once (captionCompare before the perf round). */
function compileCaptionCompare(op: PivotFilterType, a: string, b: string | undefined,
  keyOf: (label: string) => EnglishLabelKey): (label: string, budget: Budget) => boolean | undefined | 'limit' {
  if (ORDERED.has(op)) {
    if (!hasMeasuredEnglishOrder(a) || (b !== undefined && !hasMeasuredEnglishOrder(b))) return () => undefined
    const first = englishLabelKey(a), second = b === undefined ? undefined : englishLabelKey(b)
    return label => {
      if (!hasMeasuredEnglishOrder(label)) return undefined
      const key = keyOf(label)
      const firstVsLabel = compareEnglishLabelKeys(key, first), secondVsLabel = second === undefined ? undefined : compareEnglishLabelKeys(key, second)
      switch (op) {
        case 'captionGreaterThan': return firstVsLabel > 0
        case 'captionGreaterThanOrEqual': return firstVsLabel >= 0
        case 'captionLessThan': return firstVsLabel < 0
        case 'captionLessThanOrEqual': return firstVsLabel <= 0
        case 'captionBetween': return secondVsLabel !== undefined && firstVsLabel >= 0 && secondVsLabel <= 0
        default: return secondVsLabel !== undefined && (firstVsLabel < 0 || secondVsLabel > 0)   // captionNotBetween
      }
    }
  }
  const glob = (mode: Mode, negate: boolean) => {
    const matcher = compileWildcard(a, mode)
    return (label: string, budget: Budget) => {
      const matched = matcher(label, budget)
      return matched === undefined || matched === 'limit' ? matched : matched !== negate
    }
  }
  switch (op) {
    case 'captionEqual': return glob('exact', false)
    case 'captionNotEqual': return glob('exact', true)
    case 'captionBeginsWith': return glob('prefix', false)
    case 'captionNotBeginsWith': return glob('prefix', true)
    case 'captionEndsWith': return glob('suffix', false)
    case 'captionNotEndsWith': return glob('suffix', true)
    case 'captionContains': return glob('contains', false)
    case 'captionNotContains': return glob('contains', true)
    default: return () => false
  }
}

export function compileLabelFilter(filter: PivotFilter): LabelFilterMatcher {
  if (!isLabelFilter(filter)) return { match: () => undefined }
  const column = filter.autoFilter.columns[0]
  const negated = filter.type === 'captionNotEqual' || filter.type === 'captionNotContains' || filter.type === 'captionNotBeginsWith' ||
    filter.type === 'captionNotEndsWith' || filter.type === 'captionNotBetween'
  const budget: Budget = { left: LABEL_FILTER_DP_BUDGET }
  // The item's collation key, shared by every rule that compares it.
  let lastLabel: string | undefined, lastKey: EnglishLabelKey | undefined
  const keyOf = (label: string) => {
    if (label !== lastLabel || !lastKey) { lastLabel = label; lastKey = englishLabelKey(label) }
    return lastKey
  }
  // Decides a printable label; the shared preamble below handles blanks, types and printability.
  let decide: (label: string) => boolean | undefined | 'limit'
  if (column?.filters) {
    const values = column.filters.values
    const valid = values.every(printableAscii)
    const lower = new Set(values.map(candidate => candidate.toLowerCase()))
    decide = label => valid ? lower.has(label.toLowerCase()) : undefined
  } else {
    const a = criterion(filter, column, 1), b = criterion(filter, column, 2)
    const custom = column?.customFilters?.filters ?? []
    // The schema allows at most two customFilter rules (CT_CustomFilters); more is not a file Excel writes.
    if (custom.length > 2 || (!custom.length && a === undefined) || (a !== undefined && !printableAscii(a)) || (b !== undefined && !printableAscii(b))) decide = () => undefined
    else if (custom.length) {
      const rules = custom.map(rule => rule.val === undefined || !printableAscii(rule.val) ? undefined
        : compileCaptionCompare(RULE_OPS[rule.operator] ?? 'captionEqual', rule.val, undefined, keyOf))
      const and = !!column!.customFilters!.and
      decide = label => {
        if (rules.some(rule => rule === undefined)) return undefined
        // Every rule decides first (an undecided rule makes the item undecided), as before.
        let any = false, all = true
        for (const rule of rules) {
          const decision = rule!(label, budget)
          if (decision === undefined || decision === 'limit') return decision
          any ||= decision; all &&= decision
        }
        return and ? all : any
      }
    } else {
      const compare = compileCaptionCompare(filter.type, a!, b, keyOf)
      decide = label => compare(label, budget)
    }
  }
  return {
    match(value) {
      if (!value || value.type === 'missing') return negated ? undefined : false
      if (value.type !== 'string' && value.type !== 'number') return undefined
      const label = value.type === 'string' ? value.value : String(value.value)
      if (!printableAscii(label)) return undefined
      return decide(label)
    },
  }
}

/** One item against one label filter (compiles the filter; computeAxes compiles once per filter instead). */
export function matchesLabelFilter(filter: PivotFilter, value: PivotValue | undefined): boolean | undefined {
  const decision = compileLabelFilter(filter).match(value)
  return decision === 'limit' ? undefined : decision
}

function asNumber(value: PivotValue | undefined): number | undefined {
  if (value?.type === 'number') return value.value
  if (value?.type === 'date') {
    const time = Date.parse(value.value.endsWith('Z') ? value.value : `${value.value}Z`)
    return Number.isFinite(time) ? time / 86_400_000 + 25_569 : undefined
  }
  return undefined
}
/** A value filter's numeric cuts (one, or two for between and not between). */
export function valueFilterCuts(filter: PivotFilter): number[] {
  const a = Number(criterion(filter, filter.autoFilter.columns[0], 1)), bRaw = criterion(filter, filter.autoFilter.columns[0], 2)
  return bRaw === undefined ? [a] : [a, Number(bRaw)]
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
  if (value?.type === 'missing') return false
  if (value?.type !== 'date') return undefined
  const midnight = (iso: string) => /^\d{4}-\d\d-\d\d(?:T00:00:00(?:\.0+)?(?:Z|[+-]00:00)?)?$/u.test(iso)
  if (!midnight(value.value)) return undefined
  const v = asNumber(value), aValue = criterion(filter, filter.autoFilter.columns[0], 1), bValue = criterion(filter, filter.autoFilter.columns[0], 2)
  const serial = (text: string): number => {
    if (/^\d+(?:\.\d+)?$/u.test(text)) return Number(text)
    const time = Date.parse(text.endsWith('Z') ? text : `${text}Z`)
    return Number.isFinite(time) ? time / 86_400_000 + 25_569 : Number.NaN
  }
  if (v === undefined || aValue === undefined || (bValue !== undefined && !Number.isFinite(serial(bValue)))) return undefined
  // Criteria here use the 1900 serial epoch; computePivot refuses date filters in a 1904 workbook.
  const a = serial(aValue)
  const b = bValue === undefined ? undefined : serial(bValue)
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
