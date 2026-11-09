/**
 * Measured Stage-1 predicates (Task 5).
 *
 * `evaluatePredicate` decides one rule on one cell from the normalized
 * definition: literal `cellIs` comparisons per the Task-1 matrix, canonical
 * text/blank/error predicates, and statistics-backed top/bottom-N,
 * above/below-average and duplicate/unique checks. Anything the oracle has
 * not measured — blank-cell comparison semantics, mixed-type relational
 * operators, error-literal operands (unrepresentable in the Stage-3 grammar),
 * unmeasured clock/timePeriod evaluation, unmeasured rank/stdDev edges —
 * returns unresolved instead of guessing. The canonical blank formula
 * additionally passes through the evidence gate and stays unresolved until
 * `stage1-blanks-errors` is measured.
 */
import type { CfLiteral, CfOperator, CfOutcome, CfRule, CfTimePeriod } from './model'
import { hasEvidence } from './evidence'
import { cfValueKey, type CfRangeStatistics } from './statistics'
import { decodeSerial, epochToSerial, weekdayOf } from '../formula/serial'

/** Evidence gate: production reads the reviewed allowlist; tests inject fakes. */
export interface CfEvidenceGate {
  allows(id: string, variant?: string): boolean
}

export const cfRuntimeEvidence: CfEvidenceGate = {
  allows: (id, variant) => hasEvidence(id, variant),
}

export interface CfDateContext {
  readonly dateSystem?: '1900' | '1904'
  readonly timeZone?: string
  readonly epochNowMs?: number
  readonly clockSerial?: number
}

export interface CfSourceAddress {
  readonly row: number
  readonly col: number
}

const unmeasured = (extra?: string): CfOutcome<boolean> => ({
  kind: 'unresolved',
  reason: extra ?? 'cf-unmeasured-capability',
})

function inRanges(row: number, col: number, rule: CfRule): boolean {
  for (const r of rule.ranges) {
    if (row >= r.firstRow && row <= r.lastRow && col >= r.firstCol && col <= r.lastCol) return true
  }
  return false
}

function compareNumbers(operator: CfOperator, value: number, operands: readonly CfLiteral[]): boolean | undefined {
  const nums: number[] = []
  for (const operand of operands) {
    if (operand.kind !== 'number') return undefined
    nums.push(operand.value)
  }
  const [a, b] = nums
  switch (operator) {
    case 'equal': return value === a
    case 'notEqual': return value !== a
    case 'greaterThan': return value > a!
    case 'lessThan': return value < a!
    case 'greaterThanOrEqual': return value >= a!
    case 'lessThanOrEqual': return value <= a!
    case 'between': {
      const lo = Math.min(a!, b!)
      const hi = Math.max(a!, b!)
      return value >= lo && value <= hi
    }
    case 'notBetween': {
      const lo = Math.min(a!, b!)
      const hi = Math.max(a!, b!)
      return value < lo || value > hi
    }
  }
}

function evaluateCellIs(
  operator: CfOperator,
  operands: readonly CfLiteral[],
  value: CfLiteral,
  evidence: CfEvidenceGate,
): CfOutcome<boolean> {
  // Numeric cells against numeric operands: the settled comparison matrix.
  if (value.kind === 'number') {
    if (operator === 'between' || operator === 'notBetween') {
      if (operands.length !== 2) return unmeasured('cf-unsupported-rule')
      // Reversed operands normalise to [min,max] only with evidence.
      if (operands[0]!.kind !== 'number' || operands[1]!.kind !== 'number') return unmeasured()
      if (operands[0]!.value > operands[1]!.value) {
        const allowed = operator === 'between'
          ? evidence.allows('cellIs:between-reversed')
          : evidence.allows('cellIs:notBetween-reversed')
        if (!allowed) return unmeasured()
      }
      const decided = compareNumbers(operator, value.value, operands)
      return decided === undefined ? unmeasured() : decided ? { kind: 'ready', value: true } : { kind: 'no-match' }
    }
    if (operands.length !== 1) return unmeasured('cf-unsupported-rule')
    const [operand] = operands as [CfLiteral]
    if (operand!.kind === 'number') {
      const decided = compareNumbers(operator, value.value, operands)
      return decided === undefined ? unmeasured() : decided ? { kind: 'ready', value: true } : { kind: 'no-match' }
    }
    // Cross-type equality is decided false only where the manifest rules it
    // (B8/L6 strings, N6 error); every other cross-type combination is an
    // unmeasured missing case.
    if (operator === 'equal') {
      if (operand!.kind === 'string' || operand!.kind === 'error') {
        return { kind: 'no-match' }
      }
      return unmeasured()
    }
    return unmeasured()
  }
  // Error cells: only the manifest-decided numeric-operand equality is settled.
  if (value.kind === 'error') {
    if (operator === 'equal' && operands.length === 1 && operands[0]!.kind === 'number') {
      return { kind: 'no-match' }
    }
    return unmeasured()
  }
  if (value.kind === 'string') {
    if (operator === 'equal' && operands.length === 1 && operands[0]!.kind === 'string') {
      return operands[0]!.value === value.value ? { kind: 'ready', value: true } : { kind: 'no-match' }
    }
    return unmeasured()
  }
  if (value.kind === 'boolean') {
    if (operator === 'equal' && operands.length === 1 && operands[0]!.kind === 'boolean') {
      return operands[0]!.value === value.value ? { kind: 'ready', value: true } : { kind: 'no-match' }
    }
    return unmeasured()
  }
  // Unstored blank cells: blank comparison semantics are unmeasured.
  return unmeasured()
}

function foldCase(text: string): string {
  return text.toLowerCase()
}

function evaluateText(
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith',
  needle: string,
  value: CfLiteral,
  evidence: CfEvidenceGate,
): CfOutcome<boolean> {
  if (needle.length === 0) return unmeasured()
  // notContainsText on numeric cells and blanks is explicitly unmeasured.
  if (type === 'notContainsText' && (value.kind === 'number' || value.kind === 'blank')) return unmeasured()
  let haystack: string | undefined
  if (value.kind === 'string') haystack = value.value
  else if (value.kind === 'number') haystack = String(value.value)
  else return unmeasured()
  // Exact matches are decided regardless of case policy; fold-only matches
  // require the measured case-insensitive semantics.
  const foldedHay = foldCase(haystack)
  const foldedNeedle = foldCase(needle)
  const exact = (() => {
    switch (type) {
      case 'containsText': return haystack.includes(needle)
      case 'notContainsText': return haystack.includes(needle)
      case 'beginsWith': return haystack.startsWith(needle)
      case 'endsWith': return haystack.endsWith(needle)
    }
  })()
  const folded = (() => {
    switch (type) {
      case 'containsText': return foldedHay.includes(foldedNeedle)
      case 'notContainsText': return foldedHay.includes(foldedNeedle)
      case 'beginsWith': return foldedHay.startsWith(foldedNeedle)
      case 'endsWith': return foldedHay.endsWith(foldedNeedle)
    }
  })()
  // Exact hits and exact misses are decided regardless of case policy; only
  // fold-only hits need the measured case-insensitive semantics.
  if (exact) return type === 'notContainsText' ? { kind: 'no-match' } : { kind: 'ready', value: true }
  if (folded) {
    const allowed = (() => {
      switch (type) {
        case 'containsText': return evidence.allows('text:containsText-canonical')
        case 'notContainsText': return evidence.allows('text:notContainsText-canonical')
        case 'beginsWith': return evidence.allows('text:beginsWith-canonical')
        case 'endsWith': return evidence.allows('text:endsWith-canonical')
      }
    })()
    if (!allowed) return unmeasured()
    return type === 'notContainsText' ? { kind: 'no-match' } : { kind: 'ready', value: true }
  }
  return type === 'notContainsText' ? { kind: 'ready', value: true } : { kind: 'no-match' }
}

function evaluateBlanks(
  type: 'containsBlanks' | 'notContainsBlanks',
  value: CfLiteral,
): CfOutcome<boolean> {
  if (value.kind === 'blank') return type === 'containsBlanks' ? { kind: 'ready', value: true } : { kind: 'no-match' }
  if (value.kind === 'string' && value.value === '') {
    return type === 'containsBlanks' ? { kind: 'ready', value: true } : { kind: 'no-match' }
  }
  if (value.kind === 'string' && /^\s+$/.test(value.value)) {
    // Whitespace-only strings: TRIM semantics are unmeasured without a formula.
    return unmeasured()
  }
  if (value.kind === 'string') {
    // Ordinary nonempty text is non-blank.
    return type === 'containsBlanks' ? { kind: 'no-match' } : { kind: 'ready', value: true }
  }
  if (value.kind === 'number' || value.kind === 'boolean' || value.kind === 'error') {
    return type === 'containsBlanks' ? { kind: 'no-match' } : { kind: 'ready', value: true }
  }
  // Whitespace-only strings: TRIM semantics are unmeasured without a formula.
  return unmeasured()
}

function evaluateBlanksGated(
  type: 'containsBlanks' | 'notContainsBlanks',
  hasCanonicalFormula: boolean,
  value: CfLiteral,
  evidence: CfEvidenceGate,
): CfOutcome<boolean> {
  if (!hasCanonicalFormula) return evaluateBlanks(type, value)
  const allowed = type === 'containsBlanks'
    ? evidence.allows('blanks:containsBlanks-canonical')
    : evidence.allows('blanks:notContainsBlanks-canonical')
  if (!allowed) {
    return { kind: 'unresolved', reason: 'cf-unmeasured-capability' }
  }
  return evaluateBlanks(type, value)
}

function evaluateErrors(
  type: 'containsErrors' | 'notContainsErrors',
  value: CfLiteral,
): CfOutcome<boolean> {
  if (value.kind === 'error') {
    return type === 'containsErrors' ? { kind: 'ready', value: true } : { kind: 'no-match' }
  }
  if (value.kind === 'blank') return unmeasured()
  if (value.kind === 'number' || value.kind === 'string' || value.kind === 'boolean') {
    return type === 'containsErrors' ? { kind: 'no-match' } : { kind: 'ready', value: true }
  }
  return unmeasured()
}

function needStats(stats: CfRangeStatistics | undefined): CfOutcome<CfRangeStatistics> {
  if (stats === undefined) return { kind: 'unresolved', reason: 'cf-missing-statistics' }
  return { kind: 'ready', value: stats }
}

function evaluateTop(
  rank: number | undefined,
  bottom: boolean,
  percent: boolean,
  value: CfLiteral,
  stats: CfRangeStatistics | undefined,
): CfOutcome<boolean> {
  if (rank === undefined || !Number.isInteger(rank)) return unmeasured('cf-unmeasured-capability')
  if (value.kind !== 'number') return { kind: 'no-match' }
  const settled = needStats(stats)
  if (settled.kind !== 'ready') return settled
  const st = settled.value
  if (st.numericCount === 0) return { kind: 'no-match' }
  // Ruling I7: ECMA does not specify percent rounding. Math.ceil((count * rank) / 100)
  // is inferred from the Excel UI preset and unverified beyond 5 values.
  const need = percent ? Math.ceil((st.numericCount * rank) / 100) : Math.min(rank, st.numericCount)
  if (need <= 0) return { kind: 'no-match' }
  const cutoff = bottom ? st.sortedAsc[need - 1]! : st.sortedAsc[st.numericCount - need]!
  const matched = bottom ? value.value <= cutoff : value.value >= cutoff
  return matched ? { kind: 'ready', value: true } : { kind: 'no-match' }
}

function evaluateAboveAverage(
  above: boolean,
  equal: boolean,
  stdDev: number | undefined,
  value: CfLiteral,
  stats: CfRangeStatistics | undefined,
  evidence: CfEvidenceGate,
): CfOutcome<boolean> {
  if (stdDev !== undefined && equal) return { kind: 'unresolved', reason: 'cf-invalid-rule' }
  // Population-vs-sample is unmeasured: any stddev use needs evidence.
  if (stdDev !== undefined && !evidence.allows('aboveAverage:stdDev-1-pop-vs-sample')) {
    return unmeasured()
  }
  // Hidden-row inclusion/exclusion in range statistics is unmeasured: ranges
  // crossing known hidden rows require evidence.
  if (stats?.hasHiddenRows && !evidence.allows('layout:hidden-row-statistics')) {
    return unmeasured()
  }
  if (value.kind !== 'number') return { kind: 'no-match' }
  const settled = needStats(stats)
  if (settled.kind !== 'ready') return settled
  const st = settled.value
  if (st.mean === undefined || st.numericCount === 0) return { kind: 'no-match' }
  let threshold = st.mean
  if (stdDev !== undefined) {
    // Ruling I2: Excel STDEV.S (sample stddev) only; never fall back to population stddev.
    // When range contains errors or sample stddev is undefined, fail closed (no-match).
    const sd = st.stddevSample
    if (sd === undefined) return { kind: 'no-match' }
    threshold = above ? st.mean + stdDev * sd : st.mean - stdDev * sd
  }
  let matched: boolean
  if (above && !equal) matched = value.value > threshold
  else if (above && equal) matched = value.value >= threshold
  else if (!above && !equal) matched = value.value < threshold
  else matched = value.value <= threshold
  return matched ? { kind: 'ready', value: true } : { kind: 'no-match' }
}

function evaluateDuplicates(
  unique: boolean,
  value: CfLiteral,
  stats: CfRangeStatistics | undefined,
  evidence: CfEvidenceGate,
): CfOutcome<boolean> {
  // Blank cells are never flagged, per the staged hypothesis.
  if (value.kind === 'blank') return { kind: 'no-match' }
  // Text equivalence (case folding): withhold text cells without evidence.
  if (value.kind === 'string') {
    const allowed = unique
      ? evidence.allows('uniqueValues:single-range')
      : evidence.allows('duplicateValues:single-range')
    if (!allowed) return unmeasured()
  }
  const settled = needStats(stats)
  if (settled.kind !== 'ready') return settled
  const key = cfValueKey(value)
  if (key === undefined) return { kind: 'no-match' }
  const count = settled.value.frequencies.get(key) ?? 0
  const matched = unique ? count === 1 : count > 1
  return matched ? { kind: 'ready', value: true } : { kind: 'no-match' }
}

function evaluateTimePeriod(
  period: CfTimePeriod,
  value: CfLiteral,
  evidence: CfEvidenceGate,
  context?: CfDateContext,
): CfOutcome<boolean> {
  const system = context?.dateSystem ?? '1900'
  const allowed = (() => {
    if (system === '1904') {
      switch (period) {
        case 'today': return evidence.allows('timePeriod:1904-today')
        case 'yesterday': return evidence.allows('timePeriod:1904-yesterday')
        case 'tomorrow': return evidence.allows('timePeriod:1904-tomorrow')
        case 'thisWeek': return evidence.allows('timePeriod:1904-thisWeek')
        case 'thisMonth': return evidence.allows('timePeriod:1904-thisMonth')
        default: return false
      }
    }
    switch (period) {
      case 'today': return evidence.allows('timePeriod:today')
      case 'yesterday': return evidence.allows('timePeriod:yesterday')
      case 'tomorrow': return evidence.allows('timePeriod:tomorrow')
      case 'last7Days': return evidence.allows('timePeriod:last7Days')
      case 'thisWeek': return evidence.allows('timePeriod:thisWeek')
      case 'lastWeek': return evidence.allows('timePeriod:lastWeek')
      case 'nextWeek': return evidence.allows('timePeriod:nextWeek')
      case 'thisMonth': return evidence.allows('timePeriod:thisMonth')
      case 'lastMonth': return evidence.allows('timePeriod:lastMonth')
      case 'nextMonth': return evidence.allows('timePeriod:nextMonth')
    }
  })()
  if (!allowed) return unmeasured()

  if (value.kind !== 'number' || !Number.isFinite(value.value)) return { kind: 'no-match' }
  const cellDay = Math.floor(value.value)
  if (cellDay < 0 && system === '1900') return { kind: 'no-match' }

  let todaySerial: number
  if (context?.clockSerial !== undefined && Number.isFinite(context.clockSerial)) {
    todaySerial = Math.floor(context.clockSerial)
  } else {
    const epochNowMs = context?.epochNowMs ?? Date.now()
    const timeZone = context?.timeZone ?? 'UTC'
    const s = epochToSerial(epochNowMs, timeZone, system)
    if (typeof s !== 'number' || !Number.isFinite(s)) return unmeasured()
    todaySerial = Math.floor(s)
  }

  let matched = false
  switch (period) {
    case 'today':
      matched = cellDay === todaySerial
      break
    case 'yesterday':
      matched = cellDay === todaySerial - 1
      break
    case 'tomorrow':
      matched = cellDay === todaySerial + 1
      break
    case 'last7Days':
      matched = cellDay <= todaySerial && cellDay >= todaySerial - 6
      break
    case 'thisWeek': {
      const w = weekdayOf(todaySerial, 1, system)
      if (typeof w !== 'number' || !Number.isFinite(w)) return unmeasured()
      const sun = todaySerial - (w - 1)
      matched = cellDay >= sun && cellDay <= sun + 6
      break
    }
    case 'lastWeek': {
      const w = weekdayOf(todaySerial, 1, system)
      if (typeof w !== 'number' || !Number.isFinite(w)) return unmeasured()
      const sun = todaySerial - (w - 1) - 7
      matched = cellDay >= sun && cellDay <= sun + 6
      break
    }
    case 'nextWeek': {
      const w = weekdayOf(todaySerial, 1, system)
      if (typeof w !== 'number' || !Number.isFinite(w)) return unmeasured()
      const sun = todaySerial - (w - 1) + 7
      matched = cellDay >= sun && cellDay <= sun + 6
      break
    }
    case 'thisMonth': {
      const cellDec = decodeSerial(cellDay, system)
      const todayDec = decodeSerial(todaySerial, system)
      if ('kind' in cellDec || 'kind' in todayDec) return unmeasured()
      matched = cellDec.civil.month === todayDec.civil.month && cellDec.civil.year === todayDec.civil.year
      break
    }
    case 'lastMonth': {
      const cellDec = decodeSerial(cellDay, system)
      const todayDec = decodeSerial(todaySerial, system)
      if ('kind' in cellDec || 'kind' in todayDec) return unmeasured()
      const targetMonth = todayDec.civil.month === 1 ? 12 : todayDec.civil.month - 1
      const targetYear = todayDec.civil.month === 1 ? todayDec.civil.year - 1 : todayDec.civil.year
      matched = cellDec.civil.month === targetMonth && cellDec.civil.year === targetYear
      break
    }
    case 'nextMonth': {
      const cellDec = decodeSerial(cellDay, system)
      const todayDec = decodeSerial(todaySerial, system)
      if ('kind' in cellDec || 'kind' in todayDec) return unmeasured()
      const targetMonth = todayDec.civil.month === 12 ? 1 : todayDec.civil.month + 1
      const targetYear = todayDec.civil.month === 12 ? todayDec.civil.year + 1 : todayDec.civil.year
      matched = cellDec.civil.month === targetMonth && cellDec.civil.year === targetYear
      break
    }
  }

  return matched ? { kind: 'ready', value: true } : { kind: 'no-match' }
}

/**
 * Decide one rule on one cell. `stats` carries the prepared range statistics
 * for rules that need them (top10/aboveAverage/duplicateValues/uniqueValues);
 * `evidence` gates measurement-pending semantics. Never reads the wall clock
 * and never evaluates formulas: anything unmeasured returns unresolved.
 */
export function evaluatePredicate(
  rule: CfRule,
  sourceAddress: { readonly row: number; readonly col: number },
  value: CfLiteral,
  stats: CfRangeStatistics | undefined,
  evidence: CfEvidenceGate,
  dateContext?: CfDateContext,
): CfOutcome<boolean> {
  if (!inRanges(sourceAddress.row, sourceAddress.col, rule)) return { kind: 'no-match' }
  const definition = rule.definition
  if (definition.kind === 'unresolved') return { kind: 'unresolved', reason: definition.reason }
  switch (definition.kind) {
    case 'cellIs':
      return evaluateCellIs(definition.operator, definition.operands, value, evidence)
    case 'text':
      return evaluateText(definition.type, definition.needle, value, evidence)
    case 'predicate':
      switch (definition.type) {
        case 'containsBlanks':
        case 'notContainsBlanks': {
          const formulas = definition.formulas ?? []
          return evaluateBlanksGated(definition.type, formulas.length > 0, value, evidence)
        }
        case 'containsErrors':
        case 'notContainsErrors':
          return evaluateErrors(definition.type, value)
        case 'duplicateValues':
        case 'uniqueValues': {
          // Multi-rectangle duplicate/unique evaluation needs evidence; single rectangles are exact.
          if (rule.ranges.length !== 1) {
            const allowed = definition.type === 'duplicateValues'
              ? evidence.allows('duplicateValues:disjoint-sqref')
              : evidence.allows('uniqueValues:overlapping-sqref-no-double-count')
            if (!allowed) return unmeasured()
          }
          return evaluateDuplicates(definition.type === 'uniqueValues', value, stats, evidence)
        }
      }
      break
    case 'timePeriod':
      return evaluateTimePeriod(definition.period, value, evidence, dateContext)
    case 'top10': {
      // No manifest rows pin multi-rectangle ranking: always withheld.
      if (rule.ranges.length !== 1) return unmeasured()
      return evaluateTop(definition.rank, definition.bottom, definition.percent, value, stats)
    }
    case 'aboveAverage': {
      // No manifest rows pin multi-rectangle averaging: always withheld.
      if (rule.ranges.length !== 1) return unmeasured()
      return evaluateAboveAverage(definition.above, definition.equal, definition.stdDev, value, stats, evidence)
    }
  }
  return unmeasured()
}
