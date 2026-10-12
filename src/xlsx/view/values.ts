/**
 * Value authority (spec §4.4, plan S1.3): the single place that decides what a
 * populated record means for display, sort, filter, search and copy.
 *
 * A document has exactly one value policy, fixed at admission and never mixed:
 * - 'engine' (small object documents): the existing evaluator has already
 *   published `cell.value`, the file cache and the error status; this module
 *   only reads them, unchanged.
 * - 'cached' (paged documents and saved-values fallbacks): no evaluator runs.
 *   A formula's file cache is authoritative (provenance 'cached'); a formula
 *   without one is permanently `unavailable` — never blank, zero or an error.
 *
 * Both source adapters build their cells through `valueAuthority`/`laneValue`,
 * so object/paged parity under one policy is structural, not coincidental.
 */
import type { XlsxDocument } from '../types'
import { workbookValuePolicy } from '../formula/value-policy'
import {
  LANE_BLANK, LANE_BOOLEAN, LANE_ERROR, LANE_NUMBER, LANE_STRING, LANE_UNAVAILABLE,
  type FilterSpec, type OperationRefusal, type ReadOptions, type SourceFailure, type SourceStore, type ValuePolicy,
} from './contracts'
import type { SourceCell, SourceCol, SourceRect, SourceRow } from './coordinates'

export type { ValuePolicy }
export type ValueProvenance = ValuePolicy
export const FORMULA_VALUE_UNAVAILABLE = 'formula-value-unavailable' as const
/** Spec §4.4 wording, used while the formula count is not yet known. */
export const SAVED_VALUES_STATUS = 'Showing saved values; formulas are not recalculated'
/** Worker diagnostic code carrying the status on open; the same code the
 * loader uses for its document notice, so a consumer renders one of them. */
export const SAVED_VALUES_DIAGNOSTIC = 'cached-value-policy'

/** The raw fields either adapter holds for one populated record. */
export interface RawCellValue {
  value: string | number | boolean | null
  hasCachedValue?: boolean
  valueIsError?: boolean
  formula?: string
  sharedFormula?: unknown
}
export interface ValueAuthority {
  value: string | number | boolean | null
  provenance: ValueProvenance
  valueIsError: boolean
  unavailable: boolean
}
/** Shared-formula followers carry no formula text but are formulas all the same. */
export const hasFormula = (cell: RawCellValue): boolean => cell.formula !== undefined || cell.sharedFormula !== undefined
/** Never inferred from cache flags under 'engine': the engine already published a value. */
export const isValueUnavailable = (cell: RawCellValue, policy: ValuePolicy): boolean =>
  policy === 'cached' && hasFormula(cell) && cell.hasCachedValue !== true
export function valueAuthority(cell: RawCellValue, policy: ValuePolicy): ValueAuthority {
  // An unavailable value is withheld: no error identity can leak from an
  // uncached t="e" formula into comparators.
  if (isValueUnavailable(cell, policy)) return { value: null, provenance: policy, valueIsError: false, unavailable: true }
  return { value: cell.value, provenance: policy, valueIsError: cell.valueIsError === true, unavailable: false }
}

export interface LaneValue { tag: number; num: number; text: string }
/** Typed-lane classification (also the value index's cell-entry rule): an
 * error token is never confused with error-looking text, a non-finite number
 * is #NUM!, and text keeps its UTF-16 code units verbatim. */
export function laneValue(value: string | number | boolean | null | undefined, valueIsError: boolean | undefined, unavailable: boolean): LaneValue {
  if (unavailable) return { tag: LANE_UNAVAILABLE, num: 0, text: '' }
  if (valueIsError) return { tag: LANE_ERROR, num: 0, text: String(value) }
  if (value === null || value === undefined) return { tag: LANE_BLANK, num: 0, text: '' }
  if (typeof value === 'number') return Number.isFinite(value) ? { tag: LANE_NUMBER, num: value, text: '' } : { tag: LANE_ERROR, num: 0, text: '#NUM!' }
  if (typeof value === 'boolean') return { tag: LANE_BOOLEAN, num: value ? 1 : 0, text: '' }
  return { tag: LANE_STRING, num: 0, text: value }
}

/** The policy a parsed document was produced under (parseXlsx records 'cached'). */
export function documentValuePolicy(doc: XlsxDocument): ValuePolicy {
  return workbookValuePolicy(doc)
}
/** Functions that force evaluation on load whatever the cache says (design note §4.3 rule 1). */
const LOAD_VOLATILE = /(^|[^A-Za-z0-9_.])(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/i
/** Formula text kept for the volatile scan (twice Excel's 8,192 cap); longer text is treated as volatile (fail-closed). */
export const MAX_LOAD_FORMULA_CHARS = 16_384
export interface LoadFormula {
  /** The file holds a saved value for the formula. */
  cached: boolean
  /** `ca="1"` on the cell or its formula. */
  calcAlways: boolean
  /** Formula text; undefined for a shared-formula follower (no text of its own). */
  text: string | undefined
}
/**
 * Would the engine evaluate this formula at load? Uncached, calculate-always,
 * volatile, or every formula under fullCalcOnLoad. One rule for admission's
 * evaluation reserve and for N in the saved-values status: the formulas a
 * cached document actually loses.
 */
export function isLoadEvaluated(formula: LoadFormula, fullCalcOnLoad: boolean): boolean {
  if (fullCalcOnLoad || !formula.cached || formula.calcAlways) return true
  const text = formula.text
  return text !== undefined && (text.length > MAX_LOAD_FORMULA_CHARS || LOAD_VOLATILE.test(text))
}
export function countLoadEvaluatedFormulas(doc: XlsxDocument): number {
  const fullCalcOnLoad = doc.calc?.fullCalcOnLoad === true
  let n = 0
  for (const sheet of doc.sheets) for (const row of sheet.rows) for (const cell of row.cells) {
    if (hasFormula(cell) && isLoadEvaluated({ cached: cell.hasCachedValue === true, calcAlways: cell.ca === true, text: cell.formula }, fullCalcOnLoad)) n++
  }
  return n
}
/** Status line / accessible text: only cached documents have one. With a known
 * count it is the user-approved wording (plan §S9 decision 2). */
export function savedValuesStatus(policy: ValuePolicy, formulas?: number): string | undefined {
  if (policy !== 'cached') return undefined
  return formulas === undefined ? SAVED_VALUES_STATUS : `Showing saved values — ${formulas} formulas not recalculated (file too large)`
}
export interface ValueStatus {
  policy: ValuePolicy
  /** Formulas the engine would have evaluated at load (`isLoadEvaluated`); undefined while a paged import still runs. */
  formulas: number | undefined
  message: string | undefined
  /** The paged import ended without completing: `formulas` will never arrive. */
  ended?: boolean
}
export const valueStatus = (policy: ValuePolicy, formulas: number | undefined): ValueStatus =>
  ({ policy, formulas, message: savedValuesStatus(policy, formulas) })

/** Columns × body rows whose values an operation compares (spec §§4.4, 6.1). */
export interface ValueDomain { cols: readonly SourceCol[]; r0: SourceRow; r1: SourceRow }
export type ValueDomainScan =
  | { kind: 'clear' }
  | { kind: 'refused'; refusal: Extract<OperationRefusal, { reason: 'formula-value-unavailable' }> }
  | { kind: 'incomplete'; missing: readonly SourceRect[] }
  | { kind: 'error'; error: SourceFailure }
  | { kind: 'cancelled' }
/** The filter's value domain: its rule columns over the body band (headers and
 * totals are outside it). null when the band or the rule set is empty. */
export function filterValueDomain(spec: FilterSpec): ValueDomain | null {
  const r0 = spec.range.r0 + spec.range.headerRows, r1 = spec.range.r1 - spec.range.totalsRows
  if (r0 > r1 || spec.columns.size === 0) return null
  return { cols: [...spec.columns.keys()].sort((a, b) => a - b), r0, r1 }
}

const SCAN_PAGE_LIMIT = 4096
/**
 * Counts unavailable values inside one domain. Unavailable cells elsewhere
 * never block; an unknown page can never authorize an operation, so it is
 * reported incomplete rather than guessed. First address is row-major.
 */
export async function scanValueDomain(
  store: SourceStore, domain: ValueDomain, options: ReadOptions,
  /** Cooperative yield between pages, charged with the entries just scanned. */
  cooperate?: (units: number) => Promise<void>,
): Promise<ValueDomainScan> {
  let count = 0
  let first: SourceCell | null = null
  const note = (row: number, col: number) => {
    count++
    if (first === null || row < first.row || (row === first.row && col < first.col)) first = { row, col }
  }
  for (const col of domain.cols) {
    let cursor: string | undefined
    const seen = new Set<string>()
    do {
      if (options.signal.aborted) return { kind: 'cancelled' }
      const page = await store.queryColumn({ col, rows: { kind: 'intervals', intervals: [{ r0: domain.r0, r1: domain.r1 }] }, limit: SCAN_PAGE_LIMIT, cursor }, options)
      if (options.signal.aborted) return { kind: 'cancelled' }
      if (page.error) return page.error.code === 'cancelled' ? { kind: 'cancelled' } : { kind: 'error', error: page.error }
      const stamp = page.stamp, owner = options.stamp
      if (stamp.docId !== owner.docId || stamp.sheetId !== owner.sheetId || stamp.valueGen !== owner.valueGen) {
        return { kind: 'error', error: { code: 'stale-source', retryable: true, message: 'Source owner changed' } }
      }
      for (const item of page.items) {
        switch (item.kind) {
          case 'unavailable': note(item.row, col); break
          case 'unknown': return { kind: 'incomplete', missing: [{ r0: item.r0, r1: item.r1, c0: col, c1: col }] }
          case 'error': return { kind: 'error', error: { code: item.code, retryable: item.retryable, message: 'Value domain source failed' } }
          case 'lanes':
            for (let i = 0; i < item.tags.length; i++) if (item.tags[i] === LANE_UNAVAILABLE) note(item.rows[i], col)
            break
          case 'cell': case 'blank': break
        }
      }
      if (page.completeness !== 'complete') return { kind: 'incomplete', missing: [{ r0: domain.r0, r1: domain.r1, c0: col, c1: col }] }
      await cooperate?.(page.items.length)
      cursor = page.nextCursor
      if (cursor !== undefined) {
        if (seen.has(cursor)) return { kind: 'error', error: { code: 'invalid-cursor', retryable: false, message: 'Value domain cursor made no progress' } }
        seen.add(cursor)
      }
    } while (cursor !== undefined)
  }
  return first === null ? { kind: 'clear' } : { kind: 'refused', refusal: { reason: FORMULA_VALUE_UNAVAILABLE, count, firstAddress: first } }
}
