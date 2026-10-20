/**
 * Chart point states (spec 5.8; lazy chunk). Every numeric plotting input is
 * exactly one of value, blank, error(code) or unavailable, from the chart
 * part's caches and literals and from both XLSX fills (resident and paged), so
 * the same workbook yields the same arrays on every path.
 *
 * The non-value states are shared frozen objects (one per error code): a
 * sparse or error-heavy series costs one object per state, not per point, and
 * a structured clone keeps the sharing.
 */
import type { ChartPoint, ChartSeries } from './model'

export const BLANK: ChartPoint = Object.freeze({ state: 'blank' })
export const UNAVAILABLE: ChartPoint = Object.freeze({ state: 'unavailable' })

/** The worksheet error codes (Excel's set, modern codes included). */
export const CHART_ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!'] as const
/** Units kept of an error cell's code (the longest known code is 13). */
export const CHART_ERROR_CODE_LIMIT = 16
const ERRORS = new Map<string, ChartPoint>(CHART_ERROR_CODES.map(code => [code, Object.freeze({ state: 'error', code })]))

export const valuePoint = (v: number): ChartPoint => ({ state: 'value', v })
/** An error point keeps its code (#N/A distinct from the rest); known codes are shared. */
export function errorPoint(code: string): ChartPoint {
  const text = code.slice(0, CHART_ERROR_CODE_LIMIT)
  return ERRORS.get(text.toUpperCase()) ?? { state: 'error', code: text }
}
/** A known worksheet error code, or undefined (cache text that is not an error). */
export const knownError = (text: string): ChartPoint | undefined => ERRORS.get(text.trim().toUpperCase())
/** The plotted number, or undefined for every other state (today's drawing gap). */
export const pointValue = (point: ChartPoint | undefined): number | undefined => point?.state === 'value' ? point.v : undefined
/**
 * The input an analytic element computes from (spec 5.8, 7): every acquired
 * value of the role, blank and error points excluded. `refused` when the role
 * was not acquired in full (more than CHART_ANALYTIC_POINT_LIMIT points) or a
 * required point is unavailable: the element is then dropped, never computed
 * from a prefix or a partial set.
 */
export function analyticInput(series: Pick<ChartSeries, 'analytic' | 'values' | 'xValues'>, role: 'values' | 'xValues' = 'values'): { kind: 'input'; values: number[] } | { kind: 'refused'; reason: string } {
  if (series.analytic !== true) return { kind: 'refused', reason: series.analytic === 'refused' ? 'analytic-point-limit' : 'not-analytic' }
  const values: number[] = []
  for (const point of series[role] ?? []) {
    if (point.state === 'unavailable') return { kind: 'refused', reason: 'chart-point-unavailable' }
    if (point.state === 'value') values.push(point.v)
  }
  return { kind: 'input', values }
}
