/**
 * Chart point storage (spec 5.8; lazy chunk). Every numeric plotting input is
 * exactly one of value, blank, error(code) or unavailable. A role's points are
 * stored compactly (`ChartPoints`): one state byte and one float per point,
 * an error's code as an index into the known codes (or a small table of other
 * codes), optional per-point format codes, and the role's hidden-cell
 * provenance. No object is allocated per point; the `ChartPoint` union is an
 * accessor view (`pointAt`) for consumers that want one point at a time.
 */
import type { ChartPoint, ChartPoints, ChartSeries } from './model'

export const POINT_BLANK = 0, POINT_VALUE = 1, POINT_ERROR = 2, POINT_UNAVAILABLE = 3

/** The worksheet error codes (Excel's set, modern codes included). */
export const CHART_ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!'] as const
/** Units kept of an error cell's code (the longest known code is 13). */
export const CHART_ERROR_CODE_LIMIT = 16
const KNOWN = new Map<string, number>(CHART_ERROR_CODES.map((code, i) => [code, i]))
/** A known worksheet error code's index, or undefined (cache text that is not an error). */
export const knownError = (text: string): number | undefined => KNOWN.get(text.trim().toUpperCase())

/** Bytes a role of `length` points occupies (states + values, and format indices when present). */
export const pointBytes = (length: number, formats = false): number => length * (9 + (formats ? 2 : 0))
/** Bytes a string table entry is charged (UTF-16 plus overhead). */
export const stringBytes = (text: string): number => text.length * 2 + 32
/** The bytes a role's storage holds (what its owner keeps charged). */
export function storageBytes(points: ChartPoints): number {
  let bytes = points.states.byteLength + points.values.byteLength + (points.formatIndex?.byteLength ?? 0) + (points.hidden?.byteLength ?? 0)
  for (const text of points.codes ?? []) bytes += stringBytes(text)
  for (const text of points.formats ?? []) bytes += stringBytes(text)
  return bytes
}

/** Builds one role's storage in place (the caller admitted `pointBytes(length, formats)` first). */
export class PointWriter {
  readonly states: Uint8Array
  readonly values: Float64Array
  private formatIndex?: Uint16Array
  private codes?: string[]
  private formats?: string[]
  constructor(readonly length: number, withFormats = false) {
    this.states = new Uint8Array(length)
    this.values = new Float64Array(length)
    if (withFormats) this.formatIndex = new Uint16Array(length)
  }
  value(i: number, v: number, fmt?: string): void {
    this.states[i] = POINT_VALUE; this.values[i] = v
    if (fmt && this.formatIndex) {
      const formats = this.formats ??= []
      let k = formats.indexOf(fmt)
      if (k < 0 && formats.length < 0xfffe) k = formats.push(fmt) - 1
      if (k >= 0) this.formatIndex[i] = k + 1
    }
  }
  unavailable(i: number): void { this.states[i] = POINT_UNAVAILABLE }
  /** An error keeps its code (#N/A distinct from the rest); unknown codes go to a small table. */
  error(i: number, code: string): void {
    const text = code.slice(0, CHART_ERROR_CODE_LIMIT)
    let index = KNOWN.get(text.toUpperCase())
    if (index === undefined) {
      const codes = this.codes ??= []
      let k = codes.indexOf(text)
      if (k < 0) k = codes.push(text) - 1
      index = CHART_ERROR_CODES.length + k
    }
    this.states[i] = POINT_ERROR; this.values[i] = index
  }
  done(): ChartPoints {
    return {
      length: this.length, states: this.states, values: this.values,
      ...(this.codes ? { codes: this.codes } : {}),
      ...(this.formats && this.formatIndex ? { formats: this.formats, formatIndex: this.formatIndex } : {}),
    }
  }
}

/** The plotted number of point `i`, or undefined for every other state (a drawing gap). */
export function valueAt(points: ChartPoints | undefined, i: number): number | undefined {
  return points && i < points.length && points.states[i] === POINT_VALUE ? points.values[i] : undefined
}
/** The error code of point `i` (undefined unless it is an error). */
export function errorCodeAt(points: ChartPoints, i: number): string | undefined {
  if (points.states[i] !== POINT_ERROR) return undefined
  const index = points.values[i]
  return index < CHART_ERROR_CODES.length ? CHART_ERROR_CODES[index] : points.codes?.[index - CHART_ERROR_CODES.length]
}
/** Point `i` as the `ChartPoint` union (an accessor view: one object per call, never stored). */
export function pointAt(points: ChartPoints, i: number): ChartPoint {
  switch (points.states[i]) {
    case POINT_VALUE: { const k = points.formatIndex?.[i]; return k ? { state: 'value', v: points.values[i], fmt: points.formats![k - 1] } : { state: 'value', v: points.values[i] } }
    case POINT_ERROR: return { state: 'error', code: errorCodeAt(points, i) ?? '#N/A' }
    case POINT_UNAVAILABLE: return { state: 'unavailable' }
    default: return { state: 'blank' }
  }
}
/** Every point as the union (tests, diagnostics; allocates one object per point). */
export function pointList(points: ChartPoints): ChartPoint[] {
  return Array.from({ length: points.length }, (_, i) => pointAt(points, i))
}
/** Whether point `i` sits in a hidden row or column of its own source range (sorted provenance). */
export function hiddenAt(hidden: Uint32Array | undefined, i: number): boolean {
  if (!hidden?.length) return false
  let lo = 0, hi = hidden.length
  while (lo < hi) { const mid = (lo + hi) >>> 1; if (hidden[mid] < i) lo = mid + 1; else hi = mid }
  return lo < hidden.length && hidden[lo] === i
}
/** The same storage with its own provenance (no point copied). */
export const withHidden = (points: ChartPoints, hidden: Uint32Array | undefined): ChartPoints => {
  const { hidden: _old, ...rest } = points
  void _old
  return hidden?.length ? { ...rest, hidden } : rest
}

/**
 * The input an analytic element computes from (spec 5.8, 7): every acquired
 * value of the role, blank and error points excluded. `refused` when the role
 * was not acquired in full (more than CHART_ANALYTIC_POINT_LIMIT points) or a
 * required point is unavailable: the element is then dropped, never computed
 * from a prefix or a partial set.
 */
export function analyticInput(series: Pick<ChartSeries, 'analytic' | 'values' | 'xValues'>, role: 'values' | 'xValues' = 'values'): { kind: 'input'; values: number[] } | { kind: 'refused'; reason: string } {
  if (series.analytic !== true) return { kind: 'refused', reason: series.analytic === 'refused' ? 'analytic-point-limit' : 'not-analytic' }
  const points = series[role], values: number[] = []
  for (let i = 0; i < (points?.length ?? 0); i++) {
    const state = points!.states[i]
    if (state === POINT_UNAVAILABLE) return { kind: 'refused', reason: 'chart-point-unavailable' }
    if (state === POINT_VALUE) values.push(points!.values[i])
  }
  return { kind: 'input', values }
}
