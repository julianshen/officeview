/**
 * Chart dates (spec 5.1; lazy chunk): the chart epoch, epoch normalization of
 * filled date roles, date-format detection and calendar stepping for date axes.
 *
 * - The chart epoch is the first that applies: `c:chartSpace/c:date1904`
 *   (present with `val` omitted means true), the workbook's
 *   `workbookPr@date1904`, the 1900 system.
 * - Conversion between epochs goes through the repository's civil-date model
 *   (`xlsx/formula/serial.ts`): source serial -> civil date and time fraction
 *   -> destination serial. It is not a blanket +-1,462 offset. A value with no
 *   counterpart (1900 serial 0, the fictitious 1900-02-29, any 1900 date before
 *   1904-01-01 going into 1904) is unrepresentable: `undefined`, which the fill
 *   records as an unavailable point with `chart-date-unrepresentable`.
 * - Month and year steps follow the calendar from the start date (EDATE rules:
 *   the day is clamped to the month's real length), never a constant day count.
 */
import { decodeSerial, edateSerial, encodeCivil } from '../xlsx/formula/serial'
import type { ChartTimeUnit } from './model'

export type ChartEpoch = '1900' | '1904'
/** Ticks per axis (spec 6.5 work bound). */
export const CHART_TICK_LIMIT = 1000

/** The chart epoch (spec 5.1 precedence): the chart's own `c:date1904`, else the workbook's, else 1900. */
export const chartEpoch = (authored?: ChartEpoch, workbook?: ChartEpoch): ChartEpoch => authored ?? workbook ?? '1900'

/** A serial of epoch `from` in epoch `to`, through the civil model; `undefined` when it has no counterpart. */
export function convertSerial(v: number, from: ChartEpoch, to: ChartEpoch): number | undefined {
  if (from === to) return v
  if (!Number.isFinite(v) || v < 0) return undefined
  const parts = decodeSerial(v, from)
  if (!('civil' in parts) || parts.civil.special) return undefined
  const day = encodeCivil({ year: parts.civil.year, month: parts.civil.month, day: parts.civil.day }, to)
  // A 1900-system date before 1904-01-01 encodes below 1904 serial 0: it has no 1904 counterpart.
  if (typeof day !== 'number' || day < 0) return undefined
  return day + parts.fraction
}

/**
 * Whether a number format code formats dates or times: a d, m, y, h or s token (or e/g/b year forms) in its
 * first section outside quotes, escapes and brackets (an elapsed `[h]` counts). `General` and plain numbers do not.
 */
export function isDateFormat(code: string | undefined): boolean {
  if (!code) return false
  let quoted = false, escaped = false, bracket: string | undefined
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (bracket !== undefined) {
      if (ch === ']') { if (/^(h+|m+|s+)$/i.test(bracket)) return true; bracket = undefined } else bracket += ch
      continue
    }
    if (quoted) { if (ch === '"') quoted = false; continue }
    if (escaped) { escaped = false; continue }
    if (ch === ';') break
    if (ch === '"') quoted = true
    else if (ch === '\\' || ch === '_' || ch === '*') escaped = true
    else if (ch === '[') bracket = ''
    else if (code.slice(i, i + 7).toLowerCase() === 'general') i += 6
    // `E+`/`e-` is scientific notation, not a year.
    else if ((ch === 'e' || ch === 'E') && (code[i + 1] === '+' || code[i + 1] === '-')) i++
    else if (/[dmyhseg]/i.test(ch)) return true
  }
  return false
}

const civilOf = (serial: number, epoch: ChartEpoch) => {
  const parts = decodeSerial(Math.floor(serial), epoch)
  return 'civil' in parts ? parts.civil : undefined
}
/** A serial's index in a time unit: its day, its calendar month (year * 12 + month - 1) or its year. */
export function dateUnitIndex(serial: number, unit: ChartTimeUnit, epoch: ChartEpoch): number {
  if (unit === 'days') return Math.floor(serial)
  const civil = civilOf(serial, epoch)
  if (!civil) return Math.floor(serial)
  return unit === 'years' ? civil.year : civil.year * 12 + civil.month - 1
}

/** The base time unit Excel infers from the dates: years when they share month and day, months when they share the day, else days. */
export function autoBaseUnit(serials: readonly number[], epoch: ChartEpoch): ChartTimeUnit {
  if (serials.length < 2) return 'days'
  let sameDay = true, sameMonth = true
  const first = civilOf(serials[0], epoch)
  if (!first) return 'days'
  for (const serial of serials) {
    const civil = civilOf(serial, epoch)
    if (!civil) return 'days'
    sameDay &&= civil.day === first.day
    sameMonth &&= civil.month === first.month
  }
  return sameDay ? (sameMonth ? 'years' : 'months') : 'days'
}

const unitSpan = (start: number, end: number, unit: ChartTimeUnit, epoch: ChartEpoch) => dateUnitIndex(end, unit, epoch) - dateUnitIndex(start, unit, epoch)
/**
 * Major ticks of a date axis from `start` to `end` (serials), every `step` `unit`s: month and year ticks are
 * calendar steps from the start date (tick k = EDATE(start, k * step months)), day ticks `start + k * step`.
 * More than CHART_TICK_LIMIT ticks enlarge the step (a visible rendering bound; `enlarged` lets the caller diagnose it).
 */
export function dateTicks(start: number, end: number, step: number, unit: ChartTimeUnit, epoch: ChartEpoch, limit = CHART_TICK_LIMIT): { ticks: number[]; step: number; enlarged: boolean } {
  let unitStep = Number.isFinite(step) && step > 0 ? (unit === 'days' ? step : Math.max(1, Math.round(step))) : 1
  const span = Math.max(0, unitSpan(start, end, unit, epoch))
  let enlarged = false
  if (span / unitStep + 1 > limit) { unitStep = Math.ceil(span / (limit - 1)); enlarged = true }
  const ticks: number[] = []
  if (!(end >= start)) return { ticks: [start], step: unitStep, enlarged }
  if (unit === 'days') {
    for (let k = 0; ticks.length < limit; k++) { const t = start + k * unitStep; if (t > end + 1e-9) break; ticks.push(t) }
    return { ticks, step: unitStep, enlarged }
  }
  const months = unit === 'years' ? 12 * unitStep : unitStep
  const base = Math.floor(start), fraction = start - base
  for (let k = 0; ticks.length < limit; k++) {
    const day = edateSerial(base, k * months, epoch)
    if (typeof day !== 'number') break
    const t = day + fraction
    if (t > end + 1e-9) break
    ticks.push(t)
  }
  return { ticks, step: unitStep, enlarged }
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
/** A plain civil label for a date-axis tick (m/d/yyyy, Mmm yyyy or yyyy by unit). Number-format-driven labels are A2's. */
export function dateTickLabel(serial: number, unit: ChartTimeUnit, epoch: ChartEpoch): string {
  const civil = civilOf(serial, epoch)
  if (!civil) return String(serial)
  return unit === 'years' ? String(civil.year) : unit === 'months' ? `${MONTHS[civil.month - 1]} ${civil.year}` : `${civil.month}/${civil.day}/${civil.year}`
}
