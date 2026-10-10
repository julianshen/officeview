import type { CalendarDay, TypedCondition } from './contracts'
import type { DateSystem } from '../formula/types'
import { civilFromDays, daysFromCivil, decodeSerial, encodeCivil, epochToSerial, realDaysInMonth, weekdayIndex } from '../formula/serial'
export interface DateContext { dateSystem: DateSystem; timeZone: string; nowMs: number }
/** Structured-cloneable pure value; unbounded before/after use internal infinities.
 * Serialize the original rules for JSON replay, not these lowered boundaries. */
export interface DateInterval { start: number; end: number; complement: boolean }
type DateCondition = Extract<TypedCondition, { kind: 'date' | 'datePreset' }>
function numeric(result: unknown): number {
  if (typeof result !== 'number' || !Number.isFinite(result)) throw new RangeError('Invalid date or date semantics')
  return result
}
function validateDay(date: CalendarDay, system: DateSystem): number {
  if (!date || !Number.isInteger(date.year) || date.year < 1 || date.year > 9999
    || !Number.isInteger(date.month) || date.month < 1 || date.month > 12
    || !Number.isInteger(date.day)) throw new RangeError('Invalid calendar date')
  const length = system === '1900' && date.year === 1900 && date.month === 2 ? 29 : realDaysInMonth(date.year, date.month)
  if (date.day < 1 || date.day > length) throw new RangeError('Invalid calendar date')
  return numeric(encodeCivil(date, system))
}
function nextDay(date: CalendarDay, system: DateSystem): number {
  // The serial helper owns parallel-calendar normalization, including the sentinel
  // boundary 10000-01-01 needed to include all of 9999-12-31.
  return numeric(encodeCivil({ ...date, day: date.day + 1 }, system))
}
export function lowerDateCondition(condition: DateCondition, context: DateContext): DateInterval {
  const system = context.dateSystem
  if (system !== '1900' && system !== '1904') throw new RangeError('Invalid date system')
  if (condition.kind === 'date') {
    const start = validateDay(condition.first, system), end = nextDay(condition.first, system)
    switch (condition.op) {
      case 'eq': return { start, end, complement: false }
      case 'neq': return { start, end, complement: true }
      case 'before': return { start: -Infinity, end: start, complement: false }
      case 'after': return { start: end, end: Infinity, complement: false }
      case 'between': {
        if (!condition.second) throw new RangeError('Missing second date bound')
        const second = validateDay(condition.second, system)
        if (start > second) throw new RangeError('Reversed date bounds')
        return { start, end: nextDay(condition.second, system), complement: false }
      }
      default: throw new RangeError('Invalid date operator')
    }
  }
  const today = Math.floor(numeric(epochToSerial(context.nowMs, context.timeZone, system)))
  const decoded = decodeSerial(today, system)
  if ('kind' in decoded) throw new RangeError('Invalid current date')
  const date = decoded.civil.special === '1900-january-day-zero'
    ? civilFromDays(daysFromCivil(1900, 1, 1) - 1) : decoded.civil
  validateDay(date, system)
  let start: number, end: number
  switch (condition.preset) {
    case 'today': start = today; end = nextDay(date, system); break
    case 'yesterday':
      start = numeric(encodeCivil({ ...date, day: date.day - 1 }, system)); end = today; break
    case 'tomorrow':
      start = nextDay(date, system); end = numeric(encodeCivil({ ...date, day: date.day + 2 }, system)); break
    case 'thisWeek': {
      const weekday = numeric(weekdayIndex(today, system))
      start = numeric(encodeCivil({ ...date, day: date.day - weekday }, system))
      end = numeric(encodeCivil({ ...date, day: date.day - weekday + 7 }, system))
      break
    }
    case 'thisMonth':
      start = numeric(encodeCivil({ ...date, day: 1 }, system))
      end = numeric(encodeCivil({ ...date, month: date.month + 1, day: 1 }, system)); break
    case 'thisYear':
      start = numeric(encodeCivil({ year: date.year, month: 1, day: 1 }, system))
      end = numeric(encodeCivil({ year: date.year + 1, month: 1, day: 1 }, system)); break
    default: throw new RangeError('Invalid date preset')
  }
  return { start, end, complement: false }
}
export function matchesDateInterval(serial: number, interval: DateInterval): boolean {
  const inside = serial >= interval.start && serial < interval.end
  return Number.isFinite(serial) && (interval.complement ? !inside : inside)
}
