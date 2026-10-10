/**
 * serial.ts — shared Excel serial/civil date model for DATE/TIME functions and
 * the render/format path (C3 integration of the accepted pure helpers).
 *
 * ABI.md: decodeSerial(value, system): SerialParts | EvaluationError
 *         encodeCivil(civil, system): number | EvaluationError
 *
 * Model — confirmed row-by-row with executable date arithmetic BEFORE writing
 * (native targets: DATE 0/1/59/60/61/61, 366, 36526, 45657, 693962, 1904 col):
 *   1900 system: parallel calendar where 1900-02-29 exists.
 *     serial 0            = token 1900-01-00  (real 1899-12-31 slot)
 *     serial 1..59        = real 1900-01-01 .. 1900-02-28
 *     serial 60           = fictitious 1900-02-29
 *     serial 61..         = real >= 1900-03-01  (z from 1899-12-31 + 1, z=60 is Mar 1)
 *     encodeCivil(c)      = z - BASE_1900 + (real(c) >= 1900-03-01 ? 1 : 0),
 *                           with (1900,2,29) -> 60 special; day/month normalization
 *                           runs in the PARALLEL calendar (Feb-1900 length 29).
 *   1904 system: serial = real ordinal - ordinal(1904-01-01); no fictitious day.
 *   No JS Date anywhere for the 1900 tokens; pure day arithmetic only.
 */
import type { DateSystem, EvaluationError, SerialCivil, SerialParts } from './types'
import { formulaError, isEvaluationError } from './evaluator'

type SerialResult = SerialParts | EvaluationError
type Civil = { year: number; month: number; day: number }

// ---------------------------------------------------------------------------
// Proleptic-Gregorian day arithmetic (Hinnant's algorithm; no JS Date).
// ---------------------------------------------------------------------------

/** Days since 1970-01-01 (mod-allowed epoch; only differences are used). */
export function daysFromCivil(y0: number, m0: number, d0: number): number {
  let y = y0
  y -= m0 <= 2 ? 1 : 0
  const era = Math.floor((y >= 0 ? y : y - 399) / 400)
  const yoe = y - era * 400
  const mp = (m0 + (m0 > 2 ? -3 : 9)) % 12
  const doy = Math.floor((153 * mp + 2) / 5) + (d0 <= 0 ? d0 - 1 : d0 - 1)
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy
  return era * 146097 + doe - 719468 - (d0 <= 0 ? 0 : 0)
}

export function civilFromDays(z: number): Civil {
  z += 719468
  const era = Math.floor(z / 146097)
  const doe = z - era * 146097
  const yoe = Math.floor(
    (doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365,
  )
  let y = yoe + era * 400
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100))
  const mp = Math.floor((5 * doy + 2) / 153)
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1
  const m = mp + (mp < 10 ? 3 : -9)
  y += m <= 2 ? 1 : 0
  return { year: y, month: m, day: d }
}

const BASE_1900 = daysFromCivil(1899, 12, 31) // --25568; z=0 slot, token 1900-01-00
const BASE_1904 = daysFromCivil(1904, 1, 1)
const MARCH_1_1900_Z = BASE_1900 + 60 // real 1900-03-01 slot (z >= this -> +1)

/** Documented max serials: 1900 -> 9999-12-31 = 2958465; 1904 -> 2957003. */
export const MAX_SERIAL_1900 = daysFromCivil(9999, 12, 31) - BASE_1900 + 1
export const MAX_SERIAL_1904 = daysFromCivil(9999, 12, 31) - BASE_1904

function isInt(v: number): boolean {
  return typeof v === 'number' && Number.isFinite(v) && Number.isInteger(v)
}

/** Excel documented DATE year rule: year 0..1899 -> 1900 + year (both systems). */
export function applyDocumentedYearRule(year: number): number {
  if (!isInt(year) || year < 0 || year > 9999) return NaN
  return year <= 1899 ? 1900 + year : year
}

const NUM = formulaError('#NUM!')
const VALUE = formulaError('#VALUE!')

/** Real (not parallel) length of a month — EOMONTH evidence keeps Feb-1900 at 28. */
export function realDaysInMonth(year: number, month: number): number {
  switch (month) {
    case 1: case 3: case 5: case 7: case 8: case 10: case 12: return 31
    case 4: case 6: case 9: case 11: return 30
    case 2: return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0 ? 29 : 28
    default: return NaN
  }
}

/** PARALLEL-calendar month length: February 1900 holds the fictitious day. */
function parallelDaysInMonth(year: number, month: number, system: DateSystem): number {
  if (system === '1900' && year === 1900 && month === 2) return 29
  return realDaysInMonth(year, month)
}

/**
 * Normalize (year, month, day) in the parallel calendar with full overflow
 * (month > 12 / < 1, day > month length incl. day 0 / negative days).
 * Returns the token civil; 1900-02-29 tokens only exist via this path.
 *
 * Backward normalization (SPEC F1 fix): day 0 / negative days of month m
 * borrow the PREVIOUS month's length (day 0 of March = last day of February;
 * in the 1900 parallel calendar that last day is the fictitious 1900-02-29,
 * so DATE(1900,3,0)=60 per the accepted DATE(1900,2,29)=60 anchor).
 */
function normalizeCivilToToken(
  y0: number, m0: number, d0: number, system: DateSystem,
): Civil | EvaluationError {
  let y = y0
  let m = m0
  let d = d0
  let guard = 0
  for (;;) {
    if (++guard > 40000) return NUM
    if (m < 1 || m > 12) {
      y += Math.floor((m - 1) / 12)
      m = (((m - 1) % 12) + 12) % 12 + 1
      continue
    }
    const len = parallelDaysInMonth(y, m, system)
    if (d > len) { d -= len; m += 1; continue }
    if (d < 1) {
      // borrow the PREVIOUS (parallel) month's length
      m -= 1
      if (m < 1) {
        y += Math.floor((m - 1) / 12)
        m = (((m - 1) % 12) + 12) % 12 + 1
      }
      d += parallelDaysInMonth(y, m, system)
      continue
    }
    return { year: y, month: m, day: d }
  }
}

/** Parallel-calendar serial of a normalized token civil. */
function civilToSerial(c: Civil, system: DateSystem): number | EvaluationError {
  const y = c.year, m = c.month, d = c.day
  if (system === '1900') {
    if (y === 1900 && m === 2 && d === 29) return 60 // fictitious slot
    if (y === 1900 && m === 1 && d === 0) return 0 // january-day-zero token
    const z = daysFromCivil(y, m, d)
    return z - BASE_1900 + (z >= MARCH_1_1900_Z ? 1 : 0)
  }
  const z = daysFromCivil(y, m, d)
  return z - BASE_1904
}

export { civilToSerial, normalizeCivilToToken }

/**
 * ABI: encodeCivil — parallel normalization incl. 1900 day-0/fictitious tokens.
 */
export function encodeCivil(civil: SerialCivil, system: DateSystem): number | EvaluationError {
  if (!civil || typeof civil !== 'object') return VALUE
  const y0 = civil.year, m0 = civil.month, d0 = civil.day
  if (!isInt(y0) || !isInt(m0) || !isInt(d0)) return NUM
  if (civil.special === '1900-january-day-zero' && system === '1900') {
    // Token round-trips exactly through its own slot (1900,1,0).
    const tok = normalizeCivilToToken(y0, m0, d0, system)
    if (isEvaluationError(tok)) return tok
  }
  const tok = normalizeCivilToToken(y0, m0, d0, system)
  if (isEvaluationError(tok)) return tok
  const s = civilToSerial(tok, system)
  if (isEvaluationError(s)) return s
  return s
}

/** ABI: decodeSerial — daySerial, raw fraction, civil tokens (day0/fictitious). */
export function decodeSerial(value: number, system: DateSystem): SerialResult {
  if (typeof value !== 'number' || !Number.isFinite(value)) return NUM
  const daySerial = Math.trunc(value)
  const fraction = value - Math.trunc(value) // raw fractional precision, unrounded

  if (system === '1900') {
    if (daySerial === 0) {
      return {
        system, daySerial, fraction,
        civil: { year: 1900, month: 1, day: 0, special: '1900-january-day-zero' },
      }
    }
    if (daySerial === 60) {
      return {
        system, daySerial, fraction,
        civil: { year: 1900, month: 2, day: 29, special: '1900-fictitious-february-29' },
      }
    }
    // 1..59: real slot z = BASE_1900 + serial (1899-12-31 + serial)
    // >=61 / <=-1: slot shifted by the fictitious day: z = BASE_1900 + serial - 1 for >=61,
    //             z = BASE_1900 + serial for <= -1 (1899-12-30 at serial -1).
    const z = daySerial >= 61 ? BASE_1900 + daySerial - 1 : BASE_1900 + daySerial
    const c = civilFromDays(z)
    return { system, daySerial, fraction, civil: c }
  }

  const z = BASE_1904 + daySerial
  const c = civilFromDays(z)
  return { system, daySerial, fraction, civil: c }
}

/** DATE core: documented year rule + parallel normalization + serial-range gates. */
export function dateSerial(
  year: number, month: number, day: number, system: DateSystem,
): number | EvaluationError {
  if (!isInt(year) || !isInt(month) || !isInt(day)) return NUM
  const y2 = applyDocumentedYearRule(year)
  if (!isInt(y2)) return NUM
  if (y2 > 9999) return NUM
  if (system === '1904') {
    if (y2 < 1904 || y2 > 9999) return NUM // measured: DATE(1900,1,1,'1904') = #NUM!
  } else if (y2 < 1900 || y2 > 9999) return NUM
  const tok = normalizeCivilToToken(y2, month, day, system)
  if (isEvaluationError(tok)) return tok
  const s = civilToSerial(tok, system)
  if (isEvaluationError(s)) return s
  const max = system === '1900' ? MAX_SERIAL_1900 : MAX_SERIAL_1904
  if (s < 0 || s > max) return NUM
  return s
}

/** WEEKDAY index 0=Sunday in the serial system's own calendar (parallel for 1900). */
export function weekdayIndex(serial: number, system: DateSystem): number | EvaluationError {
  if (!isInt(serial)) return NUM
  const off = system === '1900' ? 6 : 5 // serial+off ≡ 0 (mod 7) on Sunday
  return ((serial + off) % 7 + 7) % 7
}

/** WEEKDAY return types 1|2|3|11..17 (documented finite set). */
export function weekdayOf(
  serial: number, type: number | undefined, system: DateSystem,
): number | EvaluationError {
  const w = weekdayIndex(serial, system)
  if (isEvaluationError(w)) return w
  const t = type === undefined ? 1 : type
  if (!isInt(t) || !(t === 1 || t === 2 || t === 3 || (t >= 11 && t <= 17))) return NUM
  if (t === 1) return w + 1
  if (t === 2 || t === 11) return ((w + 6) % 7) + 1
  if (t === 3) return (w + 6) % 7
  const start = t - 10 // 11 -> Monday(=1 designator).. 17 -> Sunday(=7)
  return ((w - start) % 7 + 7) % 7 + 1
}

/** TIME: (h,m,s)/86400 wrapped mod 1; any negative argument -> #NUM! (measured).
 *
 * QUALITY Q1 fix: finite huge arguments previously overflowed the h*3600/m*60
 * multiplication to Infinity and escaped as numeric NaN. Each argument is now
 * reduced modulo a whole number of days BEFORE scaling (a whole-day shift
 * cannot change the result mod 1), so every intermediate stays far below the
 * double range and a final finite guard makes a non-finite success impossible.
 */
export function timeFraction(h: number, m: number, s: number): number | EvaluationError {
  if (!Number.isFinite(h) || !Number.isFinite(m) || !Number.isFinite(s)) return NUM
  if (h < 0 || m < 0 || s < 0) return NUM // measured TIME(-1,0,0)=#NUM!
  const bh = h % 86400 // 86400 hours = 3600 days
  const bm = m % 86400 // 86400 minutes = 60 days
  const bs = s % 86400 // whole days of seconds
  const total = (Math.floor(bh) * 3600 + Math.floor(bm) * 60 + bs) / 86400
  const result = total - Math.floor(total)
  return Number.isFinite(result) ? result : NUM
}

/** Time-of-day parts rounded to whole seconds (measured SECOND(0.999999)=0). */
export function timeOfDayParts(fraction: number): { hour: number; minute: number; second: number } {
  const f = fraction < 0 ? fraction - Math.floor(fraction) : fraction
  const secs = Math.round(f * 86400)
  return {
    hour: Math.floor(secs / 3600) % 24,
    minute: Math.floor(secs / 60) % 60,
    second: secs % 60,
  }
}

function civilOfDaySerial(serial: number, system: DateSystem): Civil | EvaluationError {
  const d = decodeSerial(serial, system)
  if (isEvaluationError(d)) return d
  return d.civil
}

/** EDATE: civil month shift (fictitious source day = 29), day clamp to REAL length. */
export function edateSerial(serial: number, months: number, system: DateSystem): number | EvaluationError {
  if (!isInt(serial) || !Number.isFinite(serial)) return NUM
  if (!Number.isFinite(months)) return NUM
  const m32 = Math.trunc(months)
  const c = civilOfDaySerial(serial, system)
  if (isEvaluationError(c)) return c
  const tm = c.month + m32
  const y2 = c.year + Math.floor((tm - 1) / 12)
  const m2 = (((tm - 1) % 12) + 12) % 12 + 1
  if (y2 < 1 || y2 > 9999) return NUM
  const len = realDaysInMonth(y2, m2)
  const d2 = Math.min(c.day, len)
  const tok = normalizeCivilToToken(y2, m2, d2, system)
  if (isEvaluationError(tok)) return tok
  const s = civilToSerial(tok, system)
  if (isEvaluationError(s)) return s
  const max = system === '1900' ? MAX_SERIAL_1900 : MAX_SERIAL_1904
  if (s < 0 || s > max) return NUM
  return s
}

/** EOMONTH: REAL month end (measured EOMONTH(60,0)=59, not fictitious 60). */
export function eomonthSerial(serial: number, months: number, system: DateSystem): number | EvaluationError {
  if (!isInt(serial)) return NUM
  if (!Number.isFinite(months)) return NUM
  const m32 = Math.trunc(months)
  const c = civilOfDaySerial(serial, system)
  if (isEvaluationError(c)) return c
  const tm = c.month + m32
  const y2 = c.year + Math.floor((tm - 1) / 12)
  const m2 = (((tm - 1) % 12) + 12) % 12 + 1
  if (y2 < 1 || y2 > 9999) return NUM
  const len = realDaysInMonth(y2, m2)
  const tok = normalizeCivilToToken(y2, m2, len, system)
  if (isEvaluationError(tok)) return tok
  const s = civilToSerial(tok, system)
  if (isEvaluationError(s)) return s
  const max = system === '1900' ? MAX_SERIAL_1900 : MAX_SERIAL_1904
  if (s < 0 || s > max) return NUM
  return s
}

/**
 * NOW/TODAY: interpret the injected epochNowMs in the explicit resolved
 * timeZone (IANA via Intl). Deterministic: no ambient Date.now()/argless
 * new Date()/ambient zone sampling anywhere.
 */
export function epochToSerial(
  epochNowMs: number, timeZone: string, system: DateSystem,
): number | EvaluationError {
  if (!Number.isFinite(epochNowMs) || typeof timeZone !== 'string' || timeZone === '') return NUM
  let parts
  try {
    parts = new Intl.DateTimeFormat('en-US', {
      timeZone,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
      hourCycle: 'h23',
    }).formatToParts(new Date(epochNowMs)) as { type: string; value: string }[]
  } catch {
    return NUM
  }
  const get = (t: string) => Number(parts.find((p) => p.type === t)?.value)
  const y = get('year'), mo = get('month'), da = get('day')
  let hh = get('hour'), mi = get('minute'), ss = get('second')
  if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(da)) return NUM
  if (hh === 24) hh = 0 // h23 quirk guard
  if (!Number.isFinite(hh) || !Number.isFinite(mi) || !Number.isFinite(ss)) return NUM
  const day = encodeCivil({ year: y, month: mo, day: da }, system)
  if (isEvaluationError(day)) return day
  const fracSecs = hh * 3600 + mi * 60 + ss
  const fraction = (fracSecs * 1000 + (((epochNowMs % 1000) + 1000) % 1000)) / 86400000
  return day + fraction
}
