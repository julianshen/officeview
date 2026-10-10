/**
 * functions/datetime.ts — bounded scalar DATE/TIME helpers + shipping registry
 * wrappers (C3 integration of the accepted pure helpers; the algorithm bodies
 * are byte-preserving, only the type/error adapter paths differ).
 *
 * Every helper returns ElementValue with the existing typed error machinery
 * via the contracts adapter (no second model). Provenance of expectations is
 * per-row documented in tests/.
 *
 * SPEC r2 applied here (counterexamples from c3-pure-review-r1):
 *  F1  DATE backward-normalization borrows the PREVIOUS month's length
 *      (serial.ts; parallel Feb-1900 length included: DATE(1900,3,0)=60).
 *  F2  DATEVALUE text years keep literal identity 1900..9999 — DATE's
 *      0..1899 -> +1900 remap must NOT apply to literal text years.
 *  F3  DAYS numeric subtraction keeps raw fractions; numeric inputs are
 *      guarded (finite + serial range) -> #NUM!.
 *  F4  tagged EvaluationError inputs pass through all helpers untouched.
 *  F9  unsupported text-date locale is a distinguishable CAPABILITY gate
 *      exposed via dateTextParse (no workbook dispatch/cache wiring here).
 *
 * Locale caveat (root guidance): the native-families ledger did not capture
 * Excel's locale/date-order. Non-ISO slash-date outcomes in the ledger are
 * profile-specific observations, NOT a universal rule. The prototype:
 *  - always accepts ISO yyyy-mm-dd (unambiguous, measured),
 *  - applies the DOCUMENTED en-US m/d/yyyy order only when
 *    semantics.locale === 'en-US' (injected),
 *  - reports other locales as unsupported-locale capability gates (no
 *    invented order).
 *
 * UNVERIFIED boundaries kept explicit (claim-correction rule):
 *  - text-parsing rejection of the ISO literal "1900-02-29" is a MODEL
 *    decision, not a measured/documented rule (only the locale-unknown
 *    slash observation exists; no order-independence was proven).
 *  - real-calendar validity for typeable text dates in the implemented
 *    ISO/en-US paths is likewise a prototype model choice.
 */
import type {
  AstNode,
  DateSystem,
  ElementValue,
  EvaluationContext,
  EvaluationError,
  FormatResult,
  ResolvedSemantics,
} from '../types'
import { coerceToNumber, formulaError, isEvaluationError, isMatrixValue, markUnsupported } from '../evaluator'
import {
  MAX_SERIAL_1900,
  MAX_SERIAL_1904,
  civilToSerial,
  dateSerial,
  decodeSerial,
  edateSerial,
  eomonthSerial,
  epochToSerial,
  normalizeCivilToToken,
  realDaysInMonth,
  timeFraction,
  timeOfDayParts,
  weekdayOf,
} from '../serial'
import { formatCommon } from '../format'
import type { EvaluatorFn, FunctionHandler } from '../functions'

const NUM_ERR = formulaError('#NUM!') as EvaluationError
const VALUE_ERR = formulaError('#VALUE!') as EvaluationError

function numeric(v: ElementValue): number | EvaluationError {
  if (isEvaluationError(v)) return v // SPEC F4: tagged errors pass through untouched
  if (typeof v === 'number') return Number.isFinite(v) ? v : NUM_ERR
  if (typeof v === 'boolean') return v ? 1 : 0 // documented boolean->number
  return VALUE_ERR // text/blank coercion policy is an owner-side gate in this prototype
}

/** DATE(year, month, day) — documented year rule + measured boundary tokens. */
export function dateOf(
  year: ElementValue, month: ElementValue, day: ElementValue, system: DateSystem,
): ElementValue {
  const y = numeric(year)
  if (isEvaluationError(y)) return y
  const m = numeric(month)
  if (isEvaluationError(m)) return m
  const d = numeric(day)
  if (isEvaluationError(d)) return d
  // DATE truncates non-integer arguments to integers (documented behavior).
  return dateSerial(Math.trunc(y), Math.trunc(m), Math.trunc(d), system)
}

function civilParts(serial: ElementValue, system: DateSystem): ReturnType<typeof decodeSerial> {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  if (!Number.isInteger(s)) {
    // day extraction uses the truncated day part (time fraction ignored)
    return decodeSerial(Math.trunc(s), system)
  }
  return decodeSerial(s, system)
}

function validRange(s: number, system: DateSystem): boolean {
  const max = system === '1900' ? 2958465 : 2957003
  return s >= 0 && s <= max
}

/** YEAR/MONTH/DAY — serial day part -> civil token (year<0 -> #NUM!, measured). */
export function yearOf(serial: ElementValue, system: DateSystem): ElementValue {
  const c = civilParts(serial, system)
  if (isEvaluationError(c)) return c
  if (!validRange(c.daySerial, system)) return NUM_ERR
  return c.civil.year
}
export function monthOf(serial: ElementValue, system: DateSystem): ElementValue {
  const c = civilParts(serial, system)
  if (isEvaluationError(c)) return c
  if (!validRange(c.daySerial, system)) return NUM_ERR
  return c.civil.month
}
export function dayOf(serial: ElementValue, system: DateSystem): ElementValue {
  const c = civilParts(serial, system)
  if (isEvaluationError(c)) return c
  if (!validRange(c.daySerial, system)) return NUM_ERR
  return c.civil.day
}

/** TIME(h,m,s) — wraps mod 1, negative args #NUM!; args truncated to integers. */
export function timeOf(h: ElementValue, m: ElementValue, s: ElementValue): ElementValue {
  const hv = numeric(h)
  if (isEvaluationError(hv)) return hv
  const mv = numeric(m)
  if (isEvaluationError(mv)) return mv
  const sv = numeric(s)
  if (isEvaluationError(sv)) return sv
  return timeFraction(Math.trunc(hv), Math.trunc(mv), sv)
}

function fracOfDay(value: number): number {
  // NOTE: negative-fraction normalization here is a MODEL decision; only the
  // measured SECOND(0.999999)=0 whole-second rule and its timeOfDayParts
  // implementation are verified. HOUR/MINUTE at sub-second boundaries remain
  // unmeasured pending gates (never claimed as accepted Excel rules).
  return value - Math.floor(value)
}

export function hourOf(serial: ElementValue): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  return timeOfDayParts(fracOfDay(s)).hour
}
export function minuteOf(serial: ElementValue): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  return timeOfDayParts(fracOfDay(s)).minute
}
export function secondOf(serial: ElementValue): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  return timeOfDayParts(fracOfDay(s)).second
}

export function edateOf(serial: ElementValue, months: ElementValue, system: DateSystem): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  const m = numeric(months)
  if (isEvaluationError(m)) return m
  return edateSerial(Math.trunc(s), m, system)
}
export function eomonthOf(serial: ElementValue, months: ElementValue, system: DateSystem): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  const m = numeric(months)
  if (isEvaluationError(m)) return m
  return eomonthSerial(Math.trunc(s), m, system)
}
export function weekdayFn(
  serial: ElementValue, type: ElementValue | undefined, system: DateSystem,
): ElementValue {
  const s = numeric(serial)
  if (isEvaluationError(s)) return s
  let t: number | EvaluationError = undefined as unknown as number
  if (type !== undefined) {
    const tv = numeric(type)
    if (isEvaluationError(tv)) return tv
    t = Math.trunc(tv)
  }
  return weekdayOf(Math.trunc(s), t, system)
}

/**
 * Real-calendar validity for typeable text dates. NOTE (unverified boundary):
 * treating the fictitious 1900-02-29 token as untypeable in text is a MODEL
 * decision of this prototype — the accepted evidence covers only real dates;
 * ISO "1900-02-29" and slash "2/29/1900" rejections stay pending gates.
 */
function isRealCivil(y: number, m: number, d: number): boolean {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return false
  if (y < 0 || y > 9999) return false
  if (m < 1 || m > 12) return false
  return d >= 1 && d <= realDaysInMonth(y, m)
}

function serialForLiteralCivil(
  y: number, m: number, d: number, system: DateSystem,
): number | EvaluationError {
  // SPEC F2: literal text years keep their identity — DATE's 0..1899 -> +1900
  // remap must NOT apply. Primary text-year range: 1900..9999.
  if (y < 1900 || y > 9999 || !isRealCivil(y, m, d)) return VALUE_ERR
  const tok = normalizeCivilToToken(y, m, d, system)
  if (isEvaluationError(tok)) return tok
  const ser = civilToSerial(tok, system)
  if (isEvaluationError(ser)) return ser
  const max = system === '1900' ? MAX_SERIAL_1900 : MAX_SERIAL_1904
  if (ser < 0 || ser > max) return VALUE_ERR // text value outside the serial range
  return ser
}

/**
 * Text-date parse capability preflight (SPEC F9): discriminated result so the
 * owner adapter can tell a GENUINE parse error from an UNSUPPORTED-LOCALE
 * capability limit without relying on a counterfeit plain #VALUE! object.
 * Prototype-side only; no workbook dispatcher/cache integration is wired here.
 */
export type DateTextParseResult =
  | { status: 'parsed'; serial: number }
  | { status: 'invalid-date'; error: EvaluationError }
  | { status: 'unsupported-locale'; locale: string; error: EvaluationError }

export function dateTextParse(
  text: string, system: DateSystem, locale: string,
): DateTextParseResult {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text.trim())
  if (iso) {
    const s = serialForLiteralCivil(Number(iso[1]), Number(iso[2]), Number(iso[3]), system)
    if (isEvaluationError(s)) return { status: 'invalid-date', error: s }
    return { status: 'parsed', serial: s }
  }
  if (locale === 'en-US') {
    const us = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{1,4})\s*$/.exec(text)
    if (us) {
      const s = serialForLiteralCivil(Number(us[3]), Number(us[1]), Number(us[2]), system)
      if (isEvaluationError(s)) return { status: 'invalid-date', error: s }
      return { status: 'parsed', serial: s }
    }
  }
  // capability gate: this locale's text-date grammar is not implemented here
  return { status: 'unsupported-locale', locale, error: VALUE_ERR }
}

/** DATEVALUE: ISO yyyy-mm-dd always; documented en-US m/d/yyyy when locale en-US. */
export function dateValueOf(
  text: ElementValue, system: DateSystem, locale = 'en-US',
): ElementValue {
  if (isEvaluationError(text)) return text // SPEC F4
  if (typeof text !== 'string') return VALUE_ERR
  const r = dateTextParse(text, system, locale)
  if (r.status === 'parsed') return r.serial
  return r.error
}

/**
 * DAYS(end, start): numeric dates subtract RAW values (fractions preserved);
 * numeric inputs are guarded (finite + serial range 0..max) -> #NUM! outside;
 * text dates parse to INTEGER date serials via dateTextParse rules (ISO
 * always, en-US slash documented order); unsupported locales surface the
 * capability error payload (distinguishable upstream via dateTextParse).
 */
export function daysOf(
  end: ElementValue, start: ElementValue, system: DateSystem, locale = 'en-US',
): ElementValue {
  if (isEvaluationError(end)) return end // SPEC F4
  if (isEvaluationError(start)) return start
  const max = system === '1900' ? MAX_SERIAL_1900 : MAX_SERIAL_1904
  const partOf = (v: ElementValue): number | EvaluationError => {
    if (typeof v === 'string') {
      const r = dateTextParse(v, system, locale)
      return r.status === 'parsed' ? Math.trunc(r.serial) : r.error
    }
    const n = numeric(v)
    if (isEvaluationError(n)) return n
    if (n < 0 || n > max) return NUM_ERR // primary DAYS numeric range guard
    return n // raw (fractions preserved)
  }
  const e = partOf(end)
  if (isEvaluationError(e)) return e
  const st = partOf(start)
  if (isEvaluationError(st)) return st
  const diff = e - st
  return Number.isFinite(diff) ? diff : NUM_ERR
}

/** TIMEVALUE: hh:mm(:ss[.f]) with optional AM/PM (en-US measured rows). */
export function timeValueOf(text: ElementValue): ElementValue {
  if (isEvaluationError(text)) return text // SPEC F4
  if (typeof text !== 'string') return VALUE_ERR
  const m = /^\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2})(\.\d+)?)?\s*([AaPp][Mm])?\s*$/.exec(text)
  if (!m) return VALUE_ERR
  let h = Number(m[1])
  const mi = Number(m[2])
  const s = m[3] !== undefined ? Number(m[3]) + (m[4] ? Number(m[4]) : 0) : 0
  const ampm = m[5]?.toUpperCase()
  if (ampm) {
    if (h < 1 || h > 12) return VALUE_ERR
    if (ampm === 'PM' && h !== 12) h += 12
    if (ampm === 'AM' && h === 12) h = 0
  } else if (h > 23) {
    return VALUE_ERR // 24h clock gate (unmeasured beyond measured rows)
  }
  if (mi > 59 || s >= 60) return VALUE_ERR
  const total = (h * 3600 + mi * 60 + s) / 86400
  return total
}

/** TEXT: one shared formatCommon for evaluation and rendering. */
export function textOf(
  value: ElementValue, format: ElementValue, semantics: ResolvedSemantics,
): FormatResult {
  if (isEvaluationError(format)) return format // SPEC F4: tagged format arg passes through
  if (typeof format !== 'string') return VALUE_ERR
  return formatCommon(value, format, semantics)
}

/** NOW/TODAY with injected epoch + explicit zone (manual mode never calls them). */
export function nowOf(semantics: ResolvedSemantics): ElementValue {
  return epochToSerial(semantics.epochNowMs, semantics.timeZone, semantics.dateSystem)
}
export function todayOf(semantics: ResolvedSemantics): ElementValue {
  const n = epochToSerial(semantics.epochNowMs, semantics.timeZone, semantics.dateSystem)
  if (isEvaluationError(n)) return n
  return Math.trunc(n)
}

export { isEvaluationError, formulaError }

// ---------------------------------------------------------------------------
// Shipping registry wrappers.
//
// Every datetime handler resolves dateSystem/locale/timeZone/epoch from the
// active context semantics (explicit ctx.semantics wins; standalone callers get
// one captured zone + one sampled clock per context). The scalar wrappers apply
// the engine's function-specific coercion (blank -> 0, Excel numeric text,
// booleans, tagged-error passthrough); the pure helper bodies above keep their
// intentional scalar policy for the accepted algorithms.
// ---------------------------------------------------------------------------

const fallbackSemanticsCache = new WeakMap<object, ResolvedSemantics>()

function capturedRuntimeZone(): string {
  try {
    const zone = Intl.DateTimeFormat().resolvedOptions().timeZone
    return typeof zone === 'string' && zone !== '' ? zone : 'UTC'
  } catch {
    return 'UTC'
  }
}

/** Resolved semantics for the active evaluation context (never ambient). */
export function resolveSemantics(ctx?: EvaluationContext): ResolvedSemantics {
  if (ctx?.semantics) return ctx.semantics
  if (ctx) {
    let cached = fallbackSemanticsCache.get(ctx)
    if (!cached) {
      cached = {
        dateSystem: '1900',
        unicode: { version: 2, source: 'standalone-default' },
        locale: 'en-US',
        timeZone: capturedRuntimeZone(),
        epochNowMs: Date.now(),
      }
      fallbackSemanticsCache.set(ctx, cached)
    }
    return cached
  }
  return {
    dateSystem: '1900',
    unicode: { version: 2, source: 'standalone-default' },
    locale: 'en-US',
    timeZone: capturedRuntimeZone(),
    epochNowMs: Date.now(),
  }
}

/**
 * Honest gate for an UNVERIFIED matrix origin on a scalar date/time argument
 * slot: frame mark + attributable diagnostic + internal #NAME? gate shape, so
 * a valid workbook cache is retained instead of being overwritten by a
 * counterfeit #VALUE!. Shared by every scalar slot (numeric, string-date and
 * both DAYS positions).
 */
function dateMatrixGate(ctx: EvaluationContext | undefined): EvaluationError {
  markUnsupported(ctx, 'date-array-input-unverified')
  ctx?.reportFormulaIssue?.({
    kind: 'native-gate',
    gate: 'date-array-input-unverified',
    message: 'Array input to a scalar date/time function is not verified',
  })
  return formulaError('#NAME?')
}

/** Coerce one scalar argument through the engine's Excel numeric grammar. */
function numericArg(arg: AstNode, ctx: EvaluationContext | undefined, evalNode: EvaluatorFn): number | EvaluationError {
  const value = evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  // Scalar date/time array-input profiles are not verified: honest gate.
  if (isMatrixValue(value)) return dateMatrixGate(ctx)
  return coerceToNumber(value)
}

/** Mark an unsupported text-date locale as a distinguishable capability gate. */
function markDateLocaleGate(ctx: EvaluationContext | undefined, locale: string): void {
  markUnsupported(ctx, 'date-text-locale-unsupported')
  ctx?.reportFormulaIssue?.({
    kind: 'native-gate',
    gate: 'date-text-locale-unsupported',
    message: `Text-date parsing for locale '${locale}' is not implemented; ISO yyyy-mm-dd and the documented en-US order are the supported profiles`,
  })
}

function scalarStringArg(
  arg: AstNode,
  ctx: EvaluationContext | undefined,
  evalNode: EvaluatorFn,
): string | EvaluationError {
  const value = evalNode(arg, ctx)
  if (isEvaluationError(value)) return value
  // Matrix origin gates BEFORE the scalar string type check, so a matrix never
  // falls through to a counterfeit #VALUE! that would drop a valid cache.
  if (isMatrixValue(value)) return dateMatrixGate(ctx)
  if (typeof value === 'string') return value
  return formulaError('#VALUE!')
}

function oneArg(args: AstNode[]): EvaluationError | undefined {
  if (args.length !== 1) return formulaError('#VALUE!')
  return undefined
}

/** DATE(year, month, day) */
export const dateHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 3) return formulaError('#VALUE!')
  const y = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(y)) return y
  const m = numericArg(args[1], ctx, evalNode)
  if (isEvaluationError(m)) return m
  const d = numericArg(args[2], ctx, evalNode)
  if (isEvaluationError(d)) return d
  return dateOf(y, m, d, resolveSemantics(ctx).dateSystem)
}

export const yearHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return yearOf(s, resolveSemantics(ctx).dateSystem)
}

export const monthHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return monthOf(s, resolveSemantics(ctx).dateSystem)
}

export const dayHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return dayOf(s, resolveSemantics(ctx).dateSystem)
}

export const timeHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 3) return formulaError('#VALUE!')
  const h = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(h)) return h
  const m = numericArg(args[1], ctx, evalNode)
  if (isEvaluationError(m)) return m
  const s = numericArg(args[2], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return timeOf(h, m, s)
}

export const hourHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return hourOf(s)
}

export const minuteHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return minuteOf(s)
}

export const secondHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  return secondOf(s)
}

export const edateHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 2) return formulaError('#VALUE!')
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  const m = numericArg(args[1], ctx, evalNode)
  if (isEvaluationError(m)) return m
  return edateOf(s, m, resolveSemantics(ctx).dateSystem)
}

export const eomonthHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 2) return formulaError('#VALUE!')
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  const m = numericArg(args[1], ctx, evalNode)
  if (isEvaluationError(m)) return m
  return eomonthOf(s, m, resolveSemantics(ctx).dateSystem)
}

export const weekdayHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length < 1 || args.length > 2) return formulaError('#VALUE!')
  const s = numericArg(args[0], ctx, evalNode)
  if (isEvaluationError(s)) return s
  let type: number | undefined
  if (args.length === 2) {
    const t = numericArg(args[1], ctx, evalNode)
    if (isEvaluationError(t)) return t
    type = Math.trunc(t)
  }
  return weekdayFn(s, type, resolveSemantics(ctx).dateSystem)
}

export const datevalueHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const text = scalarStringArg(args[0], ctx, evalNode)
  if (isEvaluationError(text)) return text
  const semantics = resolveSemantics(ctx)
  const result = dateTextParse(text, semantics.dateSystem, semantics.locale)
  if (result.status === 'parsed') return result.serial
  if (result.status === 'unsupported-locale') {
    markDateLocaleGate(ctx, semantics.locale)
    return formulaError('#NAME?')
  }
  return result.error
}

export const timevalueHandler: FunctionHandler = (args, ctx, evalNode) => {
  const arity = oneArg(args)
  if (arity) return arity
  const text = scalarStringArg(args[0], ctx, evalNode)
  if (isEvaluationError(text)) return text
  return timeValueOf(text)
}

export const daysHandler: FunctionHandler = (args, ctx, evalNode) => {
  if (args.length !== 2) return formulaError('#VALUE!')
  const semantics = resolveSemantics(ctx)
  // One-child order: evaluate each argument exactly once, in order, and stop
  // on the first typed error or locale capability gate BEFORE touching a later
  // argument (no eager reads of a reference the first error already settles).
  const coerced: ElementValue[] = []
  for (const arg of args) {
    const value = evalNode(arg, ctx)
    if (isEvaluationError(value)) return value
    // A matrix origin on EITHER DAYS slot gates BEFORE type/coercion, and
    // BEFORE a later argument is read (one-child order preserved).
    if (isMatrixValue(value)) return dateMatrixGate(ctx)
    if (typeof value === 'string') {
      const pre = dateTextParse(value, semantics.dateSystem, semantics.locale)
      if (pre.status === 'unsupported-locale') {
        // Capability gate (not a genuine Excel error): mark the frame and
        // surface the internal/public #NAME? gate; no invented locale order.
        markDateLocaleGate(ctx, semantics.locale)
        return formulaError('#NAME?')
      }
      coerced.push(value)
      continue
    }
    const numeric = coerceToNumber(value)
    if (isEvaluationError(numeric)) return numeric
    coerced.push(numeric)
  }
  return daysOf(coerced[0], coerced[1], semantics.dateSystem, semantics.locale)
}

export const nowHandler: FunctionHandler = (args, ctx) => {
  if (args.length !== 0) return formulaError('#VALUE!')
  return nowOf(resolveSemantics(ctx))
}

export const todayHandler: FunctionHandler = (args, ctx) => {
  if (args.length !== 0) return formulaError('#VALUE!')
  return todayOf(resolveSemantics(ctx))
}

export const DATETIME_FUNCTIONS: Record<string, FunctionHandler> = {
  DATE: dateHandler,
  DATEVALUE: datevalueHandler,
  YEAR: yearHandler,
  MONTH: monthHandler,
  DAY: dayHandler,
  TIME: timeHandler,
  TIMEVALUE: timevalueHandler,
  HOUR: hourHandler,
  MINUTE: minuteHandler,
  SECOND: secondHandler,
  DAYS: daysHandler,
  EDATE: edateHandler,
  EOMONTH: eomonthHandler,
  WEEKDAY: weekdayHandler,
  TODAY: todayHandler,
  NOW: nowHandler,
}
