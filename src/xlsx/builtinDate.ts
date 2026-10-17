/** Renderer-only localized built-ins. Formula TEXT keeps its finite grammar.
 * Locale patterns are Intl-derived, not a desktop Excel measurement. */
import { formatCommon, type FormatAllocation } from './formula/format'
import { decodeSerial, timeOfDayParts } from './formula/serial'
import { isEvaluationError } from './formula/evaluator'
import type { FormatResult, ResolvedSemantics } from './formula/types'

export const BUILTIN_DATE_FORMATS = ['m/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm']
/** Admission estimate for one ephemeral formatter and its options/parts/output.
 * This is an application accounting allowance, not a bound on engine-native
 * Intl workspace. No formatter or locale cache survives the call. */
export const BUILTIN_DATE_WORK_BYTES = 32 * 1024
const DATE_OPTIONS: readonly Intl.DateTimeFormatOptions[] = [
  { year: 'numeric', month: 'numeric', day: 'numeric' },
  { year: '2-digit', month: 'short', day: 'numeric' },
  { month: 'short', day: 'numeric' },
  { year: '2-digit', month: 'short' },
]
const unsupported = (): FormatResult => ({ kind: 'unsupported', reason: 'built-in date/time locale is invalid, unavailable, or longer than 128 characters' })

export function formatBuiltinDate(value: number, id: number, semantics: ResolvedSemantics, allocation?: FormatAllocation): FormatResult {
  const code = BUILTIN_DATE_FORMATS[id - 14]
  if (code === undefined) return { kind: 'unsupported', reason: 'not a built-in date/time format' }
  if (value < 0 && semantics.dateSystem === '1900') return { kind: 'text', text: '#' }
  if (semantics.locale === 'en-US') return formatCommon(value, code, semantics, allocation)
  if (typeof semantics.locale !== 'string' || semantics.locale.length > 128) return unsupported()
  // Admission precedes native Intl construction and locale canonicalization.
  const workBytes = BUILTIN_DATE_WORK_BYTES * (id === 22 ? 2 : 1)
  allocation?.allocate(workBytes)
  try {
    let locale: string | undefined
    try { locale = Intl.DateTimeFormat.supportedLocalesOf([semantics.locale], { localeMatcher: 'lookup' })[0] } catch { return unsupported() }
    if (!locale) return unsupported()
    if (/^en-US(?:-|$)/.test(locale)) return formatCommon(value, code, { ...semantics, locale: 'en-US' }, allocation)
    const magnitude = Math.abs(value), decoded = decodeSerial(Math.trunc(magnitude), semantics.dateSystem)
    if (isEvaluationError(decoded)) return decoded
    const civil = decoded.civil, time = timeOfDayParts(magnitude - Math.trunc(magnitude))
    // A leap surrogate carries month/day safely, including the fictitious
    // 1900-02-29. Day zero is substituted below; never normalize Excel tokens.
    const surrogate = new Date(Date.UTC(2000, civil.month - 1, Math.max(1, civil.day), time.hour, time.minute, time.second))
    const render = (formatId: number): string | undefined => {
      const options: Intl.DateTimeFormatOptions = formatId <= 17 ? DATE_OPTIONS[formatId - 14] : {
        hour: 'numeric', minute: '2-digit', ...(formatId === 19 || formatId === 21 ? { second: '2-digit' as const } : {}), hourCycle: formatId <= 19 ? 'h12' : 'h23',
      }
      const formatter = new Intl.DateTimeFormat(locale, { ...options, calendar: 'gregory', numberingSystem: 'latn', timeZone: 'UTC' })
      const parts = formatter.formatToParts(surrogate)
      // The application-owned parts/output builder is bounded within the
      // admitted allowance; native Intl internals remain engine-dependent.
      if (parts.length > 32 || parts.some(part => part.value.length > 128) || parts.reduce((n, part) => n + part.value.length, 0) > 256) return undefined
      return parts.map(part => {
        if (part.type === 'year') return String(civil.year).padStart(4, '0').slice(options.year === '2-digit' ? -2 : 0)
        if (part.type === 'hour') return String(Number(part.value))
        if (part.type === 'day' && civil.day === 0) return '0'.padStart(part.value.length, '0')
        // Intl parts may use narrow no-break space before the day period;
        // retain locale punctuation but expose ordinary spaces to copy/find.
        return part.type === 'literal' ? part.value.replace(/[\u00a0\u202f]/g, ' ') : part.value
      }).join('')
    }
    const first = render(id === 22 ? 14 : id), second = id === 22 ? render(20) : ''
    if (first === undefined || second === undefined) return { kind: 'unsupported', reason: 'localized date/time parts exceed application workspace limits' }
    const text = (value < 0 ? '-' : '') + first + (id === 22 ? ' ' + second : '')
    allocation?.allocate(text.length * 2)
    return { kind: 'text', text }
  } finally { allocation?.release(workBytes) }
}
