/** Renderer-only localized built-ins. Formula TEXT keeps its finite grammar.
 * Locale patterns are Intl-derived, not a desktop Excel measurement. */
import { formatCommon, type FormatAllocation } from './formula/format'
import { decodeSerial, timeOfDayParts } from './formula/serial'
import { isEvaluationError } from './formula/evaluator'
import type { FormatResult, ResolvedSemantics } from './formula/types'

export const BUILTIN_DATE_FORMATS = ['m/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm']
/** Admission estimate for one native Intl step a cache cannot answer: locale
 * canonicalization or constructing one formatter. It is admitted before the
 * step and released after it; an application accounting allowance, not a
 * bound on engine-native Intl workspace. */
export const BUILTIN_DATE_WORK_BYTES = 32 * 1024
/** Per-call allowance for the options/parts/output builders. */
export const BUILTIN_DATE_BUILDER_BYTES = 4 * 1024
/** Canonical locales whose formatters stay cached. */
export const BUILTIN_DATE_CACHE_LOCALES = 4
/** Raw locale tags whose canonicalization verdicts stay cached. */
export const BUILTIN_DATE_CACHE_VERDICTS = 16
const DATE_OPTIONS: readonly Intl.DateTimeFormatOptions[] = [
  { year: 'numeric', month: 'numeric', day: 'numeric' },
  { year: '2-digit', month: 'short', day: 'numeric' },
  { month: 'short', day: 'numeric' },
  { year: '2-digit', month: 'short' },
]
const optionsFor = (formatId: number): Intl.DateTimeFormatOptions => formatId <= 17 ? DATE_OPTIONS[formatId - 14] : {
  hour: 'numeric', minute: '2-digit', ...(formatId === 19 || formatId === 21 ? { second: '2-digit' as const } : {}), hourCycle: formatId <= 19 ? 'h12' : 'h23',
}
// Application-level cache, module-scoped on purpose: semantics objects are
// plain values spread or rebuilt per call (render.ts, semanticsFor, public
// formatValue callers), so no document owner reaches this seam. Both maps are
// LRU by insertion order and fixed-bounded: at most 4 locales x 8 formatters
// (22 reuses 14 and 20) and 16 verdicts. The engine keeps its own ICU caches.
const formatters = new Map<string, Map<number, Intl.DateTimeFormat>>()
const verdicts = new Map<string, string | null>()
function touch<K, V>(map: Map<K, V>, key: K, value: V, limit: number): V {
  map.delete(key)
  map.set(key, value)
  if (map.size > limit) map.delete(map.keys().next().value as K)
  return value
}
/** Diagnostic/test seam: what the application-level cache holds. */
export function builtinDateCacheStats(): { locales: string[]; formatters: number; verdicts: number } {
  let count = 0
  for (const byId of formatters.values()) count += byId.size
  return { locales: [...formatters.keys()], formatters: count, verdicts: verdicts.size }
}
/** Test seam: drops every cached formatter and verdict. */
export function resetBuiltinDateCache(): void { formatters.clear(); verdicts.clear() }
/** Runs one admitted native step; the estimate ends with the step. */
function admitted<T>(allocation: FormatAllocation | undefined, step: () => T): T {
  allocation?.allocate(BUILTIN_DATE_WORK_BYTES)
  try { return step() } finally { allocation?.release(BUILTIN_DATE_WORK_BYTES) }
}
/** The canonical supported locale, or null when invalid or unavailable. */
function canonicalLocale(tag: string, allocation?: FormatAllocation): string | null {
  const cached = verdicts.get(tag)
  if (cached !== undefined) return touch(verdicts, tag, cached, BUILTIN_DATE_CACHE_VERDICTS)
  const locale = admitted(allocation, () => {
    try { return Intl.DateTimeFormat.supportedLocalesOf([tag], { localeMatcher: 'lookup' })[0] ?? null } catch { return null }
  })
  return touch(verdicts, tag, locale, BUILTIN_DATE_CACHE_VERDICTS)
}
function formatterFor(locale: string, formatId: number, allocation?: FormatAllocation): Intl.DateTimeFormat {
  const byId = formatters.get(locale) ?? new Map<number, Intl.DateTimeFormat>()
  const formatter = byId.get(formatId) ?? admitted(allocation, () => new Intl.DateTimeFormat(locale, { ...optionsFor(formatId), calendar: 'gregory', numberingSystem: 'latn', timeZone: 'UTC' }))
  byId.set(formatId, formatter)
  touch(formatters, locale, byId, BUILTIN_DATE_CACHE_LOCALES)
  return formatter
}
const unsupported = (): FormatResult => ({ kind: 'unsupported', reason: 'built-in date/time locale is invalid, unavailable, or longer than 128 characters' })

export function formatBuiltinDate(value: number, id: number, semantics: ResolvedSemantics, allocation?: FormatAllocation): FormatResult {
  const code = BUILTIN_DATE_FORMATS[id - 14]
  if (code === undefined) return { kind: 'unsupported', reason: 'not a built-in date/time format' }
  if (value < 0 && semantics.dateSystem === '1900') return { kind: 'text', text: '#' }
  if (semantics.locale === 'en-US') return formatCommon(value, code, semantics, allocation)
  if (typeof semantics.locale !== 'string' || semantics.locale.length > 128) return unsupported()
  // Admission precedes any native Intl step; a cache hit enters none.
  allocation?.allocate(BUILTIN_DATE_BUILDER_BYTES)
  try {
    const locale = canonicalLocale(semantics.locale, allocation)
    if (!locale) return unsupported()
    if (/^en-US(?:-|$)/.test(locale)) return formatCommon(value, code, { ...semantics, locale: 'en-US' }, allocation)
    const magnitude = Math.abs(value), decoded = decodeSerial(Math.trunc(magnitude), semantics.dateSystem)
    if (isEvaluationError(decoded)) return decoded
    const civil = decoded.civil, time = timeOfDayParts(magnitude - Math.trunc(magnitude))
    // A leap surrogate carries month/day safely, including the fictitious
    // 1900-02-29. Day zero is substituted below; never normalize Excel tokens.
    const surrogate = new Date(Date.UTC(2000, civil.month - 1, Math.max(1, civil.day), time.hour, time.minute, time.second))
    const render = (formatId: number): string | undefined => {
      const parts = formatterFor(locale, formatId, allocation).formatToParts(surrogate), twoDigitYear = formatId === 15 || formatId === 17
      // The application-owned parts/output builder is bounded within the
      // admitted allowance; native Intl internals remain engine-dependent.
      if (parts.length > 32 || parts.some(part => part.value.length > 128) || parts.reduce((n, part) => n + part.value.length, 0) > 256) return undefined
      return parts.map(part => {
        if (part.type === 'year') return String(civil.year).padStart(4, '0').slice(twoDigitYear ? -2 : 0)
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
  } finally { allocation?.release(BUILTIN_DATE_BUILDER_BYTES) }
}
