// @vitest-environment node
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { cellDisplayText, formatValue } from '../src/xlsx/render'
import { BUILTIN_DATE_BUILDER_BYTES, BUILTIN_DATE_CACHE_LOCALES, BUILTIN_DATE_CACHE_VERDICTS, BUILTIN_DATE_WORK_BYTES, builtinDateCacheStats, formatBuiltinDate, resetBuiltinDateCache } from '../src/xlsx/builtinDate'
import { formatCommon } from '../src/xlsx/formula/format'
import type { ResolvedSemantics } from '../src/xlsx/formula/types'
import type { XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const semantics = (locale: string, dateSystem: '1900' | '1904' = '1900'): ResolvedSemantics => ({ dateSystem, locale, timeZone: 'UTC', epochNowMs: 0, unicode: { version: 2, source: 'explicit' } })
const ids = [14, 15, 16, 17, 18, 19, 20, 21, 22]
// Localized expectations are Intl-derived presentation controls, not desktop
// Excel measurements. en-US retains the accepted Excel built-in forms.
const cases = [
  ['en-US', ['6/2/2023', '2-Jun-23', '2-Jun', 'Jun-23', '1:30 PM', '1:30:00 PM', '13:30', '13:30:00', '6/2/2023 13:30']],
  ['fr-FR', ['02/06/2023', '2 juin 23', '2 juin', 'juin 23', '1:30 PM', '1:30:00 PM', '13:30', '13:30:00', '02/06/2023 13:30']],
  ['de-DE', ['2.6.2023', '2. Juni 23', '2. Juni', 'Juni 23', '1:30 PM', '1:30:00 PM', '13:30', '13:30:00', '2.6.2023 13:30']],
  ['zh-TW', ['2023/6/2', '23年6月2日', '6月2日', '23年6月', '下午1:30', '下午1:30:00', '13:30', '13:30:00', '2023/6/2 13:30']],
] as const
// Full localized strings mirror the runtime's CLDR data: skip them where the
// runtime lacks the locale (small-ICU builds fall back to en-US by design).
const supported = (...locales: string[]) => Intl.DateTimeFormat.supportedLocalesOf(locales).length === locales.length

describe('localized built-in date/time presentation', () => {
  for (const [locale, expected] of cases) it.skipIf(locale !== 'en-US' && !supported(locale))(`${locale} localizes all built-ins 14-22 on public and sheet paths`, () => {
    const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semantics(locale) }
    expect(ids.map(id => formatValue(45079.5625, id, semantics(locale)))).toEqual(expected)
    expect(ids.map(numFmtId => cellDisplayText(sheet, { value: 45079.5625, style: { numFmtId } as XlsxCellStyle }))).toEqual(expected)
    expect((sheet as unknown as { diagnostics?: unknown[] }).diagnostics ?? []).toEqual([])
  })
  for (const [locale, zero, sixty, first1904, negative1904] of [
    ['fr-FR', '00/01/1900', '29/02/1900', '01/01/1904', '-02/01/1904'],
    ['de-DE', '0.1.1900', '29.2.1900', '1.1.1904', '-2.1.1904'],
    ['zh-TW', '1900/1/0', '1900/2/29', '1904/1/1', '-1904/1/2'],
  ]) it.skipIf(!supported(locale))(`${locale} preserves civil day 0, fictitious day 60, and 1904 instead of normalizing a JS Date`, () => {
    expect(formatValue(0, 14, semantics(locale))).toBe(zero)
    expect(formatValue(60, 14, semantics(locale))).toBe(sixty)
    expect(formatValue(0, 14, semantics(locale, '1904'))).toBe(first1904)
    expect(formatValue(-1, 14, semantics(locale, '1904'))).toBe(negative1904)
    expect(formatValue(-1, 14, semantics(locale))).toBe('#')
  })
  // ICU-independent invariants: they hold for the localized forms and for the
  // en-US fallback a runtime without the locale uses.
  it.each(['fr-FR', 'de-DE', 'zh-TW', 'ja-JP'])('%s keeps Latin digits, civil days and stored serials whatever the ICU data', locale => {
    const shifted = { ...semantics(locale), timeZone: 'Pacific/Honolulu' }
    for (const id of ids) {
      const text = formatValue(45079.5625, id, semantics(locale))
      expect(text).not.toBe('45079.5625')
      expect(text.replace(/[0-9]/g, '')).not.toMatch(/\p{Nd}/u)
      expect(formatValue(45079.5625, id, shifted)).toBe(text)
    }
    expect(formatValue(45079, 14, semantics(locale))).toContain('2023')
    for (const id of [15, 17]) expect(formatValue(45079, id, semantics(locale))).toMatch(/(^|\D)23(\D|$)/)
    expect(formatValue(0, 14, semantics(locale))).toMatch(/^(?=.*1900)(?=.*(^|\D)0{1,2}(\D|$))/)
    expect(formatValue(60, 14, semantics(locale))).toMatch(/^(?=.*1900)(?=.*(^|\D)29(\D|$))/)
    expect(formatValue(-1, 14, semantics(locale))).toBe('#')
  })
  it.skipIf(!supported('zh-TW'))('preserves localized 12 AM/PM and 24-hour midnight despite time-zone overrides', () => {
    const sem = { ...semantics('zh-TW'), timeZone: 'Pacific/Honolulu' }
    expect(formatValue(0, 18, sem)).toBe('上午12:00')
    expect(formatValue(0.5, 18, sem)).toBe('下午12:00')
    expect(formatValue(0, 20, sem)).toBe('0:00')
    expect(formatValue(0.524259259259259, 21, sem)).toBe('12:34:56')
  })
  it('keeps authored and formula finite locale capabilities unchanged', () => {
    expect(formatCommon(45079, 'yyyy-mm-dd', semantics('fr-FR')).kind).toBe('unsupported')
    const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semantics('en-US') }
    expect(cellDisplayText(sheet, { value: 45079, style: { numFmtId: 14, formatCode: 'yyyy-mm-dd' } as XlsxCellStyle })).toBe('2023-06-02')
  })
  it.each(['not_a_locale', 'zz-ZZ', 'x'.repeat(1000)])('invalid or unsupported locale %s falls back to formatted en-US and diagnoses once', locale => {
    const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semantics(locale) }
    const cell = { value: 45079, style: { numFmtId: 14 } as XlsxCellStyle }
    expect(cellDisplayText(sheet, cell)).toBe('6/2/2023')
    expect(cellDisplayText(sheet, cell)).toBe('6/2/2023')
    const diagnostics = (sheet as unknown as { diagnostics?: Array<{ feature: string; message: string }> }).diagnostics ?? []
    expect(diagnostics).toHaveLength(1)
    expect(diagnostics[0]).toMatchObject({ feature: 'format-unsupported' })
    expect(diagnostics[0].message).toMatch(/^Built-in date\/time format 'm\/d\/yyyy' is shown in en-US: .*locale/)
    expect(diagnostics[0].message).not.toMatch(/outside the finite/)
  })
  it.each(['en-us', 'en-US-u-nu-latn', 'en-US-u-ca-japanese'])('canonical US locale %s keeps the accepted fixed forms and Gregorian/Latin convention', locale => {
    expect(ids.map(id => formatValue(45079.5625, id, semantics(locale)))).toEqual(cases[0][1])
  })
})

describe('localized formatter application admission', () => {
  beforeEach(() => resetBuiltinDateCache())
  it('admits workspace before entering Intl and propagates an admission refusal', () => {
    const canonical = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    const refusal = new RangeError('test admission refused')
    expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), { allocate() { throw refusal }, release() {} })).toThrow(refusal)
    // A cache miss admits its native step separately, before entering Intl.
    const refuseMiss = { allocate(bytes: number) { if (bytes === BUILTIN_DATE_WORK_BYTES) throw refusal }, release() {} }
    expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), refuseMiss)).toThrow(refusal)
    expect(canonical).not.toHaveBeenCalled()
    canonical.mockRestore()
  })
  it.skipIf(!supported('fr-FR'))('admits a formatter construction before entering Intl once the locale verdict is cached', () => {
    const refusal = new RangeError('test admission refused')
    const refuseMiss = { allocate(bytes: number) { if (bytes === BUILTIN_DATE_WORK_BYTES) throw refusal }, release() {} }
    formatBuiltinDate(45079, 15, semantics('fr-FR'))
    const Native = Intl.DateTimeFormat, construct = vi.fn()
    Intl.DateTimeFormat = new Proxy(Native, { construct(target, args) { construct(); return Reflect.construct(target, args) } })
    try {
      expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), refuseMiss)).toThrow(refusal)
      expect(construct).not.toHaveBeenCalled()
    } finally { Intl.DateTimeFormat = Native }
  })
  it.skipIf(!supported('fr-FR'))('releases temporary estimates, leaves output charged, and never swallows output admission failure', () => {
    let live = 0, peak = 0
    const allocation = { allocate(bytes: number) { live += bytes; peak = Math.max(live, peak) }, release(bytes: number) { live -= bytes } }
    const out = formatBuiltinDate(45079.5625, 22, semantics('fr-FR'), allocation)
    expect(out).toEqual({ kind: 'text', text: '02/06/2023 13:30' })
    expect(live).toBe('02/06/2023 13:30'.length * 2)
    // Cold: canonicalization and each of the two constructions are admitted
    // and released one at a time inside the builder allowance.
    expect(peak).toBe(BUILTIN_DATE_BUILDER_BYTES + BUILTIN_DATE_WORK_BYTES)
    const refusal = new RangeError('output refused')
    const release = vi.fn()
    expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), { allocate(bytes) { if (bytes !== BUILTIN_DATE_BUILDER_BYTES) throw refusal }, release })).toThrow(refusal)
    expect(release).toHaveBeenCalledWith(BUILTIN_DATE_BUILDER_BYTES)
  })
  it('returns a typed locale capability result for invalid or unavailable input', () => {
    expect(formatBuiltinDate(1, 14, semantics('not_a_locale'))).toMatchObject({ kind: 'unsupported', reason: expect.stringContaining('locale') })
    expect(formatBuiltinDate(1, 14, semantics('zz-ZZ'))).toMatchObject({ kind: 'unsupported' })
    const canonical = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    expect(formatBuiltinDate(1, 14, semantics('x'.repeat(1000)))).toMatchObject({ kind: 'unsupported' })
    expect(canonical).not.toHaveBeenCalled()
    canonical.mockRestore()
  })
  it.skipIf(!supported('fr-FR'))('rejects oversized Intl parts before building an output outside the admitted workspace', () => {
    const parts = vi.spyOn(Intl.DateTimeFormat.prototype, 'formatToParts').mockReturnValue([{ type: 'literal', value: 'x'.repeat(1000) }])
    expect(formatBuiltinDate(1, 14, semantics('fr-FR'))).toMatchObject({ kind: 'unsupported', reason: expect.stringContaining('workspace') })
    parts.mockRestore()
  })
})

describe('application-level localized formatter cache', () => {
  beforeEach(() => resetBuiltinDateCache())
  // Counts native constructions while keeping the statics (supportedLocalesOf).
  const countConstructions = () => {
    const Native = Intl.DateTimeFormat, counter = { count: 0, restore: () => { Intl.DateTimeFormat = Native } }
    Intl.DateTimeFormat = new Proxy(Native, { construct(target, args) { counter.count++; return Reflect.construct(target, args) } })
    expect(typeof Intl.DateTimeFormat.supportedLocalesOf).toBe('function')
    return counter
  }
  it.skipIf(!supported('fr-FR'))('constructs one formatter and canonicalizes once for many cells of one locale and id', () => {
    const constructions = countConstructions(), canonical = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    try {
      const texts = Array.from({ length: 100 }, (_, i) => formatValue(45079 + i, 14, semantics('fr-FR')))
      expect(texts[0]).toBe('02/06/2023')
      expect(constructions.count).toBe(1)
      expect(canonical).toHaveBeenCalledTimes(1)
      // 22 reuses the cached 14 formatter and adds only 20.
      expect(formatValue(45079.5625, 22, semantics('fr-FR'))).toBe('02/06/2023 13:30')
      expect(constructions.count).toBe(2)
      expect(builtinDateCacheStats()).toMatchObject({ locales: ['fr-FR'], formatters: 2 })
    } finally { constructions.restore(); canonical.mockRestore() }
  })
  it.skipIf(!supported('fr-FR'))('charges a cache hit only the builder allowance and its output', () => {
    formatBuiltinDate(45079, 14, semantics('fr-FR'))
    let live = 0, peak = 0
    const allocation = { allocate(bytes: number) { live += bytes; peak = Math.max(live, peak) }, release(bytes: number) { live -= bytes } }
    expect(formatBuiltinDate(45080, 14, semantics('fr-FR'), allocation)).toEqual({ kind: 'text', text: '03/06/2023' })
    expect(live).toBe('03/06/2023'.length * 2)
    expect(peak).toBe(BUILTIN_DATE_BUILDER_BYTES + live)
  })
  it.skipIf(!supported('fr-FR', 'de-DE', 'zh-TW', 'ja-JP', 'es-ES'))('evicts the least recently used locale beyond the locale bound', () => {
    const locales = ['fr-FR', 'de-DE', 'zh-TW', 'ja-JP', 'es-ES']
    for (const locale of locales) formatValue(45079, 14, semantics(locale))
    expect(builtinDateCacheStats().locales).toEqual(locales.slice(-BUILTIN_DATE_CACHE_LOCALES))
    expect(builtinDateCacheStats().formatters).toBe(BUILTIN_DATE_CACHE_LOCALES)
    const constructions = countConstructions()
    try {
      formatValue(45079, 14, semantics('es-ES'))
      expect(constructions.count).toBe(0)
      formatValue(45079, 14, semantics('fr-FR'))
      expect(constructions.count).toBe(1)
      expect(builtinDateCacheStats().locales).toEqual(['zh-TW', 'ja-JP', 'es-ES', 'fr-FR'])
    } finally { constructions.restore() }
  })
  it('memoizes unsupported and invalid verdicts, within a bound on distinct tags', () => {
    const canonical = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    try {
      for (let i = 0; i < 3; i++) {
        expect(formatBuiltinDate(1, 14, semantics('zz-ZZ'))).toMatchObject({ kind: 'unsupported' })
        expect(formatBuiltinDate(1, 14, semantics('not_a_locale'))).toMatchObject({ kind: 'unsupported' })
      }
      expect(canonical).toHaveBeenCalledTimes(2)
      for (let i = 0; i < 100; i++) formatBuiltinDate(1, 14, semantics(`zz-Z${i}`))
      expect(builtinDateCacheStats().verdicts).toBeLessThanOrEqual(BUILTIN_DATE_CACHE_VERDICTS)
      expect(builtinDateCacheStats().formatters).toBe(0)
    } finally { canonical.mockRestore() }
  })
})
