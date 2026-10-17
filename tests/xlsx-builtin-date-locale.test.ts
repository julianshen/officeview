// @vitest-environment node
import { describe, expect, it, vi } from 'vitest'
import { cellDisplayText, formatValue } from '../src/xlsx/render'
import { BUILTIN_DATE_WORK_BYTES, formatBuiltinDate } from '../src/xlsx/builtinDate'
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

describe('localized built-in date/time presentation', () => {
  it.each(cases)('%s localizes all built-ins14-22 on public and sheet paths', (locale, expected) => {
    const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], semantics: semantics(locale) }
    expect(ids.map(id => formatValue(45079.5625, id, semantics(locale)))).toEqual(expected)
    expect(ids.map(numFmtId => cellDisplayText(sheet, { value: 45079.5625, style: { numFmtId } as XlsxCellStyle }))).toEqual(expected)
    expect((sheet as unknown as { diagnostics?: unknown[] }).diagnostics ?? []).toEqual([])
  })
  it.each([
    ['fr-FR', '00/01/1900', '29/02/1900', '01/01/1904', '-02/01/1904'],
    ['de-DE', '0.1.1900', '29.2.1900', '1.1.1904', '-2.1.1904'],
    ['zh-TW', '1900/1/0', '1900/2/29', '1904/1/1', '-1904/1/2'],
  ])('%s preserves civil day0, fictitiousday60, and1904 instead of normalizing a JS Date', (locale, zero, sixty, first1904, negative1904) => {
    expect(formatValue(0, 14, semantics(locale))).toBe(zero)
    expect(formatValue(60, 14, semantics(locale))).toBe(sixty)
    expect(formatValue(0, 14, semantics(locale, '1904'))).toBe(first1904)
    expect(formatValue(-1, 14, semantics(locale, '1904'))).toBe(negative1904)
    expect(formatValue(-1, 14, semantics(locale))).toBe('#')
  })
  it('preserves localized12AM/PM and24hour midnight despite time-zone overrides', () => {
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
    expect(diagnostics[0].message).toMatch(/locale/)
  })
})


describe('localized formatter application admission', () => {
  it('admits workspace before entering Intl and propagates an admission refusal', () => {
    const supported = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    const refusal = new RangeError('test admission refused')
    expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), { allocate() { throw refusal }, release() {} })).toThrow(refusal)
    expect(supported).not.toHaveBeenCalled()
    supported.mockRestore()
  })
  it('releases temporary estimates, leaves output charged, and never swallows output admission failure', () => {
    let live = 0, peak = 0
    const allocation = { allocate(bytes: number) { live += bytes; peak = Math.max(live, peak) }, release(bytes: number) { live -= bytes } }
    const out = formatBuiltinDate(45079.5625, 22, semantics('fr-FR'), allocation)
    expect(out).toEqual({ kind: 'text', text: '02/06/2023 13:30' })
    expect(live).toBe('02/06/2023 13:30'.length * 2)
    expect(peak).toBe(BUILTIN_DATE_WORK_BYTES * 2 + live)
    const refusal = new RangeError('output refused')
    const release = vi.fn()
    expect(() => formatBuiltinDate(45079, 14, semantics('fr-FR'), { allocate(bytes) { if (bytes !== BUILTIN_DATE_WORK_BYTES) throw refusal }, release })).toThrow(refusal)
    expect(release).toHaveBeenCalledWith(BUILTIN_DATE_WORK_BYTES)
  })
  it('returns a typed locale capability result for invalid or unavailable input', () => {
    expect(formatBuiltinDate(1, 14, semantics('not_a_locale'))).toMatchObject({ kind: 'unsupported', reason: expect.stringContaining('locale') })
    expect(formatBuiltinDate(1, 14, semantics('zz-ZZ'))).toMatchObject({ kind: 'unsupported' })
    const supported = vi.spyOn(Intl.DateTimeFormat, 'supportedLocalesOf')
    expect(formatBuiltinDate(1, 14, semantics('x'.repeat(1000)))).toMatchObject({ kind: 'unsupported' })
    expect(supported).not.toHaveBeenCalled()
    supported.mockRestore()
  })
})

it.each(['en-us', 'en-US-u-nu-latn', 'en-US-u-ca-japanese'])('canonical US locale %s keeps the accepted fixed forms and Gregorian/Latin convention', locale => {
  expect(ids.map(id => formatValue(45079.5625, id, semantics(locale)))).toEqual(cases[0][1])
})

it('rejects oversized Intl parts before building an output outside the admitted workspace', () => {
  const parts = vi.spyOn(Intl.DateTimeFormat.prototype, 'formatToParts').mockReturnValue([{ type: 'literal', value: 'x'.repeat(1000) }])
  expect(formatBuiltinDate(1, 14, semantics('fr-FR'))).toMatchObject({ kind: 'unsupported', reason: expect.stringContaining('workspace') })
  parts.mockRestore()
})
