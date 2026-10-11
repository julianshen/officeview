import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { CANONICAL_KEY_VERSION, canonicalKeyStrategyId, canonicalSearchKey, foldVariantOf, fullCaseFold, graphemeClusters, isGraphemeBoundary, isSimpleKey, isTurkicLocale, keyBegins, keyContains, keyEnds, UNICODE_VERSION, unicodeProperty } from '../src/xlsx/view/canonicalText'
import { compileKeyMatcher, DEFAULT_FILTER_WORK, TextFilterMatcher, WorkBudget } from '../src/xlsx/view/textFilter'
import { FilterEngine, type ColumnClassification, type PreparedValue } from '../src/xlsx/view/filter'
import type { TextOp } from '../src/xlsx/view/contracts'
import { oracleCanonical, oracleGraphemes, oracleMatch } from './helpers/canonical-text-oracle'

const ops: TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
const matcher = (locale = 'en-US') => new TextFilterMatcher(locale)
const match = (op: TextOp, pattern: string, text: string, locale = 'en-US', wildcards = false) => matcher(locale).match(op, pattern, wildcards, text)
const wild = (pattern: string, text: string, locale = 'en-US') => matcher(locale).match('eq', pattern, true, text)
const hex = (text: string) => Array.from(text, point => point.codePointAt(0)!.toString(16)).join(' ')

describe('pinned Unicode data', () => {
  it('pins one Unicode version independent of the host ICU', () => {
    expect(UNICODE_VERSION).toBe('16.0.0')
  })
  it('passes every GraphemeBreakTest-16.0.0 case with the production segmenter', () => {
    const fixture = JSON.parse(readFileSync(join(__dirname, 'fixtures/unicode/grapheme-break-test-16.0.0.json'), 'utf8')) as { cases: string[] }
    let passed = 0
    for (const entry of fixture.cases) {
      const clusters = entry.split('|').map(cluster => String.fromCodePoint(...cluster.split('.').map(value => parseInt(value, 16))))
      const actual = graphemeClusters(clusters.join(''))
      if (JSON.stringify(actual) !== JSON.stringify(clusters)) throw new Error(`GraphemeBreakTest ${entry}: got ${actual.map(hex).join(' | ')}`)
      passed++
    }
    expect(passed).toBe(1093)
  })
  it('exposes full and Turkic case folding from the pinned table', () => {
    expect(fullCaseFold(0x41, false)).toBe('a')
    expect(fullCaseFold(0xdf, false)).toBe('ss')
    expect(fullCaseFold(0x1e9e, false)).toBe('ss')
    expect(fullCaseFold(0xfb01, false)).toBe('fi')
    expect(fullCaseFold(0x130, false)).toBe('i\u0307')
    expect(fullCaseFold(0x49, true)).toBe('ı')
    expect(fullCaseFold(0x130, true)).toBe('i')
    expect(fullCaseFold(0x61, false)).toBeUndefined()
    // Cherokee folds to upper case under Unicode case folding, unlike toLowerCase.
    expect(fullCaseFold(0xab70, false)).toBe('Ꭰ')
  })
  it('classifies grapheme properties, pictographs, conjunct breaks and ignorables', () => {
    expect(unicodeProperty(0x0d) & 15).toBe(1)
    expect(unicodeProperty(0x200d) & 15).toBe(5)
    expect(unicodeProperty(0x1f1e6) & 15).toBe(6)
    expect(unicodeProperty(0xac00) & 15).toBe(12)
    expect(unicodeProperty(0xac01) & 15).toBe(13)
    expect(unicodeProperty(0x1f600) & 16).toBe(16)
    expect((unicodeProperty(0x915) >> 5) & 3).toBe(1)
    expect((unicodeProperty(0x94d) >> 5) & 3).toBe(2)
    for (const point of [0xad, 0x34f, 0x200b, 0x200d, 0x2060, 0xfe0f, 0xfeff, 0xe0067]) expect(unicodeProperty(point) & 128).toBe(128)
    for (const point of [0x20, 0x61, 0x301, 0x1f3fd]) expect(unicodeProperty(point) & 128).toBe(0)
  })
})

describe('canonical search key', () => {
  it.each([
    ['É', 'e\u0301'], ['e\u0301', 'e\u0301'], ['ß', 'ss'], ['ẞ', 'ss'], ['ﬁ', 'fi'], ['Ǆ', '\u01c6'],
    ['a\u200b\u00ad\ufeff\u2060b', 'ab'], ['\u200b'.repeat(40), ''], ['ᾈ', 'α\u0313ι'], ['ẛ', 's\u0307'],
    ['\ud800A', '\ud800a'], ['한', '\u1112\u1161\u11ab'],
  ])('canonicalizes %j', (text, expected) => expect(hex(canonicalSearchKey(text, 'en-US'))).toBe(hex(expected)))
  it.each([
    ['tr', 'I', 'ı'], ['tr', 'İ', 'i'], ['tr', 'I\u0307', 'i'], ['tr', 'I\u0323\u0307', 'i\u0323'], ['tr', 'I\u0301\u0307', 'ı\u0301\u0307'],
    ['az', 'İSTANBUL', 'istanbul'], ['tr_TR', 'I', 'ı'], ['TR-tr', 'I', 'ı'], ['az-Latn-AZ', 'I', 'ı'],
    ['en-US', 'İ', 'i\u0307'], ['en-US', 'I', 'i'], ['de', 'I\u0307', 'i\u0307'], ['tk', 'I', 'i'],
  ])('applies locale-aware folding for %s %j', (locale, text, expected) => expect(hex(canonicalSearchKey(text, locale))).toBe(hex(expected)))
  it('recognizes only Turkish and Azeri language subtags as Turkic', () => {
    expect(['tr', 'tr-TR', 'az', 'AZ-cyrl', 'tr_TR'].map(isTurkicLocale)).toEqual([true, true, true, true, true])
    expect(['en', 'tk', 'trv', 'de-TR', ''].map(isTurkicLocale)).toEqual([false, false, false, false, false])
  })
  it('is idempotent and agrees with the spec-text oracle over hostile samples', () => {
    for (const locale of ['en-US', 'tr', 'da-DK']) {
      for (const text of ['ǅungla İI\u0307ı', 'Straße ẞ ﬃ', 'å a\u030a Å A\u030a', 'e\u0301\u0323 ệ', '👨\u200d👩\u200d👧 🇺🇸 👍🏽', '\ud800\u0301\udc00', 'a\u200b\u200b\u200bb', 'क\u094dष\u093f กำ']) {
        const key = canonicalSearchKey(text, locale)
        expect(canonicalSearchKey(key, locale)).toBe(key)
        expect(hex(key)).toBe(hex(oracleCanonical(text, locale)))
      }
    }
  })
  it('keeps raw values unchanged and rejects non-string input', () => {
    const raw = 'Straße'
    canonicalSearchKey(raw, 'de')
    expect(raw).toBe('Straße')
    expect(() => canonicalSearchKey(5 as unknown as string, 'en')).toThrow(RangeError)
  })
})

describe('verified differences from the collator implementation (amendment table)', () => {
  it('Danish a/A and aa/aA are both equal after folding', () => {
    expect(match('eq', 'a', 'A', 'da-DK')).toBe(true)
    expect(match('eq', 'aa', 'aA', 'da-DK')).toBe(true)
    expect(match('contains', 'aa', 'xaAx', 'da-DK')).toBe(true)
  })
  it('Danish å differs from aa', () => {
    expect(match('eq', 'å', 'aa', 'da-DK')).toBe(false)
    expect(match('contains', 'å', 'aa', 'da-DK')).toBe(false)
    expect(match('eq', 'å', 'a\u030a', 'da-DK')).toBe(true)
  })
  it('Danish å differs from a+ZWSP+a, which canonicalizes to aa', () => {
    expect(match('eq', 'å', 'a\u200ba', 'da-DK')).toBe(false)
    expect(match('contains', 'å', 'a\u200ba', 'da-DK')).toBe(false)
    expect(match('eq', 'aa', 'a\u200ba', 'da-DK')).toBe(true)
  })
  it('a and U+0363 (and e and U+0364) are unequal: marks never become base letters', () => {
    expect(match('eq', 'a', '\u0363')).toBe(false)
    expect(match('contains', 'a', '\u0363')).toBe(false)
    expect(match('contains', '\u0363', 'a')).toBe(false)
    expect(match('eq', 'e', '\u0364')).toBe(false)
  })
  it('ß equals ss and becomes two wildcard units', () => {
    expect(match('eq', 'ss', 'ß')).toBe(true)
    expect(match('eq', 'STRASSE', 'straße', 'de-DE')).toBe(true)
    expect(wild('??', 'ß')).toBe(true)
    expect(wild('?', 'ß')).toBe(false)
  })
  it('ﬁ equals fi but is two graphemes, so one ? fails', () => {
    expect(match('eq', 'ﬁ', 'fi')).toBe(true)
    expect(wild('?', 'ﬁ')).toBe(false)
    expect(wild('??', 'ﬁ')).toBe(true)
    expect(match('contains', 'f', 'ﬁ')).toBe(true)
  })
  it('contains(e) is false inside both é and e+U+0301', () => {
    expect(match('contains', 'e', 'é')).toBe(false)
    expect(match('contains', 'e', 'e\u0301')).toBe(false)
    expect(match('begins', 'e', 'e\u0301x')).toBe(false)
    expect(match('contains', 'é', 'xe\u0301x')).toBe(true)
  })
  it('? matches é and e+U+0301 as one canonical grapheme', () => {
    expect(wild('?', 'é')).toBe(true)
    expect(wild('?', 'e\u0301')).toBe(true)
    expect(wild('??', 'e\u0301')).toBe(false)
  })
  it('? does not match a lone U+200B, which canonicalizes to nothing', () => {
    expect(wild('?', '\u200b')).toBe(false)
    expect(wild('??', 'a\u200b')).toBe(false)
  })
})

describe('Turkish and Azeri dotted and dotless I', () => {
  it.each(['tr-TR', 'az'])('%s folds I to ı and İ (including decomposed I+U+0307) to i', locale => {
    expect(match('eq', 'İstanbul', 'istanbul', locale)).toBe(true)
    expect(match('eq', 'I\u0307stanbul', 'istanbul', locale)).toBe(true)
    expect(match('eq', 'I', 'ı', locale)).toBe(true)
    expect(match('eq', 'I', 'i', locale)).toBe(false)
    expect(match('contains', 'ı', 'DIŞ', locale)).toBe(true)
    expect(wild('?', 'İ', locale)).toBe(true)
  })
  it('other locales use the default folding', () => {
    expect(match('eq', 'I', 'i')).toBe(true)
    expect(match('eq', 'I', 'ı')).toBe(false)
    expect(match('eq', 'İ', 'i')).toBe(false)
    expect(match('eq', 'İ', 'i\u0307')).toBe(true)
  })
})

describe('emoji, flags, lone surrogates and ignorables', () => {
  it('keeps regional-indicator flags and skin-tone sequences as single graphemes', () => {
    expect(wild('?', '🇺🇸')).toBe(true)
    expect(match('contains', '🇸🇫', '🇺🇸🇫🇷')).toBe(false)
    expect(match('contains', '🇫🇷', '🇺🇸🇫🇷')).toBe(true)
    expect(wild('?', '👍🏽')).toBe(true)
    expect(match('contains', '👍', '👍🏽')).toBe(false)
    expect(wild('?', '1\ufe0f\u20e3')).toBe(true)
  })
  it('retains ZWJ, variation selectors and tag characters so emoji sequences stay single graphemes', () => {
    // Spec §7.1 (clarified 2026-10-11) exempts U+200D, U+FE00-FE0F,
    // U+E0100-E01EF and U+E0020-E007F from default-ignorable removal.
    const family = '\u{1f468}‍\u{1f469}‍\u{1f467}'
    expect(canonicalSearchKey(family, 'en')).toBe(family)
    expect(wild('?', family)).toBe(true)
    expect(wild('???', family)).toBe(false)
    expect(match('contains', '\u{1f469}', family)).toBe(false)
    expect(match('eq', family, '\u{1f468}\u{1f469}\u{1f467}')).toBe(false)
    expect(match('eq', '❤️', '❤')).toBe(false)
    expect(wild('?', '❤️')).toBe(true)
    const england = '\u{1f3f4}\u{e0067}\u{e0062}\u{e0065}\u{e006e}\u{e0067}\u{e007f}'
    expect(wild('?', england)).toBe(true)
    expect(match('eq', england, '\u{1f3f4}')).toBe(false)
    expect(match('contains', '\u{1f3f4}', england)).toBe(false)
    expect(canonicalSearchKey('辻\u{e0100}', 'en')).toBe('辻\u{e0100}')
    // Other default-ignorables are still removed.
    expect(canonicalSearchKey('a​‌­͏⁠﻿b', 'en')).toBe('ab')
  })
  it('treats a query whose canonical form is empty as no search term', () => {
    for (const op of ops) {
      for (const pattern of ['', '​', '­⁠']) {
        for (const wildcards of [false, true]) {
          for (const text of ['abc', '', '​', '\u{1f600}']) expect(matcher().match(op, pattern, wildcards, text)).toBe(true)
        }
      }
    }
    expect(wild('?', '​')).toBe(false)
    expect(wild('*', 'abc')).toBe(true)
  })
  it('treats a lone surrogate as one unit, even before a combining mark', () => {
    expect(wild('?', '\ud800')).toBe(true)
    expect(wild('?', '\ud800\u0301')).toBe(false)
    expect(wild('??', '\ud800\u0301')).toBe(true)
    expect(match('eq', '\ud800', '\ud800')).toBe(true)
    expect(match('eq', '\ud800', '\udc00')).toBe(false)
    expect(match('contains', '\ud83d', '😀')).toBe(false)
    expect(match('begins', '\ud83d', '😀')).toBe(false)
    expect(match('ends', '\ude00', '😀')).toBe(false)
    expect(match('contains', '\ud800', 'a\ud800b')).toBe(true)
    expect(graphemeClusters('\ud800\u0301\udc00')).toEqual(['\ud800', '\u0301', '\udc00'])
  })
  it('removes default-ignorables and their runs from text and patterns', () => {
    expect(match('eq', 'a', 'a' + '\u200b'.repeat(400))).toBe(true)
    expect(match('contains', 'l·', 'abl\u200b\u200b·cd')).toBe(true)
    expect(match('eq', 'a\u200bb', 'ab')).toBe(true)
    expect(wild('a?b', 'a\u200bb')).toBe(false)
    expect(wild('a?b', 'a\u200b\u200bb')).toBe(false)
    expect(wild('\u200b?', 'a')).toBe(true)
  })
})

describe('scripts and wildcard escapes', () => {
  it('matches Hangul syllables and conjoining jamo as the same graphemes', () => {
    expect(match('eq', '한국', '\u1112\u1161\u11ab\u1100\u116e\u11a8')).toBe(true)
    expect(wild('??', '한국')).toBe(true)
    expect(match('contains', '\u1112\u1161', '한')).toBe(false)
  })
  it('keeps Indic conjuncts and Thai spacing marks inside one grapheme', () => {
    expect(wild('?', 'क\u094dष\u093f')).toBe(true)
    expect(match('contains', 'क', 'क\u094dष\u093f')).toBe(false)
    expect(wild('?', 'กำ')).toBe(true)
    expect(match('contains', 'ก', 'กำ')).toBe(false)
  })
  it('keeps strict ~ escapes and tokenizes before canonicalizing', () => {
    expect(wild('~*~?~~', '*?~')).toBe(true)
    expect(wild('a~*b', 'A*B')).toBe(true)
    expect(() => wild('~x', 'x')).toThrow(RangeError)
    expect(() => wild('abc~', 'abc')).toThrow(RangeError)
    expect(() => wild('~\u200b*', 'x')).toThrow(RangeError)
    expect(match('eq', '*?~', '*?~')).toBe(true)
  })
  it('applies wildcards to begins, ends and contains', () => {
    expect(matcher().match('begins', 'S?', true, 'straße')).toBe(true)
    expect(matcher().match('ends', '?e', true, 'straße')).toBe(true)
    expect(matcher().match('contains', 'a*s', true, 'straße')).toBe(true)
    expect(matcher().match('contains', 'e?', true, 'cafe\u0301')).toBe(false)
    expect(matcher().match('not-contains', 'ss', true, 'STRASSE')).toBe(false)
  })
})

describe('production matcher has no host collator path', () => {
  it('never constructs Intl.Collator or Intl.Segmenter while matching', () => {
    const collator = vi.spyOn(Intl, 'Collator'), segmenter = vi.spyOn(Intl as unknown as { Segmenter: () => unknown }, 'Segmenter')
    try {
      const value = new TextFilterMatcher('da-DK')
      for (const op of ops) value.match(op, 'å*?', true, 'aa å straße 👍🏽')
      expect(collator).not.toHaveBeenCalled()
      expect(segmenter).not.toHaveBeenCalled()
      expect('collator' in value).toBe(false)
    } finally { vi.restoreAllMocks() }
  })
})

describe('cacheable search keys', () => {
  const mixed: ColumnClassification = { kind: 'text', hasNumeric: true }
  it('evaluates through a caller-supplied searchKeyOf port with identical results', () => {
    const cache = new Map<string, string>()
    let computed = 0
    const searchKeyOf = (text: string, locale: string) => {
      const key = `${locale}\u0000${text}`
      let value = cache.get(key)
      if (value === undefined) { computed++; value = canonicalSearchKey(text, locale); cache.set(key, value) }
      return value
    }
    const context = { locale: 'tr-TR', dateSystem: '1900' as const, timeZone: 'UTC', nowMs: 0 }
    const plain = new FilterEngine(context), cached = new FilterEngine({ ...context, searchKeyOf })
    const values: PreparedValue[] = ['İstanbul', 'ISTANBUL', 'istanbul', 'Straße'].map(value => ({ kind: 'string', value }))
    for (const rule of [{ op: 'eq', first: 'istanbul', wildcards: false }, { op: 'begins', first: '?s', wildcards: true }, { op: 'not-contains', first: 'ss', wildcards: false }] as const) {
      const a = plain.compileRule({ kind: 'text', ...rule }, mixed), b = cached.compileRule({ kind: 'text', ...rule }, mixed)
      for (let repeat = 0; repeat < 3; repeat++) for (const value of values) expect(b(value)).toBe(a(value))
    }
    expect(computed).toBe(values.length)
    expect(plain.compileRule({ kind: 'text', op: 'eq', first: 'istanbul', wildcards: false }, mixed)({ kind: 'number', value: 1, displayText: 'İSTANBUL' })).toBe(true)
  })
  it('matches a prepared predicate against canonical keys directly', () => {
    const predicate = matcher('de').prepare('contains', 'SS', false, new WorkBudget())
    expect(predicate.key(canonicalSearchKey('Straße', 'de'), new WorkBudget())).toBe(true)
    expect(predicate.key('strasse', new WorkBudget())).toBe(true)
    expect(predicate('Straße', new WorkBudget())).toBe(true)
  })
})

describe('key-level API for the S8.2 value search', () => {
  it('versions the key function by canonicalizer, Unicode data and fold variant', () => {
    expect(canonicalKeyStrategyId('default')).toBe(`canonical-text/v${CANONICAL_KEY_VERSION}/unicode-16.0.0/default`)
    expect(canonicalKeyStrategyId('turkic')).toBe(`canonical-text/v${CANONICAL_KEY_VERSION}/unicode-16.0.0/turkic`)
    expect(['en-US', 'tr', 'az-Latn', 'de'].map(foldVariantOf)).toEqual(['default', 'turkic', 'turkic', 'default'])
    expect(canonicalSearchKey('I', 'turkic')).toBe('ı')
    expect(canonicalSearchKey('I', 'default')).toBe('i')
    expect(canonicalSearchKey('I', 'tr-TR')).toBe(canonicalSearchKey('I', 'turkic'))
  })
  it('flags simple keys whose every code-unit position is a grapheme boundary', () => {
    for (const key of ['', 'abc', 'straße'.normalize('NFD'), '漢字', 'a\r', 'a\ud800b']) expect(isSimpleKey(key)).toBe(true)
    for (const key of ['é', '\u{1f600}', '\u{1f1fa}\u{1f1f8}', '\r\n', '하', 'กำ', 'a‍']) expect(isSimpleKey(key)).toBe(false)
  })
  it('checks grapheme boundaries on keys', () => {
    const key = canonicalSearchKey('Cafés \u{1f468}‍\u{1f469}', 'en')
    expect([0, 3, 4, 5, 6, 7].map(at => isGraphemeBoundary(key, at))).toEqual([true, true, false, true, true, true])
    expect(isGraphemeBoundary(key, key.length)).toBe(true)
    expect(isGraphemeBoundary(key, key.length - 2)).toBe(false)
    expect(isGraphemeBoundary('\u{1f600}', 1)).toBe(false)
    expect(() => isGraphemeBoundary(key, key.length + 1)).toThrow(RangeError)
  })
  it('applies the boundary rule rather than plain includes', () => {
    const key = canonicalSearchKey('é', 'en')
    expect(key.includes('e')).toBe(true)
    expect(keyContains(key, 'e')).toBe(false)
    expect(keyBegins(key, 'e')).toBe(false)
    expect(keyEnds(key, '́')).toBe(false)
    expect(keyContains(canonicalSearchKey('Straße', 'de'), canonicalSearchKey('SE', 'de'))).toBe(true)
    expect(keyContains('xáa', 'a')).toBe(true)
    expect(keyContains('abc', '')).toBe(true)
    expect(keyContains('é', 'e', true)).toBe(true) // caller asserted simple: no verification
    const wildcard = compileKeyMatcher('eq', 'caf?', true, 'default')
    expect(wildcard(canonicalSearchKey('CAFÉ', 'en'))).toBe(true)
    expect(compileKeyMatcher('contains', 'e', false, 'default')(key)).toBe(false)
  })
})

describe('linear cost and no resource-limit for ordinary text', () => {
  const pools: Record<string, readonly string[]> = {
    latin: Array.from('abcdefghijklmnopqrstuvwxyz ABCDEFGHIJ.,'),
    german: ['ß', 'Straße ', 'e\u0301', 'ü', 'u\u0308', 'ﬁ', ' '],
    danish: ['aa', 'å', 'A\u030a', 'ø', 'æ', ' '],
    cjk: Array.from({ length: 150 }, (_, index) => String.fromCodePoint(0x4e00 + index * 7)),
    emoji: ['👍🏽', '🇺🇸', '👨\u200d👩\u200d👧', '😀', ' ', '\u200b'],
  }
  const sample = (pool: readonly string[], length: number) => {
    let seed = 20261011, text = ''
    while (text.length < length) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
    return text.slice(0, length).replace(/[\ud800-\udbff]$/, 'x')
  }
  it.each(Object.keys(pools))('keeps %s text up to 4000 characters within default allowances with linear work', name => {
    const value = matcher(name === 'danish' ? 'da-DK' : 'en-US')
    const work: Record<string, number[]> = {}
    for (const length of [500, 1000, 2000, 4000]) {
      const text = sample(pools[name], length)
      for (const [label, op, pattern, wildcards] of [['contains', 'contains', 'zzz', false], ['eq', 'eq', text, false], ['wildcard', 'eq', '*a*?*e*z', true], ['hit', 'contains', text.slice(length / 2, length / 2 + 8), false]] as const) {
        const used = { used: 0 }
        expect(() => value.match(op, pattern, wildcards, text, { work: used })).not.toThrow()
        expect(used.used).toBeLessThan(DEFAULT_FILTER_WORK / 100)
        ;(work[label] ??= []).push(used.used)
      }
    }
    for (const series of Object.values(work)) for (let i = 1; i < series.length; i++) expect(series[i]).toBeLessThanOrEqual(series[i - 1] * 2.5 + 100)
  })
  it('keeps literal contains linear on periodic worst cases', () => {
    const used = { used: 0 }
    expect(matcher().match('contains', 'a'.repeat(999) + 'b', false, 'a'.repeat(20000), { work: used })).toBe(false)
    expect(used.used).toBeLessThan(200000)
    const wildcard = { used: 0 }
    expect(matcher().match('eq', '*a*a*a*a*a*b', true, 'a'.repeat(20000), { work: wildcard })).toBe(false)
    expect(wildcard.used).toBeLessThan(400000)
  })
  it('reports labelled non-browser per-value cost at 100/300/1000/2000 characters', () => {
    const rows: object[] = []
    const host = process.versions.bun ? `Bun ${process.versions.bun}` : `Node ${process.versions.node}`
    for (const name of ['latin', 'german', 'cjk']) {
      const value = matcher('en-US')
      for (const length of [100, 300, 1000, 2000]) {
        const text = sample(pools[name], length)
        for (const [label, op, pattern, wildcards] of [['contains-miss', 'contains', 'zzz', false], ['equals', 'eq', text.toUpperCase(), false], ['wildcard', 'eq', '*e*?*z', true]] as const) {
          const used = { used: 0 }
          value.match(op, pattern, wildcards, text, { work: used })
          const repeats = 200, at = performance.now()
          for (let i = 0; i < repeats; i++) value.match(op, pattern, wildcards, text)
          rows.push({ host: `${host} (non-browser)`, pool: name, length, label, work: used.used, usPerValue: Number(((performance.now() - at) * 1000 / repeats).toFixed(2)) })
        }
      }
    }
    console.info('S8.5 canonical text cost ' + JSON.stringify(rows))
    expect(rows.length).toBe(36)
  })
})

describe('differential: optimized matcher versus the spec-text oracle', () => {
  const pool = [
    'a', 'A', 'e', 'E', 'é', 'É', 'e\u0301', 'E\u0301', '\u0301', '\u0308', '\u0323', '\u0307', '\u0363', '\u0364', 'ö', 'o\u0308',
    'ß', 'ẞ', 's', 'S', 'ss', 'ﬁ', 'ﬂ', 'ﬀ', 'ﬃ', 'f', 'i', 'Ǆ', 'ǅ', 'ǆ', 'I', 'ı', 'İ', 'I\u0307', 'i\u0307',
    'å', 'Å', 'a\u030a', 'aa', 'æ', 'ø', '\u200b', '\u00ad', '\ufeff', '\u2060', '\u200c', '\u200d', '\u034f', '\ufe0f',
    '👨', '👩', '👧', '👨\u200d👩\u200d👧', '🇺', '🇸', '🇫', '🇷', '🇺🇸', '👍', '🏽', '👍🏽', '😀', '❤\ufe0f', '🏴',
    '한', '\u1112', '\u1161', '\u11ab', '가', 'क', '\u094d', 'ष', '\u093f', 'त', 'क\u094dष\u093f', 'ก', 'ำ', '\u0e31', '\u0e48', 'เ',
    '\ud800', '\udc00', '\ud83d', '\ude00', '*', '?', '~', ' ', 'x', 'Z',
  ]
  const locales = ['en-US', 'tr-TR', 'az', 'da-DK', 'de-DE', 'sv-SE', 'lt']
  it('agrees on at least 100,000 seeded inputs across all six operations', () => {
    const cases = Number(process.env.S85_DIFFERENTIAL_CASES ?? 100000)
    if (!Number.isSafeInteger(cases) || cases < 100000) throw new RangeError('Differential case count must be at least 100000')
    let seed = 8501, checked = 0
    const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
    const matchers = locales.map(locale => new TextFilterMatcher(locale))
    const started = performance.now()
    for (let index = 0; index < cases; index++) {
      const wildcards = index % 2 === 1
      let text = '', pattern = ''
      if (index % 50 === 0) text = pool[random(pool.length)].repeat(20 + random(40))
      else for (let i = random(11); i > 0; i--) text += pool[random(pool.length)]
      for (let i = random(6); i > 0; i--) {
        const point = pool[random(pool.length)]
        pattern += wildcards && point === '~' ? '~' + '*?~'[random(3)] : point
      }
      if (index % 7 === 3 && text) pattern = wildcards ? text.slice(random(text.length)).replace(/~/g, '~~') : text.slice(random(text.length))
      const which = index % locales.length, locale = locales[which]
      const key = canonicalSearchKey(text, locale), query = canonicalSearchKey(pattern, locale)
      for (const op of ops) {
        const actual = matchers[which].match(op, pattern, wildcards, text, { maxWork: Number.MAX_SAFE_INTEGER })
        const expected = oracleMatch(op, pattern, wildcards, text, locale)
        // Key path: searchKeyOf-style precomputed keys and the exported matchers.
        const viaKey = compileKeyMatcher(op, pattern, wildcards, foldVariantOf(locale))(key)
        const viaLiteral = wildcards ? viaKey : (op === 'contains' ? keyContains(key, query) : op === 'begins' ? keyBegins(key, query) : op === 'ends' ? keyEnds(key, query) : viaKey)
        if (actual !== expected || viaKey !== expected || viaLiteral !== expected) throw new Error(JSON.stringify({ index, locale, op, pattern: hex(pattern), text: hex(text), wildcards, actual, viaKey, viaLiteral, expected }))
        checked++
      }
    }
    expect(checked).toBe(cases * ops.length)
    console.info(`S8.5 differential: ${cases} seeded inputs, ${checked} comparisons, ${((performance.now() - started) / 1000).toFixed(1)} s, host=${process.versions.bun ? 'Bun ' + process.versions.bun : 'Node ' + process.versions.node}`)
  }, 600000)
  it('segments canonical pool text like the host Intl.Segmenter (test-only cross-check)', () => {
    let seed = 77, checked = 0
    const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
    for (let index = 0; index < 20000; index++) {
      let text = ''
      for (let i = random(12); i > 0; i--) text += pool[random(pool.length)]
      const canonical = canonicalSearchKey(text, 'en')
      if (/\p{Unassigned}/u.test(canonical)) continue
      expect(graphemeClusters(canonical)).toEqual(oracleGraphemes(canonical))
      checked++
    }
    console.info(`S8.5 segmenter cross-check: ${checked} strings, host Unicode ${process.versions.unicode ?? 'unknown'}`)
    expect(checked).toBeGreaterThan(19000)
  })
})
