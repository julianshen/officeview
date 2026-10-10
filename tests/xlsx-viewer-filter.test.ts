import { describe, expect, it } from 'vitest'
import { classifyColumn, filterMenus, rawValueKey, FilterEngine, type PreparedValue, type ColumnClassification } from '../src/xlsx/view/filter'
import type { ColumnRule, FilterSpec, NumberOp, TextOp } from '../src/xlsx/view/contracts'
import type { ResolvedSemantics } from '../src/xlsx/formula/types'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet } from '../src/xlsx/view/visibility'
const context = { locale: 'en-US', dateSystem: '1900' as const, timeZone: 'UTC', nowMs: Date.parse('2026-03-08T12:00:00Z') }
const semantics: ResolvedSemantics = { ...context, epochNowMs: context.nowMs, unicode: { version: 2, source: 'standalone-default' } }
const engine = new FilterEngine(context)
const mixed: ColumnClassification = { kind: 'text', hasNumeric: true }
const numeric: ColumnClassification = { kind: 'number', hasNumeric: true }
const date: ColumnClassification = { kind: 'date', hasNumeric: true }
const n = (value: number, displayText?: string): PreparedValue => ({ kind: 'number', value, displayText })
const s = (value: string): PreparedValue => ({ kind: 'string', value })
const range = { r0: 1, r1: 4, c0: 0, c1: 1, headerRows: 1 as const, totalsRows: 1 as const }
const spec = (...rules: readonly [number, ColumnRule][]): FilterSpec => ({ range, columns: new Map(rules) })
describe('typed filter evaluation', () => {
  it.each([
    ['eq', 3, true], ['eq', 4, false], ['neq', 4, true], ['neq', 3, false],
    ['gt', 4, true], ['gt', 3, false], ['gte', 3, true], ['gte', 2, false],
    ['lt', 2, true], ['lt', 3, false], ['lte', 3, true], ['lte', 4, false],
    ['between', 3, true], ['between', 5, true], ['between', 2, false], ['between', 6, false],
  ] as const)('evaluates number %s on numeric cells', (op, value, expected) => {
    const match = engine.compileRule({ kind: 'number', op, first: 3, second: 5 }, numeric)
    expect(match(n(value))).toBe(expected)
    expect(match(s(String(value)))).toBe(false)
    expect(match({ kind: 'blank' })).toBe(false)
    expect(match({ kind: 'error', value: '#NUM!' })).toBe(false)
    expect(match(n(value), { maxWork: 0 })).toBe(expected)
  })
  it('includes date serials in number conditions', () => expect(engine.compileRule({ kind: 'number', op: 'gt', first: 3 }, date)(n(46089.5))).toBe(true))
  it.each(['neq', 'not-contains'] as const)('complements text %s only within evaluable cells', op => {
    const match = engine.compileRule({ kind: 'text', op, first: '2', wildcards: false }, mixed)
    expect(match(s('other'))).toBe(true)
    for (const value of [{ kind: 'blank' }, { kind: 'error', value: '#DIV/0!' }] as PreparedValue[]) expect(match(value)).toBe(false)
  })
  it('uses the formatted display lane for mixed numeric/date/boolean values', () => {
    const match = engine.compileRule({ kind: 'text', op: 'contains', first: '2', wildcards: false }, mixed)
    expect(match(n(20, '20'))).toBe(true)
    expect(match(n(0.2, '20%'))).toBe(true)
    expect(match(n(46089, '2026-03-08'))).toBe(true)
    expect(match({ kind: 'boolean', value: true, displayText: 'TRUE2' })).toBe(true)
    expect(() => match(n(20))).toThrow(/display/i)
    expect(engine.compileRule({ kind: 'text', op: 'neq', first: '0', wildcards: false }, numeric)(n(20, '20'))).toBe(false)
  })
  it('evaluates dates by the column type including General numeric cells', () => {
    const rule: ColumnRule = { kind: 'date', op: 'eq', first: { year: 2026, month: 3, day: 8 } }
    expect(engine.compileRule(rule, date)(n(46089.5))).toBe(true)
    expect(engine.compileRule(rule, mixed)({ kind: 'number', value: 46089.5, numFmtId: 14, displayText: '2026-03-08' })).toBe(false)
    for (const value of [s('2026-03-08'), { kind: 'error', value: '#NUM!' }, { kind: 'blank' }] as PreparedValue[]) expect(engine.compileRule({ ...rule, op: 'neq' }, date)(value)).toBe(false)
    expect(engine.compileRule({ kind: 'datePreset', preset: 'today' }, date)(n(46089.5))).toBe(true)
  })
  it('supports same-value cross-type compound OR and AND', () => {
    const first: ColumnRule = { kind: 'number', op: 'gt', first: 3 }
    const second: ColumnRule = { kind: 'text', op: 'eq', first: 'yes', wildcards: false }
    const or = engine.compileRule({ kind: 'compound', connective: 'or', first, second }, mixed)
    expect(or(n(5, '5'))).toBe(true)
    expect(or(s('yes'))).toBe(true)
    expect(or(s('no'))).toBe(false)
    expect(engine.compileRule({ kind: 'compound', connective: 'and', first, second }, mixed)(s('yes'))).toBe(false)
    expect(engine.compileRule({ kind: 'compound', connective: 'and', first, second: { kind: 'number', op: 'lt', first: 6 } }, numeric)(n(5))).toBe(true)
  })
  it('ANDs columns and replays serialized rule entries identically', () => {
    const input = spec([0, { kind: 'number', op: 'gt', first: 3 }], [1, { kind: 'compound', connective: 'or', first: { kind: 'text', op: 'eq', first: 'yes', wildcards: false }, second: { kind: 'number', op: 'lt', first: 0 } }])
    const replay: FilterSpec = { range: input.range, columns: new Map(JSON.parse(JSON.stringify([...input.columns]))) }
    const rows = [[n(5), s('yes')], [n(2), s('yes')], [n(5), n(-1, '-1')], [n(5), s('no')]]
    const columns = new Map([[0, numeric], [1, mixed]])
    const read = (row: number, col: number) => rows[row][col]
    expect(rows.map((_, row) => engine.compileFilter(input, columns, read)(row))).toEqual([true, false, true, false])
    expect(rows.map((_, row) => engine.compileFilter(replay, columns, read)(row))).toEqual([true, false, true, false])
  })
  it('refuses unknown/unavailable values, loading classification and S8.2 membership', () => {
    const rule: ColumnRule = { kind: 'number', op: 'neq', first: 0 }
    for (const kind of ['unknown', 'unavailable'] as const) expect(() => engine.compileRule(rule, numeric)({ kind })).toThrow(/incomplete/i)
    expect(() => engine.compileRule(rule, { kind: 'loading', hasNumeric: false })).toThrow(/incomplete/i)
    expect(() => engine.compileRule({ kind: 'values', selection: { mode: 'include', keys: { refId: 'set', count: 1, owner: { docId: 'd', sheetId: 's', valueGen: 1 } } }, showBlanks: false }, mixed)).toThrow(/S8.2/)
  })
  it('snapshots compiled compound and column classification from mutable drafts', () => {
    const classification = { ...date }
    const compound: ColumnRule = { kind: 'compound', connective: 'or', first: { kind: 'number', op: 'gt', first: 0 }, second: { kind: 'number', op: 'lt', first: 0 } }
    const match = engine.compileRule(compound, classification)
    compound.connective = 'and'
    expect(match(n(5))).toBe(true)
    const dateMatch = engine.compileRule({ kind: 'datePreset', preset: 'today' }, classification)
    classification.kind = 'text'
    expect(dateMatch(n(46089.5))).toBe(true)
  })
  it('shares bounded column/row work and cancellation through filter projection', () => {
    const columns = new Map([[0, mixed], [1, mixed]])
    const input = spec([0, { kind: 'text', op: 'eq', first: 'a', wildcards: false }], [1, { kind: 'text', op: 'eq', first: 'a', wildcards: false }])
    const match = engine.compileFilter(input, columns, () => s('a'), { maxWork: 10 })
    expect(() => match(2)).toThrow(/resource/i)
    const controller = new AbortController()
    const baseHidden = createIntervalSet([{ r0: 2, r1: 2 }])
    const numericSpec = spec([0, { kind: 'number', op: 'gt', first: 0 }])
    let result: unknown
    expect(() => { result = engine.projectFilter(numericSpec, baseHidden, createIdentityOrder(), new Map([[0, numeric]]), row => { if (row === 3) controller.abort(); return n(5) }, 6, { signal: controller.signal }) }).toThrow(/cancel/i)
    expect(result).toBeUndefined()
    expect(baseHidden.contains(2)).toBe(true)
    expect(() => engine.compileRule({ kind: 'number', op: 'eq', first: 1 }, numeric)(n(1), { signal: controller.signal })).toThrow(/cancel/i)
  })
  it('rejects invalid numeric rules and prepared nonfinite values', () => {
    for (const first of [NaN, Infinity, -Infinity]) expect(() => engine.compileRule({ kind: 'number', op: 'eq', first }, numeric)).toThrow(RangeError)
    expect(() => engine.compileRule({ kind: 'number', op: 'between', first: 5, second: 3 }, numeric)).toThrow(/reversed/i)
    expect(() => engine.compileRule({ kind: 'number', op: 'between', first: 5 }, numeric)).toThrow(RangeError)
    expect(() => engine.compileRule({ kind: 'number', op: 'bad' as NumberOp, first: 5 }, numeric)).toThrow(RangeError)
    expect(() => engine.compileRule({ kind: 'number', op: 'eq', first: 5 }, numeric)(n(NaN))).toThrow(RangeError)
  })
})
describe('literal and wildcard Unicode semantics', () => {
  it.each([
    ['eq', 'A', 'a', true], ['eq', 'e', 'é', false], ['eq', '2', '02', false], ['eq', 'é', 'e\u0301', true],
    ['begins', 'é', 'e\u0301abc', true], ['ends', 'é', 'abce\u0301', true], ['contains', 'é', 'abe\u0301cd', true],
    ['contains', '\ud800', 'a\ud800b', true], ['begins', '\ud83d', '😀', false], ['ends', '\ude00', '😀', false],
    ['neq', 'é', 'e', true], ['not-contains', 'x', 'aaa', true],
  ] as const)('%s uses literal host collation at code-point boundaries', (op, pattern, text, expected) => expect(engine.matchText(op, pattern, false, text)).toBe(expected))
  it.each([
    ['*', '', true], ['?', '😀', true], ['?', '\ud800', true], ['??', '😀', false],
    ['~*~?~~', '*?~', true], ['a*b', 'ab', true], ['a*b', 'a😀b', true], ['a?b', 'a😀b', true],
    ['é?', 'e\u0301a', true], ['fi?', 'ﬁa', true], ['a?b', 'ab', false], ['é*', 'E\u0301a', true],
    ['\u200b?', 'a', true],
  ] as const)('wildcard %s respects code points and collated literal runs', (pattern, text, expected) => expect(engine.matchText('eq', pattern, true, text)).toBe(expected))
  it('uses wildcards with prefix/suffix/subsequence and negative operators', () => {
    for (const [op, text] of [['begins', 'ab-tail'], ['ends', 'tail-ab'], ['contains', 'tail-ab-tail']] as const) expect(engine.matchText(op, 'a?', true, text)).toBe(true)
    expect(engine.matchText('not-contains', 'a?', true, 'tail-ab-tail')).toBe(false)
    expect(engine.matchText('neq', 'a?', true, 'ab')).toBe(false)
  })
  it('rejects wildcard tilde escapes of ordinary characters', () => expect(() => engine.matchText('eq', '~x', true, 'x')).toThrow(RangeError))
  it('validates malformed trailing escape on compilation', () => expect(() => engine.compileRule({ kind: 'text', op: 'eq', first: 'abc~', wildcards: true }, mixed)).toThrow(RangeError))
  it('bounds pathological wildcard work and supports cancellation with explicit outcomes', () => {
    const work = { used: 0 }, text = 'a'.repeat(160)
    expect(engine.matchText('eq', '*a*a*a*a*a*b', true, text, { work, maxWork: 10000000 })).toBe(false)
    expect(work.used).toBeGreaterThan(1000)
    expect(work.used).toBeLessThan(10000000)
    expect(() => engine.matchText('contains', 'z', false, text, { maxWork: 1 })).toThrow(/resource/i)
    expect(() => engine.matchText('eq', 'a', false, text, { maxTextLength: 10 })).toThrow(/resource/i)
    const controller = new AbortController(); controller.abort()
    expect(() => engine.matchText('eq', 'a', false, 'a', { signal: controller.signal })).toThrow(/cancel/i)
    for (const maxWork of [NaN, Infinity, -1, 1.5]) expect(() => engine.matchText('eq', 'a', false, 'a', { maxWork })).toThrow(RangeError)
    for (const maxTextLength of [NaN, Infinity, -1, 1.5]) expect(() => engine.matchText('eq', 'a', false, 'a', { maxTextLength })).toThrow(RangeError)
    expect(() => engine.matchText('bad' as TextOp, 'a', false, 'a')).toThrow(RangeError)
  })
})
describe('complete classification and exact keys', () => {
  it('classifies all body values with supported date formats and builtins', () => {
    expect(classifyColumn([n(1), n(2)], true, semantics)).toEqual(numeric)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, numFmtId: 14 }, { kind: 'blank' }, { kind: 'error', value: '#NUM!' }], true, semantics)).toEqual(date)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, formatCode: 'yyyy-mm-dd' }], true, semantics)).toEqual(date)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, formatCode: 'h:mm' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0 "days"' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0;yyyy-mm-dd' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([n(2), s('s')], true, semantics)).toEqual(mixed)
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0;yyyy-mm-dd' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: '0;0;yyyy-mm-dd' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: 'yyyy-mm-dd;0' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0\\\\d' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: 'yyyy-mm-dd' }], true, { ...semantics, locale: 'zh-TW' })).toEqual(date)
    expect(classifyColumn([s('s'), { kind: 'boolean', value: true }], true, semantics)).toEqual({ kind: 'text', hasNumeric: false })
  })
  it('treats backslashes inside quoted format literals as shared grammar literal text', () => {
    const slash = String.fromCharCode(92)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '"' + slash + '"yyyy-mm-dd' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0 "' + slash + '";yyyy-mm-dd' }], true, semantics)).toEqual(date)
  })
  it('never infers date for empty/errors or completes missing domains', () => {
    for (const values of [[], [{ kind: 'blank' }], [{ kind: 'error', value: '#NUM!' }]] as PreparedValue[][]) expect(classifyColumn(values, true, semantics)).toEqual({ kind: 'empty', hasNumeric: false })
    expect(classifyColumn([n(1)], false, semantics).kind).toBe('loading')
    for (const kind of ['unknown', 'unavailable'] as const) expect(classifyColumn([n(1), { kind }], true, semantics).kind).toBe('loading')
    expect(() => classifyColumn([n(NaN)], true, semantics)).toThrow(RangeError)
  })
  it('exposes only available typed menus', () => {
    expect(filterMenus(numeric)).toEqual(['values', 'blanks', 'number'])
    expect(filterMenus(date)).toEqual(['values', 'blanks', 'date'])
    expect(filterMenus(mixed)).toEqual(['values', 'blanks', 'text', 'number'])
    expect(filterMenus({ kind: 'empty', hasNumeric: false })).toEqual(['values', 'blanks'])
    expect(filterMenus({ kind: 'loading', hasNumeric: true })).toEqual([])
  })
  it('keys raw types and preserves UTF-16, distinct serials and error-looking text', () => {
    const values = [n(-0, '0.00'), n(1.5, '1.50'), s('foo'), s('\ud800'), s('#DIV/0!'), { kind: 'boolean', value: true }, { kind: 'error', value: '#DIV/0!' }] as PreparedValue[]
    expect(values.map(value => rawValueKey(value, mixed))).toEqual(['n:0', 'n:1.5', 's:foo', 's:\ud800', 's:#DIV/0!', 'b:1', 'e:#DIV/0!'])
    expect(rawValueKey({ kind: 'boolean', value: false }, mixed)).toBe('b:0')
    expect([n(44927, '2023-01-01'), n(44927.5, '2023-01-01')].map(value => rawValueKey(value, date))).toEqual(['d:44927', 'd:44927.5'])
    expect(rawValueKey({ kind: 'blank' }, mixed)).toBe(null)
    expect(() => rawValueKey({ kind: 'unknown' }, mixed)).toThrow(/incomplete/i)
    expect(() => rawValueKey(n(Infinity), mixed)).toThrow(RangeError)
  })
  it('bounds complete-domain scanning', () => {
    expect(() => classifyColumn([n(1), n(2)], true, semantics, { maxWork: 1 })).toThrow(/resource/i)
    const controller = new AbortController(); controller.abort()
    expect(() => classifyColumn([n(1)], true, semantics, { signal: controller.signal })).toThrow(/cancel/i)
  })
})
describe('AC-SEM-01 visibility composition', () => {
  it('protects Amount header/totals, reveals hidden body, restores replacement and Clear snapshot', () => {
    const baseHidden = createIntervalSet([{ r0: 0, r1: 0 }, { r0: 2, r1: 2 }, { r0: 4, r1: 4 }])
    const rows = [n(9), s('Amount'), n(5), n(2), n(9), n(2)]
    const read = (row: number) => rows[row]
    const columns = new Map([[0, numeric]])
    const input = spec([0, { kind: 'number', op: 'gt', first: 3 }])
    const before = JSON.stringify(rows)
    const visible = (input: FilterSpec | null) => {
      const projection = engine.projectFilter(input, baseHidden, createIdentityOrder(), columns, read, rows.length)
      return rows.map((_, row) => projection.visibility.visible(row))
    }
    expect(visible(input)).toEqual([false, true, true, false, false, true])
    expect(visible({ ...input, range: { ...range, r0: 5, r1: 5, headerRows: 0, totalsRows: 0 } })).toEqual([false, true, false, true, false, false])
    expect(visible(null)).toEqual([false, true, false, true, false, true])
    expect(visible(spec())).toEqual(visible(null))
    expect(JSON.stringify(rows)).toBe(before)
  })
})
