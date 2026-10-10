import { describe, expect, it, vi } from 'vitest'
import { ReferenceTextFilterMatcher, TextFilterMatcher } from '../src/xlsx/view/textFilter'
import { classifyColumn, filterMenus, rawValueKey, FilterEngine, WorkBudget, type PreparedValue, type ColumnClassification, type FilterLimits } from '../src/xlsx/view/filter'
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
  it('bounds each column cell independently and retains projection cancellation', () => {
    const columns = new Map([[0, mixed], [1, mixed]])
    const input = spec([0, { kind: 'text', op: 'eq', first: 'a', wildcards: false }], [1, { kind: 'text', op: 'eq', first: 'a', wildcards: false }])
    const match = engine.compileFilter(input, columns, () => s('a'), { maxWork: 10 })
    expect(match(2)).toBe(true)
    const controller = new AbortController()
    const baseHidden = createIntervalSet([{ r0: 2, r1: 2 }])
    const numericSpec = spec([0, { kind: 'number', op: 'gt', first: 0 }])
    let result: unknown
    expect(() => { result = engine.projectFilter(numericSpec, baseHidden, createIdentityOrder(), new Map([[0, numeric]]), row => { if (row === 3) controller.abort(); return n(5) }, 6, { signal: controller.signal }) }).toThrow(/cancel/i)
    expect(result).toBeUndefined()
    expect(baseHidden.contains(2)).toBe(true)
    expect(() => engine.compileRule({ kind: 'number', op: 'eq', first: 1 }, numeric)(n(1), { signal: controller.signal })).toThrow(/cancel/i)
  })
  it('shares compile-time pattern limits, column work and cancellation with the operation', () => {
    const textSpec = spec([0, { kind: 'text', op: 'eq', first: 'a'.repeat(65000), wildcards: false }])
    expect(() => engine.compileFilter(textSpec, new Map([[0, mixed]]), () => s('a'), { maxTextLength: 1 })).toThrow(/resource/i)
  })
  it('charges compilation and evaluation to the same counter', () => {
    const work = { used: 0 }
    const compiled = engine.compileFilter(spec([0, { kind: 'text', op: 'eq', first: 'abc', wildcards: false }]), new Map([[0, mixed]]), () => s('abc'), { work })
    expect(work.used).toBeGreaterThanOrEqual(3)
    const before = work.used
    compiled(2)
    expect(work.used).toBeGreaterThan(before)
  })
  it('rejects cancellation of compilation and Clear projection', () => {
    const textSpec = spec([0, { kind: 'text', op: 'eq', first: 'a', wildcards: false }])
    const controller = new AbortController(); controller.abort()
    expect(() => engine.compileFilter(textSpec, new Map([[0, mixed]]), () => s('a'), { signal: controller.signal })).toThrow(/cancel/i)
    expect(() => engine.projectFilter(null, createIntervalSet([]), createIdentityOrder(), new Map(), () => n(1), 6, { signal: controller.signal })).toThrow(/cancel/i)
  })
  it('rejects compound children outside the two typed-condition contract', () => {
    const first: ColumnRule = { kind: 'number', op: 'gt', first: 0 }
    const nested: ColumnRule = { kind: 'compound', connective: 'or', first, second: first }
    const cyclic = { kind: 'compound', connective: 'or', first, second: first } as unknown as Record<string, unknown>
    cyclic.first = cyclic
    for (const invalid of [nested, { kind: 'values' }, null, 5, cyclic]) {
      const rule = { kind: 'compound', connective: 'or', first: invalid, second: first } as ColumnRule
      expect(() => engine.compileRule(rule, numeric)).toThrow(/typed condition/i)
    }
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
  it('rejects wildcard tilde escapes of ordinary characters consistently with the formula helper', () => expect(() => engine.matchText('eq', '~x', true, 'x')).toThrow(RangeError))
  it('validates malformed trailing escape on compilation', () => expect(() => engine.compileRule({ kind: 'text', op: 'eq', first: 'abc~', wildcards: true }, mixed)).toThrow(RangeError))
  it.each([null, {}, { used: null }, { used: undefined }, { used: NaN }, { used: -1 }, { used: 0.5 }])('rejects malformed supplied work counter %j', work => {
    expect(() => engine.matchText('eq', 'a', false, 'a', { work } as FilterLimits)).toThrow(RangeError)
  })
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
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0;yyyy-mm-dd' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: '0;0;yyyy-mm-dd' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: 'yyyy-mm-dd;0' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0\\\\d' }], true, semantics)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: 'yyyy-mm-dd' }], true, { ...semantics, locale: 'zh-TW' })).toEqual(date)
    expect(classifyColumn([s('s'), { kind: 'boolean', value: true }], true, semantics)).toEqual({ kind: 'text', hasNumeric: false })
  })
  it('treats backslashes inside quoted format literals as shared grammar literal text', () => {
    const slash = String.fromCharCode(92)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '"' + slash + '"yyyy-mm-dd' }], true, semantics)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0 "' + slash + '";yyyy-mm-dd' }], true, semantics)).toEqual(numeric)
  })
  it('preserves semicolons in inactive bracket regions during active-section classification', () => {
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: '0;[foo;yyyy];0' }], true, semantics)).toEqual(numeric)
  })
  it('honors an explicit empty custom format over a date builtin', () => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId: 14, formatCode: '' }], true, semantics)).toEqual(numeric)
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
  it('bounds each classification cell', () => {
    expect(() => classifyColumn([{ kind: 'number', value: 1, formatCode: '0.00' }], true, semantics, { maxWork: 1 })).toThrow(/resource/i)
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

describe('review regressions', () => {
  it('rejects null and missing-kind rules before property access', () => {
    for (const rule of [null, {}, undefined]) expect(() => engine.compileRule(rule as unknown as ColumnRule, numeric)).toThrow(RangeError)
  })
  it('uses per-cell defaults across one million numeric rows', () => {
    const match = engine.compileFilter(spec([0, { kind: 'number', op: 'gt', first: 0 }]), new Map([[0, numeric]]), () => n(1))
    let accepted = 0
    expect(() => { for (let row = 0; row < 1000000; row++) if (match(row)) accepted++ }).not.toThrow()
    expect(accepted).toBe(1000000)
  })
  it('classifies 200k custom-format numeric cells with per-cell defaults', () => {
    function* values() { for (let row = 0; row < 200000; row++) yield { kind: 'number' as const, value: row, formatCode: '0.000' } }
    expect(classifyColumn(values(), true, semantics)).toEqual(numeric)
  })
  it('repeated compiled calls and instrumentation never become implicit cumulative limits', () => {
    const work = { used: 0 }
    const match = engine.compileRule({ kind: 'text', op: 'eq', first: 'hello', wildcards: false }, mixed, { work, maxWork: 50 })
    expect(() => { for (let row = 0; row < 10000; row++) expect(match(s('hello'))).toBe(true) }).not.toThrow()
    expect(work.used).toBeGreaterThan(50)
  })
  it.each(['[$-409]mmm d, yyyy', '[$-F800]dddd, mmmm dd, yyyy', '[$-409]h:mm:ss AM/PM', 'yyyy年m月d日', '[Red]yyyy-mm-dd', '[$-411]ge.m.d', 'e', 'g', 'b', 'h:mm', '[Color10][>=0][$USD-409]yyyy-mm-dd'])('classifies authored locale/date format %s', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode }], true, semantics)).toEqual(date)
  })
  it.each(['[h]:mm', '[hh]', '[mm]:ss', '[ss]', 'mm:ss.0', 'mmss.0', 'General', '0.00E+00', '@', '0 \"days\"', '0_d', '0*y', '0\\m', '[foo]0'])('keeps duration/numeric/text format %s numeric', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode }], true, semantics)).toEqual(numeric)
  })
  it.each([...Array.from({ length: 9 }, (_, i) => 14 + i), ...Array.from({ length: 10 }, (_, i) => 27 + i), ...Array.from({ length: 9 }, (_, i) => 50 + i)])('recognizes date builtin %s', numFmtId => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId }], true, semantics)).toEqual(date)
  })
  it.each([45, 46, 47])('keeps duration builtin %s numeric', numFmtId => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId }], true, semantics)).toEqual(numeric)
  })
})

describe('explicit scan budgets', () => {
  it('stops a whole scan only when a cumulative budget is explicitly supplied', () => {
    const budget = new WorkBudget({ maxWork: 20 })
    const match = engine.compileFilter(spec([0, { kind: 'number', op: 'gt', first: 0 }]), new Map([[0, numeric]]), () => n(1), { budget })
    expect(match(2)).toBe(true)
    expect(() => { for (let row = 0; row < 50; row++) match(row) }).toThrow(/resource/i)
    expect(() => classifyColumn([n(1), n(2)], true, semantics, { budget: new WorkBudget({ maxWork: 1 }) })).toThrow(/resource/i)
  })
  it.each(['and', 'or'] as const)('shares one cell allowance for compound %s', connective => {
    const first: ColumnRule = { kind: 'text', op: 'eq', first: 'hello', wildcards: false }
    const match = engine.compileRule({ kind: 'compound', connective, first, second: first }, mixed)
    expect(engine.compileRule(first, mixed)(s('hello'), { maxWork: 15 })).toBe(true)
    expect(() => match(s('hello'), { maxWork: 15 })).toThrow(/resource/i)
    expect(match(s('hello'), { maxWork: 100 })).toBe(true)
  })
})

describe('prepared text predicates', () => {
  it('tokenizes wildcard patterns once when compiled rather than per cell', () => {
    const work = { used: 0 }
    const match = engine.compileRule({ kind: 'text', op: 'eq', first: '*'.repeat(2000), wildcards: true }, mixed, { work })
    const before = work.used
    expect(match(s(''))).toBe(true)
    expect(work.used - before).toBeLessThanOrEqual(2000)
    const again = work.used
    expect(match(s(''))).toBe(true)
    expect(work.used - again).toBeLessThanOrEqual(2000)
  })
})

describe('exact reference differential coverage', () => {
  it.each([false, true])('matches 20k seeded cases for all six operations (wildcards=%s)', wildcards => {
    const locales = ['en-US', 'de-DE', 'da-DK', 'tr-TR', 'cs-CZ', 'es-ES']
    const matchers = locales.map(locale => [new TextFilterMatcher(locale), new ReferenceTextFilterMatcher(locale)] as const)
    const ascii = Array.from({ length: 95 }, (_, i) => String.fromCharCode(32 + i))
    const unicode = [...ascii, 'é', 'e', '\u0301', '\u200b', '\u00ad', 'ﬁ', '😀', '\ud800', '\udc00', 'å', 'İ']
    const ops: TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
    let seed = wildcards ? 173 : 107, checked = 0
    const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
    for (let index = 0; index < 20000; index++) {
      const pool = index % 3 === 0 ? ascii : unicode
      let text = '', pattern = ''
      for (let i = random(8); i > 0; i--) text += pool[random(pool.length)]
      for (let i = random(6); i > 0; i--) {
        const point = pool[random(pool.length)]
        pattern += wildcards && point === '~' ? '~~' : point
      }
      const [fast, reference] = matchers[index % locales.length]
      for (const op of ops) {
        const actual = fast.match(op, pattern, wildcards, text), expected = reference.match(op, pattern, wildcards, text)
        if (actual !== expected) throw new Error(JSON.stringify({ index, locale: locales[index % locales.length], op, pattern, text, wildcards, actual, expected }))
        checked++
      }
    }
    expect(checked).toBe(120000)
  })
  it('retains arbitrarily long ignorable endpoint candidates and locale contractions', () => {
    for (const locale of ['en-US', 'da-DK', 'cs-CZ', 'tr-TR']) {
      const fast = new TextFilterMatcher(locale), reference = new ReferenceTextFilterMatcher(locale)
      for (const [pattern, text] of [['a', 'a' + '\u200b'.repeat(400)], ['å', 'aa'], ['å', 'a\u200ba'], ['ch', 'cH'], ['I', 'i']]) {
        for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
          for (const wildcards of [false, true]) expect(fast.match(op, pattern, wildcards, text, { maxWork: Number.MAX_SAFE_INTEGER })).toBe(reference.match(op, pattern, wildcards, text, { maxWork: Number.MAX_SAFE_INTEGER }))
        }
      }
    }
  })
})

describe('sound text shortcuts', () => {
  it('finds ASCII literal witnesses before scanning every endpoint', () => {
    const work = { used: 0 }
    expect(engine.matchText('contains', 'ZzZ', false, 'a'.repeat(1000) + 'zzz', { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(true)
    expect(work.used).toBeLessThan(10000)
  })
  it('shares exact repeated-substring comparisons across a periodic ASCII miss', () => {
    const work = { used: 0 }
    expect(engine.matchText('contains', 'zzz', false, 'a'.repeat(2000), { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(false)
    expect(work.used).toBeLessThan(2100000)
  })
  it('evaluates wildcard-free compiled rules with the literal matcher', () => {
    const work = { used: 0 }
    const match = engine.compileRule({ kind: 'text', op: 'eq', first: 'a'.repeat(2000), wildcards: true }, mixed, { work })
    const before = work.used
    expect(match(s('a'.repeat(2000)))).toBe(true)
    expect(work.used - before).toBeLessThan(10000)
  })
})

describe('complete ordinary text scans and periodic differential coverage', () => {
  it('filters 10k ordinary 200-character cells with default per-cell allowances', () => {
    const text = 'hello world '.repeat(17).slice(0, 200)
    const match = engine.compileFilter(spec([0, { kind: 'text', op: 'contains', first: 'zzz', wildcards: false }]), new Map([[0, mixed]]), () => s(text))
    let rejected = 0
    expect(() => { for (let row = 0; row < 10000; row++) if (!match(row)) rejected++ }).not.toThrow()
    expect(rejected).toBe(10000)
  })
  it('preserves tailored collation on every short raw period and wildcard literal run', () => {
    const patterns = ['a', 'A', 'aa', 'å', 'ch', 'I', 'i', 'ﬁ', 'fi', 'zzz', '\u200b', '*a?', '?i*', 'a*b', '~*', '*']
    for (const locale of ['en-US', 'da-DK', 'tr-TR', 'cs-CZ']) {
      const fast = new TextFilterMatcher(locale), reference = new ReferenceTextFilterMatcher(locale)
      for (let period = 1; period <= 16; period++) {
        const text = 'aaIich0*?~ zABCD'.slice(0, period).repeat(3).slice(0, period * 3 - 1)
        for (const pattern of patterns) {
          for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
            for (const wildcards of [false, true]) expect(fast.match(op, pattern, wildcards, text)).toBe(reference.match(op, pattern, wildcards, text))
          }
        }
      }
    }
  })
})

describe('certified ASCII backend work', () => {
  it('handles a nonperiodic ASCII miss with linear candidate work on the pinned backend', () => {
    let seed = 79, text = ''
    const pool = 'abcdefghijklmnopqrstuvwxy 0123456789!'
    for (let i = 0; i < 1000; i++) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
    const work = { used: 0 }
    expect(engine.matchText('contains', 'zzz', false, text, { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(false)
    expect(work.used).toBeLessThan(10000)
  })
  it('propagates all ignorable endpoint plateaus without quadratic wildcard work', () => {
    const work = { used: 0 }, text = 'a' + '\u200b'.repeat(2000) + 'b'
    expect(engine.matchText('eq', 'a*b', true, text, { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(true)
    expect(work.used).toBeLessThan(30000)
    expect(engine.matchText('eq', 'a?b', true, 'a\u200bb')).toBe(true)
    expect(engine.matchText('eq', 'a?b', true, 'a\u200b\u200bb')).toBe(true)
  })
})

describe('compilation and certificate boundaries', () => {
  it('bounds all columns of a filter compilation with one per-call allowance', () => {
    const first: ColumnRule = { kind: 'text', op: 'eq', first: 'a'.repeat(100), wildcards: false }
    expect(() => engine.compileFilter(spec([0, first], [1, first]), new Map([[0, mixed], [1, mixed]]), () => s('a'), { maxWork: 300 })).toThrow(/resource/i)
  })
  it('falls back on unknown backend versions and browser hosts', () => {
    const original = process
    try {
      vi.stubGlobal('process', { ...original, versions: { ...original.versions, icu: '999' } })
      expect(new TextFilterMatcher('en-US').certifiedAsciiProfile).toBe(false)
      vi.stubGlobal('process', undefined)
      expect(new TextFilterMatcher('en-US').certifiedAsciiProfile).toBe(false)
    } finally { vi.unstubAllGlobals() }
  })
  it.each(['NODE_ICU_DATA', 'ICU_DATA', 'NODE_OPTIONS'])('falls back on explicit custom ICU configuration %s', variable => {
    try {
      vi.stubEnv(variable, variable === 'NODE_OPTIONS' ? '--icu-data-dir=/custom' : '/custom')
      expect(new TextFilterMatcher('en-US').certifiedAsciiProfile).toBe(false)
    } finally { vi.unstubAllEnvs() }
  })
  it('falls back on tailored locales and substituted Intl constructors', () => {
    for (const locale of ['da-DK', 'tr-TR', 'cs-CZ', 'de-DE']) expect(new TextFilterMatcher(locale).certifiedAsciiProfile).toBe(false)
    const Native = Intl.Collator
    try {
      vi.spyOn(Intl, 'Collator').mockImplementation(function (locale, options) { return new Native(locale, options) })
      expect(new TextFilterMatcher('en-US').certifiedAsciiProfile).toBe(false)
    } finally { vi.restoreAllMocks() }
  })
})
