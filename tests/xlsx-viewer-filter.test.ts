import { describe, expect, it, vi } from 'vitest'
import { DEFAULT_FILTER_WORK, ReferenceTextFilterMatcher, TextFilterMatcher } from '../src/xlsx/view/textFilter'
import { classifyColumn, filterMenus, rawValueKey, FilterEngine, WorkBudget, type PreparedValue, type ColumnClassification, type FilterLimits } from '../src/xlsx/view/filter'
import type { ColumnRule, FilterSpec, NumberOp, TextOp } from '../src/xlsx/view/contracts'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet } from '../src/xlsx/view/visibility'
const context = { locale: 'en-US', dateSystem: '1900' as const, timeZone: 'UTC', nowMs: Date.parse('2026-03-08T12:00:00Z') }
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
    expect(classifyColumn([n(1), n(2)], true)).toEqual(numeric)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, numFmtId: 14 }, { kind: 'blank' }, { kind: 'error', value: '#NUM!' }], true)).toEqual(date)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, formatCode: 'yyyy-mm-dd' }], true)).toEqual(date)
    expect(classifyColumn([n(1), { kind: 'number', value: 2, formatCode: 'h:mm' }], true)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0 "days"' }], true)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0;yyyy-mm-dd' }], true)).toEqual(numeric)
    expect(classifyColumn([n(2), s('s')], true)).toEqual(mixed)
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0;yyyy-mm-dd' }], true)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: '0;0;yyyy-mm-dd' }], true)).toEqual(numeric)
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: 'yyyy-mm-dd;0' }], true)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '0\\\\d' }], true), 'unescaped d after an escaped backslash is a definitive date token').toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: 'yyyy-mm-dd' }], true)).toEqual(date)
    expect(classifyColumn([s('s'), { kind: 'boolean', value: true }], true)).toEqual({ kind: 'text', hasNumeric: false })
  })
  it('treats backslashes inside quoted format literals as shared grammar literal text', () => {
    const slash = String.fromCharCode(92)
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode: '"' + slash + '"yyyy-mm-dd' }], true)).toEqual(date)
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode: '0 "' + slash + '";yyyy-mm-dd' }], true)).toEqual(numeric)
  })
  it('preserves semicolons in inactive bracket regions during active-section classification', () => {
    expect(classifyColumn([{ kind: 'number', value: 0, formatCode: '0;[foo;yyyy];0' }], true)).toEqual(numeric)
  })
  it('honors an explicit empty custom format over a date builtin', () => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId: 14, formatCode: '' }], true)).toEqual(numeric)
  })
  it('never infers date for empty/errors or completes missing domains', () => {
    for (const values of [[], [{ kind: 'blank' }], [{ kind: 'error', value: '#NUM!' }]] as PreparedValue[][]) expect(classifyColumn(values, true)).toEqual({ kind: 'empty', hasNumeric: false })
    expect(classifyColumn([n(1)], false).kind).toBe('loading')
    for (const kind of ['unknown', 'unavailable'] as const) expect(classifyColumn([n(1), { kind }], true).kind).toBe('loading')
    expect(() => classifyColumn([n(NaN)], true)).toThrow(RangeError)
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
    expect(() => classifyColumn([{ kind: 'number', value: 1, formatCode: '0.00' }], true, { maxWork: 1 })).toThrow(/resource/i)
    const controller = new AbortController(); controller.abort()
    expect(() => classifyColumn([n(1)], true, { signal: controller.signal })).toThrow(/cancel/i)
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
    expect(classifyColumn(values(), true)).toEqual(numeric)
  })
  it('repeated compiled calls and instrumentation never become implicit cumulative limits', () => {
    const work = { used: 0 }
    const match = engine.compileRule({ kind: 'text', op: 'eq', first: 'hello', wildcards: false }, mixed, { work, maxWork: 50 })
    expect(() => { for (let row = 0; row < 10000; row++) expect(match(s('hello'))).toBe(true) }).not.toThrow()
    expect(work.used).toBeGreaterThan(50)
  })
  it.each(['[$-409]mmm d, yyyy', '[$-F800]dddd, mmmm dd, yyyy', '[$-409]h:mm:ss AM/PM', 'yyyy年m月d日', '[Red]yyyy-mm-dd', '[$-411]ge.m.d', 'e', 'g', 'b', 'h:mm', '[Color10][>=0][$USD-409]yyyy-mm-dd'])('classifies authored locale/date format %s', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: -2, formatCode }], true)).toEqual(date)
  })
  it.each(['[h]:mm', '[hh]', '[mm]:ss', '[ss]', 'mm:ss.0', 'mmss.0', 'General', '0.00E+00', '@', '0 \"days\"', '0_d', '0*y', '0\\m', '[foo]0'])('keeps duration/numeric/text format %s numeric', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode }], true)).toEqual(numeric)
  })
  it.each([...Array.from({ length: 9 }, (_, i) => 14 + i), ...Array.from({ length: 10 }, (_, i) => 27 + i), ...Array.from({ length: 9 }, (_, i) => 50 + i)])('recognizes date builtin %s', numFmtId => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId }], true)).toEqual(date)
  })
  it.each([45, 46, 47])('keeps duration builtin %s numeric', numFmtId => {
    expect(classifyColumn([{ kind: 'number', value: 2, numFmtId }], true)).toEqual(numeric)
  })
})

describe('explicit scan budgets', () => {
  it('stops a whole scan only when a cumulative budget is explicitly supplied', () => {
    const budget = new WorkBudget({ maxWork: 20 })
    const match = engine.compileFilter(spec([0, { kind: 'number', op: 'gt', first: 0 }]), new Map([[0, numeric]]), () => n(1), { budget })
    expect(match(2)).toBe(true)
    expect(() => { for (let row = 0; row < 50; row++) match(row) }).toThrow(/resource/i)
    expect(() => classifyColumn([n(1), n(2)], true, { budget: new WorkBudget({ maxWork: 1 }) })).toThrow(/resource/i)
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
  it.each([false, true])('matches seeded cases for all six operations (wildcards=%s)', wildcards => {
    const cases = Number(process.env.S8_FILTER_DIFFERENTIAL_CASES ?? 2000)
    if (!Number.isSafeInteger(cases) || cases < 2000) throw new RangeError('Differential case count must be at least 2000')
    const locales = ['en-US', 'de-DE', 'sv-SE', 'tr-TR', 'da-DK', 'es-ES-u-co-trad', 'cs-CZ']
    const matchers = locales.map(locale => [new TextFilterMatcher(locale), new ReferenceTextFilterMatcher(locale)] as const)
    const ascii = Array.from({ length: 95 }, (_, i) => String.fromCharCode(32 + i))
    const unicode = [...ascii, 'é', 'e', '\u0301', '\u200b', '\u00ad', '\ufeff', '\u2060', 'ﬁ', 'ß', '😀', '\u200d', '漢', '\ud800', '\udc00', 'å', 'İ']
    const ops: TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
    let seed = wildcards ? 173 : 107, checked = 0
    const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
    for (let index = 0; index < cases; index++) {
      const pool = index % 3 === 0 ? ascii : unicode
      let text = '', pattern = ''
      for (let i = random(8); i > 0; i--) text += pool[random(pool.length)]
      for (let i = random(6); i > 0; i--) {
        const point = pool[random(pool.length)]
        pattern += wildcards && point === '~' ? '~~' : point
      }
      if (index % 2000 === 0) text = ['é', 'a', '漢😀', '\ud800', 'e\u0301\u200b'][Math.floor(index / 2000) % 5].repeat(32)
      const [fast, reference] = matchers[index % locales.length]
      for (const op of ops) {
        const actual = fast.match(op, pattern, wildcards, text), expected = reference.match(op, pattern, wildcards, text)
        if (actual !== expected) throw new Error(JSON.stringify({ index, locale: locales[index % locales.length], op, pattern, text, wildcards, actual, expected }))
        checked++
      }
    }
    expect(checked).toBe(cases * ops.length)
    console.info(`S8 differential: ${cases} seeded inputs, ${checked} comparisons, wildcards=${wildcards}, host=${process.versions.bun ? 'Bun ' + process.versions.bun : 'Node ' + process.versions.node}`)
  }, process.env.S8_FILTER_DIFFERENTIAL_CASES ? 600000 : undefined)
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
  it('uses the ASCII optimization on the Bun or Node host running this suite', () => {
    const matcher = new TextFilterMatcher('en-US')
    matcher.match('contains', 'zzz', false, 'hello')
    expect(matcher.certifiedAsciiProfile).toBe(true)
  })
  it('handles a nonperiodic ASCII miss with linear candidate work on a certified host', ({ skip }) => {
    const matcher = new TextFilterMatcher('en-US')
    let seed = 79, text = ''
    const pool = 'abcdefghijklmnopqrstuvwxy 0123456789!'
    for (let i = 0; i < 1000; i++) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
    const work = { used: 0 }
    expect(matcher.match('contains', 'zzz', false, text, { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(false)
    if (!matcher.certifiedAsciiProfile) skip('Actual alphabet exceeded the bounded certificate work cap; reference retained')
    // Includes the cold observed-adjacency certificate: about 120 units/char.
    expect(work.used).toBeLessThan(150000)
  })
  it('propagates single ignorables linearly and keeps runs exact through the reference', () => {
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    const work = { used: 0 }, text = 'a\u200b'.repeat(1000) + 'b'
    expect(matcher.match('eq', 'a*b', true, text, { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(true)
    expect(matcher.certifiedAsciiProfile).toBe(true)
    expect(work.used).toBeLessThan(150000)
    // A run of two or more ignorables is outside the certificate window, so it
    // must retain the exact (uncertified) reference path.
    const run = 'a' + '\u200b'.repeat(40) + 'b'
    for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
      expect(matcher.match(op, 'a*b', true, run, { maxWork: Number.MAX_SAFE_INTEGER })).toBe(reference.match(op, 'a*b', true, run, { maxWork: Number.MAX_SAFE_INTEGER }))
    }
    expect(matcher.certifiedAsciiProfile).toBe(false)
    expect(engine.matchText('eq', 'a?b', true, 'a\u200bb')).toBe(true)
    expect(engine.matchText('eq', 'a?b', true, 'a\u200b\u200bb')).toBe(true)
  })
})

describe('compilation and certificate boundaries', () => {
  it('bounds all columns of a filter compilation with one per-call allowance', () => {
    const first: ColumnRule = { kind: 'text', op: 'eq', first: 'a'.repeat(100), wildcards: false }
    expect(() => engine.compileFilter(spec([0, first], [1, first]), new Map([[0, mixed], [1, mixed]]), () => s('a'), { maxWork: 300 })).toThrow(/resource/i)
  })
  it.each([undefined, {}, { versions: { icu: '999' } }])('certifies actual alphabets without process metadata', shim => {
    try {
      vi.stubGlobal('process', shim)
      const matcher = new TextFilterMatcher('de-DE'), work = { used: 0 }
      expect(matcher.match('contains', 'zzz', false, 'é漢😀'.repeat(667), { work })).toBe(false)
      expect(matcher.certifiedAsciiProfile).toBe(true)
      expect(work.used).toBeLessThan(100000)
    } finally { vi.unstubAllGlobals() }
  })
  it('fails closed on substituted Intl constructors', () => {
    const Native = Intl.Collator
    try {
      vi.spyOn(Intl, 'Collator').mockImplementation(function (locale, options) { return new Native(locale, options) })
      expect(new TextFilterMatcher('en-US').certifiedAsciiProfile).toBe(false)
    } finally { vi.restoreAllMocks() }
  })
})

describe('zero-weight wildcard literals never rewind raw endpoints', () => {
  it.each(['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[])('preserves forward consumption for %s', op => {
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    const pattern = '?\u200b?', text = '\u200b'
    expect(matcher.match(op, pattern, true, text)).toBe(reference.match(op, pattern, true, text))
  })
  it('matches exhaustive tiny-alphabet questions, ignored literals and stars', () => {
    const strings = (alphabet: readonly string[], maximum: number) => {
      let level = [''], result = ['']
      for (let size = 1; size <= maximum; size++) { level = level.flatMap(prefix => alphabet.map(point => prefix + point)); result = result.concat(level) }
      return result
    }
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    const patterns = strings(['?', '*', '\u200b', 'a'], 5), texts = strings(['\u200b', 'a'], 4)
    let checked = 0
    for (const pattern of patterns) {
      for (const text of texts) {
        for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
          const actual = matcher.match(op, pattern, true, text), expected = reference.match(op, pattern, true, text)
          if (actual !== expected) throw new Error(JSON.stringify({ op, pattern, text, actual, expected }))
          checked++
        }
      }
    }
    expect(checked).toBe(253890)
  })
  it('keeps ignored endpoint propagation forward through long plateaus', () => {
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    for (const size of [1, 2, 3, 25, 100, 401]) {
      const text = '\u200b'.repeat(size)
      for (const pattern of ['?\u200b?', '??\u200b?', '?\u200b??', '?\u200b?\u200b?', 'a?\u200b?b']) {
        for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
          expect(matcher.match(op, pattern, true, text, { maxWork: Number.MAX_SAFE_INTEGER })).toBe(reference.match(op, pattern, true, text, { maxWork: Number.MAX_SAFE_INTEGER }))
        }
      }
    }
  })
})

describe('partial browser process shims do not govern certificates', () => {
  it.each([
    { versions: {} },
    { versions: {}, env: {} },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' } },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' }, env: {} },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' }, env: {}, execArgv: null },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' }, env: {}, execArgv: {} },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' }, env: {}, execArgv: [null] },
    { versions: { bun: '1.3.10', icu: '74.2', unicode: '15.1' }, env: null, execArgv: [] },
    { get versions(): never { throw new Error('metadata unavailable') } },
  ])('retains exact fallback for incomplete/malformed metadata', shim => {
    try {
      vi.stubGlobal('process', shim)
      const matcher = new TextFilterMatcher('en-US')
      expect(matcher.match('contains', 'zzz', false, 'é漢😀'.repeat(667))).toBe(false)
      expect(matcher.certifiedAsciiProfile).toBe(true)
      expect(matcher.match('eq', 'é', false, 'e\u0301')).toBe(true)
      expect(matcher.match('eq', '?\u200b?', true, '\u200b')).toBe(false)
    } finally { vi.unstubAllGlobals() }
  })
})

describe('fallback text preflight precedes normalization', () => {
  it('rejects oversized fallback text before making normalized copies', () => {
    const matcher = new TextFilterMatcher('da-DK'), text = 'x'.repeat(1000)
    const copy = vi.spyOn(String.prototype, 'toLowerCase')
    try {
      expect(() => matcher.match('contains', 'x', false, text, { maxTextLength: 10 })).toThrow(/resource/i)
      expect(copy.mock.contexts.some(value => typeof value === 'string' && value.length === text.length)).toBe(false)
    } finally { copy.mockRestore() }
  })
  it('checks cancellation before oversized reference text processing', () => {
    const controller = new AbortController()
    const evaluation = new WorkBudget({ signal: controller.signal, maxTextLength: 10 })
    controller.abort()
    const matcher = new ReferenceTextFilterMatcher('da-DK')
    const evaluate = matcher.prepare('contains', 'x', false, new WorkBudget())
    expect(() => evaluate('x'.repeat(1000), evaluation)).toThrow(/cancel/i)
  })
})

describe('definitive authored date tokens override numeric placeholders', () => {
  it('keeps a day token escaped by one runtime backslash numeric', () => {
    const formatCode = '0' + String.fromCharCode(92) + 'd'
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode }], true)).toEqual(numeric)
  })
  it.each(['yyyy-mm-dd 0', 'h:mm:ss 0'])('classifies definitive date/time format %s', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: 2, formatCode }], true)).toEqual(date)
  })
})


describe('classification per-value allowances', () => {
  it('allows classification policy as the third argument and resets the per-value allowance', () => {
    const work = { used: 0 }
    expect(classifyColumn([n(1), n(2), n(3)], true, { maxWork: 1, work })).toEqual(numeric)
    expect(work.used).toBe(3)
  })
})


describe('duration and time-of-day classification decision', () => {
  // R2 decision: mm:ss[.0] mirrors duration/fraction built-ins 45–47;
  // h:mm denotes time of day and remains date-typed.
  it.each(['mm:ss.0', 'mm:ss'])('keeps %s numeric', formatCode => {
    expect(classifyColumn([{ kind: 'number', value: 0.5, formatCode }], true)).toEqual(numeric)
  })
  it('keeps h:mm date-typed', () => {
    expect(classifyColumn([{ kind: 'number', value: 0.5, formatCode: 'h:mm' }], true)).toEqual(date)
  })
})


describe('exact Unicode witnesses and repeated slices', () => {
  it('finds a raw Unicode witness before exhaustive slice enumeration', () => {
    const matcher = new TextFilterMatcher('da-DK'), work = { used: 0 }
    const text = 'àé漢😀'.repeat(200) + '東京'
    expect(matcher.match('contains', '東京', false, text, { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(true)
    expect(work.used).toBeLessThan(10000)
  })
  it('shares byte-identical periodic Unicode slices without assuming collation weights', () => {
    const matcher = new TextFilterMatcher('da-DK'), work = { used: 0 }
    expect(matcher.match('contains', 'zzz', false, 'é'.repeat(2000), { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(false)
    expect(work.used).toBeLessThan(2100000)
  })
  it('shares periodic wildcard literal comparisons across code-point phases', () => {
    const matcher = new TextFilterMatcher('da-DK'), work = { used: 0 }
    expect(matcher.match('eq', '*a*a*a*b', true, 'é'.repeat(500), { work, maxWork: Number.MAX_SAFE_INTEGER })).toBe(false)
    expect(work.used).toBeLessThan(300000)
  })
  it('never treats a lone-surrogate needle inside a paired surrogate as a witness', () => {
    const matcher = new TextFilterMatcher('da-DK'), reference = new ReferenceTextFilterMatcher('da-DK')
    for (const op of ['begins', 'ends', 'contains'] as const) {
      for (const pattern of ['\ud83d', '\ude00']) expect(matcher.match(op, pattern, false, '😀')).toBe(reference.match(op, pattern, false, '😀'))
    }
  })
})


describe('actual substituted collator identity', () => {
  it('fails closed for a contextual comparator despite native-looking options', () => {
    const Native = Intl.Collator
    try {
      vi.spyOn(Intl, 'Collator').mockImplementation(function (locale, options) {
        const collator = new Native(locale, options), compare = collator.compare
        // All <=3-element strings retain native behavior. A finite probe
        // domain cannot establish the rule for longer contextual rewrites.
        Object.defineProperty(collator, 'compare', {
          value: (a: string, b: string) => compare(a.replace(/abcd/g, 'wxyz'), b.replace(/abcd/g, 'wxyz')),
        })
        return collator
      })
      const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
      expect(matcher.certifiedAsciiProfile).toBe(false)
      for (const wildcards of [false, true]) {
        for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
          expect(matcher.match(op, 'wxyz', wildcards, 'abcd')).toBe(reference.match(op, 'wxyz', wildcards, 'abcd'))
        }
      }
      expect(matcher.match('eq', 'wxyz', false, 'abcd')).toBe(true)
    } finally { vi.restoreAllMocks() }
  })
})


describe('actual alphabet certification work', () => {
  it.each(['en-US', 'de-DE', 'sv-SE', 'tr-TR', 'da-DK', 'es-ES-u-co-trad'])('certifies nonperiodic Unicode literals and raw wildcard endpoints in %s', locale => {
    const matcher = new TextFilterMatcher(locale), work = { used: 0 }
    let seed = 317, text = ''
    const pool = ['é', 'à', '漢', '東', '😀', '😄']
    for (let i = 0; i < 2000; i++) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
    expect(matcher.match('contains', 'zzz', false, text, { work })).toBe(false)
    expect(matcher.certifiedAsciiProfile).toBe(true)
    expect(work.used).toBeLessThan(100000)
    expect(matcher.match('eq', '*é*é*zzz', true, text)).toBe(false)
    expect(matcher.certifiedAsciiProfile).toBe(true)
  })
  it('fails closed for expansion and canonical mark alphabets, preserving the oracle', () => {
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    for (const [pattern, text] of [['fi', 'ﬁ'], ['ss', 'ß'], ['é', 'e\u0301'], ['å', 'aa']]) {
      for (const wildcards of [false, true]) for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
        expect(matcher.match(op, pattern, wildcards, text)).toBe(reference.match(op, pattern, wildcards, text))
      }
    }
  })
  it('charges certificate probes to the actual cell and whole-scan allowance', () => {
    const matcher = new TextFilterMatcher('sv-SE')
    expect(() => matcher.match('contains', 'zzz', false, 'é漢😀'.repeat(667), { budget: new WorkBudget({ maxWork: 50 }) })).toThrow(/resource/i)
    expect(() => matcher.match('contains', 'zzz', false, 'é漢😀'.repeat(667), { maxWork: 50 })).toThrow(/resource/i)
  })
  it('never reuses a certificate after replacing the actual comparator', () => {
    const matcher = new TextFilterMatcher('en-US')
    matcher.match('contains', 'zzz', false, 'abc')
    const compare = matcher.collator.compare
    Object.defineProperty(matcher.collator, 'compare', { value: (a: string, b: string) => compare(a.replace(/abcd/g, 'wxyz'), b.replace(/abcd/g, 'wxyz')) })
    expect(matcher.match('contains', 'wxyz', false, 'abcd')).toBe(true)
    expect(matcher.certifiedAsciiProfile).toBe(false)
  })
})

describe('bounded certificate ownership and measurements', () => {
  it('reuses actual native profile probes and certifies large alphabets by adjacency', () => {
    const matcher = new TextFilterMatcher('sv-SE'), text = 'é漢😀東à'.repeat(400)
    expect(matcher.certification.comparisons).toBe(0)
    matcher.match('contains', 'zzz', false, text)
    expect(matcher.certifiedAsciiProfile).toBe(true)
    const work = { used: 0 }
    matcher.match('contains', 'zzz', false, text, { work })
    expect(matcher.certification.comparisons).toBe(0)
    expect(matcher.certification.cacheHits).toBeGreaterThan(0)
    expect(work.used).toBeLessThan(100000)
    const huge = Array.from({ length: 80 }, (_, index) => String.fromCodePoint(0x4e00 + index)).join('').repeat(25)
    expect(matcher.match('contains', 'zzz', false, huge)).toBe(false)
    expect(matcher.certifiedAsciiProfile).toBe(true)
  })
  it('detects native pair contractions and expansion/canonical boundary differences', () => {
    const da = new TextFilterMatcher('da-DK'), en = new TextFilterMatcher('en-US')
    expect(da.match('contains', 'å', false, 'aa')).toBe(true)
    expect(da.certifiedAsciiProfile).toBe(false)
    expect(en.match('contains', 'fi', false, 'ﬁ')).toBe(true)
    expect(en.certifiedAsciiProfile).toBe(false)
    expect(en.match('contains', 'f', false, 'ﬁ')).toBe(false)
    expect(en.match('eq', 'ss', false, 'ß')).toBe(false)
    expect(en.match('eq', '?', true, 'é')).toBe(true)
    expect(en.match('eq', '?', true, 'e\u0301')).toBe(false)
    expect(en.match('eq', 'fi', false, 'ﬁ' + '\u200b'.repeat(400))).toBe(true)
  })
  it('reports cold and warm actual probe work separately from total cell work', () => {
    const rows: object[] = []
    for (const locale of ['en-US', 'de-DE', 'sv-SE', 'tr-TR', 'da-DK', 'es-ES-u-co-trad']) {
      for (const kind of ['prose', 'unicode', 'random-ascii', 'expansion'] as const) {
        let seed = 881, text = ''
        if (kind === 'prose') text = 'The quick brown fox jumps over the lazy dog. A filter examines each ordinary cell. '.repeat(25).slice(0, 2000)
        else {
          const pool = kind === 'unicode' ? ['é', 'à', '漢', '東', '😀', '😄'] : kind === 'expansion' ? ['ﬁ', 'ß', 'é', '\u0301'] : Array.from('abcdefghijklmnopqrstuvwxy 0123456789!')
          for (let index = 0; index < 2000; index++) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
        }
        const construction = performance.now(), matcher = new TextFilterMatcher(locale), constructorMs = performance.now() - construction
        for (const temperature of ['cold', 'warm'] as const) {
          const work = { used: 0 }, at = performance.now()
          let outcome = 'miss'
          try { matcher.match('contains', 'zzz', false, text, { work }) } catch (error) { if (!(error instanceof Error) || !/resource/i.test(error.message)) throw error; outcome = 'resource-limit' }
          rows.push({ locale, kind, temperature, constructorMs: Number(constructorMs.toFixed(3)), ms: Number((performance.now() - at).toFixed(3)), used: work.used, certified: matcher.certifiedAsciiProfile, ...matcher.certification, outcome })
        }
        if (kind === 'unicode') {
          const work = { used: 0 }, at = performance.now()
          expect(matcher.match('eq', '*é*é*zzz', true, text, { work })).toBe(false)
          rows.push({ locale, kind: 'unicode-wildcard', ms: Number((performance.now() - at).toFixed(3)), used: work.used, certified: matcher.certifiedAsciiProfile, ...matcher.certification })
        }
      }
    }
    console.info('S8 certificate measurements ' + JSON.stringify(rows))
    expect(rows.length).toBe(54)
  })
})


describe('step 2 observed-unit window rejection evidence', () => {
  it('clears previous probe observations for a wildcard without literals', () => {
    const matcher = new TextFilterMatcher('sv-SE')
    matcher.match('contains', 'zzz', false, 'é漢😀'.repeat(667))
    expect(matcher.certifiedAsciiProfile).toBe(true)
    expect(matcher.certification.chargedWork).toBeGreaterThan(0)
    expect(matcher.match('eq', '?', true, '😀')).toBe(true)
    expect(matcher.certifiedAsciiProfile).toBe(false)
    expect(matcher.certification).toEqual({ comparisons: 0, cacheHits: 0, chargedWork: 0 })
  })
  it('rejects the tentative non-mark-unit window after a native zero-unit alias counterexample', () => {
    const matcher = new TextFilterMatcher('en-US'), reference = new ReferenceTextFilterMatcher('en-US')
    // Attempted bound: NFKD non-mark units, actual empty-equal ignorables,
    // observed scalar expansion factor; plateau ends retain raw CP boundaries.
    // This is a rejected prototype, not shipped matching semantics.
    const units = (point: string) => matcher.collator.compare(point, '') === 0 ? 0
      : Array.from(point.normalize('NFKD')).filter(point => !/\p{M}/u.test(point)).length
    const windowMatch = (needle: string, text: string) => {
      const points = Array.from(text), alphabet = Array.from(new Set([...needle, ...text]))
      const factor = Math.max(1, ...alphabet.map(units)), needleUnits = Array.from(needle).reduce((sum, point) => sum + units(point), 0)
      const minimum = Math.ceil(needleUnits / factor), maximum = needleUnits * factor
      for (let start = 0; start <= points.length; start++) {
        let length = 0
        for (let end = start; end <= points.length; end++) {
          if (end > start) length += units(points[end - 1])
          if (length > maximum) break
          if (length >= minimum && matcher.collator.compare(needle, points.slice(start, end).join('')) === 0) return true
        }
      }
      return false
    }
    for (const [needle, text] of [['a', '\u0363'], ['e', '\u0364'], ['\u0363', 'a']]) {
      expect(matcher.collator.compare(needle, text)).toBe(0)
      expect(windowMatch(needle, text)).toBe(false)
      expect(reference.match('contains', needle, false, text)).toBe(true)
      expect(matcher.match('contains', needle, false, text)).toBe(true)
      expect(matcher.certifiedAsciiProfile).toBe(false)
    }
  })
  it('keeps ignored separators inside Danish contractions and raw accent slice boundaries exact', () => {
    const da = new TextFilterMatcher('da-DK'), en = new TextFilterMatcher('en-US')
    expect(da.match('contains', 'å', false, 'aa')).toBe(true)
    expect(da.match('contains', 'å', false, 'a\u200ba')).toBe(false)
    expect(da.match('contains', 'aa', false, 'aA')).toBe(false)
    expect(en.match('contains', 'e', false, 'é')).toBe(false)
    expect(en.match('contains', 'e', false, 'e\u0301')).toBe(true)
  })
})

describe('native-looking bound comparator ownership', () => {
  it('rejects a replaced prototype getter returning a bound JavaScript comparator', () => {
    const getter = Object.getOwnPropertyDescriptor(Intl.Collator.prototype, 'compare')!.get!
    const bound = new WeakMap<Intl.Collator, Intl.Collator['compare']>()
    try {
      vi.spyOn(Intl.Collator.prototype, 'compare', 'get').mockImplementation(function (this: Intl.Collator) {
        let compare = bound.get(this)
        if (!compare) {
          const native = getter.call(this) as Intl.Collator['compare']
          compare = function (a: string, b: string) { return native(a.replace(/abcd/g, 'wxyz'), b.replace(/abcd/g, 'wxyz')) }.bind(null)
          bound.set(this, compare)
        }
        return compare
      })
      const matcher = new TextFilterMatcher('en-US')
      expect(matcher.match('contains', 'wxyz', false, 'abcd')).toBe(true)
      expect(matcher.certifiedAsciiProfile).toBe(false)
    } finally { vi.restoreAllMocks() }
  })
})

describe('ignorable runs between context characters stay exact', () => {
  const ops: TextOp[] = ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains']
  const all = { maxWork: Number.MAX_SAFE_INTEGER }
  const agree = (locale: string, pattern: string, text: string) => {
    const fast = new TextFilterMatcher(locale), reference = new ReferenceTextFilterMatcher(locale)
    for (const op of ops) for (const wildcards of [false, true]) {
      const actual = fast.match(op, pattern, wildcards, text, all), expected = reference.match(op, pattern, wildcards, text, all)
      if (actual !== expected) throw new Error(JSON.stringify({ locale, op, pattern, text, wildcards, actual, expected }))
    }
  }
  it.each([
    ['en-US', 'l·', 'abl​​·cd'],
    ['zh', 'l·', 'xl​⁠·y'],
    ['de', 'l·', 'xl​﻿·y'],
    ['en', 'l·', 'abl​­·cd'],
    ['zh', 'el·', 'El​⁠·'],
    ['kk', 'el·', 'El​⁠·'],
    ['en-US', 'l​​·', 'abl·cd'],
    ['en-US', 'l·', 'abl​·cd'],
  ])('matches the reference for %s %j in %j', (locale, pattern, text) => agree(locale, pattern, text))
  it('pins the native oracle that the certificate must not contradict', () => {
    const reference = new ReferenceTextFilterMatcher('en-US'), fast = new TextFilterMatcher('en-US')
    const text = 'abl​​·cd'
    expect(reference.match('contains', 'l·', false, text)).toBe(false)
    expect(fast.match('contains', 'l·', false, text)).toBe(false)
    expect(fast.certifiedAsciiProfile).toBe(false)
    expect(fast.match('contains', 'cd', false, 'abcd')).toBe(true)
    expect(fast.certifiedAsciiProfile).toBe(true)
  })
  it('matches the reference over seeded ignorable-run fuzz in en, zh, de and kk', () => {
    const pool = ['​', '­', '﻿', '⁠', '‍', '·', 'l', 'e', 'a', 'L', 'E']
    const locales = ['en', 'zh', 'de', 'kk'].map(locale => [new TextFilterMatcher(locale), new ReferenceTextFilterMatcher(locale), locale] as const)
    let seed = 4241, checked = 0
    const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
    const cases = Number(process.env.S8_IGNORABLE_FUZZ_CASES ?? 3000)
    for (let index = 0; index < cases; index++) {
      let text = '', pattern = ''
      if (index % 2 === 0) {
        for (let i = 1 + random(9); i > 0; i--) text += pool[random(pool.length)]
        for (let i = 1 + random(4); i > 0; i--) pattern += pool[random(pool.length)]
      } else {
        // Structured: a known context with 0-3 ignorables between its characters.
        const contexts = ['l\u00b7', 'el\u00b7', 'L\u00b7', 'll', 'aa', 'l\u00b7l'], gap = () => { let run = ''; for (let i = random(4); i > 0; i--) run += pool[random(5)]; return run }
        for (const point of contexts[random(contexts.length)]) text += point + gap()
        text = 'ab'.slice(0, random(3)) + text + 'cd'.slice(0, random(3))
        for (const point of contexts[random(contexts.length)]) pattern += point + (random(4) === 0 ? gap() : '')
      }
      const [fast, reference, locale] = locales[index % locales.length]
      for (const op of ops) for (const wildcards of [false, true]) {
        const actual = fast.match(op, pattern, wildcards, text, all), expected = reference.match(op, pattern, wildcards, text, all)
        if (actual !== expected) throw new Error(JSON.stringify({ index, locale, op, pattern, text, wildcards, actual, expected }))
        checked++
      }
    }
    expect(checked).toBe(cases * ops.length * 2)
  })
})

describe('large alphabets are certified by observed adjacency', () => {
  const latin = Array.from('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,;-').slice(0, 50)
  const cjk = Array.from({ length: 150 }, (_, index) => String.fromCodePoint(0x4e00 + index * 7))
  const sample = (pool: readonly string[], length: number) => {
    let seed = 20261011, text = ''
    for (let i = 0; i < length; i++) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
    return text
  }
  it.each([['latin', latin], ['cjk', cjk]] as const)('keeps %s text within the default allowance at 2000 characters', (_name, pool) => {
    for (const locale of ['en-US', 'zh', 'sv-SE']) {
      for (const length of [100, 300, 1000, 2000]) {
        const matcher = new TextFilterMatcher(locale), work = { used: 0 }
        expect(matcher.match('contains', 'zzz', false, sample(pool, length), { work })).toBe(false)
        expect(matcher.certifiedAsciiProfile).toBe(true)
        expect(work.used).toBeLessThan(DEFAULT_FILTER_WORK / 5)
      }
    }
  })
  it('keeps certified large-alphabet answers equal to the reference', () => {
    for (const [pool, locale] of [[latin, 'en-US'], [cjk, 'zh'], [latin, 'da-DK']] as const) {
      const fast = new TextFilterMatcher(locale), reference = new ReferenceTextFilterMatcher(locale)
      const text = sample(pool, 120)
      for (const pattern of [text.slice(10, 14), text.slice(100), 'zzz', text.slice(3, 5) + '*' + text.slice(30, 33)]) {
        for (const op of ['eq', 'neq', 'begins', 'ends', 'contains', 'not-contains'] as TextOp[]) {
          expect(fast.match(op, pattern, true, text, { maxWork: Number.MAX_SAFE_INTEGER })).toBe(reference.match(op, pattern, true, text, { maxWork: Number.MAX_SAFE_INTEGER }))
        }
      }
    }
  })
  it('still fails explicitly or exactly when a large alphabet is inconclusive', () => {
    const da = new TextFilterMatcher('da-DK'), reference = new ReferenceTextFilterMatcher('da-DK')
    const text = sample(latin, 200) + 'aa'
    expect(da.match('contains', 'å', false, text)).toBe(reference.match('contains', 'å', false, text))
    expect(da.certifiedAsciiProfile).toBe(false)
  })
})
