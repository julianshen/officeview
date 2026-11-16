// @vitest-environment node
/**
 * Stage 5, Task 32: CF formula compile, admission, allow-list, model and parsing (Rulings S5-B, S5-SUMIF-SHAPE, S5-F, S5-G).
 * Nothing here evaluates a formula: compile is static on the AST, and parse stays behind evidence.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { getChildren, parseXml, type XmlNode } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { createCfMemory, CF_FORMULA_NODE_BYTES, CF_LIMITS } from '../src/xlsx/cf/limits'
import { withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { parseCfSheet, type CfStylesSource } from '../src/xlsx/cf/parse'
import {
  CF_FORMULA_FUNCTIONS, compileCfFormula, compileCfFormulaWith, countFormulaTokens, formulaRequiredCapabilities,
  type CfFormulaRuleContext,
} from '../src/xlsx/cf/formula'
import type { CfFormula, CfLease } from '../src/xlsx/cf/model'
import { cfDefinitionFormulas } from '../src/xlsx/cf/model'
import { FUNCTIONS, registerFormulaFunction } from '../src/xlsx/formula/functions'
import * as parserModule from '../src/xlsx/formula/parser'
import { CF_STAGE5_CAPABILITIES } from '../scripts/cf-refs/cases'
import { estimateCfPrepareBytes } from '../src/xlsx/cf/service'
import { canRuleMatchBlanks } from '../src/xlsx/cfLoader'

vi.mock('../src/xlsx/formula/parser', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/formula/parser')>()
  return { ...original, parseFormula: vi.fn(original.parseFormula) }
})
const parseSpy = vi.mocked(parserModule.parseFormula)
beforeEach(() => { parseSpy.mockClear() })

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const DXF_FILL = '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
const PARSE_SET = ['formula:dialect-base', 'formula:anchor-single', 'formula:anchor-mixed-abs'] as const
const BASE_RULE: CfFormulaRuleContext = { dialect: 'base', rangeCount: 1, sourceType: 'expression' }

const wrap = (inner: string): string => `<worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}">${inner}</worksheet>`
function stylesFor(): CfStylesSource {
  const root = parseXml(`<styleSheet xmlns="${NS}"><dxfs count="1">${DXF_FILL}</dxfs></styleSheet>`)
  return { colors: {}, dxfNodes: getChildren(getChildren(root, 'dxfs')[0] as XmlNode, 'dxf') }
}
const parseSheet = (inner: string, memory = createCfMemory()) => {
  const lease = parseCfSheet(parseXml(wrap(inner)), stylesFor(), memory)
  return { lease, rules: lease.value.rules, diagnostics: lease.value.diagnostics, memory }
}
const rule = (type: string, formula: string, sqref = 'B2:D9', attrs = '', extra = '') =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="${type}" priority="1" dxfId="0"${attrs}>${formula === '' ? '' : `<formula>${formula}</formula>`}${extra}</cfRule></conditionalFormatting>`
const compile = (text: string, memory = createCfMemory(), anchor: { row: number; col: number } | undefined = { row: 1, col: 1 }) => compileCfFormula(text, anchor, memory)
const formulaOf = (text: string): CfFormula => {
  const out = compile(text)
  if ('refused' in out) throw new Error(`refused ${text}: ${out.reason}`)
  return out.value
}

describe('compile admits the allow-list and refuses every out-of-scope node statically (S5-B)', () => {
  it.each([
    'MyName>0', '[1]Sheet1!A1>0', 'Table1[@Col]>0', 'INDIRECT("A1")', 'IF(FALSE,INDIRECT("A1"),1)', '{1,2}', '(A1,B1)', 'A1#',
    'Sheet1:Sheet3!A1', 'NOW()', 'VLOOKUP(A1,B:C,2)', 'SEARCH("a",A1)', 'A1 B1', '@A1', '#N/A', 'A1>', '',
    // Round-2 #1 and #4: a multi-cell reference only in a range slot of an admitted function, or as the whole result.
    'LEN(CONCAT(A1:A64))', 'CONCAT(A1:A3)', 'A2:A4>0', 'IF(TRUE,A2:A4,0)', 'IFERROR(A2:A4,0)', 'NOT(A2:A4)', 'IFNA(A2:A4,0)', 'ABS(A2:A4)', '-A2:A4',
    // A range whose endpoints disagree on pinning is one atom no field can hold.
    'SUM($A$2:A9)>0', 'A1:B2:C3', "'[2]Sheet'!A1", 'ROW(A1:A3)=1', 'AND(A1:A3)',
  ])('refuses %s', text => {
    const out = compile(text)
    expect(out).toMatchObject({ refused: expect.stringMatching(/^cf-(unsupported-rule|resource-limit)$/) })
    if (text !== '') expect(out).toMatchObject({ refused: 'cf-unsupported-rule' })
  })
  it('compiles the admitted shapes', () => {
    for (const text of ['MOD(A1,2)=1', 'COUNTIF($A$1:$A$50,A1)>1', 'CONCAT(A1,"x")', 'AND(A1>0,B1<5)', 'A2:A4', '$A:$A', 'IF(A1,,2)', 'A1&"x"', '=A1>0', 'SUM(A1:A3,B1)>1', '_xlfn.IFNA(A1,1)', 'TODAY()=A1', 'ROW()=2']) {
      const out = compile(text)
      expect('refused' in out ? out : undefined, text).toBeUndefined()
      if (!('refused' in out)) out.release()
    }
  })
  it('refuses a function registered after load, and an omitted argument where a range is read', () => {
    registerFormulaFunction('ZZZ', () => 1)
    expect(compile('ZZZ(A1)')).toMatchObject({ refused: 'cf-unsupported-rule' })
    expect(CF_FORMULA_FUNCTIONS.has('ZZZ')).toBe(false)
    delete FUNCTIONS['ZZZ']
    expect(compile('SUM(A1,,B1)')).toMatchObject({ refused: 'cf-unsupported-rule' })
    expect(compile('SUMIF(A1:A2,"",)')).toMatchObject({ refused: 'cf-unsupported-rule' })
  })
  it('the allow-list is frozen, every entry has its engine handler and declares geometry for every argument', () => {
    expect([...CF_FORMULA_FUNCTIONS.keys()].sort()).toEqual([
      'ABS', 'AND', 'AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COLUMN', 'CONCAT', 'COUNT', 'COUNTA', 'COUNTIF', 'COUNTIFS', 'DATE', 'DAY', 'IF', 'IFERROR', 'IFNA', 'INT',
      'ISBLANK', 'ISERROR', 'ISLOGICAL', 'ISNA', 'ISNUMBER', 'ISTEXT', 'LEFT', 'LEN', 'LOWER', 'MAX', 'MID', 'MIN', 'MOD', 'MONTH', 'NOT', 'OR', 'RIGHT', 'ROUND', 'ROW', 'SUM',
      'SUMIF', 'SUMIFS', 'TODAY', 'TRIM', 'UPPER', 'WEEKDAY', 'YEAR',
    ].sort())
    for (const [name, entry] of CF_FORMULA_FUNCTIONS) {
      expect(typeof FUNCTIONS[name], name).toBe('function')
      expect(entry.geometry.length, `${name} geometry per argument`).toBe(entry.args.length)
      for (const g of entry.geometry) if (g !== 'literal') expect(g.resizedTo).toBeLessThan(entry.args.length)
    }
    expect(() => { (CF_FORMULA_FUNCTIONS as Map<string, unknown>).set('ZZZ', {}) }).not.toThrow()
    // Withheld by default (OD-13/OD-14).
    for (const name of ['NOW', 'INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'SEQUENCE', 'FILTER', 'SORT', 'UNIQUE', 'TRANSPOSE', 'TEXT', 'DATEVALUE', 'TIMEVALUE', 'SEARCH', 'FIND', 'EXACT', 'INDIRECT', 'OFFSET', 'SUMPRODUCT', 'RAND', 'CELL']) {
      expect(CF_FORMULA_FUNCTIONS.has(name), name).toBe(false)
    }
  })
  it('admits no handler other than the one bound when the module loaded', async () => {
    const original = FUNCTIONS['LEN']
    FUNCTIONS['LEN'] = () => 0
    try { expect(compile('LEN(A1)>1')).toMatchObject({ refused: 'cf-unsupported-rule' }) } finally { FUNCTIONS['LEN'] = original! }
    expect('refused' in compile('LEN(A1)>1')).toBe(false)
  })
  it('refuses LEN(CONCAT(A1:A64)) before any evaluation: no handler runs', async () => {
    vi.resetModules()
    const spies = Object.keys(FUNCTIONS).map(name => vi.spyOn(FUNCTIONS, name))
    try {
      const fresh = await import('../src/xlsx/cf/formula')
      const out = fresh.compileCfFormula('LEN(CONCAT(A1:A64))', { row: 0, col: 0 }, createCfMemory())
      expect(out).toMatchObject({ refused: 'cf-unsupported-rule', reason: expect.stringMatching(/multi-cell/) })
      expect(spies.every(spy => spy.mock.calls.length === 0)).toBe(true)
    } finally { for (const spy of spies) spy.mockRestore() }
  })
})

describe('functions with implicit expansion compile only at literal geometry (S5-SUMIF-SHAPE)', () => {
  it.each(['SUMIF(A1:A2,"",B1)', 'AVERAGEIF(A1:A2,"",B1)', 'SUMIF(A:A,"",B1)', 'SUMIF(A1:A2,"",B1:B3)', 'SUMIF(A1:B2,"",B1:B3)', 'SUMIF(A1:A2,"",5)', 'SUMIF(A:A,"",B1:B5)'])('refuses %s', text => {
    expect(compile(text)).toMatchObject({ refused: 'cf-unsupported-rule', reason: expect.stringMatching(/sum range|resized/) })
  })
  it('admits an identical shape or an omitted sum range, and the footprint holds the effective geometry', () => {
    for (const text of ['SUMIF(A1:A2,"",B1:B2)', 'AVERAGEIF(A1:A2,"",B1:B2)', 'SUMIF(A:A,"",B:B)', 'SUMIF(A1:A2,"")', 'SUMIF(A1,"x",B1)']) expect('refused' in compile(text), text).toBe(false)
    expect(formulaOf('SUMIF(A1:A2,"",B1:B2)>0').atoms).toEqual([
      { firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 0, absRow: false, absCol: false },
      { firstRow: 0, lastRow: 1, firstCol: 1, lastCol: 1, absRow: false, absCol: false },
    ])
    expect(formulaOf('SUMIF(A:A,"",B:B)>0').atoms.map(a => [a.axis, a.firstCol, a.lastRow])).toEqual([['col', 0, 1048575], ['col', 1, 1048575]])
  })
  it('SUMIFS, AVERAGEIFS and COUNTIFS read every range as written, whatever the shapes', () => {
    for (const text of ['SUMIFS(B1:B3,A1:A3,"x")', 'AVERAGEIFS(B1:B3,A1:A3,"x",C1:C3,">1")', 'COUNTIFS(A1:A3,"x",B1:B2,1)', 'COUNTIF(A1:A3,"x")']) expect('refused' in compile(text), text).toBe(false)
    expect(formulaOf('SUMIFS(B1:B3,A1:A3,"x")>0').atoms).toHaveLength(2)
  })
})

describe('compiled formula records atoms, families, invariance and clock use', () => {
  it('a pinned aggregate is invariant with its atoms and family', () => {
    const f = formulaOf('$A$1>AVERAGE($A$2:$A$9)')
    expect(f.invariant).toBe(true)
    expect(f.atoms).toEqual([
      { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0, absRow: true, absCol: true },
      { firstRow: 1, lastRow: 8, firstCol: 0, lastCol: 0, absRow: true, absCol: true },
    ])
    expect(f.families).toEqual(['aggregate'])
    expect(f.pinned).toBe(true)
    expect(f.usesClock).toBe(false)
    expect(f.dependencySheets).toEqual([])
    expect(f.anchor).toEqual({ row: 1, col: 1 })
  })
  it('relative references and ROW() are not invariant; TODAY uses the clock; another sheet is a dependency', () => {
    expect(formulaOf('A2>0').invariant).toBe(false)
    expect(formulaOf('ROW()=2').invariant).toBe(false)
    expect(formulaOf('COLUMN()=2').invariant).toBe(false)
    expect(formulaOf('ROW($A$1)=1').invariant).toBe(true)
    expect(formulaOf('$A2>0').invariant).toBe(false)
    const today = formulaOf('A2=TODAY()')
    expect(today.usesClock).toBe(true)
    expect(today.families).toEqual(['today'])
    expect(formulaOf('Sheet2!$B$1<A2').dependencySheets).toEqual(['Sheet2'])
    expect(formulaOf("'My Sheet'!A1+MY!A1+my!B1+'My Sheet'!A2").dependencySheets).toEqual(['My Sheet', 'MY'])
  })
  it('whole-column and whole-row atoms carry their axis and pin the other axis', () => {
    expect(formulaOf('COUNTIF($A:$A,A2)>1').atoms[0]).toEqual({ firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0, absRow: true, absCol: true, axis: 'col' })
    expect(formulaOf('SUM(1:2)>0').atoms[0]).toEqual({ firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 16383, absRow: false, absCol: true, axis: 'row' })
  })
  it('records the static root kinds', () => {
    expect(formulaOf('A2>0').truth).toEqual(['boolean'])
    expect(formulaOf('MOD(A2,2)').truth).toEqual(['number'])
    expect(formulaOf('A2&"x"').truth).toEqual(['text'])
    expect(formulaOf('A2:A4').truth).toEqual(['range'])
    expect(formulaOf('A2:A4').rootArea).toBe(true)
    expect(formulaOf('IF(A1,1,"x")').truth).toEqual(['number', 'text'])
    expect(formulaOf('$A2').truth).toEqual(['blank', 'boolean', 'error', 'number', 'text'])
  })
})

describe('compile caps refuse before construction', () => {
  const sum = (n: number, ref: (i: number) => string) => Array.from({ length: n }, (_, i) => ref(i)).join('+')
  it('4,097 nodes refuse and 4,096 are admitted', () => {
    const body = (numbers: number) => `${Array.from({ length: numbers }, () => '1').join('+')}`
    // 2,048 numbers + 2,047 sums + one unary minus = 4,096 nodes (and 4,096 tokens): the limit itself.
    const atLimit = `-${body(2048)}`
    expect(countFormulaTokens(atLimit)).toBe(4096)
    expect(compile(atLimit)).toMatchObject({ value: { nodes: 4096 } })
    expect(compile(`-${body(2049)}`)).toMatchObject({ refused: 'cf-resource-limit' })
    expect(compile(`--${body(2048)}`)).toMatchObject({ refused: 'cf-resource-limit' })
  })
  it('65 atoms refuse and 64 are admitted', () => {
    expect(compile(sum(64, i => `A${i + 1}`))).toMatchObject({ value: { atoms: expect.arrayContaining([]) } })
    expect(compile(sum(65, i => `A${i + 1}`))).toMatchObject({ refused: 'cf-resource-limit' })
  })
  it('17 referenced sheets refuse and 16 are admitted', () => {
    expect(compile(sum(16, i => `S${i}!A1`))).toMatchObject({ value: { dependencySheets: expect.any(Array) } })
    expect(compile(sum(17, i => `S${i}!A1`))).toMatchObject({ refused: 'cf-resource-limit' })
  })
  it('refuses text over the text cap and non-finite literals', () => {
    expect(compile(`"${'x'.repeat(CF_LIMITS.textUnits)}"`)).toMatchObject({ refused: 'cf-resource-limit' })
    expect(compile('1E999')).toMatchObject({ refused: 'cf-unsupported-rule' })
  })
})

describe('compile reserves before parsing and releases with the rule (round-1 #3)', () => {
  const text = '$B2>$C2'
  const worst = countFormulaTokens(text) * CF_FORMULA_NODE_BYTES
  it('one byte short: parseFormula is never called and no AST exists', () => {
    const memory = createCfMemory({ limitBytes: worst - 1 })
    expect(compileCfFormula(text, { row: 1, col: 1 }, memory)).toMatchObject({ refused: 'cf-resource-limit' })
    expect(parseSpy).not.toHaveBeenCalled()
    expect(memory.currentBytes).toBe(0)
  })
  it('exactly enough: it compiles and the parse account holds the reservation', () => {
    const memory = createCfMemory({ limitBytes: worst })
    const out = compileCfFormula(text, { row: 1, col: 1 }, memory) as CfLease<CfFormula>
    expect(parseSpy).toHaveBeenCalledTimes(1)
    expect(out.bytes).toBe(worst)
    expect(memory.currentBytes).toBe(worst)
    out.release()
    expect(memory.currentBytes).toBe(0)
  })
  it('a refusal after the reservation (a parse failure, a function outside the list) returns the ledger to baseline', () => {
    const memory = createCfMemory()
    expect(compileCfFormula('INDIRECT("A1")', undefined, memory)).toMatchObject({ refused: 'cf-unsupported-rule' })
    expect(compileCfFormula('A1>', undefined, memory)).toMatchObject({ refused: 'cf-unsupported-rule' })
    expect(parseSpy).toHaveBeenCalledTimes(2)
    expect(memory.currentBytes).toBe(0)
  })
  it('the mutant with the accounting removed parses under insufficient memory, so the first assertion catches it', () => {
    parseSpy.mockClear()
    const out = compileCfFormulaWith(text, undefined, () => true)
    expect('refused' in out).toBe(false)
    expect(parseSpy).toHaveBeenCalled()
  })
  it('a cancelled streamed import and sheet disposal return the parse ledger to baseline', async () => {
    const memory = createCfMemory()
    const xml = wrap(`<conditionalFormatting sqref="B2:D9"><cfRule type="expression" priority="1" dxfId="0"><formula>${text}</formula></cfRule></conditionalFormatting>`)
    await withExactCfEvidence(PARSE_SET, async () => {
      const h = createCfStreamEvents(stylesFor(), memory, {})
      const bytes = new TextEncoder().encode(xml)
      const half = Math.floor(bytes.length / 2)
      await tokenizeBytes((async function* () { yield bytes.subarray(0, half); yield bytes.subarray(half) })(), h.events)
      // Compile ran inside finish; dispose a finished-then-released lease, and a partially captured one.
      const lease = h.finish()
      expect(cfDefinitionFormulas(lease.value.rules[0]!.definition)).toHaveLength(1)
      expect(memory.currentBytes).toBeGreaterThan(0)
      lease.release()
      expect(memory.currentBytes).toBe(0)
      const cancelled = createCfStreamEvents(stylesFor(), memory, {})
      cancelled.events.open('worksheet', ['xmlns'], [NS], 1, false)
      cancelled.events.open('conditionalFormatting', ['sqref'], ['A1'], 1, false)
      cancelled.dispose()
      expect(memory.currentBytes).toBe(0)
      // Sheet disposal of a resident parse.
      const resident = parseSheet(rule('expression', text), memory)
      expect(resident.rules[0]!.definition.kind).toBe('expression')
      expect(memory.currentBytes).toBeGreaterThan(0)
      resident.lease.release()
      expect(memory.currentBytes).toBe(0)
    })
  })
  it('a formula the sheet cannot hold refuses its rule, not the sheet', () => {
    // Each holds about 0.5 MB, and its reservation (4,096 tokens x node bytes, about 1.8 MB) must also fit beside the others
    // under the 4 MiB sheet metadata cap: the later rules refuse by themselves.
    const big = `-${Array.from({ length: 2048 }, () => '1').join('+')}`
    const rules = Array.from({ length: 8 }, (_, i) => `<conditionalFormatting sqref="B${i + 2}"><cfRule type="expression" priority="${i + 1}" dxfId="0"><formula>${big}</formula></cfRule></conditionalFormatting>`).join('')
    withExactCfEvidence(PARSE_SET, () => {
      const { rules: parsed, lease } = parseSheet(rules)
      expect(lease.value.unresolvedSheet).toBeUndefined()
      const kinds = parsed.map(r => r.definition.kind)
      const firstRefused = kinds.indexOf('unresolved')
      expect(firstRefused).toBeGreaterThan(0)
      expect(kinds.slice(0, firstRefused).every(k => k === 'expression')).toBe(true)
      expect(parsed.slice(firstRefused).every(r => r.definition.kind === 'unresolved' && r.definition.reason === 'cf-resource-limit')).toBe(true)
      lease.release()
    })
  })
})

describe('expression and operand rules parse to formula definitions behind evidence', () => {
  it('an expression rule resolves under the complete parse set, with its anchor', () => {
    withExactCfEvidence(PARSE_SET, () => {
      const { rules, lease } = parseSheet(rule('expression', '$B2&gt;$C2'))
      expect(rules[0]!.definition).toMatchObject({ kind: 'expression', sourceType: 'expression', dialect: 'base', formulas: ['$B2>$C2'] })
      const def = rules[0]!.definition as { formula: CfFormula }
      expect(def.formula.anchor).toEqual({ row: 1, col: 1 })
      expect(rules[0]!.writes.has('fill.pattern') || rules[0]!.writes.size > 0).toBe(true)
      expect(rules[0]!.dxf).toBeDefined()
      lease.release()
    })
  })
  it('without evidence, or without any one parse capability, it is unresolved with the DXF writes and one diagnostic', () => {
    const resolvedWrites = withExactCfEvidence(PARSE_SET, () => { const r = parseSheet(rule('expression', '$B2&gt;$C2')); const w = [...r.rules[0]!.writes].sort(); r.lease.release(); return w })
    for (const evidence of [[], ['formula:dialect-base'], ['formula:dialect-base', 'formula:anchor-single'], ['formula:anchor-single', 'formula:anchor-mixed-abs'], ['formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs']]) {
      parseSpy.mockClear()
      withExactCfEvidence(evidence, () => {
        const { rules, diagnostics, memory, lease } = parseSheet(rule('expression', '$B2&gt;$C2'))
        expect(rules[0]!.definition).toMatchObject({ kind: 'unresolved', type: 'expression', reason: 'cf-unsupported-rule' })
        expect([...rules[0]!.writes].sort()).toEqual(resolvedWrites)
        expect(diagnostics.filter(d => d.feature === 'cf-unsupported-rule')).toHaveLength(1)
        // The formula-independent gate fails first: nothing is parsed or reserved without evidence.
        if (evidence.length < 2) expect(parseSpy).not.toHaveBeenCalled()
        lease.release(); expect(memory.currentBytes).toBe(0)
      })
    }
  })
  it('production evidence (none for formula:*) leaves every formula rule unresolved', () => {
    const { rules, lease } = parseSheet(rule('expression', 'A1&gt;0') + rule('cellIs', '$F$1', 'B2', ' operator="greaterThan"'))
    expect(rules.map(r => r.definition.kind)).toEqual(['unresolved', 'unresolved'])
    expect(parseSpy).not.toHaveBeenCalled()
    lease.release()
  })
  it('a pinned reference needs anchor-mixed-abs; a multi-range rule needs anchor-multi-range; x14 needs its dialect; cross-sheet needs its capability', () => {
    const kinds = (evidence: readonly string[], xml: string) => withExactCfEvidence(evidence, () => { const r = parseSheet(xml); const k = r.rules.map(x => x.definition.kind); r.lease.release(); return k })
    expect(kinds(['formula:dialect-base', 'formula:anchor-single'], rule('expression', '$B2&gt;1'))).toEqual(['unresolved'])
    expect(kinds(['formula:dialect-base', 'formula:anchor-single'], rule('expression', 'B2&gt;1'))).toEqual(['expression'])
    expect(kinds(PARSE_SET, rule('expression', 'B2&gt;1', 'B2:B4 D2:D4'))).toEqual(['unresolved'])
    expect(kinds([...PARSE_SET, 'formula:anchor-multi-range'], rule('expression', 'B2&gt;1', 'B2:B4 D2:D4'))).toEqual(['expression'])
    expect(kinds(PARSE_SET, rule('expression', 'Sheet2!$B$1&lt;B2'))).toEqual(['unresolved'])
    const crossed = withExactCfEvidence([...PARSE_SET, 'formula:cross-sheet-reference'], () => { const r = parseSheet(rule('expression', 'Sheet2!$B$1&lt;B2')); const d = r.rules[0]!.definition; r.lease.release(); return d })
    expect(crossed).toMatchObject({ kind: 'expression', dependencySheets: ['Sheet2'] })
    expect(kinds([...PARSE_SET, 'formula:cross-sheet-reference'], rule('expression', "'My Sheet'!$B$1&lt;B2"))).toEqual(['unresolved'])
    expect(kinds([...PARSE_SET, 'formula:cross-sheet-reference', 'formula:cross-sheet-quoted-name'], rule('expression', "'My Sheet'!$B$1&lt;B2"))).toEqual(['expression'])
  })
  it('an out-of-scope or over-cap formula stays unresolved under complete evidence, naming its reason', () => {
    withExactCfEvidence(PARSE_SET, () => {
      const names = (xml: string) => { const r = parseSheet(xml); const d = r.rules[0]!.definition; r.lease.release(); return d }
      expect(names(rule('expression', 'INDIRECT("A1")'))).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      expect(names(rule('expression', `${Array.from({ length: 65 }, (_, i) => `A${i + 1}`).join('+')}`))).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
      expect(names(rule('expression', ''))).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      expect(names(`<conditionalFormatting sqref="B2"><cfRule type="expression" priority="1" dxfId="0"><formula>A1</formula><formula>B1</formula></cfRule></conditionalFormatting>`)).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    })
  })
  it('a standalone x14 expression compiles under the x14 dialect', () => {
    const inner = `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="1" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>$A1="x"</xm:f></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    withExactCfEvidence(PARSE_SET, () => { const r = parseSheet(inner); expect(r.rules[0]!.definition.kind).toBe('unresolved'); r.lease.release() })
    withExactCfEvidence(['formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs'], () => {
      const r = parseSheet(inner)
      expect(r.rules[0]!.definition).toMatchObject({ kind: 'expression', dialect: 'x14', sourceType: 'expression' })
      r.lease.release()
    })
  })
  it('cellIs formula operands: a reference, an expression, a mixed literal; literals stay literals', () => {
    withExactCfEvidence(PARSE_SET, () => {
      const r = parseSheet(
        rule('cellIs', '$F$1', 'B2', ' operator="greaterThan"').replace('priority="1"', 'priority="1"')
        + `<conditionalFormatting sqref="B3"><cfRule type="cellIs" priority="2" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>`
        + `<conditionalFormatting sqref="B4"><cfRule type="cellIs" priority="3" operator="between"><formula>1</formula><formula>$F$2*2</formula></cfRule></conditionalFormatting>`
        + `<conditionalFormatting sqref="B5"><cfRule type="cellIs" priority="4" operator="equal"><formula>"a"&amp;"b"</formula></cfRule></conditionalFormatting>`
        + `<conditionalFormatting sqref="B6"><cfRule type="cellIs" priority="5" operator="containsText"><formula>$F$1</formula></cfRule></conditionalFormatting>`,
      )
      const [a, b, c, d, e] = r.rules.map(x => x.definition)
      expect(a).toMatchObject({ kind: 'cellIs', operator: 'greaterThan', operands: [{ formula: { nodes: 1, atoms: [expect.objectContaining({ firstRow: 0, firstCol: 5, absRow: true })] } }] })
      expect((a as unknown as { operands: object[] }).operands[0]).not.toHaveProperty('kind')
      expect(b).toMatchObject({ kind: 'cellIs', operands: [{ kind: 'number', value: 100 }] })
      expect(c).toMatchObject({ kind: 'cellIs', operator: 'between', operands: [{ kind: 'number', value: 1 }, { formula: { nodes: 3 } }] })
      expect(d).toMatchObject({ kind: 'cellIs', operands: [{ formula: expect.anything() }] })
      expect(e).toMatchObject({ kind: 'unresolved', type: 'cellIs', reason: 'cf-unsupported-rule' })
      expect(cfDefinitionFormulas(a!)).toHaveLength(1)
      r.lease.release()
    })
    // Without evidence the literal path is untouched and the formula operand is unresolved, as before.
    const none = parseSheet(rule('cellIs', '$F$1', 'B2', ' operator="greaterThan"') + `<conditionalFormatting sqref="B3"><cfRule type="cellIs" priority="2" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>`)
    expect(none.rules.map(x => x.definition.kind)).toEqual(['unresolved', 'cellIs'])
    none.lease.release()
  })
  it('a divergent blank, error or text predicate compiles to an expression with its own type; the canonical forms keep their predicates', () => {
    withExactCfEvidence(PARSE_SET, () => {
      const one = (type: string, formula: string, attrs = '', sqref = 'D2:D9') => { const r = parseSheet(rule(type, formula, sqref, attrs)); const d = r.rules[0]!.definition; r.lease.release(); return d }
      expect(one('containsBlanks', 'ISBLANK(D2)')).toMatchObject({ kind: 'expression', sourceType: 'containsBlanks' })
      expect(one('notContainsErrors', 'NOT(ISNA(E2))', '', 'E2:E9')).toMatchObject({ kind: 'expression', sourceType: 'notContainsErrors' })
      expect(one('beginsWith', 'LEFT(D2,3)="abcd"', ' text="ab"')).toMatchObject({ kind: 'expression', sourceType: 'beginsWith' })
      expect(one('containsBlanks', 'LEN(TRIM(Z99))=0')).toMatchObject({ kind: 'expression', sourceType: 'containsBlanks' })
      // Canonical: unchanged Stage 1 predicate and text definitions.
      expect(one('containsBlanks', 'LEN(TRIM(D2))=0')).toMatchObject({ kind: 'predicate', type: 'containsBlanks' })
      expect(one('containsText', 'NOT(ISERROR(SEARCH(&quot;ab&quot;,D2)))', ' text="ab"')).toMatchObject({ kind: 'text', type: 'containsText' })
      // An absent formula is the Stage 1 never-matching rule (unresolved here), a duplicate/unique formula stays re-deferred.
      expect(one('containsBlanks', '')).toMatchObject({ kind: 'unresolved' })
      expect(one('duplicateValues', 'A1')).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      // A divergent SEARCH is refused by the allow-list, not evaluated.
      expect(one('containsText', 'SEARCH(&quot;zz&quot;,D2)', ' text="ab"')).toMatchObject({ kind: 'unresolved' })
    })
  })
})

describe('formulaRequiredCapabilities splits parse and evaluation gates (round-1 #12)', () => {
  const caps = (text: string, use: 'expression' | 'operand' | 'cfvo' | 'activation' = 'expression', rule: Partial<CfFormulaRuleContext> = {}) =>
    formulaRequiredCapabilities(formulaOf(text), use, { ...BASE_RULE, ...rule })
  it('=$B2>$C2 gates parse on dialect, anchor and mixed pinning; evaluation on the boolean truth reading', () => {
    const c = caps('$B2>$C2')
    expect([...c.parse].sort()).toEqual([...PARSE_SET].sort())
    expect(c.evaluation).toEqual(new Set(['formula:truth-boolean']))
  })
  it('a function adds its family, a multi-range rule its anchor, x14 its dialect, another sheet its capability', () => {
    expect(caps('MOD(A2,2)=1').evaluation).toEqual(new Set(['formula:fn-math', 'formula:truth-boolean']))
    expect(caps('A2>AVERAGE($A$2:$A$9)').evaluation).toEqual(new Set(['formula:fn-aggregate', 'formula:truth-boolean']))
    expect(caps('B2>1', 'expression', { rangeCount: 2 }).parse.has('formula:anchor-multi-range')).toBe(true)
    expect(caps('B2>1', 'expression', { rangeCount: 1 }).parse.has('formula:anchor-multi-range')).toBe(false)
    const x14 = caps('$B2>$C2', 'expression', { dialect: 'x14' }).parse
    expect(x14.has('formula:dialect-x14')).toBe(true); expect(x14.has('formula:dialect-base')).toBe(false)
    expect(caps('Sheet2!$B$1<A2').parse.has('formula:cross-sheet-reference')).toBe(true)
    expect(caps("'My Sheet'!A1<A2").parse.has('formula:cross-sheet-quoted-name')).toBe(true)
    expect(caps('Sheet2!A1<A2').parse.has('formula:cross-sheet-quoted-name')).toBe(false)
    expect(caps('A2=TODAY()').evaluation.has('formula:fn-today')).toBe(true)
    expect(caps('A2:A4').evaluation).toEqual(new Set(['formula:truth-range']))
    expect(caps('B2>1', 'cfvo', { rangeCount: 2 }).evaluation.has('formula:cfvo-scope-multi-range')).toBe(true)
    expect(caps('B2>1', 'cfvo', { rangeCount: 1 }).evaluation.has('formula:cfvo-scope-multi-range')).toBe(false)
  })
  it('a divergent predicate adds formula:predicate-divergent; operands and activation have their own gates', () => {
    expect(caps('ISBLANK(D2)', 'expression', { sourceType: 'containsBlanks' }).evaluation.has('formula:predicate-divergent')).toBe(true)
    expect(caps('ISBLANK(D2)').evaluation.has('formula:predicate-divergent')).toBe(false)
    expect(caps('$F$1', 'operand', { sourceType: 'cellIs' }).evaluation).toEqual(new Set(['formula:cellis-operand-reference']))
    expect(caps('F2*2', 'operand', { sourceType: 'cellIs' }).evaluation).toEqual(new Set(['formula:cellis-operand-expression', 'formula:cellis-operand-relative']))
    expect(caps('Sheet2!$A$1', 'operand', { sourceType: 'cellIs' }).evaluation.has('formula:cellis-operand-cross-sheet')).toBe(true)
    expect(caps('$F$1', 'cfvo', { sourceType: 'colorScale' }).evaluation).toEqual(new Set(['formula:cfvo-reference']))
    expect(caps('MIN($A$2:$A$9)+1', 'cfvo', { sourceType: 'colorScale' }).evaluation).toEqual(new Set(['formula:cfvo-expression', 'formula:fn-aggregate']))
    expect(caps('$F$1', 'activation', { sourceType: 'dataBar' }).evaluation.has('formula:active-present-bar')).toBe(true)
  })
  it('neither set is ever empty, and every id it can emit is an oracle capability', () => {
    const known = new Set(CF_STAGE5_CAPABILITIES.map(c => c.capabilityId))
    const texts = ['A2>0', '$A$1', 'A2:A4', 'A2&"x"', 'MOD(A2,2)=1', 'AND(A1>0,B1<5)', 'ISBLANK(A1)', 'LEN(A1)>3', 'SUM($A$1:$A$9)>1', 'COUNTIF($A:$A,A2)>1', 'YEAR(A2)=2026', 'ROW()=COLUMN()',
      'A2=TODAY()', "'My Sheet'!A1>0", 'Sheet2!$A$1>0', '5', '"x"', 'TRUE', 'IF(A1,1,"x")', 'SUMIF(A1:A2,"",B1:B2)>0']
    const uses = ['expression', 'operand', 'cfvo', 'activation'] as const
    const sources = ['expression', 'cellIs', 'containsBlanks', 'colorScale', 'dataBar', 'iconSet'] as const
    for (const text of texts) for (const use of uses) for (const sourceType of sources) for (const dialect of ['base', 'x14'] as const) for (const rangeCount of [1, 2]) {
      const c = formulaRequiredCapabilities(formulaOf(text), use, { dialect, rangeCount, sourceType })
      expect(c.parse.size, `${text}/${use}`).toBeGreaterThan(0)
      expect(c.evaluation.size, `${text}/${use}`).toBeGreaterThan(0)
      for (const id of [...c.parse, ...c.evaluation]) expect(known.has(id), `${id} (${text}, ${use})`).toBe(true)
    }
  })
})

describe('formula definitions are priced and statistical-aware', () => {
  it('estimateCfPrepareBytes adds the node-priced term once per formula rule, outside the per-cell product', () => {
    const plain = parseSheet(rule('cellIs', '100', 'B2:D9', ' operator="greaterThan"'))
    const nodes = 21
    const text = Array.from({ length: nodes }, (_, i) => `A${i + 1}`).join('+')
    const priced = withExactCfEvidence(PARSE_SET, () => parseSheet(rule('expression', text, 'B2:D9')))
    const rects = [{ firstRow: 1, lastRow: 8, firstCol: 1, lastCol: 3 }]
    const base = estimateCfPrepareBytes(plain.lease.value, rects)!
    const withFormula = estimateCfPrepareBytes(priced.lease.value, rects)!
    const formula = (priced.rules[0]!.definition as { formula: CfFormula }).formula
    expect(withFormula - base).toBe(CF_FORMULA_NODE_BYTES * formula.nodes)
    // Per-cell terms do not multiply it: the same rule over twice the cells adds the same formula term.
    const wider = [{ firstRow: 1, lastRow: 8, firstCol: 1, lastCol: 6 }]
    expect(estimateCfPrepareBytes(priced.lease.value, wider)! - estimateCfPrepareBytes(plain.lease.value, wider)!).toBe(CF_FORMULA_NODE_BYTES * formula.nodes)
    plain.lease.release(); priced.lease.release()
    const source = readFileSync(new URL('../src/xlsx/cf/service.ts', import.meta.url), 'utf8')
    const arm = source.split('\n').filter(line => line.includes('cfDefinitionFormulas(rule.definition)'))
    expect(arm).toHaveLength(2)
    expect(arm[0]).toMatch(/CF_FORMULA_NODE_BYTES \* formula\.nodes/)
    expect(arm[0]).not.toMatch(/\b\d{2,}\b/)
  })
  it('statistical() is unchanged for expressions: they read the snapshot, not range statistics', () => {
    const source = readFileSync(new URL('../src/xlsx/cf/service.ts', import.meta.url), 'utf8')
    const fn = source.slice(source.indexOf('function statistical'), source.indexOf('/** A caller-owned'))
    expect(fn).not.toMatch(/expression|formula/)
  })
})

describe('blank synthesis stays on for formula rules (S5-P)', () => {
  it('canRuleMatchBlanks is true for resolved expressions, formula operands and divergent predicates, and for the unresolved expression', () => {
    withExactCfEvidence(PARSE_SET, () => {
      const r = parseSheet(
        rule('expression', 'ISBLANK(D2)')
        + `<conditionalFormatting sqref="B3"><cfRule type="cellIs" priority="2" operator="greaterThan"><formula>$F$1</formula></cfRule></conditionalFormatting>`
        + `<conditionalFormatting sqref="B4"><cfRule type="cellIs" priority="3" operator="greaterThan"><formula>7</formula></cfRule></conditionalFormatting>`
        + `<conditionalFormatting sqref="D2:D9"><cfRule type="containsBlanks" priority="4"><formula>ISBLANK(D2)</formula></cfRule></conditionalFormatting>`,
      )
      expect(r.rules.map(x => canRuleMatchBlanks(x))).toEqual([true, true, false, true])
      r.lease.release()
    })
    const unresolved = parseSheet(rule('expression', 'ISBLANK(D2)'))
    expect(unresolved.rules[0]!.definition).toMatchObject({ kind: 'unresolved', type: 'expression' })
    expect(canRuleMatchBlanks(unresolved.rules[0]!)).toBe(true)
    unresolved.lease.release()
  })
})

describe('the formula arm never reaches the literal predicates', () => {
  it('evaluatePredicate refuses an expression and a formula operand by name', async () => {
    const { evaluatePredicate, cfRuntimeEvidence } = await import('../src/xlsx/cf/predicates')
    withExactCfEvidence(PARSE_SET, () => {
      const r = parseSheet(rule('expression', 'A1&gt;0') + `<conditionalFormatting sqref="B3"><cfRule type="cellIs" priority="2" operator="equal"><formula>$F$1</formula></cfRule></conditionalFormatting>`)
      const allow = { allows: () => true }
      void cfRuntimeEvidence
      for (const [i, row] of [[0, 1], [1, 2]] as const) {
        expect(evaluatePredicate(r.rules[i]!, { row, col: 1 }, { kind: 'number', value: 5 }, undefined, allow)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
        expect(evaluatePredicate(r.rules[i]!, { row, col: 1 }, { kind: 'blank' }, undefined, allow)).toEqual({ kind: 'unresolved', reason: 'cf-formula-unevaluated' })
      }
      r.lease.release()
    })
  })
})
