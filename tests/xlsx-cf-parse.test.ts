import { describe, expect, it } from 'vitest'
import { TextEncoder } from 'node:util'
import { getChildren, parseXml, type XmlNode } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { parseCfSheet, type CfParseMeta, type CfStylesSource } from '../src/xlsx/cf/parse'
import type { CfLease, CfRule, CfSheetRules } from '../src/xlsx/cf/model'
import { xlsxBytes } from './helpers/xlsx-package'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AA7539E650}'
const DB_EXT = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'

const DXF_FILL = `<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>`

function worksheet(inner: string, decls = `xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}"`): XmlNode {
  return parseXml(`<worksheet ${decls}>${inner}</worksheet>`)
}

function stylesFor(dxfsInner = ''): CfStylesSource {
  const root = parseXml(`<styleSheet xmlns="${NS}"><dxfs count="1">${dxfsInner}</dxfs></styleSheet>`)
  const dxfs = getChildren(root, 'dxfs')[0]
  return { colors: {}, dxfNodes: dxfs ? getChildren(dxfs, 'dxf') : [] }
}

function parse(
  wsInner: string,
  dxfsInner = DXF_FILL,
  limitBytes?: number,
  decls?: string,
  meta?: CfParseMeta,
) {
  const memory = createCfMemory(limitBytes === undefined ? {} : { limitBytes })
  const lease = parseCfSheet(worksheet(wsInner, decls), stylesFor(dxfsInner), memory, meta)
  return { lease, value: lease.value, memory }
}

function ruleByPriority(rules: readonly CfRule[], priority: number): CfRule {
  const found = rules.filter(r => r.priority === priority)
  expect(found).toHaveLength(1)
  return found[0]!
}

describe('CF rule parsing: all 18 types', () => {
  it('parses the spec example: literal cellIs, rank-3 top10, unresolved colorScale; B2>100 is not a literal', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="B2:B20 D2:D20">` +
        `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule>` +
        `</conditionalFormatting>` +
        `<conditionalFormatting sqref="B2:B20">` +
        `<cfRule type="top10" dxfId="0" priority="2" rank="3"/>` +
        `<cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/></colorScale></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(value.rules).toHaveLength(3)
    const cellIs = ruleByPriority(value.rules, 1)
    expect(cellIs.definition).toMatchObject({
      kind: 'cellIs',
      operator: 'greaterThan',
      operands: [{ kind: 'number', value: 100 }],
    })
    expect(cellIs.documentOrder).toBe(0)
    // Task 2 contract: rule provenance (dialect, formulas, dxfId) rides inside the definition.
    expect(cellIs.definition).toMatchObject({ dialect: 'base', formulas: ['100'], dxfId: 0 })
    expect(cellIs.ranges).toEqual([
      { firstRow: 1, lastRow: 19, firstCol: 1, lastCol: 1 },
      { firstRow: 1, lastRow: 19, firstCol: 3, lastCol: 3 },
    ])
    expect(cellIs.dxf?.['fill.pattern']).toBe('solid')
    const top10 = ruleByPriority(value.rules, 2)
    expect(top10.definition).toMatchObject({ kind: 'top10', rank: 3, bottom: false, percent: false })
    expect(top10.documentOrder).toBe(1)
    const scale = ruleByPriority(value.rules, 3)
    expect(scale.definition).toMatchObject({ kind: 'unresolved', type: 'colorScale' })
    expect(scale.writes.has('colorScale')).toBe(true)

    const bad = parse(
      `<conditionalFormatting sqref="B2:B3">` +
        `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>B2&gt;100</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(bad.value.rules).toHaveLength(1)
    expect(bad.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', type: 'cellIs' })
    expect(bad.value.rules[0]!.definition).toMatchObject({ reason: 'cf-unsupported-rule' })
  })

  it('accepts literal string, boolean and negative operands; trims a single leading =', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="A1:A4">` +
        `<cfRule type="cellIs" priority="1" operator="equal"><formula>"100"</formula></cfRule>` +
        `<cfRule type="cellIs" priority="2" operator="equal"><formula>TRUE</formula></cfRule>` +
        `<cfRule type="cellIs" priority="3" operator="lessThan"><formula>=-3.5</formula></cfRule>` +
        `<cfRule type="cellIs" priority="4" operator="between"><formula>1</formula><formula>3</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(ruleByPriority(value.rules, 1).definition).toMatchObject({
      kind: 'cellIs',
      operands: [{ kind: 'string', value: '100' }],
    })
    expect(ruleByPriority(value.rules, 2).definition).toMatchObject({
      operands: [{ kind: 'boolean', value: true }],
    })
    expect(ruleByPriority(value.rules, 3).definition).toMatchObject({
      operands: [{ kind: 'number', value: -3.5 }],
    })
    expect(ruleByPriority(value.rules, 4).definition).toMatchObject({
      operator: 'between',
      operands: [{ kind: 'number', value: 1 }, { kind: 'number', value: 3 }],
    })
  })

  it('marks non-literal operands and wrong formula counts unresolved without guessing', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="B2:B3">` +
        `<cfRule type="cellIs" priority="1" operator="greaterThan"><formula>$B$1</formula></cfRule>` +
        `<cfRule type="cellIs" priority="2" operator="greaterThan"><formula>SUM(B2:B3)</formula></cfRule>` +
        `<cfRule type="cellIs" priority="3" operator="between"><formula>1</formula></cfRule>` +
        `<cfRule type="cellIs" priority="4" operator="equal"><formula></formula></cfRule>` +
        `<cfRule type="expression" priority="5"><formula>B2&gt;100</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    for (const p of [1, 2, 3, 4, 5]) {
      expect(ruleByPriority(value.rules, p).definition.kind, `priority ${p}`).toBe('unresolved')
    }
    expect(ruleByPriority(value.rules, 5).definition).toMatchObject({ type: 'expression' })
  })

  it('evaluates text, blank and error predicates only for absent or canonical formulas', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="B2:O2">` +
        `<cfRule type="containsText" dxfId="0" priority="1" text="ab"><formula>NOT(ISERROR(SEARCH("ab",B2)))</formula></cfRule>` +
        `<cfRule type="containsText" dxfId="0" priority="2" text="ab"/>` +
        `<cfRule type="containsText" dxfId="0" priority="3" text="ab"><formula>NOT(ISERROR(SEARCH("zz",D2)))</formula></cfRule>` +
        `<cfRule type="beginsWith" dxfId="0" priority="4" text="ab"><formula>LEFT(E2,2)="ab"</formula></cfRule>` +
        `<cfRule type="containsBlanks" dxfId="0" priority="5"><formula>LEN(TRIM(F2))=0</formula></cfRule>` +
        `<cfRule type="containsBlanks" dxfId="0" priority="6"><formula>ISBLANK(G2)</formula></cfRule>` +
        `<cfRule type="containsErrors" dxfId="0" priority="7"><formula>ISNA(H2)</formula></cfRule>` +
        `<cfRule type="notContainsErrors" dxfId="0" priority="8"><formula>NOT(ISERROR(I2))</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(ruleByPriority(value.rules, 1).definition).toMatchObject({ kind: 'text', type: 'containsText', needle: 'ab' })
    expect(ruleByPriority(value.rules, 2).definition).toMatchObject({ kind: 'text', type: 'containsText', needle: 'ab' })
    expect(ruleByPriority(value.rules, 3).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 4).definition).toMatchObject({ kind: 'text', type: 'beginsWith', needle: 'ab' })
    expect(ruleByPriority(value.rules, 5).definition).toMatchObject({ kind: 'predicate', type: 'containsBlanks' })
    expect(ruleByPriority(value.rules, 6).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 7).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 8).definition).toMatchObject({ kind: 'predicate', type: 'notContainsErrors' })
  })

  it('keeps quoted needles exact, including whitespace and escaped quotes', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="F2:G2">` +
        `<cfRule type="containsText" priority="1" text=" ABC "><formula>NOT(ISERROR(SEARCH(" ABC ",F2)))</formula></cfRule>` +
        `<cfRule type="containsText" priority="2" text="a&quot;b"><formula>NOT(ISERROR(SEARCH("a""b",G2)))</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(ruleByPriority(value.rules, 1).definition).toMatchObject({ kind: 'text', type: 'containsText', needle: ' ABC ' })
    expect(ruleByPriority(value.rules, 2).definition).toMatchObject({ kind: 'text', type: 'containsText', needle: 'a"b' })
  })

  it('accepts canonical x14 blank/error rules and refuses activePresent on top10', () => {
    const { value } = parse(
      `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings>` +
        `<x14:conditionalFormatting><xm:sqref>B2:B3</xm:sqref>` +
        `<x14:cfRule type="containsBlanks" priority="1"><xm:f>LEN(TRIM(B2))=0</xm:f></x14:cfRule>` +
        `<x14:cfRule type="top10" priority="2" rank="3" activePresent="1"/>` +
        `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(ruleByPriority(value.rules, 1).definition).toMatchObject({ kind: 'predicate', type: 'containsBlanks' })
    expect(ruleByPriority(value.rules, 2).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
  })

  it('matches the ten canonical timePeriod patterns and refuses divergent or unknown periods', () => {
    const periods: Array<[string, string]> = [
      ['today', 'FLOOR(B2,1)=TODAY()'],
      ['yesterday', 'FLOOR(C2,1)=TODAY()-1'],
      ['tomorrow', 'FLOOR(D2,1)=TODAY()+1'],
      ['last7Days', 'AND(TODAY()-FLOOR(E2,1)&lt;=6,FLOOR(E2,1)&lt;=TODAY())'],
      ['thisWeek', 'AND(TODAY()-ROUNDDOWN(F2,0)&lt;=(WEEKDAY(TODAY())-1),ROUNDDOWN(F2,0)-TODAY()&lt;=(7-WEEKDAY(TODAY())))'],
      ['lastWeek', 'AND(TODAY()-ROUNDDOWN(G2,0)&gt;(WEEKDAY(TODAY())-1),TODAY()-ROUNDDOWN(G2,0)&lt;(WEEKDAY(TODAY())+6))'],
      ['nextWeek', 'AND(ROUNDDOWN(H2,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(H2,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))'],
      ['thisMonth', 'AND(MONTH(I2)=MONTH(TODAY()),YEAR(I2)=YEAR(TODAY()))'],
      ['lastMonth', 'AND(MONTH(J2)=MONTH(EDATE(TODAY(),0-1)),YEAR(J2)=YEAR(EDATE(TODAY(),0-1)))'],
      ['nextMonth', 'AND(MONTH(K2)=MONTH(EDATE(TODAY(),0+1)),YEAR(K2)=YEAR(EDATE(TODAY(),0+1)))'],
    ]
    const rules = periods
      .map(([p, f], i) => `<cfRule type="timePeriod" timePeriod="${p}" priority="${i + 1}"><formula>${f}</formula></cfRule>`)
      .join('')
    const { value } = parse(`<conditionalFormatting sqref="B2:K2">${rules}</conditionalFormatting>`)
    expect(value.unresolvedSheet).toBeUndefined()
    expect(value.rules).toHaveLength(10)
    for (const [p, ,] of periods) {
      const r = value.rules.find(r => (r.definition as { period?: string }).period === p)
      expect(r, p).toBeDefined()
      expect(r!.definition).toMatchObject({ kind: 'timePeriod', period: p })
    }
    const divergent = parse(
      `<conditionalFormatting sqref="B2:B3"><cfRule type="timePeriod" timePeriod="today" priority="1"><formula>FLOOR(B2,1)=TODAY()-2</formula></cfRule></conditionalFormatting>`,
    )
    expect(divergent.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    const unknown = parse(
      `<conditionalFormatting sqref="B2:B3"><cfRule type="timePeriod" timePeriod="someday" priority="1"/></conditionalFormatting>`,
    )
    expect(unknown.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
  })
})

describe('CF x14 pairing by GUID', () => {
  const pairedBase =
    `<conditionalFormatting sqref="C2:C10">` +
    `<cfRule type="dataBar" priority="2"><color rgb="FF638EC6"/><dataBar minLength="10" maxLength="90" showValue="1"><cfvo type="min"/><cfvo type="max"/></dataBar>` +
    `<extLst><ext uri="${DB_EXT}"><x14:id>DBGUID-1</x14:id></ext></extLst></cfRule></conditionalFormatting>`
  const pairedExt =
    `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting>` +
    `<x14:cfRule type="dataBar" id="DBGUID-1"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule>` +
    `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
  const standaloneExt =
    `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>D2:D5</xm:sqref>` +
    `<x14:cfRule type="dataBar" priority="4"><x14:dataBar minLength="0" maxLength="100"><x14:fillColor><x14:color rgb="FF638EC6"/></x14:fillColor><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule>` +
    `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`

  it('pairs one base with one extension into a single rule inheriting priority', () => {
    const { value } = parse(pairedBase + pairedExt)
    expect(value.unresolvedSheet).toBeUndefined()
    expect(value.rules).toHaveLength(1)
    expect(value.rules[0]).toMatchObject({ priority: 2, documentOrder: 0 })
    expect(value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', type: 'dataBar' })
    expect(value.rules[0]!.writes.has('dataBar')).toBe(true)
  })

  it('keeps standalone priority bars with fillColor without a partner', () => {
    const { value } = parse(standaloneExt)
    expect(value.unresolvedSheet).toBeUndefined()
    expect(value.rules).toHaveLength(1)
    expect(value.rules[0]).toMatchObject({ priority: 4, documentOrder: 0 })
    expect(value.rules[0]!.ranges).toEqual([{ firstRow: 1, lastRow: 4, firstCol: 3, lastCol: 3 }])
    expect(value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', type: 'dataBar' })
  })

  it('refuses orphan priority-less extensions and missing targets with a sheet sentinel', () => {
    const orphan = parse(
      `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>E2:E3</xm:sqref>` +
        `<x14:cfRule type="dataBar" id="NOPE"><x14:dataBar minLength="0" maxLength="100"/></x14:cfRule>` +
        `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
    )
    expect(orphan.value.rules).toHaveLength(0)
    expect(orphan.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-invalid-x14-link' })
    expect(orphan.value.diagnostics.some(d => d.feature === 'cf-invalid-x14-link')).toBe(true)

    const missing = parse(pairedBase)
    expect(missing.value.rules).toHaveLength(1)
    expect(missing.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-x14-link' })
    expect([...missing.value.rules[0]!.writes]).toContain('ALL')
    expect(missing.value.unresolvedSheet).toBeUndefined()
  })

  it('refuses duplicate pairing GUIDs with a sheet sentinel', () => {
    const dupBase =
      `<conditionalFormatting sqref="C2:C3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/></dataBar>` +
      `<extLst><ext uri="${DB_EXT}"><x14:id>DUP</x14:id></ext></extLst></cfRule></conditionalFormatting>` +
      `<conditionalFormatting sqref="D2:D3"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/></dataBar>` +
      `<extLst><ext uri="${DB_EXT}"><x14:id>DUP</x14:id></ext></extLst></cfRule></conditionalFormatting>`
    const { value } = parse(dupBase + pairedExt.replace(/DBGUID-1/g, 'DUP'))
    // The priority-less extension cannot be ordered once pairing is ambiguous:
    // withhold the sheet rather than guess positions.
    expect(value.rules).toHaveLength(0)
    expect(value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-invalid-x14-link' })
    expect(value.diagnostics.some(d => d.feature === 'cf-invalid-x14-link')).toBe(true)
  })

  it('rejects wrong-namespace prefixes and accepts alternate prefixes with the correct URI', () => {
    const wrong = parse(
      `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings xmlns:x14="urn:wrong">` +
        `<x14:conditionalFormatting><xm:sqref>F2:F3</xm:sqref>` +
        `<x14:cfRule type="cellIs" priority="9" operator="greaterThan"><xm:f>100</xm:f></x14:cfRule>` +
        `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
    )
    expect(wrong.value.rules).toHaveLength(0)
    expect(wrong.value.diagnostics.some(d => d.feature === 'cf-namespace')).toBe(true)

    const alt = parseXml(
      `<worksheet xmlns="${NS}" xmlns:foo="${X14}" xmlns:xm="${XM}"><extLst><ext uri="${CF_EXT}">` +
        `<foo:conditionalFormattings><foo:conditionalFormatting><xm:sqref>F2:F3</xm:sqref>` +
        `<foo:cfRule type="cellIs" priority="9" operator="greaterThan"><xm:f>100</xm:f></foo:cfRule>` +
        `</foo:conditionalFormatting></foo:conditionalFormattings></ext></extLst></worksheet>`,
    )
    const memory = createCfMemory({})
    const lease = parseCfSheet(alt, stylesFor(DXF_FILL), memory)
    expect(lease.value.rules).toHaveLength(1)
    expect(lease.value.rules[0]).toMatchObject({ priority: 9 })
    expect(lease.value.rules[0]!.definition).toMatchObject({
      kind: 'cellIs',
      operands: [{ kind: 'number', value: 100 }],
    })
  })
})

describe('CF ranges, caps and invalid attributes', () => {
  it('parses $, whole columns, whole rows and multi-range sqref', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="$B$2:$C$3 M:M 2:2 Z5">` +
        `<cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(value.rules[0]!.ranges).toEqual([
      { firstRow: 1, lastRow: 2, firstCol: 1, lastCol: 2 },
      { firstRow: 0, lastRow: 1048575, firstCol: 12, lastCol: 12 },
      { firstRow: 1, lastRow: 1, firstCol: 0, lastCol: 16383 },
      { firstRow: 4, lastRow: 4, firstCol: 25, lastCol: 25 },
    ])
  })

  it('withholds the sheet on missing or invalid scope', () => {
    for (const inner of [
      `<conditionalFormatting><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`,
      `<conditionalFormatting sqref="B2:ZZZ"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`,
      `<conditionalFormatting sqref=""><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`,
    ]) {
      const { value } = parse(inner)
      expect(value.rules, inner).toHaveLength(0)
      expect(value.unresolvedSheet).toMatchObject({ writes: 'ALL' })
    }
  })

  it('withholds the sheet past 4096 rules and past 256 ranges per rule', () => {
    const many = Array.from({ length: 4097 }, (_, i) => `<cfRule type="cellIs" priority="${i + 1}" operator="greaterThan"><formula>1</formula></cfRule>`).join('')
    const over = parse(`<conditionalFormatting sqref="A1">${many}</conditionalFormatting>`)
    expect(over.value.rules).toHaveLength(0)
    expect(over.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-resource-limit' })

    const wide = Array.from({ length: 257 }, (_, i) => `A${i + 1}`).join(' ')
    const wideRes = parse(`<conditionalFormatting sqref="${wide}"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`)
    // Per-rule interval overflow withholds only that rule (scope is known);
    // sheet-wide preconditions (rule count, total ranges, scope) use the sentinel.
    expect(wideRes.value.unresolvedSheet).toBeUndefined()
    expect(wideRes.value.rules).toHaveLength(1)
    expect(wideRes.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    // Refused before dxf resolution, so the write mask stays maximally conservative.
    expect([...wideRes.value.rules[0]!.writes]).toContain('ALL')
  })

  it('refuses over-long text, excess formulas, cfvos and custom icons at rule level', () => {
    const long = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="containsText" dxfId="0" priority="1" text="${'a'.repeat(9000)}"/></conditionalFormatting>`,
    )
    expect(long.value.unresolvedSheet).toBeUndefined()
    expect(long.value.rules).toHaveLength(1)
    expect(long.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    const manyFormulas = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan">` +
        `<formula>1</formula><formula>2</formula><formula>3</formula><formula>4</formula></cfRule></conditionalFormatting>`,
    )
    expect(manyFormulas.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    const manyCfvo = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="num" val="1"/><cfvo type="num" val="2"/><cfvo type="num" val="3"/><cfvo type="num" val="4"/><cfvo type="num" val="5"/><cfvo type="num" val="6"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(manyCfvo.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    const noType = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo val="1"/></colorScale></cfRule></conditionalFormatting>`,
    )
    expect(noType.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-cfvo' })
  })

  it('refuses invalid priorities, operators, periods and boolean spellings without guessing', () => {
    const bad = parse(
      `<conditionalFormatting sqref="A1">` +
        `<cfRule type="cellIs" priority="0" operator="greaterThan"><formula>1</formula></cfRule>` +
        `<cfRule type="cellIs" priority="2" operator="newerThan"><formula>1</formula></cfRule>` +
        `<cfRule type="cellIs" priority="3" operator="greaterThan" stopIfTrue="yes"><formula>1</formula></cfRule>` +
        `</conditionalFormatting>`,
    )
    expect(bad.value.rules).toHaveLength(3)
    for (const r of bad.value.rules) expect(r.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
    const missing = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="cellIs" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>`,
    )
    expect(missing.value.rules[0]!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
  })

  it('keeps document order across elements and equal priorities', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="20" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>` +
        `<conditionalFormatting sqref="B2"><cfRule type="cellIs" priority="20" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>`,
    )
    expect(value.rules.map(r => r.documentOrder)).toEqual([0, 1])
    expect(value.rules.map(r => r.priority)).toEqual([20, 20])
  })

  it('fails top10 and aboveAverage edges closed per the measured bounds', () => {
    const { value } = parse(
      `<conditionalFormatting sqref="A1:A9">` +
        `<cfRule type="top10" priority="1"/>` +
        `<cfRule type="top10" priority="2" rank="0"/>` +
        `<cfRule type="top10" priority="3" rank="1001"/>` +
        `<cfRule type="top10" priority="4" rank="50" percent="1"/>` +
        `<cfRule type="aboveAverage" priority="5" stdDev="4"/>` +
        `<cfRule type="aboveAverage" priority="6" stdDev="1" equalAverage="1"/>` +
        `<cfRule type="aboveAverage" priority="7"/>` +
        `<cfRule type="duplicateValues" priority="8"/>` +
        `<cfRule type="uniqueValues" priority="9"/>` +
        `</conditionalFormatting>`,
    )
    expect(value.unresolvedSheet).toBeUndefined()
    expect(ruleByPriority(value.rules, 1).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 2).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 3).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 4).definition).toMatchObject({ kind: 'top10', rank: 50, bottom: false, percent: true })
    expect(ruleByPriority(value.rules, 5).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(ruleByPriority(value.rules, 6).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
    expect(ruleByPriority(value.rules, 7).definition).toMatchObject({ kind: 'aboveAverage', above: true, equal: false })
    expect(ruleByPriority(value.rules, 8).definition).toMatchObject({ kind: 'predicate', type: 'duplicateValues' })
    expect(ruleByPriority(value.rules, 9).definition).toMatchObject({ kind: 'predicate', type: 'uniqueValues' })
  })

  it('refuses unknown types and bad dxfIds with ALL writes, and keeps stopIfTrue', () => {
    const { value, lease } = parse(
      `<conditionalFormatting sqref="B2:B3">` +
        `<cfRule type="mystery" priority="1" stopIfTrue="1"/>` +
        `<cfRule type="cellIs" dxfId="99" priority="2" operator="greaterThan"><formula>1</formula></cfRule>` +
        `<cfRule type="cellIs" dxfId="abc" priority="3" operator="greaterThan"><formula>1</formula></cfRule>` +
        `</conditionalFormatting>`,
      DXF_FILL,
    )
    expect(lease.value.rules).toHaveLength(3)
    const [mystery, badId, garbageId] = value.rules
    expect(mystery!.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect([...mystery!.writes]).toContain('ALL')
    expect(mystery!.stopIfTrue).toBe(true)
    for (const r of [badId!, garbageId!]) {
      expect(r.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-dxf' })
      expect([...r.writes]).toContain('ALL')
    }
  })

  it('admits memory before allocating and settles the rules lease exactly once', () => {
    const { lease, memory } = parse(
      `<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>`,
    )
    expect(lease.value.rules).toHaveLength(1)
    // Rules, diagnostics and metadata all charge the shared account up front.
    const retained = memory.retainedBytes
    expect(retained).toBeGreaterThan(0)
    expect(memory.currentBytes).toBeGreaterThanOrEqual(retained)
    expect(memory.peakBytes).toBeGreaterThanOrEqual(memory.currentBytes)
    const sharedHeld = memory.currentBytes - retained
    lease.release()
    // The rules charge returns exactly once; sheet-lifetime diagnostics and
    // metadata stay admitted until the sheet (and its retained handles) drop.
    expect(memory.retainedBytes).toBe(0)
    expect(memory.currentBytes).toBe(sharedHeld)
    lease.release()
    expect(memory.currentBytes).toBe(sharedHeld)
  })

  it('fails stop instead of partial output when the account is exhausted', () => {
    const memory = createCfMemory({ limitBytes: 64 })
    expect(() =>
      parseCfSheet(
        worksheet(
          `<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>`,
        ),
        stylesFor(DXF_FILL),
        memory,
      ),
    ).toThrow()
    // No partial charge survives the refusal: fail-stop, never half-admitted.
    expect(memory.currentBytes).toBe(0)
  })
})

describe('CF streaming capture parity', () => {
  const doc =
    `<worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}">` +
    `<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>` +
    `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>D2:D5</xm:sqref>` +
    `<x14:cfRule type="beginsWith" text="ab" priority="4"><xm:f>LEFT(D2,2)="ab"</xm:f></x14:cfRule>` +
    `</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>`

  async function viaStream(xml: string, chunkBytes = 1024): Promise<CfLease<CfSheetRules>> {
    const bytes = new TextEncoder().encode(xml)
    const chunks: Uint8Array[] = []
    for (let i = 0; i < bytes.length; i += chunkBytes) chunks.push(bytes.slice(i, i + chunkBytes))
    async function* source() {
      for (const c of chunks) yield c
    }
    const memory = createCfMemory({})
    const stream = createCfStreamEvents(stylesFor(DXF_FILL), memory, {})
    await tokenizeBytes(source(), stream.events)
    return stream.finish()
  }

  it('matches the resident parser rule-for-rule through the real tokenizer', async () => {
    const memory = createCfMemory({})
    const resident = parseCfSheet(parseXml(doc), stylesFor(DXF_FILL), memory, {})
    const streamed = await viaStream(doc, 7)
    const strip = (v: unknown): unknown =>
      v === undefined ? undefined : JSON.parse(JSON.stringify(v, (_k, val) => (val instanceof Set ? [...val].sort() : val)))
    expect(strip(streamed.value.rules)).toEqual(strip(resident.value.rules))
    expect(strip(streamed.value.diagnostics)).toEqual(strip(resident.value.diagnostics))
    expect(strip(streamed.value.unresolvedSheet)).toEqual(strip(resident.value.unresolvedSheet))
    streamed.release()
    resident.release()
  })

  it('fails an unclosed capture closed instead of publishing partial rules', async () => {
    // A stream that ends mid-subtree (no tokenizer close events) must withhold everything.
    const memory = createCfMemory({})
    const stream = createCfStreamEvents(stylesFor(DXF_FILL), memory, {})
    // Feed only well-formed open events for the first element, then stop.
    stream.events.open('worksheet', [], [], 0, false)
    stream.events.open('conditionalFormatting', ['sqref'], ['B2:B3'], 1, false)
    stream.events.open('cfRule', ['type', 'priority', 'operator'], ['cellIs', '1', 'greaterThan'], 3, false)
    const value = stream.finish().value
    expect(value.rules).toHaveLength(0)
    expect(value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-resource-limit' })
  })

  it('withholds the sheet when a streamed text node exceeds the pre-screen bound', async () => {
    // The streaming pre-screen is deliberately coarser than the resident path:
    // both refuse oversized input, but the streamer withholds the sheet while
    // the resident parser can still refuse at rule level. Fail-closed either way.
    const big = doc.replace('<formula>100</formula>', `<formula>${'1'.repeat(9000)}</formula>`)
    const streamed = await viaStream(big, 13)
    expect(streamed.value.rules).toHaveLength(0)
    expect(streamed.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-resource-limit' })
  })
})

describe('CF resident workbook integration', () => {
  it('attaches parsed rules and dxf source on a CF-only sheet without table metadata', async () => {
    const styles =
      `<numFmts count="0"/><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<dxfs count="1">${DXF_FILL}</dxfs>`
    const pkg = await OfficePackage.load(
      xlsxBytes({
        sheetData: `<row r="2"><c r="B2"><v>150</v></c></row>`,
        afterSheetData:
          `<conditionalFormatting sqref="B2:B3">` +
          `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule>` +
          `</conditionalFormatting>`,
        styles,
      }),
    )
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]!
    expect(sheet.conditionalFormatting?.rules).toHaveLength(1)
    expect(sheet.conditionalFormatting?.rules[0]).toMatchObject({ priority: 1, documentOrder: 0 })
    expect(sheet.conditionalFormatting?.rules[0]?.definition).toMatchObject({
      kind: 'cellIs',
      operands: [{ kind: 'number', value: 100 }],
    })
    // CF-only sheets still receive dxfs for the lazy painter.
    expect(sheet.tableStyleSource).toBeDefined()
    expect(sheet.diagnostics ?? []).toEqual([])
  })

  it('leaves CF-free sheets untouched: no rules, no dxf source, no import', async () => {
    const pkg = await OfficePackage.load(xlsxBytes({ sheetData: `<row r="1"><c r="A1"><v>1</v></c></row>` }))
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]!
    expect(sheet.conditionalFormatting).toBeUndefined()
    expect(sheet.tableStyleSource).toBeUndefined()
  })
})
