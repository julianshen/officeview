import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { parseXml } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { CF_BAR_DEFINITION_BYTES, createCfMemory } from '../src/xlsx/cf/limits'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { injectExactCfEvidence, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { MEASURED_CF_OFFICE_THEME, barRequiredCapabilities, parseCfSheet, type CfParseMeta, type CfStylesSource } from '../src/xlsx/cf/parse'
import { DATA_BAR_WRITES, type CfDefinition, type CfRule, type CfSheetRules } from '../src/xlsx/cf/model'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { resolveSheetColor } from '../src/xlsx/sheetColors'
import { xlsxBytes } from './helpers/xlsx-package'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const DB_EXT = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
const REPO_ROOT = resolve(__dirname, '..')
const LTR: CfParseMeta = { sheetRightToLeft: false }

/** Every capability the parse gate can ask for: with all of them injected only structure decides. */
const BAR_PARSE_CAPS = [
  'bar:settings-schema-defaults', 'bar:settings-paired-precedence',
  'bar:pairing-paired', 'bar:pairing-standalone', 'bar:pairing-base-only',
  'bar:direction-ltr', 'bar:direction-rtl', 'bar:direction-context-ltr-sheet',
  'bar:colors-theme-tint', 'bar:colors-indexed',
  'bar:bounds-positional', 'bar:bounds-out-of-domain', 'bar:degenerate-inverted',
]
const SETTINGS = ['bar:settings-precondition-probe', 'bar:settings-schema-defaults', 'bar:settings-paired-precedence']

const styles: CfStylesSource = { colors: { theme: MEASURED_CF_OFFICE_THEME }, dxfNodes: [] }

function worksheet(inner: string): string {
  return `<worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}">${inner}</worksheet>`
}
function parse(inner: string, meta: CfParseMeta = LTR, st: CfStylesSource = styles) {
  const memory = createCfMemory()
  const lease = parseCfSheet(parseXml(worksheet(inner)), st, memory, meta)
  return { lease, value: lease.value, memory }
}
async function stream(inner: string, meta: CfParseMeta = LTR, st: CfStylesSource = styles) {
  const memory = createCfMemory(), handle = createCfStreamEvents(st, memory, meta)
  const bytes = new TextEncoder().encode(worksheet(inner))
  await tokenizeBytes((async function* () { yield bytes })(), handle.events)
  return handle.finish()
}
const only = (rules: readonly CfRule[]): CfRule => {
  expect(rules).toHaveLength(1)
  return rules[0]!
}
const byPriority = (rules: readonly CfRule[], priority: number): CfRule => {
  const found = rules.filter(r => r.priority === priority)
  expect(found, `priority ${priority}`).toHaveLength(1)
  return found[0]!
}
const bar = (definition: CfDefinition): Extract<CfDefinition, { kind: 'dataBar' }> => {
  expect(definition.kind).toBe('dataBar')
  return definition as Extract<CfDefinition, { kind: 'dataBar' }>
}

/** A base `<cfRule type="dataBar">` (Excel's canonical order: cfvo, cfvo, color) and its optional `<x14:id>` link. */
const baseBar = (sqref: string, priority: number, body: string, guid?: string, ruleAttrs = '', barAttrs = '') =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="dataBar" priority="${priority}"${ruleAttrs}><dataBar${barAttrs}>${body}</dataBar>` +
  (guid ? `<extLst><ext uri="${DB_EXT}"><x14:id>${guid}</x14:id></ext></extLst>` : '') +
  `</cfRule></conditionalFormatting>`
const x14Block = (rules: string) =>
  `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings>${rules}</x14:conditionalFormattings></ext></extLst>`
const companion = (sqref: string, guid: string, attrs: string, body: string, ruleAttrs = '') =>
  `<x14:conditionalFormatting><x14:cfRule type="dataBar" id="${guid}"${ruleAttrs}><x14:dataBar${attrs}>${body}</x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`
const MINMAX = '<cfvo type="min"/><cfvo type="max"/>'
const BLUE = '<color rgb="FF638EC6"/>'

afterEach(() => injectExactCfEvidence(null))

describe('CF Stage 3 data-bar model and parsing (Task 19)', () => {
  it('canonical data-bar write mask is [\'dataBar\'] for both resolved and unresolved bars', () => {
    expect(DATA_BAR_WRITES).toEqual(new Set(['dataBar']))
    const xml = baseBar('B2:B6', 1, MINMAX + BLUE)
    const resolved = withExactCfEvidence(BAR_PARSE_CAPS, () => parse(xml))
    const r = only(resolved.value.rules)
    expect(r.definition.kind).toBe('dataBar')
    expect(new Set(r.writes)).toEqual(DATA_BAR_WRITES)

    const unresolved = withExactCfEvidence([], () => parse(xml))
    const u = only(unresolved.value.rules)
    expect(u.definition).toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
    expect(new Set(u.writes)).toEqual(DATA_BAR_WRITES)

    // A broken link keeps main's invalidLinkRule mask: ALL.
    const broken = withExactCfEvidence(BAR_PARSE_CAPS, () => parse(baseBar('B2:B6', 1, MINMAX + BLUE, '{NO-COMPANION}')))
    const b = only(broken.value.rules)
    expect(b.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-x14-link' })
    expect(new Set(b.writes)).toEqual(new Set(['ALL']))
  })

  it('Excel 2010 paired bar resolves to one definition with x14 precedence', () => {
    const G = '{7E970C84-77E5-6843-BF55-BA32DBB9E43A}'
    const xml =
      baseBar('B2:B6', 1, MINMAX + BLUE, G) +
      x14Block(companion('B2:B6', G, ' minLength="0" maxLength="100" gradient="0" border="1" axisPosition="middle" negativeBarBorderColorSameAsPositive="0"',
        '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/>' +
        '<x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/>'))
    // `direction` is omitted, so it is `context`: S3-K needs a known-LTR sheet and its own capability.
    const evidence = [...SETTINGS, 'bar:pairing-paired', 'bar:direction-context-ltr-sheet']
    const { value } = withExactCfEvidence(evidence, () => parse(xml))
    expect(value.unresolvedSheet).toBeUndefined()
    const rule = only(value.rules)
    expect(rule.priority).toBe(1)
    expect(bar(rule.definition)).toMatchObject({
      kind: 'dataBar', dialect: 'base',
      bounds: [{ type: 'autoMin' }, { type: 'autoMax' }],
      minLength: 0, maxLength: 100,
      fillArgb: 0xff638ec6, borderArgb: 0xff638ec6, negativeFillArgb: 0xffff0000, negativeBorderArgb: 0xffff0000, axisArgb: 0xff000000,
      gradient: false, border: true, axisPosition: 'middle', direction: 'context', showValue: true,
      negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: false, sheetRightToLeft: false,
    })
    expect(new Set(rule.writes)).toEqual(DATA_BAR_WRITES)

    // Each gate is required: without the paired-precedence setting, or the context direction, the bar withholds.
    for (const missing of ['bar:settings-paired-precedence', 'bar:pairing-paired', 'bar:direction-context-ltr-sheet', 'bar:settings-schema-defaults']) {
      const partial = withExactCfEvidence(evidence.filter(e => e !== missing), () => parse(xml))
      expect(only(partial.value.rules).definition, missing).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    }

    const none = withExactCfEvidence([], () => parse(xml))
    const u = only(none.value.rules)
    expect(u.definition).toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
    expect(new Set(u.writes)).toEqual(DATA_BAR_WRITES)
    expect(none.value.diagnostics.filter(d => d.feature === 'cf-unsupported-rule')).toHaveLength(1)
  })

  it('Excel\'s own paired form (excel-authored-databar.xlsx) keeps the base showValue and the x14 settings', async () => {
    const bytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, 'tests/fixtures/xlsx-cf/excel-authored-databar.xlsx')))
    injectExactCfEvidence([...SETTINGS, 'bar:pairing-paired', 'bar:direction-rtl'])
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    const rule = only(doc.sheets[0]!.conditionalFormatting!.rules)
    expect(bar(rule.definition)).toMatchObject({
      bounds: [{ type: 'autoMin' }, { type: 'autoMax' }], minLength: 0, maxLength: 100,
      fillArgb: 0xff638ec6, borderArgb: 0xff000000, negativeFillArgb: 0xffff0000, axisArgb: 0xff000000,
      gradient: false, border: true, direction: 'rightToLeft', axisPosition: 'middle',
      // showValue="0" lives on the base <dataBar> only; x14 precedence must not reset it to the default.
      showValue: false,
      negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true, sheetRightToLeft: false,
    })
    // The negative border is drawn in the positive border colour (same-as-positive), so no separate colour is kept.
    expect(bar(rule.definition).negativeBorderArgb).toBeUndefined()
  })

  it('Excel-resaved Stage 3 fixtures parse as Excel wrote them; auto colours Excel adds on resave stay unresolved', async () => {
    const load = async (name: string) => {
      const bytes = new Uint8Array(readFileSync(resolve(REPO_ROOT, `tests/fixtures/xlsx-cf/${name}.xlsx`)))
      return (await parseXlsx(await OfficePackage.load(bytes))).sheets[0]!.conditionalFormatting!
    }
    injectExactCfEvidence([...SETTINGS, ...BAR_PARSE_CAPS])
    // Probe: the configured rule carries <x14:axisColor auto="1"/> under a middle axis; the default control is base-only.
    const probe = await load('stage3-bar-settings-probe')
    expect(byPriority(probe.rules, 1).definition).toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
    expect(bar(byPriority(probe.rules, 2).definition)).toMatchObject({
      bounds: [{ type: 'min' }, { type: 'max' }], minLength: 10, maxLength: 90, fillArgb: 0xff638ec6, showValue: true,
      gradient: true, border: false, axisPosition: 'automatic', direction: 'context', sheetRightToLeft: false,
    })
    // Pairing: Excel resaves every bar (the standalone one included) as a paired bar with auto negative and axis colours.
    const pairing = await load('stage3-bar-pairing')
    for (const p of [1, 2, 4]) expect(byPriority(pairing.rules, p).definition, `pairing p${p}`).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(bar(byPriority(pairing.rules, 3).definition)).toMatchObject({ fillArgb: 0xff0000ff, minLength: 10, maxLength: 90 })
    // Axis/negative: explicit colours resolve; axis "none" needs no axis colour.
    const axis = await load('stage3-bar-axis-negative')
    expect(bar(byPriority(axis.rules, 2).definition)).toMatchObject({ axisPosition: 'middle', axisArgb: 0xff000000, negativeFillArgb: 0xffff0000 })
    const noAxis = bar(byPriority(axis.rules, 3).definition)
    expect(noAxis).toMatchObject({ axisPosition: 'none', negativeFillArgb: 0xffff0000 })
    expect(noAxis.axisArgb).toBeUndefined()
    expect(bar(byPriority(axis.rules, 5).definition)).toMatchObject({ negativeBarColorSameAsPositive: true })
    expect(bar(byPriority(axis.rules, 5).definition).negativeFillArgb).toBeUndefined()
    expect(bar(byPriority(axis.rules, 6).definition)).toMatchObject({
      border: true, borderArgb: 0xff00b050, negativeBarBorderColorSameAsPositive: false, negativeBorderArgb: 0xffffa500,
    })
    // Composition: every bar carries Excel's auto negative/axis colours, so none resolves; the DXF rules are untouched.
    const composition = await load('stage3-bar-composition')
    const bars = composition.rules.filter(r => r.definition.kind === 'dataBar' || (r.definition.kind === 'unresolved' && r.definition.type === 'dataBar'))
    expect(bars.length).toBeGreaterThan(0)
    for (const r of bars) expect(r.definition, `composition p${r.priority}`).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    expect(byPriority(composition.rules, 1).definition.kind).toBe('cellIs')
  })

  it('schema defaults apply to omitted attributes', () => {
    const base = baseBar('B2:B6', 1, MINMAX + BLUE)
    const caps = ['bar:settings-schema-defaults', 'bar:pairing-base-only', 'bar:direction-context-ltr-sheet']
    const baseOnly = bar(only(withExactCfEvidence(caps, () => parse(base)).value.rules).definition)
    expect(baseOnly).toMatchObject({
      minLength: 10, maxLength: 90, showValue: true, gradient: true, border: false,
      negativeBarColorSameAsPositive: false, negativeBarBorderColorSameAsPositive: true, axisPosition: 'automatic', direction: 'context',
    })
    expect(only(withExactCfEvidence(caps.filter(c => c !== 'bar:settings-schema-defaults'), () => parse(base)).value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })

    const G = '{00000000-0000-4000-8000-0000000000D0}'
    const paired = baseBar('B2:B6', 1, MINMAX + BLUE, G) + x14Block(companion('B2:B6', G, '', '<x14:cfvo type="min"/><x14:cfvo type="max"/>'))
    const pcaps = ['bar:settings-schema-defaults', 'bar:settings-paired-precedence', 'bar:pairing-paired', 'bar:direction-context-ltr-sheet']
    const x14 = bar(only(withExactCfEvidence(pcaps, () => parse(paired)).value.rules).definition)
    expect(x14).toMatchObject({
      minLength: 10, maxLength: 90, gradient: true, border: false, negativeBarColorSameAsPositive: false,
      negativeBarBorderColorSameAsPositive: true, axisPosition: 'automatic', direction: 'context', showValue: true,
    })
    // Missing optional colours stay absent: the evaluator withholds a cell that needs one (S3-F).
    expect(x14.negativeFillArgb).toBeUndefined()
    expect(x14.axisArgb).toBeUndefined()
    expect(only(withExactCfEvidence(pcaps.filter(c => c !== 'bar:settings-schema-defaults'), () => parse(paired)).value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })

  it('standalone x14 bar uses its own fillColor and priority', () => {
    const xml = x14Block(
      `<x14:conditionalFormatting><x14:cfRule type="dataBar" priority="4" id="{00000000-0000-4000-8000-0000000000A1}">` +
      `<x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="min"/><x14:cfvo type="max"/>` +
      `<x14:fillColor rgb="FF00B050"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>` +
      `<xm:sqref>D2:D5</xm:sqref></x14:conditionalFormatting>`)
    const caps = ['bar:settings-schema-defaults', 'bar:pairing-standalone', 'bar:direction-context-ltr-sheet']
    const rule = only(withExactCfEvidence(caps, () => parse(xml)).value.rules)
    expect(rule.priority).toBe(4)
    expect(rule.ranges).toEqual([{ firstRow: 1, lastRow: 4, firstCol: 3, lastCol: 3 }])
    expect(bar(rule.definition)).toMatchObject({
      dialect: 'x14', fillArgb: 0xff00b050, negativeFillArgb: 0xffff0000, axisArgb: 0xff000000,
      minLength: 0, maxLength: 100, gradient: false, showValue: true,
    })
    expect(new Set(rule.writes)).toEqual(DATA_BAR_WRITES)
    expect(only(withExactCfEvidence(caps.filter(c => c !== 'bar:pairing-standalone'), () => parse(xml)).value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    // A standalone bar without a fill colour has nothing to draw.
    const noFill = xml.replace('<x14:fillColor rgb="FF00B050"/>', '')
    expect(only(withExactCfEvidence(caps, () => parse(noFill)).value.rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })

  it('x14 literal xm:f bounds resolve; references stay unsupported', () => {
    const G = '{00000000-0000-4000-8000-0000000000F0}'
    const at = (f: string) =>
      baseBar('B2:B6', 1, '<cfvo type="num" val="0"/><cfvo type="max"/>' + BLUE, G) +
      x14Block(companion('B2:B6', G, ' minLength="0" maxLength="100"', `<x14:cfvo type="num"><xm:f>${f}</xm:f></x14:cfvo><x14:cfvo type="autoMax"/>`))
    const literal = bar(only(withExactCfEvidence(BAR_PARSE_CAPS, () => parse(at('20'))).value.rules).definition)
    expect(literal.bounds).toEqual([{ type: 'num', value: 20 }, { type: 'autoMax' }])
    for (const f of ['$A$1', 'A1', 'SUM(A1:A5)']) {
      expect(only(withExactCfEvidence(BAR_PARSE_CAPS, () => parse(at(f))).value.rules).definition, f)
        .toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
    }
    expect(only(withExactCfEvidence(BAR_PARSE_CAPS, () => parse(at('abc"'))).value.rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-cfvo' })
  })

  it('bar structure refusals', () => {
    injectExactCfEvidence(BAR_PARSE_CAPS)
    const reason = (inner: string) => only(parse(inner).value.rules).definition
    const unresolved = (r: string) => ({ kind: 'unresolved', type: 'dataBar', reason: r })
    // S3-C counts: exactly two cfvo.
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="min"/>' + BLUE))).toMatchObject(unresolved('cf-invalid-rule'))
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-invalid-rule'))
    expect(new Set(only(parse(baseBar('A1:A5', 1, '<cfvo type="min"/>' + BLUE)).value.rules).writes)).toEqual(DATA_BAR_WRITES)
    const six = parse(baseBar('A1:A5', 1, '<cfvo type="min"/><cfvo type="num" val="1"/><cfvo type="num" val="2"/><cfvo type="num" val="3"/><cfvo type="num" val="4"/><cfvo type="max"/>' + BLUE))
    expect(only(six.value.rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(six.value.diagnostics.some(d => d.feature === 'cf-resource-limit')).toBe(true)
    expect(reason(baseBar('A1:A5', 1, '<cfvo val="0"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-invalid-cfvo'))
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="bogus"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-invalid-cfvo'))
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="num"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-invalid-cfvo'))
    // Stage 5 shapes.
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="formula" val="SUM(A1:A5)"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-unsupported-rule'))
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="num" val="$A$1"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-unsupported-rule'))
    for (const spelling of ['1', 'true']) {
      expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, undefined, ` activePresent="${spelling}"`)), spelling).toMatchObject(unresolved('cf-unsupported-rule'))
    }
    const A = '{00000000-0000-4000-8000-0000000000A0}'
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, A) + x14Block(companion('A1:A5', A, '', '<x14:cfvo type="min"/><x14:cfvo type="max"/>', ' activePresent="1"'))))
      .toMatchObject(unresolved('cf-unsupported-rule'))
    // autoMin/autoMax resolve only on <x14:cfvo>.
    expect(reason(baseBar('A1:A5', 1, '<cfvo type="autoMin"/><cfvo type="max"/>' + BLUE))).toMatchObject(unresolved('cf-unsupported-rule'))
    // S3-D positional and legacy shapes: unresolved until bar:bounds-positional.
    const legacy = baseBar('A1:A5', 1, '<cfvo type="min" val="0"/><cfvo type="max"/>' + BLUE)
    const maxFirst = baseBar('A1:A5', 1, '<cfvo type="max"/><cfvo type="min"/>' + BLUE)
    withExactCfEvidence(BAR_PARSE_CAPS.filter(c => c !== 'bar:bounds-positional'), () => {
      expect(reason(legacy)).toMatchObject(unresolved('cf-unsupported-rule'))
      expect(reason(maxFirst)).toMatchObject(unresolved('cf-unsupported-rule'))
    })
    expect(bar(reason(legacy)).bounds).toEqual([{ type: 'min' }, { type: 'max' }])
    expect(bar(reason(maxFirst)).bounds).toEqual([{ type: 'max' }, { type: 'min' }])
    // S3-E statically visible degenerate bounds.
    withExactCfEvidence(BAR_PARSE_CAPS.filter(c => c !== 'bar:bounds-out-of-domain' && c !== 'bar:degenerate-inverted'), () => {
      expect(reason(baseBar('A1:A5', 1, '<cfvo type="percent" val="-10"/><cfvo type="percent" val="150"/>' + BLUE))).toMatchObject(unresolved('cf-unsupported-rule'))
      expect(reason(baseBar('A1:A5', 1, '<cfvo type="num" val="80"/><cfvo type="num" val="20"/>' + BLUE))).toMatchObject(unresolved('cf-unsupported-rule'))
      expect(bar(reason(baseBar('A1:A5', 1, '<cfvo type="num" val="20"/><cfvo type="num" val="80"/>' + BLUE))).bounds)
        .toEqual([{ type: 'num', value: 20 }, { type: 'num', value: 80 }])
    })
    // S3-F colour forms: no #000000 fallback, whatever the evidence.
    for (const colour of ['<color auto="1"/>', '<color rgb="000000FF"/>', '<color theme="99"/>', '<color indexed="64"/>', '']) {
      const def = reason(baseBar('A1:A5', 1, MINMAX + colour))
      expect(def, colour).toMatchObject(unresolved('cf-unsupported-rule'))
      expect(def).not.toHaveProperty('fillArgb')
    }
    const borderAuto = '{00000000-0000-4000-8000-0000000000B1}'
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, borderAuto) + x14Block(companion('A1:A5', borderAuto, ' border="1"', '<x14:cfvo type="min"/><x14:cfvo type="max"/><x14:borderColor auto="1"/>'))))
      .toMatchObject(unresolved('cf-unsupported-rule'))
    const borderMissing = '{00000000-0000-4000-8000-0000000000B2}'
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, borderMissing) + x14Block(companion('A1:A5', borderMissing, ' border="1"', '<x14:cfvo type="min"/><x14:cfvo type="max"/>'))))
      .toMatchObject(unresolved('cf-unsupported-rule'))
    // Invalid attribute values never default.
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, undefined, '', ' showValue="maybe"'))).toMatchObject(unresolved('cf-invalid-rule'))
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, undefined, '', ' minLength="-1"'))).toMatchObject(unresolved('cf-invalid-rule'))
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, undefined, '', ' minLength="60" maxLength="40"'))).toMatchObject(unresolved('cf-unsupported-rule'))
    const badEnum = '{00000000-0000-4000-8000-0000000000E1}'
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, badEnum) + x14Block(companion('A1:A5', badEnum, ' axisPosition="left"', '<x14:cfvo type="min"/><x14:cfvo type="max"/>'))))
      .toMatchObject(unresolved('cf-invalid-rule'))
    // A dataBar carrying a DXF is not Excel's shape: withheld, never painted with its DXF.
    expect(reason(baseBar('A1:A5', 1, MINMAX + BLUE, undefined, ' dxfId="0"'))).toMatchObject({ kind: 'unresolved', type: 'dataBar' })
    // S3-Q broken links keep main's cf-invalid-x14-link handling.
    const orphan = parse(x14Block(companion('E2:E3', '{ORPHAN}', ' minLength="0" maxLength="100"', '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>')))
    expect(orphan.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-invalid-x14-link' })
    const D = '{00000000-0000-4000-8000-0000000000DD}'
    const dup = parse(baseBar('C2:C6', 2, MINMAX + BLUE, D) + baseBar('D2:D6', 3, MINMAX + BLUE, D) +
      x14Block(companion('C2:C6', D, ' minLength="0" maxLength="100"', '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>')))
    expect(dup.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-invalid-x14-link' })
    const F = '{00000000-0000-4000-8000-0000000000FC}'
    const fillComp = parse(baseBar('E2:E6', 4, MINMAX + BLUE, F) +
      x14Block(companion('E2:E6', F, ' minLength="0" maxLength="100"', '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="FF0000FF"/>')))
    expect(fillComp.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-invalid-x14-link' })
    expect(fillComp.value.diagnostics.some(d => d.feature === 'cf-invalid-x14-link')).toBe(true)
    // S3-R: activePresent on a non-visual rule keeps main's cf-invalid-rule.
    expect(only(parse('<conditionalFormatting sqref="A1:A5"><cfRule type="top10" priority="1" rank="3" activePresent="1"/></conditionalFormatting>').value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
    // The diagnostics cap stays at 64.
    let many = ''
    for (let i = 1; i <= 100; i++) many += baseBar(`A${i}`, i, '<cfvo val="0"/><cfvo type="max"/>' + BLUE)
    expect(parse(many).value.diagnostics).toHaveLength(64)
  })

  it('context direction reads sheetView rightToLeft on both paths', async () => {
    const G = '{00000000-0000-4000-8000-0000000000C1}'
    const cf =
      baseBar('B2:B6', 1, MINMAX + BLUE) +
      `<conditionalFormatting sqref="C2:C6"><cfRule type="dataBar" priority="2"><dataBar>${MINMAX}${BLUE}</dataBar>` +
      `<extLst><ext uri="${DB_EXT}" xmlns:x14="${X14}"><x14:id>${G}</x14:id></ext></extLst></cfRule></conditionalFormatting>` +
      `<extLst><ext uri="${CF_EXT}" xmlns:x14="${X14}"><x14:conditionalFormattings xmlns:xm="${XM}">` +
      companion('C2:C6', G, ' minLength="0" maxLength="100" direction="leftToRight"', '<x14:cfvo type="min"/><x14:cfvo type="max"/>') +
      `</x14:conditionalFormattings></ext></extLst>`
    const view = (attrs: string) => `<sheetViews><sheetView${attrs} workbookViewId="0"/></sheetViews>`
    const bothPaths = async (sheetView: string, docId: string): Promise<[CfSheetRules, CfSheetRules]> => {
      const bytes = xlsxBytes({ beforeSheetData: sheetView, afterSheetData: cf })
      const resident = (await parseXlsx(await OfficePackage.load(bytes))).sheets[0]!.conditionalFormatting!
      const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId, profile: 'D' })
      try {
        await session.done
        return [resident, session.workbook.sheets[0]!.conditionalFormatting!]
      } finally { await session.release() }
    }
    injectExactCfEvidence(BAR_PARSE_CAPS)
    const rtl = await bothPaths(view(' rightToLeft="1"'), 'bar-rtl')
    for (const rules of rtl) {
      // S3-K: the viewer does not lay out RTL sheets, so a context bar there withholds whatever the evidence.
      expect(byPriority(rules.rules, 1).definition).toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
      expect(bar(byPriority(rules.rules, 2).definition)).toMatchObject({ direction: 'leftToRight', sheetRightToLeft: true })
    }
    expect(rtl[1].rules).toEqual(rtl[0].rules)

    const ltr = await bothPaths(view(''), 'bar-ltr')
    for (const rules of ltr) {
      expect(bar(byPriority(rules.rules, 1).definition)).toMatchObject({ direction: 'context', sheetRightToLeft: false })
      expect(bar(byPriority(rules.rules, 2).definition)).toMatchObject({ direction: 'leftToRight', sheetRightToLeft: false })
    }
    expect(ltr[1].rules).toEqual(ltr[0].rules)
    // Without its capability the LTR context bar withholds.
    injectExactCfEvidence(BAR_PARSE_CAPS.filter(c => c !== 'bar:direction-context-ltr-sheet'))
    for (const rules of await bothPaths(view(''), 'bar-ltr-nocap')) {
      expect(byPriority(rules.rules, 1).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    }
    // An unknown reading order (no sheet view, or a malformed flag) is never assumed LTR.
    injectExactCfEvidence(BAR_PARSE_CAPS)
    for (const [sheetView, id] of [['', 'bar-noview'], [view(' rightToLeft="maybe"'), 'bar-badview']] as const) {
      for (const rules of await bothPaths(sheetView, id)) {
        expect(byPriority(rules.rules, 1).definition, id).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
        expect(bar(byPriority(rules.rules, 2).definition).sheetRightToLeft, id).toBeUndefined()
      }
    }
  })

  it('bar colours resolve once on both paths', async () => {
    const expected = parseInt(resolveSheetColor({ theme: '4', tint: '-0.25' }, { theme: MEASURED_CF_OFFICE_THEME })!, 16)
    const xml = baseBar('B2:B6', 1, MINMAX + '<color theme="4" tint="-0.25"/>')
    const caps = ['bar:settings-schema-defaults', 'bar:pairing-base-only', 'bar:direction-context-ltr-sheet', 'bar:colors-theme-tint']
    await withExactCfEvidence(caps, async () => {
      const resident = parse(xml)
      expect(bar(only(resident.value.rules).definition).fillArgb).toBe(expected)
      const streamed = await stream(xml)
      expect(bar(only(streamed.value.rules).definition).fillArgb).toBe(expected)
      expect(streamed.value).toEqual(resident.value)
      streamed.release()
    })
    // The theme colour needs its own capability...
    expect(only(withExactCfEvidence(caps.filter(c => c !== 'bar:colors-theme-tint'), () => parse(xml)).value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    // ...and a palette other than the measured Office theme needs bar:colors-custom-theme, which no case can verify.
    const custom: CfStylesSource = { colors: { theme: { ...MEASURED_CF_OFFICE_THEME, accent1: '4472C4' } }, dxfNodes: [] }
    await withExactCfEvidence(caps, async () => {
      expect(only(parse(xml, LTR, custom).value.rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      const streamed = await stream(xml, LTR, custom)
      expect(only(streamed.value.rules).definition).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
      streamed.release()
    })
  })

  it('a bar definition is priced from the payload measurer', () => {
    const G = '{00000000-0000-4000-8000-0000000000BB}'
    const xml = baseBar('B2:B6', 1, MINMAX + BLUE, G) + x14Block(companion('B2:B6', G, ' border="1" negativeBarBorderColorSameAsPositive="0" direction="rightToLeft" axisPosition="automatic"',
      '<x14:cfvo type="percentile"><xm:f>10</xm:f></x14:cfvo><x14:cfvo type="percentile"><xm:f>90</xm:f></x14:cfvo><x14:borderColor rgb="FF000000"/>' +
      '<x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/>'))
    const def = bar(only(withExactCfEvidence(BAR_PARSE_CAPS, () => parse(xml)).value.rules).definition)
    // The bar's own fields, every optional one present, fit the derived constant (the origin is priced separately).
    const { dialect: _d, guid: _g, formulas: _f, ...own } = def
    expect(Object.keys(own)).toHaveLength(17)
    expect(measurePayloadBytes(own)).toBeLessThanOrEqual(CF_BAR_DEFINITION_BYTES)
  })

  it('barRequiredCapabilities names every parse gate and is never empty', () => {
    const def = bar(only(withExactCfEvidence(BAR_PARSE_CAPS, () => parse(baseBar('B2:B6', 1, MINMAX + BLUE))).value.rules).definition)
    expect(barRequiredCapabilities(def, [{ rgb: 'FF638EC6' }])).toEqual(new Set(['bar:pairing-base-only', 'bar:settings-schema-defaults', 'bar:direction-context-ltr-sheet']))
    expect(barRequiredCapabilities({ ...def, dialect: 'x14' }, [{ theme: '4', customTheme: true }, { indexed: '10' }]))
      .toEqual(new Set(['bar:pairing-standalone', 'bar:settings-schema-defaults', 'bar:direction-context-ltr-sheet', 'bar:colors-theme-tint', 'bar:colors-custom-theme', 'bar:colors-indexed']))
  })
})
