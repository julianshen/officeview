import { afterEach, describe, expect, it } from 'vitest'
import { getChildren, parseXml } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { CF_ICON_DEFINITION_BYTES, CF_LIMITS, createCfMemory } from '../src/xlsx/cf/limits'
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { injectExactCfEvidence, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { ICON_SETS, MEASURED_CF_OFFICE_THEME, iconRequiredCapabilities, parseCfSheet, type CfParseMeta, type CfStylesSource } from '../src/xlsx/cf/parse'
import { CF_ICON_SET_NAMES, ICON_SET_WRITES, type CfDefinition, type CfRule, type CfSheetRules } from '../src/xlsx/cf/model'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { xlsxBytes } from './helpers/xlsx-package'

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const DB_EXT = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
const LTR: CfParseMeta = { sheetRightToLeft: false }

/** Every parse-layer capability (the S4-DIALECT-EVIDENCE gates and the S4-B/C/F ones): with all of them injected only
 * structure decides. */
const ICON_PARSE_CAPS = [
  'icon:settings-schema-defaults', 'icon:dialect-base', 'icon:dialect-standalone',
  'icon:custom-mapping', 'icon:custom-noicons', 'icon:attr-percent-false', 'icon:thresholds-min-max',
]
const BASE_CAPS = ['icon:settings-schema-defaults', 'icon:dialect-base']

const styles: CfStylesSource = { colors: { theme: MEASURED_CF_OFFICE_THEME }, dxfNodes: [] }
const worksheet = (inner: string): string => `<worksheet xmlns="${NS}" xmlns:x14="${X14}" xmlns:xm="${XM}">${inner}</worksheet>`
function parse(inner: string, meta: CfParseMeta = LTR) {
  const memory = createCfMemory()
  const lease = parseCfSheet(parseXml(worksheet(inner)), styles, memory, meta)
  return { lease, value: lease.value, memory }
}
async function stream(inner: string, meta: CfParseMeta = LTR, split = 1e9) {
  const memory = createCfMemory(), handle = createCfStreamEvents(styles, memory, meta)
  const bytes = new TextEncoder().encode(worksheet(inner))
  await tokenizeBytes((async function* () { yield bytes.subarray(0, split); yield bytes.subarray(split) })(), handle.events)
  return handle.finish()
}
const only = (rules: readonly CfRule[]): CfRule => {
  expect(rules).toHaveLength(1)
  return rules[0]!
}
const icon = (definition: CfDefinition): Extract<CfDefinition, { kind: 'iconSet' }> => {
  expect(definition.kind).toBe('iconSet')
  return definition as Extract<CfDefinition, { kind: 'iconSet' }>
}

/** A base `<cfRule type="iconSet">`. */
const baseIcon = (sqref: string, priority: number, setAttrs: string, body: string, ruleAttrs = '') =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="iconSet" priority="${priority}"${ruleAttrs}><iconSet${setAttrs}>${body}</iconSet></cfRule></conditionalFormatting>`
/** Standalone priority-bearing x14 icon rules, in one `x14:conditionalFormattings` block. */
const x14Icons = (...rules: string[]) => `<extLst><ext uri="${CF_EXT}" xmlns:x14="${X14}"><x14:conditionalFormattings xmlns:xm="${XM}">${rules.join('')}</x14:conditionalFormattings></ext></extLst>`
const x14Icon = (sqref: string, priority: number, setAttrs: string, body: string, ruleAttrs = '') =>
  `<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="${priority}"${ruleAttrs}><x14:iconSet${setAttrs}>${body}</x14:iconSet></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`
const PCT3 = '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>'
const x14Num = (...vals: Array<string | [string, string]>) => vals.map(v => {
  const [val, extra] = typeof v === 'string' ? [v, ''] : v
  return `<x14:cfvo type="num"${extra}><xm:f>${val}</xm:f></x14:cfvo>`
}).join('')
const NUM3 = x14Num('0', '20', '40')
const cfIcons = (...pairs: Array<[string, string]>) => pairs.map(([set, id]) => `<x14:cfIcon iconSet="${set}" iconId="${id}"/>`).join('')
const CUSTOM = ' iconSet="3Arrows" custom="1"'
const customX14 = (icons: string, setAttrs = CUSTOM, cfvo = NUM3) => x14Icons(x14Icon('B2:B6', 3, setAttrs, cfvo + icons))
const GOOD_ICONS = cfIcons(['3Flags', '0'], ['NoIcons', '0'], ['3Stars', '2'])

afterEach(() => { injectExactCfEvidence(null) })

const reasonOf = (xml: string, caps: readonly string[] = ICON_PARSE_CAPS, meta: CfParseMeta = LTR) => {
  const rule = only(withExactCfEvidence(caps, () => parse(xml, meta)).value.rules)
  return rule
}
const unresolved = (reason: string) => ({ kind: 'unresolved', type: 'iconSet', reason })

describe('CF Stage 4 icon-set model and parsing (Task 25)', () => {
  it('the frozen set-name table lines up with the family sizes and carries NoIcons last', () => {
    expect(CF_ICON_SET_NAMES).toHaveLength(21)
    expect(Object.isFrozen(CF_ICON_SET_NAMES)).toBe(true)
    expect(CF_ICON_SET_NAMES.at(-1)).toBe('NoIcons')
    expect(new Set(CF_ICON_SET_NAMES.slice(0, 20))).toEqual(new Set(ICON_SETS.keys()))
    expect(new Set(CF_ICON_SET_NAMES).size).toBe(21)
  })

  it('canonical icon write mask is [\'iconSet\'] for resolved and unresolved icon sets', () => {
    expect(ICON_SET_WRITES).toEqual(new Set(['iconSet']))
    const xml = baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3)
    const resolved = only(withExactCfEvidence(BASE_CAPS, () => parse(xml)).value.rules)
    expect(resolved.definition.kind).toBe('iconSet')
    expect(new Set(resolved.writes)).toEqual(ICON_SET_WRITES)

    const unresolvedRule = only(withExactCfEvidence([], () => parse(xml)).value.rules)
    expect(unresolvedRule.definition).toMatchObject(unresolved('cf-unsupported-rule'))
    expect(new Set(unresolvedRule.writes)).toEqual(ICON_SET_WRITES)

    // A set carrying a DXF adds exactly the DXF's own writes, resolved or not.
    const dxfRoot = parseXml(`<styleSheet xmlns="${NS}"><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`)
    const dxfStyles: CfStylesSource = { colors: { theme: MEASURED_CF_OFFICE_THEME }, dxfNodes: getChildren(getChildren(dxfRoot, 'dxfs')[0]!, 'dxf') }
    const withDxf = baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3, ' dxfId="0"')
    for (const caps of [BASE_CAPS, []]) {
      const lease = withExactCfEvidence(caps, () => parseCfSheet(parseXml(worksheet(withDxf)), dxfStyles, createCfMemory(), LTR))
      const rule = only(lease.value.rules)
      expect(rule.definition.kind === 'iconSet', String(caps)).toBe(caps.length > 0)
      expect(rule.writes.has('iconSet')).toBe(true)
      // Exactly the DXF's own writes (a solid fill writes the `fill` slot), never more and never ALL.
      expect(new Set(rule.writes)).toEqual(new Set(['iconSet', 'fill']))
      expect(rule.writes.has('ALL')).toBe(false)
    }

    // A broken link keeps main's invalidLinkRule mask: ALL.
    const broken = baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3).replace('</cfRule>',
      `<extLst><ext uri="${DB_EXT}"><x14:id>{NO-COMPANION}</x14:id></ext></extLst></cfRule>`)
    const b = only(withExactCfEvidence(ICON_PARSE_CAPS, () => parse(broken)).value.rules)
    expect(b.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-x14-link' })
    expect(new Set(b.writes)).toEqual(new Set(['ALL']))
  })

  it('base 3Arrows resolves with schema defaults', async () => {
    const xml = baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3)
    await withExactCfEvidence(BASE_CAPS, async () => {
      const resident = parse(xml)
      const def = icon(only(resident.value.rules).definition)
      expect(def).toMatchObject({
        kind: 'iconSet', dialect: 'base', set: '3Arrows', reverse: false, showValue: true, percentAttr: true, sheetRightToLeft: false,
        thresholds: [{ bound: { type: 'percent', value: 0 }, gte: true }, { bound: { type: 'percent', value: 33 }, gte: true }, { bound: { type: 'percent', value: 67 }, gte: true }],
      })
      expect('custom' in def).toBe(false)
      const streamed = await stream(xml)
      expect(streamed.value).toEqual(resident.value)
      streamed.release()
    })
    // Each parse-layer gate is required, and without evidence the rule withholds with one deduplicated diagnostic.
    for (const missing of BASE_CAPS) {
      const partial = withExactCfEvidence(BASE_CAPS.filter(c => c !== missing), () => parse(xml))
      expect(only(partial.value.rules).definition, missing).toMatchObject(unresolved('cf-unsupported-rule'))
    }
    const none = withExactCfEvidence([], () => parse(xml))
    const rule = only(none.value.rules)
    expect(rule.definition).toMatchObject(unresolved('cf-unsupported-rule'))
    expect(new Set(rule.writes)).toEqual(ICON_SET_WRITES)
    expect(none.value.diagnostics.filter(d => d.feature === 'cf-unsupported-rule')).toHaveLength(1)
    // An omitted iconSet attribute is 3TrafficLights1 (spec line 84).
    const defaulted = baseIcon('B2:B6', 1, '', PCT3)
    expect(icon(only(withExactCfEvidence(BASE_CAPS, () => parse(defaulted)).value.rules).definition)).toMatchObject({ set: '3TrafficLights1', reverse: false, showValue: true, percentAttr: true })
    // Explicit attributes read back, in every boolean spelling.
    const explicit = baseIcon('B2:B6', 1, ' iconSet="4Arrows" reverse="true" showValue="0" percent="false"', '<cfvo type="num" val="0"/><cfvo type="num" val="10"/><cfvo type="num" val="20"/><cfvo type="num" val="30" gte="false"/>')
    expect(icon(only(withExactCfEvidence([...BASE_CAPS, 'icon:attr-percent-false'], () => parse(explicit)).value.rules).definition))
      .toMatchObject({ set: '4Arrows', reverse: true, showValue: false, percentAttr: false, thresholds: [{ gte: true }, { gte: true }, { gte: true }, { bound: { type: 'num', value: 30 }, gte: false }] })
  })

  it('standalone x14 custom set with NoIcons resolves', async () => {
    const xml = customX14(GOOD_ICONS, CUSTOM, x14Num('0', ['20', ' gte="0"'], '40'))
    await withExactCfEvidence(ICON_PARSE_CAPS, async () => {
      const resident = parse(xml)
      const rule = only(resident.value.rules)
      expect(rule.priority).toBe(3)
      const def = icon(rule.definition)
      expect(def).toMatchObject({ dialect: 'x14', set: '3Arrows', reverse: false, showValue: true, percentAttr: true })
      expect(def.custom).toEqual([{ set: '3Flags', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '3Stars', iconId: 2 }])
      expect(def.thresholds.map(t => t.gte)).toEqual([true, false, true])
      expect(def.thresholds.map(t => t.bound)).toEqual([{ type: 'num', value: 0 }, { type: 'num', value: 20 }, { type: 'num', value: 40 }])
      expect(new Set(rule.writes)).toEqual(ICON_SET_WRITES)
      const streamed = await stream(xml)
      expect(streamed.value).toEqual(resident.value)
      streamed.release()
    })
    // The dialect, the mapping and the NoIcons bucket each need their own evidence.
    for (const missing of ['icon:dialect-standalone', 'icon:custom-mapping', 'icon:custom-noicons']) {
      const partial = withExactCfEvidence(ICON_PARSE_CAPS.filter(c => c !== missing), () => parse(xml))
      expect(only(partial.value.rules).definition, missing).toMatchObject(unresolved('cf-unsupported-rule'))
    }
    // Without a NoIcons bucket the NoIcons capability is not needed.
    const plain = customX14(cfIcons(['3Flags', '0'], ['3Arrows', '1'], ['3Stars', '2']))
    const caps = ICON_PARSE_CAPS.filter(c => c !== 'icon:custom-noicons')
    expect(icon(reasonOf(plain, caps).definition).custom).toHaveLength(3)
    // A standalone x14 rule without a custom flag is a plain built-in set.
    const builtin = x14Icons(x14Icon('C2:C6', 4, ' iconSet="3Stars"', NUM3))
    const builtinRule = reasonOf(builtin, ICON_PARSE_CAPS)
    expect(icon(builtinRule.definition)).toMatchObject({ dialect: 'x14', set: '3Stars' })
  })

  it('unknown icon mappings refuse the whole contribution (S4-B)', async () => {
    const cases: Array<[string, string]> = [
      ['unknown family', customX14('', ' iconSet="7Moons"')],
      ['unknown base family', baseIcon('B2:B6', 1, ' iconSet="7Moons"', PCT3)],
      ['cfIcon iconId out of range', customX14(cfIcons(['3Arrows', '3'], ['3Arrows', '1'], ['3Arrows', '2']))],
      ['cfIcon unknown family', customX14(cfIcons(['7Moons', '0'], ['3Arrows', '1'], ['3Arrows', '2']))],
      ['NoIcons with another id', customX14(cfIcons(['NoIcons', '1'], ['3Arrows', '1'], ['3Arrows', '2']))],
      ['non-integer iconId', customX14(cfIcons(['3Arrows', '1.5'], ['3Arrows', '1'], ['3Arrows', '2']))],
      ['negative iconId', customX14(cfIcons(['3Arrows', '-1'], ['3Arrows', '1'], ['3Arrows', '2']))],
      ['custom with 2 cfIcons on a 3-icon set', customX14(cfIcons(['3Arrows', '0'], ['3Arrows', '1']))],
      ['custom with 4 cfIcons on a 3-icon set', customX14(cfIcons(['3Arrows', '0'], ['3Arrows', '1'], ['3Arrows', '2'], ['3Arrows', '2']))],
      ['cfIcons without the custom flag', customX14(GOOD_ICONS, ' iconSet="3Arrows"')],
    ]
    for (const [label, xml] of cases) {
      const rule = reasonOf(xml)
      expect(rule.definition, label).toMatchObject(unresolved('cf-unsupported-icon'))
      expect(new Set(rule.writes), label).toEqual(ICON_SET_WRITES)
      const streamed = await withExactCfEvidence(ICON_PARSE_CAPS, async () => {
        const result = await stream(xml)
        const value = result.value
        result.release()
        return value
      })
      expect(streamed.rules, label).toEqual(withExactCfEvidence(ICON_PARSE_CAPS, () => parse(xml)).value.rules)
    }
    // Every legal mapping still resolves: each family at each id, and NoIcons only at id 0.
    for (const [name, size] of ICON_SETS) {
      const icons = cfIcons([name, '0'], ['NoIcons', '0'], [name, String(size - 1)])
      expect(icon(reasonOf(customX14(icons)).definition).custom, name).toEqual([{ set: name, iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: name, iconId: size - 1 }])
    }
    // The diagnostics cap stays at 64.
    let many = ''
    for (let i = 1; i <= 100; i++) many += baseIcon(`A${i}`, i, ' iconSet="7Moons"', PCT3)
    expect(withExactCfEvidence(ICON_PARSE_CAPS, () => parse(many)).value.diagnostics).toHaveLength(64)
  })

  it('icon cfvo structure refusals', async () => {
    const cfvo2 = '<cfvo type="percent" val="0"/><cfvo type="percent" val="50"/>'
    const cfvo6 = '<cfvo type="num" val="0"/>'.repeat(6)
    const missingType = '<cfvo val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>'
    const withCfvo = (...cfvos: string[]) => baseIcon('B2:B6', 1, ' iconSet="3Arrows"', cfvos.join(''))
    const rows: Array<[string, string, string] | [string, string, string, readonly string[]]> = [
      ['2 cfvo on a 3-icon set (S4-C)', baseIcon('B2:B6', 1, ' iconSet="3Arrows"', cfvo2), 'cf-invalid-rule'],
      ['4 cfvo on a 3-icon set', baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3 + '<cfvo type="num" val="90"/>'), 'cf-invalid-rule'],
      ['no cfvo', baseIcon('B2:B6', 1, ' iconSet="3Arrows"', ''), 'cf-invalid-rule'],
      ['missing type', baseIcon('B2:B6', 1, ' iconSet="3Arrows"', missingType), 'cf-invalid-cfvo'],
      ['unknown cfvo type', withCfvo('<cfvo type="autoMin"/>', '<cfvo type="percent" val="33"/>', '<cfvo type="percent" val="67"/>'), 'cf-invalid-cfvo'],
      ['numeric cfvo without a value', withCfvo('<cfvo type="num"/>', '<cfvo type="percent" val="33"/>', '<cfvo type="percent" val="67"/>'), 'cf-invalid-cfvo'],
      ['non-numeric cfvo value', withCfvo('<cfvo type="num" val="abc"/>', '<cfvo type="percent" val="33"/>', '<cfvo type="percent" val="67"/>'), 'cf-invalid-cfvo'],
      ['bad gte', withCfvo('<cfvo type="percent" val="0" gte="maybe"/>', '<cfvo type="percent" val="33"/>', '<cfvo type="percent" val="67"/>'), 'cf-invalid-cfvo'],
      ['formula cfvo (Stage 5)', withCfvo('<cfvo type="percent" val="0"/>', '<cfvo type="formula" val="$A$1"/>', '<cfvo type="percent" val="67"/>'), 'cf-unsupported-rule'],
      ['reference value (Stage 5)', withCfvo('<cfvo type="percent" val="0"/>', '<cfvo type="num" val="$A$1"/>', '<cfvo type="percent" val="67"/>'), 'cf-unsupported-rule'],
      ['x14 xm:f reference (Stage 5)', x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Arrows"', x14Num('0', '$A$1', '40'))), 'cf-unsupported-rule'],
      ['x14 formula cfvo (Stage 5)', x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Arrows"', '<x14:cfvo type="formula"><xm:f>A1</xm:f></x14:cfvo>' + x14Num('20', '40'))), 'cf-unsupported-rule'],
      ['x14 cfvo with @val', x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Arrows"', '<x14:cfvo type="num" val="0"><xm:f>0</xm:f></x14:cfvo>' + x14Num('20', '40'))), 'cf-invalid-cfvo'],
      ['x14 custom with 2 cfvo', customX14(GOOD_ICONS, CUSTOM, x14Num('0', '20')), 'cf-invalid-rule'],
      ['activePresent (S4-P)', baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3, ' activePresent="1"'), 'cf-unsupported-rule'],
      ['x14 activePresent (S4-P)', x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Stars"', NUM3, ' activePresent="1"')), 'cf-unsupported-rule'],
      ['rule without its iconSet element', '<conditionalFormatting sqref="B2:B6"><cfRule type="iconSet" priority="1"/></conditionalFormatting>', 'cf-invalid-rule'],
      ['malformed reverse', baseIcon('B2:B6', 1, ' iconSet="3Arrows" reverse="maybe"', PCT3), 'cf-invalid-rule'],
      ['malformed showValue', baseIcon('B2:B6', 1, ' iconSet="3Arrows" showValue="2"', PCT3), 'cf-invalid-rule'],
      ['malformed percent', baseIcon('B2:B6', 1, ' iconSet="3Arrows" percent="x"', PCT3), 'cf-invalid-rule'],
      ['malformed custom', x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Arrows" custom="yes"', NUM3 + GOOD_ICONS)), 'cf-invalid-rule'],
      // Evidence-gated shapes: unresolved (cf-unsupported-rule) unless their capability is present.
      ['min/max cfvo without icon:thresholds-min-max', withCfvo('<cfvo type="min"/>', '<cfvo type="percent" val="50"/>', '<cfvo type="max"/>'), 'cf-unsupported-rule', ICON_PARSE_CAPS.filter(c => c !== 'icon:thresholds-min-max')],
      // Record-only (S4-DIALECT-CASE): `icon:dialect-base-x14-family` records Excel's outcome (it may be no icon or another
      // family), so even with it injected an x14-only family on a base rule stays withheld.
      ['base 3Stars', baseIcon('B2:B6', 1, ' iconSet="3Stars"', PCT3), 'cf-unsupported-rule'],
      ['base 3Triangles with every id incl. the recorded dialect', baseIcon('B2:B6', 1, ' iconSet="3Triangles"', PCT3), 'cf-unsupported-rule', [...ICON_PARSE_CAPS, 'icon:dialect-base-x14-family', 'icon:dialect-paired']],
      ['base 5Boxes', baseIcon('B2:B6', 1, ' iconSet="5Boxes"', '<cfvo type="num" val="0"/>'.repeat(5)), 'cf-unsupported-rule', [...ICON_PARSE_CAPS, 'icon:dialect-base-x14-family']],
      ['<iconSet percent="0"> (S4-F)', baseIcon('B2:B6', 1, ' iconSet="3Arrows" percent="0"', PCT3), 'cf-unsupported-rule', ICON_PARSE_CAPS.filter(c => c !== 'icon:attr-percent-false')],
      ['base custom set (not an Excel form)', baseIcon('B2:B6', 1, ' iconSet="3Arrows" custom="1"', PCT3 + '<cfIcon iconSet="3Flags" iconId="0"/><cfIcon iconSet="3Arrows" iconId="1"/><cfIcon iconSet="3Stars" iconId="2"/>'), 'cf-unsupported-rule'],
    ]
    for (const [label, xml, reason, caps] of rows) {
      const rule = reasonOf(xml, caps ?? ICON_PARSE_CAPS)
      expect(rule.definition, label).toMatchObject(unresolved(reason))
      expect(new Set(rule.writes), label).toEqual(ICON_SET_WRITES)
      const resident = withExactCfEvidence(caps ?? ICON_PARSE_CAPS, () => parse(xml)).value
      const streamed = await withExactCfEvidence(caps ?? ICON_PARSE_CAPS, async () => { const r = await stream(xml); const v = r.value; r.release(); return v })
      expect(streamed, label).toEqual(resident)
    }
    // min/max resolve with their capability (an unresolved-without row above is the gate, this is the pass).
    const minMax = icon(reasonOf(withCfvo('<cfvo type="min"/>', '<cfvo type="percent" val="50"/>', '<cfvo type="max" val="5"/>')).definition)
    expect(minMax.thresholds.map(t => t.bound)).toEqual([{ type: 'min' }, { type: 'percent', value: 50 }, { type: 'max' }])
    // The x14-only families resolve on their own (standalone x14) dialect.
    for (const set of ['3Triangles', '3Stars']) {
      expect(icon(reasonOf(x14Icons(x14Icon('B2:B6', 1, ` iconSet="${set}"`, NUM3))).definition).set).toBe(set)
    }
    // 6 cfvo is over the cap: refused before any definition is built, on both paths.
    for (const xml of [baseIcon('B2:B6', 1, ' iconSet="3Arrows"', cfvo6), x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Arrows"', x14Num('0', '1', '2', '3', '4', '5')))]) {
      const rule = reasonOf(xml)
      expect(rule.definition).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
      const streamed = await withExactCfEvidence(ICON_PARSE_CAPS, async () => { const r = await stream(xml); const v = r.value; r.release(); return v })
      expect(streamed.rules).toEqual(withExactCfEvidence(ICON_PARSE_CAPS, () => parse(xml)).value.rules)
    }
    expect(CF_LIMITS.cfvo).toBe(5)
    // S4-P: activePresent on a non-visual rule keeps main's cf-invalid-rule.
    expect(only(parse('<conditionalFormatting sqref="A1:A5"><cfRule type="top10" priority="1" rank="3" activePresent="1"/></conditionalFormatting>').value.rules).definition)
      .toMatchObject({ kind: 'unresolved', reason: 'cf-invalid-rule' })
  })

  it('icon sets on an RTL sheet are unresolved on both paths (S4-M)', async () => {
    const cf = baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3) + x14Icons(x14Icon('C2:C6', 2, ' iconSet="3Stars"', NUM3))
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
    injectExactCfEvidence(ICON_PARSE_CAPS)
    // Whatever the evidence, an RTL sheet withholds every icon: the viewer does not lay out RTL sheets.
    for (const [sheetView, id] of [[view(' rightToLeft="1"'), 'icon-rtl'], [view(' rightToLeft="maybe"'), 'icon-badview']] as const) {
      const both = await bothPaths(sheetView, id)
      for (const rules of both) {
        expect(rules.rules, id).toHaveLength(2)
        for (const rule of rules.rules) {
          expect(rule.definition, id).toMatchObject(unresolved('cf-unsupported-rule'))
          expect(new Set(rule.writes), id).toEqual(ICON_SET_WRITES)
        }
      }
      expect(both[1].rules, id).toEqual(both[0].rules)
    }
    // An LTR sheet (no view, an empty views element, or an explicit 0) resolves under evidence.
    for (const [sheetView, id] of [['', 'icon-noview'], ['<sheetViews/>', 'icon-emptyviews'], [view(' rightToLeft="0"'), 'icon-ltr0']] as const) {
      const both = await bothPaths(sheetView, id)
      for (const rules of both) {
        expect(icon(rules.rules.find(r => r.priority === 1)!.definition), id).toMatchObject({ set: '3Arrows', sheetRightToLeft: false })
        expect(icon(rules.rules.find(r => r.priority === 2)!.definition), id).toMatchObject({ set: '3Stars', sheetRightToLeft: false })
      }
      expect(both[1].rules, id).toEqual(both[0].rules)
    }
    // A caller that cannot tell the reading order never gets an icon.
    expect(reasonOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3), ICON_PARSE_CAPS, {}).definition).toMatchObject(unresolved('cf-unsupported-rule'))
    expect(reasonOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3), ICON_PARSE_CAPS, { sheetRightToLeft: true }).definition).toMatchObject(unresolved('cf-unsupported-rule'))
  })

  it('iconRequiredCapabilities is never empty and names each feature', () => {
    const defOf = (xml: string, caps: readonly string[] = ICON_PARSE_CAPS) => icon(reasonOf(xml, caps).definition)
    const base = defOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows"', PCT3))
    expect(iconRequiredCapabilities(base)).toEqual(new Set(['icon:dialect-base', 'icon:settings-schema-defaults']))
    // Even a fully explicit set needs its dialect; the schema defaults stay required while `percent` is true (S4-F).
    expect(iconRequiredCapabilities(base, { defaultedAttributes: false })).toEqual(new Set(['icon:dialect-base', 'icon:settings-schema-defaults']))
    const explicit = defOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows" percent="0" reverse="0" showValue="1"', PCT3))
    expect(iconRequiredCapabilities(explicit, { defaultedAttributes: false })).toEqual(new Set(['icon:dialect-base', 'icon:attr-percent-false']))
    expect(iconRequiredCapabilities(explicit)).toEqual(new Set(['icon:dialect-base', 'icon:attr-percent-false', 'icon:settings-schema-defaults']))

    const standalone = defOf(x14Icons(x14Icon('B2:B6', 1, ' iconSet="3Stars"', NUM3)))
    expect(iconRequiredCapabilities(standalone).has('icon:dialect-standalone')).toBe(true)
    expect(iconRequiredCapabilities(standalone).has('icon:dialect-base')).toBe(false)
    expect(iconRequiredCapabilities(standalone).has('icon:custom-mapping')).toBe(false)

    const custom = defOf(customX14(cfIcons(['3Flags', '0'], ['3Arrows', '1'], ['3Stars', '2'])))
    expect(iconRequiredCapabilities(custom).has('icon:custom-mapping')).toBe(true)
    expect(iconRequiredCapabilities(custom).has('icon:custom-noicons')).toBe(false)
    const noIcons = defOf(customX14(GOOD_ICONS))
    expect(iconRequiredCapabilities(noIcons).has('icon:custom-mapping')).toBe(true)
    expect(iconRequiredCapabilities(noIcons).has('icon:custom-noicons')).toBe(true)

    // No record-only id is ever a parse gate (a base x14-only family is withheld outright, see the structure refusals).
    for (const def of [base, standalone, custom]) {
      expect(iconRequiredCapabilities(def).has('icon:dialect-base-x14-family')).toBe(false)
      expect(iconRequiredCapabilities(def).has('icon:dialect-paired')).toBe(false)
    }

    expect(iconRequiredCapabilities(defOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows" percent="0"', PCT3))).has('icon:attr-percent-false')).toBe(true)
    expect(iconRequiredCapabilities(base).has('icon:attr-percent-false')).toBe(false)
    const minMax = defOf(baseIcon('B2:B6', 1, ' iconSet="3Arrows"', '<cfvo type="min"/><cfvo type="percent" val="50"/><cfvo type="max"/>'))
    expect(iconRequiredCapabilities(minMax).has('icon:thresholds-min-max')).toBe(true)

    // Never empty, and no evaluation-layer id (thresholds, families, reverse) is a parse gate (S4-D, S4-E).
    for (const def of [base, explicit, standalone, custom, noIcons, minMax, defOf(baseIcon('B2:B6', 1, ' iconSet="4Arrows" reverse="1"', '<cfvo type="num" val="0" gte="0"/>'.repeat(4)))]) {
      const required = iconRequiredCapabilities(def)
      expect(required.size).toBeGreaterThan(0)
      for (const cap of required) expect(cap, cap).not.toMatch(/^icon:(thresholds-(?!min-max$)|families-|reverse|bounds-|degenerate-|participation-|scope-|composition-)/)
    }
  })

  it('icon definition bytes are measurer-derived (Lesson 6)', () => {
    // The largest definition: 5 thresholds of the longest bound type, 5 custom refs and the longest set names.
    const longest = '3TrafficLights1'
    expect(CF_ICON_SET_NAMES.every(n => n.length <= longest.length)).toBe(true)
    const largest = {
      kind: 'iconSet', set: longest, thresholds: Array.from({ length: 5 }, () => ({ bound: { type: 'percentile' as const, value: 0 }, gte: false })),
      reverse: false, showValue: false, percentAttr: false,
      custom: Array.from({ length: 5 }, () => ({ set: longest, iconId: 0 })), sheetRightToLeft: false,
    }
    expect(CF_ICON_DEFINITION_BYTES).toBe(measurePayloadBytes(largest))
    // A real five-icon custom definition parsed from XML fits.
    const xml = x14Icons(x14Icon('B2:B6', 1, ' iconSet="5Boxes" custom="1" reverse="1" showValue="0" percent="0"', x14Num('0', '1', '2', '3', '4') +
      cfIcons(['5Boxes', '4'], ['NoIcons', '0'], ['5Rating', '3'], ['5Arrows', '2'], ['5Quarters', '4'])))
    const def = icon(reasonOf(xml).definition)
    const { dialect: _d, guid: _g, formulas: _f, ...own } = def
    expect(def.thresholds).toHaveLength(5)
    expect(measurePayloadBytes(own)).toBeLessThanOrEqual(CF_ICON_DEFINITION_BYTES)
    expect(CF_ICON_DEFINITION_BYTES).toBeGreaterThan(0)
  })
})

describe('Stage 5 formula thresholds and activation (Task 37)', () => {
  const FX = ['formula:dialect-base', 'formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs', 'formula:cfvo-reference', 'icon:thresholds-formula',
    'formula:active-present-icon']
  const set = (cfvos: string) => `<conditionalFormatting sqref="B2:B5"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows">${cfvos}</iconSet></cfRule></conditionalFormatting>`
  const cfvos = '<cfvo type="percent" val="0"/><cfvo type="formula" val="$F$1"/><cfvo type="num" val="$F$2"/>'
  it('a formula and a reference num threshold compile under their capabilities, gte kept', async () => {
    const def = withExactCfEvidence([...ICON_PARSE_CAPS, ...FX], () => only(parse(set(cfvos)).value.rules).definition)
    expect(def.kind).toBe('iconSet')
    const t = (def as Extract<CfDefinition, { kind: 'iconSet' }>).thresholds
    expect(t.map(x => x.bound.type === 'formula' ? `formula:${x.bound.cfvo}` : x.bound.type)).toEqual(['percent', 'formula:formula', 'formula:num'])
    expect(t.every(x => x.gte)).toBe(true)
    const streamed = await withExactCfEvidence([...ICON_PARSE_CAPS, ...FX], () => stream(set(cfvos)))
    expect(streamed.value.rules.map(r => r.definition)).toEqual([def])
  })
  it('without the family or the dialect capability, main\'s cf-unsupported-rule', () => {
    for (const missing of ['icon:thresholds-formula', 'formula:dialect-base', 'formula:cfvo-reference']) {
      const def = withExactCfEvidence([...ICON_PARSE_CAPS, ...FX].filter(c => c !== missing), () => only(parse(set(cfvos)).value.rules).definition)
      expect(def, missing).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
    }
  })
  it('a standalone activePresent set carries its activation; a base one stays refused', () => {
    const x14 = `<extLst><ext uri="${CF_EXT}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1" activePresent="1" id="{00000000-0000-4000-8000-0000000000C1}"><xm:f>$A2&gt;0</xm:f><x14:iconSet iconSet="3Arrows"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    const def = withExactCfEvidence([...ICON_PARSE_CAPS, ...FX], () => only(parse(x14).value.rules).definition)
    expect(def).toMatchObject({ kind: 'iconSet', activePresent: true, activation: { pinned: true } })
    const base = `<conditionalFormatting sqref="B2:B5"><cfRule type="iconSet" priority="1" activePresent="1"><formula>$A2&gt;0</formula><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>`
    expect(withExactCfEvidence([...ICON_PARSE_CAPS, ...FX], () => only(parse(base).value.rules).definition)).toMatchObject({ kind: 'unresolved', reason: 'cf-unsupported-rule' })
  })
})
