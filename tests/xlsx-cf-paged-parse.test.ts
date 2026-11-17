import { describe, expect, it, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { createCfMemory } from '../src/xlsx/cf/limits'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { parseCfSheet, type CfParseMeta } from '../src/xlsx/cf/parse'
import { withExactCfEvidence, withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { xlsxBytes, xlsxEntries, SS_NS } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'
import * as formatting from '../src/xlsx/storage/pagedFormatting'
import * as capturing from '../src/xlsx/cf/capture'
const styles = {
  colors: {
    theme: {
      accent1: '4F81BD',
      accent2: 'C0504D',
    },
  },
  dxfNodes: [],
}
const base = '<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>'
const wrap = (s: string) => `<worksheet xmlns="${SS_NS}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${s}</worksheet>`
async function stream(xml: string, split = 1e9, meta: CfParseMeta = {}) {
  const memory = createCfMemory(), handle = createCfStreamEvents(styles, memory, meta)
  const bytes = new TextEncoder().encode(xml)
  await tokenizeBytes((async function* () { yield bytes.subarray(0, split); yield bytes.subarray(split) })(), handle.events)
  const lease = handle.finish()
  return { lease, memory, handle }
}
describe('paged CF capture', () => {
  it('resident and streamed CF are deep equal at every chunk boundary', async () => {
    const twoColorScale = '<conditionalFormatting sqref="D2:D10"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const threeColorScale = '<conditionalFormatting sqref="E2:E10"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>'
    const themeScale = '<conditionalFormatting sqref="F2:F10"><cfRule type="colorScale" priority="5"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4" tint="-0.4"/><color theme="5" tint="0.4"/></colorScale></cfRule></conditionalFormatting>'
    const x14Scale = '<x14:conditionalFormatting><x14:cfRule type="colorScale" priority="6" id="{01234567-89AB-CDEF-0123-456789ABCDE0}"><x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:color rgb="FFFF0000"/><x14:color rgb="FF00FF00"/></x14:colorScale></x14:cfRule><xm:sqref>G1:G10</xm:sqref></x14:conditionalFormatting>'
    // Data bars: a paired bar whose companion arrives after every base rule (with x14 attributes, colours and an
    // xm:f bound that must survive capture), a base-only bar, an autoMin-on-base refusal and a standalone x14 bar.
    const pairId = '{3B3B0019-0000-4000-8000-000000000001}'
    const pairedBase = `<conditionalFormatting sqref="H2:H6"><cfRule type="dataBar" priority="7"><dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>${pairId}</x14:id></ext></extLst></cfRule></conditionalFormatting>`
    const baseOnlyBar = '<conditionalFormatting sqref="I2:I6"><cfRule type="dataBar" priority="8"><dataBar><cfvo type="percentile" val="10"/><cfvo type="num" val="90"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting>'
    const autoMinBase = '<conditionalFormatting sqref="J2:J6"><cfRule type="dataBar" priority="9"><dataBar><cfvo type="autoMin"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting>'
    const companion = `<x14:conditionalFormatting><x14:cfRule type="dataBar" id="${pairId}"><x14:dataBar minLength="0" maxLength="100" gradient="0" border="1" direction="leftToRight" axisPosition="middle" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF00B050"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFFA500"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>H2:H6</xm:sqref></x14:conditionalFormatting>`
    const standaloneBar = '<x14:conditionalFormatting><x14:cfRule type="dataBar" priority="10" id="{3B3B0019-0000-4000-8000-000000000002}"><x14:dataBar minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="1" axisPosition="none"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="FF00B050"/></x14:dataBar></x14:cfRule><xm:sqref>K2:K6</xm:sqref></x14:conditionalFormatting>'
    // Icon sets: a base 3Arrows set, a base x14-only family (3Stars) and a standalone x14 custom set carrying a NoIcons
    // bucket, a `gte` flag and an `xm:f` threshold. Splits fall inside every <x14:cfIcon>, <x14:cfvo> and <xm:f> too.
    const baseIconSet = '<conditionalFormatting sqref="L2:L6"><cfRule type="iconSet" priority="11"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>'
    const baseStars = '<conditionalFormatting sqref="M2:M6"><cfRule type="iconSet" priority="12"><iconSet iconSet="3Stars" reverse="1" showValue="0"><cfvo type="num" val="1"/><cfvo type="num" val="2" gte="0"/><cfvo type="num" val="3"/></iconSet></cfRule></conditionalFormatting>'
    const customIconSet = '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="13" id="{3B3B0019-0000-4000-8000-000000000003}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num" gte="0"><xm:f>20</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>40</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>N2:N6</xm:sqref></x14:conditionalFormatting>'
    const standaloneStars = '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="14" id="{3B3B0019-0000-4000-8000-000000000004}"><x14:iconSet iconSet="3Stars" percent="0"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>O2:O6</xm:sqref></x14:conditionalFormatting>'
    // Stage 5 (Task 33): one rule Excel's own `today` formula, one divergent formula and one absent formula, split everywhere.
    const timePeriodCanonical = '<conditionalFormatting sqref="P2:P6"><cfRule type="timePeriod" priority="15" timePeriod="today"><formula>FLOOR(P2,1)=TODAY()</formula></cfRule></conditionalFormatting>'
    const timePeriodDivergent = '<conditionalFormatting sqref="Q2:Q6"><cfRule type="timePeriod" priority="16" timePeriod="today"><formula>Q2=TODAY()-3</formula></cfRule></conditionalFormatting>'
    const timePeriodAbsent = '<conditionalFormatting sqref="R2:R6"><cfRule type="timePeriod" priority="17" timePeriod="lastMonth"/></conditionalFormatting>'
    const xml = wrap(
      base +
        timePeriodCanonical +
        timePeriodDivergent +
        timePeriodAbsent +
        baseIconSet +
        baseStars +
        twoColorScale +
        threeColorScale +
        themeScale +
        pairedBase +
        baseOnlyBar +
        autoMinBase +
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="2" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>A1=&quot;茶&quot;</xm:f></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting>' +
        x14Scale +
        companion +
        standaloneBar +
        customIconSet +
        standaloneStars +
        '</x14:conditionalFormattings></ext></extLst>',
    )
    // Additive evidence: the scales keep production evidence, and the bars resolve, so a streamed bar that lost an
    // x14 attribute, colour or bound would differ from the resident one instead of both withholding alike.
    const meta: CfParseMeta = { sheetRightToLeft: false }
    await withInjectedCfEvidence(['bar:settings-schema-defaults', 'bar:settings-paired-precedence', 'bar:pairing-paired', 'bar:pairing-standalone',
      'bar:pairing-base-only', 'bar:direction-ltr', 'bar:direction-context-ltr-sheet',
      'icon:settings-schema-defaults', 'icon:dialect-base', 'icon:dialect-standalone', 'icon:dialect-base-x14-family', 'icon:dialect-paired', 'icon:custom-mapping',
      'icon:custom-noicons', 'icon:attr-percent-false'], async () => {
      const expected = parseCfSheet(parseXml(xml), styles, createCfMemory(), meta).value
      const barOf = (priority: number) => expected.rules.find(r => r.priority === priority)!.definition
      expect(barOf(7)).toMatchObject({ kind: 'dataBar', bounds: [{ type: 'num', value: 20 }, { type: 'autoMax' }], showValue: false,
        direction: 'leftToRight', borderArgb: 0xff00b050, negativeBorderArgb: 0xffffa500, axisArgb: 0xff000000 })
      expect(barOf(8)).toMatchObject({ kind: 'dataBar', bounds: [{ type: 'percentile', value: 10 }, { type: 'num', value: 90 }] })
      expect(barOf(9)).toMatchObject({ kind: 'unresolved', type: 'dataBar', reason: 'cf-unsupported-rule' })
      expect(barOf(10)).toMatchObject({ kind: 'dataBar', dialect: 'x14', bounds: [{ type: 'autoMin' }, { type: 'autoMax' }] })
      // A canonical or absent formula keeps the attribute predicate; the divergent one is withheld, never the attribute (S5-G).
      expect(barOf(15)).toMatchObject({ kind: 'timePeriod', period: 'today' })
      expect(barOf(16)).toMatchObject({ kind: 'unresolved', type: 'timePeriod', reason: 'cf-unsupported-rule' })
      expect(barOf(17)).toMatchObject({ kind: 'timePeriod', period: 'lastMonth' })
      const iconOf = (priority: number) => expected.rules.find(r => r.priority === priority)!.definition
      expect(iconOf(11)).toMatchObject({ kind: 'iconSet', dialect: 'base', set: '3Arrows', showValue: true, percentAttr: true })
      // A base x14-only family is a recorded dialect outcome, so it is withheld whatever the evidence.
      expect(iconOf(12)).toMatchObject({ kind: 'unresolved', type: 'iconSet', reason: 'cf-unsupported-rule' })
      expect(iconOf(13)).toMatchObject({ kind: 'iconSet', dialect: 'x14', set: '3Arrows',
        custom: [{ set: '3Flags', iconId: 0 }, { set: 'NoIcons', iconId: 0 }, { set: '3Stars', iconId: 2 }],
        thresholds: [{ bound: { type: 'num', value: 0 }, gte: true }, { bound: { type: 'num', value: 20 }, gte: false }, { bound: { type: 'num', value: 40 }, gte: true }] })
      expect(iconOf(14)).toMatchObject({ kind: 'iconSet', dialect: 'x14', set: '3Stars', percentAttr: false })
      const bytes = new TextEncoder().encode(xml)
      // Every split point is exercised, so chunk boundaries fall inside each <x14:dataBar> too.
      for (let i = 0; i <= bytes.length; i++) {
        const result = await stream(xml, i, meta)
        expect(result.lease.value).toEqual(expected)
        result.lease.release(); expect(result.memory.currentBytes).toBe(0)
      }
    })
  })
  it('resident and streamed formula rules are deep equal at every chunk boundary', async () => {
    // Stage 5 (Task 32): an expression, a formula operand, a divergent predicate, a quoted cross-sheet name and a standalone x14
    // `<xm:f>`, with boundaries falling inside each `<formula>` and `<xm:f>`; unresolved shapes ride along unchanged.
    const expression = '<conditionalFormatting sqref="B2:D9"><cfRule type="expression" priority="1"><formula>$B2&gt;$C2</formula></cfRule></conditionalFormatting>'
    const operand = '<conditionalFormatting sqref="F2:F9"><cfRule type="cellIs" priority="2" operator="between"><formula>$H$1</formula><formula>$H$2*2</formula></cfRule></conditionalFormatting>'
    const divergent = '<conditionalFormatting sqref="J2:J9"><cfRule type="containsBlanks" priority="3"><formula>ISBLANK(J2)</formula></cfRule></conditionalFormatting>'
    const quoted = `<conditionalFormatting sqref="L2:L9"><cfRule type="expression" priority="4"><formula>'My Sheet'!$A$1&lt;L2</formula></cfRule></conditionalFormatting>`
    const refused = '<conditionalFormatting sqref="N2:N9"><cfRule type="expression" priority="5"><formula>INDIRECT("A1")</formula></cfRule></conditionalFormatting>'
    const x14 = '<x14:conditionalFormatting><x14:cfRule type="expression" priority="6" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>$P2="茶"</xm:f></x14:cfRule><xm:sqref>P2:P9</xm:sqref></x14:conditionalFormatting>'
    // Task 33: a divergent timePeriod formula is an expression under the formula reading, and Excel's own formula stays a timePeriod.
    const tpDivergent = '<conditionalFormatting sqref="R2:R9"><cfRule type="timePeriod" priority="7" timePeriod="today"><formula>R2=TODAY()-3</formula></cfRule></conditionalFormatting>'
    const tpCanonical = '<conditionalFormatting sqref="T2:T9"><cfRule type="timePeriod" priority="8" timePeriod="thisMonth"><formula>AND(MONTH(T2)=MONTH(TODAY()),YEAR(T2)=YEAR(TODAY()))</formula></cfRule></conditionalFormatting>'
    const xml = wrap(expression + operand + divergent + quoted + refused + tpDivergent + tpCanonical + `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>${x14}</x14:conditionalFormattings></ext></extLst>`)
    await withExactCfEvidence(['formula:dialect-base', 'formula:dialect-x14', 'formula:anchor-single', 'formula:anchor-mixed-abs', 'formula:cross-sheet-reference', 'formula:cross-sheet-quoted-name', 'formula:predicate-timeperiod-divergent'], async () => {
      const expected = parseCfSheet(parseXml(xml), styles, createCfMemory()).value
      expect(expected.rules.map(r => r.definition.kind)).toEqual(['expression', 'cellIs', 'expression', 'expression', 'unresolved', 'expression', 'timePeriod', 'expression'])
      expect(expected.rules[5]!.definition).toMatchObject({ kind: 'expression', sourceType: 'timePeriod' })
      expect(expected.rules[3]!.definition).toMatchObject({ dependencySheets: ['My Sheet'] })
      expect(expected.rules[1]!.definition).toMatchObject({ operands: [{ formula: { nodes: 1 } }, { formula: { nodes: 3 } }] })
      const bytes = new TextEncoder().encode(xml)
      for (let i = 0; i <= bytes.length; i++) {
        const result = await stream(xml, i)
        expect(result.lease.value).toEqual(expected)
        result.lease.release(); expect(result.memory.currentBytes).toBe(0)
      }
    })
    // Without evidence both paths stay unresolved, deep-equal (pinned under exactly no evidence, Ruling S5-STUB).
    const none = withExactCfEvidence([], () => parseCfSheet(parseXml(xml), styles, createCfMemory()).value)
    expect(none.rules.map(r => r.definition.kind)).toEqual(['unresolved', 'unresolved', 'unresolved', 'unresolved', 'unresolved', 'unresolved', 'timePeriod', 'unresolved'])
    expect((await withExactCfEvidence([], () => stream(xml, 100))).lease.value).toEqual(none)
  })
  it.each(['formula', 'color', 'dxf'])('ten thousand singleton duplicates do not build ten thousand nodes: %s', async tag => {
    const xml = wrap(`<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1">${(`<${tag}>1</${tag}>`).repeat(10_000)}</cfRule></conditionalFormatting>`)
    // Pinned under exactly no evidence (Ruling S5-STUB): the stub stays unresolved whatever Stage 5 publishes.
    const expected = withExactCfEvidence([], () => parseCfSheet(parseXml(xml), styles, createCfMemory()).value)
    const result = await withExactCfEvidence([], () => stream(xml))
    if (tag === 'formula') expect(result.lease.value.rules[0]?.definition.kind).toBe('unresolved')
    else expect(result.lease.value.unresolvedSheet).toBeDefined()
    expect(result.lease.value).toEqual(expected)
    expect(result.handle.stats.nodes).toBeLessThan(20)
    expect(result.handle.stats.textUnits).toBeLessThan(20)
    result.lease.release()
  })
  it('overflow before sqref is known withholds sheet CF', async () => {
    const result = await stream(wrap(`<conditionalFormatting><cfRule>${'<formula>1</formula>'.repeat(10_000)}</cfRule></conditionalFormatting>`))
    expect(result.lease.value.unresolvedSheet).toBeDefined(); expect(result.lease.value.rules).toEqual([])
    result.lease.release()
  })
  it('cancelled capture releases partial construction without publishing a result', () => {
    const memory = createCfMemory(), h = createCfStreamEvents(styles, memory, {})
    h.events.open('worksheet', ['xmlns'], [SS_NS], 1, false)
    h.events.open('conditionalFormatting', ['sqref'], ['A1'], 1, false)
    expect(memory.currentBytes).toBeGreaterThan(0)
    h.dispose(); expect(memory.currentBytes).toBe(0); expect(() => h.finish()).toThrow()
  })
  it.each(['beforeSheetData', 'afterSheetData'] as const)('late CF reconciles DXFs before sheet readiness: %s', async position => {
    const bytes = xlsxBytes({ [position]: base.replace('priority="1"', 'priority="1" dxfId="0"'), styles: '<cellXfs><xf/></cellXfs><dxfs><dxf><font><b/></font></dxf></dxfs>' })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer]) }, { docId: `cf-${position}`, profile: 'D' })
    try {
      await session.done
      const sheet = session.workbook.sheets[0]!
      expect(sheet.state).toBe('complete')
      expect(sheet.metadata.hasConditionalFormatting).toBe(true)
      expect(sheet.conditionalFormatting?.rules).toHaveLength(1)
      expect(sheet.conditionalFormatting?.rules[0]?.dxf).toBeDefined()
      expect(sheet.cfOwner?.bytes).toBeGreaterThan(0)
    } finally { await session.release() }
    expect(session.workbook.sheets[0]!.cfOwner).toBeUndefined()
  })
})

it('loads raw DXFs only once for multiple CF sheets and never for a plain sheet', async () => {
  const loader = vi.spyOn(formatting, 'pagedCfStyles')
  const plain = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({})).buffer]) }, { docId: 'plain-cf', profile: 'D' })
  await plain.done; await plain.release(); expect(loader).not.toHaveBeenCalled()
  const entries = xlsxEntries({ afterSheetData: base })
  const workbook = entries.find(e => e.name === 'xl/workbook.xml')!
  workbook.data = String(workbook.data).replace('</sheets>', '<sheet name="T" sheetId="2" r:id="rId4"/></sheets>')
  const rels = entries.find(e => e.name === 'xl/_rels/workbook.xml.rels')!
  rels.data = String(rels.data).replace('</Relationships>', '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
  entries.push({ name: 'xl/worksheets/sheet2.xml', data: wrap(base) })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(buildZip(entries).bytes).buffer]) }, { docId: 'two-cf', profile: 'D' })
  try {
    await session.done
    expect(session.workbook.sheets.map(s => s.conditionalFormatting?.rules.length)).toEqual([1, 1])
    expect(loader).toHaveBeenCalledTimes(1)
  } finally { await session.release(); loader.mockRestore() }
})
it('cancel during the paged CF pass drains the partial collector and publishes nothing', async () => {
  const controller = new AbortController()
  const original = capturing.createCfStreamEvents
  let memory: ReturnType<typeof createCfMemory> | undefined
  const spy = vi.spyOn(capturing, 'createCfStreamEvents').mockImplementation((styles, account, meta) => {
    memory = account
    const h = original(styles, account, meta), open = h.events.open
    h.events.open = (...args) => { open(...args); if (args[0] === 'cfRule') controller.abort() }
    return h
  })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: base })).buffer]) }, { docId: 'cancel-cf', profile: 'D', signal: controller.signal })
  try {
    expect((await session.done).kind).toBe('cancelled')
    expect(session.workbook.sheets[0]?.conditionalFormatting).toBeUndefined()
    expect(memory?.currentBytes).toBe(0)
  } finally { await session.release(); spy.mockRestore() }
})
it('keeps excluded AlternateContent and wrong-namespace paths out of effective CF', async () => {
  for (const inner of [
    `<AlternateContent><Choice>${base}</Choice></AlternateContent>`,
    '<extLst><ext xmlns="urn:wrong" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>',
  ]) {
    const result = await stream(wrap(inner))
    expect(result.lease.value.rules).toEqual([])
    // Diagnostics may describe the malformed wrapper at different granularity.
    expect(parseCfSheet(parseXml(wrap(inner)), styles, createCfMemory()).value.rules).toEqual([])
    result.lease.release()
  }
})
it('refuses CF before constructing unadmitted raw DXFs, preserving the base import', async () => {
  const { ConstructionMemory } = await import('../src/xlsx/storage/constructionMemory')
  const { PackageError } = await import('../src/core/package-error')
  const original = ConstructionMemory.prototype.share
  const allocator = vi.spyOn(ConstructionMemory.prototype, 'share').mockImplementation(function (this: InstanceType<typeof ConstructionMemory>, name, options) {
    const share = original.call(this, name, options)
    if (name === 'CF raw DXF XML') return { ...share, need() { throw new PackageError('resource-limit', 'test admission refusal') } }
    return share
  })
  const loader = vi.spyOn(formatting, 'pagedCfStyles')
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: base })).buffer]) }, { docId: 'refuse-cf', profile: 'D' })
  try {
    expect((await session.done).kind).toBe('complete')
    expect(loader).not.toHaveBeenCalled()
    expect(session.workbook.sheets[0]?.conditionalFormatting?.unresolvedSheet).toBeDefined()
  } finally { await session.release(); allocator.mockRestore(); loader.mockRestore() }
})
it('publishes late x14 definitions with the same normalized model as resident parsing', async () => {
  const extension = '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormatting><x14:cfRule type="expression" priority="1" id="{01234567-89AB-CDEF-0123-456789ABCDEF}"><xm:f>A1&gt;2</xm:f><x14:dxf><x14:font><x14:b/></x14:font></x14:dxf></x14:cfRule><xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(xlsxBytes({ afterSheetData: extension })).buffer]) }, { docId: 'x14-cf', profile: 'D' })
  try {
    await session.done
    const sheet = session.workbook.sheets[0]!
    const expected = parseCfSheet(parseXml(wrap(extension)), styles, createCfMemory(), { sheetId: sheet.sheetId, partPath: sheet.partPath })
    expect(sheet.metadata.hasConditionalFormatting).toBe(true)
    expect(sheet.conditionalFormatting).toEqual(expected.value)
    expected.release()
  } finally { await session.release() }
})
it('worker admission precedes capture growth and retained ownership lasts until explicit release', async () => {
  const { ConstructionMemory } = await import('../src/xlsx/storage/constructionMemory')
  const { MemoryPool } = await import('../src/xlsx/view/tileCache')
  const construction = new ConstructionMemory(new MemoryPool({ limitBytes: 1024 * 1024 }), 'cf-byte-test')
  const memory = capturing.createPagedCfMemory(construction)
  const h = capturing.createCfStreamEvents(styles, memory, {})
  const bytes = new TextEncoder().encode(wrap(base))
  await tokenizeBytes((async function* () { yield bytes })(), h.events)
  expect(memory.constructionBytes).toBeGreaterThan(0)
  const lease = h.finish()
  expect(memory.constructionBytes).toBe(0)
  expect(memory.retainedBytes).toBe(lease.bytes)
  h.dispose(); expect(memory.retainedBytes).toBe(lease.bytes)
  expect(construction.transientBytes).toBeGreaterThanOrEqual(lease.bytes)
  memory.close(); expect(memory.retainedBytes).toBeGreaterThan(0)
  lease.release(); expect(memory.currentBytes).toBe(0)
  expect(construction.transientBytes).toBe(0)
  construction.releaseAll()
})
it('resident duplicate preflight reads only a bounded overflow witness, without copying the child list', () => {
  let reads = 0
  const formulas = new Proxy(Array.from({ length: 10_000 }, () => ({ '@attrs': {}, '#text': '1' })), {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  const worksheet = { '@attrs': {}, conditionalFormatting: { '@attrs': { sqref: 'A1' }, cfRule: { '@attrs': { type: 'expression', priority: '1' }, formula: formulas } } }
  const lease = parseCfSheet(worksheet, styles, createCfMemory())
  expect(lease.value.rules[0]?.definition.kind).toBe('unresolved')
  expect(reads).toBeLessThan(16)
  lease.release()
})

describe('regression bot P2: paged CF dateContext', () => {
  it('keeps local timeZone and epochNowMs in default dateContext', () => {
    const fs = require('node:fs')
    const src = fs.readFileSync('src/xlsx/view/pagedSource.ts', 'utf8')
    expect(src).toMatch(/timeZone: Intl\.DateTimeFormat\(\)\.resolvedOptions\(\)\.timeZone \?\? 'UTC'/)
    // Ruling S5-M: the instant is sampled once per preparation (default: the wall clock) and shared by predicates and formulas.
    expect(src).toMatch(/const instant = \(options\?\.now \?\? Date\.now\)\(\)/)
    expect(src).toMatch(/epochNowMs: instant/)
  })
})
