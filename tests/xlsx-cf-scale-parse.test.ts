import { afterEach, describe, expect, it } from 'vitest'
import { parseXml } from '../src/core/xml'
import { tokenizeBytes } from '../src/core/xml-stream'
import { createCfDiagnostics, createCfMemory } from '../src/xlsx/cf/limits'
import {
  parseCfSheet,
  MEASURED_CF_OFFICE_THEME,
  scaleRequiredCapabilities,
  type CfScaleBoundInput,
  type CfScaleColorInput,
  type CfStylesSource,
} from '../src/xlsx/cf/parse'
import { createCfStreamEvents } from '../src/xlsx/cf/capture'
import { resolveSheetColor } from '../src/xlsx/sheetColors'
import { injectExactCfEvidence, withExactCfEvidence } from '../src/xlsx/cf/evidence'
import { COLOR_SCALE_WRITES, type CfRule } from '../src/xlsx/cf/model'

const defaultStyles: CfStylesSource = {
  colors: {
    theme: {
      accent1: '4472C4',
      accent2: 'ED7D31',
      accent3: 'A5A5A5',
      accent4: 'FFC000',
      accent5: '5B9BD5',
      accent6: '70AD47',
    },
  },
  dxfNodes: [],
}

function parse(xmlText: string, styles: CfStylesSource = defaultStyles) {
  const memory = createCfMemory()
  const lease = parseCfSheet(
    parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
        `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ` +
        `xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
        xmlText +
        `</worksheet>`,
    ),
    styles,
    memory,
  )
  return { lease, value: lease.value, memory }
}

function ruleByPriority(rules: readonly CfRule[], priority: number): CfRule {
  const found = rules.filter(r => r.priority === priority)
  expect(found).toHaveLength(1)
  return found[0]!
}

afterEach(() => injectExactCfEvidence(null))

describe('CF Stage 2 colour-scale model and parsing (Task 13)', () => {
  it('canonical colour-scale write mask is [\'colorScale\', \'fill\'] for both resolved and unresolved scales', () => {
    // 1. Unresolved scale (without evidence)
    injectExactCfEvidence([])
    const unevidenced = parse(
      `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>` +
        `<cfvo type="min"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale>` +
        `</cfRule>` +
        `</conditionalFormatting>`,
    )
    const unevidencedRule = unevidenced.value.rules[0]!
    expect(unevidencedRule.definition.kind).toBe('unresolved')
    expect(new Set(unevidencedRule.writes)).toEqual(new Set(['colorScale', 'fill']))
    expect(COLOR_SCALE_WRITES).toEqual(new Set(['colorScale', 'fill']))

    // 2. Resolved scale (with evidence injected)
    withExactCfEvidence(['scale:thresholds-min-max'], () => {
      const evidenced = parse(
        `<conditionalFormatting sqref="A1:A10">` +
          `<cfRule type="colorScale" priority="1">` +
          `<colorScale>` +
          `<cfvo type="min"/><cfvo type="max"/>` +
          `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
          `</colorScale>` +
          `</cfRule>` +
          `</conditionalFormatting>`,
      )
      const evidencedRule = evidenced.value.rules[0]!
      expect(evidencedRule.definition.kind).toBe('colorScale')
      expect(new Set(evidencedRule.writes)).toEqual(new Set(['colorScale', 'fill']))

      // 3. S2-I two-scale withholding on one cell:
      // A higher-priority scale's colorScale slot conflicts with a lower scale on the same cell;
      // following S2-I's text exactly ("two resolved scales on one cell withhold the fill"),
      // conflicting colorScale and fill slots ensure both scales withhold the fill on that cell.
      const twoScales = parse(
        `<conditionalFormatting sqref="B2:B10">` +
          `<cfRule type="colorScale" priority="1">` +
          `<colorScale>` +
          `<cfvo type="min"/><cfvo type="max"/>` +
          `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
          `</colorScale>` +
          `</cfRule>` +
          `<cfRule type="colorScale" priority="2">` +
          `<colorScale>` +
          `<cfvo type="min"/><cfvo type="max"/>` +
          `<color rgb="FF0000FF"/><color rgb="FFFFFF00"/>` +
          `</colorScale>` +
          `</cfRule>` +
          `</conditionalFormatting>`,
      )
      expect(twoScales.value.rules).toHaveLength(2)
      const r1 = twoScales.value.rules[0]!
      const r2 = twoScales.value.rules[1]!
      expect(new Set(r1.writes)).toEqual(new Set(['colorScale', 'fill']))
      expect(new Set(r2.writes)).toEqual(new Set(['colorScale', 'fill']))

      // Overlapping write masks share both 'colorScale' and 'fill' slots
      const commonSlots = [...r1.writes].filter(w => r2.writes.has(w)).sort()
      expect(commonSlots).toEqual(['colorScale', 'fill'])
    })
  })

  it('spec example scale resolves to three ordered points', () => {
    const xml =
      `<conditionalFormatting sqref="B2:B20 D2:D20">` +
      `<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="B2:B20">` +
      `<cfRule type="top10" dxfId="0" priority="2" rank="3"/>` +
      `<cfRule type="colorScale" priority="3">` +
      `<colorScale>` +
      `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
      `<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>` +
      `</colorScale>` +
      `</cfRule>` +
      `</conditionalFormatting>`

    // Without evidence: yields cf-unsupported-rule and scale's fill is tainted
    injectExactCfEvidence([])
    const withoutEvidence = parse(xml)
    expect(withoutEvidence.value.rules).toHaveLength(3)
    const scaleUnresolved = ruleByPriority(withoutEvidence.value.rules, 3)
    expect(scaleUnresolved.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })
    expect(scaleUnresolved.writes.has('colorScale')).toBe(true)
    expect(scaleUnresolved.writes.has('fill')).toBe(true)

    // With evidence injected: §3 example yields min / percentile 50 / max with points
    // FFF8696B, FFFFEB84, FF63BE7B.
    withExactCfEvidence(
      [
        'scale:thresholds-min-max',
        'scale:thresholds-percentile-inc-exc',
        'scale:interpolation-spec-palette',
      ],
      () => {
        const withEvidence = parse(xml)
        expect(withEvidence.value.rules).toHaveLength(3)
        const scaleResolved = ruleByPriority(withEvidence.value.rules, 3)
        expect(scaleResolved.definition.kind).toBe('colorScale')
        if (scaleResolved.definition.kind === 'colorScale') {
          expect(scaleResolved.definition.points).toHaveLength(3)
          expect(scaleResolved.definition.points[0]).toEqual({
            bound: { type: 'min' },
            argb: 0xfff8696b,
          })
          expect(scaleResolved.definition.points[1]).toEqual({
            bound: { type: 'percentile', value: 50 },
            argb: 0xffffeb84,
          })
          expect(scaleResolved.definition.points[2]).toEqual({
            bound: { type: 'max' },
            argb: 0xff63be7b,
          })
        }
        expect(scaleResolved.writes.has('colorScale')).toBe(true)
        expect(scaleResolved.writes.has('fill')).toBe(true)
      },
    )
  })

  it('scale structure refusals', () => {
    injectExactCfEvidence([])

    // 1. 2 cfvo + 3 colors -> cf-invalid-rule
    const mismatch = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(mismatch.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-invalid-rule',
    })
    expect(mismatch.value.diagnostics.some(d => d.feature === 'cf-invalid-rule')).toBe(true)

    // 2. missing type -> cf-invalid-cfvo
    const missingType = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo val="0"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(missingType.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-invalid-cfvo',
    })
    expect(missingType.value.diagnostics.some(d => d.feature === 'cf-invalid-cfvo')).toBe(true)

    // 3. num val="$A$1", type="formula", activePresent="1" -> cf-unsupported-rule
    const cellRef = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="num" val="$A$1"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(cellRef.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })

    const formulaCfvo = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="formula" val="SUM(A1:A5)"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(formulaCfvo.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })

    const activePres = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1" activePresent="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(activePres.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })

    // 4. min val="0" and mid-list min -> unresolved without evidence
    const legacyMin = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min" val="0"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(legacyMin.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })
    // Under evidence it resolves:
    withExactCfEvidence(
      ['scale:thresholds-min-max', 'scale:thresholds-legacy-min-val-0'],
      () => {
        const resolvedLegacy = parse(
          `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
            `<cfvo type="min" val="0"/><cfvo type="max"/>` +
            `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
            `</colorScale></cfRule></conditionalFormatting>`,
        )
        expect(resolvedLegacy.value.rules[0]!.definition.kind).toBe('colorScale')
      },
    )

    const maxFirst = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="max"/><cfvo type="min"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(maxFirst.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })
    // Under evidence it resolves:
    withExactCfEvidence(
      ['scale:thresholds-min-max', 'scale:thresholds-min-midlist-max-first'],
      () => {
        const resolvedPos = parse(
          `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
            `<cfvo type="max"/><cfvo type="min"/>` +
            `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
            `</colorScale></cfRule></conditionalFormatting>`,
        )
        expect(resolvedPos.value.rules[0]!.definition.kind).toBe('colorScale')
      },
    )

    // 5. 6 cfvo -> cf-resource-limit before construction
    const overCfvo = parse(
      `<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="num" val="1"/><cfvo type="num" val="2"/><cfvo type="num" val="3"/><cfvo type="num" val="4"/><cfvo type="max"/>` +
        `<color rgb="FFFF0000"/><color rgb="FF00FF00"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
    )
    expect(overCfvo.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      reason: 'cf-resource-limit',
    })
    expect(overCfvo.value.diagnostics.some(d => d.feature === 'cf-resource-limit')).toBe(true)

    // 6. Standalone x14 scale, literal or reference xm:f -> cf-unsupported-rule (OD-3, S2-K)
    const x14Xml =
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AA7539E650}"><x14:conditionalFormattings><x14:conditionalFormatting>` +
      `<x14:cfRule type="colorScale" priority="1" id="{01234567-89AB-CDEF-0123-456789ABCDEF}">` +
      `<x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:color rgb="FFFF0000"/><x14:color rgb="FF00FF00"/></x14:colorScale>` +
      `</x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    const x14Scale = parse(x14Xml)
    expect(x14Scale.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      reason: 'cf-unsupported-rule',
    })

    // 7. Diagnostics are deduplicated and the sheet cap stays at 64
    const diagMem = createCfMemory()
    const diag = createCfDiagnostics(diagMem)
    for (let i = 0; i < 100; i++) {
      diag.add({ kind: 'unsupported', feature: 'cf-invalid-cfvo', message: 'cfvo missing type' })
    }
    expect(diag.entries).toHaveLength(1)
    diag.dispose()

    let dupXml = ''
    for (let i = 1; i <= 100; i++) {
      dupXml += `<conditionalFormatting sqref="A${i}"><cfRule type="colorScale" priority="${i}"><colorScale><cfvo val="0"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>`
    }
    const dupParsed = parse(dupXml)
    expect(dupParsed.value.diagnostics).toHaveLength(64) // sheet cap stays at 64
  })

  it('endpoint colours resolve once on both paths', async () => {
    const fixtureTheme = {
      // The palette the scale colours were measured against; any other palette withholds (colors-custom-theme).
      theme: MEASURED_CF_OFFICE_THEME,
    }
    const styles = { colors: fixtureTheme, dxfNodes: [] }

    // Expected ARGB computed once through resolveSheetColor under fixture theme
    const expectedHex = resolveSheetColor({ theme: '4', tint: '-0.4' }, fixtureTheme)
    expect(expectedHex).toBeDefined()
    const expectedArgb = parseInt(expectedHex!, 16)

    const xmlBody =
      `<conditionalFormatting sqref="B2:B10">` +
      `<cfRule type="colorScale" priority="1">` +
      `<colorScale>` +
      `<cfvo type="min"/><cfvo type="max"/>` +
      `<color theme="4" tint="-0.4"/><color rgb="FFFFFFFF"/>` +
      `</colorScale>` +
      `</cfRule>` +
      `</conditionalFormatting>`

    await withExactCfEvidence(
      ['scale:thresholds-min-max', 'scale:colors-theme-tint'],
      async () => {
      // 1. Resident path
      const resident = parse(xmlBody, styles)
      expect(resident.value.rules).toHaveLength(1)
      const resRule = resident.value.rules[0]!
      expect(resRule.definition.kind).toBe('colorScale')
      if (resRule.definition.kind === 'colorScale') {
        expect(resRule.definition.points[0]!.argb).toBe(expectedArgb)
      }

      // 2. Paged path (streaming)
      const fullXml =
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
        `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ` +
        `xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
        xmlBody +
        `</worksheet>`

      const memory = createCfMemory()
      const handle = createCfStreamEvents(styles, memory, {})
      const bytes = new TextEncoder().encode(fullXml)
      await tokenizeBytes((async function* () { yield bytes })(), handle.events)
      const streamedLease = handle.finish()
      const streamedRule = streamedLease.value.rules[0]!
      expect(streamedRule.definition.kind).toBe('colorScale')
      if (streamedRule.definition.kind === 'colorScale') {
        expect(streamedRule.definition.points[0]!.argb).toBe(expectedArgb)
      }
      expect(streamedLease.value).toEqual(resident.value)
      streamedLease.release()
    })

    // auto and alpha other than FF resolve to unresolved, and no #000000 fallback appears
    injectExactCfEvidence([])
    const autoScale = parse(
      `<conditionalFormatting sqref="C1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="max"/>` +
        `<color auto="1"/><color rgb="FFFFFFFF"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
      styles,
    )
    expect(autoScale.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })
    // Ensure no #000000 fallback was registered
    expect(autoScale.value.rules[0]!.definition.kind).toBe('unresolved')

    const alphaScale = parse(
      `<conditionalFormatting sqref="D1"><cfRule type="colorScale" priority="1"><colorScale>` +
        `<cfvo type="min"/><cfvo type="max"/>` +
        `<color rgb="00F8696B"/><color rgb="FFFFFFFF"/>` +
        `</colorScale></cfRule></conditionalFormatting>`,
      styles,
    )
    expect(alphaScale.value.rules[0]!.definition).toMatchObject({
      kind: 'unresolved',
      type: 'colorScale',
      reason: 'cf-unsupported-rule',
    })
    expect(alphaScale.value.rules[0]!.definition.kind).toBe('unresolved')
  })

  describe('fail-closed colour-scale capability authorization', () => {
    it('(a) theme + percentile with only the theme evidence -> unresolved, cf-unsupported-rule', () => {
      const xml =
        `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>` +
        `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
        `<color theme="4" tint="-0.4"/><color rgb="FFFFFFFF"/><color rgb="FF000000"/>` +
        `</colorScale>` +
        `</cfRule>` +
        `</conditionalFormatting>`
      withExactCfEvidence(['scale:colors-theme-tint'], () => {
        const res = parse(xml)
        expect(res.value.rules[0]!.definition).toMatchObject({
          kind: 'unresolved',
          type: 'colorScale',
          reason: 'cf-unsupported-rule',
        })
      })
    })

    it('(b) the spec palette with only scale:thresholds-percentile-inc-exc -> unresolved', () => {
      const xml =
        `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>` +
        `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
        `<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>` +
        `</colorScale>` +
        `</cfRule>` +
        `</conditionalFormatting>`
      withExactCfEvidence(['scale:thresholds-percentile-inc-exc'], () => {
        const res = parse(xml)
        expect(res.value.rules[0]!.definition).toMatchObject({
          kind: 'unresolved',
          type: 'colorScale',
          reason: 'cf-unsupported-rule',
        })
      })
    })

    it('(c) the same bounds with different colours behave identically to the spec palette under identical evidence', () => {
      const diffColoursXml =
        `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>` +
        `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
        `<color rgb="FF0000FF"/><color rgb="FFFFFFFF"/><color rgb="FFFF0000"/>` +
        `</colorScale>` +
        `</cfRule>` +
        `</conditionalFormatting>`
      // With only percentile evidence: unresolved
      withExactCfEvidence(['scale:thresholds-percentile-inc-exc'], () => {
        const res = parse(diffColoursXml)
        expect(res.value.rules[0]!.definition).toMatchObject({
          kind: 'unresolved',
          type: 'colorScale',
          reason: 'cf-unsupported-rule',
        })
      })
      // With full required set: resolved
      const fullEvidence = [
        'scale:thresholds-min-max',
        'scale:thresholds-percentile-inc-exc',
        'scale:interpolation-spec-palette',
      ]
      withExactCfEvidence(fullEvidence, () => {
        const res = parse(diffColoursXml)
        expect(res.value.rules[0]!.definition.kind).toBe('colorScale')
      })
    })

    it('(d) the full required set -> resolved', () => {
      const specXml =
        `<conditionalFormatting sqref="A1:A10">` +
        `<cfRule type="colorScale" priority="1">` +
        `<colorScale>` +
        `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` +
        `<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>` +
        `</colorScale>` +
        `</cfRule>` +
        `</conditionalFormatting>`
      const fullEvidence = [
        'scale:thresholds-min-max',
        'scale:thresholds-percentile-inc-exc',
        'scale:interpolation-spec-palette',
      ]
      withExactCfEvidence(fullEvidence, () => {
        const res = parse(specXml)
        expect(res.value.rules[0]!.definition.kind).toBe('colorScale')
        if (res.value.rules[0]!.definition.kind === 'colorScale') {
          expect(res.value.rules[0]!.definition.points).toHaveLength(3)
          expect(res.value.rules[0]!.definition.points[0]!.argb).toBe(0xfff8696b)
          expect(res.value.rules[0]!.definition.points[1]!.argb).toBe(0xffffeb84)
          expect(res.value.rules[0]!.definition.points[2]!.argb).toBe(0xff63be7b)
        }
      })
    })

    it('(e) a table test of scaleRequiredCapabilities for each feature', () => {
      const cases: Array<{
        name: string
        bounds: CfScaleBoundInput[]
        colours: CfScaleColorInput[]
        expected: string[]
      }> = [
        {
          name: '2-point min/max with standard RGB',
          bounds: [{ type: 'min' }, { type: 'max' }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-min-max'],
        },
        {
          name: '2-point num bounds',
          bounds: [{ type: 'num', value: 20 }, { type: 'num', value: 80 }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-num-bounds'],
        },
        {
          name: '2-point percent bounds',
          bounds: [{ type: 'percent', value: 10 }, { type: 'percent', value: 90 }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-percent'],
        },
        {
          name: '2-point percentile bounds',
          bounds: [{ type: 'percentile', value: 25 }, { type: 'percentile', value: 75 }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-percentile-inc-exc'],
        },
        {
          name: 'out-of-bounds percent/percentile',
          bounds: [{ type: 'percent', value: -10 }, { type: 'percentile', value: 150 }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: [
            'scale:thresholds-percent',
            'scale:thresholds-percentile-inc-exc',
            'scale:thresholds-percent-percentile-out-of-bounds',
          ],
        },
        {
          name: 'legacy min val="0"',
          bounds: [{ type: 'min', val: '0' }, { type: 'max' }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-min-max', 'scale:thresholds-legacy-min-val-0'],
        },
        {
          name: 'positional max first / min mid-list',
          bounds: [{ type: 'max' }, { type: 'min' }],
          colours: ['FFFF0000', 'FF00FF00'],
          expected: ['scale:thresholds-min-max', 'scale:thresholds-min-midlist-max-first'],
        },
        {
          name: '3-point scale interpolation (spec palette)',
          bounds: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
          colours: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
          expected: [
            'scale:interpolation-spec-palette',
            'scale:thresholds-min-max',
            'scale:thresholds-percentile-inc-exc',
          ],
        },
        {
          name: 'theme+tint colour form',
          bounds: [{ type: 'min' }, { type: 'max' }],
          colours: [{ theme: '4', tint: '-0.4' }, { rgb: 'FFFFFFFF' }],
          expected: ['scale:thresholds-min-max', 'scale:colors-theme-tint'],
        },
        {
          name: 'indexed colour form',
          bounds: [{ type: 'min' }, { type: 'max' }],
          colours: [{ indexed: '10' }, { indexed: '12' }],
          expected: ['scale:thresholds-min-max', 'scale:colors-indexed'],
        },
        {
          name: 'alpha other than FF colour form',
          bounds: [{ type: 'min' }, { type: 'max' }],
          colours: ['00F8696B', '0063BE7B'],
          expected: ['scale:thresholds-min-max', 'scale:colors-alpha-00'],
        },
        {
          name: 'auto colour form',
          bounds: [{ type: 'min' }, { type: 'max' }],
          colours: ['auto', 'auto'],
          expected: ['scale:thresholds-min-max', 'scale:colors-auto'],
        },
      ]

      for (const c of cases) {
        const actual = scaleRequiredCapabilities(c.bounds, c.colours)
        expect(actual, c.name).toEqual(new Set(c.expected))
      }
    })
  })
})
