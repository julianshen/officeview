import { describe, expect, it } from 'vitest'
import { getChildren, parseXml, type XmlNode } from '../src/core/xml'
import type { SheetColorContext } from '../src/xlsx/sheetColors'
import { parseCfDxf, resolveCfCellStyle } from '../src/xlsx/cf/dxf'

const colors: SheetColorContext = {}

function dxfNode(inner: string): XmlNode {
  const root = parseXml(
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs><dxf>${inner}</dxf></dxfs></styleSheet>`,
  )
  return getChildren(getChildren(root, 'dxfs')[0], 'dxf')[0]!
}

function x14DxfNode(inner: string): XmlNode {
  const root = parseXml(
    `<root xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dxf>${inner}</x14:dxf></root>`,
  )
  return getChildren(root, 'dxf')[0]!
}

describe('CF strict dxf parsing', () => {
  it('parses a supported fill+font dxf with presence preserved', () => {
    const { properties, writes, unsupported } = parseCfDxf(
      dxfNode(
        `<font><b val="1"/><color rgb="FF9C0006"/></font>` +
          `<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>`,
      ),
      colors,
    )
    expect(unsupported).toEqual([])
    expect(properties['font.bold']).toBe(true)
    expect(properties['font.color']).toBe('FF9C0006')
    expect(properties['fill.pattern']).toBe('solid')
    expect(properties['fill.fg']).toBe('FFFFC7CE')
    expect(writes.has('font.bold')).toBe(true)
    expect(writes.has('font.color')).toBe(true)
    expect(writes.has('fill')).toBe(true)
    expect(writes.has('font.italic')).toBe(false)
  })

  it('a dxf solid fill with only fgColor paints no fill (Ruling T11-B)', () => {
    const { properties, unsupported } = parseCfDxf(
      dxfNode(`<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>`),
      colors,
    )
    expect(unsupported).toEqual([])
    expect(properties['fill.pattern']).toBe('solid')
    expect(properties['fill.fg']).toBe('FFFFC7CE')
    expect('fill.bg' in properties).toBe(false)
    const style = resolveCfCellStyle(undefined, { properties }, colors)
    expect(style.fillColor).toBeUndefined()
    // M1: A fgColor-only dxf leaves base fill untouched (does not clear to white/undefined)
    const baseWithFill = { fillColor: 'FFFFFF00', numFmtId: 0 }
    const stylePreserved = resolveCfCellStyle(baseWithFill, { properties }, colors)
    expect(stylePreserved.fillColor).toBe('FFFFFF00')
  })

  it('accepts native Excel preset dxf fill without patternType (Ruling C2)', () => {
    const { properties, unsupported } = parseCfDxf(
      dxfNode(`<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>`),
      colors,
    )
    expect(unsupported).toEqual([])
    expect(properties['fill.bg']).toBe('FFFFC7CE')
    const style = resolveCfCellStyle(undefined, { properties }, colors)
    expect(style.fillColor).toBe('FFFFC7CE')
  })

  it('accepts color auto="1" and benign font children condense/extend/family/scheme/charset (Ruling C2)', () => {
    const { properties, unsupported } = parseCfDxf(
      dxfNode(
        `<font><b val="1"/><family val="2"/><scheme val="minor"/><charset val="1"/><condense val="0"/><extend val="0"/><color auto="1"/></font>` +
          `<border><bottom style="thin"><color auto="1"/></bottom></border>`,
      ),
      colors,
    )
    expect(unsupported).toEqual([])
    expect(properties['font.bold']).toBe(true)
    expect(properties['border.bottom.style']).toBe('thin')
  })

  it('a dxf solid fill with bgColor paints fill (Ruling T11-B)', () => {
    const { properties, unsupported } = parseCfDxf(
      dxfNode(`<fill><patternFill patternType="solid"><bgColor rgb="FFFFC7CE"/></patternFill></fill>`),
      colors,
    )
    expect(unsupported).toEqual([])
    expect(properties['fill.pattern']).toBe('solid')
    expect(properties['fill.bg']).toBe('FFFFC7CE')
    const style = resolveCfCellStyle(undefined, { properties }, colors)
    expect(style.fillColor).toBe('FFFFC7CE')
  })

  it('preserves explicit clears versus omission', () => {
    const cleared = parseCfDxf(
      dxfNode(
        `<font><u val="none"/><strike val="0"/></font>` +
          `<fill><patternFill patternType="none"/></fill>` +
          `<border><top style="none"/><bottom style="thin"><color rgb="FF000000"/></bottom></border>`,
      ),
      colors,
    )
    expect(cleared.unsupported).toEqual([])
    expect(cleared.properties['font.underline']).toBe('none')
    expect(cleared.properties['font.strike']).toBe(false)
    expect(cleared.properties['fill.pattern']).toBe('none')
    expect(cleared.properties['border.top.style']).toBe('none')
    expect(cleared.properties['border.bottom.style']).toBe('thin')
    expect(cleared.writes.has('font.underline')).toBe(true)
    expect(cleared.writes.has('border.top')).toBe(true)
    expect(cleared.writes.has('border.bottom')).toBe(true)

    const omitted = parseCfDxf(dxfNode(`<font><b val="1"/></font>`), colors)
    expect(omitted.unsupported).toEqual([])
    expect('font.underline' in omitted.properties).toBe(false)
    expect('font.strike' in omitted.properties).toBe(false)
    expect('fill.pattern' in omitted.properties).toBe(false)
    expect('border.top.style' in omitted.properties).toBe(false)
    expect(omitted.properties['font.bold']).toBe(true)
  })

  it('rejects a mixed rule dxf: red fill plus diagonal, gradient, layout alignment', () => {
    const { properties, writes, unsupported } = parseCfDxf(
      dxfNode(
        `<font><color rgb="FFFF0000"/></font>` +
          `<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill><gradientFill><stop position="0"/><stop position="1"/></gradientFill></fill>` +
          `<border><diagonalUp val="1"/><diagonal style="thin"><color rgb="FF000000"/></diagonal></border>` +
          `<alignment horizontal="center"/><protection locked="1"/>`,
      ),
      colors,
    )
    expect(unsupported.length).toBeGreaterThan(0)
    expect(unsupported.join(' ')).toContain('gradientFill')
    expect(unsupported.join(' ')).toContain('diagonal')
    expect(unsupported.join(' ')).toContain('alignment')
    expect(unsupported.join(' ')).toContain('protection')
    // The whole dxf is unusable: callers must refuse the rule, never cherry-pick the red fill.
    expect(writes.size).toBe(0)
    expect(Object.keys(properties)).toHaveLength(0)
  })

  it('rejects each unsupported visual in isolation', () => {
    const cases: Array<[string, string]> = [
      [`<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill><gradientFill/></fill>`, 'gradientFill'],
      [`<fill><patternFill patternType="gray125"><fgColor rgb="FFFF0000"/></patternFill></fill>`, 'gray125'],
      [`<fill><patternFill patternType="lightDown"><fgColor rgb="FFFF0000"/></patternFill></fill>`, 'patternType'],
      [`<border><diagonal style="thin"/></border>`, 'diagonal'],
      [`<border diagonalUp="1"><top style="thin"/></border>`, 'diagonalUp'],
      [`<border><vertical style="thin"/></border>`, 'vertical'],
      [`<border><horizontal style="thin"/></border>`, 'horizontal'],
      [`<alignment wrapText="1"/>`, 'alignment'],
      [`<protection hidden="1"/>`, 'protection'],
      [`<extLst><ext uri="{GUID}"/></extLst>`, 'extLst'],
      [`<mystery foo="1"/>`, 'mystery'],
        [`<font><u val="heavy"/></font>`, 'underline'],
        [`<font><sz val="abc"/></font>`, 'font size'],
        [`<font><color rgb="ZZZ"/></font>`, 'color'],
    ]
    for (const [inner, marker] of cases) {
      const r = parseCfDxf(dxfNode(inner), colors)
      expect(r.unsupported.length, `${marker}: ${inner}`).toBeGreaterThan(0)
      expect(r.writes.size, `${marker}: writes must be empty`).toBe(0)
    }
  })

  it('keeps measured underline and border-none values', () => {
    for (const u of ['single', 'double', 'singleAccounting', 'doubleAccounting']) {
      const r = parseCfDxf(dxfNode(`<font><u val="${u}"/></font>`), colors)
      expect(r.unsupported).toEqual([])
      expect(r.properties['font.underline']).toBe(u)
    }
    const bare = parseCfDxf(dxfNode(`<font><u/></font>`), colors)
    expect(bare.unsupported).toEqual([])
    expect(bare.properties['font.underline']).toBe('single')
    const sideNoStyle = parseCfDxf(dxfNode(`<border><top/></border>`), colors)
    expect(sideNoStyle.unsupported).toEqual([])
    expect('border.top.style' in sideNoStyle.properties).toBe(false)
    // Degenerate empty elements carry no claim and are legacy-ignored, not refusals.
    const emptyColor = parseCfDxf(dxfNode(`<font><color/></font>`), colors)
    expect(emptyColor.unsupported).toEqual([])
    expect('font.color' in emptyColor.properties).toBe(false)
    const emptyName = parseCfDxf(dxfNode(`<font><name val=""/></font>`), colors)
    expect(emptyName.unsupported).toEqual([])
  })

  it('encodes theme, tint and indexed colours symbolically', () => {
    const theme = parseCfDxf(dxfNode(`<font><color theme="4" tint="0.5"/></font>`), colors)
    expect(theme.unsupported).toEqual([])
    expect(theme.properties['font.color']).toBe('theme:4:0.5')
    const indexed = parseCfDxf(dxfNode(`<font><color indexed="9"/></font>`), { indexed: [] })
    expect(indexed.unsupported).toEqual([])
    expect(indexed.properties['font.color']).toBe('indexed:9')
    const rgb6 = parseCfDxf(dxfNode(`<font><color rgb="9C0006"/></font>`), colors)
    expect(rgb6.unsupported).toEqual([])
    expect(rgb6.properties['font.color']).toBe('FF9C0006')
  })

  it('reads numFmt codes and maps overlong codes to General', () => {
    const ok = parseCfDxf(dxfNode(`<numFmt numFmtId="164" formatCode="0.00"/>`), colors)
    expect(ok.unsupported).toEqual([])
    expect(ok.properties['numFmt.id']).toBe(164)
    expect(ok.properties['numFmt.code']).toBe('0.00')
    expect(ok.writes.has('numFmt')).toBe(true)
    const long = parseCfDxf(dxfNode(`<numFmt numFmtId="165" formatCode="${'x'.repeat(300)}"/>`), colors)
    expect(long.unsupported).toEqual([])
    expect(long.properties['numFmt.id']).toBe(0)
    expect('numFmt.code' in long.properties).toBe(false)
  })

  it('parses inline x14:dxf by local name with the same strictness', () => {
    const ok = parseCfDxf(x14DxfNode(`<x14:font><x14:b val="1"/></x14:font>`), colors)
    expect(ok.unsupported).toEqual([])
    expect(ok.properties['font.bold']).toBe(true)
    const bad = parseCfDxf(
      x14DxfNode(`<x14:fill><x14:patternFill patternType="solid"/><x14:gradientFill/></x14:fill>`),
      colors,
    )
    expect(bad.unsupported.length).toBeGreaterThan(0)
    expect(bad.writes.size).toBe(0)
  })

  it('loads referenced dxfs on a CF-only sheet with no table, pivot or slicer metadata', () => {
    const styles = parseXml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<numFmts count="0"/><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
        `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
        `<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
        `</styleSheet>`,
    )
    const dxfs = getChildren(styles, 'dxfs')[0]
    expect(dxfs).toBeDefined()
    // No tableStyles, no pivot, no slicer parts anywhere in this styles part.
    expect(getChildren(styles, 'tableStyles')).toHaveLength(0)
    const dxf = getChildren(dxfs, 'dxf')[0]!
    const r = parseCfDxf(dxf, {})
    expect(r.unsupported).toEqual([])
    expect(r.properties['fill.fg']).toBe('FFFFC7CE')
  })
})


describe('CF strict dxf review-t3-r1 findings', () => {
  it('F8 refuses unrepresentable sizes, styles, themes and booleans without cherry-picking', () => {
    const cases: Array<[string, string]> = [
      [
        `<font><sz val="Infinity"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>`,
        'infinite size with red fill',
      ],
      [`<font><sz val="-2"/></font>`, 'negative size'],
      [`<border><left style="bogus"/></border>`, 'unknown border style'],
      [`<border><top style="THIN"/></border>`, 'wrong-case border style'],
      [`<font><color theme="999"/></font>`, 'unresolvable theme index'],
      [`<font><color theme="4" tint="5"/></font>`, 'out-of-range tint'],
      [`<font><color theme="4" tint="abc"/></font>`, 'non-numeric tint'],
      [`<font><b val="bogus"/></font>`, 'non-boolean bold'],
      [`<font><strike val="yes"/></font>`, 'non-boolean strike'],
      [`<font><italic val="maybe"/></font>`, 'non-boolean italic'],
    ]
    for (const [inner, marker] of cases) {
      const r = parseCfDxf(dxfNode(inner), colors)
      expect(r.unsupported.length, marker).toBeGreaterThan(0)
      expect(r.writes.size, `${marker}: writes must be empty`).toBe(0)
      expect(Object.keys(r.properties), `${marker}: properties must be empty`).toHaveLength(0)
    }
  })

  it('F8 keeps lexical booleans and bounded themes working', () => {
    for (const [val, expected] of [['1', true], ['true', true], ['TRUE', true], ['0', false], ['false', false], ['FALSE', false]] as const) {
      const r = parseCfDxf(dxfNode(`<font><b val="${val}"/></font>`), colors)
      expect(r.unsupported).toEqual([])
      expect(r.properties['font.bold']).toBe(expected)
    }
    const maxTheme = parseCfDxf(dxfNode(`<font><color theme="11" tint="-1"/></font>`), colors)
    expect(maxTheme.unsupported).toEqual([])
    expect(maxTheme.properties['font.color']).toBe('theme:11:-1')
  })
})

describe('CF fix round 2: review-t3-r2 residuals', () => {
  it('F6 accepts every renderer-supported border style and rejects the rest', () => {
    // Mirrors src/xlsx/render.ts BORDER_STYLES plus the explicit none clear.
    const supported = [
      'none',
      'thin',
      'hair',
      'dotted',
      'dashed',
      'dashDot',
      'dashDotDot',
      'medium',
      'mediumDashed',
      'mediumDashDot',
      'mediumDashDotDot',
      'slantDashDot',
      'thick',
      'double',
    ]
    for (const style of supported) {
      const r = parseCfDxf(dxfNode(`<border><left style="${style}"/></border>`), colors)
      expect(r.unsupported, style).toEqual([])
      expect(r.properties['border.left.style'], style).toBe(style)
      expect(r.writes.has('border.left'), style).toBe(true)
    }
    for (const bad of ['bogus', 'THIN', 'MediumDashed', 'medium-dashDot']) {
      const r = parseCfDxf(dxfNode(`<border><left style="${bad}"/></border>`), colors)
      expect(r.unsupported.length, JSON.stringify(bad)).toBeGreaterThan(0)
      expect(r.writes.size).toBe(0)
      expect(Object.keys(r.properties)).toHaveLength(0)
    }
  })
})
