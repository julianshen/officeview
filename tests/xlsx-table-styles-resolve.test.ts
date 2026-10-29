// @vitest-environment node
/**
 * Table/pivot style model (pivots Task 5): styles.xml extras and role
 * resolution. Fixtures and colours are the Task 3 Excel 16.114 references
 * (docs/validation/xlsx-pivots.md).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, test } from 'vitest'
import { parseXml } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { parseXlsx } from '../src/xlsx/parse'
import { resolveSheetColor } from '../src/xlsx/sheetColors'
import { parseStylesExtras } from '../src/xlsx/tableStyles/parseStylesExtras'
import { orderRoles, PIVOT_STYLE_PRECEDENCE, resolveRoleStyle } from '../src/xlsx/tableStyles/resolve'
import type { TableStyleDef, TableStyleDiagnostic, TableStyleElementType } from '../src/xlsx/tableStyles/model'

const fixturePath = (name: string) => fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}`, import.meta.url))
const theme = parseThemeContext()
async function extrasOf(name: string) {
  const zip = await JSZip.loadAsync(readFileSync(fixturePath(name)))
  return parseStylesExtras(parseXml(await zip.file('xl/styles.xml')!.async('string')), {})
}
function extrasFrom(inner: string, colors = {}) {
  return parseStylesExtras(parseXml(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${inner}</styleSheet>`), colors)
}
const COLORS = JSON.parse(readFileSync(fixturePath('recipes/pivot-precedence-colors.json'), 'utf8')) as Record<string, string>

describe('parseStylesExtras', () => {
  test('a parsed custom pivot style resolves its element dxfs', async () => {
    const extras = await extrasOf('pivot-custom-style.xlsx')
    expect(extras.dxfs).toHaveLength(5)
    expect(extras.defaultTableStyle).toBe('TableStyleMedium2')
    expect(extras.defaultPivotStyle).toBe('PivotStyleLight16')
    expect(extras.defaultSlicerStyle).toBe('SlicerStyleLight1')
    const def = extras.tableStyles.get('OfficeViewPivot')!
    expect(def).toMatchObject({ name: 'OfficeViewPivot', pivot: true, table: false })
    expect(Object.keys(def.elements).sort()).toEqual(['firstColumn', 'firstRowSubheading', 'headerRow', 'totalRow', 'wholeTable'])
    expect(def.elements.headerRow!.dxf).toEqual({ font: { bold: true, color: { rgb: 'FFFFFFFF' } }, fill: { pattern: 'solid', bg: { rgb: 'FF1F4E78' } } })
    const header = resolveRoleStyle(def, ['wholeTable', 'headerRow'], undefined, theme)
    expect(header).toMatchObject({ bold: true, color: 'FFFFFFFF', fillColor: 'FF1F4E78' })
    expect(header.borders).toEqual({ left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' })
    expect(header.borderColors).toEqual({ left: 'FF7F7F7F', right: 'FF7F7F7F', top: 'FF7F7F7F', bottom: 'FF7F7F7F' })
    expect(resolveRoleStyle(def, ['firstRowSubheading'], undefined, theme)).toMatchObject({ bold: true, italic: true, fillColor: 'FFFCE4D6' })
  })

  test('indexed colours resolve against the workbook palette; theme colours stay symbolic', () => {
    const extras = extrasFrom('<colors><indexedColors>' + '<rgbColor rgb="FF000000"/>'.repeat(10) + '<rgbColor rgb="FF123456"/></indexedColors></colors>'
      + '<dxfs count="1"><dxf><font><color indexed="10"/></font><fill><patternFill><bgColor theme="4" tint="-0.25"/></patternFill></fill></dxf></dxfs>',
    { indexed: [...Array(10).fill('FF000000'), 'FF123456'] })
    expect(extras.dxfs[0]).toEqual({ font: { color: { rgb: 'FF123456' } }, fill: { bg: { theme: 4, tint: -0.25 } } })
  })

  test('x14 slicer styles read their own x14 dxfs', () => {
    const extras = extrasFrom(`<dxfs count="0"/><extLst>
      <ext uri="{46F421CA-312F-682f-3DD2-61675219B42D}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFB8CCE4"/></patternFill></fill></dxf></x14:dxfs></ext>
      <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="Mine"><x14:slicerStyle name="Mine"><x14:slicerStyleElements><x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/><x14:slicerStyleElement type="bogus" dxfId="0"/></x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext></extLst>`)
    expect(extras.defaultSlicerStyle).toBe('Mine')
    expect(extras.slicerStyles.get('Mine')).toEqual({ name: 'Mine', elements: { selectedItemWithData: { dxf: { fill: { bg: { rgb: 'FFB8CCE4' } } } } } })
  })
})

describe('resolveRoleStyle precedence', () => {
  const fill = (def: TableStyleDef, roles: TableStyleElementType[]) => resolveRoleStyle(def, orderRoles(roles, 'pivot'), undefined, theme).fillColor?.slice(2)

  test('wholeTable < firstRowStripe < firstColumn < headerRow (§18.8.40)', async () => {
    const def = (await extrasOf('measure-precedence.xlsx')).tableStyles.get('OfficeViewPrecedence')!
    expect(fill(def, ['wholeTable'])).toBe(COLORS.wholeTable)
    expect(fill(def, ['firstRowStripe', 'wholeTable'])).toBe(COLORS.firstRowStripe)
    expect(fill(def, ['firstColumn', 'firstRowStripe', 'wholeTable'])).toBe(COLORS.firstColumn)
    expect(fill(def, ['headerRow', 'firstColumn', 'firstRowStripe', 'wholeTable'])).toBe(COLORS.headerRow)
  })

  test('the order agrees with every pair Excel 16.114 showed (measure-precedence)', async () => {
    const def = (await extrasOf('measure-precedence.xlsx')).tableStyles.get('OfficeViewPrecedence')!
    const wins: Array<[TableStyleElementType, TableStyleElementType]> = [
      ['firstRowStripe', 'firstColumnStripe'], ['secondRowStripe', 'secondColumnStripe'],
      ['firstColumn', 'firstRowStripe'], ['firstColumn', 'secondRowStripe'],
      ['firstRowSubheading', 'firstColumn'], ['lastColumn', 'firstRowSubheading'],
      ['totalRow', 'lastColumn'], ['totalRow', 'firstColumn'],
      ['firstHeaderCell', 'headerRow'], ['firstColumnSubheading', 'headerRow'], ['secondColumnSubheading', 'headerRow'],
      ['lastColumn', 'firstColumnSubheading'], ['lastColumn', 'secondColumnSubheading'],
    ]
    for (const [winner, loser] of wins) {
      expect(fill(def, [loser, winner]), `${winner} > ${loser}`).toBe(COLORS[winner])
      expect(fill(def, [winner, loser, 'wholeTable']), `${winner} > ${loser}`).toBe(COLORS[winner])
    }
    for (const element of PIVOT_STYLE_PRECEDENCE.slice(1)) {
      if (COLORS[element]) expect(fill(def, ['wholeTable', element])).toBe(COLORS[element])
    }
  })
})

describe('resolveRoleStyle borders', () => {
  const def: TableStyleDef = { name: 'Banded', pivot: true, table: true, elements: {
    wholeTable: { dxf: { border: { top: { style: 'thick', color: { rgb: 'FF000000' } }, bottom: { style: 'thick' }, left: { style: 'medium' }, right: { style: 'medium' }, horizontal: { style: 'thin', color: { theme: 4 } }, vertical: { style: 'dotted' } } } },
    headerRow: { dxf: { border: { bottom: { style: 'double' }, top: { style: 'none' } } } },
  } }
  const edges = (top: boolean, bottom: boolean, left: boolean, right: boolean) => ({ element: 'wholeTable' as const, edges: { top, bottom, left, right } })

  test('an inner horizontal border becomes a top/bottom edge on interior cells only', () => {
    expect(resolveRoleStyle(def, [edges(false, false, false, false)], undefined, theme).borders).toEqual({ top: 'thin', bottom: 'thin', left: 'dotted', right: 'dotted' })
    expect(resolveRoleStyle(def, [edges(true, false, true, false)], undefined, theme).borders).toEqual({ top: 'thick', bottom: 'thin', left: 'medium', right: 'dotted' })
    expect(resolveRoleStyle(def, [edges(false, true, false, true)], undefined, theme).borders).toEqual({ top: 'thin', bottom: 'thick', left: 'dotted', right: 'medium' })
    // A bare role is a one-cell band: outer edges only, no inner lines.
    expect(resolveRoleStyle(def, ['wholeTable'], undefined, theme).borders).toEqual({ top: 'thick', bottom: 'thick', left: 'medium', right: 'medium' })
    const interior = resolveRoleStyle(def, [edges(false, false, false, false)], undefined, theme)
    expect(interior.borderColors).toEqual({ top: resolveSheetColor({ theme: '4' }), bottom: resolveSheetColor({ theme: '4' }) })
  })

  test('a higher element replaces lines it sets and an explicit none clears one', () => {
    const header = resolveRoleStyle(def, [edges(true, false, true, true), 'headerRow'], undefined, theme)
    expect(header.borders).toEqual({ bottom: 'double', left: 'medium', right: 'medium' })
    expect(header.borderColors).toBeUndefined()
  })
})

describe('resolveRoleStyle: own s= layering (measure-cell-override)', () => {
  test("the cell's applied font and fill win; what its format does not set comes from the style", async () => {
    const extras = await extrasOf('pivot-custom-style.xlsx')
    const def = extras.tableStyles.get('OfficeViewPivot')!
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(fixturePath('measure-cell-override.xlsx'))))
    const cells = doc.sheets.flatMap(s => s.rows.flatMap(r => r.cells))
    const own = cells.find(c => c.styleIndex === 4)!.style!
    const neutral = cells.find(c => c.styleIndex === 2)!.style!
    expect(own).toMatchObject({ applyFont: true, applyFill: true, bold: true, color: 'FFC00000', fillColor: 'FF92D050' })
    // Excel drew the cell's fill and bold red font over the style; the style's
    // borders (wholeTable), which the format does not set, still show.
    const resolved = resolveRoleStyle(def, ['wholeTable', 'firstRowSubheading'], own, theme)
    expect(resolved).toMatchObject({ bold: true, color: 'FFC00000', fillColor: 'FF92D050', italic: false })
    expect(resolved.borders).toEqual({ left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' })
    expect(resolved.applyFont).toBeUndefined()
    // A neutral pivot format (font 0, no fill, left-aligned) leaves the style's
    // font and fill, and keeps its own alignment.
    const header = resolveRoleStyle(def, ['wholeTable', 'headerRow'], neutral, theme)
    expect(header).toMatchObject({ bold: true, color: 'FFFFFFFF', fillColor: 'FF1F4E78', horizontal: 'left' })
    const sub = resolveRoleStyle(def, ['wholeTable', 'firstRowSubheading'], neutral, theme)
    expect(sub).toMatchObject({ bold: true, italic: true, fillColor: 'FFFCE4D6', fontSizePt: 11 })
  })

  test('own borders win per side; an own number format other than General wins', () => {
    const def: TableStyleDef = { name: 'N', pivot: true, table: true, elements: { wholeTable: { dxf: { border: { top: { style: 'thin' }, bottom: { style: 'thin' } }, numFmt: { id: 164, code: '0.0' } } } } }
    const r = resolveRoleStyle(def, ['wholeTable'], { numFmtId: 0, borders: { top: 'double' } }, theme)
    expect(r.borders).toEqual({ top: 'double', bottom: 'thin' })
    expect(r).toMatchObject({ numFmtId: 164, formatCode: '0.0' })
    expect(resolveRoleStyle(def, ['wholeTable'], { numFmtId: 4 }, theme).numFmtId).toBe(4)
  })
})

describe('resolveRoleStyle fills and fonts', () => {
  test('a custom dxf that sets font name, font size and a solid pattern fill resolves to them', () => {
    const extras = extrasFrom(`<dxfs count="1"><dxf><font><sz val="14"/><name val="Arial"/><u val="double"/><strike/></font><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor theme="4" tint="0.5"/></patternFill></fill></dxf></dxfs>
      <tableStyles count="1"><tableStyle name="Custom" pivot="0"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>`)
    const def = extras.tableStyles.get('Custom')!
    expect(def).toMatchObject({ pivot: false, table: true })
    expect(resolveRoleStyle(def, ['wholeTable'], undefined, theme)).toMatchObject({
      fontName: 'Arial', fontSizePt: 14, underline: 'double', strike: true,
      fillColor: resolveSheetColor({ theme: '4', tint: '0.5' }, { theme: theme.palette }),
    })
  })

  test('a non-solid pattern fill is painted as its foreground colour, with a diagnostic', () => {
    const extras = extrasFrom(`<dxfs count="1"><dxf><fill><patternFill patternType="lightUp"><fgColor rgb="FFFF0000"/><bgColor rgb="FF0000FF"/></patternFill></fill></dxf></dxfs>
      <tableStyles count="1"><tableStyle name="Hatched"><tableStyleElement type="headerRow" dxfId="0"/></tableStyle></tableStyles>`)
    const diagnostics: TableStyleDiagnostic[] = []
    const def = extras.tableStyles.get('Hatched')!
    expect(resolveRoleStyle(def, ['headerRow'], undefined, theme, diagnostics).fillColor).toBe('FFFF0000')
    resolveRoleStyle(def, ['headerRow'], undefined, theme, diagnostics)
    expect(diagnostics).toEqual([expect.objectContaining({ kind: 'approximate', feature: 'table-style-pattern-fill' })])
  })
})
