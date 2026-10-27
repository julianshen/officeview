import { describe, expect, it } from 'vitest'
import { parseXml } from '../src/core/xml'
import { parseSparklineGroups } from '../src/xlsx/sparklines/parseGroups'
import { parseA1Area } from '../src/xlsx/ranges'
import { resolveSheetColor, type SheetColorContext } from '../src/xlsx/sheetColors'
import { parseStyles, parseXlsx } from '../src/xlsx/parse'
import { OfficePackage } from '../src/core/zip'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xm = 'http://schemas.microsoft.com/office/excel/2006/main'
const group = (attrs = '', inner = '') => `<x14:sparklineGroup ${attrs}>${inner}<x14:sparklines><x14:sparkline><xm:f>A1:A2</xm:f><xm:sqref>G2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`
const parse = (body: string, color: SheetColorContext = {}) => {
  const diagnostics: Array<{ kind: 'unsupported' | 'parse'; feature: string; message: string }> = []
  const root = parseXml(`<root xmlns:x14="${x14}" xmlns:xm="${xm}" xmlns:a14="${x14}" xmlns:xmm="${xm}">${body}</root>`)
  const node = (root.sparklineGroups ?? root) as import('../src/core/xml').XmlNode
  return { groups: parseSparklineGroups(node, color, diagnostics), diagnostics }
}

describe('XLSX sparkline group parsing', () => {
  it('applies defaults when attributes are absent', () => {
    const { groups } = parse(group())
    expect(groups[0]).toMatchObject({ type: 'line', lineWeightPt: 0.75, displayEmptyCellsAs: 'zero', markers: false, high: false, low: false, first: false, last: false, negative: false, displayXAxis: false, displayHidden: false, dateAxis: false, rightToLeft: false, minAxisType: 'individual', maxAxisType: 'individual' })
  })

  it('accepts boolean spellings 1, 0, true and false', () => {
    const { groups } = parse(group('markers="1" high="true" low="0" first="false" last="1" negative="true" displayXAxis="0" displayHidden="true" dateAxis="1" rightToLeft="false"'))
    expect(groups[0]).toMatchObject({ markers: true, high: true, low: false, first: false, last: true, negative: true, displayXAxis: false, displayHidden: true, dateAxis: true, rightToLeft: false })
  })

  it('treats unknown boolean spellings as false with a parse diagnostic', () => {
    const { groups, diagnostics } = parse(group('markers="yes"'))
    expect(groups[0].markers).toBe(false)
    expect(groups[0].refused).toBeUndefined()
    expect(diagnostics).toMatchObject([{ kind: 'parse', feature: 'sparkline-boolean' }])
  })

  it('falls back from an invalid lineWeight with a diagnostic', () => {
    const { groups, diagnostics } = parse(group('lineWeight="0"'))
    expect(groups[0].lineWeightPt).toBe(0.75)
    expect(diagnostics).toMatchObject([{ kind: 'parse', feature: 'sparkline-line-weight' }])
  })

  it('matches alternate prefixes by namespace URI and ignores wrong-namespace names', () => {
    const body = `<a14:sparklineGroups xmlns:bad="urn:wrong">${group('type="column"').replace(/x14:/g, 'a14:').replace(/xm:/g, 'xmm:')}<bad:sparklineGroup type="stacked"/></a14:sparklineGroups>`
    const { groups } = parse(body)
    expect(groups).toHaveLength(1)
    expect(groups[0].type).toBe('column')
    expect(groups[0].sparklines[0].host).toEqual({ row: 1, col: 6 })
  })

  it('ignores manual bounds under individual/group axes and diagnoses once', () => {
    const fixtureStyle = group('minAxisType="group" maxAxisType="group" manualMin="-10" manualMax="80"')
    const { groups, diagnostics } = parse(fixtureStyle)
    expect(groups[0]).toMatchObject({ minAxisType: 'group', maxAxisType: 'group' })
    expect(groups[0].refused).toBeUndefined()
    expect(groups[0].manualMin).toBeUndefined()
    expect(groups[0].manualMax).toBeUndefined()
    expect(diagnostics).toEqual([{ kind: 'parse', feature: 'sparkline-axis-manual-ignored', message: 'manualMin/manualMax ignored unless their axis type is custom' }])
  })

  it('ignores a manual bound under an individual axis without refusing the group', () => {
    const { groups, diagnostics } = parse(group('manualMin="2"'))
    expect(groups[0].refused).toBeUndefined()
    expect(diagnostics).toHaveLength(1)
    expect(diagnostics[0].feature).toBe('sparkline-axis-manual-ignored')
  })

  it('refuses a custom axis when its manual bound is absent', () => {
    const { groups } = parse(group('minAxisType="custom"'))
    expect(groups[0].refused).toBe('sparkline-axis-nonconforming')
  })

  it('fails closed for an unknown type', () => {
    const { groups } = parse(group('type="radar"'))
    expect(groups[0].refused).toBe('sparkline-type')
  })

  it('requires sqref to be one cell', () => {
    for (const sqref of ['G2:G3', 'G2 H2']) {
      const body = group('', '').replace('<xm:sqref>G2</xm:sqref>', `<xm:sqref>${sqref}</xm:sqref>`)
      const { groups, diagnostics } = parse(body)
      expect(groups[0].sparklines[0].host).toBeUndefined()
      expect(diagnostics.length).toBeGreaterThan(0)
    }
  })

  it('diagnoses missing reference elements', () => {
    const { groups, diagnostics } = parse('<x14:sparklineGroup><x14:sparklines><x14:sparkline/></x14:sparklines></x14:sparklineGroup>')
    expect(groups[0].sparklines[0]).toMatchObject({ host: undefined, data: undefined })
    expect(diagnostics.map(item => item.feature)).toEqual(['sparkline-data-reference', 'sparkline-host-reference'])
  })

  it('parses quoted sheet names and absolute A1 references', () => {
    expect(parseA1Area("'My Sheet'!$B$2:$F$2")).toEqual({ sheet: 'My Sheet', r0: 1, c0: 1, r1: 1, c1: 5 })
  })

  it('parses the group-level date range when dateAxis is enabled', () => {
    const body = group('dateAxis="1"', '<xm:f>Sheet1!$B$1:$F$1</xm:f>').replace('<xm:f>A1:A2</xm:f>', '<xm:f>Sheet1!A1:E1</xm:f>')
    const { groups } = parse(body)
    expect(groups[0].dateRange).toEqual({ sheet: 'Sheet1', r0: 0, c0: 1, r1: 0, c1: 5 })
  })

  it('keeps date-axis staging silent for malformed or unused date ranges', () => {
    const staged = parse(group('dateAxis="1"', '<xm:f>[1]Sheet1!A1:A3</xm:f>'))
    expect(staged.groups[0].dateRange).toBeUndefined()
    expect(staged.diagnostics.some(item => item.feature === 'sparkline-date-range')).toBe(false)
    const unused = parse(group('dateAxis="0"', '<xm:f>[1]Sheet1!A1:A3</xm:f>'))
    expect(unused.diagnostics.some(item => item.feature === 'sparkline-date-range')).toBe(false)
    const { groups, diagnostics } = staged
    expect(groups[0].dateRange).toBeUndefined()
    expect(diagnostics).toEqual([])
  })

  it('refuses multi-range and external references', () => {
    for (const formula of ['Sheet1!A1:A3 Sheet1!B1:B3', '[1]Sheet1!A1:A3']) {
      const { groups, diagnostics } = parse(group().replace('<xm:f>A1:A2</xm:f>', `<xm:f>${formula}</xm:f>`))
      expect(groups[0].sparklines[0].data).toBeUndefined()
      expect(diagnostics.length).toBeGreaterThan(0)
    }
  })

  it('refuses two-dimensional and over-limit point ranges', () => {
    const twoDimensional = parse(group().replace('<xm:f>A1:A2</xm:f>', '<xm:f>A1:B2</xm:f>'))
    expect(twoDimensional.groups[0].sparklines[0].data).toBeUndefined()
    const tooMany = parse(group().replace('<xm:f>A1:A2</xm:f>', '<xm:f>A1:A16385</xm:f>'))
    expect(tooMany.groups[0].sparklines[0].data).toBeUndefined()
    expect(tooMany.diagnostics.some(item => item.feature === 'sparkline-point-limit')).toBe(true)
  })

  it('resolves theme tint and indexed colours from the supplied workbook context', () => {
    const color: SheetColorContext = { theme: { accent1: '#123456' }, indexed: ['000000', 'FFFFFF', '010203', '040506', '070809', '0A0B0C', '0D0E0F', '101112', '131415', '161718', '191A1B', '1C1D1E', 'ABCDEF'] }
    const body = group('', '<x14:colorSeries theme="4" tint="-0.5"/><x14:colorNegative indexed="12"/>')
    const { groups } = parse(body, color)
    expect(groups[0].colors.series).toBe(resolveSheetColor({ theme: '4', tint: '-0.5' }, color))
    expect(groups[0].colors.negative).toBe(resolveSheetColor({ indexed: '12' }, color))
  })

  it('exposes the indexed palette through Styles for resident sparkline colours', async () => {
    const path = resolve(process.cwd(), 'tests/fixtures/xlsx-sparklines/line-colours-indexed.xlsx')
    const pkg = await OfficePackage.load(readFileSync(path))
    const styles = await parseStyles(pkg)
    const doc = await parseXlsx(pkg)
    const group = doc.sheets[0].sparklineGroups?.[0]
    expect(group?.colors.negative).toBeDefined()
    expect(group?.colors.negative).toBe(resolveSheetColor({ indexed: '10' }, styles.colorContext))
  })

  it('leaves missing colour elements absent', () => {
    expect(parse(group()).groups[0].colors.axis).toBeUndefined()
  })

  it('attaches resident sparkline groups from the reference fixture', async () => {
    const path = resolve(process.cwd(), 'tests/fixtures/xlsx-sparklines/line-basic.xlsx')
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(path)))
    expect(doc.sheets[0].sparklineGroups).toHaveLength(1)
  })
})
