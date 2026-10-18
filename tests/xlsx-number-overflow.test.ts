/**
 * S13 V4 (Excel fidelity): a number/date/time whose formatted text does not
 * fit its cell is shown as '#' repeated to fill the cell width — never clipped
 * digits and never spilled into a neighbour. General first drops decimals and
 * then tries scientific notation (formatGeneral) before falling back to '#'.
 * Text cells keep overflowing exactly as before. The accessible / find value
 * (cellDisplayText) stays the real formatted value.
 *
 * Fixed 7 px glyphs pin the measurement independently of host fonts; cell
 * padding is 3 px on each side, so a W px column fits floor((W - 6) / 7) chars.
 */
import { describe, expect, it } from 'vitest'
import { cellDisplayText, renderSheet } from '../src/xlsx/render'
import type { ResolvedSemantics } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const GLYPH = 7
function paint(value: XlsxCell['value'], widthPx: number, style?: Partial<XlsxCellStyle>, dateSystem: '1900' | '1904' = '1900'): string[] {
  const painted: string[] = []
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, styleIndex: 0, value, ...(style ? { style: { numFmtId: 0, ...style } as XlsxCellStyle } : {}) }
  const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell] }], semantics: { dateSystem, locale: 'en-US', timeZone: 'UTC', epochNowMs: 0, unicode: { version: 2, source: 'standalone-default' } } as ResolvedSemantics }
  const ctx = new Proxy({ font: '', measureText: (s: string) => ({ width: s.length * GLYPH }), fillText: (s: string) => painted.push(s) },
    { get: (o, k) => k in o ? o[k as keyof typeof o] : () => {} }) as unknown as CanvasRenderingContext2D
  renderSheet(sheet, ctx, { colWidthsPx: [widthPx], rowHeightsPx: [20], widthPx, heightPx: 20 })
  return painted
}
const hashes = (widthPx: number) => '#'.repeat(Math.max(1, Math.floor((widthPx - 6) / GLYPH)))

describe('numeric overflow renders as a full cell of #', () => {
  it('the f2 date column: yyyy-mm-dd (10 chars) in a 64 px column', () => {
    // 2023-06-02 is serial 45079 in the 1900 system.
    expect(paint(45079, 64, { formatCode: 'yyyy-mm-dd' })).toEqual([hashes(64)])
    expect(hashes(64)).toBe('########')
    expect(paint(45079, 80, { formatCode: 'yyyy-mm-dd' })).toEqual(['2023-06-02']) // 70 <= 74 fits
  })
  it('stream F item 6 (user decision): built-in 14 is the locale short date, so F2 dates fit at Excel\'s default width where Excel fits them', () => {
    // Excel default column: baseColWidth 8, Calibri 11 MDW 7 px -> 8 chars + 5 px
    // padding = 61 px, rounded up to 64 px; interior 64 - 3 - 3 = 58 px (+0.5).
    // Built-in 14 is Excel's locale short date (en-US m/d/yyyy). Calibri advances
    // (1/2048 em): digit 1038, '/' 791; at 11 pt (x 14.667 / 2048):
    // "4/5/2023" = 6 x 1038 + 2 x 791 = 7810 -> 55.93 px (fits),
    // "3/15/2023" = 63.36 px and "12/30/2023" = 70.80 px (####, as in Excel).
    const CALIBRI: Record<string, number> = { '/': 791, '#': 1020 }
    const px = (s: string) => [...s].reduce((sum, ch) => sum + (CALIBRI[ch] ?? 1038), 0) / 2048 * 11 * 96 / 72
    const paintCalibri = (value: number) => {
      const painted: string[] = []
      const cell: XlsxCell = { ref: 'B2', row: 0, col: 0, styleIndex: 0, value, style: { numFmtId: 14, fontSizePt: 11 } as XlsxCellStyle }
      const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell] }] }
      const ctx = new Proxy({ font: '', measureText: (s: string) => ({ width: px(s) }), fillText: (s: string) => painted.push(s) },
        { get: (o, k) => k in o ? o[k as keyof typeof o] : () => {} }) as unknown as CanvasRenderingContext2D
      renderSheet(sheet, ctx, { colWidthsPx: [64], rowHeightsPx: [20], widthPx: 64, heightPx: 20 })
      return { painted, text: cellDisplayText(sheet, cell) }
    }
    expect(px('4/5/2023')).toBeCloseTo(55.93, 2)
    expect(paintCalibri(45021)).toEqual({ painted: ['4/5/2023'], text: '4/5/2023' }) // 2023-04-05: fits
    const hashes = '#'.repeat(Math.floor(58 / px('#')))
    expect(paintCalibri(45000)).toEqual({ painted: [hashes], text: '3/15/2023' }) // 63.4 px > 58.5
    expect(paintCalibri(45290)).toEqual({ painted: [hashes], text: '12/30/2023' })
  })
  it('built-in date/time formats 14-22 use Excel\'s en-US locale forms (ECMA-376 18.8.30)', () => {
    const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [] }
    const text = (numFmtId: number) => cellDisplayText(sheet, { value: 45079.5625, style: { numFmtId } as XlsxCellStyle }) // 2023-06-02 13:30
    expect([14, 15, 16, 17, 18, 19, 20, 21, 22].map(text)).toEqual(['6/2/2023', '2-Jun-23', '2-Jun', 'Jun-23', '1:30 PM', '1:30:00 PM', '13:30', '13:30:00', '6/2/2023 13:30'])
  })
  it('built-in date and time formats', () => {
    // Built-in 14 renders the en-US locale short date (8 chars here).
    expect(paint(45079, 40, { numFmtId: 14 })).toEqual([hashes(40)])
    expect(paint(45079, 80, { numFmtId: 14 })).toEqual(['6/2/2023'])
    expect(paint(0.524259259259259, 30, { numFmtId: 21 })).toEqual([hashes(30)]) // "12:34:56"
  })
  it('negative numbers and percentages', () => {
    expect(paint(-12345.67, 50, { numFmtId: 4 })).toEqual([hashes(50)]) // "-12,345.67"
    expect(paint(-12345.67, 90, { numFmtId: 4 })).toEqual(['-12,345.67'])
    expect(paint(0.123456, 40, { numFmtId: 10 })).toEqual([hashes(40)]) // "12.35%" = 42 px > 34
    expect(paint(0.123456, 50, { numFmtId: 10 })).toEqual(['12.35%'])
  })
  it('General drops decimals, then tries scientific, and only then shows #', () => {
    expect(paint(3.14159265358979, 50)).toEqual(['3.1416'])
    expect(paint(123456789, 50)).toEqual(['1E+08']) // 9 digits do not fit 44 px; "1E+08" (35 px) does
    expect(paint(123456789, 30)).toEqual([hashes(30)]) // "1E+08" = 35 px > 24: nothing fits
  })
  it('a single # still appears in a column too narrow for one glyph', () => {
    expect(paint(45079, 8, { formatCode: 'yyyy-mm-dd' })).toEqual(['#'])
  })
  it('text cells keep overflowing (painted in full); fitting numbers are unchanged', () => {
    expect(paint('a long text value', 30)).toEqual(['a long text value'])
    expect(paint(42, 30)).toEqual(['42'])
  })
  it('the accessible / find value stays the real formatted value', () => {
    const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [] }
    expect(cellDisplayText(sheet, { value: 45079, style: { numFmtId: 0, formatCode: 'yyyy-mm-dd' } as XlsxCellStyle })).toBe('2023-06-02')
  })
})

describe('negative 1900 date display', () => {
  it.each([40, 160])('fills a %s px cell for built-in and authored negative dates/times', width => {
    expect(paint(-1, width, { numFmtId: 14 })).toEqual([hashes(width)])
    expect(paint(-0.5, width, { numFmtId: 20 })).toEqual([hashes(width)])
    expect(paint(-1, width, { formatCode: 'yyyy-mm-dd' })).toEqual([hashes(width)])
    expect(paint(-0.5, width, { formatCode: 'h:mm' })).toEqual([hashes(width)])
  })
  it('keeps 1904 signed dates and selected literal/numeric sections', () => {
    expect(paint(-1, 160, { numFmtId: 14 }, '1904')).toEqual(['-1/2/1904'])
    expect(paint(-1, 160, { formatCode: 'yyyy-mm-dd' }, '1904')).toEqual(['-1904-01-02'])
    expect(paint(-1, 160, { formatCode: 'yyyy-mm-dd;0' })).toEqual(['1'])
    expect(paint(-1, 160, { formatCode: '0" days"' })).toEqual(['-1 days'])
    expect(paint(-1, 160, { formatCode: '0;"hidden date yyyy"' })).toEqual(['hidden date yyyy'])
    expect(paint(-1, 160, { formatCode: '[<0]yyyy-mm-dd;0' })).toEqual([hashes(160)])
    expect(paint(-1, 160, { formatCode: '[>0]yyyy-mm-dd;0' })).toEqual(['-1'])
    expect(paint(-1, 160, { formatCode: '[Red]yyyy-mm-dd' })).toEqual([hashes(160)])
    expect(paint(-1, 160, { formatCode: '0\\d\\a\\y\\s' })).toEqual(['-1days'])
  })
})

it('admits negative-date section parsing through the caller allocation observer', async () => {
  const { negativeDateOverflow } = await import('../src/xlsx/cellFormat')
  const { dateSection } = await import('../src/xlsx/formula/format')
  const refusal = new RangeError('parser workspace refused')
  const allocation = { allocate() { throw refusal }, release() {} }
  const sem = { dateSystem: '1900', locale: 'en-US' } as ResolvedSemantics
  expect(() => negativeDateOverflow(-1, 'yyyy-mm-dd', sem, allocation)).toThrow(refusal)
  expect(() => dateSection('yyyy-mm-dd', allocation)).toThrow(refusal)
})

it('admits only retained negative-date sections and refuses oversized retained padding before building', async () => {
  const { negativeDateOverflow } = await import('../src/xlsx/cellFormat')
  const code = '0;'.repeat(2000) + '0_ '.repeat(2000)
  const refusal = new RangeError('section workspace exceeds remaining budget')
  const allocation = { allocate(bytes: number) { if (bytes > 64 * 1024) throw refusal }, release() {} }
  // E stops after four sections: the tail is never constructed.
  expect(() => negativeDateOverflow(-1, code, { dateSystem: '1900' } as ResolvedSemantics, allocation)).not.toThrow()
  expect(() => negativeDateOverflow(-1, '0_ '.repeat(3000) + ';0', { dateSystem: '1900' } as ResolvedSemantics, allocation)).toThrow(refusal)
})
