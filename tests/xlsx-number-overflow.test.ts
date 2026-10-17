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
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

const GLYPH = 7
function paint(value: XlsxCell['value'], widthPx: number, style?: Partial<XlsxCellStyle>): string[] {
  const painted: string[] = []
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, styleIndex: 0, value, ...(style ? { style: { numFmtId: 0, ...style } as XlsxCellStyle } : {}) }
  const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell] }] }
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
  it('stream F item 6: the F2 date column at Excel\'s default width is ##### by Excel\'s own width model', () => {
    // Excel default column: baseColWidth 8, Calibri 11 MDW 7 px -> 8 chars + 5 px
    // padding = 61 px, rounded up to a multiple of 8 = 64 px; the cell interior is
    // 64 - 3 - 3 = 58 px (+0.5 px fit tolerance). Built-in numFmt 14 renders the
    // ISO date (product convention), "2023-03-15": Calibri advances (1/2048 em)
    // 8 digits x 1038 + 2 hyphens x 627 = 9558 -> 9558 / 2048 x 11 pt x 96/72
    // = 68.45 px > 58.5 px (and > 64 px even with no padding), so Excel would show
    // '#' too: the simulators' ####### in column B is correct, not a bug.
    const CALIBRI: Record<string, number> = { '-': 627, '#': 1020 }
    const px = (s: string) => [...s].reduce((sum, ch) => sum + (CALIBRI[ch] ?? 1038), 0) / 2048 * 11 * 96 / 72
    expect(px('2023-03-15')).toBeCloseTo(68.45, 2)
    const painted: string[] = []
    const cell: XlsxCell = { ref: 'B2', row: 0, col: 0, styleIndex: 0, value: 45000, style: { numFmtId: 14, fontSizePt: 11 } as XlsxCellStyle }
    const sheet: XlsxSheet = { name: 'S', cols: [], merges: [], mergeRanges: [], rows: [{ index: 0, cells: [cell] }] }
    const ctx = new Proxy({ font: '', measureText: (s: string) => ({ width: px(s) }), fillText: (s: string) => painted.push(s) },
      { get: (o, k) => k in o ? o[k as keyof typeof o] : () => {} }) as unknown as CanvasRenderingContext2D
    renderSheet(sheet, ctx, { colWidthsPx: [64], rowHeightsPx: [20], widthPx: 64, heightPx: 20 })
    expect(painted).toEqual(['#'.repeat(Math.floor(58 / px('#')))])
    expect(cellDisplayText(sheet, cell)).toBe('2023-03-15')
  })
  it('built-in date and time formats', () => {
    // Built-in 14 renders the ISO date in this formatter profile (10 chars).
    expect(paint(45079, 40, { numFmtId: 14 })).toEqual([hashes(40)])
    expect(paint(45079, 80, { numFmtId: 14 })).toEqual(['2023-06-02'])
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
