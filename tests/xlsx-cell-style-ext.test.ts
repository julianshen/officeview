// @vitest-environment node
/**
 * Cell-style extensions (pivots Task 5): underline, strike, font name and
 * alignment indent, parsed from styles.xml and painted by the cell painter.
 * Pixel assertions compare a cell with the extension against the same cell
 * without it, so they hold whichever fallback face the host has.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, test } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { layoutCellBody, paintCellLayout, textMeasurer } from '../src/xlsx/render'
import { cellFontString, CELL_FONT_STACK } from '../src/xlsx/cellFont'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache, type PlacedCell } from '../src/xlsx/view/textLayout'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'
import { buildXlsx } from '../src/testdata/ooxml-builders'

const fixture = (name: string) => readFileSync(fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}`, import.meta.url)))

const W = 160, H = 30
function paint(style: Partial<XlsxCellStyle>, value: string | number = 'Hamburg', sheet: Partial<XlsxSheet> = {}) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, W, H)
  const cell: XlsxCell = { ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style: { numFmtId: 0, fontSizePt: 11, ...style } }
  const body = layoutCellBody(ctx, sheet as XlsxSheet, cell, 0, 0, W, H, textMeasurer(ctx))
  paintCellLayout(ctx, body, { width: W, height: H })
  const data = ctx.getImageData(0, 0, W, H).data
  const dark = (x: number, y: number) => { const i = (y * W + x) * 4; return data[i] + data[i + 1] + data[i + 2] < 384 }
  return { body, dark }
}
/** Dark pixels per row. */
const rowInk = (p: ReturnType<typeof paint>) => Array.from({ length: H }, (_, y) => Array.from({ length: W }, (_, x) => p.dark(x, y) ? 1 : 0).reduce<number>((a, b) => a + b, 0))
const leftmostInk = (p: ReturnType<typeof paint>) => {
  for (let x = 0; x < W; x++) for (let y = 0; y < H; y++) if (p.dark(x, y)) return x
  return -1
}
const baselineOf = (p: ReturnType<typeof paint>) => {
  const t = p.body.text!
  return t.kind === 'plain' ? t.ty : t.layout.lines[0].segments[0].transform?.f ?? t.layout.lines[0].baseline
}

function stylesXml(fonts: string, xfs: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts>${fonts}</fonts>
<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders><border><left/><right/><top/><bottom/></border></borders>
<cellXfs>${xfs}</cellXfs></styleSheet>`
}
async function withStyles(styles: string, cells: number): Promise<XlsxSheet> {
  const bytes = await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: Array.from({ length: cells }, (_, i) => ({ ref: `${String.fromCharCode(65 + i)}1`, v: 1, style: i })) }] }])
  const zip = await JSZip.loadAsync(bytes)
  zip.file('xl/styles.xml', styles)
  return (await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))).sheets[0]
}

describe('styles.xml: underline, strike, font name and indent', () => {
  test('font u/strike/name and alignment indent reach the cell style; absent and off stay undefined', async () => {
    const sheet = await withStyles(stylesXml(
      '<font><sz val="11"/><name val="Calibri"/></font>'
      + '<font><u/><sz val="11"/><name val="Arial"/></font>'
      + '<font><u val="double"/><strike/><sz val="11"/><name val="Calibri"/></font>'
      + '<font><u val="singleAccounting"/><strike val="0"/><sz val="11"/></font>'
      + '<font><u val="none"/><sz val="11"/></font>',
      '<xf fontId="0"/><xf fontId="1" applyFont="1"/><xf fontId="2"><alignment horizontal="left" indent="2"/></xf>'
      + '<xf fontId="3" applyFill="1"><alignment indent="0"/></xf><xf fontId="4"/>'), 5)
    const styles = sheet.rows[0].cells.map(c => c.style!)
    expect(styles[0]).toMatchObject({ fontName: 'Calibri' })
    expect(styles[0].underline).toBeUndefined()
    expect(styles[0].strike).toBeUndefined()
    expect(styles[0].indent).toBeUndefined()
    expect(styles[0].ownFont).toBeUndefined()
    expect(styles[1]).toMatchObject({ underline: 'single', fontName: 'Arial', ownFont: true })
    expect(styles[2]).toMatchObject({ underline: 'double', strike: true, indent: 2, horizontal: 'left' })
    expect(styles[3]).toMatchObject({ underline: 'singleAccounting', ownFont: true })
    expect('applyFill' in styles[3]).toBe(false)
    expect(styles[3].strike).toBeUndefined()
    expect(styles[3].indent).toBeUndefined()
    expect(styles[4].underline).toBeUndefined()
  })

  test('measure-indent.xlsx: G1:G4 carry indent 0..3', async () => {
    const doc = await parseXlsx(await OfficePackage.load(fixture('measure-indent.xlsx')))
    const g = doc.sheets.flatMap(s => s.rows.flatMap(r => r.cells)).filter(c => c.col === 6 && c.row < 4).sort((a, b) => a.row - b.row)
    expect(g.map(c => c.style?.indent ?? 0)).toEqual([0, 1, 2, 3])
  })
})

describe('cell painter: underline, strike, indent, font name', () => {
  test('a single underline adds a line of ink just below the baseline', () => {
    const plain = paint({}), underlined = paint({ underline: 'single' })
    expect(underlined.body.text).toMatchObject({ decoration: { underline: 'single' } })
    const base = Math.round(baselineOf(plain))
    const a = rowInk(plain), b = rowInk(underlined)
    const gained = b.map((n, y) => n - a[y])
    const lineRow = gained.findIndex(n => n > 20)
    expect(lineRow).toBeGreaterThan(base)
    expect(lineRow).toBeLessThanOrEqual(base + 4)
    // Nothing above the baseline changes.
    for (let y = 0; y < base; y++) expect(b[y]).toBe(a[y])
  })

  test('a double underline paints two lines; accounting underlines span the cell', () => {
    const plain = rowInk(paint({}))
    const lines = (style: Partial<XlsxCellStyle>) => rowInk(paint(style)).map((n, y) => n - plain[y]).filter(n => n > 20)
    expect(lines({ underline: 'double' })).toHaveLength(2)
    const single = Math.max(...lines({ underline: 'single' }))
    const accounting = lines({ underline: 'singleAccounting' })
    expect(accounting).toHaveLength(1)
    expect(accounting[0]).toBeGreaterThan(single + 40)
    expect(lines({ underline: 'doubleAccounting' })).toHaveLength(2)
  })

  test('strike paints a line through the text body, above the baseline', () => {
    const plain = paint({}), struck = paint({ strike: true })
    const base = Math.round(baselineOf(plain))
    const a = rowInk(plain), b = rowInk(struck)
    const gained = b.map((n, y) => n - a[y])
    const row = gained.findIndex(n => n > 10)
    expect(row).toBeGreaterThan(base - 11)
    expect(row).toBeLessThan(base - 1)
    for (let y = base; y < H; y++) expect(b[y]).toBe(a[y])
  })

  test('underline and strike also paint for laid-out (aligned) text', () => {
    const plain = rowInk(paint({ horizontal: 'center' }))
    const both = rowInk(paint({ horizontal: 'center', underline: 'single', strike: true }))
    expect(both.map((n, y) => n - plain[y]).filter(n => n > 10)).toHaveLength(2)
  })

  test('indent moves left-aligned text right by level x maximum digit width', () => {
    const level0 = leftmostInk(paint({ horizontal: 'left' }))
    expect(leftmostInk(paint({ horizontal: 'left', indent: 1 })) - level0).toBe(7)
    expect(leftmostInk(paint({ horizontal: 'left', indent: 3 })) - level0).toBe(21)
    // The sheet's own maximum digit width (column metrics) sets the step.
    expect(leftmostInk(paint({ horizontal: 'left', indent: 2 }, 'Hamburg', { columnMetrics: { maxDigitWidthPx: 9, defaultWidthPx: 80 } })) - level0).toBe(18)
    // General alignment ignores indent (Excel indents left, right and distributed only).
    expect(leftmostInk(paint({ indent: 2 }))).toBe(level0)
  })

  test('indent moves right-aligned text left', () => {
    const rightmost = (p: ReturnType<typeof paint>) => { for (let x = W - 1; x >= 0; x--) for (let y = 0; y < H; y++) if (p.dark(x, y)) return x; return -1 }
    expect(rightmost(paint({ horizontal: 'right' })) - rightmost(paint({ horizontal: 'right', indent: 2 }))).toBe(14)
  })

  test('font name goes ahead of the cell font stack; Calibri and absent are the stack', () => {
    expect(cellFontString({ fontName: 'Arial' }, 11)).toBe(`11pt "Arial", ${CELL_FONT_STACK}`)
    expect(cellFontString({ fontName: 'Calibri' }, 11)).toBe(`11pt ${CELL_FONT_STACK}`)
    expect(cellFontString({}, 11)).toBe(`11pt ${CELL_FONT_STACK}`)
    expect(cellFontString({ fontName: 'Evil"; x' }, 11)).toBe(`11pt "Evil x", ${CELL_FONT_STACK}`)
    const plain = paint({ fontName: 'Times New Roman' })
    expect(plain.body.text).toMatchObject({ kind: 'plain', font: `11pt "Times New Roman", ${CELL_FONT_STACK}` })
    const laid = paint({ fontName: 'Times New Roman', horizontal: 'center' }).body.text
    expect(laid?.kind === 'layout' && laid.body.paragraphs[0].runs[0].fontFamily).toBe('Times New Roman')
  })

  test('the layout cache keys underline, strike, indent and font name', () => {
    const owner: CacheIdentity = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
    const host = createCellLayoutHost(createCanvas(200, 40).getContext('2d') as unknown as CanvasRenderingContext2D)
    const cache = new LayoutCache()
    const placed = (style: XlsxCellStyle): PlacedCell => ({ cell: { ref: 'A1', row: 0, col: 0, value: 'x', styleIndex: 0, style }, source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 20 })
    const base: XlsxCellStyle = { numFmtId: 0, horizontal: 'left' }
    const bodies = [base, { ...base, underline: 'single' as const }, { ...base, strike: true }, { ...base, indent: 1 }, { ...base, fontName: 'Arial' }]
      .map(style => cache.layout(owner, placed(style), host))
    expect(new Set(bodies).size).toBe(5)
    expect(cache.stats.built).toBe(5)
  })
})
