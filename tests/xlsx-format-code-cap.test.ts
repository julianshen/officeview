/**
 * Excel's own limit on a custom number-format code is 255 UTF-16 units. Longer
 * authored codes are treated as General on every path (static render, the
 * interactive viewport and the paged/worker import); a 255-unit code is
 * honoured. Paint selects the text ('@') section from the formatter's own
 * report and the section split stops at the fourth unquoted ';', so even a
 * directly constructed long code builds at most four sections.
 */
import { describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXml } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'
import { cellDisplayText, renderSheet } from '../src/xlsx/render'
import { workbookFormatting } from '../src/xlsx/storage/pagedFormatting'
import { createCellLayoutHost } from '../src/xlsx/view/renderViewport'
import { LayoutCache, type CellLayoutHost, type LayoutAdmission } from '../src/xlsx/view/textLayout'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxCellStyle, XlsxSheet } from '../src/xlsx/types'

// '0' + '"' + n × 'x' + '"': 255 units at n = 252, 256 at n = 253.
const code = (units: number) => `0"${'x'.repeat(units - 3)}"`
const CODE_255 = code(255), CODE_256 = code(256)
const escape = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
const stylesXml = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="${escape(CODE_255)}"/><numFmt numFmtId="165" formatCode="${escape(CODE_256)}"/><numFmt numFmtId="14" formatCode="${escape(code(300))}"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`
const workbookXml = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'

async function workbook(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
  zip.file('xl/workbook.xml', workbookXml)
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>3</v></c><c r="B1" s="2"><v>1234.5</v></c><c r="C1" s="3"><v>1234.5</v></c></row></sheetData></worksheet>')
  zip.file('xl/styles.xml', stylesXml)
  return zip.generateAsync({ type: 'uint8array' })
}

const owner: CacheIdentity = { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }
function admission(limit = Infinity) {
  let held = 0
  const api: LayoutAdmission = {
    admit: bytes => { if (held + bytes > limit) return false; held += bytes; return true },
    release: bytes => { held -= bytes },
    hold: bytes => { held += bytes },
    cacheChanged: () => {},
  }
  return api
}
/** Paints one cell through the admitted LayoutCache; `formatted` counts the
 * bytes admitted by the formatting observer alone (not the record bound). */
function paint(sheet: XlsxSheet, cell: XlsxCell, width = 4000, gate = admission(), formatted = { bytes: 0 }) {
  const inner = createCellLayoutHost(createCanvas(400, 40).getContext('2d') as unknown as CanvasRenderingContext2D, sheet)
  const host: CellLayoutHost = {
    ...inner,
    paintedText: (target, w, outer) => inner.paintedText!(target, w, { ...outer, admit: bytes => { formatted.bytes += bytes; return outer.admit(bytes) } }),
  }
  const cache = new LayoutCache()
  cache.beginBatch(gate)
  try { return cache.layout(owner, { cell, source: { row: cell.row, col: cell.col }, x: 0, y: 0, width, height: 20 }, host) }
  finally { cache.endBatch() }
}

describe('custom number-format codes longer than 255 UTF-16 units are General', () => {
  it('static: the parsed style keeps a 255-unit code and drops a longer one to General', async () => {
    const doc = await parseXlsx(await OfficePackage.load(await workbook()))
    const sheet = doc.sheets[0]
    const [a1, b1, c1] = sheet.rows[0].cells
    expect(a1.style?.formatCode).toBe(CODE_255)
    expect(cellDisplayText(sheet, a1)).toBe(`3${'x'.repeat(252)}`)
    expect(b1.style).toMatchObject({ numFmtId: 0 })
    expect(b1.style?.formatCode).toBeUndefined()
    expect(cellDisplayText(sheet, b1)).toBe('1234.5')
    // An overlong code that overrides a built-in id is General too, not the built-in date.
    expect(c1.style).toMatchObject({ numFmtId: 0 })
    expect(cellDisplayText(sheet, c1)).toBe('1234.5')
    sheet.cols = [{ min: 0, max: 2, widthChars: 18 }]
    const ctx = createCanvas(400, 40).getContext('2d') as unknown as CanvasRenderingContext2D
    const texts: string[] = []
    const fill = ctx.fillText.bind(ctx)
    ctx.fillText = (text: string, x: number, y: number) => { texts.push(text); fill(text, x, y) }
    renderSheet(sheet, ctx)
    expect(texts.filter(text => text === '1234.5')).toHaveLength(2)
  })

  it('interactive: the admitted viewport paints the parsed styles the same way', async () => {
    const doc = await parseXlsx(await OfficePackage.load(await workbook()))
    const sheet = doc.sheets[0]
    const [a1, b1] = sheet.rows[0].cells
    expect(paint(sheet, a1).formattedText).toBe(`3${'x'.repeat(252)}`)
    expect(paint(sheet, b1, 200).formattedText).toBe('1234.5')
  })

  it('worker: the paged import style table applies the same limit', async () => {
    const formatting = await workbookFormatting(parseXml(workbookXml), stylesXml)
    expect(formatting.styles[1]).toMatchObject({ numFmtId: 164, formatCode: CODE_255 })
    expect(formatting.styles[2]).toMatchObject({ numFmtId: 0 })
    expect(formatting.styles[2].formatCode).toBeUndefined()
    expect(formatting.styles[3]).toMatchObject({ numFmtId: 0 })
    const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
    expect(cellDisplayText(sheet, { value: 1234.5, style: formatting.styles[2] })).toBe('1234.5')
    expect(cellDisplayText(sheet, { value: 3, style: formatting.styles[1] })).toBe(`3${'x'.repeat(252)}`)
  })
})

describe('paint selects the text section without re-parsing the code', () => {
  const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
  const cell = (value: number, style: XlsxCellStyle): XlsxCell => ({ ref: 'A1', row: 0, col: 0, value, styleIndex: 0, style })

  it('an 8,192-semicolon code builds at most four sections on the admitted paint path', () => {
    const formatted = { bytes: 0 }
    const body = paint(sheet, cell(3, { numFmtId: 164, formatCode: `0${';'.repeat(8192)}` }), 400, admission(), formatted)
    expect(body.formattedText).toBe('3')
    // Four 8-byte section slots, not 8,193 (65,544 bytes): formatting stays small.
    expect(formatted.bytes).toBeGreaterThan(0)
    expect(formatted.bytes).toBeLessThan(4096)
  })

  it('the conditional variant ([>0]0 + 8,192 semicolons) is bounded the same way', () => {
    const formatted = { bytes: 0 }
    const body = paint(sheet, cell(3, { numFmtId: 164, formatCode: `[>0]0${';'.repeat(8192)}` }), 400, admission(), formatted)
    expect(body.formattedText).toBe('3')
    expect(formatted.bytes).toBeLessThan(4096)
  })

  it('the section check before the overflow fill is observed by the same admission', () => {
    const style: XlsxCellStyle = { numFmtId: 164, formatCode: `0;-"${'q'.repeat(65_536)}"0` }
    let display = 0
    const text = cellDisplayText(sheet, { value: -3, style }, { allocate: bytes => { display += bytes }, release: () => {} })
    const formatted = { bytes: 0 }
    expect(paint(sheet, cell(-3, style), 400, admission(), formatted).formattedText).toMatch(/^#+$/)
    expect(text).toHaveLength(65_538)
    // The paint-side section check copies the quoted literal again (at least
    // 2 bytes per unit): those copies are admitted, never made unobserved.
    expect(formatted.bytes - display).toBeGreaterThanOrEqual(2 * 65_536)
  })

  it('a number shown through an @ section keeps its text when it does not fit (preformatted and inline)', () => {
    for (const formatCode of ['@', '"n"@']) {
      const body = paint(sheet, cell(123456789, { numFmtId: 49, formatCode }), 20)
      expect(body.formattedText).not.toMatch(/^#+$/)
    }
    // A numeric section of the same code still overflows to '#'.
    expect(paint(sheet, cell(123456789, { numFmtId: 164, formatCode: '0;@' }), 20).formattedText).toMatch(/^#+$/)
  })
})
