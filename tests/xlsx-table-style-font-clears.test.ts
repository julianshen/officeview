// @vitest-environment node
/**
 * PR #55 round 5 P2: an own (non-Normal) cell font that states `<u val="none"/>` or
 * `<strike val="0"/>` clears a table style's underline and strike; a font without those
 * elements leaves the style's decoration (resolve.ts: an own font wins for every property
 * it states). Plain cells paint the same either way.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { layoutCellBody, textMeasurer } from '../src/xlsx/render'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'

/** table-medium-2 restyled by a custom table style whose wholeTable is underlined and struck.
 * A2 uses a Calibri font that clears both (`u none`, `strike 0`); B2 a Calibri font without them;
 * G1, outside the table, the clearing font. */
async function decoratedTable(): Promise<Uint8Array<ArrayBuffer>> {
  const archive = await JSZip.loadAsync(readFileSync(resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/table-medium-2.xlsx')))
  const edit = async (path: string, change: (xml: string) => string) => {
    const xml = await archive.file(path)!.async('string'), next = change(xml)
    expect(next).not.toBe(xml)
    archive.file(path, next)
  }
  await edit('xl/styles.xml', xml => xml
    .replace('<fonts count="2">', '<fonts count="4">')
    .replace('</fonts>', '<font><u val="none"/><strike val="0"/><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/></font></fonts>')
    .replace('<cellXfs count="3">', '<cellXfs count="5">')
    .replace('</cellXfs>', '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>')
    .replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><font><u/><strike/></font></dxf></dxfs>')
    .replace(/<tableStyles count="0"([^/]*)\/>/, '<tableStyles count="1"$1><tableStyle name="OVDecor" pivot="0" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>'))
  await edit('xl/tables/table1.xml', xml => xml.replace('name="TableStyleMedium2"', 'name="OVDecor"'))
  await edit('xl/worksheets/sheet1.xml', xml => xml
    .replace('<c r="A2" t="inlineStr">', '<c r="A2" s="3" t="inlineStr">')
    .replace('<c r="B2" t="inlineStr">', '<c r="B2" s="4" t="inlineStr">')
    .replace('</row><row r="2">', '<c r="G1" s="3" t="inlineStr"><is><t>Plain</t></is></c></row><row r="2">'))
  return new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
}

const decoration = (sheet: XlsxSheet, cell: XlsxCell) => {
  const ctx = createCanvas(120, 20).getContext('2d') as unknown as CanvasRenderingContext2D
  const body = layoutCellBody(ctx, sheet, cell, 0, 0, 120, 20, textMeasurer(ctx))
  return body.text?.decoration ?? null
}

describe('an own font clears a table style decoration it states off', () => {
  it('resident: u none and strike 0 clear the style; an absent element keeps it', async () => {
    const input = await decoratedTable()
    const doc = await parseXlsx(await OfficePackage.load(input))
    await ensureTableStyleEngine(doc.sheets)
    const sheet = doc.sheets[0]!
    const cellAt = (row: number, col: number) => sheet.rows.find(r => r.index === row)!.cells.find(c => c.col === col)!
    const styled = (cell: XlsxCell) => ({ ...cell, style: sheet.regionStyles!.styleAt(cell.row, cell.col, cell.style, 'screen') })
    const cleared = cellAt(1, 0), absent = cellAt(1, 1)
    expect(cleared.style?.ownFont).toBe(true)
    expect(absent.style?.ownFont).toBe(true)
    // 79ff888f: the clears parse as absent and the style's underline and strike come back.
    expect(styled(cleared).style).not.toHaveProperty('underline')
    expect(styled(cleared).style).not.toHaveProperty('strike')
    expect(decoration(sheet, styled(cleared))).toBeNull()
    expect(styled(absent).style).toMatchObject({ underline: 'single', strike: true })
    expect(decoration(sheet, styled(absent))).toMatchObject({ underline: 'single', strike: true })
    // Outside the table the clearing font paints as before: no decoration.
    const plain = cellAt(0, 6)
    expect(sheet.regionStyles!.styleAt(0, 6, plain.style, 'screen') ?? plain.style).toBe(plain.style)
    expect(decoration(sheet, plain)).toBeNull()
  })

  it('paged: the worker style table carries the same clears', async () => {
    const session = await openImportSession({ kind: 'blob', blob: new Blob([await decoratedTable()]) }, { docId: 'font-clears', profile: 'D' })
    try {
      await session.done
      const styles = session.workbook.formatting.styles
      expect(styles[3]).toMatchObject({ ownFont: true, fontCleared: { underline: true, strike: true } })
      expect(styles[4]).toMatchObject({ ownFont: true })
      expect(styles[4]).not.toHaveProperty('fontCleared')
    } finally { await session.release() }
  })
})
