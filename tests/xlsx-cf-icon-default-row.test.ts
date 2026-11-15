// @vitest-environment node
/**
 * Task 28 round 5, Ruling S4-DEFAULT-ROW: a row without `ht` takes `<sheetFormatPr defaultRowHeight>` as its authored
 * height for the icon frame (S4-ROW-EXACT), read by the one `sheetFormatPr` parser both adapters share
 * (`columnWidth.ts` `sheetFormat`). `zeroHeight` hides such rows in Excel and `thickTop`/`thickBottom` change their
 * height, so either withholds; `customHeight` only marks the value as set by hand. Absent: withheld.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { defaultRowPtOf, sheetFormat } from '../src/xlsx/columnWidth'
import { cfRowPtOfSheet } from '../src/xlsx/cf/render'
import { cfRowPtOfOverrides } from '../src/xlsx/cf/text'

describe('S4-DEFAULT-ROW: the default row height', () => {
  it('sheetFormat reads defaultRowHeight and the flags; defaultRowPtOf admits only a plain default', () => {
    expect(sheetFormat({ defaultRowHeight: '15', customHeight: '1', zeroHeight: '0', thickTop: 'false' })).toMatchObject({ defaultRowHeight: 15, customHeight: true, zeroHeight: false, thickTop: false })
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: '15' }))).toBe(15)
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: '15', customHeight: '1' }))).toBe(15)
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: '15', zeroHeight: '1' }))).toBeUndefined()
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: '15', thickTop: '1' }))).toBeUndefined()
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: '15', thickBottom: 'true' }))).toBeUndefined()
    expect(defaultRowPtOf(sheetFormat({}))).toBeUndefined()
    expect(defaultRowPtOf(sheetFormat({ defaultRowHeight: 'x' }))).toBeUndefined()
  })

  async function workbook(format: string | undefined) {
    const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'D', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }, { r: 2, cells: [{ ref: 'A2', v: 2 }] }] }]))
    const path = 'xl/worksheets/sheet1.xml'
    let xml = await zip.file(path)!.async('string')
    if (format) xml = xml.replace('<sheetData>', `${format}<sheetData>`)
    xml = xml.replace('<row r="2">', '<row r="2" ht="30">')
    zip.file(path, xml)
    return zip.generateAsync({ type: 'uint8array' })
  }

  it.each(['resident', 'paged'] as const)('%s: rows without ht take defaultRowHeight; rows with ht keep it; absent withholds', async kind => {
    for (const [format, want] of [['<sheetFormatPr defaultRowHeight="15"/>', 15], ['<sheetFormatPr baseColWidth="10"/>', undefined], [undefined, undefined], ['<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/>', undefined]] as const) {
      const bytes = await workbook(format)
      let source
      if (kind === 'resident') {
        const doc = await parseXlsx(await OfficePackage.load(bytes))
        const sheet = doc.sheets[0]!
        expect(sheet.defaultRowHeightPt, `${format}`).toBe(want)
        const rowPt = cfRowPtOfSheet(sheet)
        expect([rowPt(0), rowPt(1), rowPt(7)], `${format}`).toEqual([want, 30, want])
        source = createObjectWorkbookSource(doc, 'd-resident')
      } else {
        const session = await openImportSession({ kind: 'blob', blob: new Blob([Uint8Array.from(bytes)]) }, { docId: 'd-paged', profile: 'D' })
        expect((await session.done).kind).toBe('complete')
        source = createPagedWorkbookSource(session)
      }
      try {
        const g = source.geometry('1')
        expect(g.defaultRowPt, `${kind} ${format}`).toBe(want)
        const rowPt = cfRowPtOfOverrides(g.rowOverrides, g.defaultRowPt)
        expect([rowPt(0), rowPt(1), rowPt(7)], `${kind} ${format}`).toEqual([want, 30, want])
      } finally { await source.release() }
    }
  })
})
