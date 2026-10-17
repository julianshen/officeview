import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createScrollWindow, logicalToViewport } from '../src/xlsx/view/scrollWindow'

async function workbook() {
  const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'Join', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }, { r: 2, cells: [{ ref: 'B2', v: 2 }] }, { r: 12, cells: [{ ref: 'C12', v: 12 }] }] }]))
  const path = 'xl/worksheets/sheet1.xml'
  let xml = await zip.file(path)!.async('string')
  xml = xml.replace('<sheetData>', '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="2" topLeftCell="D12" state="frozen"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10"/><cols><col min="1" max="1" width="26" customWidth="1"/></cols><sheetData>')
    .replace('<row r="1">', '<row r="1" ht="12.75">')
    .replace('<row r="2">', '<row r="2" ht="30">')
  expect(xml).toContain('<row r="1" ht="12.75">')
  expect(xml).toContain('<row r="2" ht="30">')
  zip.file(path, xml)
  return zip.generateAsync({ type: 'uint8array' })
}

it.each(['resident', 'paged'] as const)('%s frozen geometry preserves Excel non64 defaults and stored heights', async kind => {
  const bytes = await workbook()
  let source
  if (kind === 'resident') {
    const doc = await parseXlsx(await OfficePackage.load(bytes))
    source = createObjectWorkbookSource(doc, 'join-resident')
  } else {
    const session = await openImportSession({ kind: 'blob', blob: new Blob([Uint8Array.from(bytes)]) }, { docId: 'join-paged', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    source = createPagedWorkbookSource(session)
  }
  try {
    const g = source.geometry('1')
    // Calibri11 MDW7: explicit26chars=182px; baseColWidth10 ->80px defaults.
    expect([g.cols.size(0), g.cols.size(1), g.cols.start(2), g.cols.start(3)]).toEqual([182, 80, 262, 342])
    // Stored ht applies without customHeight:12.75pt=17px,30pt=40px.
    expect([g.rows.size(0), g.rows.size(1), g.rows.start(2)]).toEqual([17, 40, 57])
    expect((g as typeof g & { frozen?: unknown }).frozen).toEqual({ rows: 2, cols: 2, topRow: 11, leftCol: 3 })
    const visibility = visibilityFromHidden(createIntervalSet([]))
    const m = createSheetMetrics({ rows: g.rows, cols: g.cols, order: createIdentityOrder(), visibility })
    const window = createScrollWindow(m, { width: 800, height: 400 }, { frozen: (g as typeof g & { frozen?: unknown }).frozen, offset: { x: 100, y: 200 } } as never)
    expect((window as typeof window & { frozen?: unknown }).frozen).toEqual({ rows: 2, cols: 2, left: 262, top: 57 })
    expect(logicalToViewport(window, { x: 362, y: 257 })).toEqual({ x: 262, y: 57 })
  } finally { await source.release() }
})
