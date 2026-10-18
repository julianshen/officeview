import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { frozenPaneFrom } from '../src/xlsx/frozenPane'
import { initialOrigin } from '../src/xlsx/view/frozenPanes'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'

const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1000, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 64 }) })
const viewport = { width: 640, height: 400 }

it('row-only freeze preserves the authored horizontal scroll to G30', () => {
  const frozen = frozenPaneFrom({ ySplit: '1', state: 'frozen', topLeftCell: 'G30' })!
  const window = createScrollWindow(metrics, viewport, { frozen })
  const origin = initialOrigin(frozen, window.frozen, metrics)
  expect(origin).toEqual({ x: 6 * 64, y: 29 * 20 - 20 })
})

it('column-only freeze preserves the authored vertical scroll to G30', () => {
  const frozen = frozenPaneFrom({ xSplit: '1', state: 'frozen', topLeftCell: 'G30' })!
  const window = createScrollWindow(metrics, viewport, { frozen })
  expect(initialOrigin(frozen, window.frozen, metrics)).toEqual({ x: 6 * 64 - 64, y: 29 * 20 })
})

it('paged import ignores frozen panes belonging only to a later sheetView', async () => {
  const zip = await JSZip.loadAsync(await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', v: 1 }] }] }]))
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file('xl/worksheets/sheet1.xml', xml.replace('<sheetData>', '<sheetViews><sheetView workbookViewId="0"/><sheetView workbookViewId="1"><pane ySplit="2" topLeftCell="A3" state="frozen"/></sheetView></sheetViews><sheetData>'))
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const resident = await parseXlsx(await OfficePackage.load(bytes))
  expect(resident.sheets[0].frozen).toBeUndefined()
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'review-first-sheet-view', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = (createPagedWorkbookSource as (...args: unknown[]) => ReturnType<typeof createPagedWorkbookSource>)(session)
  expect(source.geometry('1').frozen).toEqual(resident.sheets[0].frozen)
})
