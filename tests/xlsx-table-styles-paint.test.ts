// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test, vi } from 'vitest'
import { parseXml } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { getPaintables } from '../src/render/paint'
import * as tableStyleLoader from '../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { regionStylesFor } from '../src/xlsx/tableStyleLoader'

const fixture = (name: string) => readFileSync(fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}`, import.meta.url)))
const parsed = async (name: string) => parseXlsx(await OfficePackage.load(fixture(name)))
const styleSheet = (doc: Awaited<ReturnType<typeof parsed>>, options: { header?: boolean; totals?: boolean } = {}) => {
  const sheet = doc.sheets[0]!
  const table = sheet.tables![0]!
  table.headerRowCount = options.header === false ? 0 : 1
  table.totalsRowCount = options.totals ? 1 : 0
  table.styleInfo = { name: 'OVTest', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false }
  sheet.tableStyleSource = {
    root: parseXml(`<styleSheet><dxfs>
      <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFEEEEFF"/></patternFill></fill></dxf>
      <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>
      <dxf><fill><patternFill patternType="solid"><bgColor rgb="FF0000FF"/></patternFill></fill></dxf>
      <dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>
      </dxfs><tableStyles><tableStyle name="OVTest" table="1">
      <tableStyleElement type="wholeTable" dxfId="0"/>
      <tableStyleElement type="firstRowStripe" dxfId="1" size="2"/>
      <tableStyleElement type="secondRowStripe" dxfId="2" size="2"/>
      <tableStyleElement type="headerRow" dxfId="2"/>
      <tableStyleElement type="totalRow" dxfId="3"/>
      </tableStyle></tableStyles></styleSheet>`),
    colors: {},
  }
  return sheet
}
const paint = async (doc: Awaited<ReturnType<typeof parsed>>) => {
  const pages = await getPaintables(doc)
  try {
    const canvas = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
    pages[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
    return { canvas, pages }
  } catch (error) { pages.dispose(); throw error }
}
const pixel = (canvas: ReturnType<typeof createCanvas>, x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data)

describe('table style painting', () => {
  test('renderSheet without preload stays unstyled and diagnoses unavailable engine', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    const metrics = computeMetrics(sheet)
    const canvas = createCanvas(metrics.widthPx, metrics.heightPx)
    renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
    expect(pixel(canvas, 20, 10)).not.toEqual([255, 0, 0, 255])
    expect(sheet.diagnostics).toContainEqual(expect.objectContaining({ feature: 'table-style-engine-unavailable' }))
  })

  test('paints an empty cell inside a styled table extent', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    styleSheet(doc)
    doc.sheets[0]!.rows[2]!.cells = doc.sheets[0]!.rows[2]!.cells.filter(cell => cell.col !== 2)
    const { canvas, pages } = await paint(doc)
    try {
      expect(pixel(canvas, 2 * 64 + 20, 2 * 20 + 10)).toEqual([255, 0, 0, 255])
      expect(pixel(canvas, 20, 5 * 20 + 10)).toEqual([255, 0, 0, 255])
    }
    finally { pages.dispose() }
  })

  test('viewport styling includes stored and synthetic empty table cells', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    sheet.rows[2]!.cells = sheet.rows[2]!.cells.filter(cell => cell.col !== 2)
    await tableStyleLoader.preloadTableStyles()
    const geometry = buildSheetGeometry(sheet)
    expect(geometry.kind).toBe('ready')
    const canvas = createCanvas(320, 120)
    const result = renderSheetViewport(sheet, canvas.getContext('2d') as never, geometry, { box: { x: 0, y: 0, width: 320, height: 120 }, zoom: 1, deviceScale: 1 }, {}, {
      owner: { docId: 'table-style', sheetId: '1', valueGen: 0, fontGen: 0, viewGen: 0 }, cache: new LayoutCache(),
    })
    expect(result.kind).toBe('detail')
    expect(pixel(canvas, 2 * 64 + 20, 2 * 20 + 10)).toEqual([255, 0, 0, 255])
  })

  test('built-in styles still resolve when a workbook has no styles.xml source', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = doc.sheets[0]!
    sheet.tableStyleSource = undefined
    await tableStyleLoader.preloadTableStyles()
    expect(regionStylesFor(sheet)?.styleAt(0, 0, undefined)?.fillColor).toBeDefined()
  })

  test('headerRowCount zero suppresses the header role', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc, { header: false })
    sheet.rows[0]!.cells = []
    const withTable = await paint(doc)
    try { expect(pixel(withTable.canvas, 20, 10)).toEqual([255, 0, 0, 255]) }
    finally { withTable.pages.dispose() }
    const headerDoc = await parsed('table-medium-2.xlsx')
    const headerSheet = styleSheet(headerDoc)
    headerSheet.rows[0]!.cells = []
    const header = await paint(headerDoc)
    try { expect(pixel(header.canvas, 20, 10)).toEqual([0, 0, 255, 255]) }
    finally { header.pages.dispose() }
    expect(sheet.tables![0]!.headerRowCount).toBe(0)
  })

  test('paints totals rows and size-two stripes', async () => {
    const doc = await parsed('table-totals-row.xlsx')
    const sheet = styleSheet(doc, { totals: true })
    sheet.rows.at(-1)!.cells = []
    sheet.rows[2]!.cells = []
    sheet.rows[4]!.cells = []
    const { canvas, pages } = await paint(doc)
    try {
      const sheet = doc.sheets[0]!
      const metrics = computeMetrics(sheet)
      const last = sheet.rows.length - 1
      expect(pixel(canvas, 20, Math.floor(metrics.rowHeightsPx.slice(0, last).reduce((a, b) => a + b, 0) + 10))).toEqual([0, 255, 0, 255])
      expect(pixel(canvas, 20, 2 * 20 + 10)).toEqual([255, 0, 0, 255])
      expect(pixel(canvas, 20, 4 * 20 + 10)).toEqual([0, 0, 255, 255])
    } finally { pages.dispose() }
  })

  test('a table without a style name is pixel-equal to a direct render', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    doc.sheets[0]!.tables![0]!.styleInfo = undefined
    const { canvas, pages } = await paint(doc)
    const metrics = computeMetrics(doc.sheets[0]!)
    const direct = createCanvas(metrics.widthPx, metrics.heightPx)
    renderSheet(doc.sheets[0]!, direct.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
    try { expect(canvas.toBuffer('raw')).toEqual(direct.toBuffer('raw')) }
    finally { pages.dispose() }
  })

  test('plain worksheets and unstyled tables render as before without requesting the lazy chunk', async () => {
    const spy = vi.spyOn(tableStyleLoader, 'preloadTableStyles')
    try {
      for (const unstyledTable of [true, false]) {
        const doc = await parsed('table-medium-2.xlsx')
        if (unstyledTable) doc.sheets[0]!.tables![0]!.styleInfo = undefined
        else doc.sheets[0]!.tables = undefined
        const pages = await getPaintables(doc)
        const metrics = computeMetrics(doc.sheets[0]!)
        const direct = createCanvas(metrics.widthPx, metrics.heightPx)
        renderSheet(doc.sheets[0]!, direct.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
        const canvas = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
        pages[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
        try { expect(canvas.toBuffer('raw')).toEqual(direct.toBuffer('raw')) }
        finally { pages.dispose() }
      }
      expect(spy).not.toHaveBeenCalled()
    } finally { spy.mockRestore() }
  })
})
