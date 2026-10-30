// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test } from 'vitest'
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
    expect(result.records.some(record => record.source.row === 2 && record.source.col === 2)).toBe(true)
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

  test('alternating stripe sizes default independently and far-row roles use constant work', async () => {
    await tableStyleLoader.preloadTableStyles()
    const { tableCellRoles } = await import('../src/xlsx/tableStyles/regionLayer')
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    sheet.tables![0]!.extent = { ...sheet.tables![0]!.extent, rows: 100_000 }
    const source = sheet.tableStyleSource!
    const root = parseXml(`<styleSheet><dxfs/><tableStyles><tableStyle name="OVTest" table="1">
      <tableStyleElement type="firstRowStripe" dxfId="0" size="2"/>
      </tableStyle></tableStyles></styleSheet>`)
    sheet.tableStyleSource = { ...source, root }
    const layer = regionStylesFor(sheet)!
    const start = performance.now()
    for (let row = 98_000; row < 99_000; row++) layer.styleAt(row, 1, undefined)
    const elapsed = performance.now() - start
    expect(elapsed).toBeLessThan(50)

    const table = sheet.tables![0]!
    for (const row of [98_001, 98_002, 98_997, 98_998]) {
      const actual = tableCellRoles(table, row, 1, { elements: { firstRowStripe: { type: 'firstRowStripe', dxfId: 0, size: 2 } } } as never)
        .find(role => typeof role === 'object' && (role.element === 'firstRowStripe' || role.element === 'secondRowStripe'))
      let offset = 0, size = 2
      const index = row - (table.extent.firstRow + table.headerRowCount)
      while (offset + size <= index) { offset += size; size = size === 2 ? 1 : 2 }
      const expected = Math.floor((index - offset) / size) === 0 && size === 2 ? 'firstRowStripe' : 'secondRowStripe'
      expect(typeof actual === 'object' ? actual.element : actual).toBe(expected)
    }
  })

  test('styleAt over a 16,000 by 10 table scales linearly past the memo cap', async () => {
    await tableStyleLoader.preloadTableStyles()
    const { buildTableRegions } = await import('../src/xlsx/tableStyles/regionLayer')
    const { parseStylesExtras } = await import('../src/xlsx/tableStyles/parseStylesExtras')
    const { parseThemeContext } = await import('../src/drawing/style')
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    const extras = parseStylesExtras(sheet.tableStyleSource!.root, sheet.tableStyleSource!.colors)
    const original = sheet.tables![0]!
    const measure = (rows: number) => {
      const table = { ...original, extent: { ...original.extent, rows, cols: 10 } }
      const layer = buildTableRegions([table], extras, parseThemeContext())!
      const started = performance.now()
      for (let row = 0; row < rows; row++) for (let col = 0; col < 10; col++) layer.styleAt(row, col, undefined)
      return performance.now() - started
    }
    const small = measure(4_000)
    const large = measure(16_000)
    expect(large).toBeLessThan(4 * small + 500)
    expect(large).toBeLessThan(3_000)
  })

  test('a 20,000 by 10 table keeps its memo within the documented per-layer bound', async () => {
    await tableStyleLoader.preloadTableStyles()
    const { buildTableRegions, REGION_STYLE_LAYER_MAX_BYTES } = await import('../src/xlsx/tableStyles/regionLayer')
    const { parseStylesExtras } = await import('../src/xlsx/tableStyles/parseStylesExtras')
    const { parseThemeContext } = await import('../src/drawing/style')
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc), extras = parseStylesExtras(sheet.tableStyleSource!.root, sheet.tableStyleSource!.colors)
    const table = { ...sheet.tables![0]!, extent: { ...sheet.tables![0]!.extent, rows: 20_000, cols: 10 } }
    const layer = buildTableRegions([table], extras, parseThemeContext())!
    for (let row = 0; row < 20_000; row++) for (let col = 0; col < 10; col++) layer.styleAt(row, col, undefined)
    expect(layer.retainedBytes).toBeLessThanOrEqual(REGION_STYLE_LAYER_MAX_BYTES)
  })

  test('metrics do not create a diagnostics array without a diagnostic', async () => {
    await tableStyleLoader.preloadTableStyles()
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    sheet.diagnostics = undefined
    computeMetrics(sheet)
    expect(sheet.diagnostics).toBeUndefined()
  })

  test('a resolved styled table is clipped to static grid limits without refusal', async () => {
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    const original = sheet.tables![0]!
    original.extent = { ...original.extent, rows: 1_000_000, cols: 1 }
    await tableStyleLoader.preloadTableStyles()
    const metrics = computeMetrics(sheet)
    expect(metrics.rowHeightsPx).toHaveLength(16_384)
    expect(metrics.exportLimits?.some(limit => limit.resource === 'rows') ?? false).toBe(false)
    expect(metrics.diagnostics).toContainEqual(expect.objectContaining({ code: 'table-style-region-clipped' }))
    const canvas = createCanvas(100, 100)
    expect(() => renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics, undefined, undefined, { x: 0, y: 0, width: 100, height: 100 })).not.toThrow()
    const fitStyled = await parsed('table-medium-2.xlsx')
    styleSheet(fitStyled)
    const fitUnstyled = await parsed('table-medium-2.xlsx')
    fitUnstyled.sheets[0]!.tables![0]!.styleInfo = undefined
    expect(computeMetrics(fitStyled.sheets[0]!).exportLimits).toEqual(computeMetrics(fitUnstyled.sheets[0]!).exportLimits)
    expect(computeMetrics(fitStyled.sheets[0]!).exportLimits).toBeUndefined()
  })

  test('unknown style is resolved once and diagnosed once per table', async () => {
    await tableStyleLoader.preloadTableStyles()
    const doc = await parsed('table-medium-2.xlsx')
    const sheet = styleSheet(doc)
    sheet.tables![0]!.styleInfo = { name: 'NotAStyle', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false }
    expect(regionStylesFor(sheet)).toBeUndefined()
    expect(regionStylesFor(sheet)).toBeUndefined()
    expect(sheet.regionStylesResolved).toBe(true)
    expect(sheet.diagnostics?.filter(d => d.feature === 'table-style-unknown')).toHaveLength(1)
    const metrics = computeMetrics(sheet)
    const canvas = createCanvas(metrics.widthPx, metrics.heightPx)
    renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
    expect(sheet.diagnostics?.some(d => d.feature === 'table-style-engine-unavailable')).toBe(false)
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

})
