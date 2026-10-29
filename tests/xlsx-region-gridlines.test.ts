// @vitest-environment node
/**
 * Excel's gridlines over styled regions (src/xlsx/regionGrid.ts; docs/validation/xlsx-pivots.md
 * "Gridlines over pivots and styled tables"), in both the static painter (page mode, export) and the
 * interactive viewport. Every expectation cites the Excel 16.114 reference it is read from
 * (tests/fixtures/xlsx-pivots/refs/*.png, 1 px per pt, gridlines #D4D4D4; ours paint #D0D0D0).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { regionGrid } from '../src/xlsx/regionGrid'
import type { XlsxSheet } from '../src/xlsx/types'

type Canvas = ReturnType<typeof createCanvas>
const fixture = (name: string) => readFileSync(fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}.xlsx`, import.meta.url)))
async function load(name: string): Promise<XlsxSheet> {
  const doc = await parseXlsx(await OfficePackage.load(fixture(name)))
  await ensureTableStyleEngine(doc.sheets)
  return doc.sheets.find(sheet => sheet.pivotTables?.length || sheet.tables?.length)!
}
const hex = (canvas: Canvas, x: number, y: number) => {
  const [r, g, b] = canvas.getContext('2d').getImageData(Math.floor(x), Math.floor(y), 1, 1).data
  return [r, g, b].map(v => v!.toString(16).padStart(2, '0')).join('').toUpperCase()
}
/** The sheet painted by both painters at 1 CSS px per px, with its column and row starts. */
function paintBoth(sheet: XlsxSheet): { painters: Array<[string, Canvas]>; x: (col: number) => number; y: (row: number) => number } {
  const m = computeMetrics(sheet)
  const x = (col: number) => m.colWidthsPx.slice(0, col).reduce((a, b) => a + b, 0)
  const y = (row: number) => m.rowHeightsPx.slice(0, row).reduce((a, b) => a + b, 0)
  const still = createCanvas(Math.ceil(m.widthPx), Math.ceil(m.heightPx))
  renderSheet(sheet, still.getContext('2d') as unknown as CanvasRenderingContext2D, m)
  const geometry = buildSheetGeometry(sheet)
  const live = createCanvas(still.width, still.height)
  const result = renderSheetViewport(sheet, live.getContext('2d') as never, geometry, { box: { x: 0, y: 0, width: live.width, height: live.height }, zoom: 1, deviceScale: 1 }, {}, { cache: new LayoutCache() })
  expect(result.kind).toBe('detail')
  return { painters: [['static', still], ['viewport', live]], x, y }
}

describe('a pivot draws no gridlines inside its location (refs/pivot-light-16.png, location A3:I22)', () => {
  test('inner column and row edges are hidden; the outer edges and the rows above keep their gridlines', async () => {
    const sheet = await load('pivot-light-16')
    const { painters, x, y } = paintBoth(sheet)
    const mid = (row: number) => y(row) + 7
    for (const [name, canvas] of painters) {
      // Ref x 170 (B|C) at y 100 (row 7, Apples, white body): #FFFFFF.
      expect(hex(canvas, x(2), mid(6)), `${name} B|C in row 7`).toBe('FFFFFF')
      // Ref x 0 and x 728 (I|J) at y 100: the location's outer edges keep their gridlines (the grid
      // ends at I here, so the left edge is sampled).
      expect(hex(canvas, x(0), mid(6)), `${name} sheet edge in row 7`).toBe('D0D0D0')
      // Ref y 105 (rows 7|8) at x 300: no line between two body rows; ref x 170 at y 20 (row 2, above): #D4D4D4.
      expect(hex(canvas, x(2) + 30, y(7)), `${name} rows 7|8`).toBe('FFFFFF')
      expect(hex(canvas, x(2), mid(1)), `${name} B|C in row 2`).toBe('D0D0D0')
    }
  })

  test('the rule needs no style engine: the location comes from the eager pivot region', async () => {
    const sheet = await load('pivot-light-16')
    const rule = regionGrid({ ...sheet, pivotTables: undefined, tables: undefined, regionStyles: undefined, regionStylesResolved: true }, () => undefined)!
    expect(sheet.pivotRegions?.[0]?.ref).toBe('A3:I22')
    expect(rule.line(6, 1, 6, 2)).toBeNull() // B7|C7, inside
    expect(rule.line(6, 8, 6, 9)).toBeUndefined() // I7|J7, the right edge
    expect(rule.line(1, 1, 2, 1)).toBeUndefined() // B2|B3, the top edge
  })
})

describe('a styled region cell\'s fill covers the gridlines on its edges', () => {
  test('columns: covered in striped rows, drawn in white rows (refs/table-medium-2.png x 53)', async () => {
    const sheet = await load('table-medium-2')
    const { painters, x, y } = paintBoth(sheet)
    for (const [name, canvas] of painters) {
      const band = hex(canvas, x(1) - 2, y(1) + 3)
      expect(band, `${name} band fill`).not.toBe('FFFFFF')
      // Ref x 53: the band fill #C0E6F5 at y 16–29 (row 2), #D4D4D4 at y 31–44 (row 3).
      expect(hex(canvas, x(1), y(1) + 3), `${name} A2|B2`).toBe(band)
      expect(hex(canvas, x(1), y(2) + 3), `${name} A3|B3`).toBe('D0D0D0')
    }
  })

  test('rows: a striped row\'s fill covers the line below it, a white row\'s does not (refs/table-light-1.png x 10)', async () => {
    const sheet = await load('table-light-1')
    const { painters, x, y } = paintBoth(sheet)
    for (const [name, canvas] of painters) {
      const band = hex(canvas, x(1) - 4, y(1) + 3)
      expect(band, `${name} band fill`).not.toBe('FFFFFF')
      // Ref x 10: y 16–30 is the row-2 fill #D9D9D9 including the row 2|3 line at y 30; row 3 white
      // from y 31; row 4's fill starts on its own top line at y 45.
      expect(hex(canvas, x(1) - 4, y(2)), `${name} rows 2|3`).toBe(band)
      expect(hex(canvas, x(1) - 4, y(2) + 3), `${name} row 3`).toBe('FFFFFF')
      expect(hex(canvas, x(1) - 4, y(3)), `${name} rows 3|4`).toBe(band)
    }
  })

  test('sheets without tables or pivots have no rule', async () => {
    const sheet = await load('table-light-1')
    expect(regionGrid({ ...sheet, tables: undefined, regionStyles: undefined, regionStylesResolved: true }, () => undefined)).toBeUndefined()
  })
})
