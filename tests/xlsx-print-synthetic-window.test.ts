// @vitest-environment node
/**
 * PR #55 P1: the static painter adds synthetic cells for empty styled table cells.
 * A print page asks renderSheet for one window; synthetic cells are created only
 * there (widened to merge anchors reaching in), never for the whole table per page.
 * The styled-region gridline rule and the region style of real cells are likewise
 * looked up only beside what the page shows.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test, vi } from 'vitest'
const { OfficePackage } = await import('../src/core/zip')
const { parseXlsx } = await import('../src/xlsx/parse')
const { computeMetrics, computePrintMetrics, renderPrintPage, renderSheet, syntheticCellWork } = await import('../src/xlsx/render')
const { preloadTableStyles, regionStylesFor } = await import('../src/xlsx/tableStyleLoader')
type XlsxSheet = import('../src/xlsx/types').XlsxSheet

const fixture = readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/table-medium-2.xlsx', import.meta.url)))

/** The fixture's TableStyleMedium2 table stretched to rows × cols, its body left empty
 * except the header row and one value in the far corner. */
async function stretchedTable(rows: number, cols: number): Promise<XlsxSheet> {
  await preloadTableStyles()
  const doc = await parseXlsx(await OfficePackage.load(fixture))
  const sheet = doc.sheets[0]!
  const table = sheet.tables![0]!
  table.extent = { ...table.extent, rows, cols }
  table.columns = Array.from({ length: cols }, (_, index) => ({ id: String(index + 1), name: `C${index + 1}`, index }))
  const header = sheet.rows.find(row => row.index === 0)!
  header.cells = Array.from({ length: cols }, (_, col) => ({ ref: '', row: 0, col, value: `C${col + 1}`, styleIndex: 0, style: { numFmtId: 0 } }))
  sheet.rows = [header, { index: rows - 1, cells: [{ ref: '', row: rows - 1, col: cols - 1, value: 1, styleIndex: 0, style: { numFmtId: 0 } }] }]
  return sheet
}

/** Indices whose band meets the open window (from, to), counted from the metrics themselves. */
function meeting(sizes: number[], from: number, to: number): number {
  let at = 0, count = 0
  for (const size of sizes) { if (at < to && at + size > from) count++; at += size }
  return count
}
const edge = (sizes: number[], index: number) => sizes.slice(0, index).reduce((a, b) => a + b, 0)

/** Counts the synthetic-cell probes the static painter makes: styleAt with no own style
 * (every real cell here carries one). */
function countSyntheticProbes(sheet: XlsxSheet) {
  const layer = regionStylesFor(sheet)!
  expect(layer).toBeDefined()
  const spy = vi.spyOn(layer, 'styleAt')
  return { probes: () => spy.mock.calls.filter(call => call[2] === undefined).length, calls: () => spy.mock.calls.length, reset: () => spy.mockClear() }
}

describe('static print pages create synthetic styled cells for their own window only', () => {
  test('a 1,000 × 1,000 empty styled table, printed one page at a time', async () => {
    const sheet = await stretchedTable(1000, 1000)
    const m = computeMetrics(sheet)
    expect(m.rowHeightsPx.length).toBeGreaterThanOrEqual(1000)
    expect(m.colWidthsPx.length).toBeGreaterThanOrEqual(1000)
    const print = computePrintMetrics(sheet, m, 96)
    const counter = countSyntheticProbes(sheet)
    for (const [pageCol, pageRow] of [[0, 0], [3, 7], [print.pagesWide - 1, print.pagesTall - 1]] as const) {
      counter.reset()
      const before = { ...syntheticCellWork }
      const x = pageCol * print.printable.width / print.scale, y = pageRow * print.printable.height / print.scale
      const pageCols = meeting(m.colWidthsPx, x, x + print.printable.width / print.scale), pageRows = meeting(m.rowHeightsPx, y, y + print.printable.height / print.scale)
      const window = pageCols * pageRows
      const canvas = createCanvas(print.paperPx.width, print.paperPx.height)
      renderPrintPage(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m, print, pageCol, pageRow)
      expect(counter.probes()).toBeGreaterThan(0)
      // Every style lookup this page makes: synthetic probes, the painted cells'
      // region style and the gridline rule's cells beside the page's lines, each
      // bounded by the window widened one band per side. c94cbe8c made 998,999
      // synthetic probes per page; 248ddb0e still made ~57k gridline lookups and
      // one per header cell.
      expect(counter.calls()).toBeLessThanOrEqual(4 * (pageCols + 2) * (pageRows + 2))
      const created = syntheticCellWork.created - before.created
      expect(created).toBeGreaterThan(0)
      expect(created).toBeLessThanOrEqual(window)
      // Only rows that gain a synthetic cell are copied: at most the two real ones.
      expect(syntheticCellWork.rowsCopied - before.rowsCopied).toBeLessThanOrEqual(2)
    }
  })
})

describe('a page window paints exactly what the full sheet paints there', () => {
  const crop = (canvas: ReturnType<typeof createCanvas>, x: number, y: number, w: number, h: number) => Buffer.from(canvas.getContext('2d').getImageData(x, y, w, h).data.buffer)

  test('top-left inside a styled merge and the banded interior', async () => {
    const sheet = await stretchedTable(40, 14)
    // An empty merge in the banded body: its anchor sits above and left of the page.
    sheet.mergeRanges = [{ minRow: 5, minCol: 3, maxRow: 12, maxCol: 7 }]
    sheet.merges = ['D6:H13']
    const m = computeMetrics(sheet)
    const full = createCanvas(m.widthPx, m.heightPx)
    renderSheet(sheet, full.getContext('2d') as unknown as CanvasRenderingContext2D, m)
    // Page windows by real edges: the first starts inside the merge (column F, row 9),
    // the second inside the banded body below it, the third at A1 with the header.
    const at = (col: number, row: number, dx: number, dy: number, width: number, height: number) => ({ x: edge(m.colWidthsPx, col) + dx, y: edge(m.rowHeightsPx, row) + dy, width, height })
    const pageOf = (vp: { x: number; y: number; width: number; height: number }) => {
      const page = createCanvas(vp.width, vp.height)
      renderSheet(sheet, page.getContext('2d') as unknown as CanvasRenderingContext2D, m, undefined, undefined, vp)
      return page
    }
    for (const vp of [at(5, 8, 31, 9, 300, 200), at(4, 15, 7, 3, 250, 240), at(0, 0, 0, 0, 200, 100)]) {
      const before = syntheticCellWork.created
      const page = pageOf(vp)
      expect({ vp, same: crop(page, 0, 0, vp.width, vp.height).equals(crop(full, vp.x, vp.y, vp.width, vp.height)) }).toEqual({ vp, same: true })
      // The window's cells plus the merge anchor reaching in from outside it.
      expect(syntheticCellWork.created - before).toBeLessThanOrEqual(meeting(m.colWidthsPx, vp.x, vp.x + vp.width) * meeting(m.rowHeightsPx, vp.y, vp.y + vp.height) + 1)
    }
    // Pages exactly on cell edges, where the region gridline rule must still consult
    // the off-page cells beside the edge lines. The only pixels allowed to differ are
    // the line crossings on the edge row and column: an off-page cell's border stroke
    // reaches half a pixel into the page there and the static cull drops that cell
    // (pre-existing since c94cbe8c; see pr55-p1-report.md, follow-up). A wrong band
    // span would differ along the edge between crossings.
    const colEdges = new Set(m.colWidthsPx.map((_, index) => edge(m.colWidthsPx, index)))
    const rowEdges = new Set(m.rowHeightsPx.map((_, index) => edge(m.rowHeightsPx, index)))
    for (const vp of [at(2, 3, 0, 0, edge(m.colWidthsPx, 9) - edge(m.colWidthsPx, 2), edge(m.rowHeightsPx, 20) - edge(m.rowHeightsPx, 3)), at(1, 9, 0, 0, edge(m.colWidthsPx, 5) - edge(m.colWidthsPx, 1), edge(m.rowHeightsPx, 39) - edge(m.rowHeightsPx, 9))]) {
      const a = pageOf(vp).getContext('2d').getImageData(0, 0, vp.width, vp.height).data
      const b = full.getContext('2d').getImageData(vp.x, vp.y, vp.width, vp.height).data
      const stray: Array<[number, number]> = []
      let differing = 0
      for (let i = 0; i < a.length; i += 4) {
        if (a[i] === b[i] && a[i + 1] === b[i + 1] && a[i + 2] === b[i + 2] && a[i + 3] === b[i + 3]) continue
        differing++
        const px = (i / 4) % vp.width, py = Math.floor(i / 4 / vp.width)
        if (!((py === 0 && colEdges.has(vp.x + px)) || (px === 0 && rowEdges.has(vp.y + py)))) stray.push([px, py])
      }
      expect({ vp, stray }).toEqual({ vp, stray: [] })
      expect(differing).toBeLessThan(vp.width / 64 + vp.height / 20 + 2)
    }
  })
})
