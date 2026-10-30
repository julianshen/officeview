// @vitest-environment node
/**
 * PR #55 round 7 P2: the viewport's styled-region gridline rule keeps the planner's
 * sub-detail aggregation. In a mixed plan (normal rows beside thousands of rows, or
 * columns, too small to show) it consults the region only along detail rows and
 * columns; the summarized runs stroke as the plan's ordinary gridlines do there. The
 * detail rows' gridline suppression and fill cover are unchanged.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { preloadTableStyles, regionStylesFor } from '../src/xlsx/tableStyleLoader'
import { belowDetail, buildSheetGeometry, planViewport } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxSheet } from '../src/xlsx/types'
import type { LogicalViewport } from '../src/xlsx/view/contracts'

const fixture = readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/table-medium-2.xlsx', import.meta.url)))
/** Sub-detail rows and columns: 2 px each at zoom 1 (detail needs 4 device px). The planner turns a frame
 * with more visible indexes than its 4,096 gridline instances wholly coarse, so a mixed plan holds at most
 * a few thousand; this is near that ceiling. */
const TINY = 3_000

/** The fixture's TableStyleMedium2 table: header and first body row/column normal, then TINY tiny ones. */
async function tinyTable(axis: 'rows' | 'cols'): Promise<XlsxSheet> {
  await preloadTableStyles()
  const doc = await parseXlsx(await OfficePackage.load(fixture))
  const sheet = doc.sheets[0]!
  const table = sheet.tables![0]!
  const rows = axis === 'rows' ? TINY + 2 : 20, cols = axis === 'cols' ? TINY + 2 : 6
  table.extent = { ...table.extent, rows, cols }
  table.columns = Array.from({ length: cols }, (_, index) => ({ id: String(index + 1), name: `C${index + 1}`, index }))
  const header = { index: 0, cells: Array.from({ length: Math.min(cols, 6) }, (_, col) => ({ ref: '', row: 0, col, value: `C${col + 1}`, styleIndex: 0 })) }
  sheet.rows = [header]
  if (axis === 'rows') for (let r = 2; r < rows; r++) sheet.rows.push({ index: r, cells: [], heightPt: 1.5 })
  else sheet.cols = [{ min: 2, max: cols - 1, widthChars: 0.3 }]
  return sheet
}

function paint(sheet: XlsxSheet, box: LogicalViewport['box']) {
  const geometry = buildSheetGeometry(sheet)
  const viewport: LogicalViewport = { box, zoom: 1, deviceScale: 1 }
  const canvas = createCanvas(box.width, box.height)
  const result = renderSheetViewport(sheet, canvas.getContext('2d') as never, geometry, viewport, {}, { cache: new LayoutCache() })
  return { canvas, result, geometry, plan: planViewport(geometry, viewport, 'D') }
}

afterEach(() => vi.restoreAllMocks())

describe.each(['rows', 'cols'] as const)('a mixed plan of normal and sub-detail %s', axis => {
  it('consults the region along detail lines only, and keeps the detail suppression', async () => {
    const sheet = await tinyTable(axis)
    const geometry = buildSheetGeometry(sheet)
    expect(geometry.kind).toBe('ready')
    if (geometry.kind !== 'ready') return
    const metrics = axis === 'rows' ? geometry.metrics.rows : geometry.metrics.cols
    expect(metrics.size(2)).toBeGreaterThan(0)
    expect(belowDetail(metrics.size(2), 1)).toBe(true)
    expect(belowDetail(metrics.size(1), 1)).toBe(false)
    // The whole table in one pane: two normal rows (columns) and every tiny one.
    const extent = metrics.start(TINY + 2)
    const box = axis === 'rows' ? { x: 0, y: 0, width: 400, height: Math.ceil(extent) + 20 } : { x: 0, y: 0, width: Math.ceil(extent) + 20, height: 400 }
    const layer = regionStylesFor(sheet)!
    const lookups = vi.spyOn(layer, 'styleAt')
    const mixed = paint(sheet, box)
    expect(mixed.plan.kind === 'detail' && mixed.plan.coarse.length).toBeGreaterThan(0)
    expect(mixed.result.kind).toBe('detail')
    // Bounded by the plan's detail lines (a handful of rows x columns), not the 3,000 summarized ones.
    // 124f6089: more than 3,000 lookups per crossing gridline.
    expect(lookups.mock.calls.length).toBeLessThan(500)

    // The detail rows (columns) paint exactly as in a pure-detail pane holding only them.
    const near = axis === 'rows' ? { x: 0, y: 0, width: 400, height: Math.floor(metrics.start(2)) } : { x: 0, y: 0, width: Math.floor(metrics.start(2)), height: 400 }
    const pure = paint(sheet, near)
    expect(pure.plan.kind === 'detail' && pure.plan.coarse.length).toBe(0)
    // Below the header row: the mixed frame's coarse notice ('Zoom in…', baseline 20) sits on the header.
    const inset = 2, top = 24
    const crop = (canvas: ReturnType<typeof createCanvas>) => Buffer.from(canvas.getContext('2d').getImageData(0, top, near.width - (axis === 'cols' ? inset : 0), near.height - top - (axis === 'rows' ? inset : 0)).data.buffer)
    expect(crop(mixed.canvas).equals(crop(pure.canvas))).toBe(true)
  }, 60_000)
})
