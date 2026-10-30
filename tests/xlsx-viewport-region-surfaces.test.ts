// @vitest-environment node
/**
 * PR #55 P2 (interactive counterpart of the print fix): synthetic styled surfaces —
 * empty cells of a styled table that paint a fill or border — are records. They are
 * admitted against the planner's frame and tile caps (degrading the tile the way the
 * planner degrades real cells), never enumerated in summarized sub-detail runs, and
 * an empty styled merge keeps its paint when only a covered part is in view.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { describe, expect, test } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { preloadTableStyles } from '../src/xlsx/tableStyleLoader'
import { belowDetail, buildSheetGeometry, frameBudgetFor, planViewport, type SheetGeometryResult } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxSheet } from '../src/xlsx/types'
import type { LogicalViewport } from '../src/xlsx/view/contracts'

const fixture = readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/table-medium-2.xlsx', import.meta.url)))

/** The fixture's TableStyleMedium2 table stretched to rows × cols. `fill` stores a
 * value in every body cell; otherwise the body is empty (header row only). */
async function stretchedTable(rows: number, cols: number, fill = false): Promise<XlsxSheet> {
  await preloadTableStyles()
  const doc = await parseXlsx(await OfficePackage.load(fixture))
  const sheet = doc.sheets[0]!
  const table = sheet.tables![0]!
  table.extent = { ...table.extent, rows, cols }
  table.columns = Array.from({ length: cols }, (_, index) => ({ id: String(index + 1), name: `C${index + 1}`, index }))
  const header = { index: 0, cells: Array.from({ length: cols }, (_, col) => ({ ref: '', row: 0, col, value: `C${col + 1}`, styleIndex: 0 })) }
  sheet.rows = [header, ...(fill ? Array.from({ length: rows - 1 }, (_, i) => ({ index: i + 1, cells: Array.from({ length: cols }, (_, col) => ({ ref: '', row: i + 1, col, value: 1, styleIndex: 0 })) })) : [])]
  return sheet
}

const render = (sheet: XlsxSheet, geometry: SheetGeometryResult, viewport: LogicalViewport, options: Parameters<typeof renderSheetViewport>[5] = {}) => {
  const canvas = createCanvas(Math.ceil(viewport.box.width * viewport.zoom), Math.ceil(viewport.box.height * viewport.zoom))
  const ctx = canvas.getContext('2d')
  ctx.scale(viewport.zoom, viewport.zoom)
  const result = renderSheetViewport(sheet, ctx as never, geometry, viewport, {}, { cache: new LayoutCache(), ...options })
  return { canvas, result }
}
const pixel = (canvas: ReturnType<typeof createCanvas>, x: number, y: number) => Array.from(canvas.getContext('2d').getImageData(x, y, 1, 1).data)

describe('synthetic styled surfaces obey the planner record caps', () => {
  // At zoom 0.2 a 64 x 20 cell projects to 12.8 x 4 device px: still detail, and one
  // 512-device-pixel tile (2,560 logical px) holds 40 x 128 = 5,120 cells.
  const tile: LogicalViewport = { box: { x: 0, y: 0, width: 2560, height: 2560 }, zoom: 0.2, deviceScale: 1 }

  test.each([
    ['D', false, 'more than 5000 records in one 512-device-pixel tile'],
    ['M', false, 'more than 2500 records in one 512-device-pixel tile'],
    ['D', true, 'more than 5000 cell records in the frame'],
    ['M', true, 'more than 2500 cell records in the frame'],
  ] as const)('an empty styled region over the cap degrades like stored cells (%s, budget %s)', async (profile, budgeted, message) => {
    const empty = await stretchedTable(140, 44)
    const geometry = buildSheetGeometry(empty)
    expect(geometry.kind).toBe('ready')
    if (geometry.kind !== 'ready') return
    expect(belowDetail(geometry.metrics.rows.size(5), 0.2)).toBe(false)
    expect(belowDetail(geometry.metrics.cols.size(5), 0.2)).toBe(false)
    const plan = planViewport(geometry, tile, profile)
    // The planner sees only the header row: detail, well under the cap, nothing summarized.
    expect(plan.kind).toBe('detail')
    if (plan.kind !== 'detail') return
    expect(plan.coarse).toEqual([])
    expect(plan.work.records).toBeLessThan(100)
    const { result } = render(empty, geometry, tile, { profile, ...(budgeted ? { frameBudget: frameBudgetFor(profile) } : {}) })
    // 3ebdf5c4: kind 'detail' with 5,120 records while work.records counted only the header.
    expect(result.kind).toBe('coarse')
    expect(result.records).toEqual([])
    expect(result.diagnostics).toContainEqual({ code: 'detail-limit', message })

    // The same table with every cell stored degrades the same way in the planner.
    const stored = await stretchedTable(140, 44, true)
    const storedGeometry = buildSheetGeometry(stored)
    const storedResult = render(stored, storedGeometry, tile, { profile, ...(budgeted ? { frameBudget: frameBudgetFor(profile) } : {}) }).result
    expect(storedResult.kind).toBe('coarse')
    expect(storedResult.diagnostics).toEqual(result.diagnostics)
    expect(storedResult.coarse).toEqual(result.coarse)
  })

  test('surfaces under the cap are charged to the frame budget', async () => {
    const sheet = await stretchedTable(60, 20)
    const geometry = buildSheetGeometry(sheet)
    const viewport: LogicalViewport = { box: { x: 0, y: 0, width: 640, height: 400 }, zoom: 1, deviceScale: 1 }
    const plan = planViewport(geometry, viewport, 'D')
    expect(plan.kind).toBe('detail')
    const budget = frameBudgetFor('D')
    const { result } = render(sheet, geometry, viewport, { frameBudget: budget })
    expect(result.kind).toBe('detail')
    expect(result.work.records).toBeGreaterThan(plan.kind === 'detail' ? plan.work.records : 0)
    expect(result.work.records).toBe(result.records.length)
    // 3ebdf5c4 charged only the planner's header cells.
    expect(budget.records).toBe(frameBudgetFor('D').records - result.work.records)
  })

  test('a mixed window enumerates no surfaces in its summarized sub-detail column', async () => {
    const sheet = await stretchedTable(30, 12)
    sheet.cols = [{ min: 4, max: 4, widthChars: 0.3 }]
    const geometry = buildSheetGeometry(sheet)
    expect(geometry.kind).toBe('ready')
    if (geometry.kind !== 'ready') return
    const narrow = [...Array(12).keys()].find(col => geometry.metrics.cols.size(col) > 0 && belowDetail(geometry.metrics.cols.size(col), 1))
    expect(narrow).toBeDefined()
    const viewport: LogicalViewport = { box: { x: 0, y: 0, width: 800, height: 600 }, zoom: 1, deviceScale: 1 }
    const plan = planViewport(geometry, viewport, 'D')
    expect(plan.kind === 'detail' && plan.coarse.length).toBeGreaterThan(0)
    const { result } = render(sheet, geometry, viewport)
    expect(result.kind).toBe('detail')
    expect(result.records.some(record => record.source.col === narrow! - 1)).toBe(true)
    expect(result.records.filter(record => record.source.col === narrow)).toEqual([])
  })
})

describe('an empty styled merge keeps its paint when only a covered part is visible', () => {
  test('viewports inside D6:H13 paint what a viewport holding its anchor paints there', async () => {
    const sheet = await stretchedTable(40, 14)
    sheet.mergeRanges = [{ minRow: 5, minCol: 3, maxRow: 12, maxCol: 7 }]
    sheet.merges = ['D6:H13']
    const geometry = buildSheetGeometry(sheet)
    expect(geometry.kind).toBe('ready')
    if (geometry.kind !== 'ready') return
    const left = geometry.metrics.cols.start(3), top = geometry.metrics.rows.start(5)
    const right = geometry.metrics.cols.start(8), bottom = geometry.metrics.rows.start(13)
    const whole = render(sheet, geometry, { box: { x: 0, y: 0, width: 1024, height: 512 }, zoom: 1, deviceScale: 1 })
    const fill = pixel(whole.canvas, (left + right) / 2, (top + bottom) / 2)
    expect(fill).not.toEqual([255, 255, 255, 255])
    // Only covered cells in view (F9 and around): 3ebdf5c4 left this white.
    const inside = { x: left + 140, y: top + 50, width: 100, height: 60 }
    const covered = render(sheet, geometry, { box: inside, zoom: 1, deviceScale: 1 })
    expect(covered.result.records.some(record => record.source.row === 5 && record.source.col === 3)).toBe(true)
    expect(pixel(covered.canvas, 50, 30)).toEqual(fill)
    // Two tiles split across the merge, neither holding the anchor, each match the
    // whole-view render where they overlap it.
    const split = left + 200
    for (const box of [{ x: left + 100, y: top + 40, width: split - left - 100, height: 100 }, { x: split, y: top + 40, width: right - split + 30, height: 100 }]) {
      const part = render(sheet, geometry, { box, zoom: 1, deviceScale: 1 })
      for (const [dx, dy] of [[5, 5], [box.width - 6, 50], [Math.floor(box.width / 2), box.height - 6]]) {
        expect({ box, at: [dx, dy], px: pixel(part.canvas, dx, dy) }).toEqual({ box, at: [dx, dy], px: pixel(whole.canvas, box.x + dx, box.y + dy) })
      }
    }
  })
})
