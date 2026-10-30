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
const pivotFixture = readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/pivot-formats-override.xlsx', import.meta.url)))
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
  it('bounds its region lookups by the row and column classes, and keeps the detail suppression', async () => {
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
    // Bounded by the class pairs (header, two stripes x a few columns), not the 3,000 summarized rows.
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

/** pivot-formats-override with its location reaching row 330; rows 30-329 are 1 px (summarized at zoom 1). */
async function longPivotSheet() {
  await preloadTableStyles()
  const doc = await parseXlsx(await OfficePackage.load(pivotFixture))
  const sheet = doc.sheets.find(item => item.pivotTables?.length)!
  const def = sheet.pivotTables![0]!
  const [from] = def.location.ref.split(':')
  def.location = { ...def.location, ref: `${from}:I330` }
  const known = new Set(sheet.rows.map(row => row.index))
  for (let r = 30; r < 330; r++) if (!known.has(r)) sheet.rows.push({ index: r, cells: [], heightPt: 0.75 })
  sheet.rows.sort((a, b) => a.index - b.index)
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('geometry')
  expect(belowDetail(geometry.metrics.rows.size(100), 1)).toBe(true)
  const box = { x: 0, y: 0, width: Math.ceil(geometry.metrics.cols.start(10)) + 20, height: Math.ceil(geometry.metrics.rows.start(330)) + 10 }
  const layer = regionStylesFor(sheet)
  const lookups = layer ? vi.spyOn(layer, 'styleAt') : undefined
  const { canvas, plan } = paint(sheet, box)
  expect(plan.kind === 'detail' && plan.coarse.length).toBeGreaterThan(0)
  if (lookups) expect(lookups.mock.calls.length).toBeLessThan(500)
  const y = Math.round((geometry.metrics.rows.start(30) + geometry.metrics.rows.start(330)) / 2)
  return { at: (px: number) => Array.from(canvas.getContext('2d').getImageData(px, y, 1, 1).data), geometry }
}

describe('a summarized run inside a pivot', () => {
  it('keeps its interior gridlines hidden', async () => {
    const { at, geometry } = await longPivotSheet()
    // B|C, halfway down the summarized rows. 6906332c: the run was skipped and the line drawn through the pivot.
    const x = Math.round(geometry.metrics.cols.start(2))
    expect(at(x)).toEqual(at(x + 4))
  }, 60_000)

  it("keeps the pivot's outer edge", async () => {
    const { at, geometry } = await longPivotSheet()
    // I|J, the right edge of A3:I330: one neighbour only is inside, so the line stays (regionGrid.ts).
    // e6bec43b: hidden by rectangle containment, coarse fill [236,239,243].
    expect(at(Math.round(geometry.metrics.cols.start(9)))).toEqual([208, 208, 208, 255])
  }, 60_000)
})

describe('a long summarized run inside a banded table', () => {
  it('keeps the gridlines of its unfilled stripes and hides those under filled ones', async () => {
    const sheet = await tinyTable('rows')
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('geometry')
    const layer = regionStylesFor(sheet)!
    const box = { x: 0, y: 0, width: 400, height: Math.ceil(geometry.metrics.rows.start(TINY + 2)) + 20 }
    const lookups = vi.spyOn(layer, 'styleAt')
    const { canvas, plan } = paint(sheet, box)
    expect(plan.kind === 'detail' && plan.coarse.length).toBeGreaterThan(0)
    expect(lookups.mock.calls.length).toBeLessThan(500)
    lookups.mockRestore()
    // B|C at each summarized row's second pixel: gridline grey exactly where the row's stripe has no fill.
    const x = Math.round(geometry.metrics.cols.start(2))
    const column = canvas.getContext('2d').getImageData(x, 0, 1, box.height).data
    let grey = 0, wrong = 0
    for (let row = 2; row < TINY + 2; row++) {
      const y = Math.round(geometry.metrics.rows.start(row)) + 1
      const isGrey = column[y * 4] === 208 && column[y * 4 + 1] === 208 && column[y * 4 + 2] === 208
      const filled = Boolean(layer.styleAt(row, 2, undefined)?.fillColor)
      if (isGrey) grey++
      if (isGrey === filled) wrong++
    }
    // e6bec43b: none grey (hidden by rectangle); 124f6089 and now: the unfilled half of the 3,000 rows.
    expect(wrong).toBe(0)
    expect(grey).toBe(TINY / 2)
  }, 60_000)
})

describe('region class keys', () => {
  it('equal row and column keys resolve to the same fill, on every styled fixture sheet', async () => {
    const { readdirSync } = await import('node:fs')
    const { ensureTableStyleEngine } = await import('../src/xlsx/tableStyleLoader')
    const dir = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
    let checked = 0, keyed = 0
    for (const name of readdirSync(dir).filter(file => file.endsWith('.xlsx'))) {
      const doc = await parseXlsx(await OfficePackage.load(readFileSync(dir + name)))
      await ensureTableStyleEngine(doc.sheets)
      for (const sheet of doc.sheets) {
        const layer = sheet.regionStyles
        if (!layer) continue
        const seen = new Map<string, string>()
        for (const rect of layer.rects) for (let row = rect.minRow; row <= Math.min(rect.maxRow, rect.minRow + 60); row++) for (let col = rect.minCol; col <= Math.min(rect.maxCol, rect.minCol + 30); col++) {
          const rowKey = layer.rowKey?.(row), colKey = layer.colKey?.(col)
          checked++
          if (rowKey === undefined || colKey === undefined) continue
          keyed++
          const fill = layer.styleAt(row, col, undefined)?.fillColor ?? ''
          const key = `${rowKey}\u0000${colKey}`
          if (!seen.has(key)) seen.set(key, fill)
          expect({ name, sheet: sheet.name, row, col, fill }).toEqual({ name, sheet: sheet.name, row, col, fill: seen.get(key) })
        }
      }
    }
    expect(keyed).toBeGreaterThan(1000)
    expect(checked).toBeGreaterThanOrEqual(keyed)
  }, 120_000)
})
