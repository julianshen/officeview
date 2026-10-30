// @vitest-environment node
/**
 * Spec §8 hot path (Task 14 fix round 1): a sheet with no pivot and no styled table pays one guard per paint for the
 * styled-region gridline rule (src/xlsx/regionGrid.ts), never the rule itself (its callback, arrays and
 * regionStylesFor lookup). Both painters are checked, with a styled sheet as the positive control.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { afterEach, describe, expect, test, vi } from 'vitest'

vi.mock('../src/xlsx/regionGrid', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/regionGrid')>()
  return { ...original, regionGrid: vi.fn(original.regionGrid) }
})

vi.mock('../src/xlsx/view/textLayout', async importOriginal => {
  const original = await importOriginal<typeof import('../src/xlsx/view/textLayout')>()
  return { ...original, buildCellRecords: vi.fn(original.buildCellRecords) }
})

const { regionGrid, mayHaveRegionGrid } = await import('../src/xlsx/regionGrid')
const { buildCellRecords } = await import('../src/xlsx/view/textLayout')
const { planViewport } = await import('../src/xlsx/view/viewport')
const { ensureTableStyleEngine } = await import('../src/xlsx/tableStyleLoader')
const { computeMetrics, renderSheet } = await import('../src/xlsx/render')
const { buildSheetGeometry } = await import('../src/xlsx/view/viewport')
const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
const { LayoutCache } = await import('../src/xlsx/view/textLayout')
const { OfficePackage } = await import('../src/core/zip')
const { parseXlsx } = await import('../src/xlsx/parse')
type Sheet = import('../src/xlsx/types').XlsxSheet

function paintBoth(sheet: Sheet): void {
  const m = computeMetrics(sheet)
  renderSheet(sheet, createCanvas(200, 100).getContext('2d') as unknown as CanvasRenderingContext2D, m)
  const geometry = buildSheetGeometry(sheet)
  for (const x of [0, 100]) {
    const result = renderSheetViewport(sheet, createCanvas(100, 100).getContext('2d') as never, geometry, { box: { x, y: 0, width: 100, height: 100 }, zoom: 1, deviceScale: 1 }, {}, { cache: new LayoutCache() })
    expect(result.kind).toBe('detail')
  }
}

afterEach(() => { vi.mocked(regionGrid).mockClear(); vi.mocked(buildCellRecords).mockClear() })

/** The placements the viewport painter lays out for one tile of `sheet`, and the plan's own cells. */
function tilePlacements(sheet: Sheet) {
  const geometry = buildSheetGeometry(sheet)
  const viewport = { box: { x: 0, y: 0, width: 200, height: 100 }, zoom: 1, deviceScale: 1 }
  const plan = planViewport(geometry, viewport)
  expect(plan.kind).toBe('detail')
  renderSheetViewport(sheet, createCanvas(200, 100).getContext('2d') as never, geometry, viewport, {}, { cache: new LayoutCache(), preparedPlan: plan })
  return { placed: vi.mocked(buildCellRecords).mock.calls.at(-1)![0], cells: plan.kind === 'detail' ? plan.cells : [] }
}

describe('the styled-region gridline rule costs a plain sheet only its guard', () => {
  test('no pivot, no table: neither painter builds the rule or looks up region styles for it', () => {
    const sheet: Sheet = { name: 'Plain', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 1, styleIndex: 0, style: { numFmtId: 0, fillColor: 'FFFFFF00' } }] }], cols: [], merges: [], mergeRanges: [] } as unknown as Sheet
    expect(mayHaveRegionGrid(sheet)).toBe(false)
    paintBoth(sheet)
    // regionGrid is where the rule allocates and calls regionStylesFor; behind the guard it is never reached.
    expect(regionGrid).not.toHaveBeenCalled()
  })

  test('a plain sheet lays out the plan\'s own cells: no per-tile copy (final review M1)', () => {
    const sheet: Sheet = { name: 'Plain', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 1, styleIndex: 0, style: { numFmtId: 0, fillColor: 'FFFFFF00' } }] }], cols: [], merges: [], mergeRanges: [] } as unknown as Sheet
    const { placed, cells } = tilePlacements(sheet)
    expect(cells.length).toBeGreaterThan(0)
    expect(placed).toBe(cells)
  })

  test('control: a pivot sheet lays out region-styled copies', async () => {
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/pivot-light-16.xlsx', import.meta.url)))))
    await ensureTableStyleEngine(doc.sheets)
    const { placed, cells } = tilePlacements(doc.sheets.find(s => s.pivotTables?.length)!)
    expect(placed).not.toBe(cells)
  })

  test('a sheet whose only table has no style name pays only the guard too', async () => {
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/table-medium-2.xlsx', import.meta.url)))))
    const sheet = doc.sheets.find(s => s.tables?.length)!
    sheet.tables = sheet.tables!.map(table => ({ ...table, styleInfo: undefined }))
    await ensureTableStyleEngine(doc.sheets)
    expect(sheet.regionStyles).toBeUndefined()
    expect(mayHaveRegionGrid(sheet)).toBe(false)
    paintBoth(sheet)
    expect(regionGrid).not.toHaveBeenCalled()
  })

  test('control: a pivot sheet builds the rule in both painters', async () => {
    const doc = await parseXlsx(await OfficePackage.load(readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/pivot-light-16.xlsx', import.meta.url)))))
    await ensureTableStyleEngine(doc.sheets)
    const sheet = doc.sheets.find(s => s.pivotTables?.length)!
    expect(mayHaveRegionGrid(sheet)).toBe(true)
    paintBoth(sheet)
    expect(vi.mocked(regionGrid).mock.calls.length).toBe(3) // one static paint, two viewport tiles
  })
})
