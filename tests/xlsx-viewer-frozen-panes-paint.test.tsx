/** Frozen panes in the interactive pane: tiles per region, divider, semantics and headers. */
import { describe, expect, it, vi } from 'vitest'
import { render, within } from '@testing-library/react'
import { BITMAP_BUDGETS, TILE_DEVICE_PX } from '../src/xlsx/view/tileCache'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { GRID_HEADER_SIZE, viewportHeaderBox } from '../src/components/xlsx/GridHeaders'
import { createScrollWindow, setScrollOffset } from '../src/xlsx/view/scrollWindow'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'

vi.mock('../src/xlsx/view/renderViewport', () => ({
  renderSheetViewport: () => ({ kind: 'detail', diagnostics: [], work: { records: 1, gridlines: 0 } }),
}))
const { SheetTileRenderer, TILE_BYTES } = await import('../src/components/xlsx/sheetTiles')

const T = TILE_DEVICE_PX
describe('frozen pane paint', () => {
  it('draws the body, top strip, left strip and corner from the same tiles with their own origins, plus the divider', () => {
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const canvas = document.createElement('canvas')
    canvas.width = 640; canvas.height = 400
    const pane = canvas.getContext('2d')!
    const draws = vi.spyOn(pane, 'drawImage'), fills = vi.spyOn(pane, 'fillRect'), clips = vi.spyOn(pane, 'clip')
    const origin = { x: 5 * T + 100, y: 5 * T + 100 }
    renderer.present(pane, {
      sheet: {} as never, geometry: {} as never, paletteKey: '', layoutCache: {} as never, merges: {} as never,
      identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 },
      origin, viewport: { width: 640, height: 400 }, zoom: 1, dpr: 1, extent: { width: 64 * T, height: 64 * T },
      frozen: { rows: 1, cols: 1, left: 64, top: 20 },
    })
    const at = draws.mock.calls.map(call => [call[1], call[2]])
    expect(at).toContainEqual([-100, -100]) // body: tile (5,5)
    expect(at).toContainEqual([-100, 0]) // top strip: tile (5,0)
    expect(at).toContainEqual([0, -100]) // left strip: tile (0,5)
    expect(at).toContainEqual([0, 0]) // corner: tile (0,0)
    expect(clips.mock.calls.length).toBe(draws.mock.calls.length)
    // The divider sits on the cell boundary, in the body's first pixel: never over frozen-strip glyphs.
    expect(fills.mock.calls).toContainEqual([64, 0, 1, 400])
    expect(fills.mock.calls).toContainEqual([0, 20, 640, 1])
    expect(fills.mock.calls.some(call => call[0] === 63 || call[1] === 19)).toBe(false)
    renderer.dispose()
  })

  it('a tile shown by several regions is painted and charged once per frame (frame budget debited once)', () => {
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const canvas = document.createElement('canvas')
    canvas.width = 640; canvas.height = 400
    const pane = canvas.getContext('2d')!
    const draws = vi.spyOn(pane, 'drawImage')
    renderer.present(pane, {
      sheet: {} as never, geometry: {} as never, paletteKey: '', layoutCache: {} as never, merges: {} as never,
      identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 },
      origin: { x: 100, y: 100 }, viewport: { width: 400, height: 300 }, zoom: 1, dpr: 1, extent: { width: 64 * T, height: 64 * T },
      frozen: { rows: 1, cols: 1, left: 64, top: 20 },
    })
    // Body, top, left and corner all show tile (0,0) here: four draws, one paint, records counted once.
    expect(draws.mock.calls.filter(call => call[1] <= 0 && call[2] <= 0).length).toBeGreaterThanOrEqual(4)
    expect(renderer.stats.basePaints).toBe(new Set(draws.mock.calls.map(call => call[0])).size)
    expect(renderer.stats.frameRecords).toBe(renderer.stats.basePaints)
    renderer.dispose()
  })

  it('without panes paints exactly as before (no clip, no divider)', () => {
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const canvas = document.createElement('canvas')
    canvas.width = 100; canvas.height = 100
    const pane = canvas.getContext('2d')!
    const clips = vi.spyOn(pane, 'clip'), draws = vi.spyOn(pane, 'drawImage')
    renderer.present(pane, {
      sheet: {} as never, geometry: {} as never, paletteKey: '', layoutCache: {} as never, merges: {} as never,
      identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 },
      origin: { x: 10, y: 10 }, viewport: { width: 100, height: 100 }, zoom: 1, dpr: 1, extent: { width: 64 * T, height: 64 * T },
    })
    expect(clips).not.toHaveBeenCalled()
    expect(draws.mock.calls.map(call => [call[1], call[2]])).toEqual([[-10, -10]])
    renderer.dispose()
  })
})

describe('frozen pane semantics and headers', () => {
  const visibility = visibilityFromHidden(createIntervalSet([]))
  const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }), order: createIdentityOrder(), visibility })
  const viewport = { width: 400, height: 240 }
  const window = setScrollOffset(createScrollWindow(projection, viewport, { frozen: { rows: 1, cols: 1 } }), projection, viewport, { x: 800, y: 2400 })

  it('keeps row 1 and column A in the accessible grid after scrolling, beside the body cells', () => {
    const props = gridFixture({ projection, view: { baseHidden: createIntervalSet([]), visibility, filter: null, sort: null }, window, viewport, selection: { ranges: [], active: null, anchor: null } as never, body: null })
    const ui = render(<GridSemantics {...props} />), grid = ui.getByRole('grid')
    const rows = within(grid).getAllByRole('row').map(row => row.getAttribute('aria-rowindex'))
    // Header row, frozen row 1, then the body from row 2 + 2400/24 = row 102.
    expect(rows.slice(0, 3)).toEqual(['1', '2', String(2400 / 24 + 2 + 1)])
    expect(ui.getAllByRole('gridcell', { name: /^A1,/ })).toHaveLength(1)
    const colHeaders = within(grid).getAllByRole('columnheader')
    expect(colHeaders[0].querySelector('svg[data-icon="corner"]')).toBeTruthy()
    expect(colHeaders.map(header => header.textContent?.trim().charAt(0)).filter(Boolean).slice(0, 2)).toEqual(['A', 'L'])
  })

  it('headers of scrolling rows/columns never overlap the pinned strip', () => {
    expect(viewportHeaderBox(projection, window, viewport, 'column', 0)).toMatchObject({ x: 0, width: 80 })
    // Column K (index 10) starts at logical 800 = the scroll origin: it is under the strip, clipped to start at 80.
    expect(viewportHeaderBox(projection, window, viewport, 'column', 10)).toMatchObject({ x: 80, width: 0 })
    expect(viewportHeaderBox(projection, window, viewport, 'column', 11)).toMatchObject({ x: 80, width: 80 })
    expect(viewportHeaderBox(projection, window, viewport, 'row', 0)).toMatchObject({ y: 0, height: 24 })
    expect(GRID_HEADER_SIZE).toBeGreaterThan(0)
  })
})
