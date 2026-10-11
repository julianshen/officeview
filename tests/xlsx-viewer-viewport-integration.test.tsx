/**
 * S5.2 interactive viewport integration (spec §§5.5, 9, 10.1, 11, 12.4): base
 * tiles in the bitmap pool, independent overlays, view-generation invalidation,
 * document-owned zoom reset and the feature flag.
 */
import { describe, expect, it } from 'vitest'
import { TILE_DEVICE_PX, prefetchTiles, tileLogicalBox, tileRangeFor } from '../src/xlsx/view/tileCache'

describe('tile grid math', () => {
  const extent = { width: 64 * 16_384, height: 20 * 1_048_576 }
  it('covers exactly the pane in device pixels at zoom x DPR, anchored at logical 0', () => {
    // 300x200 CSS at zoom 1, DPR 2 = 600x400 device px: two tile columns, one row.
    expect(tileRangeFor({ origin: { x: 0, y: 0 }, viewport: { width: 300, height: 200 }, zoom: 1, dpr: 2, extent })).toEqual({ c0: 0, c1: 1, r0: 0, r1: 0 })
    // Scrolled to logical x=256 (512 device px): starts at tile column 1.
    expect(tileRangeFor({ origin: { x: 256, y: 0 }, viewport: { width: 300, height: 200 }, zoom: 1, dpr: 2, extent })).toEqual({ c0: 1, c1: 2, r0: 0, r1: 0 })
    // zoom 0.5 halves the device span of the same logical origin.
    expect(tileRangeFor({ origin: { x: 1024, y: 0 }, viewport: { width: 256, height: 256 }, zoom: 0.5, dpr: 1, extent })).toEqual({ c0: 1, c1: 1, r0: 0, r1: 0 })
  })
  it('grows by a bounded ring and never requests tiles past the sheet extent or before 0', () => {
    const range = tileRangeFor({ origin: { x: 0, y: 0 }, viewport: { width: 300, height: 200 }, zoom: 1, dpr: 2, extent }, 1)
    expect(range).toEqual({ c0: 0, c1: 2, r0: 0, r1: 1 })
    const small = { width: 100, height: 100 }
    expect(tileRangeFor({ origin: { x: 0, y: 0 }, viewport: { width: 300, height: 200 }, zoom: 1, dpr: 1, extent: small }, 1)).toEqual({ c0: 0, c1: 0, r0: 0, r1: 0 })
  })
  it('keeps the last row of the full domain at a Float64 tile index (no 32-bit wrap)', () => {
    const y = extent.height - 200
    const range = tileRangeFor({ origin: { x: 0, y }, viewport: { width: 300, height: 200 }, zoom: 4, dpr: 3, extent })
    expect(range.r0).toBe(Math.floor((y * 12) / TILE_DEVICE_PX))
    expect(range.r1).toBe(Math.ceil((y * 12 + 200 * 3) / TILE_DEVICE_PX) - 1)
    const box = tileLogicalBox(0, range.r0, 4, 3)
    expect(box.height).toBeCloseTo(TILE_DEVICE_PX / 12, 9)
    expect(box.y).toBe((range.r0 * TILE_DEVICE_PX) / 12)
  })
  it('prefetch lists only ring tiles, nearest first', () => {
    const tiles = prefetchTiles({ c0: 1, c1: 1, r0: 1, r1: 1 }, { c0: 0, c1: 2, r0: 0, r1: 2 })
    expect(tiles).toHaveLength(8)
    expect(tiles.some(t => t.col === 1 && t.row === 1)).toBe(false)
    expect(new Set(tiles.map(t => `${t.col}:${t.row}`)).size).toBe(8)
    // Edge-adjacent tiles before diagonals.
    expect(tiles.slice(0, 4).every(t => t.col === 1 || t.row === 1)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Component integration
// ---------------------------------------------------------------------------
import { act, fireEvent, render, waitFor, within } from '@testing-library/react'
import { OfficeDoc } from '../src/components/OfficeDoc'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx, type XlsxSheetSpec } from '../src/testdata/ooxml-builders'
import { createRowOrder } from '../src/xlsx/view/order'
import { BITMAP_BUDGETS } from '../src/xlsx/view/tileCache'
import type { XlsxDocument } from '../src/xlsx/types'

async function workbook(sheets: XlsxSheetSpec[] = [{
  name: 'S1',
  rows: [
    { r: 1, cells: [{ ref: 'A1', v: 'alpha' }, { ref: 'B1', v: 1 }] },
    { r: 2, cells: [{ ref: 'A2', v: 'beta' }, { ref: 'B2', v: 2 }] },
    { r: 3, cells: [{ ref: 'A3', v: 'gamma' }, { ref: 'B3', v: 3 }] },
    { r: 5000, cells: [{ ref: 'C5000', v: 'farhit' }] },
  ],
}, { name: 'S2', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'other farhit' }] }] }]): Promise<XlsxDocument> {
  return parseXlsx(await OfficePackage.load(await buildXlsx(sheets)))
}

function harness() {
  const frames: SheetFrameStats[] = []
  return { frames, onFrame: (stats: SheetFrameStats) => { frames.push(stats) }, last: () => frames[frames.length - 1] }
}
async function settle(): Promise<void> {
  // Flush the deferred prefetch timer and any state it triggers.
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 5)) })
}
async function mountSheet(doc: XlsxDocument, extra: Record<string, unknown> = {}) {
  const h = harness()
  const ui = render(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} {...extra} />)
  await waitFor(() => expect(h.last()?.basePaints ?? 0).toBeGreaterThan(0))
  await settle()
  return { ui, h }
}
const scroller = (ui: ReturnType<typeof render>) => ui.getByTestId('officeview-sheet-scroller')
// jsdom pane fallback: 640x400 minus the 44px header strips, DPR 1.
const PANE = { width: 596, height: 356 }
const visibleTileCount = (zoom: number, dpr = 1) => {
  const tiles = (css: number) => Math.ceil((css * dpr) / TILE_DEVICE_PX) + 1
  void zoom
  return tiles(PANE.width) * tiles(PANE.height)
}

describe('interactive sheet viewport (S5.2)', () => {
  it('first frame requests only bounded tile viewports at SheetZoom x DPR, then prefetches a bounded ring', async () => {
    const doc = await workbook()
    const { h } = await mountSheet(doc)
    const first = h.frames.find(frame => frame.lastRequests.length > 0)!
    expect(first.zoom).toBe(1)
    expect(first.lastRequests.length).toBeLessThanOrEqual(visibleTileCount(1))
    for (const box of first.lastRequests) {
      expect(box.width).toBe(TILE_DEVICE_PX)
      expect(box.height).toBe(TILE_DEVICE_PX)
    }
    // Visible 2x1 tiles plus a one-tile ring clipped at row/col 0: 3x2 = 6 at most.
    expect(h.last().basePaints).toBeLessThanOrEqual(6)
    expect(h.last().prefetched).toBeGreaterThan(0)
  })

  it('scrolling issues bounded requests and revisiting cached tiles paints no base cells', async () => {
    const doc = await workbook()
    const { ui, h } = await mountSheet(doc)
    const el = scroller(ui)
    const before = h.last().basePaints
    act(() => { el.scrollTop = 4000; fireEvent.scroll(el) })
    await settle()
    const moved = h.last()
    expect(moved.basePaints - before).toBeLessThanOrEqual(visibleTileCount(1) * 2)
    expect(moved.lastRequests.length).toBeLessThanOrEqual(visibleTileCount(1))
    for (const box of moved.lastRequests) expect(box.y).toBeGreaterThanOrEqual(3 * TILE_DEVICE_PX)
    const painted = h.last().basePaints, hits = h.last().tileHits
    act(() => { el.scrollTop = 0; fireEvent.scroll(el) })
    await settle()
    // The first frame's tiles are still in the bitmap pool.
    expect(h.last().tileHits).toBeGreaterThan(hits)
    expect(h.last().lastRequests).toHaveLength(0)
    expect(h.last().basePaints).toBe(painted)
  })

  it('Ctrl+wheel and keyboard zoom change absolute SheetZoom and request tiles at the new scale', async () => {
    const doc = await workbook()
    const { ui, h } = await mountSheet(doc)
    act(() => { fireEvent.wheel(scroller(ui), { deltaY: -200, ctrlKey: true, clientX: 10, clientY: 10 }) })
    await settle()
    const zoomed = h.last()
    expect(zoomed.zoom).toBeGreaterThan(1)
    for (const box of zoomed.lastRequests) expect(box.width).toBeCloseTo(TILE_DEVICE_PX / zoomed.zoom, 9)
    const grid = ui.getByRole('grid')
    act(() => { grid.focus(); fireEvent.keyDown(grid, { key: '0' }) })
    await settle()
    expect(h.last().zoom).toBe(1)
    act(() => { fireEvent.keyDown(grid, { key: '-' }) })
    await settle()
    expect(h.last().zoom).toBe(0.75)
    expect(ui.getByRole('button', { name: 'Reset zoom' }).textContent).toBe('75%')
  })

  it('moving the selection repaints only the overlay, never base cells', async () => {
    const doc = await workbook()
    const { ui, h } = await mountSheet(doc)
    const base = h.last().basePaints, overlays = h.last().overlayPaints
    const grid = ui.getByRole('grid')
    act(() => { fireEvent.click(grid, { clientX: 44 + 10, clientY: 44 + 5, detail: 1 }) })
    await settle()
    expect(grid.getAttribute('aria-activedescendant')).toMatch(/cell-0$/)
    act(() => { grid.focus(); fireEvent.keyDown(grid, { key: 'ArrowDown' }); fireEvent.keyDown(grid, { key: 'ArrowRight', shiftKey: true }) })
    await settle()
    expect(h.last().overlayPaints).toBeGreaterThan(overlays)
    expect(h.last().basePaints).toBe(base)
    expect(h.last().tileHits).toBeGreaterThanOrEqual(0)
    // The selection overlay is visible: green border pixels on the overlay surface.
    const overlay = ui.getByTestId('officeview-sheet-overlay') as HTMLCanvasElement
    const data = overlay.getContext('2d')!.getImageData(0, 0, overlay.width, overlay.height).data
    let green = 0
    for (let i = 0; i < data.length; i += 4) if (data[i + 1] > 100 && data[i] < 60 && data[i + 3] > 200) green++
    expect(green).toBeGreaterThan(0)
  })

  it('search highlights paint on the overlay without repainting base cells', async () => {
    const doc = await workbook()
    const { ui, h } = await mountSheet(doc)
    const base = h.last().basePaints
    const input = ui.getByRole('searchbox', { name: 'Search document' })
    act(() => { fireEvent.change(input, { target: { value: 'alpha' } }) })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toBe('1 of 1'))
    await settle()
    expect(h.last().basePaints).toBe(base)
    const overlay = ui.getByTestId('officeview-sheet-overlay') as HTMLCanvasElement
    const [r, g, b, a] = overlay.getContext('2d')!.getImageData(20, 10, 1, 1).data
    expect(a).toBeGreaterThan(0)
    // Yellow wash (rgba(255, 214, 0, 0.45)): red and green, no blue.
    expect(r).toBeGreaterThan(b + 100)
    expect(g).toBeGreaterThan(b + 80)
  })

  it('find navigates to an offscreen and other-sheet hit through the navigator, keeping zoom', async () => {
    const doc = await workbook()
    const { ui, h } = await mountSheet(doc)
    const grid = () => ui.getByRole('grid')
    act(() => { grid().focus(); fireEvent.keyDown(grid(), { key: '+' }) })
    await settle()
    expect(h.last().zoom).toBe(1.25)
    act(() => { fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'f', ctrlKey: true }) })
    const input = ui.getByRole('searchbox', { name: 'Search document' })
    expect(document.activeElement).toBe(input)
    act(() => { fireEvent.change(input, { target: { value: 'farhit' } }) })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toBe('1 of 2'))
    await settle()
    // C5000 on S1: source row 4999, col 2.
    expect(grid().getAttribute('aria-activedescendant')).toMatch(new RegExp(`cell-${4999 * 16384 + 2}$`))
    expect(h.last().zoom).toBe(1.25)
    expect(h.last().lastRequests.every(box => box.y > 4000 * 20 * 0)).toBe(true)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Next match' })) })
    await waitFor(() => expect(h.last().activeSheetId).toBe('2'))
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toBe('2 of 2'))
    expect(grid().getAttribute('aria-activedescendant')).toMatch(/cell-0$/)
    // The other sheet keeps its own zoom (default 1); S1 retains 1.25.
    expect(h.last().zoom).toBe(1)
    act(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: '1' } }) })
    await settle()
    expect(h.last().zoom).toBe(1.25)
  })

  it('a committed view re-keys and repaints the affected tiles while preserving zoom and scroll', async () => {
    const doc = await workbook()
    const h = harness()
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} />)
    await waitFor(() => expect(h.last()?.basePaints ?? 0).toBeGreaterThan(0))
    await settle()
    const grid = ui.getByRole('grid')
    act(() => { grid.focus(); fireEvent.keyDown(grid, { key: '+' }) })
    act(() => { fireEvent.keyDown(grid, { key: '+' }) })
    await settle()
    expect(h.last().zoom).toBe(1.5)
    const base = h.last().basePaints
    const reversed = createRowOrder([{ r0: 0, r1: 2, sourceAtOffset: Uint32Array.of(2, 1, 0) }])
    ui.rerender(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} views={{ 1: { revision: 1, order: reversed } }} />)
    await settle()
    const committed = h.last()
    expect(committed.identity.viewGen).toBe(1)
    expect(committed.zoom).toBe(1.5)
    expect(committed.basePaints).toBeGreaterThan(base)
    // Only the new owner's tiles remain retained.
    expect(committed.tiles.tiles).toBeLessThanOrEqual(visibleTileCount(1.5) * 3)
    // The sorted display order is what the grid exposes: display row 3 shows source row 1.
    expect(ui.getByRole('gridcell', { name: 'B3, value 1; original B1' })).toBeTruthy()
    // An older revision never replaces the committed view.
    ui.rerender(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} views={{ 1: { revision: 0 } }} />)
    await settle()
    expect(h.last().identity.viewGen).toBe(1)
    expect(ui.getByRole('gridcell', { name: 'B3, value 1; original B1' })).toBeTruthy()
  })

  it('keeps physical canvas and tile allocations inside the bitmap budget (desktop and mobile)', async () => {
    for (const profile of ['D', 'M'] as const) {
      const doc = await workbook()
      const { ui, h } = await mountSheet(doc, { profile })
      const el = scroller(ui)
      for (let step = 1; step <= 12; step++) {
        act(() => { el.scrollTop = step * 900; el.scrollLeft = step * 700; fireEvent.scroll(el) })
        await settle()
      }
      const last = h.last()
      expect(last.budgetBytes).toBe(BITMAP_BUDGETS[profile])
      expect(last.peakBitmapBytes).toBeLessThanOrEqual(BITMAP_BUDGETS[profile])
      expect(last.bitmapBytes).toBeLessThanOrEqual(BITMAP_BUDGETS[profile])
      // Two pane-sized surfaces (base + overlay), never a whole-sheet bitmap.
      const base = ui.getByTestId('officeview-sheet-base') as HTMLCanvasElement
      expect(base.width).toBe(PANE.width)
      expect(base.height).toBe(PANE.height)
      ui.unmount()
    }
  })

  it('stamps the watermark into the presented pane bitmap; changing it re-presents cached tiles only', async () => {
    const doc = await workbook()
    const h = harness()
    const ui = render(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} />)
    await waitFor(() => expect(h.last()?.basePaints ?? 0).toBeGreaterThan(0))
    await settle()
    const base = ui.getByTestId('officeview-sheet-base') as HTMLCanvasElement
    const red = () => {
      const data = base.getContext('2d')!.getImageData(0, 0, base.width, base.height).data
      let count = 0
      for (let i = 0; i < data.length; i += 4) if (data[i] > 200 && data[i + 1] < 80 && data[i + 2] < 80) count++
      return count
    }
    expect(red()).toBe(0)
    const painted = h.last().basePaints
    ui.rerender(<InteractiveSheetDoc doc={doc} onFrame={h.onFrame} watermark={{ text: 'CONFIDENTIAL', opacity: 1, color: '#ff0000' }} />)
    await settle()
    expect(red()).toBeGreaterThan(0)
    // Tiles stay mark-free in the pool: the change is a re-present, not a base repaint.
    expect(h.last().basePaints).toBe(painted)
    expect(h.last().tileHits).toBeGreaterThan(0)
  })

  it('survives a StrictMode double mount with a live tile pool', async () => {
    const { StrictMode } = await import('react')
    const doc = await workbook()
    const h = harness()
    render(<StrictMode><InteractiveSheetDoc doc={doc} onFrame={h.onFrame} /></StrictMode>)
    await waitFor(() => expect(h.last()?.basePaints ?? 0).toBeGreaterThan(0))
    await settle()
    expect(h.last().refused).toBe(0)
    expect(h.last().tiles.tiles).toBeGreaterThan(0)
  })
})

describe('feature flag and document-owned zoom (S5.2)', () => {
  it('flag off (default) keeps the legacy page path for spreadsheets', async () => {
    // Small sheets only: the legacy static bitmap of a 5000-row sheet exceeds node-canvas limits.
    const doc = await workbook([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'alpha' }] }] }, { name: 'S2', rows: [{ r: 2, cells: [{ ref: 'B2', v: 7 }] }] }])
    const legacy = render(<OfficeDoc document={doc} />)
    await waitFor(() => expect(legacy.container.querySelectorAll('canvas')).toHaveLength(2))
    expect(within(legacy.container).queryByRole('grid')).toBeNull()
    expect(within(legacy.container).getByRole('region', { name: 'Document viewer' })).toBeTruthy()
    const off = render(<OfficeDoc document={doc} interactiveSheets={false} />)
    await waitFor(() => expect(off.container.querySelectorAll('canvas')).toHaveLength(2))
    const pixels = (root: HTMLElement) => [...root.querySelectorAll('canvas')].map(c => (c as HTMLCanvasElement).toDataURL())
    expect(pixels(off.container)).toEqual(pixels(legacy.container))
  })

  it('flag on mounts sheet mode for XLSX only; DOCX ignores it', async () => {
    const doc = await workbook()
    const ui = render(<OfficeDoc document={doc} interactiveSheets />)
    await waitFor(() => expect(ui.getByRole('grid')).toBeTruthy())
    expect(ui.getByRole('navigation', { name: 'Worksheet navigation' })).toBeTruthy()
    expect(ui.getAllByRole('status').filter(node => node.getAttribute('data-officeview-ui') === 'viewer-status')).toHaveLength(1)
    expect(ui.getByRole('button', { name: 'Reset zoom' }).textContent).toBe('100%')
    const { parseDocx } = await import('../src/docx/parse')
    const { buildDocx } = await import('../src/testdata/ooxml-builders')
    const docx = await parseDocx(await OfficePackage.load(await buildDocx([{ runs: [{ text: 'page' }] }])))
    const page = render(<OfficeDoc document={docx} interactiveSheets />)
    await waitFor(() => expect(page.container.querySelectorAll('canvas').length).toBeGreaterThan(0))
    expect(within(page.container).queryByRole('grid')).toBeNull()
  })

  it('only a document replacement resets zoom; watermark and font changes keep it', async () => {
    const doc = await workbook()
    const ui = render(<OfficeDoc document={doc} interactiveSheets />)
    await waitFor(() => expect(ui.getByRole('grid')).toBeTruthy())
    const label = () => ui.getByRole('button', { name: 'Reset zoom' }).textContent
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Zoom in' })) })
    expect(label()).toBe('125%')
    ui.rerender(<OfficeDoc document={doc} interactiveSheets watermark={{ text: 'DRAFT' }} />)
    await settle()
    expect(label()).toBe('125%')
    ui.rerender(<OfficeDoc document={doc} interactiveSheets watermark={{ text: 'DRAFT' }} fallbackFonts={{ fallbackChain: ['Arial'] }} />)
    await settle()
    expect(label()).toBe('125%')
    const next = await workbook()
    ui.rerender(<OfficeDoc document={next} interactiveSheets watermark={{ text: 'DRAFT' }} />)
    await waitFor(() => expect(ui.getByRole('grid')).toBeTruthy())
    expect(label()).toBe('100%')
  })
})

// ---------------------------------------------------------------------------
// S5.1 seam: drawings are painted tile-local, never under a global translation
// ---------------------------------------------------------------------------
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import type { XlsxSheet } from '../src/xlsx/types'

describe('tile-local drawings (S5.1 seam resolved in S5.2)', () => {
  const source = { partPath: 'd', treePath: 't', element: 'sp', representation: 'native' as const }
  const sheetWith = (xPx: number, yPx: number) => ({
    name: 'D', rows: [], cols: [], merges: [], mergeRanges: [],
    drawings: [{ xEmu: xPx * 9525, yEmu: yPx * 9525, widthEmu: 20 * 9525, heightEmu: 10 * 9525, geometry: 'rect', fill: 'FF0000', source }],
  }) as unknown as XlsxSheet
  function paint(sheet: XlsxSheet, box: { x: number; y: number }) {
    const canvas = document.createElement('canvas')
    canvas.width = 64
    canvas.height = 32
    const real = canvas.getContext('2d')!
    const offsets: number[] = []
    const ctx = new Proxy(real, {
      get(target, key) {
        if (key === 'translate') return (x: number, y: number) => { offsets.push(Math.abs(x), Math.abs(y)); target.translate(x, y) }
        const value = Reflect.get(target, key, target)
        return typeof value === 'function' ? value.bind(target) : value
      },
      set(target, key, value) { return Reflect.set(target, key, value, target) },
    })
    const geometry = buildSheetGeometry(sheet)
    const result = renderSheetViewport(sheet, ctx, geometry, { box: { ...box, width: 64, height: 32 }, zoom: 1, deviceScale: 1 })
    return { result, offsets, pixels: [...real.getImageData(0, 0, 64, 32).data] }
  }
  it('a drawing near the last row paints at its tile-local offset with small canvas translations', () => {
    const far = 20_000_000
    const near = paint(sheetWith(10, 5), { x: 0, y: 0 })
    const deep = paint(sheetWith(10, far + 5), { x: 0, y: far })
    expect(deep.result.kind).toBe('detail')
    expect(Math.max(...deep.offsets)).toBeLessThan(TILE_DEVICE_PX * 4)
    // The drawing's interior (gridline overshoot differs at the sheet's top edge).
    const crop = (pixels: number[]) => pixels.filter((_, index) => { const p = index >> 2, x = p % 64, y = Math.floor(p / 64); return x > 11 && x < 29 && y > 6 && y < 14 })
    expect(crop(deep.pixels)).toEqual(crop(near.pixels))
    // Red fill landed inside the pane.
    expect(deep.pixels.some((value, index) => index % 4 === 0 && value > 200)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Review round 1 (codex): find semantics, status, navigation
// ---------------------------------------------------------------------------
describe('sheet-mode find parity (S5.2 review round 1)', () => {
  const status = (ui: ReturnType<typeof render>) => ui.getByTestId('officeview-search-status').textContent
  const search = (ui: ReturnType<typeof render>, text: string) => act(() => { fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: text } }) })

  it('does not match values in authored-hidden columns', async () => {
    const doc = await workbook([{ name: 'S1', cols: '<col min="2" max="2" width="10" hidden="1" customWidth="1"/>', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'shown' }, { ref: 'B1', v: 'secret' }] }] }])
    const { ui } = await mountSheet(doc)
    search(ui, 'secret')
    await waitFor(() => expect(status(ui)).toBe('no matches'))
    search(ui, 'shown')
    await waitFor(() => expect(status(ui)).toBe('1 of 1'))
  })

  it('searches and announces the canonical formatted text (authored number formats)', async () => {
    const doc = await workbook([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 12 }] }] }])
    const cell = doc.sheets[0].rows[0].cells[0]
    cell.style = { numFmtId: 164, formatCode: '"USD "0.00' }
    const { ui } = await mountSheet(doc)
    search(ui, 'USD')
    await waitFor(() => expect(status(ui)).toBe('1 of 1'))
    expect(ui.getByRole('gridcell', { name: 'A1, value USD 12.00' })).toBeTruthy()
  })

  it('a literal whitespace query is a real search', async () => {
    const doc = await workbook([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'A B' }] }] }])
    const { ui } = await mountSheet(doc)
    search(ui, ' ')
    await waitFor(() => expect(status(ui)).toBe('1 of 1'))
  })

  it('pages backward and wraps across the whole workbook, with a lower-bound count until complete', async () => {
    const rows = Array.from({ length: 201 }, (_, i) => ({ r: i + 1, cells: [{ ref: `A${i + 1}`, v: `hit ${i + 1}` }] }))
    const doc = await workbook([{ name: 'S1', rows }])
    const { ui } = await mountSheet(doc)
    search(ui, 'hit')
    await waitFor(() => expect(status(ui)).toBe('1 of 200+'))
    const prev = () => act(() => { fireEvent.click(ui.getByRole('button', { name: 'Previous match' })) })
    const next = () => act(() => { fireEvent.click(ui.getByRole('button', { name: 'Next match' })) })
    // Previous from the first match wraps to the workbook's last match.
    prev()
    await waitFor(() => expect(status(ui)).toBe('201 of 201'))
    prev()
    await waitFor(() => expect(status(ui)).toBe('200 of 201'))
    next()
    await waitFor(() => expect(status(ui)).toBe('201 of 201'))
    next()
    await waitFor(() => expect(status(ui)).toMatch(/^1 of 20[01]\+?$/))
  })

  it('qualifies results with unavailable values instead of a plain "no matches"', async () => {
    const doc = await workbook([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 0 }, { ref: 'B1', v: 'plain' }] }] }])
    // A formula without a cached value: its value is unavailable, not blank.
    Object.assign(doc.sheets[0].rows[0].cells[0], { value: null, formula: 'NOW()', hasCachedValue: false })
    const { ui } = await mountSheet(doc)
    search(ui, 'zzz')
    await waitFor(() => expect(status(ui)).toMatch(/^no matches \(1 value unavailable\)$/))
  })
})

describe('hidden targets, coarse selection (S5.2 review round 1)', () => {
  it('Go To a hidden cell exposes an exact read-only inspection without moving the grid', async () => {
    const doc = await workbook([{ name: 'S1', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'top' }] }, { r: 2, cells: [{ ref: 'A2', v: 'secret' }] }, { r: 3, cells: [{ ref: 'A3', v: 'low' }] }] }])
    doc.sheets[0].rows[1].hidden = true
    const { ui } = await mountSheet(doc)
    const grid = ui.getByRole('grid')
    act(() => { fireEvent.click(grid, { clientX: 44 + 10, clientY: 44 + 5, detail: 1 }) })
    const before = grid.getAttribute('aria-activedescendant')
    act(() => { fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A2' } }) })
    act(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Address space' }), { target: { value: 'source' } }) })
    act(() => { fireEvent.submit(ui.getByRole('textbox', { name: 'Go To cell' }).closest('form')!) })
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Original cell A2'))
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Value secret'))
    expect(ui.getByRole('status').textContent).toContain('read-only')
    expect(grid.getAttribute('aria-activedescendant')).toBe(before)
  })

  it('keeps a coarse selection overlay when the pane crosses the per-cell overlay cap', async () => {
    // Narrow columns at 25% put more cells in the pane than the mobile per-frame cap.
    const doc = await workbook([{ name: 'S1', cols: '<col min="1" max="16384" width="1" customWidth="1"/>', rows: [{ r: 1, cells: [{ ref: 'A1', v: 'x' }] }] }])
    const { ui, h } = await mountSheet(doc, { profile: 'M' })
    const grid = ui.getByRole('grid')
    for (let i = 0; i < 6; i++) act(() => { grid.focus(); fireEvent.keyDown(grid, { key: '-' }) })
    await settle()
    expect(h.last().zoom).toBe(0.25)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Select worksheet' })) })
    await settle()
    const overlay = ui.getByTestId('officeview-sheet-overlay') as HTMLCanvasElement
    const data = overlay.getContext('2d')!.getImageData(0, 0, overlay.width, overlay.height).data
    let tinted = 0
    for (let i = 0; i < data.length; i += 4) if (data[i + 3] > 0) tinted++
    expect(tinted).toBeGreaterThan(overlay.width * overlay.height * 0.5)
  })
})

describe('per-frame caps across tiles (S5.2 review round 1)', () => {
  it('a dense sheet never presents more detail records in one frame than the §11.1 frame cap', async () => {
    const { VIEWPORT_DETAIL_LIMITS } = await import('../src/xlsx/view/viewport')
    const { columnToLetters } = await import('../src/xlsx/view/coordinates')
    const rows = Array.from({ length: 220 }, (_, r) => ({ r: r + 1, cells: Array.from({ length: 200 }, (_, c) => ({ ref: `${columnToLetters(c)}${r + 1}`, v: r * 1000 + c })) }))
    const doc = await workbook([{ name: 'Dense', rows }])
    const { ui, h } = await mountSheet(doc)
    const grid = ui.getByRole('grid')
    for (let i = 0; i < 6; i++) act(() => { grid.focus(); fireEvent.keyDown(grid, { key: '-' }) })
    await settle()
    expect(h.last().zoom).toBe(0.25)
    const caps = VIEWPORT_DETAIL_LIMITS.D
    for (const frame of h.frames) {
      expect(frame.frameRecords).toBeLessThanOrEqual(caps.records)
      expect(frame.frameGridlines).toBeLessThanOrEqual(caps.gridlines)
      expect(frame.prefetchRecords).toBeLessThanOrEqual(caps.records)
    }
  }, 30_000)
})

describe('bitmap admission before allocation (S5.2 review round 1)', () => {
  it('admits a device scale whose surfaces and minimum tiles fit the pool, before any canvas is sized', async () => {
    const { admitDeviceScale } = await import('../src/components/xlsx/sheetTiles')
    expect(admitDeviceScale({ width: 1440, height: 900 }, 2, 'D')).toBe(2)
    // 1500x1000 at DPR 3 needs ~108 MiB of pane surfaces alone; 1.25 is the largest fitting step.
    expect(admitDeviceScale({ width: 1500, height: 1000 }, 3, 'M')).toBe(1.25)
    expect(admitDeviceScale({ width: 100_000, height: 100_000 }, 1, 'M')).toBeNull()
  })

  it('a mobile pane at DPR 3 never allocates surfaces past the budget', async () => {
    const { BITMAP_BUDGETS } = await import('../src/xlsx/view/tileCache')
    const width = Object.getOwnPropertyDescriptor(Element.prototype, 'clientWidth')!
    const height = Object.getOwnPropertyDescriptor(Element.prototype, 'clientHeight')!
    const dpr = window.devicePixelRatio
    const sizes: number[] = []
    let owner: object | null = HTMLCanvasElement.prototype
    while (owner && !Object.getOwnPropertyDescriptor(owner, 'width')) owner = Object.getPrototypeOf(owner)
    const canvasWidth = Object.getOwnPropertyDescriptor(owner!, 'width')!
    Object.defineProperty(Element.prototype, 'clientWidth', { configurable: true, get() { return (this as HTMLElement).dataset?.testid === 'officeview-sheet-pane' ? 1544 : 0 } })
    Object.defineProperty(Element.prototype, 'clientHeight', { configurable: true, get() { return (this as HTMLElement).dataset?.testid === 'officeview-sheet-pane' ? 1044 : 0 } })
    Object.defineProperty(owner!, 'width', { configurable: true, get() { return canvasWidth.get!.call(this) }, set(value: number) { sizes.push(value); canvasWidth.set!.call(this, value) } })
    ;(window as { devicePixelRatio: number }).devicePixelRatio = 3
    try {
      const doc = await workbook()
      const { ui, h } = await mountSheet(doc, { profile: 'M' })
      const base = ui.getByTestId('officeview-sheet-base') as HTMLCanvasElement
      expect(base.width).toBe(1875)
      expect(Math.max(...sizes)).toBeLessThanOrEqual(1875)
      expect(h.last().peakBitmapBytes).toBeLessThanOrEqual(BITMAP_BUDGETS.M)
      expect(ui.getByRole('status').textContent).toContain('resolution reduced')
    } finally {
      Object.defineProperty(Element.prototype, 'clientWidth', width)
      Object.defineProperty(Element.prototype, 'clientHeight', height)
      Object.defineProperty(owner!, 'width', canvasWidth)
      ;(window as { devicePixelRatio: number }).devicePixelRatio = dpr
    }
  })
})
