import { expect, it } from 'vitest'
import { BITMAP_BUDGETS } from '../src/xlsx/view/tileCache'
import { SheetTileRenderer, TILE_BYTES, admitDeviceScale, paneSurfaceBytes } from '../src/components/xlsx/sheetTiles'
import { buildSheetGeometry, MergeQueryCache } from '../src/xlsx/view/viewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import type { XlsxSheet } from '../src/xlsx/types'


it.each([false, true])('pending frame surfaces remain within the admitted bitmap limit with a warm cache (frozen=%s)', frozen => {
  const viewport = { width: 1600, height: 1200 }, surfaceBytes = paneSurfaceBytes(viewport, 1)
  expect(admitDeviceScale(viewport, 1, 'M')).toBe(1)
  let live = 0, peak = 0
  const renderer = new SheetTileRenderer('M', surfaceBytes, () => {
    const canvas = document.createElement('canvas'); canvas.width = 512; canvas.height = 512
    live += TILE_BYTES; peak = Math.max(peak, surfaceBytes + live)
    let open = true
    return { canvas, width: 512, height: 512, close() { if (open) { open = false; live -= TILE_BYTES; canvas.width = 0; canvas.height = 0 } } }
  })
  const target = document.createElement('canvas'); target.width = viewport.width; target.height = viewport.height
  const ctx = target.getContext('2d')!
  const sheet: XlsxSheet = { name: 'S', rows: [], cols: [], merges: [], mergeRanges: [] }
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('geometry')
  const frame = { sheet, geometry, paletteKey: '', layoutCache: new LayoutCache(), merges: new MergeQueryCache(),
    identity: { docId: 'r2', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 }, origin: { x: 0, y: 0 }, viewport, zoom: 1, dpr: 1,
    extent: { width: 1e6, height: 1e6 }, ...(frozen ? { frozen: { rows: 1, cols: 1, left: 64, top: 20 } } : {}) }
  for (let i = 0; i < 8; i++) renderer.present(ctx, { ...frame, origin: { x: 4096 * i, y: 4096 * i } })
  peak = surfaceBytes + live
  renderer.present(ctx, { ...frame, origin: { x: 50100, y: 50100 }, pending: (box: { x: number; y: number; width: number; height: number }) => [box] })
  console.log({ actualPeak: peak, reportedPeak: renderer.stats.peakBitmapBytes, budget: BITMAP_BUDGETS.M, end: surfaceBytes + live })
  renderer.dispose()
  expect(live).toBe(0)
  expect(peak).toBeLessThanOrEqual(BITMAP_BUDGETS.M)
})
