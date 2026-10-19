/**
 * S10.2 main ledger, canvasText pool: Chrome keeps the shaping of every distinct
 * text a 2D context drew for as long as the context lives, so each retained
 * tile canvas charges its distinct (font, text) pairs. The charge must end when
 * the tile cache closes the tile (eviction under the bitmap budget, invalidation,
 * disposal), and a fresh tile drawing the same text charges again.
 */
import { describe, expect, it, vi } from 'vitest'
import { TILE_CANVAS_BYTES, shapeCacheBytes } from '../src/xlsx/view/mainLedger'
import { BITMAP_BUDGETS, TILE_DEVICE_PX } from '../src/xlsx/view/tileCache'

// Each tile draws one text of its own and one text every tile shares (twice,
// to show repeats within a tile count once).
vi.mock('../src/xlsx/view/renderViewport', () => ({
  renderSheetViewport: (_sheet: unknown, ctx: CanvasRenderingContext2D, _geometry: unknown, viewport: { box: { x: number; y: number } }) => {
    ctx.font = '12px sans-serif'
    ctx.fillText(`tile ${viewport.box.x},${viewport.box.y}`, 0, 10)
    ctx.fillText('shared', 0, 20)
    ctx.measureText('shared')
    return { kind: 'detail', diagnostics: [], work: { records: 1, gridlines: 0 } }
  },
}))

const { SheetTileRenderer, TILE_BYTES } = await import('../src/components/xlsx/sheetTiles')

// One font per tile: two distinct texts in one shaping cache.
const tileBytes = (col: number) => TILE_CANVAS_BYTES + shapeCacheBytes(2, `tile ${col * TILE_DEVICE_PX},0`.length + 'shared'.length)

function frameAt(col: number) {
  return {
    sheet: {} as never, geometry: {} as never, paletteKey: '', layoutCache: {} as never, merges: {} as never,
    identity: { docId: 'd', sheetId: 's', valueGen: 0, fontGen: 0, viewGen: 0 },
    origin: { x: col * TILE_DEVICE_PX, y: 0 }, viewport: { width: 10, height: 10 },
    zoom: 1, dpr: 1, extent: { width: 64 * TILE_DEVICE_PX, height: TILE_DEVICE_PX },
  }
}

describe('canvasText charges follow the tile lifetime', () => {
  it('evicted tiles release their charge; a fresh tile with the same text charges again; dispose releases all', () => {
    // A pool of three tiles: past its high watermark (2.625 tiles) it evicts
    // down to its low watermark (2.25), so it holds two tiles steadily.
    const surfaceBytes = BITMAP_BUDGETS.D - TILE_BYTES - 3 * TILE_BYTES
    const renderer = new SheetTileRenderer('D', surfaceBytes)
    const pane = document.createElement('canvas').getContext('2d')!

    renderer.present(pane, frameAt(0))
    expect(renderer.cache.stats().tiles).toBe(1)
    expect(renderer.textCacheBytes()).toBe(tileBytes(0))

    renderer.present(pane, frameAt(1))
    expect(renderer.textCacheBytes()).toBe(tileBytes(0) + tileBytes(1))

    // Tile 2 evicts tile 0 (LRU): its charge ends with it.
    renderer.present(pane, frameAt(2))
    expect(renderer.cache.stats().tiles).toBe(2)
    expect(renderer.textCacheBytes()).toBe(tileBytes(1) + tileBytes(2))

    // Tile 0 again: a fresh canvas shapes the same texts again and charges again
    // (evicting tile 1).
    renderer.present(pane, frameAt(0))
    expect(renderer.textCacheBytes()).toBe(tileBytes(2) + tileBytes(0))

    renderer.dispose()
    expect(renderer.textCacheBytes()).toBe(0)
  })
})
