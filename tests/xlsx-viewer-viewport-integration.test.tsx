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
