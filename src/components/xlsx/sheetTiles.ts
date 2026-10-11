/**
 * Base-tile painter for the interactive sheet viewport (plan S5.2, spec §§11.1,
 * 11.3, 12.4). Base tiles are 512×512 device-pixel Canvas2D surfaces keyed by
 * the complete cache owner plus zoom/DPR/backend/grid cell and retained in the
 * bitmap pool through TileCache. Presenting a frame draws cached tiles into the
 * pane surface; only a missing tile calls renderSheetViewport. Overlays
 * (selection, search, watermark) live on their own surface, so moving them never
 * reaches this module.
 */
import type { XlsxSheet } from '../../xlsx/types'
import type { CacheIdentity, LogicalViewport, ProfileId } from '../../xlsx/view/contracts'
import { renderSheetViewport, type RenderSheetAssets } from '../../xlsx/render'
import type { LayoutCache } from '../../xlsx/view/textLayout'
import type { MergeQueryCache, SheetGeometryResult } from '../../xlsx/view/viewport'
import {
  BITMAP_BUDGETS, MemoryPool, RGBA_BYTES, TILE_DEVICE_PX, TILE_PREFETCH_RING, TileCache,
  prefetchTiles, tileLogicalBox, tileRangeFor, type TileIndexRange, type TileKey, type TileSurface,
} from '../../xlsx/view/tileCache'

export const TILE_BACKEND = 'canvas2d'
export interface CanvasTile extends TileSurface { readonly canvas: HTMLCanvasElement; close(): void }
export interface SheetFrame {
  sheet: XlsxSheet
  geometry: SheetGeometryResult
  identity: CacheIdentity
  /** Float64 logical origin of the pane. */
  origin: { x: number; y: number }
  /** Pane size in viewport CSS px. */
  viewport: { width: number; height: number }
  zoom: number
  dpr: number
  extent: { width: number; height: number }
  assets?: RenderSheetAssets
  layoutCache: LayoutCache
  merges: MergeQueryCache
}
export interface SheetTileStats {
  /** renderSheetViewport calls: the only base-cell painting. */
  basePaints: number
  tileHits: number
  prefetched: number
  /** Tiles the bitmap pool refused (painted once, then closed). */
  refused: number
  /** Logical viewports requested by the latest frame (bounded by its tile count). */
  lastRequests: LogicalViewport[]
  /** Peak accounted bitmap bytes: retained tiles + pane surfaces + one in-flight tile. */
  peakBitmapBytes: number
}
const TILE_BYTES = TILE_DEVICE_PX * TILE_DEVICE_PX * RGBA_BYTES

export function createCanvasTile(): CanvasTile {
  const canvas = globalThis.document.createElement('canvas')
  canvas.width = TILE_DEVICE_PX
  canvas.height = TILE_DEVICE_PX
  return {
    canvas,
    get width() { return canvas.width },
    get height() { return canvas.height },
    // Releasing the backing store is the Canvas2D equivalent of ImageBitmap.close.
    close() { canvas.width = 0; canvas.height = 0 },
  }
}

export class SheetTileRenderer {
  readonly stats: SheetTileStats = { basePaints: 0, tileHits: 0, prefetched: 0, refused: 0, lastRequests: [], peakBitmapBytes: 0 }
  readonly budgetBytes: number
  readonly cache: TileCache
  private readonly pool: MemoryPool

  /** `surfaceBytes`: the pane and overlay surfaces, charged before any tile. */
  constructor(readonly profile: ProfileId, readonly surfaceBytes: number, private readonly createTile: () => CanvasTile = createCanvasTile) {
    this.budgetBytes = BITMAP_BUDGETS[profile]
    // One in-flight tile is held outside the pool while it paints.
    const limitBytes = this.budgetBytes - surfaceBytes - TILE_BYTES
    if (!(limitBytes >= TILE_BYTES)) throw new RangeError('Viewport surfaces exceed the bitmap budget')
    this.pool = new MemoryPool({ limitBytes })
    this.cache = new TileCache(this.pool)
  }

  bitmapBytes(): number { return this.cache.stats().bytes + this.surfaceBytes }

  /** Tiles intersecting the pane (and the prefetch ring around them). */
  ranges(frame: SheetFrame): { visible: TileIndexRange; ring: TileIndexRange } {
    const input = { origin: frame.origin, viewport: frame.viewport, zoom: frame.zoom, dpr: frame.dpr, extent: frame.extent }
    return { visible: tileRangeFor(input), ring: tileRangeFor(input, TILE_PREFETCH_RING) }
  }

  /** Paint the pane from tiles. `target` is the device-pixel pane surface; it is
   * repainted only here, so overlay changes never touch it. Returns the device
   * pixel snap applied (the overlay applies the same bias). */
  present(target: CanvasRenderingContext2D, frame: SheetFrame): { bias: { x: number; y: number } } {
    const scale = frame.zoom * frame.dpr
    const px = Math.round(frame.origin.x * scale), py = Math.round(frame.origin.y * scale)
    const { visible } = this.ranges(frame)
    this.stats.lastRequests = []
    target.save()
    try {
      target.setTransform(1, 0, 0, 1, 0, 0)
      target.fillStyle = '#ffffff'
      target.fillRect(0, 0, target.canvas.width, target.canvas.height)
      for (let row = visible.r0; row <= visible.r1; row++) {
        for (let col = visible.c0; col <= visible.c1; col++) {
          const tile = this.tile(frame, col, row, true)
          // Float64 global device offsets minus the snapped origin: small integers.
          target.drawImage(tile.surface.canvas, col * TILE_DEVICE_PX - px, row * TILE_DEVICE_PX - py)
          if (!tile.retained) tile.surface.close()
        }
      }
    } finally { target.restore() }
    // Drop tiles of an older owner/zoom/DPR only after the new frame is up, so the
    // last committed frame stays presented while the replacement paints (§10.1).
    this.cache.invalidate(key => !sameIdentity(key.identity, frame.identity) || key.zoom !== frame.zoom || key.dpr !== frame.dpr)
    return { bias: { x: (frame.origin.x * scale - px) / frame.dpr, y: (frame.origin.y * scale - py) / frame.dpr } }
  }

  /** Paint up to `limit` missing ring tiles (bounded adjacent prefetch). */
  prefetch(frame: SheetFrame, limit = Number.POSITIVE_INFINITY): number {
    const { visible, ring } = this.ranges(frame)
    let painted = 0
    for (const { col, row } of prefetchTiles(visible, ring)) {
      if (painted >= limit) break
      if (this.cache.get(this.key(frame, col, row))) continue
      const tile = this.tile(frame, col, row, false)
      if (!tile.retained) { tile.surface.close(); break }
      painted++
      this.stats.prefetched++
    }
    return painted
  }

  dispose(): void { this.cache.dispose() }

  private key(frame: SheetFrame, col: number, row: number): TileKey {
    return { identity: frame.identity, backend: TILE_BACKEND, zoom: frame.zoom, dpr: frame.dpr, col, row }
  }

  private tile(frame: SheetFrame, col: number, row: number, visible: boolean): { surface: CanvasTile; retained: boolean } {
    const key = this.key(frame, col, row)
    const cached = this.cache.get(key) as CanvasTile | undefined
    if (cached) { if (visible) this.stats.tileHits++; return { surface: cached, retained: true } }
    const surface = this.createTile()
    this.stats.peakBitmapBytes = Math.max(this.stats.peakBitmapBytes, this.bitmapBytes() + TILE_BYTES)
    const ctx = surface.canvas.getContext('2d')
    const box = tileLogicalBox(col, row, frame.zoom, frame.dpr)
    const viewport: LogicalViewport = { box, zoom: frame.zoom, deviceScale: frame.dpr }
    if (visible) this.stats.lastRequests.push(viewport)
    if (ctx) {
      const scale = frame.zoom * frame.dpr
      ctx.setTransform(scale, 0, 0, scale, 0, 0)
      this.stats.basePaints++
      renderSheetViewport(frame.sheet, ctx, frame.geometry, viewport, frame.assets, {
        profile: this.profile, owner: frame.identity, cache: frame.layoutCache, merges: frame.merges,
      })
    }
    const retained = this.cache.put(key, surface)
    if (!retained) this.stats.refused++
    return { surface, retained }
  }
}

function sameIdentity(a: CacheIdentity, b: CacheIdentity): boolean {
  return a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen && a.fontGen === b.fontGen && a.viewGen === b.viewGen
}
