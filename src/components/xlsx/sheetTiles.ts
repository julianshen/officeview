/**
 * Base-tile painter for the interactive sheet viewport (plan S5.2, spec §§11.1,
 * 11.3, 12.4). Base tiles are 512×512 device-pixel Canvas2D surfaces keyed by
 * the complete cache owner plus zoom/DPR/backend/grid cell and retained in the
 * bitmap pool through TileCache. Presenting a frame draws cached tiles into the
 * pane surface; only a missing tile calls renderSheetViewport. Overlays
 * (selection, search) live on their own surface, so moving them never reaches
 * this module; the watermark is stamped into the pane after the tiles.
 */
import type { XlsxSheet } from '../../xlsx/types'
import type { CacheIdentity, LogicalViewport, ProfileId } from '../../xlsx/view/contracts'
import type { RenderSheetAssets } from '../../xlsx/render'
import { renderSheetViewport } from '../../xlsx/view/renderViewport'
import { paintWatermark, type ResolvedWatermark } from '../../core/watermark'
import type { LayoutCache } from '../../xlsx/view/textLayout'
import { frameBudgetFor, type FrameBudget, type MergeQueryCache, type SheetGeometryResult } from '../../xlsx/view/viewport'
import {
  BITMAP_BUDGETS, bitmapPoolPlan, MemoryPool, RGBA_BYTES, TILE_DEVICE_PX, TILE_PREFETCH_RING, TileCache,
  prefetchTiles, tileLogicalBox, tileRangeFor, type TileIndexRange, type TileKey, type TileSurface,
} from '../../xlsx/view/tileCache'

export const TILE_BACKEND = 'canvas2d'
export interface CanvasTile extends TileSurface {
  readonly canvas: HTMLCanvasElement
  close(): void
  /** Detail work the tile embodies, charged whenever it is presented. */
  cost?: { records: number; gridlines: number }
}
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
  /** Stamped into the pane bitmap after the tiles (tiles stay mark-free and
   * cacheable), so it travels with the presented pixels like the page path. */
  watermark?: ResolvedWatermark
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
  /** Detail records/gridline instances presented by the latest frame and by the
   * latest prefetch batch; each stays within the §11.1 per-frame caps. */
  frameRecords: number
  frameGridlines: number
  prefetchRecords: number
  /** Largest record count any single tile has built (§11.1 per-tile cap). */
  maxTileRecords: number
}
const TILE_BYTES = TILE_DEVICE_PX * TILE_DEVICE_PX * RGBA_BYTES

/** Device scales tried, highest first, when the native DPR does not fit (§12.4). */
const DEVICE_SCALE_STEPS = [3, 2.5, 2, 1.5, 1.25, 1, 0.75, 0.5]
/** Pane surface bytes (base + overlay) at a device scale. */
export function paneSurfaceBytes(viewport: { width: number; height: number }, dpr: number): number {
  return bitmapPoolPlan({ cssWidth: viewport.width, cssHeight: viewport.height, dpr }).overlayBytes * 2
}
/**
 * Admission BEFORE any backing store is sized (§12.4): the largest device scale
 * not above `dpr` whose two pane surfaces, base tiles with a one-tile ring and
 * one in-flight replacement row/column fit the profile's bitmap budget; null when
 * none fits (the caller shows a resource outcome instead of allocating).
 */
export function admitDeviceScale(viewport: { width: number; height: number }, dpr: number, profile: ProfileId): number | null {
  const budget = BITMAP_BUDGETS[profile]
  for (const scale of [dpr, ...DEVICE_SCALE_STEPS.filter(step => step < dpr)]) {
    let plan
    try { plan = bitmapPoolPlan({ cssWidth: viewport.width, cssHeight: viewport.height, dpr: scale }) } catch { continue }
    if (plan.overlayBytes * 2 + plan.baseBytes + plan.replacementBytes <= budget) return scale
  }
  return null
}

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
  readonly stats: SheetTileStats = { basePaints: 0, tileHits: 0, prefetched: 0, refused: 0, lastRequests: [], peakBitmapBytes: 0, frameRecords: 0, frameGridlines: 0, prefetchRecords: 0, maxTileRecords: 0 }
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
    const budget = frameBudgetFor(this.profile)
    let records = 0, gridlines = 0
    target.save()
    try {
      target.setTransform(1, 0, 0, 1, 0, 0)
      target.fillStyle = '#ffffff'
      target.fillRect(0, 0, target.canvas.width, target.canvas.height)
      for (let row = visible.r0; row <= visible.r1; row++) {
        for (let col = visible.c0; col <= visible.c1; col++) {
          const tile = this.tile(frame, col, row, true, budget)
          records += tile.surface.cost?.records ?? 0
          gridlines += tile.surface.cost?.gridlines ?? 0
          // Float64 global device offsets minus the snapped origin: small integers.
          target.drawImage(tile.surface.canvas, col * TILE_DEVICE_PX - px, row * TILE_DEVICE_PX - py)
          if (!tile.retained) tile.surface.close()
        }
      }
      if (frame.watermark) {
        target.setTransform(frame.dpr, 0, 0, frame.dpr, 0, 0)
        paintWatermark(target, { widthPx: frame.viewport.width, heightPx: frame.viewport.height }, frame.watermark)
      }
    } finally { target.restore() }
    // Measured from the tiles' own work, independent of the budget bookkeeping.
    this.stats.frameRecords = records
    this.stats.frameGridlines = gridlines
    // Drop tiles of an older owner/zoom/DPR only after the new frame is up, so the
    // last committed frame stays presented while the replacement paints (§10.1).
    this.cache.invalidate(key => !sameIdentity(key.identity, frame.identity) || key.zoom !== frame.zoom || key.dpr !== frame.dpr)
    return { bias: { x: (frame.origin.x * scale - px) / frame.dpr, y: (frame.origin.y * scale - py) / frame.dpr } }
  }

  /** Paint up to `limit` missing ring tiles (bounded adjacent prefetch). */
  prefetch(frame: SheetFrame, limit = Number.POSITIVE_INFINITY): number {
    const { visible, ring } = this.ranges(frame)
    // Prefetch is its own bounded batch with its own per-frame budget.
    const budget = frameBudgetFor(this.profile)
    let painted = 0, records = 0
    for (const { col, row } of prefetchTiles(visible, ring)) {
      if (painted >= limit) break
      if (this.cache.get(this.key(frame, col, row))) continue
      const tile = this.tile(frame, col, row, false, budget)
      records += tile.surface.cost?.records ?? 0
      if (!tile.retained) { tile.surface.close(); break }
      painted++
      this.stats.prefetched++
    }
    this.stats.prefetchRecords = records
    return painted
  }

  dispose(): void { this.cache.dispose() }

  private key(frame: SheetFrame, col: number, row: number): TileKey {
    return { identity: frame.identity, backend: TILE_BACKEND, zoom: frame.zoom, dpr: frame.dpr, col, row }
  }

  private tile(frame: SheetFrame, col: number, row: number, visible: boolean, budget: FrameBudget): { surface: CanvasTile; retained: boolean } {
    const key = this.key(frame, col, row)
    const cached = this.cache.get(key) as CanvasTile | undefined
    const cost = cached?.cost
    // A cached tile counts its detail work against this frame; one that no longer
    // fits is repainted against what is left (coarse past the budget).
    if (cached && cost && cost.records <= budget.records && cost.gridlines <= budget.gridlines) {
      budget.records -= cost.records
      budget.gridlines -= cost.gridlines
      if (visible) this.stats.tileHits++
      return { surface: cached, retained: true }
    }
    const surface = this.createTile()
    this.stats.peakBitmapBytes = Math.max(this.stats.peakBitmapBytes, this.bitmapBytes() + TILE_BYTES)
    const ctx = surface.canvas.getContext('2d')
    const box = tileLogicalBox(col, row, frame.zoom, frame.dpr)
    const viewport: LogicalViewport = { box, zoom: frame.zoom, deviceScale: frame.dpr }
    if (visible) this.stats.lastRequests.push(viewport)
    const full = budget.records >= frameBudgetFor(this.profile).records && budget.gridlines >= frameBudgetFor(this.profile).gridlines
    let kind: string = 'pending'
    let work = { records: 0, gridlines: 0 }
    if (ctx) {
      const scale = frame.zoom * frame.dpr
      ctx.setTransform(scale, 0, 0, scale, 0, 0)
      this.stats.basePaints++
      const result = renderSheetViewport(frame.sheet, ctx, frame.geometry, viewport, frame.assets, {
        profile: this.profile, owner: frame.identity, cache: frame.layoutCache, merges: frame.merges, frameBudget: budget,
      })
      kind = result.kind
      if (result.kind === 'detail') work = { records: result.work.records, gridlines: result.work.gridlines }
      this.stats.maxTileRecords = Math.max(this.stats.maxTileRecords, work.records)
    }
    surface.cost = work
    // A tile made coarse only because earlier tiles spent the frame budget is not
    // the tile's own content: present it, but never cache it.
    if (kind === 'coarse' && !full) return { surface, retained: false }
    const retained = this.cache.put(key, surface)
    if (!retained) this.stats.refused++
    return { surface, retained }
  }
}

function sameIdentity(a: CacheIdentity, b: CacheIdentity): boolean {
  return a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen && a.fontGen === b.fontGen && a.viewGen === b.viewGen
}
