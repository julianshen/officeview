/**
 * Stage 4a Task 13 (S4 §6.5): the renderer swap. On a stream frame the sheet renderer never paints a cell: a missing
 * tile is requested from the server through a `StreamTileSource` (batched per frame, at most 64 per batch and 2
 * batches in flight, visible tiles first), painted with the pending fill meanwhile and never cached; an arrived tile is
 * cached under the same `TileKey`. No watermark is stamped on a stream frame. Zoom settles to the 11 steps and the
 * dpr to the {1, 2} buckets.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { DecodedTile, TilesRequest, TilesResult } from '../src/stream/remote/sheet-channel'
import type { StreamSheetLayout, TilePalette } from '../src/stream/sheet-wire'
import { streamSheetGeometry } from '../src/xlsx/view/streamGeometry'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { MergeQueryCache } from '../src/xlsx/view/viewport'
import { BITMAP_BUDGETS, TILE_DEVICE_PX, tileRangeFor, type TileIndexRange, type TileKey } from '../src/xlsx/view/tileCache'
import type { CacheIdentity } from '../src/xlsx/view/contracts'
import { dprBucket, nearestZoomStep } from '../src/xlsx/view/sheetZoom'

vi.mock('../src/xlsx/view/renderViewport', async original => {
  const actual = await original<typeof import('../src/xlsx/view/renderViewport')>()
  return { ...actual, renderSheetViewport: vi.fn(actual.renderSheetViewport) }
})
vi.mock('../src/core/watermark', async original => {
  const actual = await original<typeof import('../src/core/watermark')>()
  return { ...actual, paintWatermark: vi.fn(actual.paintWatermark) }
})
const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')
const { paintWatermark, normalizeWatermark } = await import('../src/core/watermark')
const { SheetTileRenderer, TILE_BYTES } = await import('../src/components/xlsx/sheetTiles')
const { createStreamTileSource } = await import('../src/components/xlsx/streamTiles')
type CanvasTile = import('../src/components/xlsx/sheetTiles').CanvasTile
type SheetFrame = import('../src/components/xlsx/sheetTiles').SheetFrame

afterEach(() => { vi.useRealTimers(); vi.clearAllMocks() })

const T = TILE_DEVICE_PX
const ID: CacheIdentity = { docId: 'doc', sheetId: '0', valueGen: 0, fontGen: 0, viewGen: 3 }
const key = (col: number, row: number, o: { palette?: TilePalette; zoom?: number; dpr?: number } = {}): TileKey =>
  ({ identity: ID, paletteKey: o.palette ?? 'light', backend: 'canvas2d', zoom: o.zoom ?? 1, dpr: o.dpr ?? 1, col, row })
const range = (c0: number, c1: number, r0: number, r1: number): TileIndexRange => ({ c0, c1, r0, r1 })

interface Call { req: TilesRequest; onTile: (t: DecodedTile) => void; signal?: AbortSignal; resolve: (r: TilesResult) => void; reject: (e: unknown) => void }
/** A server stand-in: every `tiles` batch is held until the test delivers or settles it. */
function fakeServer() {
  const calls: Call[] = []
  const tiles = vi.fn((req: TilesRequest, onTile: (t: DecodedTile) => void, signal?: AbortSignal) => new Promise<TilesResult>((resolve, reject) => {
    const call: Call = { req, onTile, signal, resolve, reject }
    calls.push(call)
    signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true })
  }))
  const deliver = (call: Call, coords = call.req.tiles) => {
    for (const t of coords) call.onTile({ sheetIndex: Number(call.req.sheetId), viewGen: call.req.viewGen, zoomStep: call.req.zoomStep, dpr: call.req.dpr, col: t.col, row: t.row, png: new Uint8Array([call.req.palette === 'dark' ? 2 : 1]) })
  }
  return { calls, tiles, deliver }
}
const decoded: CanvasTile[] = []
const decode = vi.fn(async (png: Uint8Array): Promise<CanvasTile> => {
  const canvas = document.createElement('canvas')
  canvas.width = T; canvas.height = T
  const tile: CanvasTile = { canvas, get width() { return canvas.width }, get height() { return canvas.height }, close: vi.fn(() => { canvas.width = 0; canvas.height = 0 }) }
  ;(tile as unknown as { png: number }).png = png[0]
  decoded.push(tile)
  return tile
})
/** Flushes the source's per-frame timer and the decode promises it awaits. */
async function tick(ms = 0) {
  await vi.advanceTimersByTimeAsync(ms)
  for (let i = 0; i < 5; i++) await Promise.resolve()
}
function source(o: { palette?: () => TilePalette; viewGen?: () => number } = {}) {
  const server = fakeServer()
  const src = createStreamTileSource({
    sheetId: '0', viewGen: o.viewGen ?? (() => 3), palette: o.palette ?? (() => 'light'), tiles: server.tiles, decode,
    now: () => Date.now(), setTimer: setTimeout,
  })
  return { server, src }
}

const LAYOUT: StreamSheetLayout = {
  sheetId: '0', viewGen: 3, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
  merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 0, totalRows: 0 },
}
function streamFrame(stream: SheetFrame['stream'], o: Partial<SheetFrame> = {}): SheetFrame {
  const geometry = streamSheetGeometry(LAYOUT)
  return {
    sheet: { name: 'Data', rows: [], cols: [], merges: [], mergeRanges: [] } as never, geometry, paletteKey: 'light', identity: ID,
    origin: { x: 0, y: 0 }, viewport: { width: 1200, height: 700 }, zoom: 1, dpr: 1,
    extent: { width: geometry.metrics.cols.totalCss(), height: geometry.metrics.rows.totalCss() },
    layoutCache: new LayoutCache(), merges: new MergeQueryCache(64), stream, ...o,
  }
}
function pane(frame: SheetFrame) {
  const canvas = document.createElement('canvas')
  canvas.width = Math.ceil(frame.viewport.width * frame.dpr); canvas.height = Math.ceil(frame.viewport.height * frame.dpr)
  return canvas.getContext('2d')!
}
const visibleOf = (frame: SheetFrame) => tileRangeFor({ origin: frame.origin, viewport: frame.viewport, zoom: frame.zoom, dpr: frame.dpr, extent: frame.extent })
const count = (r: TileIndexRange) => (r.c1 - r.c0 + 1) * (r.r1 - r.r0 + 1)
const inRange = (r: TileIndexRange) => (t: { col: number; row: number }) => t.col >= r.c0 && t.col <= r.c1 && t.row >= r.r0 && t.row <= r.r1

describe('stream tiles (S4 §6.5)', () => {
  it('a missing tile is requested, not painted', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const frame = streamFrame(src)
    renderer.present(pane(frame), frame)
    const visible = visibleOf(frame)
    expect(renderer.stats.basePaints).toBe(0)
    expect(vi.mocked(renderSheetViewport)).toHaveBeenCalledTimes(0)
    expect(renderer.stats.pendingTiles).toBe(count(visible))
    expect(renderer.cache.stats().tiles).toBe(0)
    await tick()
    const asked = server.calls.flatMap(c => c.req.tiles.map(t => `${t.col},${t.row}`))
    const expected: string[] = []
    for (let row = visible.r0; row <= visible.r1; row++) for (let col = visible.c0; col <= visible.c1; col++) expected.push(`${col},${row}`)
    expect(asked.sort()).toEqual(expected.sort())
    // A second present before anything arrives asks for nothing again.
    renderer.present(pane(frame), frame)
    await tick()
    expect(server.calls.flatMap(c => c.req.tiles)).toHaveLength(expected.length)
    expect(server.calls[0]!.req).toMatchObject({ sheetId: '0', viewGen: 3, zoomStep: 5, dpr: 1, palette: 'light' })
    renderer.dispose(); src.dispose()
  })

  it('an arriving tile is cached and a repaint scheduled', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    const arrive = vi.fn()
    src.onArrive(arrive)
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const frame = streamFrame(src, { viewport: { width: 300, height: 200 } })
    renderer.present(pane(frame), frame)
    await tick()
    expect(server.calls).toHaveLength(1)
    expect(server.calls[0]!.req.tiles).toEqual([{ col: 0, row: 0 }])
    server.deliver(server.calls[0]!)
    server.calls[0]!.resolve({ refused: [] })
    await tick()
    expect(arrive).toHaveBeenCalledTimes(1)
    const ctx = pane(frame), draws = vi.spyOn(ctx, 'drawImage')
    renderer.present(ctx, frame)
    expect(renderer.stats.tileHits).toBe(1)
    expect(renderer.stats.pendingTiles).toBe(0)
    expect(renderer.stats.basePaints).toBe(0)
    expect(draws).toHaveBeenCalledWith(decoded.at(-1)!.canvas, 0, 0)
    // Cached under the same TileKey: the next frame is a cache hit, nothing is requested again.
    renderer.present(pane(frame), frame)
    await tick()
    expect(renderer.stats.tileHits).toBe(2)
    expect(server.calls).toHaveLength(1)
    expect(src.stats.outstanding).toBe(0)
    renderer.dispose(); src.dispose()
  })

  it('visible before the ring', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    const visible = range(4, 6, 4, 5), ring = range(3, 7, 3, 6)
    // The ring is asked for first: the visible tiles still go out first, in a batch of their own.
    for (let row = ring.r0; row <= ring.r1; row++) for (let col = ring.c0; col <= ring.c1; col++) if (!inRange(visible)({ col, row })) src.request(key(col, row), false)
    for (let row = visible.r0; row <= visible.r1; row++) for (let col = visible.c0; col <= visible.c1; col++) src.request(key(col, row), true)
    await tick()
    expect(server.calls.length).toBeGreaterThan(0)
    expect(server.calls[0]!.req.tiles).toHaveLength(count(visible))
    expect(server.calls[0]!.req.tiles.every(inRange(visible))).toBe(true)
    src.dispose()
  })

  it('bounded in flight', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    for (let i = 0; i < 300; i++) src.request(key(i % 20, Math.floor(i / 20)), true)
    await tick(); await tick(16)
    const open = () => server.calls.filter(c => !(c as { done?: boolean }).done)
    expect(open().length).toBeLessThanOrEqual(2)
    for (const c of server.calls) expect(c.req.tiles.length).toBeLessThanOrEqual(64)
    expect(src.stats.requested).toBeLessThanOrEqual(128)
    // Draining: each settled batch lets the next go, never more than two at once, every key once.
    const seen = new Set<string>()
    for (let guard = 0; guard < 50 && seen.size < 300; guard++) {
      const live = open()
      expect(live.length).toBeLessThanOrEqual(2)
      const c = live[0]
      if (!c) { await tick(1000); continue }
      for (const t of c.req.tiles) { const k = `${t.col},${t.row}`; expect(seen.has(k)).toBe(false); seen.add(k) }
      server.deliver(c); c.resolve({ refused: [] }); (c as { done?: boolean }).done = true
      await tick(); await tick(1000)
    }
    expect(seen.size).toBe(300)
    for (const c of server.calls) expect(c.req.tiles.length).toBeLessThanOrEqual(64)
    expect(src.stats.outstanding).toBe(0)
    src.dispose()
  })

  it('scrolling away cancels', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    for (let col = 0; col < 3; col++) src.request(key(col, 0), true)
    await tick()
    const first = server.calls[0]!
    expect(first.signal?.aborted).toBe(false)
    // Still partly visible: kept.
    src.retain(range(2, 4, 0, 1), range(1, 5, 0, 2))
    expect(first.signal?.aborted).toBe(false)
    // Disjoint from the batch: cancelled.
    src.retain(range(40, 42, 40, 41), range(39, 43, 39, 42))
    expect(first.signal?.aborted).toBe(true)
    await tick()
    expect(src.stats.outstanding).toBe(0)
    expect(src.stats.dropped).toBe(3)
    // A cancelled tile is asked for again when it is wanted again.
    src.request(key(0, 0), true)
    await tick(1000)
    expect(server.calls.at(-1)!.req.tiles).toEqual([{ col: 0, row: 0 }])
    src.dispose()
  })

  it('rate-limited tiles are retried, none stuck pending', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    for (let col = 0; col < 8; col++) src.request(key(col, 0), true)
    await tick()
    const first = server.calls[0]!
    const refused = first.req.tiles.slice(5)
    server.deliver(first, first.req.tiles.slice(0, 5))
    first.resolve({ refused: refused.map(t => ({ ...t, code: 'rate-limited' })), retryAfterMs: 500 })
    await tick()
    expect(refused).toHaveLength(3)
    expect(server.calls).toHaveLength(1)
    expect(src.stats.outstanding).toBe(3)
    await tick(499)
    expect(server.calls).toHaveLength(1)
    await tick(1)
    expect(server.calls).toHaveLength(2)
    expect(server.calls[1]!.req.tiles).toEqual(refused.map(t => ({ col: t.col, row: t.row })))
    server.deliver(server.calls[1]!)
    server.calls[1]!.resolve({ refused: [] })
    await tick()
    for (let col = 0; col < 8; col++) expect(src.take(key(col, 0))).toBeDefined()
    expect(src.stats.outstanding).toBe(0)
    src.dispose()
  })

  it('a tile ~0.9 s late (a worker recycle) is not a failure', async () => {
    vi.useFakeTimers()
    const { server, src } = source()
    for (let col = 0; col < 4; col++) src.request(key(col, 0), true)
    await tick()
    const c = server.calls[0]!
    server.deliver(c, c.req.tiles.slice(0, 3))
    await tick(900)
    server.deliver(c, c.req.tiles.slice(3))
    c.resolve({ refused: [] })
    await tick()
    for (let col = 0; col < 4; col++) expect(src.take(key(col, 0))).toBeDefined()
    expect(server.calls).toHaveLength(1)
    expect(src.stats.dropped).toBe(0)
    expect(src.stats.outstanding).toBe(0)
    src.dispose()
  })

  it('no watermark on a stream frame', async () => {
    vi.useFakeTimers()
    const { src } = source()
    const renderer = new SheetTileRenderer('D', BITMAP_BUDGETS.D - TILE_BYTES * 64)
    const frame = streamFrame(src, { watermark: normalizeWatermark({ text: 'CONFIDENTIAL' })! })
    renderer.present(pane(frame), frame)
    expect(vi.mocked(paintWatermark)).toHaveBeenCalledTimes(0)
    // The local path still stamps it (the guard is the stream, not the frame).
    const local = { ...frame, stream: undefined, viewport: { width: 10, height: 10 } }
    renderer.present(pane(local), local)
    expect(vi.mocked(paintWatermark)).toHaveBeenCalledTimes(1)
    renderer.dispose(); src.dispose()
  })

  it('the source mirrors the wire constants it does not import', async () => {
    const wire = await import('../src/stream/sheet-wire')
    const channel = await import('../src/stream/remote/sheet-channel')
    const tiles = await import('../src/components/xlsx/streamTiles')
    const { SHEET_ZOOM_STEPS } = await import('../src/xlsx/view/sheetZoom')
    expect(tiles.STREAM_BATCH_MAX).toBe(wire.TILE_BATCH_MAX)
    expect(tiles.STREAM_RETRY_FALLBACK_MS).toBe(channel.TILE_RETRY_FALLBACK_MS)
    expect([...SHEET_ZOOM_STEPS]).toEqual([...wire.TILE_ZOOM_STEPS])
  })

  it('zoom steps and dpr buckets', () => {
    expect(nearestZoomStep(0.8)).toBe(0.75)
    expect(nearestZoomStep(0.29)).toBe(0.25)
    expect(nearestZoomStep(9)).toBe(4)
    expect(dprBucket(0.9)).toBe(1)
    expect(dprBucket(1)).toBe(1)
    expect(dprBucket(1.25)).toBe(2)
    expect(dprBucket(3)).toBe(2)
  })

  it('theme switch', async () => {
    vi.useFakeTimers()
    let palette: TilePalette = 'light'
    const { server, src } = source({ palette: () => palette })
    src.request(key(0, 0), true)
    await tick()
    const old = server.calls[0]!
    expect(old.req.palette).toBe('light')
    // The theme switches while the light tile is still on its way.
    palette = 'dark'
    src.request(key(0, 0, { palette: 'dark' }), true)
    await tick(1000)
    const fresh = server.calls.find(c => c.req.palette === 'dark')
    expect(fresh?.req.tiles).toEqual([{ col: 0, row: 0 }])
    server.deliver(old); old.resolve({ refused: [] })
    await tick()
    // The late light tile is never handed out under the dark key.
    expect(src.take(key(0, 0, { palette: 'dark' }))).toBeUndefined()
    server.deliver(fresh!); fresh!.resolve({ refused: [] })
    await tick()
    const dark = src.take(key(0, 0, { palette: 'dark' }))
    expect((dark as unknown as { png: number }).png).toBe(2)
    expect(src.stats.outstanding).toBe(0)
    src.dispose()
  })
})
