/**
 * Stage 4a (S4 §6.5): the tile source of a streamed sheet. The sheet renderer asks it for the tiles a frame is
 * missing (`request`), and takes them once they have arrived (`take`); it never paints a cell of a stream sheet.
 *
 * - **Batched per frame.** Requests made while a frame is presented go out together on the next timer turn, at most
 *   64 coordinates per batch (`TILE_BATCH_MAX`) and at most 2 batches in flight (R4: ≤ 128 outstanding tiles), the
 *   visible tiles before the prefetch ring (a batch is all visible, or all ring).
 * - **One tuple per batch.** A batch is keyed by the pair (zoom step, dpr bucket), the `viewGen` and the palette of
 *   its keys; an arrival is filed under the key it was requested with, never under the frame's current one, so a
 *   late tile of an older palette or zoom is never handed out under a newer key (Review Focus 3). A request under a
 *   new tuple drops what was queued for the old one and cancels its batches.
 * - **Scrolling away cancels.** `retain` drops queued tiles outside the visible range and the ring, closes arrived
 *   tiles there, and cancels a batch none of whose tiles is wanted any more (`sheet-cancel`, via the signal).
 * - **Rate-limited tiles are retried.** A `rate-limited` refusal re-queues the tile after the result's
 *   `retryAfterMs` as the channel delivers it (already clamped, with its 1 s fallback); no tile stays pending
 *   (Review Focus 2). Any other refusal drops the tile, with a 1 s cooldown before it is asked for again, so a
 *   refused tile never loops a request per frame.
 * - **No deadline of its own.** The channel's 5 s reply deadline restarts with every fragment, so a tile that takes
 *   ~0.9 s after a server worker recycle is simply late, never a failure.
 * - **Paced under the server's message budget.** Every `tiles` send and every cancel is one message toward the
 *   per-sid `'message'` budget (120 per 10 s, 20 per s, breach → 4008): this source spends at most
 *   `MESSAGE_PACE` of it and waits for room, so neither a fast scroll nor a theme flip can close the session.
 *
 * React-free; type-only imports from the stream runtime (the viewer chunk gains no stream module).
 */
import type { RemoteOfficeDoc } from '../../stream/remote/remote-doc'
import type { DecodedTile, TilesResult } from '../../stream/remote/sheet-channel'
import type { TilePalette } from '../../stream/sheet-wire'
import type { TileIndexRange, TileKey } from '../../xlsx/view/tileCache'
import { SHEET_ZOOM_STEPS } from '../../xlsx/view/sheetZoom'
import type { CanvasTile } from './sheetTiles'

type RemoteTiles = NonNullable<RemoteOfficeDoc['sheets']>['tiles']

/** `sheet-wire.ts` `TILE_BATCH_MAX` (pinned equal by the tests; not imported, to keep the stream runtime out). */
export const STREAM_BATCH_MAX = 64
/** R4: at most two batches in flight per sheet source. */
export const STREAM_BATCHES_IN_FLIGHT = 2
/** `sheet-channel.ts` `TILE_RETRY_FALLBACK_MS`: a `rate-limited` result without `retryAfterMs`, and the cooldown of a
 * tile refused for any other reason. */
export const STREAM_RETRY_FALLBACK_MS = 1_000
/** This source's share of the per-sid `'message'` budget (server: 20 per 1 s, 120 per 10 s). The rest is headroom
 * for the session's other messages (layout, RENEW, print). */
export const MESSAGE_PACE = Object.freeze([{ perWindow: 8, windowMs: 1_000 }, { perWindow: 80, windowMs: 10_000 }])

export interface StreamTileSource {
  /** Asks for a missing tile (`visible`: in the pane, else the prefetch ring). Repeats are free. */
  request(key: TileKey, visible: boolean): void
  /** The arrived tile of `key`, handed over once (the caller then owns and closes it). */
  take(key: TileKey): CanvasTile | undefined
  /** Keeps only the tiles in `visible` or `ring`: the rest are dropped, closed or cancelled. */
  retain(visible: TileIndexRange, ring: TileIndexRange): void
  /** Called after each tile arrives (the viewer coalesces repaints); returns the unsubscribe. */
  onArrive(cb: () => void): () => void
  readonly stats: { requested: number; outstanding: number; dropped: number }
  /** Closes every held tile and cancels what is in flight (added beside the brief's interface: unmount, sheet switch). */
  dispose(): void
}

export interface StreamTileSourceOptions {
  sheetId: string
  viewGen: () => number
  palette: () => TilePalette
  tiles: RemoteTiles
  decode: (png: Uint8Array) => Promise<CanvasTile>
  now?: () => number
  setTimer?: typeof setTimeout
}

type State = 'queued' | 'flight' | 'decoding' | 'retry' | 'arrived'
interface Batch {
  readonly tuple: string
  readonly palette: TilePalette
  readonly viewGen: number
  readonly ids: string[]
  readonly abort: AbortController
  settled: boolean
}
interface Entry { readonly id: string; readonly tuple: string; readonly key: TileKey; visible: boolean; state: State; batch?: Batch; tile?: CanvasTile }

const tupleOf = (k: TileKey) => `${k.identity.docId}|${k.identity.sheetId}|${k.identity.valueGen}|${k.identity.fontGen}|${k.identity.viewGen}|${k.paletteKey}|${k.backend}|${k.zoom}|${k.dpr}`
const cell = (tuple: string, col: number, row: number) => `${tuple}|${col},${row}`
const within = (r: TileIndexRange, k: TileKey) => k.col >= r.c0 && k.col <= r.c1 && k.row >= r.r0 && k.row <= r.r1

/** A sliding-window send pacer (every window must have room). */
class Pacer {
  private readonly sent: number[] = []
  constructor(private readonly now: () => number) {}
  /** 0 when a message may go now, else the ms until one may. */
  wait(): number {
    const t = this.now()
    let wait = 0
    for (const w of MESSAGE_PACE) {
      const recent = this.sent.filter(at => at > t - w.windowMs)
      if (recent.length >= w.perWindow) wait = Math.max(wait, recent[recent.length - w.perWindow]! + w.windowMs - t)
    }
    return wait
  }
  take(): boolean {
    if (this.wait() > 0) return false
    const t = this.now()
    this.sent.push(t)
    const horizon = t - Math.max(...MESSAGE_PACE.map(w => w.windowMs))
    while (this.sent.length && this.sent[0]! <= horizon) this.sent.shift()
    return true
  }
}

export function createStreamTileSource(o: StreamTileSourceOptions): StreamTileSource {
  const now = o.now ?? (() => performance.now())
  const setTimer = o.setTimer ?? setTimeout
  const pacer = new Pacer(now)
  const entries = new Map<string, Entry>()
  const inFlight = new Set<Batch>()
  const cooldown = new Map<string, number>()
  const listeners = new Set<() => void>()
  let tuple = ''
  let flushTimer = false
  let closed = false
  const counts = { requested: 0, dropped: 0 }

  const drop = (e: Entry, cool = false) => {
    if (entries.get(e.id) !== e) return
    entries.delete(e.id)
    e.tile?.close()
    e.tile = undefined
    counts.dropped++
    if (cool) cooldown.set(e.id, now() + STREAM_RETRY_FALLBACK_MS)
  }
  const cancel = (b: Batch) => {
    if (b.settled || b.abort.signal.aborted) return
    // A cancel is a message too: without room the batch runs out instead (at most 64 tiles, filed under their keys).
    if (pacer.take()) b.abort.abort()
  }
  const schedule = (ms = 0) => {
    if (flushTimer || closed) return
    flushTimer = true
    setTimer(() => { flushTimer = false; flush() }, ms)
  }

  /** A new tuple (zoom step, dpr bucket, palette, viewGen): what the old one queued, held or sent is no longer wanted. */
  const adopt = (next: string) => {
    if (next === tuple) return
    tuple = next
    for (const e of [...entries.values()]) if (e.tuple !== next && e.state !== 'flight' && e.state !== 'decoding') drop(e)
    for (const b of inFlight) if (b.tuple !== next) cancel(b)
  }

  const flush = () => {
    if (closed) return
    while (inFlight.size < STREAM_BATCHES_IN_FLIGHT) {
      const queued = [...entries.values()].filter(e => e.state === 'queued' && e.tuple === tuple)
      if (!queued.length) return
      const visible = queued.filter(e => e.visible)
      const pick = (visible.length ? visible : queued).slice(0, STREAM_BATCH_MAX)
      const wait = pacer.wait()
      if (wait > 0) { schedule(wait); return }
      pacer.take()
      send(pick)
    }
  }

  const send = (pick: Entry[]) => {
    const k = pick[0]!.key
    const zoomStep = SHEET_ZOOM_STEPS.indexOf(k.zoom)
    const batch: Batch = { tuple: pick[0]!.tuple, palette: o.palette(), viewGen: k.identity.viewGen, ids: pick.map(e => e.id), abort: new AbortController(), settled: false }
    for (const e of pick) { e.state = 'flight'; e.batch = batch }
    counts.requested += pick.length
    inFlight.add(batch)
    const onTile = (t: DecodedTile) => {
      if (closed || batch.settled || batch.abort.signal.aborted) return
      const e = entries.get(cell(batch.tuple, t.col, t.row))
      if (!e || e.batch !== batch || e.state !== 'flight') return
      // The theme or the view moved on while it was on its way: it is not this frame's tile.
      if (o.palette() !== batch.palette || o.viewGen() !== batch.viewGen) { drop(e); return }
      e.state = 'decoding'
      o.decode(t.png).then(tile => {
        if (closed || entries.get(e.id) !== e || e.state !== 'decoding') { tile.close(); return }
        // Decoded after the frame moved to another tuple: nobody will take it.
        if (e.tuple !== tuple) { e.tile = tile; drop(e); return }
        e.state = 'arrived'
        e.tile = tile
        for (const cb of [...listeners]) { try { cb() } catch { /* a listener never breaks the source */ } }
      }, () => drop(e, true))
    }
    let result: Promise<TilesResult>
    try {
      result = o.tiles({ sheetId: o.sheetId, viewGen: batch.viewGen, zoomStep, dpr: k.dpr as 1 | 2, palette: batch.palette, tiles: pick.map(e => ({ col: e.key.col, row: e.key.row })) }, onTile, batch.abort.signal)
    } catch (error) { result = Promise.reject(error) }
    result.then(r => settle(batch, r), () => settle(batch, { refused: [] }))
  }

  const settle = (batch: Batch, r: TilesResult) => {
    batch.settled = true
    inFlight.delete(batch)
    if (closed) return
    const retry: Entry[] = []
    for (const x of r.refused ?? []) {
      const e = entries.get(cell(batch.tuple, x.col, x.row))
      if (!e || e.batch !== batch || e.state !== 'flight') continue
      // A batch of an older tuple (its cancel was paced out) is not retried: nothing would take its tiles.
      if (x.code === 'rate-limited' && batch.tuple === tuple) { e.state = 'retry'; retry.push(e) } else drop(e, x.code !== 'rate-limited')
    }
    // Neither delivered nor refused (cancelled, or the request failed): not pending any more; asked again when wanted.
    for (const id of batch.ids) { const e = entries.get(id); if (e && e.batch === batch && e.state === 'flight') drop(e) }
    if (retry.length) {
      setTimer(() => {
        if (closed) return
        for (const e of retry) {
          if (entries.get(e.id) !== e || e.state !== 'retry') continue
          if (e.tuple !== tuple) { drop(e); continue }
          e.state = 'queued'
          e.batch = undefined
        }
        flush()
      }, r.retryAfterMs ?? STREAM_RETRY_FALLBACK_MS)
    }
    flush()
  }

  return {
    request(key, visible) {
      if (closed || key.identity.sheetId !== o.sheetId) return
      if (!SHEET_ZOOM_STEPS.includes(key.zoom) || (key.dpr !== 1 && key.dpr !== 2) || key.identity.viewGen !== o.viewGen()) return
      const t = tupleOf(key), id = cell(t, key.col, key.row)
      adopt(t)
      const e = entries.get(id)
      if (e) { if (visible && !e.visible) e.visible = true; return }
      const until = cooldown.get(id)
      if (until !== undefined) { if (now() < until) return; cooldown.delete(id) }
      entries.set(id, { id, tuple: t, key, visible, state: 'queued' })
      schedule()
    },
    take(key) {
      const id = cell(tupleOf(key), key.col, key.row)
      const e = entries.get(id)
      if (!e || e.state !== 'arrived' || !e.tile) return undefined
      entries.delete(id)
      const tile = e.tile
      e.tile = undefined
      return tile
    },
    retain(visible, ring) {
      const wanted = (k: TileKey) => within(visible, k) || within(ring, k)
      for (const e of [...entries.values()]) if ((e.state === 'queued' || e.state === 'retry' || e.state === 'arrived') && !wanted(e.key)) drop(e)
      for (const b of inFlight) {
        if (b.ids.some(id => { const e = entries.get(id); return e !== undefined && e.batch === b && e.state === 'flight' && wanted(e.key) })) continue
        cancel(b)
      }
    },
    onArrive(cb) {
      listeners.add(cb)
      return () => { listeners.delete(cb) }
    },
    get stats() {
      let outstanding = 0
      for (const e of entries.values()) if (e.state !== 'arrived') outstanding++
      return { requested: counts.requested, outstanding, dropped: counts.dropped }
    },
    dispose() {
      if (closed) return
      closed = true
      for (const b of inFlight) cancel(b)
      for (const e of entries.values()) e.tile?.close()
      entries.clear()
      listeners.clear()
    },
  }
}
