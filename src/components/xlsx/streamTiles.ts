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
 *   (Review Focus 2).
 * - **Other failures are classified** (`failureKind`). A transient one (`policy-refused`, which is also how the
 *   channel's 5 s reply deadline arrives; `session-rotated`; a decode failure) drops the tile and, when the 1 s
 *   cooldown ends, asks for it again by itself if it is still visible (through the pacer, like any request), at
 *   most `STREAM_RETRY_LIMIT - 1` times in a row: the `STREAM_RETRY_LIMIT`-th consecutive failure leaves it on the
 *   pending fill. A permanent one (`unsupported-op`, `stale-view`, `layout-too-large`, and any code not known to be
 *   transient: fail closed) is never retried by itself. A stopped tile is asked for again only after a user-driven
 *   repaint, i.e. a frame whose visible range or tuple differs (an arrival's repaint changes neither).
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
/** The consecutive transient failures of one tile key that stop its automatic retries. */
export const STREAM_RETRY_LIMIT = 3
/** Refusal codes a retry can cure. The channel's reply deadline refuses the outstanding tiles `policy-refused`, and a
 * tile's render failure on the server is `policy-refused` too; `session-rotated` is a RENEW race. Every other code is
 * permanent (fail closed): explicitly `unsupported-op`, `stale-view` and `layout-too-large`. */
const TRANSIENT_CODES: ReadonlySet<string> = new Set(['policy-refused', 'session-rotated'])
/** The internal code of a tile whose PNG would not decode (transient). */
export const DECODE_FAILED = 'decode-failed'
export function failureKind(code: string): 'transient' | 'permanent' {
  return code === DECODE_FAILED || TRANSIENT_CODES.has(code) ? 'transient' : 'permanent'
}

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
  clearTimer?: typeof clearTimeout
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
  const setTimeoutFn = o.setTimer ?? setTimeout, clearTimeoutFn = o.clearTimer ?? clearTimeout
  // Every timer this source arms, so dispose clears them all.
  const timers = new Set<ReturnType<typeof setTimeout>>()
  const setTimer = (fn: () => void, ms: number) => {
    const handle: ReturnType<typeof setTimeout> = setTimeoutFn(() => { timers.delete(handle); fn() }, ms)
    timers.add(handle)
  }
  const pacer = new Pacer(now)
  const entries = new Map<string, Entry>()
  const inFlight = new Set<Batch>()
  const cooldown = new Map<string, number>()
  /** Consecutive transient failures per tile id (cleared by an arrival or a user-driven repaint). */
  const failures = new Map<string, number>()
  /** Tile ids not retried by themselves (a permanent refusal, or the retry limit): asked again only after a user-driven
   * repaint. */
  const stopped = new Set<string>()
  let lastVisible: TileIndexRange | undefined
  const listeners = new Set<() => void>()
  let tuple = ''
  let flushTimer = false
  let closed = false
  const counts = { requested: 0, dropped: 0 }

  /** `failed`: the refusal or decode code when the tile failed (classified), else it was merely not wanted. */
  const drop = (e: Entry, failed?: string) => {
    if (entries.get(e.id) !== e) return
    entries.delete(e.id)
    e.tile?.close()
    e.tile = undefined
    counts.dropped++
    if (failed === undefined) return
    cooldown.set(e.id, now() + STREAM_RETRY_FALLBACK_MS)
    const n = (failures.get(e.id) ?? 0) + 1
    failures.set(e.id, n)
    if (failureKind(failed) === 'permanent' || n >= STREAM_RETRY_LIMIT) { stopped.add(e.id); return }
    // Transient: once the cooldown ends, asked again by itself if it is still visible and its tuple still current.
    setTimer(() => {
      if (closed || e.tuple !== tuple || stopped.has(e.id) || entries.has(e.id)) return
      if (!(lastVisible ? within(lastVisible, e.key) : e.visible)) return
      cooldown.delete(e.id)
      entries.set(e.id, { id: e.id, tuple: e.tuple, key: e.key, visible: true, state: 'queued' })
      flush()
    }, STREAM_RETRY_FALLBACK_MS)
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
    // A new tuple is a user-driven repaint (zoom, theme, view): nothing of the old one stays stopped.
    stopped.clear()
    failures.clear()
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
        failures.delete(e.id)
        e.tile = tile
        for (const cb of [...listeners]) { try { cb() } catch { /* a listener never breaks the source */ } }
      }, () => drop(e, DECODE_FAILED))
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
      if (x.code === 'rate-limited') { if (batch.tuple === tuple) { e.state = 'retry'; retry.push(e) } else drop(e) } else drop(e, x.code)
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
      if (stopped.has(id)) return
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
      // A changed visible range is a user-driven repaint (a scroll): stopped tiles may be asked for again.
      if (lastVisible && (lastVisible.c0 !== visible.c0 || lastVisible.c1 !== visible.c1 || lastVisible.r0 !== visible.r0 || lastVisible.r1 !== visible.r1)) { stopped.clear(); failures.clear() }
      lastVisible = visible
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
      for (const handle of timers) clearTimeoutFn(handle)
      timers.clear()
      for (const b of inFlight) cancel(b)
      for (const e of entries.values()) e.tile?.close()
      entries.clear()
      listeners.clear()
    },
  }
}
