/**
 * Piece 5 (plan D1, D12, D23): one session's show-start encode queue over the shared cache.
 *
 * - **Order** (D1): slide order from the start unit to the last slide, then the slides before it; within a slide, the
 *   items in `media[]` order. Items are deduplicated by key (`encodeKeyId`): one part can yield several `mediaId`s.
 *   A cursor outside `[current, next]` re-queues from that unit; a forward move releases the passed slides'
 *   references (`cache.passed`).
 * - **One loop per session, one awaited key at a time** (spec §6). Encode-ahead acquires `ahead`; on `{ wait }` it
 *   waits for room (a release, a delivery, an eviction), a cursor move or a request, holding no admission place. A
 *   `request` for a key that is cached (`peek`) answers at once; any other goes to the front of the loop, which
 *   acquires it in `request` mode (the only mode that drops, evicts or refuses `media-busy`).
 * - **Agnostic of slide kind.** The queue works from whatever items `items(unit)` yields (still or animated slide).
 * - **Origin** (D23): the session's first reference to a key fixes its `EncodeOrigin`, kept across drops and
 *   re-acquires; `request` returns it.
 * - **`end()`**: cancels the session's queued items and requests, `cache.dropSession(sid)` (an encode another
 *   session waits on keeps running), and resolves once every encode left with no waiter has settled (its child gone,
 *   its part and output files disposed: `cache.quiesce()`). Cache entries remain.
 */
import type { MediaRefusal } from '../../../src/stream/media-types'
import type { MediaAdmission } from './admission'
import type { EncodeCache, EncodeOrigin, EncodeOutcome, EncodeRun } from './cache'
import type { FfmpegInfo } from './ffmpeg'
import type { MediaItem } from './jobs'
import { encodeKeyId, type EncodeKey } from './params'
import type { MediaTemp } from './temp'
import type { EncodeResult, TranscodeSeams, transcode } from './transcode'

export type MediaRequestResult = { ok: true; path: string; result: EncodeResult & { ok: true }; origin: EncodeOrigin } | { ok: false; reason: MediaRefusal }
export interface SessionMediaQueue {
  /** Keys enqueued: not cached, not joined. */
  start(unit: number): Promise<number>
  cursor(unit: number): void
  request(mediaId: string): Promise<MediaRequestResult>
  end(): Promise<void>
  /** The mediaId of the item the loop awaits now. */
  readonly running: string | undefined
}
export interface SessionMediaQueueOptions {
  sid: string
  items: (unit: number) => Promise<MediaItem[]>
  part: (partPath: string) => Promise<{ path: string; dispose(): Promise<void> }>
  units: number
  info: FfmpegInfo
  temp: MediaTemp
  admission: MediaAdmission
  cache: EncodeCache /* the shared one */
  transcode: typeof transcode
  seams?: TranscodeSeams
}

interface KeyState { key: EncodeKey; readied: boolean }
interface Pick { item: MediaItem; id: string }

const unavailable = { ok: false, reason: 'media-unavailable' } as const

export function createSessionMediaQueue(o: SessionMediaQueueOptions): SessionMediaQueue {
  const { sid, cache } = o
  const ctl = new AbortController()
  const byUnit = new Map<number, MediaItem[]>()
  const byId = new Map<string, MediaItem>()
  const keys = new Map<string, KeyState>()
  const origins = new Map<string, EncodeOrigin>()
  const requests = new Map<string, { item: MediaItem; waiting: Array<(r: MediaRequestResult) => void> }>()
  let order: number[] = []
  let cur = 0
  let ended = false, dead = false
  let running: string | undefined
  let loop: Promise<void> | undefined
  let loaded: Promise<void> | undefined
  let wake: () => void = () => {}
  let woken = new Promise<void>((ok) => { wake = ok })
  const poke = () => { const w = wake; woken = new Promise<void>((ok) => { wake = ok }); w() }

  const rotate = (unit: number) => Array.from({ length: o.units }, (_, i) => (unit + i) % o.units)
  const load = () => (loaded ??= (async () => {
    for (let u = 0; u < o.units && !ended; u++) {
      let items: MediaItem[] = []
      try { items = await o.items(u) } catch { /* a slide whose refs cannot be read has no items */ }
      byUnit.set(u, items)
      for (const it of items) {
        byId.set(it.mediaId, it)
        const id = encodeKeyId(it.key)
        if (!keys.has(id)) keys.set(id, { key: it.key, readied: false })
      }
    }
  })())
  const note = (id: string, origin: EncodeOrigin) => { if (!origins.has(id)) origins.set(id, origin) }
  const answer = (id: string, r: EncodeOutcome): MediaRequestResult => {
    if (!r.ok) return { ok: false, reason: r.reason }
    if (!('path' in r.media)) return { ok: false, reason: r.media.refused }
    return { ok: true, path: r.media.path, result: r.media.result, origin: origins.get(id) ?? 'hit' }
  }
  const settleRequests = (id: string, r: MediaRequestResult) => {
    const q = requests.get(id)
    if (!q) return
    requests.delete(id)
    for (const f of q.waiting) f(r)
  }
  const failAll = () => { for (const id of [...requests.keys()]) settleRequests(id, unavailable) }

  /** slot + part + transcode for `item`, on a miss; `mode` is the starter's. */
  const runFor = (item: MediaItem, mode: 'ahead' | 'request'): EncodeRun => (signal) => {
    return (async () => {
      const s = o.admission.slot(sid, mode, signal)
      if ('refused' in s) return { ok: false as const, reason: s.refused }
      const lease = await s
      let part: { path: string; dispose(): Promise<void> } | undefined
      let out: ReturnType<MediaTemp['reserve']> | undefined
      try {
        part = await o.part(item.partPath)
        if (signal.aborted) throw signal.reason
        out = o.temp.reserve('m4a')
        const r = await o.transcode(o.info, item.job(out.path, part.path), signal, o.seams)
        if (!r.ok) return { ok: false as const, reason: r.reason }
        const file = out
        out = undefined
        return { ok: true as const, path: file.path, result: r, file }
      } finally {
        await out?.dispose().catch(() => {})
        await part?.dispose().catch(() => {})
        lease.release()
      }
    })()
  }

  const nextAhead = (): Pick | undefined => {
    for (const u of order) for (const item of byUnit.get(u) ?? []) {
      const id = encodeKeyId(item.key)
      if (!keys.get(id)!.readied) return { item, id }
    }
    return undefined
  }

  const pump = async () => {
    while (!ended && !dead) {
      const req = requests.entries().next().value as [string, { item: MediaItem }] | undefined
      const pick: Pick | undefined = req ? { item: req[1].item, id: req[0] } : nextAhead()
      if (!pick) { await woken; continue }
      const mode = req ? 'request' : 'ahead'
      const ks = keys.get(pick.id)!
      const a = cache.acquire(sid, pick.item.unit, ks.key, runFor(pick.item, mode), { mode, signal: ctl.signal, cursor: cur, order })
      if ('wait' in a) { await Promise.race([cache.room(), woken]); continue }
      note(pick.id, a.origin)
      running = pick.item.mediaId
      const r = await a.done
      running = undefined
      ks.readied = true
      if (ended) break
      settleRequests(pick.id, answer(pick.id, r))
      if (!r.ok && r.revoked) { dead = true; failAll() }
    }
  }
  const kick = () => { loop ??= pump(); poke() }

  return {
    async start(unit) {
      if (ended || dead) return 0
      if (!Number.isInteger(unit) || unit < 0 || unit >= o.units) throw new RangeError('media: start unit out of range')
      cur = unit
      order = rotate(unit)
      await load()
      if (ended) return 0
      let enqueued = 0
      const seen = new Set<string>()
      for (const u of order) for (const item of byUnit.get(u) ?? []) {
        const id = encodeKeyId(item.key)
        if (seen.has(id)) continue
        seen.add(id)
        const ks = keys.get(id)!
        if (ks.readied) continue
        if (cache.peek(ks.key)) {
          // Cached: take the reference now (within the look-ahead), never enqueued.
          const a = cache.acquire(sid, u, ks.key, runFor(item, 'ahead'), { mode: 'ahead', signal: ctl.signal, cursor: cur, order })
          if (!('wait' in a)) { note(id, a.origin); ks.readied = true }
        } else if (!cache.inflight(ks.key)) enqueued++
      }
      kick()
      return enqueued
    },
    cursor(unit) {
      if (ended || !Number.isInteger(unit) || unit < 0 || unit >= o.units || unit === cur) return
      const prev = cur
      const at = order.indexOf(cur)
      const next = order.length ? order[(at + 1) % order.length] : undefined
      cur = unit
      if (unit !== next) order = rotate(unit) // a jump re-queues from that unit (D1)
      if (unit > prev) cache.passed(sid, unit)
      poke()
    },
    async request(mediaId) {
      if (ended || dead) return unavailable
      await load()
      const item = byId.get(mediaId)
      if (!item || ended || dead) return unavailable
      const id = encodeKeyId(item.key)
      const ks = keys.get(id)!
      if (cache.peek(ks.key) && !requests.has(id)) {
        const a = cache.acquire(sid, item.unit, ks.key, runFor(item, 'request'), { mode: 'request', signal: ctl.signal, cursor: cur, order })
        if (!('wait' in a)) { note(id, a.origin); return answer(id, await a.done) }
      }
      return await new Promise<MediaRequestResult>((ok) => {
        const q = requests.get(id)
        if (q) q.waiting.push(ok)
        else requests.set(id, { item, waiting: [ok] })
        kick()
      })
    },
    async end() {
      if (ended) { await loop; return }
      ended = true
      ctl.abort(new DOMException('media: session ended', 'AbortError'))
      cache.dropSession(sid)
      failAll()
      poke()
      await loop
      // Every encode left with no waiter (this session's departure included): wait until it has settled and cleaned up.
      await cache.quiesce()
    },
    get running() { return running },
  }
}
