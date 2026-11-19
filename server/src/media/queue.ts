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
 *   its part and output files disposed: `cache.quiesce()`), waiting at most `MEDIA_QUIESCE_MS`. Cache entries remain.
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
  /** How long `end()` waits for abandoned encodes to settle (default `MEDIA_QUIESCE_MS`). */
  quiesceMs?: number
}

/** `end()`'s bound on waiting for abandoned encodes: one session's end never hangs on a stuck encode (any session's).
 * A killed group settles in milliseconds; a run still unsettled past this is left to finish (and clean up) alone. */
export const MEDIA_QUIESCE_MS = 5000

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
  /** `gen` counts pokes, so a poke that lands while the loop is busy elsewhere is never lost. */
  let gen = 0
  const poke = () => { gen++; const w = wake; woken = new Promise<void>((ok) => { wake = ok }); w() }

  const rotate = (unit: number) => Array.from({ length: o.units }, (_, i) => (unit + i) % o.units)
  const load = () => (loaded ??= (async () => {
    for (let u = 0; u < o.units && !ended; u++) {
      let items: MediaItem[] = []
      try { items = await o.items(u) } catch { /* a slide whose refs cannot be read has no items */ }
      byUnit.set(u, items)
      for (const it of items) {
        docId ??= it.key.docId
        byId.set(it.mediaId, it)
        const id = encodeKeyId(it.key)
        if (!keys.has(id)) keys.set(id, { key: it.key, readied: false })
      }
    }
  })())
  const note = (id: string, origin: EncodeOrigin) => { if (!origins.has(id)) origins.set(id, origin) }
  let docId: string | undefined
  const revoked = () => dead || (docId !== undefined && cache.revoked(docId))
  const stop = () => { dead = true; failAll(); poke() }
  /** One pending `cache.revocation()` per queue (renewed after it fires), so waits do not pile up resolvers. */
  let revP: Promise<void> | undefined
  const revocation = () => (revP ??= cache.revocation().then(() => { revP = undefined }))
  /** One requester's answer. An ok answer pins the session's reference until its `delivered` (one pin each). */
  const answer = (id: string, key: EncodeKey, r: EncodeOutcome): MediaRequestResult => {
    if (!r.ok) return { ok: false, reason: r.reason }
    if (!('path' in r.media)) return { ok: false, reason: r.media.refused }
    if (!cache.pin(sid, key)) return unavailable
    return { ok: true, path: r.media.path, result: r.media.result, origin: origins.get(id) ?? 'hit' }
  }
  const settleRequests = (id: string, key: EncodeKey, r: EncodeOutcome | undefined) => {
    const q = requests.get(id)
    if (!q) return
    requests.delete(id)
    for (const f of q.waiting) f(r ? answer(id, key, r) : unavailable)
  }
  const failAll = () => { for (const [id, q] of [...requests]) settleRequests(id, q.item.key, undefined) }
  /** Whether a settled acquisition left the key cached (ok or a cached refusal): only then is it readied and its
   * origin recorded. A `media-busy`, `media-timeout` or `media-unavailable` result leaves neither (fix round 1). */
  const cached = (key: EncodeKey, r: EncodeOutcome) => r.ok || cache.peek(key) !== undefined

  /** part + transcode for `item`, on a miss, under `lease` (held by the caller) or a slot taken in `mode`. */
  const runFor = (item: MediaItem, mode: 'ahead' | 'request', held?: { lease: { release(): void }; used: boolean }): EncodeRun => (signal) => {
    if (held) held.used = true
    return (async () => {
      let lease = held?.lease
      if (!lease) {
        const s = o.admission.slot(sid, mode, signal)
        if ('refused' in s) return { ok: false as const, reason: s.refused }
        lease = await s
      }
      let part: { path: string; dispose(): Promise<void> } | undefined
      let out: ReturnType<MediaTemp['reserve']> | undefined
      try {
        if (signal.aborted) throw signal.reason
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

  /** Keys whose encode-ahead failed uncached (timeout, unavailable): not readied (a request encodes again), but not
   * retried by encode-ahead either, which would loop. */
  const aheadFailed = new Set<string>()
  const nextAhead = (): Pick | undefined => {
    for (const u of order) for (const item of byUnit.get(u) ?? []) {
      const id = encodeKeyId(item.key)
      if (!keys.get(id)!.readied && !aheadFailed.has(id)) return { item, id }
    }
    return undefined
  }

  /**
   * Encode-ahead's slot for a likely miss, taken BEFORE the flight exists (fix round 1): a parked encode-ahead holds
   * no cache entry, so a request (this session's or another's) never waits behind it or joins it past admission. The
   * wait gives way to a request, a cursor move or a revocation (undefined).
   */
  const aheadSlot = async (): Promise<{ release(): void } | undefined> => {
    const give = new AbortController()
    const onEnd = () => give.abort()
    ctl.signal.addEventListener('abort', onEnd, { once: true })
    const s = o.admission.slot(sid, 'ahead', give.signal)
    if ('refused' in s) return undefined // never for 'ahead'
    const lease = s.then((l) => l, () => undefined)
    const won = await Promise.race([lease, woken.then(() => 'yield' as const), revocation().then(() => 'yield' as const)])
    ctl.signal.removeEventListener('abort', onEnd)
    if (won === 'yield' || won === undefined) {
      give.abort()
      void lease.then((l) => l?.release()) // granted meanwhile: hand it back
      return undefined
    }
    return won
  }

  const pump = async () => {
    while (!ended && !revoked()) {
      const g = gen
      const req = requests.entries().next().value as [string, { item: MediaItem }] | undefined
      const pick: Pick | undefined = req ? { item: req[1].item, id: req[0] } : nextAhead()
      if (!pick) { if (gen === g) await Promise.race([woken, revocation()]); continue }
      const mode = req ? 'request' : 'ahead'
      const ks = keys.get(pick.id)!
      let held: { lease: { release(): void }; used: boolean } | undefined
      if (mode === 'ahead' && !cache.peek(ks.key) && !cache.inflight(ks.key)) {
        const lease = await aheadSlot()
        if (!lease) continue // a request, cursor move, revocation or end comes first
        if (gen !== g) { lease.release(); continue } // poked while the slot was granted: pick again
        held = { lease, used: false }
      }
      const a = cache.acquire(sid, pick.item.unit, ks.key, runFor(pick.item, mode, held), { mode, signal: ctl.signal, cursor: cur, order })
      if (held && !held.used) held.lease.release() // not a miss after all
      if ('wait' in a) { if (gen === g) await Promise.race([cache.room(), woken, revocation()]); continue }
      running = pick.item.mediaId
      const r = await a.done
      running = undefined
      if (ended) break
      if (cached(ks.key, r)) { note(pick.id, a.origin); ks.readied = true } else if (mode === 'ahead') aheadFailed.add(pick.id)
      if (!r.ok && r.revoked) { settleRequests(pick.id, ks.key, r); stop(); break }
      settleRequests(pick.id, ks.key, r)
    }
    if (revoked()) { dead = true; failAll() }
  }
  const kick = () => { loop ??= pump(); poke() }

  return {
    async start(unit) {
      if (ended || revoked()) return 0
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
          if (!('wait' in a)) {
            const r = await a.done // a hit: settled
            if (cached(ks.key, r)) { note(id, a.origin); ks.readied = true }
          }
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
      if (ended || revoked()) return unavailable
      await load()
      const item = byId.get(mediaId)
      if (!item || ended || revoked()) return unavailable
      const id = encodeKeyId(item.key)
      const ks = keys.get(id)!
      if (cache.peek(ks.key) && !requests.has(id)) {
        const a = cache.acquire(sid, item.unit, ks.key, runFor(item, 'request'), { mode: 'request', signal: ctl.signal, cursor: cur, order })
        if (!('wait' in a)) {
          const r = await a.done
          if (cached(ks.key, r)) { note(id, a.origin); ks.readied = true }
          if (!r.ok && r.revoked) stop()
          return answer(id, ks.key, r)
        }
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
      let timer: ReturnType<typeof setTimeout> | undefined
      await Promise.race([cache.quiesce(), new Promise<void>((ok) => { timer = setTimeout(ok, o.quiesceMs ?? MEDIA_QUIESCE_MS) })])
      clearTimeout(timer)
    },
    get running() { return running },
  }
}
