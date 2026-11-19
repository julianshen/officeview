/**
 * Piece 5 (plan D12, D23): the server's one shared encode cache, its references, single flight and eviction.
 *
 * An entry is keyed by `encodeKeyId(key)` (D23: `{docId, sourceHash, paramsSha, ffmpegVersion}`), so every session of
 * one document shares one encode per part, and two documents never share one. It is in flight (one `EncodeRun`, its
 * waiters) or done (a file, or a cached refusal).
 *
 * - **Single flight.** `acquire` answers `hit` (done), `joined` (in flight) or `miss` (absent: this call starts the
 *   run). Each `acquire` that returns a `done` promise is a waiter; the run's signal fires when the last waiter has
 *   gone (its own signal, `dropSession`), never because its starter left. An abandoned flight leaves the map at once,
 *   so a later `acquire` starts afresh; its late settle is discarded (an ok output is deleted).
 * - **What is cached** (an allowlist): an ok output, and the refusals `media-too-large` and `media-unreadable`
 *   (content-determined; answered for every session without re-encoding, X6). Anything else from the run
 *   (`media-timeout`, `media-busy`, `media-unavailable`, a throw) goes to every waiter and removes the flight.
 * - **References** are `(sid, key)` pairs, taken by `acquire` and released by `delivered`, `passed`, a request's
 *   drop under pressure, or `dropSession`. A cached refusal takes none (nothing to deliver). A shared entry counts in
 *   full against every referencing session's look-ahead (`sessionBytes`); an entry in flight counts its 32 MiB
 *   reservation there and against the server total (`serverBytes`) until it settles to its real size.
 * - **Pressure** (D12). Encode-ahead never drops or evicts: over its look-ahead, or (on a miss) without server room,
 *   it gets `{ wait: true }`. A request first drops the session's own farthest-ahead references to done entries
 *   (distance from `cursor` along `order`; ties: the latest-taken first; never the requested key), then, on a miss,
 *   evicts unreferenced done entries least-recently-used first, then answers `media-busy`. A request still over its
 *   look-ahead after dropping everything droppable proceeds: only the server cap refuses (D12 "no deadlock").
 * - **Idle TTL.** A done entry with no reference is deleted `idleTtlMs` after its last use or release, on the injected
 *   clock. The sweep runs at the top of every public call and on an unref'd timer; referenced entries never expire.
 * - **Revocation.** `dropDocument` deletes every entry of the document at once, referenced or not, aborts its
 *   flights, and answers their waiters `media-unavailable`. A delivery already reading a done file must check
 *   `peek(key)` between chunks: once dropped it answers undefined, and the delivery fails `media-unavailable`.
 *
 * Every entry file is one `MediaTemp` ledger row, owned by the cache (`EncodeRun`'s `file`, else adopted).
 */
import { unlink } from 'node:fs/promises'
import type { MediaRefusal } from '../../../src/stream/media-types'
import { encodeKeyId, type EncodeKey } from './params'
import type { MediaFile, MediaTemp } from './temp'
import { AUDIO_MAX_BYTES, type EncodeResult } from './transcode'

export const MEDIA_IDLE_TTL_MS = 600_000
export const MEDIA_SESSION_LOOKAHEAD = 256 << 20, MEDIA_SERVER_CACHE = 2 * 1024 ** 3 /* not 2 << 30: that is -2^31 in int32 */
/** What an encode in flight reserves against the look-ahead and the server cache (the audio output cap). */
export const MEDIA_INFLIGHT_RESERVE = AUDIO_MAX_BYTES

export type CachedMedia = { path: string; bytes: number; result: EncodeResult & { ok: true } } | { refused: 'media-too-large' | 'media-unreadable'; bytes: 0 }
export type EncodeOrigin = 'hit' | 'joined' | 'miss'
/** slot + part + transcode, built by the queue; called only on a miss. `file`: the output's ledger handle, which the
 * cache takes over (absent: the cache adopts `path`). */
export type EncodeRun = (signal: AbortSignal /* fires when the last waiter is gone */) => Promise<{ ok: true; path: string; result: EncodeResult & { ok: true }; file?: MediaFile } | { ok: false; reason: MediaRefusal }>
export type EncodeOutcome = { ok: true; media: CachedMedia } | { ok: false; reason: MediaRefusal }
export type Acquired = { wait: true /* ahead only: no room under the look-ahead or the server cap */ } | { origin: EncodeOrigin; done: Promise<EncodeOutcome> }

export interface EncodeCache {
  /** Takes the (sid, key) reference and a waiter; single flight per key. */
  acquire(sid: string, unit: number, key: EncodeKey, run: EncodeRun, o: { mode: 'ahead' | 'request'; signal: AbortSignal; cursor: number; order: readonly number[] }): Acquired
  /** Done entries only; no reference. */
  peek(key: EncodeKey): CachedMedia | undefined
  /** Whether the key is encoding now (a later `acquire` joins it). */
  inflight(key: EncodeKey): boolean
  /** Releases that session's reference. */
  delivered(sid: string, key: EncodeKey): void
  /** Releases references of slides the viewer has passed (`unit < beforeUnit`). */
  passed(sid: string, beforeUnit: number): void
  /** Releases every reference and waiter of the session; aborts encodes left with no waiter; leaves entries cached. */
  dropSession(sid: string): void
  /** Revocation or close: deletes every entry with this docId at once, referenced or not; aborts its in-flight encodes
   * (resolves once they have settled); a delivery in flight then fails media-unavailable. */
  dropDocument(docId: string): Promise<void>
  /** Resolves on the next release, settle or eviction, for waiting encode-ahead. */
  room(): Promise<void>
  /** Deletes expired entries now (the timer and every public call run it too). */
  sweep(): Promise<void>
  stats(): { entries: number; bytes: number; referenced: number; inflight: number }
  referencesOf(sid: string): number
  /** Deletes every entry file, aborts every flight and resolves once they have settled. */
  close(): Promise<void>
}

const CACHED_REFUSALS: ReadonlySet<MediaRefusal> = new Set(['media-too-large', 'media-unreadable'])

interface Waiter { sid: string; settle(o: EncodeOutcome): void }
interface Entry {
  id: string
  key: EncodeKey
  state: 'inflight' | 'done'
  bytes: number
  media?: CachedMedia
  file?: MediaFile
  touched: number
  refs: Set<string>
  waiters: Set<Waiter>
  ctl?: AbortController
}
interface Ref { unit: number; seq: number }

export function createEncodeCache(o: { temp: MediaTemp; sessionBytes: number; serverBytes: number; idleTtlMs?: number /* default MEDIA_IDLE_TTL_MS */; now?: () => number /* test seam; the sweep runs on it */ }): EncodeCache {
  const now = o.now ?? Date.now
  const ttl = o.idleTtlMs ?? MEDIA_IDLE_TTL_MS
  const entries = new Map<string, Entry>()
  const sessions = new Map<string, Map<string, Ref>>()
  const runs = new Map<Promise<void>, string /* docId */>()
  let rooms: Array<() => void> = []
  let seq = 0
  let closed = false

  const notify = () => { const r = rooms; rooms = []; for (const f of r) f() }
  const serverTotal = () => { let n = 0; for (const e of entries.values()) n += e.bytes; return n }
  const sessionTotal = (sid: string) => { let n = 0; for (const id of sessions.get(sid)?.keys() ?? []) n += entries.get(id)?.bytes ?? 0; return n }
  const refsOf = (sid: string) => { let m = sessions.get(sid); if (!m) sessions.set(sid, m = new Map()); return m }
  const disposeFile = (e: Entry) => (e.file ? e.file.dispose().catch(() => {}) : Promise.resolve())

  const release = (sid: string, id: string) => {
    const m = sessions.get(sid)
    if (!m?.delete(id)) return
    if (!m.size) sessions.delete(sid)
    const e = entries.get(id)
    if (e) { e.refs.delete(sid); e.touched = Math.max(e.touched, now()) }
  }
  /** Removes `e` from the map and every session; its waiters get `reason`; its flight is aborted. */
  const remove = (e: Entry, reason: MediaRefusal): Promise<void> => {
    if (entries.get(e.id) !== e) return Promise.resolve()
    entries.delete(e.id)
    for (const sid of e.refs) { const m = sessions.get(sid); m?.delete(e.id); if (m && !m.size) sessions.delete(sid) }
    e.refs.clear()
    const ws = [...e.waiters]; e.waiters.clear()
    for (const w of ws) w.settle({ ok: false, reason })
    e.ctl?.abort(new DOMException('media: encode abandoned', 'AbortError'))
    notify()
    return disposeFile(e)
  }
  const expired = (e: Entry, t: number) => e.state === 'done' && e.refs.size === 0 && t - e.touched >= ttl
  const sweep = (): Promise<void> => {
    const t = now()
    const gone: Promise<void>[] = []
    for (const e of [...entries.values()]) if (expired(e, t)) gone.push(remove(e, 'media-unavailable'))
    return Promise.all(gone).then(() => {})
  }
  const timer = setInterval(() => { void sweep() }, Math.max(1000, Math.min(ttl, 60_000)))
  ;(timer as { unref?: () => void }).unref?.()

  /** D12 rule 2: the session's own farthest-ahead references to done entries, until `need` fits. */
  const dropFarthest = (sid: string, need: number, cursor: number, order: readonly number[], except: string) => {
    const m = sessions.get(sid)
    if (!m) return
    const n = order.length
    const at = order.indexOf(cursor)
    const dist = (unit: number) => { const i = order.indexOf(unit); return i < 0 || at < 0 ? Number.POSITIVE_INFINITY : (i - at + n) % n }
    const cands = [...m.entries()]
      .filter(([id]) => id !== except && entries.get(id)?.state === 'done' && entries.get(id)!.bytes > 0)
      .sort(([, a], [, b]) => (dist(b.unit) - dist(a.unit)) || (b.seq - a.seq))
    for (const [id] of cands) {
      if (sessionTotal(sid) + need <= o.sessionBytes) break
      release(sid, id)
    }
  }
  /** D12: unreferenced done entries, least-recently-used first, until `need` fits under the server cap. */
  const evict = (need: number) => {
    const cands = [...entries.values()].filter((e) => e.state === 'done' && e.refs.size === 0 && e.bytes > 0).sort((a, b) => a.touched - b.touched)
    for (const e of cands) {
      if (serverTotal() + need <= o.serverBytes) break
      void remove(e, 'media-unavailable')
    }
  }

  const start = (e: Entry, run: EncodeRun) => {
    const ctl = new AbortController()
    e.ctl = ctl
    const p = (async () => {
      let r: Awaited<ReturnType<EncodeRun>> | undefined
      try { r = await run(ctl.signal) } catch { r = undefined }
      if (entries.get(e.id) !== e) { // abandoned, dropped or closed meanwhile: discard
        if (r?.ok) await (r.file ? r.file.dispose() : unlink(r.path)).catch(() => {})
        return
      }
      e.ctl = undefined
      if (r?.ok) {
        const bytes = r.result.bytes
        e.state = 'done'; e.bytes = bytes; e.touched = now()
        e.media = { path: r.path, bytes, result: r.result }
        try { e.file = r.file ?? o.temp.adopt(r.path, bytes) } catch { await remove(e, 'media-unavailable'); return }
        const ws = [...e.waiters]; e.waiters.clear()
        for (const w of ws) w.settle({ ok: true, media: e.media })
        notify()
        return
      }
      const reason: MediaRefusal = r ? r.reason : 'media-unavailable'
      if (!CACHED_REFUSALS.has(reason)) { await remove(e, reason); return }
      e.state = 'done'; e.bytes = 0; e.touched = now()
      e.media = { refused: reason as 'media-too-large' | 'media-unreadable', bytes: 0 }
      for (const sid of e.refs) { const m = sessions.get(sid); m?.delete(e.id); if (m && !m.size) sessions.delete(sid) }
      e.refs.clear()
      const ws = [...e.waiters]; e.waiters.clear()
      for (const w of ws) w.settle({ ok: false, reason })
      notify()
    })()
    runs.set(p, e.key.docId)
    void p.finally(() => runs.delete(p))
  }

  const addWaiter = (e: Entry, sid: string, signal: AbortSignal): Promise<EncodeOutcome> => new Promise((ok) => {
    let settled = false
    const onAbort = () => leave({ ok: false, reason: 'media-unavailable' })
    const w: Waiter = { sid, settle: (r) => { if (settled) return; settled = true; signal.removeEventListener('abort', onAbort); ok(r) } }
    const leave = (r: EncodeOutcome) => {
      if (!e.waiters.delete(w)) return
      w.settle(r)
      if (e.state === 'inflight' && e.waiters.size === 0) void remove(e, 'media-unavailable')
    }
    if (signal.aborted) { ok({ ok: false, reason: 'media-unavailable' }); return }
    e.waiters.add(w)
    signal.addEventListener('abort', onAbort, { once: true })
  })

  const outcome = (e: Entry): EncodeOutcome => {
    const m = e.media!
    return 'refused' in m ? { ok: false, reason: m.refused } : { ok: true, media: m }
  }

  return {
    acquire(sid, unit, key, run, a) {
      void sweep()
      if (closed) return { origin: 'miss', done: Promise.resolve({ ok: false, reason: 'media-unavailable' }) }
      const id = encodeKeyId(key)
      let e = entries.get(id)
      if (e?.state === 'done' && 'refused' in e.media!) { e.touched = Math.max(e.touched, now()); return { origin: 'hit', done: Promise.resolve(outcome(e)) } }
      const m = refsOf(sid)
      const held = m.has(id)
      const need = held ? 0 : (e ? e.bytes : MEDIA_INFLIGHT_RESERVE)
      if (sessionTotal(sid) + need > o.sessionBytes) {
        if (a.mode === 'ahead') { if (!m.size) sessions.delete(sid); return { wait: true } }
        dropFarthest(sid, need, a.cursor, a.order, id)
      }
      if (!e && serverTotal() + MEDIA_INFLIGHT_RESERVE > o.serverBytes) {
        if (a.mode === 'request') evict(MEDIA_INFLIGHT_RESERVE)
        if (serverTotal() + MEDIA_INFLIGHT_RESERVE > o.serverBytes) {
          if (!m.size) sessions.delete(sid)
          return a.mode === 'ahead' ? { wait: true } : { origin: 'miss', done: Promise.resolve({ ok: false, reason: 'media-busy' }) }
        }
      }
      let origin: EncodeOrigin
      const fresh = !e
      if (!e) {
        e = { id, key: { ...key }, state: 'inflight', bytes: MEDIA_INFLIGHT_RESERVE, touched: now(), refs: new Set(), waiters: new Set() }
        entries.set(id, e)
        origin = 'miss'
      } else origin = e.state === 'done' ? 'hit' : 'joined'
      if (!held) { m.set(id, { unit, seq: ++seq }); e.refs.add(sid) }
      e.touched = Math.max(e.touched, now())
      if (e.state === 'done') return { origin, done: Promise.resolve(outcome(e)) }
      const done = addWaiter(e, sid, a.signal)
      if (fresh) {
        if (a.signal.aborted) void remove(e, 'media-unavailable')
        else start(e, run)
      }
      return { origin, done }
    },
    peek(key) {
      void sweep()
      const e = entries.get(encodeKeyId(key))
      return e?.state === 'done' ? e.media : undefined
    },
    inflight(key) { return entries.get(encodeKeyId(key))?.state === 'inflight' },
    delivered(sid, key) { void sweep(); release(sid, encodeKeyId(key)); notify() },
    passed(sid, beforeUnit) {
      void sweep()
      for (const [id, r] of [...(sessions.get(sid)?.entries() ?? [])]) if (r.unit < beforeUnit) release(sid, id)
      notify()
    },
    dropSession(sid) {
      for (const id of [...(sessions.get(sid)?.keys() ?? [])]) release(sid, id)
      sessions.delete(sid)
      for (const e of [...entries.values()]) {
        const mine = [...e.waiters].filter((w) => w.sid === sid)
        for (const w of mine) { e.waiters.delete(w); w.settle({ ok: false, reason: 'media-unavailable' }) }
        if (mine.length && e.state === 'inflight' && e.waiters.size === 0) void remove(e, 'media-unavailable')
      }
      notify()
    },
    async dropDocument(docId) {
      const gone = [...entries.values()].filter((e) => e.key.docId === docId).map((e) => remove(e, 'media-unavailable'))
      const flights = [...runs].filter(([, d]) => d === docId).map(([p]) => p)
      await Promise.all([...gone, ...flights])
    },
    room() { return new Promise<void>((ok) => { rooms.push(ok) }) },
    sweep,
    stats() {
      void sweep()
      let done = 0, referenced = 0, inflight = 0
      for (const e of entries.values()) { if (e.state === 'done') done++; else inflight++; if (e.refs.size) referenced++ }
      return { entries: done, bytes: serverTotal(), referenced, inflight }
    },
    referencesOf(sid) { return sessions.get(sid)?.size ?? 0 },
    async close() {
      closed = true
      clearInterval(timer)
      const gone = [...entries.values()].map((e) => remove(e, 'media-unavailable'))
      sessions.clear()
      notify()
      await Promise.all([...gone, ...runs.keys()])
    },
  }
}
