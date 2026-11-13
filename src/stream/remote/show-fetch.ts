/**
 * §12a.7 show slide cache and request budget (pure; lazy remote path only).
 *
 * A show paces its unit requests against the server's `'unit'` rate class
 * (60/min per sid; third breach in 60 s closes 4008). The budget is the
 * session's (R28, `io.budget`): one sliding log per StreamSession, stamped by
 * the session on every server-charged unit send of any caller (viewer paints,
 * this and earlier shows, validation and resume re-sends), so a show restart
 * or a resume never starts from an empty window. `charge()` (R16, once per
 * charged send) only tells the fetcher a request has gone out: a request
 * issued but not yet charged holds a share of the budget until it does.
 * Without `io.budget` (standalone use) the fetcher keeps a private log of its
 * own charges, `budgetPerMin` (48) per 60 s. The first `want` entry (the navigation target) is always next
 * in line, paced (ruling R14): a new target goes out at once when no target
 * went out in the last TARGET_GAP_MS, else once it has been unchanged for
 * TARGET_SETTLE_MS, so a held arrow costs about one request per 400 ms and
 * the slide it stops on goes out within 250 ms. Prefetch waits until the
 * target is cached or has failed and the wanted set has been unchanged for
 * PREFETCH_SETTLE_MS, and is issued only while more than RESERVE requests
 * remain in the budget. A `rate-limited` refusal pauses issuing BACKOFF_MS;
 * the session's log then holds every send until the server window has room.
 * Units no longer wanted, or requested at a replaced bucket, are aborted; a
 * result that still arrives is closed, never kept. Bitmaps live in an LRU of
 * `maxCached`; a bucket change keeps old bitmaps as fallbacks until replaced.
 * Every bitmap removed (evicted, replaced, dropped, disposed) is closed.
 */
import { StreamSessionError, type StreamErrorCode } from './errors'
import { UnitSendLog, type UnitBudgetView } from './unit-budget'

/** The error codes `onError` reports (the session's typed codes). */
export type RemoteErrorCode = StreamErrorCode

/** Something drawable via `drawImage` (a canvas or an ImageBitmap); `close` frees it. */
export interface ImageBitmapLike {
  readonly width: number
  readonly height: number
  close?(): void
}

export interface ShowFetchIo {
  /** Paints `unit`; calls `charge()` once per unit request it sends (the server charges each). */
  paint(unit: number, scale: number, dpr: number, signal: AbortSignal, charge: () => void): Promise<ImageBitmapLike>
  /** R28: the session's unit send budget (it stamps every charged send itself). */
  budget?: UnitBudgetView
}

export interface ShowFetchClock {
  now(): number
  /** Calls `cb` after `ms`; returns the cancel. Absent ⇒ the global setTimeout. */
  wake?(ms: number, cb: () => void): () => void
}

export interface ShowFetcher {
  /** Bitmap for `unit` at the current bucket, or the best cached one at another bucket (scaled), or undefined. */
  cached(unit: number): ImageBitmapLike | undefined
  /** Wanted set changes on every navigation: `[target, ...prefetch]`, target first. */
  want(units: readonly number[]): void
  setBucket(scale: number, dpr: number): void
  onReady(cb: (unit: number) => void): () => void
  onError(cb: (unit: number, code: RemoteErrorCode) => void): () => void
  dispose(): void
}

const WINDOW_MS = 60_000
/** Requests kept back from prefetch for navigation targets. */
const RESERVE = 8
/** A failed unit is not re-issued sooner; a `rate-limited` refusal pauses all issuing this long. */
const BACKOFF_MS = 2_000
/** R14 pacing: a target is issued at once if no target went out this recently… */
const TARGET_GAP_MS = 400
/** …else once it has been the target this long. */
const TARGET_SETTLE_MS = 250
/** Prefetch waits for the wanted set to be unchanged this long. */
const PREFETCH_SETTLE_MS = 1_000

interface Entry { readonly unit: number; readonly key: string; readonly bmp: ImageBitmapLike }
interface Req { readonly key: string; readonly ctl: AbortController; charged: boolean }

const bucketKey = (scale: number, dpr: number) => `${scale}@${dpr}`
const close = (b: ImageBitmapLike) => { try { b.close?.() } catch { /* already released */ } }

export function createShowFetcher(io: ShowFetchIo, clock: ShowFetchClock, opts: { budgetPerMin?: number; maxCached?: number } = {}): ShowFetcher {
  const budget = opts.budgetPerMin ?? 48
  const maxCached = opts.maxCached ?? 8
  const wake = clock.wake ?? ((ms: number, cb: () => void) => { const h = setTimeout(cb, ms); return () => clearTimeout(h) })

  let disposed = false
  let wanted: number[] = []
  /** When `wanted[0]` / the wanted set last changed. */
  let targetSince = -Infinity
  let wantedSince = -Infinity
  let lastTargetIssue = -Infinity
  /** No bucket ⇒ nothing is issued (the caller sets one before the first paint). */
  let bucket: { scale: number; dpr: number; key: string } | undefined
  /** Standalone only: a private log of this fetcher's own charges. */
  const own = io.budget ? undefined : new UnitSendLog(() => clock.now(), budget, WINDOW_MS)
  const shared: UnitBudgetView = io.budget ?? own!
  /** `${unit}|${bucketKey}` → entry; Map order is LRU order (oldest first). */
  const cache = new Map<string, Entry>()
  const inflight = new Map<number, Req>()
  /** Unit → when its last request failed; cleared by a `want` ≥ BACKOFF_MS later. */
  const failed = new Map<number, number>()
  let pausedUntil = -Infinity
  let pendingWake: { at: number; cancel: () => void } | undefined
  const readyCbs = new Set<(unit: number) => void>()
  const errorCbs = new Set<(unit: number, code: RemoteErrorCode) => void>()

  const entryOf = (unit: number, key: string) => cache.get(`${unit}|${key}`)

  const schedule = (now: number, at: number) => {
    if (pendingWake && pendingWake.at <= at) return
    pendingWake?.cancel()
    const w = { at, cancel: () => {} }
    pendingWake = w
    w.cancel = wake(Math.max(0, at - now), () => { if (pendingWake === w) pendingWake = undefined; pump() })
  }

  const remove = (id: string, e: Entry) => { cache.delete(id); close(e.bmp) }

  const store = (unit: number, key: string, bmp: ImageBitmapLike) => {
    for (const [id, e] of cache) if (e.unit === unit) remove(id, e) // replaces any older bucket's fallback
    cache.set(`${unit}|${key}`, { unit, key, bmp })
    while (cache.size > maxCached) {
      // Oldest entry of a unit not wanted; only when every entry is wanted, the oldest.
      let victim: [string, Entry] | undefined
      for (const kv of cache) if (!wanted.includes(kv[1].unit)) { victim = kv; break }
      victim ??= cache.entries().next().value!
      remove(victim[0], victim[1])
    }
  }

  const issue = (unit: number) => {
    const b = bucket!
    const req: Req = { key: b.key, ctl: new AbortController(), charged: false }
    inflight.set(unit, req)
    // The session stamps its own log; a private log is stamped even after an abort
    // (a send that went out was charged by the server).
    const charge = () => { req.charged = true; if (!disposed) own?.record() }
    let p: Promise<ImageBitmapLike>
    try { p = io.paint(unit, b.scale, b.dpr, req.ctl.signal, charge) } catch (e) { p = Promise.reject(e) }
    // An aborted request is already out of `inflight`; whatever it settles with is ours to drop.
    p.then((bmp) => {
      if (disposed || req.ctl.signal.aborted) return close(bmp)
      inflight.delete(unit)
      store(unit, req.key, bmp)
      pump()
      dispatch(readyCbs, (cb) => cb(unit))
    }, (err: unknown) => {
      if (disposed || req.ctl.signal.aborted) return
      inflight.delete(unit)
      const code: RemoteErrorCode = err instanceof StreamSessionError ? err.code : 'unit-failed'
      // A session refusal pauses every request (a third breach closes the socket);
      // the unit itself is not failed and is re-issued when the pause ends.
      if (code === 'rate-limited') pausedUntil = clock.now() + BACKOFF_MS
      else failed.set(unit, clock.now())
      pump()
      dispatch(errorCbs, (cb) => cb(unit, code))
    })
  }

  /** Calls each listener still subscribed, stopping once a listener disposes the fetcher. */
  function dispatch<T>(set: Set<T>, call: (cb: T) => void) {
    for (const cb of [...set]) {
      if (disposed) return
      if (set.has(cb)) call(cb)
    }
  }

  function pump() {
    if (disposed) return
    const now = clock.now()
    for (const [unit, req] of inflight) {
      if (req.key !== bucket?.key || !wanted.includes(unit)) { inflight.delete(unit); req.ctl.abort() }
    }
    if (!bucket || !wanted.length) return
    if (now < pausedUntil) return schedule(now, pausedUntil)
    const key = bucket.key
    const settled = (unit: number) => !!entryOf(unit, key) || failed.has(unit)
    const outOfBudget = (reserve: number) => {
      let unsent = 0 // issued, not yet sent: their first send is spoken for
      for (const req of inflight.values()) if (!req.charged) unsent++
      if (shared.room(unsent) > reserve) return false
      const ms = shared.waitMs(unsent + reserve)
      // Shared: requests waiting in the session free room without telling the fetcher, so look again.
      if (Number.isFinite(ms)) schedule(now, now + ms)
      else if (!own) schedule(now, now + BACKOFF_MS)
      return true
    }
    const target = wanted[0]!
    if (!settled(target) && !inflight.has(target)) {
      if (now - lastTargetIssue < TARGET_GAP_MS && now - targetSince < TARGET_SETTLE_MS) return schedule(now, targetSince + TARGET_SETTLE_MS)
      if (outOfBudget(0)) return
      lastTargetIssue = now
      issue(target)
    }
    if (!settled(target)) return
    if (now - wantedSince < PREFETCH_SETTLE_MS) return schedule(now, wantedSince + PREFETCH_SETTLE_MS)
    for (const unit of wanted.slice(1)) {
      if (settled(unit) || inflight.has(unit)) continue
      if (outOfBudget(RESERVE)) return
      issue(unit)
    }
  }

  return {
    cached(unit) {
      let id = bucket ? `${unit}|${bucket.key}` : undefined
      let e = id === undefined ? undefined : cache.get(id)
      if (!e) {
        for (const [k, c] of cache) if (c.unit === unit && (!e || c.bmp.width > e.bmp.width)) { id = k; e = c }
      }
      if (!e || id === undefined) return undefined
      cache.delete(id); cache.set(id, e) // most recently used
      return e.bmp
    },
    want(units) {
      if (disposed) return
      const now = clock.now()
      const next: number[] = []
      for (const u of units) if (Number.isInteger(u) && u >= 0 && !next.includes(u)) next.push(u)
      for (const u of next) { const at = failed.get(u); if (at !== undefined && now - at >= BACKOFF_MS) failed.delete(u) }
      if (next[0] !== wanted[0]) targetSince = now
      if (next.length !== wanted.length || next.some((u, i) => u !== wanted[i])) wantedSince = now
      wanted = next
      pump()
    },
    setBucket(scale, dpr) {
      if (!(scale > 0) || !(dpr > 0) || !Number.isFinite(scale) || !Number.isFinite(dpr)) throw new RangeError(`bad bucket ${scale}@${dpr}`)
      if (disposed) return
      const key = bucketKey(scale, dpr)
      if (bucket?.key === key) return
      bucket = { scale, dpr, key }
      pump()
    },
    onReady(cb) { if (disposed) return () => {}; readyCbs.add(cb); return () => { readyCbs.delete(cb) } },
    onError(cb) { if (disposed) return () => {}; errorCbs.add(cb); return () => { errorCbs.delete(cb) } },
    dispose() {
      if (disposed) return
      disposed = true
      pendingWake?.cancel()
      pendingWake = undefined
      for (const req of inflight.values()) req.ctl.abort()
      inflight.clear()
      for (const e of cache.values()) close(e.bmp)
      cache.clear()
      readyCbs.clear()
      errorCbs.clear()
      wanted = []
    },
  }
}
