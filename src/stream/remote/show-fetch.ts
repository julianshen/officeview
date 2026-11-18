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
 * remain in the budget. A track or Morph pair (Task 12, `room(kind)`) needs more than TRACK_RESERVE / MORPH_RESERVE
 * remaining and no target waiting to go out: under it the step plays `step-deferred`, the Morph fades. A `rate-limited` refusal pauses issuing BACKOFF_MS;
 * the session's log then holds every send until the server window has room.
 * Units no longer wanted, or requested at a replaced bucket, are aborted; a
 * result that still arrives is closed, never kept. Bitmaps live in an LRU of
 * `maxCached`; a bucket change keeps old bitmaps as fallbacks until replaced.
 * Every bitmap removed (evicted, replaced, dropped, disposed) is closed.
 *
 * Piece 4 (Task 11): a unit `opts.animated` names is fetched as a show bundle (`io.bundle`, at the show's
 * `opts.motion`) under the same pacing, LRU and budget; `bundle(unit)` reads it as `cached` reads a bitmap.
 * A still refusal (`still`, `show-budget`, `show-limits`, `show-plan-timeout`) makes that unit a piece-3
 * still for the rest of the show: it is re-requested through `paint`. `admission` (the client's own
 * transient refusal) retries the unit after BACKOFF_MS without pausing anything else. Bundles share a
 * per-show byte budget (`opts.showBytes`, 128 MiB, with `opts.extraBytes` — tracks and Morph pairs held
 * elsewhere): over it, the least recently used bundles that are not the target go first; a target over it
 * on its own plays its still (`show-budget`). Every bundle removed is released.
 */
import { StreamSessionError, type StreamErrorCode } from './errors'
import type { ShowBundle } from './show-bundle'
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
  /** Piece 4: one show bundle of an animated `unit` (one charged send; `charge()` once it goes out). */
  bundle?(unit: number, scale: number, dpr: number, motion: 'full' | 'reduced', signal: AbortSignal, charge: () => void): Promise<ShowBundle>
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
  /** Piece 4: the show bundle of `unit` (at the current bucket, else the best cached one), or undefined. */
  bundle(unit: number): ShowBundle | undefined
  /** Σ decodedBytes of the bundles cached. */
  bundleBytes(): number
  /**
   * Room for `bytes` more held outside the cache (a track frame): the least recently used bundles that are not the
   * target are evicted until they fit; `false` when they do not fit even then.
   */
  makeRoom(bytes: number): boolean
  /**
   * PR #78 round 8: `unit` plays its piece-3 still for the rest of the show (as after a still refusal): its bundle
   * is released and the unit is asked for through `paint` (no new wire).
   */
  demote(unit: number): void
  /**
   * A charged request besides the fetcher's own (the current slide's `track`, a Morph pair) may go now: no rate pause,
   * no target waiting to go out, and more budget left than that kind's reserve (strict order: target bundle > track >
   * Morph pair > prefetch). Default `track`.
   */
  room(kind?: 'track' | 'morph'): boolean
  /** Such a request was refused `rate-limited`: every request pauses, as for the fetcher's own. */
  paused(): void
  /** Wanted set changes on every navigation: `[target, ...prefetch]`, target first. */
  want(units: readonly number[]): void
  setBucket(scale: number, dpr: number): void
  onReady(cb: (unit: number) => void): () => void
  onError(cb: (unit: number, code: RemoteErrorCode) => void): () => void
  dispose(): void
}

const WINDOW_MS = 60_000
/**
 * Budget kept back from prefetch: the reserve the navigation target, the current slide's tracks and Morph pairs share.
 * Inside it the order is strict — target (needs any room), track (more than TRACK_RESERVE left), Morph pair (more than
 * MORPH_RESERVE), prefetch (more than RESERVE) — so a run of animated slides starves the optional work first.
 */
export const RESERVE = 8
export const MORPH_RESERVE = 4
export const TRACK_RESERVE = 2
/** A failed unit is not re-issued sooner; a `rate-limited` refusal pauses all issuing this long. */
export const BACKOFF_MS = 2_000
/** R14 pacing: a target is issued at once if no target went out this recently… */
const TARGET_GAP_MS = 400
/** …else once it has been the target this long. */
const TARGET_SETTLE_MS = 250
/** Prefetch waits for the wanted set to be unchanged this long. */
const PREFETCH_SETTLE_MS = 1_000
/** The per-show budget of decoded show assets and plan records (spec §5). */
const SHOW_BYTES = 128 * 1024 * 1024
/** A show refusal that plays the piece-3 still (Task 10). */
const STILL_CODES: ReadonlySet<string> = new Set(['still', 'show-budget', 'show-limits', 'show-plan-timeout'])

interface Entry { readonly unit: number; readonly key: string; readonly bmp?: ImageBitmapLike; readonly bundle?: ShowBundle }
interface Req { readonly key: string; readonly ctl: AbortController; charged: boolean }

const bucketKey = (scale: number, dpr: number) => `${scale}@${dpr}`
const close = (b: ImageBitmapLike) => { try { b.close?.() } catch { /* already released */ } }
const drop = (e: { bmp?: ImageBitmapLike; bundle?: ShowBundle }) => { if (e.bmp) close(e.bmp); if (e.bundle) { try { e.bundle.release() } catch { /* already released */ } } }

export interface ShowFetchOptions {
  budgetPerMin?: number
  maxCached?: number
  /** Piece 4: the units fetched as show bundles (with `io.bundle`). */
  animated?(unit: number): boolean
  motion?: 'full' | 'reduced'
  /** The per-show byte budget of bundles (default 128 MiB). */
  showBytes?: number
  /** Bytes the show holds outside the cache (tracks, Morph pairs) that count against `showBytes`. */
  extraBytes?(): number
}

export function createShowFetcher(io: ShowFetchIo, clock: ShowFetchClock, opts: ShowFetchOptions = {}): ShowFetcher {
  const budget = opts.budgetPerMin ?? 48
  const maxCached = opts.maxCached ?? 8
  const showBytes = opts.showBytes ?? SHOW_BYTES
  /** Units refused a bundle with a still code: piece-3 stills for the rest of the show. */
  const stills = new Set<number>()
  /** Units refused `admission`: re-issued once their time comes (nothing else pauses). */
  const retryAt = new Map<number, number>()
  /**
   * Prefetch candidates evicted for bytes (fix round 1): not asked for again until they become the target or the
   * wanted set changes, so neighbours that cannot coexist are not charged for over and over.
   */
  const evicted = new Set<number>()
  const asBundle = (unit: number) => !!io.bundle && !!opts.animated?.(unit) && !stills.has(unit)
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
  /** Issued, not yet sent: their first send is spoken for. */
  const unsentCount = () => { let n = 0; for (const req of inflight.values()) if (!req.charged) n++; return n }
  /** The navigation target has yet to go out (paced, or out of budget): nothing optional may take its slot. */
  const targetWaiting = () => {
    const t = wanted[0]
    if (t === undefined || !bucket) return false
    return !entryOf(t, bucket.key) && !failed.has(t) && !retryAt.has(t) && !inflight.has(t)
  }

  const schedule = (now: number, at: number) => {
    if (pendingWake && pendingWake.at <= at) return
    pendingWake?.cancel()
    const w = { at, cancel: () => {} }
    pendingWake = w
    w.cancel = wake(Math.max(0, at - now), () => { if (pendingWake === w) pendingWake = undefined; pump() })
  }

  const remove = (id: string, e: Entry) => { cache.delete(id); drop(e) }
  /** Evicts the least recently used non-target bundles until `bytes` more fit the show budget; whether they do. */
  function trim(bytes: number): boolean {
    const over = () => heldBytes() + (opts.extraBytes?.() ?? 0) + bytes > showBytes
    while (over()) {
      let victim: [string, Entry] | undefined
      for (const kv of cache) if (kv[1].bundle && kv[1].unit !== wanted[0]) { victim = kv; break }
      if (!victim) break
      evicted.add(victim[1].unit)
      remove(victim[0], victim[1])
    }
    return !over()
  }
  const heldBytes = () => { let n = 0; for (const e of cache.values()) n += e.bundle?.decodedBytes ?? 0; return n }

  /** Stores the result of `unit`; `false` when a bundle that is the target does not fit the show budget on its own. */
  const store = (unit: number, key: string, got: { bmp?: ImageBitmapLike; bundle?: ShowBundle }): boolean => {
    for (const [id, e] of cache) if (e.unit === unit) remove(id, e) // replaces any older bucket's fallback
    cache.set(`${unit}|${key}`, { unit, key, ...got })
    while (cache.size > maxCached) {
      // Oldest entry of a unit not wanted; only when every entry is wanted, the oldest.
      let victim: [string, Entry] | undefined
      for (const kv of cache) if (!wanted.includes(kv[1].unit)) { victim = kv; break }
      victim ??= cache.entries().next().value!
      remove(victim[0], victim[1])
    }
    if (!got.bundle) return true
    // The per-show budget: the least recently used bundles that are not the target go first.
    const over = () => heldBytes() + (opts.extraBytes?.() ?? 0) > showBytes
    trim(0)
    const id = `${unit}|${key}`
    if (over() && cache.has(id)) {
      // Still over with only the target's bundle left: the target plays its still.
      remove(id, cache.get(id)!)
      stills.add(unit)
      return false
    }
    return cache.has(id)
  }

  const issue = (unit: number) => {
    const b = bucket!
    const req: Req = { key: b.key, ctl: new AbortController(), charged: false }
    inflight.set(unit, req)
    // The session stamps its own log; a private log is stamped even after an abort
    // (a send that went out was charged by the server).
    const charge = () => { req.charged = true; if (!disposed) own?.record() }
    const bundled = asBundle(unit)
    let p: Promise<{ bmp?: ImageBitmapLike; bundle?: ShowBundle }>
    try {
      p = bundled
        ? io.bundle!(unit, b.scale, b.dpr, opts.motion ?? 'full', req.ctl.signal, charge).then((bundle) => ({ bundle }))
        : io.paint(unit, b.scale, b.dpr, req.ctl.signal, charge).then((bmp) => ({ bmp }))
    } catch (e) { p = Promise.reject(e) }
    // An aborted request is already out of `inflight`; whatever it settles with is ours to drop.
    p.then((got) => {
      if (disposed || req.ctl.signal.aborted) return drop(got)
      inflight.delete(unit)
      const kept = store(unit, req.key, got)
      pump()
      if (kept) dispatch(readyCbs, (cb) => cb(unit))
      else dispatch(errorCbs, (cb) => cb(unit, 'show-budget'))
    }, (err: unknown) => {
      if (disposed || req.ctl.signal.aborted) return
      inflight.delete(unit)
      const code: RemoteErrorCode = err instanceof StreamSessionError ? err.code : 'unit-failed'
      // A session refusal pauses every request (a third breach closes the socket);
      // the unit itself is not failed and is re-issued when the pause ends.
      if (code === 'rate-limited') pausedUntil = clock.now() + BACKOFF_MS
      // A still refusal: the unit is a piece-3 still from now on (re-issued as a unit at once).
      else if (bundled && STILL_CODES.has(code)) stills.add(unit)
      // The client's own admission refusal: this unit alone waits.
      else if (code === 'admission') retryAt.set(unit, clock.now() + BACKOFF_MS)
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
    for (const [unit, at] of retryAt) {
      if (at <= now) retryAt.delete(unit)
      else schedule(now, at)
    }
    const settled = (unit: number) => !!entryOf(unit, key) || failed.has(unit) || retryAt.has(unit)
    const outOfBudget = (reserve: number) => {
      const unsent = unsentCount()
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
      if (settled(unit) || inflight.has(unit) || evicted.has(unit)) continue
      if (outOfBudget(RESERVE)) return
      issue(unit)
    }
  }

  /** The entry of `unit` at the current bucket, else the best one cached (`size` ranks them); touched as most recently used. */
  function lookup(unit: number, has: (e: Entry) => boolean, size: (e: Entry) => number): Entry | undefined {
    let id = bucket ? `${unit}|${bucket.key}` : undefined
    let e = id === undefined ? undefined : cache.get(id)
    if (e && !has(e)) e = undefined
    if (!e) {
      for (const [k, c] of cache) if (c.unit === unit && has(c) && (!e || size(c) > size(e))) { id = k; e = c }
    }
    if (!e || id === undefined) return undefined
    cache.delete(id); cache.set(id, e) // most recently used
    return e
  }

  return {
    cached(unit) { return lookup(unit, (e) => !!e.bmp, (e) => e.bmp!.width)?.bmp },
    bundle(unit) { return lookup(unit, (e) => !!e.bundle, (e) => e.bundle!.decodedBytes)?.bundle },
    bundleBytes: heldBytes,
    makeRoom: (bytes) => !disposed && trim(bytes),
    demote(unit) {
      if (disposed || stills.has(unit) || !asBundle(unit)) return
      stills.add(unit)
      const req = inflight.get(unit)
      if (req) { inflight.delete(unit); req.ctl.abort() }
      for (const [id, e] of [...cache]) if (e.unit === unit) remove(id, e)
      pump()
    },
    room(kind = 'track') {
      if (disposed || clock.now() < pausedUntil || targetWaiting()) return false
      return shared.room(unsentCount()) > (kind === 'morph' ? MORPH_RESERVE : TRACK_RESERVE)
    },
    paused() {
      if (disposed) return
      pausedUntil = clock.now() + BACKOFF_MS
      pump()
    },
    want(units) {
      if (disposed) return
      const now = clock.now()
      const next: number[] = []
      for (const u of units) if (Number.isInteger(u) && u >= 0 && !next.includes(u)) next.push(u)
      for (const u of next) { const at = failed.get(u); if (at !== undefined && now - at >= BACKOFF_MS) failed.delete(u) }
      if (next[0] !== wanted[0]) targetSince = now
      if (next.length !== wanted.length || next.some((u, i) => u !== wanted[i])) { wantedSince = now; evicted.clear() }
      if (next[0] !== undefined) evicted.delete(next[0])
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
      for (const e of cache.values()) drop(e)
      cache.clear()
      readyCbs.clear()
      errorCbs.clear()
      wanted = []
    },
  }
}
