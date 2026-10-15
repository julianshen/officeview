/**
 * S10.2 / spec §12.4: the document-scoped main-thread application ledger.
 *
 * One ledger per open document holds every main-owned allocation the viewer
 * keeps, by pool and owner, and its high-water mark. Charges are absolute
 * (`set` replaces the owner's previous amount), so an owner that re-measures
 * never double-counts itself; two owners holding copies of the same data (a
 * transferred reply payload and the row cache decoded from it) both count, as
 * both are live until the payload is released.
 *
 * Pools: the remote row cache, committed view buffers from worker deltas,
 * in-flight reply payloads (drawing replies attributed separately), copy sessions, controls value pages/runs, the
 * layout and merge caches, and an estimate of the owned grid DOM
 * (DOM_ELEMENT_BYTES per owned element; see below). Canvas backing stores are
 * the bitmap pool's, not this one's (§12.4).
 */
export type MainPool = 'rowCache' | 'viewBuffers' | 'replies' | 'drawings' | 'copy' | 'controls' | 'layout' | 'merges' | 'dom' | 'geometry' | 'structure' | 'document' | 'canvasText'
export const MAIN_POOLS: readonly MainPool[] = ['rowCache', 'viewBuffers', 'replies', 'drawings', 'copy', 'controls', 'layout', 'merges', 'dom', 'geometry', 'structure', 'document', 'canvasText']

/**
 * Estimated bytes per owned grid DOM element: the element, its attribute and
 * aria strings, its inline style, the text node and the React fiber that
 * renders it. Calibrated in Chrome 155 with the perf page's calibrateDom
 * (owned-gridcell-shaped div + text span, post-GC heap delta incl. embedder):
 * 2,981-3,184 B per element over 1,001-8,001 elements; rounded up.
 */
export const DOM_ELEMENT_BYTES = 3200
/** Element classes of the viewer DOM (src/components/xlsx/domCost.ts). */
export type DomClass = 'cell' | 'header' | 'other'
/**
 * Per-element bytes by class: the element, its Blink layout/style objects, its
 * React fiber(s) with an alternate after re-renders, its React element and
 * props. Rows are placed by transform (no layout shift), so each row is its own
 * paint layer; rows count as 'other'.
 * Calibrated in Chrome 155 with the perf page's calibrateGrid, which mounts the
 * real GridSemantics tree (cells, header select/sort/filter/more buttons with
 * handlers, status and live regions):
 * - 12 viewports from 400x300 to 2560x1440, rows and columns varied
 *   independently, 2 re-renders each with a moved selection;
 * - median of 3 forced-GC heap deltas per viewport;
 * - least-squares fit with a fixed term (~390 KiB, mostly the calibration's
 *   own two 1M-row visibility bitsets).
 * The fit (after the layout-shift fix) gave cell 3,847 B, header 4,337 B and
 * other 5,667 B, with residuals within +-104 KiB. Rounded up.
 */
export const DOM_CLASS_BYTES: Readonly<Record<DomClass, number>> = Object.freeze({ cell: 3900, header: 4400, other: 5700 })

/**
 * Main-heap bytes of Chrome's text-shaping cache held by a 2D context: one
 * blink::FrameShapeCache per font a context used, holding a shaped
 * PlainTextNode per distinct (text, direction) it drew or measured, for as long
 * as the context lives. Calibrated from Chrome 155 heap snapshots:
 * - after the F2 scroll: 111 caches, 8,905 entries, 14.9 MiB;
 * - at F4 loaded: 25 caches, 327 entries.
 * Per entry: PlainTextNode 1,240 B, ShapeResult and run records ~120 B, ~8 B of
 * glyph data per UTF-16 unit. Per cache: its hash table, ~218 B per bucket, at
 * least 64 buckets, grown to keep the load at or below 3/4.
 */
export const SHAPE_ENTRY_BYTES = 1360
export const SHAPE_UNIT_BYTES = 8
export const SHAPE_BUCKET_BYTES = 218
export const SHAPE_MIN_BUCKETS = 64
/**
 * Main-heap bytes of one retained tile canvas outside its pixels: the element,
 * its 2D context and the context's resource bookkeeping. Measured in Chrome 155:
 * - in the REAL viewer (scripts/calibrate-viewer-tiles.ts: InteractiveSheetDoc
 *   scrolled through text-free tiles, heap growth the ledger does not explain
 *   per added tile): 2,117-4,411 B per tile over 5 rounds, median 2,662 B;
 * - on about:blank (512x512 canvases with a 2D context, one fill and one stroke
 *   each, held): 4,529 B per canvas at 100 canvases, 3,695-3,706 B at 200-400.
 * Charged at the highest figure, rounded up.
 */
export const TILE_CANVAS_BYTES = 4600
/** Estimated shaping-cache bytes of one font's cache holding `entries` texts of `units` UTF-16 units in all. */
export function shapeCacheBytes(entries: number, units: number): number {
  if (entries <= 0) return 0
  let buckets = SHAPE_MIN_BUCKETS
  while (entries * 4 > buckets * 3) buckets *= 2
  return entries * SHAPE_ENTRY_BYTES + units * SHAPE_UNIT_BYTES + buckets * SHAPE_BUCKET_BYTES
}

export interface MainLedgerSnapshot {
  bytes: number
  peakBytes: number
  byPool: Record<MainPool, number>
  /** Pool totals at the moment the high-water mark was reached. */
  atPeak: Record<MainPool, number>
}

const emptyPools = (): Record<MainPool, number> => ({ rowCache: 0, viewBuffers: 0, replies: 0, drawings: 0, copy: 0, controls: 0, layout: 0, merges: 0, dom: 0, geometry: 0, structure: 0, document: 0, canvasText: 0 })

/** The part of FinalizationRegistry the ledger uses (injectable for tests). */
export interface CollectionWatcher { register(target: object, token: string): void }
let lifetimeSerial = 0
/** One owner's claim on a held charge (chargeWhileHeld). Releasing twice is a no-op. */
export interface LedgerHold { release(): void }
/** A held lifetime charge: its targets, bytes (one or more lifetime tokens: an
 * adoption may be split across reservations) and current holders. */
interface HeldCharge { pool: MainPool; bytes: number; targets: readonly object[]; tokens: string[]; holders: number; ended: boolean }

export class MainMemoryLedger {
  private readonly charges = new Map<string, { pool: MainPool; bytes: number }>()
  /** Lifetime charges: token -> registered targets still alive. */
  private readonly alive = new Map<string, { pool: MainPool; owner: string; remaining: number }>()
  private readonly watcher: CollectionWatcher | null
  constructor(watcher?: CollectionWatcher) {
    const Registry = (globalThis as unknown as { FinalizationRegistry?: new (cleanup: (token: string) => void) => CollectionWatcher }).FinalizationRegistry
    this.watcher = watcher ?? (Registry ? new Registry(token => this.collected(token)) : null)
  }
  /** Test hook / registry callback: one registered target of `token` was collected. */
  collected(token: string): void {
    const entry = this.alive.get(token)
    if (!entry || --entry.remaining > 0) return
    this.alive.delete(token)
    this.lifetimes.delete(token)
    this.set(entry.pool, entry.owner, 0)
  }
  /**
   * Charges `bytes` for as long as ANY of `targets` is reachable: the charge is
   * released only after every target has been garbage-collected. A consumer
   * that keeps a nested page (not the reply that carried it) keeps the charge.
   * Without a collection watcher (very old hosts) the charge is never released
   * early: the ledger over-counts rather than under-counts. At least one target
   * is required: a charge nothing holds could never end (or be settled).
   */
  chargeWhileReachable(pool: MainPool, targets: readonly object[], bytes: number): string {
    if (targets.length === 0) throw new RangeError(`lifetime charge for ${pool} needs at least one target`)
    const token = `life${++lifetimeSerial}`
    const owner = token
    this.set(pool, owner, bytes)
    this.lifetimes.set(token, pool)
    for (const target of targets) if (!this.leases.has(target)) this.leases.set(target, token)
    if (!this.watcher) return token
    this.alive.set(token, { pool, owner, remaining: targets.length })
    for (const target of targets) this.watcher.register(target, token)
    return token
  }
  /**
   * A lifetime charge with explicit owners (spec §§10.1, 12.4 overlap). Until
   * its first hold() it is a reachability charge; once held, it ends when its
   * last holder releases, or, as a safety net for a holder that never
   * releases, when every target is collected. Every consumer that keeps the
   * allocation (a committed display, a copy, a pending operation) takes its
   * own hold. Targets must belong to this allocation alone (never a shared
   * singleton such as the identity order), as any of them names it.
   */
  chargeWhileHeld(pool: MainPool, targets: readonly object[], bytes: number): string {
    const token = this.chargeWhileReachable(pool, targets, bytes)
    this.holdable(pool, targets, token, bytes)
    return token
  }
  private holdable(pool: MainPool, targets: readonly object[], token: string, bytes: number): void {
    // The remaining parts of one split adoption join the charge its first part made.
    const part = this.held.get(targets[0])
    if (part && !part.ended && !part.holders && part.pool === pool && part.targets === targets) { part.tokens.push(token); part.bytes += bytes; return }
    const charge: HeldCharge = { pool, bytes, targets, tokens: [token], holders: 0, ended: false }
    // A target whose earlier held charge ended now names this one.
    for (const target of targets) if (!this.held.get(target) || this.held.get(target)!.ended) this.held.set(target, charge)
  }
  /**
   * Another owner of the held charge `target` belongs to; null when it has none
   * (a reachability-only or unknown allocation). An owner arriving after every
   * earlier one released (a remount, a copy started from a display still on
   * screen) holds the allocation again, so its bytes are charged again.
   */
  hold(target: object): LedgerHold | null {
    const charge = this.held.get(target)
    if (!charge) return null
    if (charge.ended) {
      const token = `life${++lifetimeSerial}`
      charge.ended = false
      charge.tokens = [token]
      this.set(charge.pool, token, charge.bytes)
      this.lifetimes.set(token, charge.pool)
      for (const item of charge.targets) if (!this.hasLifetime(item)) this.leases.set(item, token)
      if (this.watcher) {
        this.alive.set(token, { pool: charge.pool, owner: token, remaining: charge.targets.length })
        for (const item of charge.targets) this.watcher.register(item, token)
      }
    }
    charge.holders++
    let held = true
    return { release: () => {
      if (!held) return
      held = false
      if (--charge.holders === 0 && !charge.ended) this.endHeld(charge)
    } }
  }
  private endHeld(charge: HeldCharge): void { charge.ended = true; for (const token of charge.tokens) this.end(token) }
  private readonly held = new WeakMap<object, HeldCharge>()
  /** Transfers already-admitted bytes without a double charge or refund gap;
   * `held` makes the adopted charge holdable, as chargeWhileHeld. */
  adoptWhileReachable(fromPool: MainPool, fromOwner: string, pool: MainPool, targets: readonly object[], bytes: number, commit?: () => void, held = false): string {
    const key = `${fromPool}\u0000${fromOwner}`, charge = this.charges.get(key)
    if (!targets.length || !Number.isSafeInteger(bytes) || bytes < 0 || !charge || bytes > charge.bytes) throw new RangeError('Invalid lifetime adoption')
    const token = `life${++lifetimeSerial}`
    // Register and publish while the existing reservation still covers every
    // byte. A failure leaves that reservation untouched for its own cleanup.
    if (this.watcher) for (const target of targets) this.watcher.register(target, token)
    commit?.()
    const remaining = charge.bytes - bytes
    if (remaining) this.charges.set(key, { pool: fromPool, bytes: remaining })
    else this.charges.delete(key)
    this.charges.set(`${pool}\u0000${token}`, { pool, bytes })
    this.pools[fromPool] -= bytes
    this.pools[pool] += bytes
    // Total and high-water bytes never change during ownership adoption.
    this.lifetimes.set(token, pool)
    for (const target of targets) if (!this.leases.has(target)) this.leases.set(target, token)
    if (this.watcher) this.alive.set(token, { pool, owner: token, remaining: targets.length })
    if (held) this.holdable(pool, targets, token, bytes)
    return token
  }
  /** Whether an allocation already has a live reachability owner. */
  hasLifetime(target: object): boolean {
    const token = this.leases.get(target)
    return token !== undefined && this.lifetimes.has(token)
  }
  /**
   * Ends the lifetime charge that `target` (any registered target of it) holds,
   * at once: its consumer has decoded the payload into its own structure and
   * that owner now carries the charge. The consumer calls this AFTER charging
   * the owner, so the bytes are never uncharged in between. A payload nobody
   * settles (still held raw) keeps its reachability lease as a safety net.
   * Later collection of the target is then a no-op.
   */
  settle(target: object): void {
    const token = this.leases.get(target)
    if (token === undefined) return
    this.leases.delete(target)
    const charge = this.held.get(target)
    if (charge?.tokens.includes(token)) this.endHeld(charge)
    else this.end(token)
  }
  /** Ends one lifetime charge now; later collection of its targets is a no-op. */
  private end(token: string): void {
    const pool = this.lifetimes.get(token)
    if (pool === undefined) return
    this.lifetimes.delete(token)
    this.alive.delete(token)
    this.set(pool, token, 0)
  }
  private readonly leases = new WeakMap<object, string>()
  private readonly lifetimes = new Map<string, MainPool>()
  private readonly pools = emptyPools()
  private peakPools = emptyPools()
  bytes = 0
  peakBytes = 0

  /** The owner's current bytes in `pool` (absolute); 0 releases it. */
  set(pool: MainPool, owner: string, bytes: number): void {
    if (!Number.isFinite(bytes) || bytes < 0) throw new RangeError(`ledger charge for ${pool}/${owner} must be a finite non-negative byte count`)
    const key = `${pool}\u0000${owner}`
    const old = this.charges.get(key)?.bytes ?? 0
    if (bytes === 0) this.charges.delete(key)
    else this.charges.set(key, { pool, bytes })
    this.pools[pool] += bytes - old
    this.bytes += bytes - old
    if (this.bytes > this.peakBytes) { this.peakBytes = this.bytes; this.peakPools = { ...this.pools } }
  }
  ownerBytes(pool: MainPool, owner: string): number { return this.charges.get(`${pool}\u0000${owner}`)?.bytes ?? 0 }
  release(pool: MainPool, owner: string): void { this.set(pool, owner, 0) }
  /** Releases every owner of `pool` (or of every pool). */
  releaseAll(pool?: MainPool): void {
    for (const [key, charge] of [...this.charges]) if (pool === undefined || charge.pool === pool) this.set(charge.pool, key.slice(key.indexOf('\u0000') + 1), 0)
  }
  snapshot(): MainLedgerSnapshot {
    return { bytes: this.bytes, peakBytes: this.peakBytes, byPool: { ...this.pools }, atPeak: { ...this.peakPools } }
  }
}

/**
 * Conservative byte estimate of a structured-clone payload (a worker reply):
 * ArrayBuffer views by byteLength (each buffer once), strings at 2 B per code
 * unit plus a header, numbers/booleans 8 B, objects/arrays a header plus their
 * slots. Walks at most `maxNodes` values; the remainder is extrapolated from
 * the average so huge payloads stay bounded in cost.
 */
export function estimatePayloadBytes(value: unknown, maxNodes = 200_000): number {
  const seen = new Set<ArrayBufferLike>()
  const stack: unknown[] = [value]
  let bytes = 0, nodes = 0, pending = 0
  while (stack.length) {
    const v = stack.pop()
    if (++nodes > maxNodes) { pending = stack.length + 1; break }
    if (v === null || v === undefined) { bytes += 8; continue }
    switch (typeof v) {
      case 'string': bytes += 16 + v.length * 2; continue
      case 'number': case 'boolean': case 'bigint': bytes += 8; continue
      case 'object': break
      default: bytes += 8; continue
    }
    if (ArrayBuffer.isView(v)) { if (!seen.has(v.buffer)) { seen.add(v.buffer); bytes += v.buffer.byteLength } continue }
    if (v instanceof ArrayBuffer) { if (!seen.has(v)) { seen.add(v); bytes += v.byteLength } continue }
    if (Array.isArray(v)) { bytes += 16 + v.length * 8; for (const item of v) stack.push(item); continue }
    if (v instanceof Map) { bytes += 32 + v.size * 24; for (const [k, item] of v) { stack.push(k); stack.push(item) }; continue }
    if (v instanceof Set) { bytes += 32 + v.size * 16; for (const item of v) stack.push(item); continue }
    const entries = Object.entries(v as Record<string, unknown>)
    bytes += 24 + entries.length * 16
    for (const [, item] of entries) stack.push(item)
  }
  return pending ? Math.ceil(bytes + (bytes / nodes) * pending) : bytes
}

/**
 * Objects a consumer may keep from a structured payload, breadth-first: the
 * root and nested objects/arrays/typed arrays down to `depth` levels, at most
 * the first 8 elements of any array, and at most `max` objects in all. A charge
 * registered on them lasts until every one of them is collected, so a consumer
 * that keeps one of these sub-objects keeps the charge; an object outside the
 * caps (deeper, past an array's eighth element, or past `max`) does not. Today
 * no consumer keeps such a detached sub-object: the row cache, views and value
 * index decode and settle, and a kept scan page is within the caps.
 */
export function leaseTargets(value: unknown, depth = 4, max = 64): object[] {
  const out: object[] = []
  const queue: { v: unknown; d: number }[] = [{ v: value, d: 0 }]
  const seen = new Set<object>()
  while (queue.length && out.length < max) {
    const { v, d } = queue.shift()!
    if (!v || typeof v !== 'object' || seen.has(v)) continue
    seen.add(v)
    out.push(v)
    if (d >= depth || ArrayBuffer.isView(v) || v instanceof ArrayBuffer) continue
    const children = Array.isArray(v) ? v.slice(0, 8) : Object.values(v as Record<string, unknown>)
    for (const child of children) queue.push({ v: child, d: d + 1 })
  }
  return out
}

/** Retained bytes of a sparse axis: its typed summaries plus one segment record (~72 B) each. */
export function axisBytes(axis: object | undefined): number {
  const d = (axis as { diagnostics?: { summaryBytes?: number; segmentCount?: number } } | undefined)?.diagnostics
  return d ? (d.summaryBytes ?? 0) + (d.segmentCount ?? 0) * 72 : 0
}
/** Retained bytes of a row projection: its typed arrays plus one section record (~96 B) each. */
export function projectionBytes(metrics: object | undefined): number {
  const d = (metrics as { diagnostics?: { typedArrayBytes?: number; sectionCount?: number } } | undefined)?.diagnostics
  return d ? (d.typedArrayBytes ?? 0) + (d.sectionCount ?? 0) * 96 : 0
}
