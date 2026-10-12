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
 * in-flight reply payloads, copy sessions, controls value pages/runs, the
 * layout and merge caches, and an estimate of the owned grid DOM
 * (DOM_ELEMENT_BYTES per owned element; see below). Canvas backing stores are
 * the bitmap pool's, not this one's (§12.4).
 */
export type MainPool = 'rowCache' | 'viewBuffers' | 'replies' | 'copy' | 'controls' | 'layout' | 'merges' | 'dom' | 'geometry' | 'structure' | 'document'
export const MAIN_POOLS: readonly MainPool[] = ['rowCache', 'viewBuffers', 'replies', 'copy', 'controls', 'layout', 'merges', 'dom', 'geometry', 'structure', 'document']

/**
 * Estimated bytes per owned grid DOM element: the element, its attribute and
 * aria strings, its inline style, the text node and the React fiber that
 * renders it. Calibrated in Chrome 155 with the perf page's calibrateDom
 * (owned-gridcell-shaped div + text span, post-GC heap delta incl. embedder):
 * 2,981-3,184 B per element over 1,001-8,001 elements; rounded up.
 */
export const DOM_ELEMENT_BYTES = 3200

export interface MainLedgerSnapshot {
  bytes: number
  peakBytes: number
  byPool: Record<MainPool, number>
  /** Pool totals at the moment the high-water mark was reached. */
  atPeak: Record<MainPool, number>
}

const emptyPools = (): Record<MainPool, number> => ({ rowCache: 0, viewBuffers: 0, replies: 0, copy: 0, controls: 0, layout: 0, merges: 0, dom: 0, geometry: 0, structure: 0, document: 0 })

/** The part of FinalizationRegistry the ledger uses (injectable for tests). */
export interface CollectionWatcher { register(target: object, token: string): void }
let lifetimeSerial = 0

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
   * early: the ledger over-counts rather than under-counts.
   */
  chargeWhileReachable(pool: MainPool, targets: readonly object[], bytes: number): string {
    const token = `life${++lifetimeSerial}`
    const owner = token
    this.set(pool, owner, bytes)
    if (targets.length) this.lifetimes.set(token, pool)
    for (const target of targets) if (!this.leases.has(target)) this.leases.set(target, token)
    if (!this.watcher || targets.length === 0) return token
    this.alive.set(token, { pool, owner, remaining: targets.length })
    for (const target of targets) this.watcher.register(target, token)
    return token
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
 * Objects a consumer may keep from a structured payload: the root and nested
 * objects/arrays/typed arrays down to `depth`, at most `max`. Registering them
 * all keeps a reply's charge until no part of it is reachable.
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
