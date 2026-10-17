/**
 * Ingestion and index construction memory (S13.4, spec §12.4: "peak allocation
 * during sort/import/transfer ... counts, not just steady-state retained data").
 *
 * The paged importer's PageBuilder buffers, pages waiting for their commit,
 * pending shared-string chunks and encode scratch are transient shares of the
 * document MemoryPool; page summaries (the overview index) are persistent,
 * non-evictable entries. Every share is charged before the allocation it
 * covers, and a pool that cannot fit it refuses with a typed `resource-limit`
 * (reason `memory-budget`) instead of over-allocating.
 *
 * The document pool only exists once the viewer session binds it, after the
 * first page committed, so a store starts on a provisional pool of the
 * profile's worker budget. `rebind` moves every live charge into the document
 * pool atomically: either all of them move and the provisional pool closes, or
 * none moves, the old ownership stays intact and bind throws the typed
 * refusal (the viewer then refuses the open with resource-limit).
 *
 * Byte figures are conservative estimates of JS heap held. They are pinned by
 * tests/xlsx-viewer-construction-memory.test.ts against a Bun heap probe
 * (scripts/construction-memory-probe.ts, heapSize after a full GC; heapSize
 * already includes extraMemorySize). Review measurements on Bun 1.3.10:
 * 72.6-83.7 B per numeric builder record, 123.6-127.3 B per record holding a
 * ~12-unit string, 187.4-191.6 B per summary; the constants add headroom for
 * array-capacity growth (up to 1.5x):
 * - builder: 96 B per record, 48 B per non-empty inline string plus 2 B per
 *   UTF-16 code unit of text and spilled values, and per formula-meta object
 *   192 B plus 2 B per character of its JSON (at least its strings' content);
 * - a built page waiting for commit: payloadBytes, which counts the same meta
 *   ownership, plus its spilled values;
 * - encode scratch: the exact encoded buffer plus the meta JSON string, both
 *   sized without allocating and charged before either exists;
 * - a pending shared string: 2 B per code unit plus 64 B;
 * - a page summary: 256 B.
 * They are Bun/JSC figures, not a Chrome heap certification (S10.2).
 */
import { PackageError } from '../../core/package-error'
import type { MemoryPool, PoolCharge, PoolEntry } from '../view/tileCache'

export const BUILDER_RECORD_BYTES = 96
export const TEXT_STRING_BYTES = 48
export const META_OBJECT_BYTES = 192
export const PENDING_STRING_BYTES = 64
/** One PageSummary: the object, its ids, its RectAggregate and its index slot. */
export const PAGE_SUMMARY_BYTES = 256
/** Shares grow in steps so a per-record check rarely touches the pool. */
const STEP_BYTES = 64 * 1024

export function memoryRefusal(what: string, bytes: number, limitBytes: number): PackageError {
  return new PackageError('resource-limit', `The worker memory budget of ${limitBytes} bytes cannot hold ${what} (${bytes} bytes); the import stopped instead of over-allocating. Close other spreadsheets or open a smaller file.`, undefined, 'memory-budget')
}

export interface ConstructionShare {
  /** Bytes currently charged (at least the last successful need). */
  readonly bytes: number
  /** Sets the share's absolute size before allocating it; throws typed when the pool cannot fit it. */
  need(bytes: number): void
  /** Idempotent. */
  release(): void
}

interface Share { what: string; charge: PoolCharge; live: boolean; untilReleased: boolean }

export class ConstructionMemory {
  private pool: MemoryPool
  /** The store's own pool until a document pool is bound; closed on the first rebind. */
  private provisional: MemoryPool | undefined
  private readonly shares = new Set<Share>()
  private readonly persistent = new Map<string, { what: string; bytes: number; entry: PoolEntry }>()

  constructor(provisional: MemoryPool, private readonly docId: string) {
    this.pool = provisional
    this.provisional = provisional
  }

  get transientBytes(): number {
    let n = 0
    for (const s of this.shares) n += s.charge.bytes
    return n
  }
  get persistentBytes(): number {
    let n = 0
    for (const p of this.persistent.values()) n += p.bytes
    return n
  }
  persistentOf(key: string): number { return this.persistent.get(key)?.bytes ?? 0 }
  /** The live pool entry charging a persistent key (for an owner that takes it over at close). */
  persistentEntry(key: string): PoolEntry | undefined { return this.persistent.get(key)?.entry }

  private usable(): void {
    // A closed pool means the document was closed: the import is being torn down.
    if (this.pool.closed) throw new PackageError('cancelled', 'the document memory pool is closed')
  }
  private entryKey(key: string): string { return `construction:${this.docId}:${key}` }

  /**
   * `untilReleased`: for a holder that outlives close/abandon (the detached
   * leftover sweep keeps its grouping reachable until it returns). releaseAll
   * then refuses its later needs, which stops the holder at its next check,
   * but keeps its charge until the holder releases it.
   */
  share(what: string, options: { untilReleased?: boolean } = {}): ConstructionShare {
    const share: Share = { what, charge: this.pool.charge(), live: true, untilReleased: options.untilReleased ?? false }
    this.shares.add(share)
    const owner = this
    return {
      get bytes() { return share.charge.bytes },
      need(bytes: number) {
        if (!share.live) throw new PackageError('cancelled', `${what} share already released`)
        owner.usable()
        if (bytes <= share.charge.bytes && bytes > share.charge.bytes - 2 * STEP_BYTES) return
        const stepped = Math.ceil(bytes / STEP_BYTES) * STEP_BYTES
        if (share.charge.set(stepped) || share.charge.set(bytes)) return
        throw memoryRefusal(what, bytes, owner.pool.limitBytes)
      },
      release() {
        if (!owner.shares.has(share)) return
        share.live = false
        share.charge.release()
        owner.shares.delete(share)
      },
    }
  }

  /** Charges a persistent, non-evictable entry at an absolute size (replacing its previous size). */
  persist(key: string, bytes: number, what: string): void {
    this.usable()
    const entry = this.pool.retain(this.entryKey(key), { tier: 'source', bytes, evictable: false })
    if (!entry) {
      if (this.pool.closed) throw new PackageError('cancelled', 'the document memory pool is closed')
      throw memoryRefusal(what, bytes, this.pool.limitBytes)
    }
    // retain() replaced (released) the previous entry of this key in the same pool.
    this.persistent.set(key, { what, bytes, entry })
  }

  /**
   * Moves every live charge into `pool`, atomically. When the pool cannot hold
   * all of them, the partial transfer is undone, the old ownership stays
   * charged where it was and the typed memory-budget refusal is thrown.
   */
  rebind(pool: MemoryPool): void {
    if (pool === this.pool) return
    const charges: Array<[Share, PoolCharge]> = []
    const entries: Array<[string, PoolEntry]> = []
    const undo = () => {
      for (const [, c] of charges) c.release()
      for (const [, e] of entries) e.release()
    }
    for (const share of this.shares) {
      const next = pool.charge()
      charges.push([share, next])
      if (!next.set(share.charge.bytes)) { undo(); throw memoryRefusal(share.what, share.charge.bytes, pool.limitBytes) }
    }
    for (const [key, p] of this.persistent) {
      const next = pool.retain(this.entryKey(key), { tier: 'source', bytes: p.bytes, evictable: false })
      if (!next) { undo(); throw memoryRefusal(p.what, p.bytes, pool.limitBytes) }
      entries.push([key, next])
    }
    for (const [share, next] of charges) { share.charge.release(); share.charge = next }
    for (const [key, next] of entries) {
      const p = this.persistent.get(key)!
      p.entry.release()
      p.entry = next
    }
    this.pool = pool
    // The provisional pool is private to the store; nothing else holds it.
    this.provisional?.close()
    this.provisional = undefined
  }

  /** Close/abandon: every transient share (except an untilReleased one's charge) and persistent entry is released. */
  releaseAll(): void {
    for (const share of [...this.shares]) {
      share.live = false
      if (share.untilReleased) continue
      share.charge.release()
      this.shares.delete(share)
    }
    for (const p of this.persistent.values()) p.entry.release()
    this.persistent.clear()
  }
}
