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
 * pool; a pool that cannot hold them is not an error at bind time (the session
 * must still open) but makes the next construction charge refuse, ending the
 * import with the typed failure.
 *
 * Byte figures are estimates of JS heap held, documented here so the model can
 * be checked against a profile later (S10):
 * - builder: 64 B per record (six number-array slots and a string reference),
 *   2 B per UTF-16 code unit of text and spilled values, 2 B per formula-meta
 *   JSON character plus 96 B per meta object (the decoded-page allowance);
 * - a built page waiting for commit: payloadBytes of its typed lanes plus its
 *   spilled values;
 * - encode scratch: the exact encoded buffer plus the meta JSON string;
 * - a pending shared string: 2 B per code unit plus 32 B of string header.
 */
import { PackageError } from '../../core/package-error'
import type { MemoryPool, PoolCharge, PoolEntry } from '../view/tileCache'

export const BUILDER_RECORD_BYTES = 64
export const META_OBJECT_BYTES = 96
export const PENDING_STRING_BYTES = 32
/** One PageSummary: the object, its ids and its RectAggregate. */
export const PAGE_SUMMARY_BYTES = 192
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

interface Share { what: string; charge: PoolCharge; live: boolean }

export class ConstructionMemory {
  private pool: MemoryPool
  /** The store's own pool until a document pool is bound; closed on the first rebind. */
  private provisional: MemoryPool | undefined
  private readonly shares = new Set<Share>()
  private readonly persistent = new Map<string, { what: string; bytes: number; entry: PoolEntry | null }>()
  /** Set when a rebind could not move every charge: the next charge refuses. */
  private refused: { what: string; bytes: number; limitBytes: number } | undefined

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

  private usable(): void {
    if (this.refused) throw memoryRefusal(this.refused.what, this.refused.bytes, this.refused.limitBytes)
    // A closed pool means the document was closed: the import is being torn down.
    if (this.pool.closed) throw new PackageError('cancelled', 'the document memory pool is closed')
  }

  share(what: string): ConstructionShare {
    const share: Share = { what, charge: this.pool.charge(), live: true }
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
        if (!share.live) return
        share.live = false
        share.charge.release()
        owner.shares.delete(share)
      },
    }
  }

  /** Charges a persistent, non-evictable entry at an absolute size (replacing its previous size). */
  persist(key: string, bytes: number, what: string): void {
    this.usable()
    const previous = this.persistent.get(key)
    const entry = this.pool.retain(`construction:${this.docId}:${key}`, { tier: 'source', bytes, evictable: false })
    if (!entry) {
      if (this.pool.closed) throw new PackageError('cancelled', 'the document memory pool is closed')
      throw memoryRefusal(what, bytes, this.pool.limitBytes)
    }
    // retain() replaced the previous entry of this key in the same pool.
    if (previous?.entry && previous.entry !== entry && previous.entry.live) previous.entry.release()
    this.persistent.set(key, { what, bytes, entry })
  }

  /** Moves every live charge into `pool`. Never throws; a shortfall refuses the next charge. */
  rebind(pool: MemoryPool): void {
    if (pool === this.pool) return
    this.pool = pool
    for (const share of this.shares) {
      const next = pool.charge()
      if (!next.set(share.charge.bytes)) this.refused ??= { what: share.what, bytes: share.charge.bytes, limitBytes: pool.limitBytes }
      share.charge.release()
      share.charge = next
    }
    for (const [key, p] of this.persistent) {
      p.entry?.release()
      p.entry = pool.retain(`construction:${this.docId}:${key}`, { tier: 'source', bytes: p.bytes, evictable: false })
      if (!p.entry) this.refused ??= { what: p.what, bytes: p.bytes, limitBytes: pool.limitBytes }
    }
    // The provisional pool is private to the store; nothing else holds it.
    this.provisional?.close()
    this.provisional = undefined
  }

  /** Close/abandon: every transient share and persistent entry is released. */
  releaseAll(): void {
    for (const share of this.shares) { share.live = false; share.charge.release() }
    this.shares.clear()
    for (const p of this.persistent.values()) p.entry?.release()
    this.persistent.clear()
  }
}
