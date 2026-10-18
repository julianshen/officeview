/**
 * One ledger for the aggregate import budgets (spec §12.4, design note §4.3).
 *
 * It is also the ExpandedBudget that openPart charges before it hands out a
 * step's bytes, so the expanded-queue bound and the per-load expanded total
 * are enforced while decoding, not after the fact.
 */
import { PackageError } from '../../core/package-error'
import type { ExpandedBudget } from '../../core/zip-stream'
import type { ProfileId } from '../view/contracts'

export const POOLS = ['compressedSpool', 'expandedQueue', 'pages', 'sharedStrings', 'blobRecords', 'decodedBuffers', 'objectModel'] as const
export type Pool = typeof POOLS[number]
export type QuotaLimits = Readonly<Record<Pool, number>>

const MiB = 1024 * 1024
/** Proposed, to be frozen in S10. compressedSpool/pages are storage-backed.
 * pages covers the paged cell cap (2^25 D / 2^24 M) at the measured 26.9 B per
 * stored cell (S1.2: F2 full, 8,238,808 cells, 211.0 MiB of pages); the
 * earlier 128/64 MiB refused F2 with a quota error. */
const QUOTA_LIMITS: Record<ProfileId, QuotaLimits> = {
  D: { compressedSpool: 256 * MiB, expandedQueue: 8 * MiB, pages: 1024 * MiB, sharedStrings: 64 * MiB, blobRecords: 32 * MiB, decodedBuffers: 128 * MiB, objectModel: 256 * MiB },
  M: { compressedSpool: 128 * MiB, expandedQueue: 8 * MiB, pages: 512 * MiB, sharedStrings: 32 * MiB, blobRecords: 16 * MiB, decodedBuffers: 64 * MiB, objectModel: 128 * MiB },
}

export function quotaLimitsFor(profile: ProfileId): QuotaLimits {
  return QUOTA_LIMITS[profile]
}

// blobRecords are oversized cell values written to the page backend (spec
// §4.3.4), so exhausting them is a storage quota refusal like pages.
const STORAGE_POOLS: ReadonlySet<Pool> = new Set<Pool>(['compressedSpool', 'pages', 'blobRecords'])

export interface Reservation {
  readonly pool: Pool
  readonly bytes: number
  /** Idempotent. */
  release(): void
}

export interface QuotaSnapshot {
  current: Record<Pool, number>
  peak: Record<Pool, number>
  totalExpanded: number
  released: boolean
}

export interface QuotaLedger extends ExpandedBudget {
  /** Throws PackageError 'quota' (storage pools) or 'resource-limit' (memory pools). */
  reserve(pool: Pool, bytes: number): Reservation
  /** An ExpandedBudget for one part: charges this ledger and a per-part cap. */
  partBudget(partId: string, maxPartBytes: number): ExpandedBudget
  /** The consumer finished with the last yielded step (end of part, failure). */
  drainQueue(): void
  snapshot(): QuotaSnapshot
  /** Cancellation/close: every reservation is dropped and later charges fail. */
  releaseAll(): void
}

function zeroes(): Record<Pool, number> {
  const out = {} as Record<Pool, number>
  for (const pool of POOLS) out[pool] = 0
  return out
}

function checkBytes(bytes: number): void {
  if (!Number.isSafeInteger(bytes) || bytes < 0) throw new RangeError(`byte count must be a non-negative safe integer, got ${bytes}`)
}

export function createQuotaLedger(limits: QuotaLimits, options: { maxTotalExpandedBytes: number }): QuotaLedger {
  const current = zeroes()
  const peak = zeroes()
  const live = new Set<{ pool: Pool; bytes: number }>()
  let totalExpanded = 0
  let released = false

  const add = (pool: Pool, bytes: number, partId?: string): void => {
    if (released) throw new PackageError('cancelled', 'import ledger already released', partId)
    if (current[pool] + bytes > limits[pool]) {
      const code = STORAGE_POOLS.has(pool) ? 'quota' : 'resource-limit'
      throw new PackageError(code, `${pool} budget of ${limits[pool]} bytes exceeded (${current[pool]} held, ${bytes} requested)`, partId)
    }
    current[pool] += bytes
    if (current[pool] > peak[pool]) peak[pool] = current[pool]
  }

  // openPart yields views of ONE reused output buffer, so a step's bytes are
  // queued only until the next pull: the queue holds the latest step.
  let queued = 0
  const charge = (bytes: number, partId?: string): void => {
    checkBytes(bytes)
    if (released) throw new PackageError('cancelled', 'import ledger already released', partId)
    if (totalExpanded + bytes > options.maxTotalExpandedBytes) {
      throw new PackageError('resource-limit', `expanded total of ${options.maxTotalExpandedBytes} bytes exceeded`, partId)
    }
    current.expandedQueue -= queued
    queued = 0
    add('expandedQueue', bytes, partId)
    queued = bytes
    totalExpanded += bytes
  }

  return {
    charge: (bytes) => charge(bytes),
    partBudget(partId, maxPartBytes) {
      let part = 0
      return {
        charge(bytes) {
          checkBytes(bytes)
          if (part + bytes > maxPartBytes) throw new PackageError('resource-limit', `part budget of ${maxPartBytes} bytes exceeded`, partId)
          charge(bytes, partId)
          part += bytes
        },
      }
    },
    drainQueue() {
      current.expandedQueue -= queued
      queued = 0
    },
    reserve(pool, bytes) {
      checkBytes(bytes)
      add(pool, bytes)
      const entry = { pool, bytes }
      live.add(entry)
      return {
        pool,
        bytes,
        release() {
          if (!live.delete(entry)) return
          current[pool] -= bytes
        },
      }
    },
    snapshot() {
      return { current: { ...current }, peak: { ...peak }, totalExpanded, released }
    },
    releaseAll() {
      live.clear()
      queued = 0
      for (const pool of POOLS) current[pool] = 0
      released = true
    },
  }
}
