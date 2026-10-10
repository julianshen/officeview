/**
 * Application allocation ledger for the real-browser harness (S0.3).
 *
 * Tracks charged/retained/peak bytes per pool (main, worker, bitmap) with
 * owner-keyed release. Charging the same owner twice accumulates; release
 * frees everything that owner charged. Used by the dev page to account for
 * canvas backing stores (bitmap), source/grid copies (main) and later worker
 * buffers (worker).
 */
import type { AllocationPool, PoolAccounting } from './types'

const POOLS: readonly AllocationPool[] = ['main', 'worker', 'bitmap']

interface OwnerCharge {
  pool: AllocationPool
  bytes: number
}

export class AllocationLedger {
  private readonly retainedBytes = new Map<AllocationPool, number>()
  private readonly peakBytes = new Map<AllocationPool, number>()
  private readonly chargedBytes = new Map<AllocationPool, number>()
  private readonly owners = new Map<string, OwnerCharge[]>()

  charge(pool: AllocationPool, bytes: number, owner: string): void {
    if (!Number.isFinite(bytes) || bytes < 0) {
      throw new Error(`ledger charge for ${owner} must be a finite non-negative byte count, got ${bytes}`)
    }
    this.retainedBytes.set(pool, (this.retainedBytes.get(pool) ?? 0) + bytes)
    this.chargedBytes.set(pool, (this.chargedBytes.get(pool) ?? 0) + bytes)
    const retainedNow = this.retainedBytes.get(pool) ?? 0
    if (retainedNow > (this.peakBytes.get(pool) ?? 0)) this.peakBytes.set(pool, retainedNow)
    const entries = this.owners.get(owner) ?? []
    entries.push({ pool, bytes })
    this.owners.set(owner, entries)
  }

  /** Free every charge owned by `owner`; false when the owner is unknown. */
  release(owner: string): boolean {
    const entries = this.owners.get(owner)
    if (!entries) return false
    for (const { pool, bytes } of entries) {
      this.retainedBytes.set(pool, Math.max(0, (this.retainedBytes.get(pool) ?? 0) - bytes))
    }
    this.owners.delete(owner)
    return true
  }

  releaseAll(): void {
    for (const pool of POOLS) this.retainedBytes.set(pool, 0)
    this.owners.clear()
  }

  retained(pool: AllocationPool): number {
    return this.retainedBytes.get(pool) ?? 0
  }

  peak(pool: AllocationPool): number {
    return this.peakBytes.get(pool) ?? 0
  }

  accounting(): Record<AllocationPool, PoolAccounting> {
    const snapshot = {} as Record<AllocationPool, PoolAccounting>
    for (const pool of POOLS) {
      snapshot[pool] = {
        retained: this.retainedBytes.get(pool) ?? 0,
        peak: this.peakBytes.get(pool) ?? 0,
        charged: this.chargedBytes.get(pool) ?? 0,
      }
    }
    return snapshot
  }
}
