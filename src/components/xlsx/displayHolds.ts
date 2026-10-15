import type { LedgerHold } from '../../xlsx/view/mainLedger'

/**
 * Spec §§10.1, 12.4: one viewer's holds on the charged buffers its sheet
 * entries display (view order and mask, row projection). A replaced entry
 * keeps its holds until React commits a display that no longer shows it: an
 * interrupted or discarded render never refunds. `extra` holds an entry for a
 * continuation (a copy session) until that continuation ends. Releases are
 * idempotent; an unmount releases everything and a remount holds the current
 * entries again.
 */
export class DisplayHolds<E extends object> {
  private readonly current = new Map<E, LedgerHold[] | null>()
  private retired: { at: number; holds: LedgerHold[] | null }[] = []
  constructor(private readonly acquire: (entry: E) => LedgerHold[]) {}
  /** A new entry is current: hold its buffers (during render is safe, it only over-counts). */
  own(entry: E): void { if (!this.current.has(entry)) this.current.set(entry, this.acquire(entry)) }
  /** `entry` was replaced; its holds end with the commit of display `at` (or a later one). */
  retire(entry: E, at: number): void {
    if (!this.current.has(entry)) return
    this.retired.push({ at, holds: this.current.get(entry)! })
    this.current.delete(entry)
  }
  /** Display `serial` is committed: entries replaced at or before it are no longer shown. */
  committed(serial: number): void {
    const keep: typeof this.retired = []
    for (const item of this.retired) if (item.at <= serial) release(item.holds); else keep.push(item)
    this.retired = keep
  }
  /** A continuation's own hold on an entry; call the result when it ends. */
  extra(entry: E): () => void {
    const holds = this.acquire(entry)
    return () => release(holds.splice(0))
  }
  dispose(): void {
    for (const [entry, holds] of this.current) { release(holds); this.current.set(entry, null) }
    for (const item of this.retired.splice(0)) release(item.holds)
  }
  /** A remount (StrictMode) shows the same entries again. */
  restore(): void { for (const [entry, holds] of this.current) if (!holds) this.current.set(entry, this.acquire(entry)) }
}
function release(holds: LedgerHold[] | null): void { for (const hold of holds ?? []) hold.release() }
