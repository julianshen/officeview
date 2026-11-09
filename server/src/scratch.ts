/** §14 scratch lease: at most `limit` per-session final-raster encodes in
 * flight; FIFO with at most `maxWaiters` queued; closing rejects every waiter
 * and refuses later acquires. A release hands its slot straight to the head
 * waiter, so a racing acquire can never push the count past `limit`. */
export class LeaseClosed extends Error { constructor() { super('scratch lease closed'); this.name = 'LeaseClosed' } }
export class LeaseFull extends Error { constructor() { super('scratch lease queue full'); this.name = 'LeaseFull' } }

export class ScratchLease {
  private active = 0
  private closed = false
  private waiters: Array<{ ok: () => void; fail: (e: Error) => void }> = []
  constructor(private readonly limit = 2, private readonly maxWaiters = Infinity, private readonly onChange: (delta: number) => void = () => {}) {}
  get inFlight() { return this.active }

  /** Resolves with an idempotent release once a slot is held. */
  async acquire(): Promise<() => void> {
    if (this.closed) throw new LeaseClosed()
    if (this.active < this.limit) { this.active++; this.onChange(1) }
    else {
      if (this.waiters.length >= this.maxWaiters) throw new LeaseFull()
      // The slot is transferred by release(); `active` already counts it.
      await new Promise<void>((ok, fail) => this.waiters.push({ ok, fail }))
      if (this.closed) { this.active--; this.onChange(-1); throw new LeaseClosed() }
    }
    let released = false
    return () => {
      if (released) return
      released = true
      const next = this.closed ? undefined : this.waiters.shift()
      if (next) next.ok()                        // hand-off: active unchanged
      else { this.active--; this.onChange(-1) }
    }
  }
  close(): void {
    if (this.closed) return
    this.closed = true
    for (const w of this.waiters.splice(0)) w.fail(new LeaseClosed())
  }
}
