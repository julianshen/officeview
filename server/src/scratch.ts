/** §14 scratch lease: at most `limit` per-session final-raster encodes in
 * flight; FIFO; closing rejects every waiter and refuses later acquires. */
export class LeaseClosed extends Error { constructor() { super('scratch lease closed'); this.name = 'LeaseClosed' } }

export class ScratchLease {
  private active = 0
  private closed = false
  private waiters: Array<{ ok: () => void; fail: (e: Error) => void }> = []
  constructor(private readonly limit = 2, private readonly onChange: (delta: number) => void = () => {}) {}
  get inFlight() { return this.active }

  /** Resolves with an idempotent release once a slot is free. */
  async acquire(): Promise<() => void> {
    if (this.closed) throw new LeaseClosed()
    if (this.active >= this.limit) await new Promise<void>((ok, fail) => this.waiters.push({ ok, fail }))
    this.active++; this.onChange(1)
    let released = false
    return () => {
      if (released) return
      released = true
      this.active--; this.onChange(-1)
      if (!this.closed) this.waiters.shift()?.ok()
    }
  }
  close(): void {
    if (this.closed) return
    this.closed = true
    for (const w of this.waiters.splice(0)) w.fail(new LeaseClosed())
  }
}
