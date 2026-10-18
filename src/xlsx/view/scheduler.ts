/**
 * Cooperative scheduling shared by the resident worker kernels (spec §10.2,
 * plan S4.2): preparation, populated/projection scans and the S7.2 sort engine
 * all yield on the same budget, at least every 10 ms or 4,096 work units,
 * whichever arrives first.
 *
 * Deterministic by injection: `now` and `yield` replace the host clock and the
 * host task queue, so tests drive every interleaving explicitly and assert yield
 * counts, never elapsed time.
 */

/** Work units (rows, entries or comparisons) between yields at most. */
export const YIELD_WORK_UNITS = 4096
/** Wall-clock budget between yields, in milliseconds of the injected clock. */
export const YIELD_INTERVAL_MS = 10
/** Work units between cancellation/staleness checkpoints at most (spec §12.1
 * amortization clause); each checkpoint also samples the clock. */
export const CHECK_WORK_UNITS = 256

export type YieldFunction<Phase extends string = string> = (phase: Phase) => Promise<void>
export interface CooperativeOptions<Phase extends string = string> {
  now?: () => number
  yield?: YieldFunction<Phase>
  /** Opt-in for unit-at-a-time kernels (§12.1). Other callers keep per-call checks. */
  checkEvery?: 1 | 256
  /** Owner boundary may interrupt a stalled host/injected yield. */
  awaitYield?: (work: Promise<void>) => Promise<void>
  /** Clock reading the first slice starts at (default: now). */
  start?: number
}
export interface HostYield { yield: () => Promise<void>; dispose: () => void }
type Allocate = <T>(create: () => T) => T

type YieldHost = {
  setImmediate?: (callback: () => void) => unknown
  MessageChannel?: typeof MessageChannel
  process?: { versions?: { node?: string } }
  setTimeout: (callback: () => void, ms: number) => unknown
}
/**
 * One real macrotask turn for pacing (spec §10.2), for `host`:
 *  - Node (and jsdom on Node): its native setImmediate. Node's immediate queue
 *    gives timers/I/O a turn; a chain of port messages can drain within one
 *    poll phase there.
 *  - elsewhere: one shared MessageChannel. A posted message is an ordinary
 *    task, FIFO with already queued tasks (a cancel message is serviced), and
 *    never clamped like a nested 0 ms timer (>= 4 ms in browsers).
 *    A bundled setImmediate polyfill is not trusted, and scheduler.yield is not
 *    used: its continuation outranks queued tasks.
 *  - a 0 ms timer only where neither exists.
 */
export function macrotaskYield(host: YieldHost = globalThis as unknown as YieldHost): () => Promise<void> {
  if (host.process?.versions?.node && host.setImmediate) {
    const immediate = host.setImmediate
    return () => new Promise<void>(resolve => { immediate(resolve) })
  }
  if (host.MessageChannel) {
    let channel: MessageChannel | undefined
    const waiting: (() => void)[] = []
    return () => new Promise<void>(resolve => {
      if (!channel) {
        channel = new host.MessageChannel!()
        // Each message resumes the oldest waiter: one per posted yield.
        channel.port1.onmessage = () => { waiting.shift()?.() }
      }
      waiting.push(resolve)
      channel.port2.postMessage(null)
    })
  }
  return () => new Promise<void>(resolve => { host.setTimeout(resolve, 0) })
}
let shared: { yield: () => Promise<void>; immediate: unknown; channel: unknown } | undefined
/** The host's shared macrotask yield (see macrotaskYield), re-resolved only
 * if the host's yield primitives themselves change (test stubs). */
export function hostTask(): Promise<void> {
  const host = globalThis as unknown as YieldHost
  if (!shared || shared.immediate !== host.setImmediate || shared.channel !== host.MessageChannel) {
    shared = { yield: macrotaskYield(host), immediate: host.setImmediate, channel: host.MessageChannel }
  }
  return shared.yield()
}

/** Real task boundary for the host: the shared macrotask yield (hostTask).
 * Nothing per operation to allocate or dispose (`allocate` is kept for
 * callers that account allocations; the shared channel is created once). */
export function createHostYield(_allocate: Allocate = create => create()): HostYield {
  return { yield: hostTask, dispose: () => {} }
}

/**
 * One operation's yield budget. `due(units)` charges completed work and samples
 * the clock at the selected check interval; a caller that sees `true` awaits `pause(phase)`. `check()` is
 * an override hook run around every pause (the sort engine throws its typed
 * cancellation/stale outcomes from it).
 */
export class CooperativeSlice<Phase extends string = string> {
  private work = 0
  private last: number
  private taken = 0
  private unchecked = 0
  private readonly checkEvery: number
  readonly now: () => number
  readonly schedule: YieldFunction<Phase>
  readonly dispose: () => void
  constructor(options: CooperativeOptions<Phase>, allocate?: Allocate) {
    this.checkEvery = options.checkEvery === 256 ? 256 : 1
    this.now = options.now ?? (() => performance.now())
    this.last = options.start ?? this.now()
    const host = options.yield ? undefined : createHostYield(allocate)
    const schedule = options.yield ?? host!.yield
    this.schedule = options.awaitYield ? phase => options.awaitYield!(schedule(phase)) : schedule
    this.dispose = host?.dispose ?? (() => {})
  }
  /** Yields taken so far (diagnostic; tests bound it). */
  get yields(): number { return this.taken }
  check(): void {}
  due(units = 1): boolean {
    this.work += units
    this.unchecked += units
    if (this.unchecked < this.checkEvery && this.work < YIELD_WORK_UNITS) return false
    this.unchecked = 0
    this.check()
    if (this.checkEvery === 1 && this.work >= YIELD_WORK_UNITS) return true
    const now = this.now()
    return this.work >= YIELD_WORK_UNITS || now - this.last >= YIELD_INTERVAL_MS
  }
  async pause(phase: Phase): Promise<void> {
    this.check()
    await this.schedule(phase)
    this.check()
    this.taken++
    this.work = 0
    this.unchecked = 0
    this.last = this.now()
  }
  /** Yields if any work is unyielded: used at a handoff to another kernel with
   * its own budget, so the combined stretch never exceeds one budget. */
  async flush(phase: Phase): Promise<void> {
    if (this.work > 0) await this.pause(phase)
  }
  /** Charges `units` and yields when the budget is spent. */
  async step(phase: Phase, units = 1): Promise<void> {
    if (this.due(units)) await this.pause(phase)
  }
}

/** Yield budget a batched pass charges work to; CooperativeSlice satisfies it. */
export interface BatchSlice {
  due(units: number): boolean
  pause(phase: string): Promise<void>
  flush(phase: string): Promise<void>
}
/**
 * Work batching shared by column passes, the owned value adapter's decode and
 * the shipped SourceStores' request decode (spec §§10.2, 12.1): a batch ends at
 * the next checkpoint, at most CHECK_WORK_UNITS units apart, where `check`
 * (cancellation and staleness; it throws) runs and the slice samples its
 * clock; the slice pauses with a real yield once YIELD_WORK_UNITS units or its
 * time budget are spent. Without a caller slice, a host-yield slice is created
 * at the first checkpoint and disposed with the budget.
 */
export class CheckpointBudget {
  private work = 0
  private unchecked = 0
  private owned: CooperativeSlice<string> | undefined
  /** The operation's start: an owned slice created later keeps this time origin. */
  private readonly start: number | undefined
  constructor(
    private readonly check: () => void,
    private slice?: BatchSlice,
    private readonly pacing: Pick<CooperativeOptions, 'now' | 'yield'> = {},
    private readonly phase = 'filter',
  ) {
    if (!slice) this.start = (pacing.now ?? (() => performance.now()))()
  }
  /** Units the next batch may take. */
  get batch(): number { return Math.max(1, Math.min(CHECK_WORK_UNITS - this.unchecked, YIELD_WORK_UNITS - this.work)) }
  /** Charges units; true when a pause is due. */
  spend(units: number): boolean {
    this.work += units
    this.unchecked += units
    if (this.unchecked < CHECK_WORK_UNITS && this.work < YIELD_WORK_UNITS) return false
    const since = this.unchecked
    this.unchecked = 0
    this.check()
    // The slice counts the same units and owns the clock and the host yield.
    this.slice ??= this.owned = new CooperativeSlice<string>({ ...this.pacing, start: this.start })
    return this.slice.due(since)
  }
  /** A check now (a final check without a pause). Units since the last
   * checkpoint stay counted toward the next one and the slice. */
  checkpoint(): void {
    this.check()
  }
  /** Ends a pass with a final check; units below the next checkpoint are
   * charged to a caller slice, so the next pass sharing it starts after them. */
  async finish(): Promise<void> {
    const since = this.unchecked
    this.unchecked = 0
    this.check()
    if (since && this.slice && this.slice !== this.owned && this.slice.due(since)) await this.pause()
  }
  async pause(): Promise<void> {
    await this.slice!.pause(this.phase)
    this.work = 0
    this.unchecked = 0
    this.check()
  }
  /** Yields work an earlier pass left unyielded on a caller slice. */
  async flush(): Promise<void> {
    if (this.slice) await this.slice.flush(this.phase)
    this.check()
  }
  dispose(): void { this.owned?.dispose() }
}
