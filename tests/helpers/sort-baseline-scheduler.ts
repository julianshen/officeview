// Frozen fb39d72 scheduler for the test-only sort oracle.
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

export type YieldFunction<Phase extends string = string> = (phase: Phase) => Promise<void>
export interface CooperativeOptions<Phase extends string = string> {
  now?: () => number
  yield?: YieldFunction<Phase>
}
export interface HostYield { yield: () => Promise<void>; dispose: () => void }
type Allocate = <T>(create: () => T) => T

/** Real task boundary for the host. Node's immediate queue gives timers/I/O a
 * turn between yields (a chain of port messages can drain within one poll
 * phase); browsers prefer scheduler.yield, then task-backed MessageChannel
 * ports, then a timer. Each operation owns its ports and disposes them. */
export function createHostYield(allocate: Allocate = create => create()): HostYield {
  const host = globalThis as typeof globalThis & {
    scheduler?: { yield: () => Promise<void> }
    setImmediate?: (callback: () => void) => unknown
  }
  if (host.setImmediate) return {
    yield: () => new Promise<void>(resolve => host.setImmediate!(resolve)),
    dispose: () => {},
  }
  if (host.scheduler?.yield) return { yield: () => host.scheduler!.yield(), dispose: () => {} }
  if (typeof MessageChannel !== 'undefined') {
    const channel = allocate(() => new MessageChannel())
    let resume: (() => void) | undefined
    channel.port1.onmessage = () => {
      const resolve = resume
      resume = undefined
      resolve?.()
    }
    return {
      yield: () => new Promise<void>(resolve => {
        resume = resolve
        channel.port2.postMessage(null)
      }),
      dispose: () => {
        channel.port1.onmessage = null
        channel.port1.close()
        channel.port2.close()
      },
    }
  }
  return {
    yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)),
    dispose: () => {},
  }
}

/**
 * One operation's yield budget. `due(units)` charges completed work and samples
 * the clock once; a caller that sees `true` awaits `pause(phase)`. `check()` is
 * an override hook run around every pause (the sort engine throws its typed
 * cancellation/stale outcomes from it).
 */
export class CooperativeSlice<Phase extends string = string> {
  private work = 0
  private last: number
  private taken = 0
  readonly now: () => number
  readonly schedule: YieldFunction<Phase>
  readonly dispose: () => void
  constructor(options: CooperativeOptions<Phase>, allocate?: Allocate) {
    this.now = options.now ?? (() => performance.now())
    this.last = this.now()
    const host = options.yield ? undefined : createHostYield(allocate)
    this.schedule = options.yield ?? host!.yield
    this.dispose = host?.dispose ?? (() => {})
  }
  /** Yields taken so far (diagnostic; tests bound it). */
  get yields(): number { return this.taken }
  check(): void {}
  due(units = 1): boolean {
    this.check()
    this.work += units
    return this.work >= YIELD_WORK_UNITS || this.now() - this.last >= YIELD_INTERVAL_MS
  }
  async pause(phase: Phase): Promise<void> {
    this.check()
    await this.schedule(phase)
    this.check()
    this.taken++
    this.work = 0
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
