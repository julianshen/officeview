/**
 * Show clocks (slideshow spec 6.2): the scheduler reads time only through a
 * `ShowClock`, so tests drive it and pausing stops the timeline.
 */
import type { ShowClock } from './types'

/** A clock that moves only when told (tests, and every injected timer in CI). */
export class ManualClock implements ShowClock {
  private t = 0
  private stopped = false
  get paused(): boolean { return this.stopped }
  now(): number { return this.t }
  /** Moves time forward by `ms` unless paused. */
  advance(ms: number): void {
    if (!Number.isFinite(ms) || ms < 0) throw new RangeError(`ManualClock.advance: ${ms}`)
    if (!this.stopped) this.t += ms
  }
  pause(): void { this.stopped = true }
  resume(): void { this.stopped = false }
}

/**
 * The production clock: show time advances by the `requestAnimationFrame`
 * timestamps of the window it is given (realm-safe; never the global
 * window), starting at 0 on the first frame. Paused time adds nothing, even
 * when no frames arrive during the pause: `resume` rebases on the window's
 * `performance.now()`.
 * Pausing on a hidden document is wired by the session (S1-10/S1-19).
 */
export class RafClock implements ShowClock {
  private t = 0
  private last: number | undefined
  private stopped = false
  private request: number | undefined
  constructor(private readonly win: Window) { this.schedule() }
  get paused(): boolean { return this.stopped }
  now(): number { return this.t }
  pause(): void { this.stopped = true }
  /**
   * Rebases on the window's `performance.now()` (the rAF timeline), so a pause
   * with no frames at all (a hidden document) never counts as show time.
   */
  resume(): void {
    if (!this.stopped) return
    this.stopped = false
    this.last = this.win.performance.now()
  }
  dispose(): void {
    if (this.request !== undefined) this.win.cancelAnimationFrame(this.request)
    this.request = undefined
  }
  private schedule(): void {
    this.request = this.win.requestAnimationFrame((stamp) => {
      this.request = undefined
      if (this.last !== undefined && !this.stopped && stamp > this.last) this.t += stamp - this.last
      // A frame stamped before the resume point never moves the anchor back.
      this.last = this.last === undefined ? stamp : Math.max(this.last, stamp)
      this.schedule()
    })
  }
}
