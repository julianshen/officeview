/**
 * The Presenter View elapsed timer and wall clock (slideshow spec 5.2
 * "Elapsed timer, with pause and reset, plus the wall-clock time").
 *
 * The timer is pure arithmetic over readings of an injected monotonic clock
 * (gc.md: every clock is injected and driven). It is not the show clock:
 * pausing the show (S or `+`) does not stop it, and its own pause does not
 * stop the show, as in PowerPoint.
 */

export interface ElapsedTimer {
  /** Milliseconds counted at reading `t`. */
  elapsed(t: number): number
  /** Stops counting at `t` (idempotent). */
  pause(t: number): void
  /** Additive to the brief: counting resumes from `t` (idempotent). */
  resume(t: number): void
  /** Back to zero at `t`; a paused timer stays paused at zero. */
  reset(t: number): void
  readonly paused: boolean
}

export function createElapsedTimer(t0: number): ElapsedTimer {
  /** Counted before the current run. */
  let banked = 0
  /** The current run's start, or undefined while paused. */
  let since: number | undefined = t0
  return {
    elapsed: (t) => Math.max(0, banked + (since === undefined ? 0 : t - since)),
    pause(t) {
      if (since === undefined) return
      banked = Math.max(0, banked + t - since)
      since = undefined
    },
    resume(t) {
      if (since === undefined) since = t
    },
    reset(t) {
      banked = 0
      if (since !== undefined) since = t
    },
    get paused() { return since === undefined },
  }
}

const two = (n: number) => String(n).padStart(2, '0')

/** `H:MM:SS`, as PowerPoint's presenter timer shows it. */
export function formatElapsed(ms: number): string {
  const s = Math.floor(Math.max(0, Number.isFinite(ms) ? ms : 0) / 1000)
  return `${Math.floor(s / 3600)}:${two(Math.floor(s / 60) % 60)}:${two(s % 60)}`
}

/** The wall-clock time in the user's locale, hours and minutes. */
export function formatWallClock(ms: number): string {
  return new Date(ms).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
}
