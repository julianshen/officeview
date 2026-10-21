/**
 * Auto-advance and the kiosk restart timer (slideshow spec 4.4; the
 * auto-advance rule is the spec's working assumption, pending D4 captures).
 *
 * Every time here is show-clock time: pausing the show clock (S or `+`, or a
 * hidden document) stops both timers with no extra wiring.
 */
import type { DiagnosticSink, ShowSettings, ShowSlide } from '../types'
import type { SlideCursor, SlideTimeline } from './steps'

export interface AutoAdvance {
  /**
   * `advClick` on the arriving slide's transition: `false` means mouse clicks
   * do not advance. Keyboard, API and presenter next are never gated, so a
   * deck cannot get stuck.
   */
  readonly clickAdvances: boolean
  /**
   * The slide should advance now: `useTimings` is on, the slide has `advTm`,
   * `advTm` ms have passed since its transition ended, no click step is left
   * unplayed, and no finite main-sequence effect is running. Indefinite
   * effects, interactive sequences and media never hold it.
   */
  due(cursor: SlideCursor, transitionEndT: number, t: number): boolean
}

export function createAutoAdvance(slide: ShowSlide, settings: ShowSettings, timeline: SlideTimeline): AutoAdvance {
  const advTm = slide.transition?.advTmMs
  const timed = settings.useTimings && typeof advTm === 'number' && Number.isFinite(advTm) && advTm >= 0
  return {
    clickAdvances: slide.transition?.advClick ?? true,
    due(cursor, transitionEndT, t) {
      if (!timed || !(t >= transitionEndT + advTm)) return false
      if (cursor.mode === 'static') return true
      if (cursor.step < timeline.steps) return false
      return !(timeline.busyUntil(cursor) > t)
    },
  }
}

export interface KioskTimer {
  restartDue(t: number): boolean
  /** Call on every restart, including automatic loop restarts: the counter starts again. */
  restarted(t: number): void
}

/**
 * Kiosk: the show restarts from its first slide `kioskRestartMs` after the
 * most recent start or restart. Not a kiosk: never due. A restart of 0 (or an
 * invalid one) would restart on every frame, so it is refused (never due)
 * and diagnosed `kiosk-restart`.
 */
export function createKioskTimer(settings: ShowSettings, startT: number, sink?: DiagnosticSink): KioskTimer {
  const ms = settings.kioskRestartMs
  let enabled = settings.type === 'kiosk'
  if (enabled && !(Number.isFinite(ms) && ms > 0)) {
    enabled = false
    sink?.add({ code: 'kiosk-restart', message: `Kiosk restart ${ms} ms would restart on every frame; the show does not restart on its own` })
  }
  let last = startT
  return {
    restartDue: (t) => enabled && t - last >= ms,
    restarted(t) { last = t },
  }
}
