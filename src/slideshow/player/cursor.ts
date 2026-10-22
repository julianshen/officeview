/**
 * The show cursor (slideshow spec 4.3): it hides after 3 s without movement
 * and comes back when the mouse moves. The timer is injected (gc.md: every
 * clock is driven); it is not the show clock, so pausing the show does not
 * keep the cursor on screen.
 */

export const CURSOR_HIDE_MS = 3000

export interface CursorTimer {
  set(cb: () => void, ms: number): unknown
  clear(id: unknown): void
}

export interface CursorHider {
  readonly hidden: boolean
  dispose(): void
}

/** Hides the cursor over `el` after `CURSOR_HIDE_MS` without a `pointermove`/`mousemove`; listeners use `el`'s own realm. */
export function createCursorHider(el: HTMLElement, timer: CursorTimer, ms = CURSOR_HIDE_MS): CursorHider {
  let hidden = false
  let id: unknown
  let disposed = false
  const arm = () => {
    if (id !== undefined) timer.clear(id)
    id = timer.set(() => {
      id = undefined
      if (disposed) return
      hidden = true
      el.setAttribute('data-officeview-show-cursor', 'hidden')
    }, ms)
  }
  const moved = () => {
    if (disposed) return
    if (hidden) {
      hidden = false
      el.removeAttribute('data-officeview-show-cursor')
    }
    arm()
  }
  el.addEventListener('pointermove', moved)
  el.addEventListener('mousemove', moved)
  arm()
  return {
    get hidden() { return hidden },
    dispose() {
      if (disposed) return
      disposed = true
      if (id !== undefined) timer.clear(id)
      el.removeEventListener('pointermove', moved)
      el.removeEventListener('mousemove', moved)
      el.removeAttribute('data-officeview-show-cursor')
    },
  }
}
