/**
 * The lazy slideshow chunk (slideshow spec 8). Never imported statically from eager code.
 *
 * `runShow` takes over from the eager start sequence (src/components/present.ts)
 * once the chunk has loaded: the session controller (player/session.ts)
 * honours the start contract (surface from the full-screen outcome, exit
 * semantics, events, cleanup exactly once) and runs the audience surface.
 * `viewerCallbacks` is called by the eager `startPresent`: keep the export.
 */
import type { ShowStart, ShowToken, SlideShowCallbacks } from '../components/present'
import { PAGE_PADDING_PX } from '../components/paged/metrics'
import { createSession, type SessionController } from './player/session'

export type { ShowStart, ShowToken, SessionController }
export type { SessionState, ShowInput } from './player/session'

/**
 * The viewer navigation a show needs, found from the host (a child of the
 * viewer's theme root, beside the page region): the slide in view (the page
 * under the region's vertical middle), scrolling the region to a slide (its
 * own scrollTop; never scrollIntoView, so host ancestors never scroll) and
 * focus restore (the element focused at start if still connected and not
 * the body, else the region). Kept lazy so the eager start stays small.
 */
export function viewerCallbacks(host: HTMLElement, back: Element | null): Pick<SlideShowCallbacks, 'currentSlide' | 'scrollToSlide' | 'restoreFocus'> {
  const region = () => host.parentElement?.querySelector<HTMLElement>(':scope > [data-officeview-root]') ?? null
  const pages = (r: HTMLElement) => [...r.querySelectorAll('canvas')]
  return {
    currentSlide() {
      const r = region()
      if (!r) return 0
      const box = r.getBoundingClientRect()
      const mid = box.top + box.height / 2
      const list = pages(r)
      const i = list.findIndex((c) => c.getBoundingClientRect().bottom > mid)
      return i < 0 ? Math.max(0, list.length - 1) : i
    },
    scrollToSlide(i) {
      const r = region()
      const page = r && pages(r)[i]
      if (!r || !page) return
      r.scrollTop += page.getBoundingClientRect().top - r.getBoundingClientRect().top - PAGE_PADDING_PX
    },
    restoreFocus() {
      // The body (an API-started show with nothing focused) is not a start element.
      const own = back?.ownerDocument
      const el = back?.isConnected && back !== own?.body && back !== own?.documentElement ? back : region()
      ;(el as HTMLElement | null)?.focus?.()
    },
  }
}

/**
 * The lazy start's handoff (spec 4.1, 4.2): the session controller and audience surface (player/session.ts).
 * `isCurrent` defaults to "the start's signal has not been aborted" (kept out of the eager start: S1-20 byte recovery).
 */
export function runShow(start: Omit<ShowStart, 'isCurrent'> & Partial<Pick<ShowStart, 'isCurrent'>>): Promise<SessionController> {
  return createSession({ isCurrent: () => !start.signal.aborted, ...start })
}
