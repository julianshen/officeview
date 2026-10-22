/**
 * The lazy slideshow chunk (slideshow spec 8). Never imported statically from eager code.
 *
 * `runShow` takes over from the eager start sequence (src/components/present.ts)
 * once the chunk has loaded. Slideshow plan S1-17 lands it as a stub session:
 * it honours the start contract (surface from the full-screen outcome, exit
 * semantics, events, cleanup exactly once) and paints a placeholder surface.
 * S1-19 replaces the body with the session controller and audience surface.
 */
import { SlideShowError, type ShowStart, type ShowToken, type SlideShowCallbacks, type SlideShowSession } from '../components/present'
import { PAGE_PADDING_PX } from '../components/paged/metrics'

export type { ShowStart, ShowToken }

/** The running show. S1-19 adds `dispatch(input)` and `state`. */
export interface SessionController extends SlideShowSession {
  readonly surface: 'fullscreen' | 'in-window'
}

/**
 * The viewer navigation a show needs, found from the host (a child of the
 * viewer's theme root, beside the page region): the slide in view (the page
 * under the region's vertical middle), scrolling the region to a slide (its
 * own scrollTop; never scrollIntoView, so host ancestors never scroll) and
 * focus restore (the element focused at start if still connected, else the
 * region). Kept lazy so the eager start stays small.
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
      const el = back?.isConnected ? back : region()
      ;(el as HTMLElement | null)?.focus?.()
    },
  }
}

export async function runShow(start: ShowStart): Promise<SessionController> {
  const { host, callbacks: cb, signal, doc } = start
  const count = doc.slides.length
  const win = host.ownerDocument
  if (start.presenterView === true) {
    cb.onDiagnostic?.({ code: 'presenter-view-unavailable', message: 'Presenter View is not available yet; presenting audience only.' })
  }
  const first = start.from === 'beginning' ? 0 : start.from === 'current' ? cb.currentSlide() : start.from
  let slideIndex = Math.min(Math.max(0, first), Math.max(0, count - 1))
  let surface: 'fullscreen' | 'in-window' = 'in-window'
  let exited = false
  let resolveEnded!: () => void
  const ended = new Promise<void>((resolve) => { resolveEnded = resolve })

  // Placeholder surface: its presence ends the host's busy (:empty) indicator.
  const stage = win.createElement('div')
  stage.setAttribute('data-officeview-show-surface', '')
  stage.setAttribute('aria-live', 'polite')
  const label = () => { stage.textContent = `Slide ${slideIndex + 1} of ${count}` }

  const goTo = (i: number) => {
    if (exited) return
    if (!Number.isInteger(i) || i < 0 || i >= count) {
      cb.onDiagnostic?.({ code: 'invalid-slide', message: `goTo(${i}) is outside the deck; ignored.` })
      return
    }
    slideIndex = i
    label()
    cb.onSlideChange?.(i, 0)
  }

  const session: SessionController = {
    get slideIndex() { return slideIndex },
    get buildIndex() { return 0 },
    get surface() { return surface },
    get presenter() { return 'none' as const },
    next: () => { if (slideIndex + 1 < count) goTo(slideIndex + 1) },
    previous: () => { if (slideIndex > 0) goTo(slideIndex - 1) },
    goTo,
    exit() {
      if (exited) return
      exited = true
      host.replaceChildren()
      start.cleanup()
      cb.scrollToSlide(slideIndex)
      cb.restoreFocus()
      cb.onEnd?.()
      resolveEnded()
    },
    ended,
  }

  if (!start.isCurrent()) {
    start.cleanup()
    throw new SlideShowError('cancelled')
  }
  signal.addEventListener('abort', () => session.exit())

  start.fullscreen.then((outcome) => {
    if (exited || !start.isCurrent()) return
    if (outcome === 'entered') { surface = 'fullscreen'; return }
    surface = 'in-window'
    cb.onDiagnostic?.({ code: 'fullscreen-refused', message: 'Full screen was refused; the show runs in the window.' })
    const notice = win.createElement('div')
    notice.setAttribute('role', 'status')
    notice.setAttribute('data-officeview-show-notice', '')
    notice.textContent = 'Full screen is not available here; presenting in the window. '
    const retry = win.createElement('button')
    retry.type = 'button'
    retry.textContent = 'Click for full screen'
    retry.onclick = () => {
      void start.retryFullscreen().then((r) => { if (r === 'entered' && !exited && start.isCurrent()) { surface = 'fullscreen'; notice.remove() } })
    }
    const dismiss = win.createElement('button')
    dismiss.type = 'button'
    dismiss.textContent = 'Dismiss'
    dismiss.onclick = () => notice.remove()
    notice.append(retry, dismiss)
    host.append(notice)
  })

  label()
  host.append(stage)
  cb.onStart?.()
  return session
}
