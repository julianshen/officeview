/**
 * Presenter View's screen controls (slideshow spec 5.1 placement, swap and
 * the drag hint), wired over S2-3's screens.ts.
 *
 * The `PlacementTarget` is the show's popup and `ShowStart.retryFullscreen`.
 * A start whose preload hint prepared both screens (window-management
 * already granted) begins placed once its full screen entered, so "Swap
 * displays" works at once. Otherwise, where the API exists, "Use both
 * screens" asks for the permission on its click; where it does not, the
 * one-time drag hint shows in the popup.
 *
 * A click inside the popup gives the opener no user activation, so the
 * second full-screen request is usually refused there. The refusal leaves
 * both windows as they were and puts a retry in the opener's show host:
 * its click is an opener gesture, and the placement is applied again.
 */
import type { PresentFullscreenOptions, PreparedPlacement, ScreenLike } from '../../components/present'
import type { DiagnosticSink } from '../types'
import { applyPlacement, showDragHint, swapDisplays, useBothScreens, type PlacementTarget } from './screens'

export interface PresenterPlacement {
  /** "Swap displays" is offered: both screens are placed. */
  readonly canSwap: boolean
  /** "Use both screens" is offered: the API exists and nothing is placed yet. */
  readonly canUseBoth: boolean
  /** A request is in flight (the permission prompt may be open). */
  readonly busy: boolean
  swap(): void
  useBoth(): void
  /** Shows the one-time drag hint (in the popup) where placement is unavailable. */
  hint(target: Element | Document): void
  dispose(): void
}

export interface PlacementDeps {
  popup: Window
  /** The viewer window (the popup's opener): permission requests and retries run against it. */
  opener: Window
  /** The show host in the opener, where the retry notice goes. */
  host: HTMLElement
  retryFullscreen(options?: PresentFullscreenOptions): Promise<'entered' | 'refused'>
  /** The start's full-screen outcome; a prepared placement counts once it entered. */
  fullscreen: Promise<'entered' | 'refused'>
  prepared: PreparedPlacement | undefined
  sink: DiagnosticSink
  /** Called whenever the offer or the busy state changes. */
  changed(): void
}

const hasApi = (w: Window) => {
  try { return typeof (w as unknown as { getScreenDetails?: unknown }).getScreenDetails === 'function' } catch { return false }
}

export function createPresenterPlacement(d: PlacementDeps): PresenterPlacement {
  const target: PlacementTarget = { popup: d.popup, retryFullscreen: d.retryFullscreen }
  let busy = false
  let live = true
  let notice: HTMLElement | undefined
  const { presenterScreen, audienceScreen } = d.prepared ?? {}
  if (presenterScreen && audienceScreen) {
    void d.fullscreen.then((r) => {
      if (!live || r !== 'entered' || target.state) return
      target.state = { presenter: presenterScreen, audience: audienceScreen }
      d.changed()
    })
  }

  const dropNotice = () => { notice?.remove(); notice = undefined }

  /** The opener-side retry: its click has the opener's activation. */
  const offerRetry = (presenter: ScreenLike, audience: ScreenLike) => {
    if (!live) return
    dropNotice()
    const doc = d.host.ownerDocument
    const n = notice = doc.createElement('div')
    n.setAttribute('role', 'status')
    n.setAttribute('data-officeview-show-notice', '')
    n.setAttribute('data-placement', '')
    n.textContent = 'The browser needs a click in this window to move the slide show. '
    const go = doc.createElement('button')
    go.type = 'button'
    go.textContent = 'Move the slide show'
    go.onclick = () => {
      void run(applyPlacement(target, presenter, audience).then((r) => {
        if (r === 'placed' && notice === n) dropNotice()
        return r
      }))
    }
    const dismiss = doc.createElement('button')
    dismiss.type = 'button'
    dismiss.textContent = 'Dismiss'
    dismiss.onclick = () => { if (notice === n) dropNotice() }
    n.append(go, dismiss)
    d.host.append(n)
  }

  const run = <T>(p: Promise<T>): Promise<T | undefined> => {
    busy = true
    d.changed()
    return p.finally(() => {
      busy = false
      if (live) d.changed()
    }).catch(() => undefined)
  }

  return {
    get canSwap() { return !!target.state },
    get canUseBoth() { return !target.state && hasApi(d.opener) },
    get busy() { return busy },
    swap() {
      const s = target.state
      if (!s || busy) return
      void run(swapDisplays(target, d.sink).then((r) => {
        if (r === 'refused') offerRetry(s.audience, s.presenter)
        else dropNotice()
        return r
      }))
    },
    useBoth() {
      if (busy) return
      // getScreenDetails is called synchronously inside this click (S2-3), on the opener.
      void run(useBothScreens(d.opener, target, d.sink).then((r) => {
        if ('plan' in r && r.outcome === 'refused') offerRetry(r.plan.presenterScreen, r.plan.audienceScreen)
        else dropNotice()
        return r
      }))
    },
    hint(at) {
      if (!(presenterScreen && audienceScreen) && !target.state && !hasApi(d.opener)) showDragHint(at)
    },
    dispose() {
      live = false
      dropNotice()
    },
  }
}
