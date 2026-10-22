/**
 * Slideshow Presenter View: Window Management and multi-screen placement (spec §5.1).
 *
 * Handles screen detection via the Window Management API, window placement planning,
 * audience fullscreen requests on external screens, popup positioning on primary screens,
 * fallback diagnostics, and display swapping.
 */
import type { DiagnosticSink, SlideShowDiagnostic } from '../types'

/** A representation of a screen (standard ScreenDetailed / Screen). */
export interface ScreenLike {
  readonly isPrimary?: boolean
  readonly isInternal?: boolean
  readonly left?: number
  readonly top?: number
  readonly width?: number
  readonly height?: number
  readonly availLeft?: number
  readonly availTop?: number
  readonly availWidth?: number
  readonly availHeight?: number
  readonly label?: string
}

/** Screen details returned by the Window Management API (`getScreenDetails`). */
export interface ScreenDetailsLike {
  readonly screens: ReadonlyArray<ScreenLike>
  readonly currentScreen?: ScreenLike
}

/** The computed plan for window and audience placement. */
export type PlacementPlan =
  | { popupFeatures: string; audienceScreen: ScreenLike; manual?: false }
  | { manual: true; reason: 'api-missing' | 'permission-denied' | 'single-screen' }

/**
 * Plans window and audience placement based on available screen details.
 *
 * When two or more screens are available:
 * - Presenter popup opens on the primary screen.
 * - Audience view goes fullscreen on the external screen.
 *
 * When unavailable, permission is denied, or only a single screen exists,
 * returns a manual fallback plan with the specific reason.
 */
export function planPlacement(
  details: ScreenDetailsLike | undefined,
  reason?: 'api-missing' | 'permission-denied',
): PlacementPlan {
  if (!details) {
    return { manual: true, reason: reason ?? 'api-missing' }
  }

  const screens = details.screens ?? []
  if (screens.length <= 1) {
    return { manual: true, reason: 'single-screen' }
  }

  const primary = screens.find((s) => s.isPrimary) ?? screens[0]
  const external = screens.find((s) => s !== primary) ?? screens[1]

  const left = primary.availLeft ?? primary.left ?? 0
  const top = primary.availTop ?? primary.top ?? 0
  const width = primary.availWidth ?? primary.width ?? 1024
  const height = primary.availHeight ?? primary.height ?? 768

  const popupFeatures = `left=${left},top=${top},width=${width},height=${height}`

  return {
    popupFeatures,
    audienceScreen: external,
  }
}

/**
 * Requests screen details from the Window Management API (`getScreenDetails`),
 * triggering a user prompt on click gesture if not yet granted.
 */
export async function useBothScreens(win: Window): Promise<ScreenDetailsLike | undefined> {
  try {
    const getScreenDetails =
      (win as any).getScreenDetails ??
      (win.navigator as any)?.getScreenDetails
    if (typeof getScreenDetails !== 'function') {
      return undefined
    }
    const details = await getScreenDetails.call(win)
    return details as ScreenDetailsLike
  } catch {
    return undefined
  }
}

/**
 * Injected slot for S2-1's eager preparedPlacement slot.
 * S2-1 will expose this on Document / present.ts; until S2-1 lands,
 * we store it in a WeakMap and allow setter injection.
 */
type PlacementSetter = (doc: Document, plan: PlacementPlan | undefined) => void
type PlacementGetter = (doc: Document) => PlacementPlan | undefined

const defaultSlot = new WeakMap<Document, PlacementPlan>()
let customSetter: PlacementSetter | undefined
let customGetter: PlacementGetter | undefined

export function setPreparedPlacementSlot(setter?: PlacementSetter, getter?: PlacementGetter): void {
  customSetter = setter
  customGetter = getter
}

export function getPreparedPlacement(doc: Document): PlacementPlan | undefined {
  if (customGetter) return customGetter(doc)
  return (doc as any)._preparedPlacement ?? defaultSlot.get(doc)
}

/**
 * Computes placement plan from preload hint details and stores it in
 * S2-1's eager preparedPlacement slot.
 */
export function preparePlacement(doc: Document, details: ScreenDetailsLike | undefined): void {
  const plan = planPlacement(details)
  if (customSetter) {
    customSetter(doc, plan)
  } else {
    defaultSlot.set(doc, plan)
    ;(doc as any)._preparedPlacement = plan
  }
}

/**
 * Session representation for swapping displays.
 */
export interface ScreenSessionLike {
  presenterScreen?: ScreenLike
  audienceScreen?: ScreenLike
  screens?: {
    presenter?: ScreenLike
    audience?: ScreenLike
    details?: ScreenDetailsLike
  }
  presenterWindow?: Window | {
    moveTo?(x: number, y: number): void
    resizeTo?(width: number, height: number): void
    screen?: ScreenLike
  }
  retryFullscreen?(options?: { screen?: ScreenLike }): Promise<any>
  requestFullscreen?(options?: { screen?: ScreenLike }): Promise<any>
  host?: HTMLElement
  onSwapScreens?(newAudienceScreen?: ScreenLike, newPresenterScreen?: ScreenLike): void
}

/**
 * Swaps the presenter popup and audience displays.
 * Moves the presenter window to the external display, and requests audience
 * fullscreen on the other display.
 */
export function swapDisplays(session: ScreenSessionLike): void {
  let presenterScreen: ScreenLike | undefined
  let audienceScreen: ScreenLike | undefined

  if (session.screens?.presenter && session.screens?.audience) {
    presenterScreen = session.screens.presenter
    audienceScreen = session.screens.audience
    session.screens.presenter = audienceScreen
    session.screens.audience = presenterScreen
  } else if (session.presenterScreen && session.audienceScreen) {
    presenterScreen = session.presenterScreen
    audienceScreen = session.audienceScreen
    session.presenterScreen = audienceScreen
    session.audienceScreen = presenterScreen
  }

  const newAudienceScreen = presenterScreen
  const newPresenterScreen = audienceScreen

  if (session.presenterWindow && newPresenterScreen) {
    const left = newPresenterScreen.availLeft ?? newPresenterScreen.left ?? 0
    const top = newPresenterScreen.availTop ?? newPresenterScreen.top ?? 0
    const width = newPresenterScreen.availWidth ?? newPresenterScreen.width ?? 1024
    const height = newPresenterScreen.availHeight ?? newPresenterScreen.height ?? 768
    session.presenterWindow.moveTo?.(left, top)
    session.presenterWindow.resizeTo?.(width, height)
  }

  if (newAudienceScreen) {
    if (typeof session.retryFullscreen === 'function') {
      void session.retryFullscreen({ screen: newAudienceScreen })
    } else if (typeof session.requestFullscreen === 'function') {
      void session.requestFullscreen({ screen: newAudienceScreen })
    } else if (session.host && typeof (session.host as any).requestFullscreen === 'function') {
      void (session.host as any).requestFullscreen({ screen: newAudienceScreen })
    }
  }

  session.onSwapScreens?.(newAudienceScreen, newPresenterScreen)
}

/**
 * Emits diagnostic for placement failure reasons to the diagnostic sink.
 * Codes: 'permission-denied' | 'api-missing' | 'single-screen'.
 * None is 'popup-blocked'.
 */
export function diagnosePlacement(
  target: 'api-missing' | 'permission-denied' | 'single-screen' | PlacementPlan,
  sink: DiagnosticSink,
): SlideShowDiagnostic | undefined {
  const reason = typeof target === 'string' ? target : target.manual ? target.reason : undefined
  if (!reason) return undefined

  let message = ''
  switch (reason) {
    case 'permission-denied':
      message = 'Window management permission was denied; manual placement required.'
      break
    case 'api-missing':
      message = 'Window Management API is not supported in this browser; manual placement required.'
      break
    case 'single-screen':
      message = 'Only a single screen was detected; cannot place presenter and audience on separate displays.'
      break
  }

  const diagnostic: SlideShowDiagnostic = { code: reason, message }
  sink.add(diagnostic)
  return diagnostic
}

export const DRAG_HINT_STORAGE_KEY = 'officeview:drag-hint-dismissed'

export interface DragHintOptions {
  storage?: Storage | null
  onDismiss?(): void
}

/**
 * Checks whether the one-time drag hint should be shown.
 */
export function shouldShowDragHint(storage: Storage | null = typeof sessionStorage !== 'undefined' ? sessionStorage : null): boolean {
  try {
    return storage?.getItem(DRAG_HINT_STORAGE_KEY) !== '1'
  } catch {
    return true
  }
}

/**
 * Marks the drag hint as dismissed in storage.
 */
export function markDragHintDismissed(storage: Storage | null = typeof sessionStorage !== 'undefined' ? sessionStorage : null): void {
  try {
    storage?.setItem(DRAG_HINT_STORAGE_KEY, '1')
  } catch {}
}

/**
 * Creates the one-time drag hint DOM element for other browsers (spec §5.1).
 * Returns undefined if already dismissed.
 */
export function createDragHint(
  doc: Document,
  options?: DragHintOptions,
): HTMLElement | undefined {
  const storage = options?.storage !== undefined ? options.storage : (doc.defaultView?.sessionStorage ?? null)
  if (!shouldShowDragHint(storage)) {
    return undefined
  }

  const hint = doc.createElement('div')
  hint.setAttribute('role', 'status')
  hint.setAttribute('data-officeview-drag-hint', '')
  hint.textContent = 'Drag this presenter window to your other display. '

  const dismiss = doc.createElement('button')
  dismiss.type = 'button'
  dismiss.textContent = 'Dismiss'
  dismiss.onclick = () => {
    markDragHintDismissed(storage)
    hint.remove()
    options?.onDismiss?.()
  }

  hint.append(dismiss)
  return hint
}

/**
 * Displays the one-time drag hint by appending it to the target document or element.
 */
export function showDragHint(
  target: HTMLElement | Document,
  options?: DragHintOptions,
): HTMLElement | undefined {
  const doc = target instanceof Document ? target : target.ownerDocument
  const root = target instanceof Document ? target.body : target
  const hint = createDragHint(doc, options)
  if (hint && root) {
    root.append(hint)
  }
  return hint
}
