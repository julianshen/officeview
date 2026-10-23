/**
 * Presenter View placement with the Window Management API (slideshow spec 5.1).
 *
 * - **Prepared (already granted).** The preload hint (`prepareScreens`, called
 *   by the toolbar after the chunk loads) reads screen details without a
 *   prompt (`grantedScreenDetails`) and stores the plan in S2-2's eager `preparedPlacement` slot
 *   (src/components/present.ts). The synchronous start then opens the popup
 *   on the presenter's screen and requests full screen on the audience
 *   screen; `'auto'` opens the popup only when that slot holds a plan.
 * - **Not yet granted.** The start proceeds single-screen (or, from Alt+F5,
 *   with a default-size popup). "Use both screens" (`useBothScreens`) asks
 *   for the permission on its click and then places the windows.
 * - **Swap displays** (`swapDisplays`) exchanges the two screens.
 *
 * Best effort throughout (owner decision): nothing here throws or rejects.
 * Full screen is always requested first; the popup is moved and the state
 * flipped only once it entered, so a refusal never leaves the popup hidden
 * behind the audience. Denial, a missing API, a single screen and other
 * failures are diagnosed separately, never as `popup-blocked`.
 */
import { preparedPlacement, type PresentFullscreenOptions, type ScreenDetails, type ScreenLike } from '../../components/present'
import type { DiagnosticSink, SlideShowDiagnostic } from '../types'

export type { ScreenDetails, ScreenLike }

export type ManualReason = 'api-missing' | 'permission-denied' | 'single-screen'
export type PlacementFailure = ManualReason | 'failed'

/** Where the presenter popup and the audience go. `popupFeatures` and `audienceScreen` are the eager `PreparedPlacement`. */
export interface Placement { popupFeatures: string; presenterScreen: ScreenLike; audienceScreen: ScreenLike }
export type PlacementPlan = Placement | { manual: true; reason: ManualReason }

/** What placement acts on: the session's presenter popup and `ShowStart.retryFullscreen` (supplied by S2-4/S2-5). */
export interface PlacementTarget {
  /** The presenter popup; null or undefined in single-screen mode. */
  popup?: Window | null
  retryFullscreen(options?: PresentFullscreenOptions): Promise<'entered' | 'refused'>
  /** The current assignment, set by a successful placement; undefined until then. */
  state?: { presenter: ScreenLike; audience: ScreenLike }
}

const finite = (...v: Array<number | undefined>): number => Math.round(v.find((n) => typeof n == 'number' && Number.isFinite(n)) ?? 0)
/** A screen with a usable work area; anything else is not counted (no fabricated geometry). */
const usable = (s: ScreenLike | undefined): s is ScreenLike =>
  !!s && Number.isFinite(s.availWidth) && Number.isFinite(s.availHeight) && s.availWidth > 0 && s.availHeight > 0

/** A screen's work area in screen coordinates (offsets may be negative). */
export function screenBox(s: ScreenLike): { left: number; top: number; width: number; height: number } {
  return { left: finite(s.availLeft, s.left), top: finite(s.availTop, s.top), width: Math.round(s.availWidth), height: Math.round(s.availHeight) }
}

/**
 * The presenter goes on the primary screen (PowerPoint's default; else the
 * screen the opener is on, else the first), the audience on another one,
 * preferring an external screen (`isInternal === false`) and then the
 * platform's order. `undefined` details mean the API is missing: callers
 * that cannot tell "missing" from "not granted" use `requestScreens`.
 *
 * `width`/`height` are the work area. They are inner sizes in `window.open`,
 * so Chromium clamps the window to the work area; `applyPlacement` resizes
 * with outer sizes.
 */
export function planPlacement(details: ScreenDetails | undefined): PlacementPlan {
  if (!details) return { manual: true, reason: 'api-missing' }
  const screens = (details.screens ?? []).filter(usable)
  if (screens.length < 2) return { manual: true, reason: 'single-screen' }
  const current = details.currentScreen && screens.includes(details.currentScreen) ? details.currentScreen : undefined
  const presenter = screens.find((s) => s.isPrimary) ?? current ?? screens[0]
  const others = screens.filter((s) => s !== presenter)
  const audience = others.find((s) => s.isInternal === false) ?? others[0]
  const b = screenBox(presenter)
  return { popupFeatures: `popup=yes,left=${b.left},top=${b.top},width=${b.width},height=${b.height}`, presenterScreen: presenter, audienceScreen: audience }
}

/**
 * Stores the plan in the eager slot the synchronous start reads, or clears
 * it, so a screen unplugged since the last hint cannot survive.
 */
export function preparePlacement(doc: Document, details: ScreenDetails | undefined): void {
  try {
    const plan = planPlacement(details)
    if ('manual' in plan) preparedPlacement.delete(doc)
    else preparedPlacement.set(doc, { popupFeatures: plan.popupFeatures, audienceScreen: plan.audienceScreen, presenterScreen: plan.presenterScreen })
  } catch {
    preparedPlacement.delete(doc)
  }
}

/**
 * Screen details only where the window-management permission is already
 * granted, so it never prompts; undefined otherwise (not granted, no API, no
 * permission name as in Firefox and Safari). Realm-safe: it reads `win`'s
 * own navigator. Never rejects.
 */
export function grantedScreenDetails(win: Window | null | undefined): Promise<ScreenDetails | undefined> {
  return Promise.resolve()
    .then(() => win!.navigator.permissions.query({ name: 'window-management' as PermissionName }))
    .then((p) => p.state == 'granted' ? (win as unknown as { getScreenDetails(): Promise<ScreenDetails> }).getScreenDetails() : undefined)
    .catch(() => undefined)
}

/**
 * The preload hint (spec 4.1): prepares the placement from already-granted
 * details of `doc`'s window, or clears it. Never rejects. (It does not go
 * through the eager `defaultPlatform.screenDetails`, so the eager entry
 * exports nothing more for it.)
 */
export function prepareScreens(doc: Document, screenDetails: () => Promise<ScreenDetails | undefined> = () => grantedScreenDetails(doc.defaultView)): Promise<void> {
  return Promise.resolve()
    .then(() => screenDetails())
    .then((d) => preparePlacement(doc, d), () => preparePlacement(doc, undefined))
}

export type ScreenRequest = { details: ScreenDetails } | { reason: 'api-missing' | 'permission-denied' | 'failed' }

const failure = (e: unknown): ScreenRequest => ({ reason: (e as { name?: unknown } | null)?.name == 'NotAllowedError' ? 'permission-denied' : 'failed' })

/**
 * `getScreenDetails()` on `win`, called synchronously (inside the caller's
 * click activation), which prompts where the permission is not yet granted.
 * Pass the viewer window (the popup's opener), never the popup itself.
 * A missing function is `api-missing`, `NotAllowedError` is
 * `permission-denied`, anything else (a `SecurityError` from a frame without
 * `allow="window-management"`) is `failed`.
 */
export function requestScreens(win: Window): Promise<ScreenRequest> {
  try {
    const get = (win as unknown as { getScreenDetails?: () => Promise<ScreenDetails> }).getScreenDetails
    if (typeof get != 'function') return Promise.resolve({ reason: 'api-missing' })
    return Promise.resolve(get.call(win)).then((details) => ({ details }), failure)
  } catch (e) {
    return Promise.resolve(failure(e))
  }
}

/** Moves and sizes the popup to a screen's work area (outer size); best effort. */
function movePopup(popup: Window | null | undefined, s: ScreenLike): void {
  try {
    if (!popup || popup.closed) return
    const b = screenBox(s)
    popup.moveTo(b.left, b.top)
    popup.resizeTo(b.width, b.height)
  } catch { /* the popup is best effort */ }
}

/**
 * Full screen on `audience` first; only when it entered, the popup goes to
 * `presenter` and `target.state` is updated. A refused, rejecting or
 * throwing request leaves both windows and the state as they were. The
 * request needs a gesture in the opener: a click in the popup does not
 * activate the opener, so a refusal is retried on a later gesture.
 */
export function applyPlacement(target: PlacementTarget, presenter: ScreenLike, audience: ScreenLike): Promise<'placed' | 'refused'> {
  return Promise.resolve()
    .then(() => target.retryFullscreen({ screen: audience }))
    .then((r) => {
      if (r != 'entered') return 'refused' as const
      target.state = { presenter, audience }
      movePopup(target.popup, presenter)
      return 'placed' as const
    }, () => 'refused' as const)
}

export type BothScreensResult = { plan: Placement; outcome: 'placed' | 'refused' } | { reason: PlacementFailure }

/**
 * "Use both screens": asks for the permission on this click (pass the viewer
 * window), then places the windows. The plan is also prepared for the next
 * start. Failures are diagnosed (when a sink is given) and returned; a
 * refused full screen keeps the plan for `applyPlacement` on a later gesture.
 * The promise stays pending while the browser's prompt is open, so the
 * control shows its busy state until it settles.
 */
export function useBothScreens(win: Window, target: PlacementTarget, sink?: DiagnosticSink): Promise<BothScreensResult> {
  const doc = (() => { try { return win.document as Document | undefined } catch { return undefined } })()
  return requestScreens(win).then((r): Promise<BothScreensResult> | BothScreensResult => {
    const plan: PlacementPlan = 'reason' in r ? { manual: true, reason: r.reason as ManualReason } : planPlacement(r.details)
    if (doc) preparePlacement(doc, 'details' in r ? r.details : undefined)
    if ('manual' in plan) {
      if (sink) diagnosePlacement(plan.reason, sink)
      return { reason: plan.reason }
    }
    return applyPlacement(target, plan.presenterScreen, plan.audienceScreen).then((outcome) => {
      if (outcome == 'refused' && sink) diagnosePlacement('fullscreen-refused', sink)
      return { plan, outcome }
    })
  }).catch((): BothScreensResult => ({ reason: 'failed' }))
}

/**
 * "Swap displays": the audience goes full screen on the presenter's screen
 * and, once that entered, the popup moves to the old audience screen.
 * Resolves `'refused'` (diagnosed) and changes nothing when the request is
 * refused or there is no placement yet. Never rejects.
 */
export function swapDisplays(target: PlacementTarget, sink?: DiagnosticSink): Promise<'placed' | 'refused'> {
  const s = target.state
  if (!s) return Promise.resolve('refused')
  return applyPlacement(target, s.audience, s.presenter).then((outcome) => {
    if (outcome == 'refused' && sink) diagnosePlacement('fullscreen-refused', sink)
    return outcome
  })
}

const MESSAGES: Record<PlacementFailure | 'fullscreen-refused', string> = {
  'permission-denied': 'Window management permission was denied; drag the presenter window to the other screen.',
  'api-missing': 'This browser cannot place windows on screens; drag the presenter window to the other screen.',
  'single-screen': 'Only one screen was found; presenting on this screen.',
  failed: 'Screen details are not available here; drag the presenter window to the other screen.',
  'fullscreen-refused': 'Full screen on the other display needs another click.',
}

/**
 * Diagnoses a placement outcome: `placement-permission-denied`,
 * `placement-api-missing`, `placement-single-screen`, `placement-failed` or
 * `placement-fullscreen-refused` (never `popup-blocked`). A placed plan
 * diagnoses nothing.
 */
export function diagnosePlacement(target: PlacementFailure | 'fullscreen-refused' | PlacementPlan, sink: DiagnosticSink): SlideShowDiagnostic | undefined {
  const reason = typeof target == 'string' ? target : 'manual' in target ? target.reason : undefined
  if (!reason) return
  const d: SlideShowDiagnostic = { code: `placement-${reason}`, message: MESSAGES[reason] }
  sink.add(d)
  return d
}

export const DRAG_HINT_STORAGE_KEY = 'officeview:drag-hint-dismissed'

export interface DragHintOptions {
  /** Defaults to the document's origin `localStorage`; null keeps no record. */
  storage?: Storage | null
  onDismiss?(): void
}

/**
 * The origin's localStorage: a popup's sessionStorage is a copy that dies
 * with the popup, so a dismissal there would never reach the next show.
 * Blocked storage (a sandboxed frame) throws on access: no record then.
 */
function storageOf(doc: Document | null | undefined): Storage | null {
  try {
    return doc?.defaultView?.localStorage ?? null
  } catch {
    return null
  }
}

export function shouldShowDragHint(storage: Storage | null): boolean {
  try {
    return storage?.getItem(DRAG_HINT_STORAGE_KEY) !== '1'
  } catch {
    return true
  }
}

export function markDragHintDismissed(storage: Storage | null): void {
  try {
    storage?.setItem(DRAG_HINT_STORAGE_KEY, '1')
  } catch { /* not remembered: the hint shows again next time */ }
}

/**
 * The one-time drag hint for browsers without window placement (spec 5.1),
 * created in `doc` (realm-safe). Undefined once dismissed.
 */
export function createDragHint(doc: Document, options?: DragHintOptions): HTMLElement | undefined {
  const storage = options?.storage !== undefined ? options.storage : storageOf(doc)
  if (!shouldShowDragHint(storage)) return
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
 * Appends the hint to an element, or to a document's body. Realm-safe: a
 * popup's document is another realm, so it is told apart by `nodeType`.
 * Never throws.
 */
export function showDragHint(target: Element | Document, options?: DragHintOptions): HTMLElement | undefined {
  try {
    const isDoc = target.nodeType === 9
    const doc = isDoc ? (target as Document) : target.ownerDocument
    const root = isDoc ? (target as Document).body : target
    if (!doc || !root) return
    const hint = createDragHint(doc, options)
    if (hint) root.append(hint)
    return hint
  } catch {
    return undefined
  }
}
