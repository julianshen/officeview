/**
 * The eager slideshow start (slideshow spec 4.1, 4.2, 8): the host handle
 * types, the per-document show coordinator and the synchronous start
 * sequence. Everything after the chunk import lives in src/slideshow (lazy).
 *
 * Bytes here are measured against the eager budget: keep it lean, and keep
 * comments off class members, object-literal properties and JSX props
 * (the library build keeps those verbatim; slideshow plan S1-1).
 */
import type { OfficeDocSource } from './OfficeDoc'
import type { PptxDocument, PptxPackageHandle } from '../pptx/types'
import type { SlideShowDiagnostic } from '../slideshow/types'
import type { WatermarkOptions } from '../core/watermark'
import { pptxPackageHandle } from '../pptx/parse'

export type { SlideShowDiagnostic }

export type SlideShowErrorCode = 'gesture-required' | 'not-presentable' | 'already-presenting' | 'load-failed' | 'cancelled' | 'invalid-slide'

export class SlideShowError extends Error {
  constructor(readonly code: SlideShowErrorCode, message: string = code) {
    super(message)
  }
}

export interface PresentOptions {
  from?: 'beginning' | 'current' | number
  /** 'auto' (default): Presenter View when a second screen can be used, else audience only. */
  presenterView?: boolean | 'auto'
}

export interface SlideShowSession {
  /** 0-based, in deck order (hidden slides keep their index). */
  readonly slideIndex: number
  /** Click checkpoint within the slide. */
  readonly buildIndex: number
  readonly surface: 'fullscreen' | 'in-window'
  readonly presenter: 'window' | 'swap' | 'none'
  next(): void
  previous(): void
  goTo(slideIndex: number): void
  /** Idempotent. */
  exit(): void
  /** Resolves exactly once when the show exits for any reason; never rejects. */
  readonly ended: Promise<void>
}

/** The `ref` handle of `OfficeDoc` and `OfficeFile`. */
export interface OfficeViewerHandle {
  /** Call synchronously inside a user-gesture handler (spec 4.2). */
  present(opts?: PresentOptions): Promise<SlideShowSession>
}

/** Window Management (spec 5.1), not yet in the TypeScript DOM lib. */
export interface ScreenLike { availLeft?: number; availTop?: number; availWidth: number; availHeight: number; left?: number; top?: number; width?: number; height?: number; isPrimary?: boolean; isInternal?: boolean; label?: string }
export interface ScreenDetails { readonly screens: readonly ScreenLike[]; readonly currentScreen?: ScreenLike }
export type PresentFullscreenOptions = FullscreenOptions & { screen?: ScreenLike }
/** What the preload hint prepared for the synchronous start: popup placement features and the audience screen. */
export interface PreparedPlacement { popupFeatures: string; audienceScreen?: ScreenLike }

export interface PresentPlatform {
  hasActivation(win: Window): boolean | undefined
  requestFullscreen(el: Element, options?: PresentFullscreenOptions): Promise<void>
  exitFullscreen(doc: Document): Promise<void>
  fullscreenElement(doc: Document): Element | null
  open(win: Window, url: string, features: string): Window | null
  /** Screen details only where the window-management permission is already granted (no prompt); else undefined. */
  screenDetails?(): Promise<ScreenDetails | undefined>
}

// async: a missing or throwing requestFullscreen becomes a rejection (in-window).
export const defaultPlatform: PresentPlatform = {
  hasActivation: (w) => (w.navigator as { userActivation?: { isActive: boolean } }).userActivation?.isActive,
  requestFullscreen: async (el, o) => el.requestFullscreen(o),
  exitFullscreen: (d) => d.exitFullscreen(),
  fullscreenElement: (d) => d.fullscreenElement,
  open: (w, u, f) => w.open(u, '_blank', f),
  screenDetails: async () => (await navigator.permissions.query({ name: 'window-management' as PermissionName })).state == 'granted' ? (window as unknown as { getScreenDetails(): Promise<ScreenDetails> }).getScreenDetails() : undefined,
}

/**
 * Per document: the placement the preload hint prepared (spec 5.1; S2-3
 * fills it), so the synchronous start can open the popup on the presenter's
 * screen and request full screen on the audience screen without awaiting.
 */
export const preparedPlacement = new WeakMap<Document, PreparedPlacement>()

export interface ShowToken {
  release(): void
  readonly released: boolean
  /** True for the token's host element or a descendant. */
  owns(el: Element | null): boolean
}

const tokens = new WeakMap<Document, ShowToken>()

/** One show per document: undefined while another start or show holds the token. */
export function acquireShowToken(host: HTMLElement): ShowToken | undefined {
  const d = host.ownerDocument
  if (tokens.has(d)) return
  const t = {
    released: false,
    release() {
      tokens.get(d) == t && tokens.delete(d)
      t.released = true
    },
    owns: (el: Element | null) => !!el && host.contains(el),
  }
  tokens.set(d, t)
  return t
}

/** Bytes, a Blob or a URL the viewer can read again; undefined for one-shot sources. */
export type ReexportableSource = Blob | ArrayBuffer | Uint8Array | string

export interface SlideShowCallbacks {
  onStart?(): void
  onEnd?(): void
  onSlideChange?(slideIndex: number, buildIndex: number): void
  onDiagnostic?(d: SlideShowDiagnostic): void
  currentSlide(): number
  restoreFocus(): void
  scrollToSlide(i: number): void
}

export type SlideShowMotion = 'auto' | 'full' | 'reduced'

/** What the eager start hands the lazy chunk (`runShow`, src/slideshow). */
export interface ShowStart {
  doc: PptxDocument
  handle?: PptxPackageHandle
  source?: ReexportableSource
  host: HTMLElement
  token: ShowToken
  platform: PresentPlatform
  /** The synchronous request's outcome, so a late refusal reaches the session. */
  fullscreen: Promise<'entered' | 'refused'>
  /** Only from a new user gesture ("Click for full screen"); defaults to the prepared audience screen. */
  retryFullscreen(options?: PresentFullscreenOptions): Promise<'entered' | 'refused'>
  /** The provisional presenter window opened in the gesture: undefined when not wanted, null (or closed) when blocked. */
  popup?: Window | null
  from: 'beginning' | 'current' | number
  presenterView?: boolean | 'auto'
  motion: SlideShowMotion
  watermark?: WatermarkOptions
  callbacks: SlideShowCallbacks
  /** Aborted when the show must end: Esc, browser full-screen exit, unmount, document change, allowPresent=false. */
  signal: AbortSignal
  generation: number
  isCurrent(): boolean
  /** Exit cleanup, exactly once: (popup,) full screen if owned, hide the host, release the token. */
  cleanup(): void
}

export interface PresentRequest {
  doc: PptxDocument | OfficeDocSource
  host: HTMLElement
  opts?: PresentOptions
  platform: PresentPlatform
  generation: number
  source?: ReexportableSource
  /** The viewer's events; the lazy chunk supplies any viewer navigation left out (slide in view, scroll, focus restore). */
  callbacks: Partial<SlideShowCallbacks>
  allow: boolean
  motion: SlideShowMotion
  watermark?: WatermarkOptions
  /** The viewer's document epoch: aborting it cancels the start or ends the show. */
  signal?: AbortSignal
}

/** @internal Tests defer or fail the chunk through this. */
export const showLoader = { load: () => import('../slideshow') }

/**
 * The start sequence. Everything up to the import runs synchronously inside
 * the caller's gesture: checks, token, show the host, requestFullscreen.
 */
export function startPresent(req: PresentRequest): Promise<SlideShowSession> {
  const { doc, host, opts = {}, platform: P } = req
  const from = opts.from ?? 'beginning'
  const d = host.ownerDocument, w = d.defaultView!
  const code: SlideShowErrorCode | 0 = !req.allow || !('slides' in doc) ? 'not-presentable'
    : typeof from == 'number' && !doc.slides[from] ? 'invalid-slide'
    : P.hasActivation(w) === false ? 'gesture-required' : 0
  const token = code ? 0 : acquireShowToken(host)
  if (!token) return Promise.reject(new SlideShowError(code || 'already-presenting'))
  // Spec 4.1: the presenter popup first, then full screen, in the same gesture.
  // 'auto' wants it only with a prepared audience screen (window-management already
  // granted), so a plain start never spends its activation on a popup (controller ruling).
  const pv = opts.presenterView ?? 'auto', placed = preparedPlacement.get(d), screen = placed?.audienceScreen
  const popup = pv === true || pv == 'auto' && screen ? P.open(w, 'about:blank', placed?.popupFeatures ?? 'popup') : undefined
  let done: unknown, handed: unknown, entered: unknown, pending: unknown = 1
  const ac = new AbortController(), signal = ac.signal
  const end = () => ac.abort()
  const exit = () => {
    token.owns(P.fullscreenElement(d)) && P.exitFullscreen(d).catch(end)
  }
  // A start can fail before the browser has entered full screen: then the
  // exit waits for the synchronous request to settle.
  const cleanup = () => {
    if (done) return
    done = 1
    end()
    popup?.close()
    pending ? fullscreen.then(exit) : exit()
    host.style.visibility = 'hidden'
    token.release()
  }
  req.signal?.addEventListener('abort', end, { signal })
  // The opener unloading (even during load) takes the provisional popup with it.
  popup && w?.addEventListener('pagehide', () => popup.close(), { signal })
  d.addEventListener('fullscreenchange', () => {
    const el = P.fullscreenElement(d)
    el ? token.owns(el) && (entered = 1) : entered && end()
  }, { signal })
  host.onkeydown = (e) => {
    e.stopPropagation()
    e.key == 'Escape' && end()
  }
  host.style.visibility = 'visible'
  const back = d.activeElement
  const retryFullscreen = (o = screen && { screen }) => P.requestFullscreen(host, o).then(() => 'entered' as const, () => 'refused' as const)
  const fullscreen = retryFullscreen().finally(() => pending = 0)
  host.focus()
  return Promise.race([
    new Promise<never>((_, no) => signal.addEventListener('abort', () => handed || no(new SlideShowError('cancelled')))),
    showLoader.load().then((m) => {
      if (signal.aborted) return
      handed = 1
      return m.runShow({
        ...req, ...opts, doc: doc as PptxDocument, handle: pptxPackageHandle(doc as PptxDocument), token, fullscreen, retryFullscreen, from, signal, cleanup, popup,
        callbacks: { ...m.viewerCallbacks(host, back), ...req.callbacks },
      })
    }),
  ]).catch((e) => {
    cleanup()
    throw e instanceof SlideShowError ? e : new SlideShowError('load-failed', 'Slideshow could not load. Reload to retry.')
  }) as Promise<SlideShowSession>
}
