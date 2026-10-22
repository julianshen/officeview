/**
 * The Presenter View window (slideshow spec 5.1 second window and closing,
 * 5.3 injected popup stylesheet, 7 popup boundary).
 *
 * The eager start opens a same-origin `about:blank` popup inside the user
 * gesture (src/components/present.ts); once the chunk has loaded the session
 * adopts it here:
 * - a theme-root container is written into the popup's own document, so the
 *   scoped `[data-ov-theme-root]` rules apply there;
 * - officeview's stylesheet is injected: every stylesheet of the opener is
 *   copied, a linked one as a `<link>` to the same URL, an inline one as a
 *   `<style>` with the same text. Host nonces are never replayed into a
 *   window that receives deck content: the host's policy decides. The popup
 *   inherits the opener's CSP, so a refusal (`securitypolicyviolation` for
 *   a style directive, or a stylesheet `error`) is diagnosed once as
 *   `presenter-styles-blocked`; the controls stay usable unstyled;
 * - listeners go on the popup's own window and document (realm-safe);
 * - it closes on the opener's `pagehide`, and reports when the presenter
 *   closes it (the session then falls back to single-screen).
 *
 * The popup is not an isolation boundary (spec 7): it is same-origin with an
 * opener. Only deck content is rendered into it, and text goes in as text.
 */
import { themeRootProps } from '../../components/theme'
import type { DiagnosticSink } from '../types'

export interface PresenterWindow {
  readonly doc: Document
  readonly root: HTMLElement
  /** Idempotent. */
  close(): void
  /** Called once, when the popup closes for any reason (the presenter, `close()`, the opener unloading). */
  onClosed(cb: () => void): void
}

export const PRESENTER_STYLES_BLOCKED = 'presenter-styles-blocked'

export function adoptPresenterWindow(popup: Window, opener: Window, sink: DiagnosticSink): PresenterWindow {
  const doc = popup.document
  const listeners = new AbortController()
  const on = { signal: listeners.signal }
  const closedCbs: Array<() => void> = []
  let closed = false

  doc.title = 'Presenter View'
  const root = doc.createElement('div')
  for (const [k, v] of Object.entries(themeRootProps('system', 'presenter'))) root.setAttribute(k, v)
  root.setAttribute('data-officeview-presenter', '')
  root.tabIndex = -1
  doc.body.replaceChildren(root)

  const blocked = () => sink.add({ code: PRESENTER_STYLES_BLOCKED, message: "The host's Content Security Policy refused officeview's stylesheet in the presenter window; its controls are unstyled but usable." })
  doc.addEventListener('securitypolicyviolation', (e) => {
    if (String((e as SecurityPolicyViolationEvent).effectiveDirective ?? '').startsWith('style-src')) blocked()
  }, on)
  for (const sheet of [...opener.document.styleSheets]) {
    const node = sheet.ownerNode as Element | null
    let copy: HTMLLinkElement | HTMLStyleElement
    if (sheet.href) {
      const link = doc.createElement('link')
      link.rel = 'stylesheet'
      link.href = sheet.href
      copy = link
    } else if (node?.localName === 'style') {
      copy = doc.createElement('style')
      copy.textContent = node.textContent
    } else continue
    copy.addEventListener('error', blocked, on)
    doc.head.append(copy)
  }

  const finish = () => {
    if (closed) return
    closed = true
    listeners.abort()
    for (const cb of closedCbs.splice(0)) cb()
  }
  const close = () => {
    finish()
    if (!popup.closed) popup.close()
  }
  popup.addEventListener('pagehide', finish, on)
  opener.addEventListener('pagehide', close, on)

  return {
    doc,
    root,
    close,
    onClosed(cb) {
      if (closed) cb()
      else closedCbs.push(cb)
    },
  }
}
