/**
 * Hyperlinks in the show (slideshow spec 4.3): slide jumps, show jumps and
 * `http:`/`https:`/`mailto:` URLs opened in a new browsing context with
 * `noopener`. Every other scheme and action was refused by the reader
 * (`read/links.ts`) and is diagnosed here when clicked, never run.
 */
import type { DiagnosticSink, LinkSpec } from '../types'

/** What a link drives: the session's navigation (host-API semantics: a link is never a "click advance"). */
export interface LinkSession {
  goTo(slideIndex: number): void
  next(): void
  previous(): void
  first(): void
  last(): void
  exit(): void
}

export interface LinkOptions {
  sink?: DiagnosticSink
  slideIndex?: number
  /** Opens a URL; the eager platform's `open` in production (testable; never the global window). */
  open?: (win: Window, url: string, features: string) => Window | null
}

const SAFE = new Set(['http:', 'https:', 'mailto:'])

/** Runs `link` in `session`. URLs open in `win` with `noopener`; a refused link is diagnosed and does nothing else. */
export function runLink(link: LinkSpec, session: LinkSession, win: Window, options: LinkOptions = {}): void {
  const { sink, slideIndex } = options
  const refuse = (reason: string) => sink?.add({ code: 'link-refused', ...(slideIndex !== undefined ? { slideIndex } : {}), effect: `link:${link.spid}`, message: `Hyperlink refused (${reason})` })
  const a = link.action
  switch (a.kind) {
    case 'slide': session.goTo(a.index); return
    case 'jump':
      switch (a.verb) {
        case 'nextslide': session.next(); return
        case 'previousslide': session.previous(); return
        case 'firstslide': session.first(); return
        case 'lastslide': session.last(); return
        case 'endshow': session.exit(); return
      }
      return
    case 'url': {
      // Re-checked here: a caller-built LinkSpec never reaches window.open with another scheme.
      let protocol = ''
      try { protocol = new URL(a.href).protocol } catch { /* refused below */ }
      if (!SAFE.has(protocol)) { refuse(`scheme ${protocol || 'invalid'}`); return }
      const open = options.open ?? ((w: Window, url: string, features: string) => w.open(url, '_blank', features))
      try {
        open(win, a.href, 'noopener')
      } catch (e) {
        sink?.add({ code: 'link-open-failed', ...(slideIndex !== undefined ? { slideIndex } : {}), effect: `link:${link.spid}`, message: `The link could not be opened (${e instanceof Error ? e.message : String(e)})` })
      }
      return
    }
    case 'refused': refuse(a.reason); return
  }
}
