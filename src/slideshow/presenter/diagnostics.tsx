/**
 * The Presenter View diagnostics panel (slideshow spec 5.2, 6.4): approximate,
 * pending and refused effects listed in a collapsible panel, never on the
 * audience picture. The show's sink already deduplicates by code, slide and
 * effect, so each entry appears once; the panel only groups them.
 */
import type { ReactElement } from 'react'
import { DIAGNOSTIC_CODES, type DiagnosticClass } from '../diagnostics'
import type { SlideShowDiagnostic } from '../types'

export type DiagnosticKind = 'refused' | 'approximate' | 'pending' | 'other'

const KINDS: ReadonlyArray<[DiagnosticKind, RegExp, string]> = [
  ['refused', /refused|unknown|unsupported|blocked|limit|invalid|unreadable|failed|missing|denied/, 'Refused'],
  ['approximate', /approx|stand-in|whole|fallback|inherited|scale|ladder|stepdown|placeholder/, 'Approximate'],
  ['pending', /pending|autoplay|requires|kiosk|prime/, 'Pending'],
  ['other', /(?:)/, 'Other notes'],
]

/**
 * The panel group says what the presenter sees; the catalogue class (`DIAGNOSTIC_CODES`) says why. A catalogued
 * code follows its class: `refused` → Refused, `approximate` and `stand-in` → Approximate, `info` → Other notes.
 * `resource` (platform or memory pressure) covers outcomes the presenter sees differently, so each such code the
 * presenter is likely to meet is placed here: Refused when the thing did not happen (no popup, no styles, a clip
 * not decoded), Approximate when it happened in a reduced form, Pending when a click or a retry can still start
 * it. These entries may only refine a `resource` code; tests/slideshow-presenter-view.test.tsx checks that every
 * other entry agrees with the class. Anything else (other `resource` codes, uncatalogued codes) is read from its wording.
 */
export const PANEL_KIND: Readonly<Record<string, DiagnosticKind>> = {
  'presenter-paint': 'refused',
  'presenter-thumbnail-approximate': 'approximate',
  'presenter-styles-blocked': 'refused',
  'presenter-window-failed': 'refused',
  'presenter-surface': 'approximate',
  'preview-prime': 'pending',
  'popup-blocked': 'refused',
  'fullscreen-refused': 'refused',
  'placement-permission-denied': 'refused',
  'placement-api-missing': 'refused',
  'placement-single-screen': 'other',
  'placement-failed': 'refused',
  'placement-fullscreen-refused': 'refused',
  'media-autoplay': 'pending',
  'media-codec': 'refused',
  'media-error': 'refused',
  'text-build-whole': 'approximate',
}

const BY_CLASS: Readonly<Partial<Record<DiagnosticClass, DiagnosticKind>>> = {
  refused: 'refused', approximate: 'approximate', 'stand-in': 'approximate', info: 'other',
}

/** The panel group of a diagnostic: an explicit entry first, then the catalogue class, else read from its code (and effect). */
export function diagnosticKind(d: SlideShowDiagnostic): DiagnosticKind {
  const own = (o: object) => Object.prototype.hasOwnProperty.call(o, d.code)
  const known = own(PANEL_KIND) ? PANEL_KIND[d.code] : own(DIAGNOSTIC_CODES) ? BY_CLASS[DIAGNOSTIC_CODES[d.code]] : undefined
  if (known) return known
  const s = `${d.code} ${d.effect ?? ''}`
  return KINDS.find(([, re]) => re.test(s))![0]
}

export function DiagnosticsPanel({ list }: { list: readonly SlideShowDiagnostic[] }): ReactElement {
  return (
    <details data-officeview-presenter-diagnostics="">
      <summary>{`Effect notes (${list.length})`}</summary>
      {list.length === 0 && <p>Every effect on the slides so far plays as authored.</p>}
      {KINDS.map(([kind, , title]) => {
        const items = list.filter((d) => diagnosticKind(d) === kind)
        if (!items.length) return null
        return (
          <section key={kind} data-kind={kind} aria-label={title}>
            <h3>{title}</h3>
            <ul>
              {items.map((d, i) => <li key={i}>{d.slideIndex === undefined ? d.message : `Slide ${d.slideIndex + 1}: ${d.message}`}</li>)}
            </ul>
          </section>
        )
      })}
    </details>
  )
}
