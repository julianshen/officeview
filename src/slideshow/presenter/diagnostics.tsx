/**
 * The Presenter View diagnostics panel (slideshow spec 5.2, 6.4): approximate,
 * pending and refused effects listed in a collapsible panel, never on the
 * audience picture. The show's sink already deduplicates by code, slide and
 * effect, so each entry appears once; the panel only groups them.
 */
import type { ReactElement } from 'react'
import type { SlideShowDiagnostic } from '../types'

export type DiagnosticKind = 'refused' | 'approximate' | 'pending' | 'other'

const KINDS: ReadonlyArray<[DiagnosticKind, RegExp, string]> = [
  ['refused', /refused|unknown|unsupported|blocked|limit|invalid|unreadable|failed|missing|denied/, 'Refused'],
  ['approximate', /approx|stand-in|whole|fallback|inherited|scale|ladder|stepdown|placeholder/, 'Approximate'],
  ['pending', /pending|autoplay|requires|kiosk|prime/, 'Pending'],
  ['other', /(?:)/, 'Other notes'],
]

/** The panel group of a diagnostic, read from its code (and effect). */
export function diagnosticKind(d: SlideShowDiagnostic): DiagnosticKind {
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
