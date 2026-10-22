/**
 * The Presenter View basics inside the popup (slideshow spec 5.2 notes,
 * slide n of m and controls; 9 labelled controls and keyboard parity). S2-5
 * builds the full Presenter View (mirror, next preview, timer, grid) on the
 * same window; this is the part the window lifecycle needs to be usable.
 *
 * Plain DOM created through the popup's own document (realm-safe), updated
 * synchronously from the session. Notes and labels are text nodes: deck text
 * is never parsed as HTML. Inline layout goes through the CSSOM only, which
 * a host CSP never blocks.
 */
import type { NoteParagraph } from '../types'
import type { PresenterWindow } from './window'

export interface PresenterView {
  slide: number | 'end'
  count: number
  notes: readonly NoteParagraph[]
}

export interface PresenterActions {
  next(): void
  previous(): void
  exit(): void
  /** Keys on the popup's document: the session applies the audience key table (spec 5.2 accessibility). */
  key(e: KeyboardEvent): void
}

export interface PresenterControls {
  sync(view: PresenterView): void
  dispose(): void
}

export function mountPresenterControls(w: PresenterWindow, act: PresenterActions): PresenterControls {
  const doc = w.doc
  const listeners = new AbortController()
  const el = <K extends keyof HTMLElementTagNameMap>(tag: K, attr: string) => {
    const e = doc.createElement(tag)
    e.setAttribute(attr, '')
    return e
  }
  const section = el('section', 'data-officeview-presenter-view')
  section.setAttribute('aria-label', 'Presenter View')
  const status = el('p', 'data-officeview-presenter-status')
  const bar = el('div', 'data-officeview-presenter-controls')
  bar.setAttribute('role', 'toolbar')
  bar.setAttribute('aria-label', 'Slide controls')
  const button = (label: string, text: string, run: () => void) => {
    const b = doc.createElement('button')
    b.type = 'button'
    b.setAttribute('aria-label', label)
    b.textContent = text
    b.addEventListener('click', run, { signal: listeners.signal })
    return b
  }
  bar.append(button('Previous slide', 'Previous', act.previous), button('Next slide', 'Next', act.next), button('End show', 'End show', act.exit))
  const notes = el('div', 'data-officeview-presenter-notes')
  notes.setAttribute('role', 'region')
  notes.setAttribute('aria-label', 'Speaker notes')
  notes.tabIndex = 0
  section.append(status, bar, notes)
  w.root.append(section)
  doc.addEventListener('keydown', act.key, { signal: listeners.signal })
  w.root.focus()

  let key = ''
  let shownNotes: readonly NoteParagraph[] | undefined
  return {
    sync(view) {
      const next = `${view.slide}/${view.count}`
      if (next !== key) {
        key = next
        status.textContent = view.slide === 'end' ? 'End of slide show' : `Slide ${view.slide + 1} of ${view.count}`
      }
      if (view.notes === shownNotes) return
      shownNotes = view.notes
      notes.replaceChildren(...view.notes.map((n) => {
        const p = doc.createElement('p')
        // A bullet character as authored; an auto-number scheme name reads as a plain bullet here.
        p.textContent = n.bullet ? `${n.bullet.length <= 2 ? n.bullet : '\u2022'} ${n.text}` : n.text
        if (n.level) p.style.marginInlineStart = `${n.level * 1.5}em`
        return p
      }))
    },
    dispose() {
      listeners.abort()
      section.remove()
    },
  }
}
