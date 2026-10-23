/**
 * The speaker notes pane (slideshow spec 5.2): the notes slide's body as
 * text (React text children, never deck-supplied HTML), one block per
 * paragraph with its bullet and level, A+ and A− size controls, scrolling.
 *
 * Keys typed or scrolled in the notes stay there, except Esc (spec 4.3): the
 * session's presenter key rule recognises `[data-officeview-presenter-notes]`.
 * With allowCopy=false the notes are not selectable (spec 6.3).
 */
import { useState, type ReactElement } from 'react'
import type { NoteParagraph } from '../types'

/** Notes text sizes (CSS px) the A+ and A− controls step through. */
export const NOTES_SIZES: readonly number[] = [12, 14, 16, 18, 20, 24, 28, 32, 40, 48]
/** 18 px. */
export const NOTES_DEFAULT = 3

/** A bullet character as authored; an auto-number scheme name reads as a plain bullet. */
const bulletText = (n: NoteParagraph) => (n.bullet ? `${n.bullet.length <= 2 ? n.bullet : '•'} ${n.text}` : n.text)

export interface NotesPaneProps {
  notes: readonly NoteParagraph[]
  /** False with allowCopy=false. */
  selectable: boolean
}

export function NotesPane({ notes, selectable }: NotesPaneProps): ReactElement {
  const [size, setSize] = useState(NOTES_DEFAULT)
  return (
    <div data-officeview-presenter-notes-pane="">
      <div data-officeview-presenter-notes-tools="" role="group" aria-label="Notes text size">
        <button type="button" aria-label="Smaller notes text" disabled={size <= 0} onClick={() => setSize((s) => Math.max(0, s - 1))}>A−</button>
        <button type="button" aria-label="Larger notes text" disabled={size >= NOTES_SIZES.length - 1} onClick={() => setSize((s) => Math.min(NOTES_SIZES.length - 1, s + 1))}>A+</button>
      </div>
      <div
        data-officeview-presenter-notes=""
        role="region"
        aria-label="Speaker notes"
        tabIndex={0}
        style={{ fontSize: `${NOTES_SIZES[size]}px`, ...(selectable ? {} : { userSelect: 'none', WebkitUserSelect: 'none' }) }}
      >
        {notes.length
          ? notes.map((n, i) => <p key={i} style={n.level ? { marginInlineStart: `${n.level * 1.5}em` } : undefined}>{bulletText(n)}</p>)
          : <p data-officeview-presenter-notes-empty="">No notes for this slide</p>}
      </div>
    </div>
  )
}
