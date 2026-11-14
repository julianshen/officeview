/**
 * The speaker notes pane (slideshow spec 5.2): the notes slide's body as
 * text (React text children, never deck-supplied HTML), one block per
 * paragraph with its bullet and level, A+ and A− size controls, scrolling.
 *
 * Keys typed or scrolled in the notes stay there, except Esc (spec 4.3): the
 * session's presenter key rule recognises `[data-officeview-presenter-notes]`.
 * With allowCopy=false the notes are not selectable (spec 6.3).
 *
 * A streamed deck's notes arrive as a server-drawn image (`ImageNotes`): a
 * canvas in the same region, never selectable or copyable, with a status
 * line instead while it loads or after it failed.
 */
import { useLayoutEffect, useRef, useState, type ReactElement, type SyntheticEvent } from 'react'
import type { NoteParagraph } from '../types'

/** Notes text sizes (CSS px) the A+ and A− controls step through. */
export const NOTES_SIZES: readonly number[] = [12, 14, 16, 18, 20, 24, 28, 32, 40, 48]
/** 18 px. */
export const NOTES_DEFAULT = 3

/** A bullet character as authored; an auto-number scheme name reads as a plain bullet. */
const bulletText = (n: NoteParagraph) => (n.bullet ? `${n.bullet.length <= 2 ? n.bullet : '•'} ${n.text}` : n.text)

/**
 * Notes as a server-drawn image (a streamed deck: no notes text on the client). `paint` sizes and draws the
 * canvas it is handed, and is called only in state `'ready'`, once per notes object: return the same object
 * while the notes are unchanged.
 */
export interface ImageNotes {
  readonly kind: 'image'
  readonly state: 'loading' | 'ready' | 'error'
  paint(canvas: HTMLCanvasElement): void
}

export interface NotesPaneProps {
  notes: readonly NoteParagraph[] | ImageNotes
  /** False with allowCopy=false. */
  selectable: boolean
}

const refuse = (e: SyntheticEvent) => e.preventDefault()
/** The server draws notes 720 CSS px wide at any scale: show the backing at that
 * size (never its device pixels), shrinking to a narrower pane. */
const IMAGE_SIZE = { display: 'block', width: '100%', maxWidth: '720px', height: 'auto' } as const

/** The image variant: never selectable, no copy menu or drag, whatever the copy policy says. */
function ImageNotesView({ notes }: { notes: ImageNotes }): ReactElement {
  const canvas = useRef<HTMLCanvasElement>(null)
  const ready = notes.state === 'ready'
  useLayoutEffect(() => {
    if (ready && canvas.current) notes.paint(canvas.current)
  }, [notes, ready])
  return (
    <div
      data-officeview-presenter-notes=""
      role="region"
      aria-label="Speaker notes"
      tabIndex={0}
      style={{ userSelect: 'none', WebkitUserSelect: 'none' }}
      onCopy={refuse}
      onCut={refuse}
    >
      {ready
        ? <canvas ref={canvas} role="img" aria-label="Speaker notes (image)" draggable={false} style={IMAGE_SIZE} onContextMenu={refuse} onDragStart={refuse} />
        : <p data-officeview-presenter-notes-status="">{notes.state === 'error' ? 'Notes could not be loaded' : 'Loading notes…'}</p>}
    </div>
  )
}

export function NotesPane({ notes, selectable }: NotesPaneProps): ReactElement {
  const [size, setSize] = useState(NOTES_DEFAULT)
  // A+ and A− size text; an image has none.
  if ('kind' in notes) return <div data-officeview-presenter-notes-pane=""><ImageNotesView notes={notes} /></div>
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
