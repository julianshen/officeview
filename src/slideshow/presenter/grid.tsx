/**
 * "See all slides" (slideshow spec 5.2 controls): every slide of the deck,
 * hidden ones included and marked, as buttons that jump to the slide (the
 * host-API `goTo`, so a hidden slide can be shown, spec 4.3).
 *
 * Thumbnails are leased presenter surfaces (S2-4 `grid(i)`), capped at 48:
 * only the scrolled-to window holds canvases, and a thumbnail scrolled out
 * of it is handed back with `releaseGrid` so its slot opens. The window is
 * computed from the scroller's own geometry on `scroll` (no
 * IntersectionObserver: the popup realm and jsdom both may lack it).
 */
import { useEffect, useLayoutEffect, useRef, useState, type ReactElement } from 'react'

/** One grid cell in CSS px: a 160 × 90 thumbnail with its label and gaps. */
export const GRID_CELL = { w: 176, h: 128 }
/** The S2-4 surfaces' thumbnail cap. */
export const GRID_WINDOW_MAX = 48

/**
 * The deck indices `[start, end)` that hold thumbnails for a scroller box:
 * the visible rows plus one row either side, never more than 48. A box with
 * no layout (not yet laid out, or jsdom) holds the first 48.
 */
export function gridWindow(box: { top: number; width: number; height: number }, count: number): [number, number] {
  if (!(box.height > 0)) return [0, Math.min(count, GRID_WINDOW_MAX)]
  const cols = Math.max(1, Math.floor(box.width / GRID_CELL.w))
  const start = Math.max(0, Math.floor(box.top / GRID_CELL.h) - 1) * cols
  const rows = Math.ceil(box.height / GRID_CELL.h) + 2
  return [Math.min(start, count), Math.min(count, start + Math.min(GRID_WINDOW_MAX, rows * cols))]
}

export interface SlideGridProps {
  count: number
  /** The deck index on screen (`aria-current`), or -1 at the end of the show. */
  current: number
  hidden(i: number): boolean
  /** The leased thumbnail canvas for deck index `i`. */
  thumb(i: number): HTMLCanvasElement
  /** Hands the thumbnail back. */
  release(i: number): boolean
  paint(i: number, canvas: HTMLCanvasElement): void
  onJump(i: number): void
  /** The grid covers the controls, so it carries its own close button. */
  onClose?(): void
}

const label = (i: number, hidden: boolean) => `Slide ${i + 1}${hidden ? ' (hidden)' : ''}`

export function SlideGrid(p: SlideGridProps): ReactElement {
  const scroller = useRef<HTMLDivElement>(null)
  const [range, setRange] = useState<[number, number]>(() => gridWindow({ top: 0, width: 0, height: 0 }, p.count))
  const held = useRef(new Map<number, HTMLCanvasElement>())
  const slots = useRef(new Map<number, HTMLElement>())
  const props = useRef(p)
  props.current = p

  const measure = () => {
    const el = scroller.current
    if (!el) return
    const next = gridWindow({ top: el.scrollTop, width: el.clientWidth, height: el.clientHeight }, props.current.count)
    setRange((r) => (r[0] === next[0] && r[1] === next[1] ? r : next))
  }

  // Lease the window, release what left it: never more than 48 held at once.
  useLayoutEffect(() => {
    const [start, end] = range
    const map = held.current
    for (const [i, c] of map) {
      if (i >= start && i < end) continue
      c.remove()
      props.current.release(i)
      map.delete(i)
    }
    for (let i = start; i < end; i++) {
      if (map.has(i)) continue
      const c = props.current.thumb(i)
      map.set(i, c)
      c.setAttribute('aria-hidden', 'true')
      slots.current.get(i)?.append(c)
      if (c.width > 0) props.current.paint(i, c)
    }
  }, [range])

  // Every thumbnail goes back with the grid.
  useEffect(() => () => {
    for (const [i, c] of held.current) { c.remove(); props.current.release(i) }
    held.current.clear()
  }, [])

  // Focus the current slide when the grid opens (spec 9: focus management inside Presenter View);
  // the window follows the scroller (a listener on the scroller's own node: realm-safe).
  useLayoutEffect(() => {
    const box = scroller.current
    measure()
    const el = box?.querySelector<HTMLElement>('[aria-current="true"]') ?? box?.querySelector<HTMLElement>('button')
    el?.focus()
    box?.addEventListener('scroll', measure, { passive: true })
    return () => box?.removeEventListener('scroll', measure)
  }, [])

  const items: ReactElement[] = []
  for (let i = 0; i < p.count; i++) {
    const hidden = p.hidden(i)
    items.push(
      <li key={i}>
        <button
          type="button"
          data-slide={i}
          data-hidden={hidden ? '' : undefined}
          aria-label={label(i, hidden)}
          aria-current={i === p.current ? 'true' : undefined}
          onClick={() => p.onJump(i)}
        >
          <span
            data-officeview-presenter-thumb=""
            ref={(el) => {
              if (el) {
                slots.current.set(i, el)
                const c = held.current.get(i)
                if (c && c.parentNode !== el) el.append(c)
              } else slots.current.delete(i)
            }}
          />
          <span aria-hidden="true">{label(i, hidden)}</span>
        </button>
      </li>,
    )
  }
  return (
    <div data-officeview-presenter-grid="" ref={scroller} role="region" aria-label="All slides">
      {p.onClose && <button type="button" data-officeview-presenter-grid-close="" aria-label="Close all slides" onClick={p.onClose}>Close</button>}
      <ul>{items}</ul>
    </div>
  )
}
