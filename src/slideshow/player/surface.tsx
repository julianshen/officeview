/**
 * The audience surface (slideshow spec 4.1 end-of-show, 4.3 black/white
 * screen; checkpoint 1: static slides and cut transitions).
 *
 * Plain DOM built through the host's own document (realm-safe: Stage 2 moves
 * surfaces between windows), updated synchronously, so the session's state
 * and the page never disagree. The host is a child of the viewer's theme
 * root, so the scoped `[data-ov-theme-root] [data-officeview-show-*]` rules
 * in officeview.css apply.
 *
 * The output canvas is the permanent reserve's surface: the terminal path
 * (spec 7, ladder step 3) paints into it directly. The engine
 * (`player/stage.ts`) mounts its compositor canvases right after it, under
 * the black/white screen and the end-of-show slide. The live region, the
 * visible-text list and the trigger/link buttons are `player/a11y.tsx`.
 */

export interface SurfaceView {
  /** The deck index shown, or the end-of-show slide. */
  slide: number | 'end'
  count: number
  screen: 'slide' | 'black' | 'white'
}

export interface AudienceSurface {
  readonly root: HTMLElement
  readonly canvas: HTMLCanvasElement
  show(view: SurfaceView): void
  dispose(): void
}

export const END_OF_SHOW = 'End of slide show, click to exit.'

/** Builds the surface inside `host` (appended; the host's other children, such as the full-screen notice, are left alone). */
export function createAudienceSurface(host: HTMLElement, backing: { w: number; h: number }, onClick: (e: MouseEvent) => void): AudienceSurface {
  const doc = host.ownerDocument
  const root = doc.createElement('div')
  root.setAttribute('data-officeview-show-surface', '')
  const canvas = doc.createElement('canvas')
  canvas.setAttribute('data-officeview-show-canvas', '')
  canvas.setAttribute('aria-hidden', 'true')
  canvas.width = backing.w
  canvas.height = backing.h
  const blank = doc.createElement('div')
  blank.hidden = true
  const end = doc.createElement('div')
  end.setAttribute('data-officeview-show-end', '')
  end.textContent = END_OF_SHOW
  end.hidden = true
  root.append(canvas, blank, end)
  root.addEventListener('click', onClick)
  host.append(root)
  return {
    root,
    canvas,
    show(view) {
      root.setAttribute('data-slide', String(view.slide))
      root.setAttribute('data-screen', view.screen)
      const isEnd = view.slide === 'end'
      canvas.hidden = isEnd
      end.hidden = !isEnd
      blank.hidden = view.screen === 'slide'
      if (view.screen === 'slide') blank.removeAttribute('data-officeview-show-blank')
      else blank.setAttribute('data-officeview-show-blank', view.screen)
    },
    dispose() {
      root.removeEventListener('click', onClick)
      root.remove()
    },
  }
}
