/**
 * Overlay geometry shared by the page renderer and the search/selection layers.
 * Rectangles are in page (document) coordinates.
 */

export interface HighlightRect {
  x: number
  y: number
  width: number
  height: number
}

/** Everything drawn on top of a page's own content. */
export interface PageOverlay {
  /** Search matches on this page. */
  highlights: HighlightRect[]
  /** Rectangles of the match currently being navigated to. */
  activeMatch: HighlightRect[]
  /** The user's text selection. */
  selection: HighlightRect[]
  /**
   * Changes whenever any of the above change; the page repaints when it does.
   */
  key: string
}

export function emptyOverlay(): PageOverlay {
  return { highlights: [], activeMatch: [], selection: [], key: '' }
}