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

export interface GridOverlayRecord { source: { row: number; col: number }; bounds: HighlightRect }
/** Viewport-record seam: S5 supplies already projected records, never full-range geometry. */
export function gridSelectionOverlay(selection: { contains(at: { row: number; col: number }): boolean }, records: readonly GridOverlayRecord[], viewport: HighlightRect): HighlightRect[] {
  const out: HighlightRect[] = []
  for (const record of records) {
    if (!selection.contains(record.source)) continue
    const { bounds } = record, x = Math.max(bounds.x, viewport.x), y = Math.max(bounds.y, viewport.y)
    const width = Math.min(bounds.x + bounds.width, viewport.x + viewport.width) - x
    const height = Math.min(bounds.y + bounds.height, viewport.y + viewport.height) - y
    if (width > 0 && height > 0) out.push({ x, y, width, height })
  }
  return out
}
export function paintGridSelectionOverlay(ctx: CanvasRenderingContext2D, rects: readonly HighlightRect[], scale = 1): void {
  if (!Number.isFinite(scale) || scale <= 0) throw new RangeError('Invalid overlay scale')
  ctx.save()
  ctx.fillStyle = 'rgba(0, 128, 0, 0.12)'
  ctx.strokeStyle = '#008000'
  ctx.lineWidth = 2 / scale
  for (const rect of rects) { ctx.fillRect(rect.x, rect.y, rect.width, rect.height); ctx.strokeRect(rect.x, rect.y, rect.width, rect.height) }
  ctx.restore()
}
