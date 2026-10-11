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
export interface GridBorderSegment { x0: number; y0: number; x1: number; y1: number }
export interface GridSelectionGeometry { fills: readonly HighlightRect[]; borders: readonly GridBorderSegment[] }
/** S5 supplies bounded projected cells plus neighbor overscan. Compute the outline
 * before clipping so neither shared cell edges nor viewport cuts become borders.
 * pixelScale is required: device pixels per record coordinate. S5 must pass
 * sheet zoom × DPR for sheet-space records (DPR alone for already zoomed
 * CSS-space records); a missing or non-positive scale throws. */
export function gridSelectionOverlay(selection: { contains(at: { row: number; col: number }): boolean }, records: readonly GridOverlayRecord[], viewport: HighlightRect, pixelScale: number): GridSelectionGeometry {
  if (!Number.isFinite(pixelScale) || pixelScale <= 0) throw new RangeError('Invalid overlay pixel scale')
  const snap = (value: number) => Math.round(value * pixelScale)
  const fills: HighlightRect[] = [], borders: GridBorderSegment[] = [], seen = new Set<string>()
  const horizontal = new Map<number, Map<number, number>>(), vertical = new Map<number, Map<number, number>>()
  const edge = (map: Map<number, Map<number, number>>, at: number, lo: number, hi: number, sign: number) => {
    const events = map.get(at) ?? new Map<number, number>()
    events.set(lo, (events.get(lo) ?? 0) + sign)
    events.set(hi, (events.get(hi) ?? 0) - sign)
    map.set(at, events)
  }
  for (const record of records) {
    if (!selection.contains(record.source)) continue
    const { bounds } = record
    if (bounds.width <= 0 || bounds.height <= 0) continue
    // Key edges in integer device pixels, using the same endpoints for fills.
    // A non-empty cell narrower than a device pixel still paints one pixel
    // (the S5 pixel group); same-pixel neighbours then coalesce through `seen`.
    const x0 = snap(bounds.x), y0 = snap(bounds.y)
    const x1 = Math.max(x0 + 1, snap(bounds.x + bounds.width)), y1 = Math.max(y0 + 1, snap(bounds.y + bounds.height))
    const key = `${x0}:${y0}:${x1}:${y1}`
    if (seen.has(key)) continue
    seen.add(key)
    edge(horizontal, y0, x0, x1, -1)
    edge(horizontal, y1, x0, x1, 1)
    edge(vertical, x0, y0, y1, -1)
    edge(vertical, x1, y0, y1, 1)
    const x = Math.max(x0 / pixelScale, viewport.x), y = Math.max(y0 / pixelScale, viewport.y)
    const width = Math.min(x1 / pixelScale, viewport.x + viewport.width) - x
    const height = Math.min(y1 / pixelScale, viewport.y + viewport.height) - y
    if (width > 0 && height > 0) fills.push({ x, y, width, height })
  }
  const outline = (map: Map<number, Map<number, number>>, across: boolean) => {
    for (const [pixel, events] of map) {
      const at = pixel / pixelScale
      const near = across ? viewport.y : viewport.x, far = near + (across ? viewport.height : viewport.width)
      if (at < near || at > far) continue
      let active = 0, previous: number | undefined
      const spans: { lo: number; hi: number }[] = []
      for (const [position, delta] of [...events].sort((a, b) => a[0] - b[0])) {
        if (previous !== undefined && active && position > previous) {
          const last = spans[spans.length - 1]
          if (last && last.hi === previous) last.hi = position
          else spans.push({ lo: previous, hi: position })
        }
        active += delta
        previous = position
      }
      for (const span of spans) {
        const start = across ? viewport.x : viewport.y, end = start + (across ? viewport.width : viewport.height)
        const lo = Math.max(start, span.lo / pixelScale), hi = Math.min(end, span.hi / pixelScale)
        if (lo < hi) borders.push(across ? { x0: lo, y0: at, x1: hi, y1: at } : { x0: at, y0: lo, x1: at, y1: hi })
      }
    }
  }
  outline(horizontal, true)
  outline(vertical, false)
  return { fills, borders }
}
export function paintGridSelectionOverlay(ctx: CanvasRenderingContext2D, geometry: GridSelectionGeometry, scale = 1): void {
  if (!Number.isFinite(scale) || scale <= 0) throw new RangeError('Invalid overlay scale')
  ctx.save()
  ctx.fillStyle = 'rgba(0, 128, 0, 0.12)'
  ctx.strokeStyle = '#008000'
  ctx.lineWidth = 2 / scale
  for (const rect of geometry.fills) ctx.fillRect(rect.x, rect.y, rect.width, rect.height)
  if (geometry.borders.length) {
    ctx.beginPath()
    for (const edge of geometry.borders) { ctx.moveTo(edge.x0, edge.y0); ctx.lineTo(edge.x1, edge.y1) }
    ctx.stroke()
  }
  ctx.restore()
}
