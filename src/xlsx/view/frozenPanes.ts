/**
 * Frozen-pane viewport math (pure; no DOM). The scroll window keeps ONE
 * logical origin; leading frozen rows/columns are painted with a zero origin
 * on their axis, so the scrolling body at viewport x ≥ frozen width shows
 * logical `origin.x + x/zoom` exactly as an unfrozen pane does. At origin 0
 * the body therefore starts right after the frozen columns (no duplicate),
 * and the logical range [frozen, origin + frozen) is what scrolled away
 * underneath the pinned strip, as in Excel.
 *
 * Regions: body (scrolls on both axes), top strip (frozen rows, scrolls on x),
 * left strip (frozen columns, scrolls on y) and corner (fixed). Each region is
 * a viewport clip rectangle plus the effective logical origin for its pixels,
 * so tiles, overlays and hit tests all use `(logical - origin) * zoom`.
 */
import type { AxisMetrics } from './contracts'

/** Requested freeze counts and the effective logical extents (0: axis not frozen at this size). */
export interface FrozenExtent { rows: number; cols: number; left: number; top: number }
export interface PanePoint { x: number; y: number }
export interface PaneRect { x: number; y: number; width: number; height: number }
export interface PaneRegion { kind: 'body' | 'top' | 'left' | 'corner'; clip: PaneRect; origin: PanePoint }

/** A frozen strip wider/taller than this share of the pane is not pinned at
 * that zoom/size (the body would have no room to scroll). */
export const FROZEN_VIEWPORT_FRACTION = 0.6

export function frozenExtent(request: { rows: number; cols: number } | undefined, metrics: { rows: AxisMetrics; cols: AxisMetrics }, viewport: { width: number; height: number }, zoom: number): FrozenExtent | undefined {
  if (!request || (!(request.rows > 0) && !(request.cols > 0))) return undefined
  const extent = (count: number, axis: AxisMetrics, size: number) => {
    if (!(count > 0) || !Number.isSafeInteger(count) || count >= axis.logicalCount) return 0
    const end = axis.start(count)
    return Number.isFinite(end) && end > 0 && end * zoom <= size * FROZEN_VIEWPORT_FRACTION ? end : 0
  }
  return Object.freeze({ rows: request.rows, cols: request.cols, left: extent(request.cols, metrics.cols, viewport.width), top: extent(request.rows, metrics.rows, viewport.height) })
}

/** Viewport CSS size of the pinned strips. */
export function frozenViewportSize(frozen: FrozenExtent | undefined, zoom: number): PanePoint {
  return { x: (frozen?.left ?? 0) * zoom, y: (frozen?.top ?? 0) * zoom }
}

/** Effective logical origin of the region containing a viewport point. */
export function originAtViewport(frozen: FrozenExtent | undefined, origin: PanePoint, zoom: number, at: PanePoint): PanePoint {
  const size = frozenViewportSize(frozen, zoom)
  return { x: at.x < size.x ? 0 : origin.x, y: at.y < size.y ? 0 : origin.y }
}

/** Effective logical origin for a logical point (frozen rows/columns use 0). */
export function originAtLogical(frozen: FrozenExtent | undefined, origin: PanePoint, at: PanePoint): PanePoint {
  return { x: frozen && at.x < frozen.left ? 0 : origin.x, y: frozen && at.y < frozen.top ? 0 : origin.y }
}

/** Visible logical intervals per axis: the frozen strip (when pinned) then the body. */
export function visibleSpans(frozen: FrozenExtent | undefined, origin: PanePoint, zoom: number, viewport: { width: number; height: number }): { x: Array<[number, number]>; y: Array<[number, number]> } {
  const axis = (pinned: number, start: number, size: number): Array<[number, number]> => {
    const end = start + size / zoom
    if (!(pinned > 0)) return [[start, end]]
    const body: [number, number] = [start + pinned, end]
    return body[1] > body[0] ? [[0, pinned], body] : [[0, pinned]]
  }
  return { x: axis(frozen?.left ?? 0, origin.x, viewport.width), y: axis(frozen?.top ?? 0, origin.y, viewport.height) }
}

/** Paint regions, body first (frozen strips are drawn over it). Without a pinned axis: one body region. */
export function paneRegions(frozen: FrozenExtent | undefined, origin: PanePoint, zoom: number, viewport: { width: number; height: number }): PaneRegion[] {
  const size = frozenViewportSize(frozen, zoom)
  const fx = Math.min(viewport.width, size.x), fy = Math.min(viewport.height, size.y)
  const regions: PaneRegion[] = [{ kind: 'body', clip: { x: fx, y: fy, width: viewport.width - fx, height: viewport.height - fy }, origin: { x: origin.x, y: origin.y } }]
  if (fy > 0) regions.push({ kind: 'top', clip: { x: fx, y: 0, width: viewport.width - fx, height: fy }, origin: { x: origin.x, y: 0 } })
  if (fx > 0) regions.push({ kind: 'left', clip: { x: 0, y: fy, width: fx, height: viewport.height - fy }, origin: { x: 0, y: origin.y } })
  if (fx > 0 && fy > 0) regions.push({ kind: 'corner', clip: { x: 0, y: 0, width: fx, height: fy }, origin: { x: 0, y: 0 } })
  return regions.filter(region => region.clip.width > 0 && region.clip.height > 0)
}

/** Scroll-into-view on one axis with a pinned extent: a cell inside the frozen
 * strip never scrolls; any other cell must clear the strip. Returns the new origin. */
export function ensureAxisVisible(origin: number, start: number, length: number, spanCss: number, pinned: number): number {
  if (pinned > 0 && start + length <= pinned) return origin
  const span = spanCss - pinned
  if (!(span > 0)) return origin
  const bodyOrigin = origin + pinned
  let next = bodyOrigin
  if (length > span) next = Math.max(start, Math.min(bodyOrigin, start + length - span))
  else if (start < bodyOrigin) next = start
  else if (start + length > bodyOrigin + span) next = start + length - span
  return Math.max(0, next - pinned)
}

/** Origin that puts a target at the top-left of the scrolling body (Go To);
 * a target inside the frozen strip keeps the current origin on that axis. */
export function bodyOriginFor(origin: PanePoint, frozen: FrozenExtent | undefined, target: PanePoint): PanePoint {
  return {
    x: frozen && frozen.left > 0 ? (target.x < frozen.left ? origin.x : Math.max(0, target.x - frozen.left)) : target.x,
    y: frozen && frozen.top > 0 ? (target.y < frozen.top ? origin.y : Math.max(0, target.y - frozen.top)) : target.y,
  }
}

/** Initial origin from the authored top-left cell of the scrolling pane. */
export function initialOrigin(frozen: { topRow?: number; leftCol?: number } | undefined, extent: FrozenExtent | undefined, metrics: { rows: AxisMetrics; cols: AxisMetrics }): PanePoint {
  const at = (index: number | undefined, axis: AxisMetrics, pinned: number) => index !== undefined && index < axis.logicalCount ? Math.max(0, axis.start(index) - pinned) : 0
  return { x: at(frozen?.leftCol, metrics.cols, extent?.left ?? 0), y: at(frozen?.topRow, metrics.rows, extent?.top ?? 0) }
}
