/**
 * Viewport zoom/pan math for the canvas document viewer. Pure functions so
 * the gesture behavior can be tested without a DOM.
 *
 * Coordinate model: the content (page stack) is laid out at "fit width"
 * scale. The viewport applies translate(pan) scale(zoom) about the origin.
 * `zoom` 1 means fit-to-width; 2 means twice that size.
 */

export interface Point {
  x: number
  y: number
}

export interface Viewport {
  /** Visible viewport size in CSS px. */
  width: number
  height: number
  /** Full content size in CSS px at fit scale. */
  contentWidth: number
  contentHeight: number
}

export interface Transform {
  zoom: number
  panX: number
  panY: number
}

export const MIN_ZOOM = 1
export const MAX_ZOOM = 6

export function clampZoom(zoom: number, min = MIN_ZOOM, max = MAX_ZOOM): number {
  if (!Number.isFinite(zoom)) return min
  return Math.min(max, Math.max(min, zoom))
}

export function distance(a: Point, b: Point): number {
  return Math.hypot(a.x - b.x, a.y - b.y)
}

export function midpoint(a: Point, b: Point): Point {
  return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }
}

/**
 * Keep the content inside (or at least covering) the viewport: when zoomed in,
 * panning past the content edge is clamped so you never lose the page. On an
 * axis where the content is shorter than the viewport it stays top-aligned.
 * At zoom 1 (fit) there is nothing to pan — the content fits the width.
 */
export function clampPan(transform: Transform, viewport: Viewport): Transform {
  const { zoom, panX, panY } = transform
  const z = clampZoom(zoom)
  if (z <= MIN_ZOOM) return { zoom: z, panX: 0, panY: 0 }
  const minX = Math.min(0, viewport.width - viewport.contentWidth * z)
  const maxX = Math.max(0, viewport.width - viewport.contentWidth * z)
  const minY = Math.min(0, viewport.height - viewport.contentHeight * z)
  const maxY = Math.max(0, viewport.height - viewport.contentHeight * z)
  return {
    zoom: z,
    panX: Math.min(maxX, Math.max(minX, panX)),
    panY: Math.min(maxY, Math.max(minY, panY)),
  }
}

/**
 * Pinch: scale by the ratio of finger distance, keeping the point under the
 * fingers anchored. `start` is the transform when the gesture began.
 */
export function pinchTransform(
  start: Transform,
  startDistance: number,
  currentDistance: number,
  currentCenter: Point,
  viewport: Viewport,
): Transform {
  const ratio = startDistance > 0 ? currentDistance / startDistance : 1
  const zoom = clampZoom(start.zoom * ratio)
  const actualRatio = zoom / start.zoom
  // anchor the gesture midpoint: content point under it must not move
  const panX = currentCenter.x - (currentCenter.x - start.panX) * actualRatio
  const panY = currentCenter.y - (currentCenter.y - start.panY) * actualRatio
  return clampPan({ zoom, panX, panY }, viewport)
}

/** Single-finger/mouse drag: pan by the pointer delta. */
export function panTransform(start: Transform, delta: Point, viewport: Viewport): Transform {
  return clampPan(
    { zoom: start.zoom, panX: start.panX + delta.x, panY: start.panY + delta.y },
    viewport,
  )
}

/**
 * Zoom toward a viewport point (wheel, double-tap, buttons): the content
 * point under `anchor` stays under it.
 */
export function zoomAt(
  start: Transform,
  nextZoom: number,
  anchor: Point,
  viewport: Viewport,
): Transform {
  const zoom = clampZoom(nextZoom)
  const ratio = zoom / start.zoom
  const panX = anchor.x - (anchor.x - start.panX) * ratio
  const panY = anchor.y - (anchor.y - start.panY) * ratio
  return clampPan({ zoom, panX, panY }, viewport)
}

/** Zoom step for buttons/keys. */
export function stepZoom(current: number, direction: 1 | -1): number {
  const next = direction > 0 ? current * 1.5 : current / 1.5
  return clampZoom(Number(next.toFixed(4)))
}

/** CSS transform string for a wrapper element. */
export function transformCss(t: Transform): string {
  if (t.zoom <= MIN_ZOOM && t.panX === 0 && t.panY === 0) return 'none'
  return `translate3d(${t.panX}px, ${t.panY}px, 0) scale(${t.zoom})`
}
