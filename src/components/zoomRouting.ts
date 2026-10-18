/**
 * Zoom-input routing for the paged viewer (spec §6.9, §6.11). Every zoom
 * input (buttons, Ctrl/Cmd-wheel, pinch, keyboard, double-click, fit) goes
 * through these pure functions, which move between native scroll and pan
 * atomically (scrollPan) and apply the one fit-exit rule. A small lazy module:
 * <OfficeDoc> requests it on its first render, and the gestures it serves
 * need painted pages, so it is in place before they can happen.
 */
import { clampPan, clampZoom, type Point, type Transform, type Viewport } from '../core/zoom'
import type { PageSpec } from '../render/paint'
import { PAGE_PADDING_PX as PAD } from './paged/metrics'
import { contentYAt, enterZoom, exitZoom, pageOffsets, remapContentY } from './scrollPan'

export interface ZoomContext {
  transform: Transform
  fit: boolean
  /** The container's scrollTop (or one about to be applied). */
  scrollTop: number
  specs: readonly PageSpec[]
  naturalWidth: number
  fitWidth: number
  gap: number
  viewport: Viewport
}
/** What OfficeDoc commits: transform, fit flag and scrollTop in one go. */
export interface ZoomOutcome { transform: Transform; fit: boolean; scrollTop: number }

const IDENTITY: Transform = { zoom: 1, panX: 0, panY: 0 }

/** The current view as a pan on the natural layout: scrollTop folded in; in
 * fit mode the zoom is z0 = fit / natural and the content point under `at`
 * is remapped by page index plus page-local offset (R3-4). */
export function viewAt(c: ZoomContext, at: Point): Transform {
  const t = c.transform
  if (!c.fit) return { ...t, panY: t.panY - c.scrollTop }
  const z0 = c.fitWidth / c.naturalWidth
  const y = remapContentY(contentYAt(at.y, { scrollTop: c.scrollTop, panY: 0, zoom: 1, padding: PAD }), pageOffsets(c.specs, c.fitWidth, c.gap), pageOffsets(c.specs, c.naturalWidth, c.gap))
  return { zoom: z0, panX: 0, panY: at.y - PAD - y * z0 }
}

/** Zoom `view` to `next` about `at` on the natural layout; fit mode ends.
 * Above 1 is a pan with scrollTop 0 (enterZoom); at 1 the native scroll
 * position that keeps the anchored point (exitZoom). null: nothing changes. */
export function zoomTo(c: ZoomContext, view: Transform, next: number, at: Point): ZoomOutcome | null {
  const zoom = clampZoom(next)
  if (zoom === view.zoom && !c.fit) return null
  const o = pageOffsets(c.specs, c.naturalWidth, c.gap)
  const n = o.top.length
  const contentHeight = n > 0 ? o.top[n - 1] + o.height[n - 1] : 0
  const from = { scrollTop: 0, panY: view.panY, zoom: view.zoom, padding: PAD }
  if (zoom > 1) {
    const panX = enterZoom({ ...from, panY: view.panX }, zoom, at.x).panY
    const transform = clampPan({ zoom, panX, panY: enterZoom(from, zoom, at.y).panY }, { ...c.viewport, contentWidth: c.naturalWidth, contentHeight })
    return { transform, fit: false, scrollTop: 0 }
  }
  return { transform: IDENTITY, fit: false, scrollTop: exitZoom(from, at.y, contentHeight + 2 * PAD - c.viewport.height).scrollTop }
}

/** Fit on keeps the content at the viewport top; fit off is a zoom to 1. */
export function toggleFit(c: ZoomContext): ZoomOutcome | null {
  const top = { x: 0, y: 0 }
  const view = viewAt(c, top)
  if (c.fit) return zoomTo(c, view, 1, top)
  const y = contentYAt(0, { scrollTop: 0, panY: view.panY, zoom: view.zoom, padding: PAD })
  return { transform: IDENTITY, fit: true, scrollTop: remapContentY(y, pageOffsets(c.specs, c.naturalWidth, c.gap), pageOffsets(c.specs, c.fitWidth, c.gap)) + PAD }
}
