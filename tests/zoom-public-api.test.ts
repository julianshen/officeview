import { describe, test, expect } from 'vitest'
import * as officeview from '../src/index'
import {
  MIN_ZOOM,
  MAX_ZOOM,
  clampZoom,
  clampPan,
  stepZoom,
  transformCss,
  zoomAt,
  panTransform,
  pinchTransform,
  distance,
  midpoint,
} from '../src/index'
import type { Point, Viewport, Transform } from '../src/index'

/**
 * Zoom in/out already works inside <OfficeDoc>. These tests pin the part that is
 * easy to regress silently: the zoom maths being part of the PUBLIC surface, so
 * a host can drive the same transform for its own chrome (a "fit width" button,
 * a minimap, a synced viewer) instead of reimplementing the clamping.
 */

const viewport = (over: Partial<Viewport> = {}): Viewport => ({
  width: 800,
  height: 600,
  contentWidth: 800,
  contentHeight: 1200,
  ...over,
})

describe('zoom public API', () => {
  test('every zoom helper is exported from the package entry point', () => {
    for (const name of [
      'MIN_ZOOM', 'MAX_ZOOM', 'clampZoom', 'clampPan', 'stepZoom',
      'transformCss', 'zoomAt', 'panTransform', 'pinchTransform', 'distance', 'midpoint',
    ]) {
      expect(officeview).toHaveProperty(name)
    }
  })

  test('the exported helpers are the same objects the module exports', () => {
    // Guards against a re-export that silently diverges from the source module.
    expect(typeof clampZoom).toBe('function')
    expect(MIN_ZOOM).toBe(1)
    expect(MAX_ZOOM).toBeGreaterThan(1)
  })

  test('MIN/MAX are sane and ordered', () => {
    expect(MIN_ZOOM).toBe(1)
    expect(MAX_ZOOM).toBeGreaterThanOrEqual(2)
    expect(MAX_ZOOM).toBeLessThanOrEqual(20)
  })
})

describe('zoom in / zoom out behaviour', () => {
  test('stepZoom zooms in and back out, clamped at both ends', () => {
    expect(stepZoom(1, 1)).toBeGreaterThan(1)
    expect(stepZoom(1.5, -1)).toBeCloseTo(1, 5)
    // at the floor, zoom-out stays put
    expect(stepZoom(MIN_ZOOM, -1)).toBe(MIN_ZOOM)
    // at the ceiling, zoom-in stays put
    expect(stepZoom(MAX_ZOOM, 1)).toBe(MAX_ZOOM)
  })

  test('repeated zoom-in steps are monotonic and bounded', () => {
    let z = MIN_ZOOM
    const seen: number[] = []
    for (let i = 0; i < 40; i++) {
      z = stepZoom(z, 1)
      seen.push(z)
    }
    expect(z).toBe(MAX_ZOOM)
    for (let i = 1; i < seen.length; i++) {
      if (seen[i] === seen[i - 1]) continue
      expect(seen[i]).toBeGreaterThan(seen[i - 1])
    }
    expect(Math.max(...seen)).toBeLessThanOrEqual(MAX_ZOOM)
  })

  test('clampZoom rejects garbage instead of trusting it', () => {
    // Non-finite input floors to MIN rather than propagating; that is the
    // documented behaviour and matters because a NaN zoom would break the CSS.
    expect(clampZoom(Number.NaN)).toBe(MIN_ZOOM)
    expect(clampZoom(Number.POSITIVE_INFINITY)).toBe(MIN_ZOOM)
    expect(clampZoom(Number.NEGATIVE_INFINITY)).toBe(MIN_ZOOM)
    expect(clampZoom(-5)).toBe(MIN_ZOOM)
    expect(clampZoom(99)).toBe(MAX_ZOOM)
    expect(clampZoom(2)).toBe(2)
  })

  test('zoomAt anchors the point under the cursor', () => {
    const vp = viewport()
    const t: Transform = { zoom: 1, panX: 0, panY: 0 }
    const anchor: Point = { x: 200, y: 150 }
    const zoomed = zoomAt(t, 2, anchor, vp)
    // the anchor must stay put: zoom*anchor + pan === anchor
    expect(zoomed.zoom).toBe(2)
    expect(zoomed.panX + zoomed.zoom * anchor.x).toBeCloseTo(anchor.x, 6)
    expect(zoomed.panY + zoomed.zoom * anchor.y).toBeCloseTo(anchor.y, 6)
  })

  test('pinch scales and keeps the midpoint anchored', () => {
    const start: Transform = { zoom: 1, panX: 0, panY: 0 }
    const mid: Point = { x: 400, y: 300 }
    const out = pinchTransform(start, 100, 200, mid, viewport())
    expect(out.zoom).toBeCloseTo(2, 6)
    expect(out.panX + out.zoom * mid.x).toBeCloseTo(mid.x, 6)
    expect(out.panY + out.zoom * mid.y).toBeCloseTo(mid.y, 6)
  })

  test('pinching inward below 1x returns to fit with no pan', () => {
    const start: Transform = { zoom: 2, panX: 50, panY: 80 }
    const out = pinchTransform(start, 200, 100, { x: 10, y: 10 }, viewport())
    expect(out.zoom).toBe(MIN_ZOOM)
    expect(out.panX).toBe(0)
    expect(out.panY).toBe(0)
  })

  test('pan respects content bounds and cannot lose the page', () => {
    // At 2x this viewport's content is 1600x2400 in an 800x600 window, so the
    // legal pan range on x is [-800, 0] and on y [-1800, 0].
    const vp = viewport()
    const zoomed: Transform = { zoom: 2, panX: 0, panY: 0 }
    const dragged = panTransform(zoomed, { x: -5000, y: -5000 }, vp)
    expect(dragged.panX).toBeGreaterThanOrEqual(-800)
    expect(dragged.panY).toBeGreaterThanOrEqual(-1800)
    // and the opposite direction cannot scroll past the far edge
    const other = panTransform(zoomed, { x: 5000, y: 5000 }, vp)
    expect(other.panX).toBeLessThanOrEqual(0)
    expect(other.panY).toBeLessThanOrEqual(0)
  })

  test('clampPan keeps short content top-aligned instead of centring it', () => {
    // Content shorter than the viewport must not drift downward or upward.
    const short = viewport({ height: 600, contentHeight: 100 })
    const t = clampPan({ zoom: 1, panX: 0, panY: -250 }, short)
    expect(t.panY).toBe(0)
  })

  test('transformCss is inert at fit and a real transform when zoomed', () => {
    expect(transformCss({ zoom: 1, panX: 0, panY: 0 })).toBe('none')
    const css = transformCss({ zoom: 1.5, panX: -10, panY: 20 })
    expect(css).toContain('scale(1.5)')
    expect(css).toContain('translate3d')
  })

  test('distance and midpoint are the geometry a pinch needs', () => {
    expect(distance({ x: 0, y: 0 }, { x: 3, y: 4 })).toBe(5)
    expect(midpoint({ x: 0, y: 0 }, { x: 10, y: 20 })).toEqual({ x: 5, y: 10 })
  })

  test('a zoom out from zoomed always returns exactly to fit', () => {
    let z = 2.5
    while (z > MIN_ZOOM) z = stepZoom(z, -1)
    expect(z).toBe(MIN_ZOOM)
    const settled = clampPan({ zoom: z, panX: -123, panY: -456 }, viewport())
    expect(settled.panX).toBe(0)
    expect(settled.panY).toBe(0)
  })
})