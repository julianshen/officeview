import { describe, test, expect } from 'vitest'
import {
  clampPan,
  clampZoom,
  MAX_ZOOM,
  MIN_ZOOM,
  panTransform,
  pinchTransform,
  stepZoom,
  transformCss,
  zoomAt,
  type Viewport,
} from '../src/core/zoom'

const viewport: Viewport = {
  width: 400,
  height: 800,
  contentWidth: 360, // fit width
  contentHeight: 3000, // several pages tall
}

const fit = { zoom: 1, panX: 0, panY: 0 }

describe('zoom math', () => {
  test('clampZoom bounds and rejects garbage', () => {
    expect(clampZoom(0.2)).toBe(MIN_ZOOM)
    expect(clampZoom(99)).toBe(MAX_ZOOM)
    expect(clampZoom(Number.NaN)).toBe(MIN_ZOOM)
  })

  test('clampPan keeps short content top-aligned instead of drifting', () => {
    // content 100px tall at 1.2x = 120px in an 800px viewport: no vertical
    // room to pan, so it stays at the top
    const t = clampPan({ zoom: 1.2, panX: -500, panY: -500 }, { ...viewport, contentHeight: 100 })
    expect(t.panY).toBe(0)
  })

  test('clampPan prevents losing the content at zoom 1', () => {
    expect(clampPan({ zoom: 1, panX: 200, panY: -200 }, viewport)).toEqual(fit)
  })

  test('pinch scales and keeps the gesture point anchored', () => {
    // fingers start 100px apart, end 200px apart => 2x
    const start = { zoom: 1, panX: 0, panY: 0 }
    const next = pinchTransform(start, 100, 200, { x: 200, y: 400 }, viewport)
    expect(next.zoom).toBeCloseTo(2, 5)
    // content point under (200,400) stays there
    const beforeX = (200 - start.panX) / start.zoom
    const afterX = (200 - next.panX) / next.zoom
    expect(afterX).toBeCloseTo(beforeX, 5)
  })

  test('pinch in below 1x clamps to fit and resets pan', () => {
    const next = pinchTransform({ zoom: 2, panX: -100, panY: -50 }, 200, 50, { x: 100, y: 200 }, viewport)
    expect(next.zoom).toBe(1)
    expect(next.panX).toBe(0)
    expect(next.panY).toBe(0)
  })

  test('pan respects content bounds when zoomed', () => {
    // at zoom 2 the content is 720x6000 in a 400x800 viewport:
    // panX range [-320, 0], panY range [-5200, 0]
    const t = panTransform({ zoom: 2, panX: 0, panY: 0 }, { x: 9999, y: 9999 }, viewport)
    expect(t.panX).toBe(0)
    expect(t.panY).toBe(0)
    const t2 = panTransform({ zoom: 2, panX: 0, panY: 0 }, { x: -9999, y: -9999 }, viewport)
    expect(t2.panX).toBe(-320)
    expect(t2.panY).toBe(-5200)
  })

  test('zoomAt anchors on the cursor', () => {
    const next = zoomAt(fit, 2, { x: 100, y: 200 }, viewport)
    // (100-0)/1 == (100-panX)/2
    expect((100 - next.panX) / next.zoom).toBeCloseTo(100, 5)
    expect(next.zoom).toBe(2)
  })

  test('stepZoom multiplies/divides and clamps', () => {
    expect(stepZoom(1, 1)).toBe(1.5)
    expect(stepZoom(1, -1)).toBe(1) // clamps at fit
    let z = 1
    for (let i = 0; i < 20; i++) z = stepZoom(z, 1)
    expect(z).toBe(MAX_ZOOM)
  })

  test('transformCss emits none at fit and a 3d transform when zoomed', () => {
    expect(transformCss(fit)).toBe('none')
    expect(transformCss({ zoom: 1.5, panX: 10, panY: -20 })).toBe('translate3d(10px, -20px, 0) scale(1.5)')
  })
})
