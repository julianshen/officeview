import { describe, expect, it } from 'vitest'
import { MIN_ZOOM, MAX_ZOOM, clampZoom } from '../src/core/zoom'
import { SHEET_MIN_ZOOM, SHEET_MAX_ZOOM } from '../src/xlsx/view/contracts'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow, refreshScrollWindow, scrollOffset, viewportToLogical } from '../src/xlsx/view/scrollWindow'
import { clampSheetZoom, createSheetZoomState, reduceSheetZoom, sheetZoomFor, zoomSheetAt, keyboardSheetZoom, wheelSheetZoom, pinchSheetZoom, SHEET_ZOOM_STEPS, snapSheetZoom } from '../src/xlsx/view/sheetZoom'

const viewport = { width: 375, height: 667 }
const metrics = createSheetMetrics({
  rows: createAxisMetrics({ logicalCount: 1_048_576, defaultSizeCss: 20 }),
  cols: createAxisMetrics({ logicalCount: 16_384, defaultSizeCss: 64 }),
})

describe('absolute sheet zoom and document ownership (AC-VIEW-04)', () => {
  it('starts each sheet at absolute 100%, independent of authored width or viewport', () => {
    const state = createSheetZoomState('doc-a', 'wide')
    expect(sheetZoomFor(state)).toBe(1)
    expect(sheetZoomFor(state, 'another')).toBe(1)
    expect(state.zoomBySheet.size).toBe(0)
  })

  it.each([
    [-1, 0.25], [0.25, 0.25], [0.5, 0.5], [1.37, 1.37], [4, 4], [6, 4],
    [NaN, 0.25], [Infinity, 0.25], [-Infinity, 0.25],
  ])('clamps %s with explicit sheet bounds to %s', (input, expected) => {
    expect(clampSheetZoom(input)).toBe(expected)
  })

  it('keeps the page-document constants and default clamp policy unchanged', () => {
    expect([SHEET_MIN_ZOOM, SHEET_MAX_ZOOM]).toEqual([0.25, 4])
    expect([MIN_ZOOM, MAX_ZOOM]).toEqual([1, 6])
    expect(clampZoom(0.25)).toBe(1)
    expect(clampZoom(8)).toBe(6)
  })

  it('retains independent sheet zoom through switches without mutating old state', () => {
    const initial = createSheetZoomState('doc-a', 'a')
    const a = reduceSheetZoom(initial, { kind: 'setZoom', docId: 'doc-a', sheetId: 'a', zoom: 0.5 })
    const b = reduceSheetZoom(a, { kind: 'switchSheet', docId: 'doc-a', sheetId: 'b' })
    expect(sheetZoomFor(b)).toBe(1)
    const zoomedB = reduceSheetZoom(b, { kind: 'setZoom', docId: 'doc-a', sheetId: 'b', zoom: 3 })
    const back = reduceSheetZoom(zoomedB, { kind: 'switchSheet', docId: 'doc-a', sheetId: 'a' })
    expect(sheetZoomFor(back)).toBe(0.5)
    expect(sheetZoomFor(back, 'b')).toBe(3)
    expect(sheetZoomFor(initial)).toBe(1)
    expect(initial.zoomBySheet.size).toBe(0)
  })

  it.each(['sort', 'filter', 'values', 'font', 'revision'] as const)('retains zoom on %s publication', reason => {
    const state = reduceSheetZoom(createSheetZoomState('doc-a', 'a'), { kind: 'setZoom', docId: 'doc-a', sheetId: 'a', zoom: 1.37 })
    const published = reduceSheetZoom(state, { kind: 'sheetChanged', docId: 'doc-a', sheetId: 'a', reason })
    expect(sheetZoomFor(published)).toBe(1.37)
    expect(published).toBe(state)
  })

  it('clamps stored zoom and resets every sheet only on explicit document replacement', () => {
    const a = reduceSheetZoom(createSheetZoomState('doc-a', 'a'), { kind: 'setZoom', docId: 'doc-a', sheetId: 'a', zoom: 7 })
    const state = reduceSheetZoom(a, { kind: 'setZoom', docId: 'doc-a', sheetId: 'b', zoom: 0.5 })
    expect(sheetZoomFor(state)).toBe(4)
    const replacement = reduceSheetZoom(state, { kind: 'replaceDocument', docId: 'doc-a', nextDocId: 'doc-b', activeSheetId: 'a' })
    expect(replacement.docId).toBe('doc-b')
    expect(replacement.activeSheetId).toBe('a')
    expect(sheetZoomFor(replacement)).toBe(1)
    expect(sheetZoomFor(replacement, 'b')).toBe(1)
    expect(replacement.zoomBySheet.size).toBe(0)
    expect(sheetZoomFor(state)).toBe(4)
  })

  it('ignores stale document actions including late replacement', () => {
    const state = reduceSheetZoom(createSheetZoomState('doc-b', 'a'), { kind: 'setZoom', docId: 'doc-b', sheetId: 'a', zoom: 2 })
    for (const action of [
      { kind: 'setZoom', docId: 'doc-a', sheetId: 'a', zoom: 0.25 },
      { kind: 'switchSheet', docId: 'doc-a', sheetId: 'b' },
      { kind: 'sheetChanged', docId: 'doc-a', sheetId: 'a', reason: 'values' },
      { kind: 'replaceDocument', docId: 'doc-a', nextDocId: 'doc-c', activeSheetId: 'c' },
    ] as const) expect(reduceSheetZoom(state, action)).toBe(state)
    expect(sheetZoomFor(state)).toBe(2)
  })
})

describe('sheet gestures through ScrollWindow (AC-VIEW-04)', () => {
  it('leaves a 1000px authored sheet at 100% with horizontal scroll', () => {
    const wide = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 20 }),
      cols: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 100 }),
    })
    const window = createScrollWindow(wide, viewport)
    expect(window.zoom).toBe(1)
    expect(window.physicalExtent.width).toBe(1000)
    expect(window.physicalExtent.width).toBeGreaterThan(viewport.width)
  })

  it('keeps absolute 100% across refresh and viewport resize in a 375px container', () => {
    const narrow = { width: 375, height: 667 }
    const wide = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 20 }),
      cols: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 100 }),
    })
    const window = createScrollWindow(wide, narrow)
    expect(window.zoom).toBe(1)
    const refreshed = refreshScrollWindow(window, wide, narrow)
    expect(refreshed.zoom).toBe(1)
    expect(refreshed.physicalExtent.width).toBe(1000)
    const resized = refreshScrollWindow(window, wide, { width: 200, height: 400 })
    expect(resized.zoom).toBe(1)
    expect(resized.physicalExtent.width).toBe(1000)
    expect(resized.physicalExtent.width).toBeGreaterThan(200)
  })

  it.each([0.25, 0.5, 1.37, 4])('changes to absolute %s preserving the pointer at a segmented offset', zoom => {
    const anchor = { x: 111.5, y: 232.25 }
    const start = createScrollWindow(metrics, viewport, { zoom: 1, deviceScale: 3, offset: { x: 700_000, y: 19_000_000 } })
    const before = viewportToLogical(start, anchor)
    expect(before).toEqual({ x: 700_111.5, y: 19_000_232.25 })
    const next = zoomSheetAt(start, metrics, viewport, zoom, anchor)
    expect(next.zoom).toBe(zoom)
    expect(next.deviceScale).toBe(3)
    expect(viewportToLogical(next, anchor).x).toBeCloseTo(before.x, 6)
    expect(viewportToLogical(next, anchor).y).toBeCloseTo(before.y, 6)
    expect(next.physicalExtent.width).toBeLessThanOrEqual(1_000_000)
    expect(next.physicalExtent.height).toBeLessThanOrEqual(1_000_000)
  })

  it('uses the actual clamped zoom when preserving an anchor', () => {
    const anchor = { x: 40, y: 60 }
    const start = createScrollWindow(metrics, viewport, { offset: { x: 700_000, y: 19_000_000 } })
    const before = viewportToLogical(start, anchor)
    const next = zoomSheetAt(start, metrics, viewport, 100, anchor)
    expect(next.zoom).toBe(4)
    expect(viewportToLogical(next, anchor)).toEqual(before)
  })

  it('steps keyboard zoom around the viewport center and permits zooming below 100%', () => {
    const center = { x: viewport.width / 2, y: viewport.height / 2 }
    const start = createScrollWindow(metrics, viewport, { offset: { x: 1000, y: 2000 } })
    const before = viewportToLogical(start, center)
    expect(before).toEqual({ x: 1187.5, y: 2333.5 })
    const next = keyboardSheetZoom(start, metrics, viewport, -1)
    expect(next.zoom).toBe(0.75)
    expect(viewportToLogical(next, center).x).toBeCloseTo(before.x, 10)
    expect(viewportToLogical(next, center).y).toBeCloseTo(before.y, 10)
    const up = keyboardSheetZoom(next, metrics, viewport, 1, center)
    expect(up.zoom).toBe(1)
  })

  it('snaps keyboard zoom to the documented step table with exact inverse round trips', () => {
    for (const step of SHEET_ZOOM_STEPS) {
      const window = createScrollWindow(metrics, viewport, { zoom: step })
      if (step > SHEET_MIN_ZOOM) {
        const down = keyboardSheetZoom(window, metrics, viewport, -1)
        expect(down.zoom).toBeLessThan(step)
        expect(keyboardSheetZoom(down, metrics, viewport, 1).zoom).toBe(step)
      } else {
        expect(keyboardSheetZoom(window, metrics, viewport, -1).zoom).toBe(step)
      }
      if (step < SHEET_MAX_ZOOM) {
        const up = keyboardSheetZoom(window, metrics, viewport, 1)
        expect(up.zoom).toBeGreaterThan(step)
        expect(keyboardSheetZoom(up, metrics, viewport, -1).zoom).toBe(step)
      } else {
        expect(keyboardSheetZoom(window, metrics, viewport, 1).zoom).toBe(step)
      }
    }
    const offTable = createScrollWindow(metrics, viewport, { zoom: 1.37 })
    expect(keyboardSheetZoom(offTable, metrics, viewport, 1).zoom).toBe(1.5)
    expect(keyboardSheetZoom(offTable, metrics, viewport, -1).zoom).toBe(1.25)
  })

  it.each([{ ctrlKey: true }, { metaKey: true }])('zooms modified wheel about its pointer (%j)', modifier => {
    const anchor = { x: 120, y: 250 }
    const start = createScrollWindow(metrics, viewport, { offset: { x: 1000, y: 2000 } })
    const before = viewportToLogical(start, anchor)
    const next = wheelSheetZoom(start, metrics, viewport, { deltaY: -200, ...modifier }, anchor)
    expect(next.zoom).toBeCloseTo(Math.exp(0.4), 12)
    expect(viewportToLogical(next, anchor).x).toBeCloseTo(before.x, 10)
    expect(viewportToLogical(next, anchor).y).toBeCloseTo(before.y, 10)
    const extreme = wheelSheetZoom(start, metrics, viewport, { deltaY: -Number.MAX_VALUE, ...modifier }, anchor)
    expect(extreme.zoom).toBe(4)
  })

  it('ignores unmodified or nonfinite wheel input', () => {
    const start = createScrollWindow(metrics, viewport)
    const anchor = { x: 50, y: 50 }
    expect(wheelSheetZoom(start, metrics, viewport, { deltaY: -100 }, anchor)).toBe(start)
    expect(wheelSheetZoom(start, metrics, viewport, { deltaY: NaN, ctrlKey: true }, anchor)).toBe(start)
    expect(wheelSheetZoom(start, metrics, viewport, { deltaY: Infinity, metaKey: true }, anchor)).toBe(start)
  })

  it.each([0.25, 4])('saturates extreme wheel deltas at the exact bounds from %s', zoom => {
    const start = createScrollWindow(metrics, viewport, { zoom })
    const anchor = { x: 50, y: 50 }
    expect(wheelSheetZoom(start, metrics, viewport, { deltaY: -Number.MAX_VALUE, ctrlKey: true }, anchor).zoom).toBe(4)
    expect(wheelSheetZoom(start, metrics, viewport, { deltaY: Number.MAX_VALUE, ctrlKey: true }, anchor).zoom).toBe(0.25)
  })

  it('pinches from gesture-start distance and carries the logical anchor to a moved midpoint', () => {
    const first = { x: 100, y: 200 }
    const current = { x: 140, y: 230 }
    const start = createScrollWindow(metrics, viewport, { zoom: 0.5, offset: { x: 700_000, y: 19_000_000 } })
    const before = viewportToLogical(start, first)
    expect(before).toEqual({ x: 700_200, y: 19_000_400 })
    const next = pinchSheetZoom(start, metrics, viewport, 100, 300, first, current)
    expect(next.zoom).toBe(1.5)
    expect(viewportToLogical(next, current).x).toBeCloseTo(before.x, 6)
    expect(viewportToLogical(next, current).y).toBeCloseTo(before.y, 6)
    const clamped = pinchSheetZoom(start, metrics, viewport, 100, 10_000, first, current)
    expect(clamped.zoom).toBe(4)
    expect(viewportToLogical(clamped, current).x).toBeCloseTo(before.x, 6)
    expect(viewportToLogical(clamped, current).y).toBeCloseTo(before.y, 6)
  })

  it.each([[0, 100], [-1, 100], [100, 0], [100, -1], [NaN, 100], [100, Infinity]])('ignores invalid pinch distances %s/%s', (first, current) => {
    const start = createScrollWindow(metrics, viewport)
    expect(pinchSheetZoom(start, metrics, viewport, first, current, { x: 20, y: 30 }, { x: 30, y: 40 })).toBe(start)
  })

  it('clamps at content edges when a requested anchor cannot be retained', () => {
    const start = createScrollWindow(metrics, viewport)
    const next = zoomSheetAt(start, metrics, viewport, 0.25, { x: 100, y: 200 })
    expect(next.zoom).toBe(0.25)
    expect(scrollOffset(next)).toEqual({ x: 0, y: 0 })
  })

  it.each([NaN, Infinity, -Infinity])('rejects non-finite zoom %s instead of silently clamping', (zoom) => {
    const start = createScrollWindow(metrics, viewport)
    expect(() => zoomSheetAt(start, metrics, viewport, zoom, { x: 10, y: 10 })).toThrow(RangeError)
    expect(start.zoom).toBe(1)
  })
})

describe('snapSheetZoom hardening', () => {
  it('rejects non-finite zoom instead of snapping to a bound', () => {
    for (const bad of [NaN, Infinity, -Infinity]) {
      expect(() => snapSheetZoom(bad, 1)).toThrow(RangeError)
      expect(() => snapSheetZoom(bad, -1)).toThrow(RangeError)
    }
  })

  it('keeps the step table immutable at runtime', () => {
    expect(Object.isFrozen(SHEET_ZOOM_STEPS)).toBe(true)
    expect(() => (SHEET_ZOOM_STEPS as number[]).push(5)).toThrow(TypeError)
  })
})
