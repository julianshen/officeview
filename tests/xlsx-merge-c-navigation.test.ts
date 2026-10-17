import { describe, expect, it } from 'vitest'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createScrollWindow, logicalToViewport, navigateKey, navigateTo, type NavigationContext } from '../src/xlsx/view/scrollWindow'

const visibility = visibilityFromHidden(createIntervalSet([]))
const view: NavigationContext['view'] = { visibility, baseHidden: createIntervalSet([]), filter: null, selection: { ranges: [], active: null, anchor: null } }
const viewport = { width: 640, height: 352, inset: { right: 56, bottom: 128 } }
const frozen = { rows: 1, cols: 1 }
function metrics(variable = false) {
  return createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, ...(variable ? { overrides: [{ start: 0, end: 0, sizeCss: 30 }, { start: 1, end: 1, sizeCss: 25 }, { start: 2, end: 2, sizeCss: 35 }] } : {}) }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order: createIdentityOrder(), visibility })
}

describe('independent C + cosmetics navigation seam', () => {
  it('Go To a pinned corner keeps both existing body offsets', () => {
    const m = metrics(), start = createScrollWindow(m, viewport, { offset: { x: 300, y: 400 }, frozen })
    const result = navigateTo(start, { metrics: m, viewport, view }, { addressSpace: 'source', at: { row: 0, col: 0 } })
    expect(result.kind).toBe('visible')
    expect(result.window.logicalOrigin).toEqual({ x: 300, y: 400 })
    expect(logicalToViewport(result.window, { x: 0, y: 0 })).toEqual({ x: 0, y: 0 })
  })

  it('ArrowUp brings a row hidden under the strip to its first body pixel', () => {
    const m = metrics(), start = createScrollWindow(m, viewport, { offset: { x: 0, y: 400 }, frozen })
    const result = navigateKey(start, { metrics: m, viewport, view }, 'ArrowUp', { row: 21, col: 3 })
    expect(result.target).toEqual({ row: 20, col: 3 })
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 380 })
    expect(logicalToViewport(result.window, { x: 192, y: 400 })).toEqual({ x: 192, y: 20 })
  })

  it('ArrowDown aligns the leading BODY row after making a below-fold target visible', () => {
    const m = metrics(), start = createScrollWindow(m, viewport, { frozen })
    const result = navigateKey(start, { metrics: m, viewport, view }, 'ArrowDown', { row: 23, col: 3 })
    // Target row24: 480..500. Raw minimal origin148 cuts the first body row at168.
    // Origin160 starts the body at180 (row9), keeping target320..340 in352px.
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 160 })
    expect(logicalToViewport(result.window, { x: 192, y: 480 })).toEqual({ x: 192, y: 320 })
    expect(logicalToViewport(result.window, { x: 192, y: 180 }).y).toBe(20)
  })

  it('aligns variable-height rows at the body boundary, not at logicalOrigin alone', () => {
    const m = metrics(true), vp = { width: 640, height: 172 }, start = createScrollWindow(m, vp, { frozen: { rows: 2, cols: 1 } })
    expect(m.rows.start(2)).toBe(55)
    expect(m.rows.start(8)).toBe(190)
    expect(m.rows.start(12)).toBe(270)
    const result = navigateKey(start, { metrics: m, viewport: vp, view }, 'ArrowDown', { row: 11, col: 3 })
    // Target270..290; minimum origin118 leaves body173 cutting row7 (170..190).
    // Origin135 puts row8 at body y55 and target at135..155 within172.
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 135 })
    expect(logicalToViewport(result.window, { x: 192, y: 190 }).y).toBe(55)
    expect(logicalToViewport(result.window, { x: 192, y: 270 }).y).toBe(135)
  })

  it('Ctrl+End with frozen strips clears the desktop zoom pill at the worksheet corner', () => {
    const m = metrics(), start = createScrollWindow(m, viewport, { frozen })
    const result = navigateKey(start, { metrics: m, viewport, view }, 'End', { row: 0, col: 0 }, { ctrl: true })
    expect(result.target).toEqual({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    // Whole leading rows: origin20971180; horizontal pill clearance aligns at1048000.
    expect(result.window.logicalOrigin).toEqual({ x: 1048000, y: 20971180 })
    const at = logicalToViewport(result.window, { x: 1048512, y: 20971500 })
    expect(at).toEqual({ x: 512, y: 320 })
    expect(at.x + 64).toBeLessThanOrEqual(640 - 56)
    expect(at.y + 20).toBeLessThanOrEqual(352)
    expect(at.x).toBeGreaterThanOrEqual(64)
    expect(at.y).toBeGreaterThanOrEqual(20)
  })

  it('frozen Go To XFD1048576 takes the body top-left, then clears the desktop zoom pill', () => {
    const m = metrics(), start = createScrollWindow(m, viewport, { frozen })
    const result = navigateTo(start, { metrics: m, viewport, view }, { addressSpace: 'source', at: { row: MAX_ROWS - 1, col: MAX_COLS - 1 } })
    expect(result.kind).toBe('visible')
    // Same corner as Ctrl+End: the body-origin window alone would leave the
    // target under the 56x128 pill, so the pinned branch still runs ensureVisible.
    expect(result.window.logicalOrigin).toEqual({ x: 1048000, y: 20971180 })
    const at = logicalToViewport(result.window, { x: 1048512, y: 20971500 })
    expect(at).toEqual({ x: 512, y: 320 })
    expect(at.x + 64).toBeLessThanOrEqual(640 - 56)
  })

  it('a pinned-column target keeps an inset-extended origin past the uninset maximum', () => {
    const m = metrics()
    // The inset lets the window scroll past cols.total - width (1047936).
    const start = createScrollWindow(m, viewport, { frozen, offset: { x: 1048000, y: 0 } })
    expect(start.logicalOrigin.x).toBe(1048000)
    const result = navigateKey(start, { metrics: m, viewport, view }, 'Home', { row: 5, col: MAX_COLS - 1 })
    expect(result.target).toEqual({ row: 5, col: 0 })
    expect(result.window.logicalOrigin).toEqual({ x: 1048000, y: 0 })
  })

  it('without frozen panes a visible Go To target preserves row1 and origin0', () => {
    const m = metrics(), start = createScrollWindow(m, viewport)
    const result = navigateTo(start, { metrics: m, viewport, view }, { addressSpace: 'source', at: { row: 1, col: 0 } })
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 0 })
    expect(logicalToViewport(result.window, { x: 0, y: 20 })).toEqual({ x: 0, y: 20 })
  })

  it('without frozen panes below-fold Go To retains the parent whole-row result', () => {
    const m = metrics(), start = createScrollWindow(m, viewport)
    const result = navigateTo(start, { metrics: m, viewport, view }, { addressSpace: 'source', at: { row: 24, col: 0 } })
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 160 })
    expect(logicalToViewport(result.window, { x: 0, y: 480 })).toEqual({ x: 0, y: 320 })
  })

  it('without frozen panes variable-height ArrowDown keeps the existing leading-row alignment', () => {
    const m = metrics(true), vp = { width: 640, height: 172 }, start = createScrollWindow(m, vp)
    const result = navigateKey(start, { metrics: m, viewport: vp, view }, 'ArrowDown', { row: 11, col: 3 })
    // Target270..290; raw118 cuts110..130, so the whole first row starts at130.
    expect(result.window.logicalOrigin).toEqual({ x: 0, y: 130 })
    expect(logicalToViewport(result.window, { x: 192, y: 270 }).y).toBe(140)
  })
})
