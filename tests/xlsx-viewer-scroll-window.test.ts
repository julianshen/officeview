import { describe, expect, it } from 'vitest'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createRowOrder, createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, createRowVisibility, projectVisibility, normalizeFilter } from '../src/xlsx/view/visibility'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
import { MAX_PHYSICAL_SCROLL_CSS_PX, type FilterSpec, type GridSelection, type ScrollWindow } from '../src/xlsx/view/contracts'
import { createScrollWindow, scrollOffset, setScrollOffset, scrollBy, panScrollWindow, scrollPhysicalTo, recenterScrollWindow, refreshScrollWindow, setScrollZoom, logicalToViewport, viewportToLogical, logicalToBitmap, bitmapToLogical, logicalToSegment, sourceToViewport, hitTestViewport, visibleLogicalViewport, navigateTo, navigateKey, navigatorRequestFromA1, type NavigationContext } from '../src/xlsx/view/scrollWindow'

const viewport = { width: 375, height: 600 }
function fullMetrics() { return createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) }) }
function bounded(window: ScrollWindow) {
  expect(window.physicalExtent.width).toBeLessThanOrEqual(MAX_PHYSICAL_SCROLL_CSS_PX)
  expect(window.physicalExtent.height).toBeLessThanOrEqual(MAX_PHYSICAL_SCROLL_CSS_PX)
  expect(window.physicalPosition.x).toBeGreaterThanOrEqual(0)
  expect(window.physicalPosition.y).toBeGreaterThanOrEqual(0)
  expect(window.physicalPosition.x).toBeLessThanOrEqual(Math.max(0, window.physicalExtent.width - viewport.width))
  expect(window.physicalPosition.y).toBeLessThanOrEqual(Math.max(0, window.physicalExtent.height - viewport.height))
}
function contextFor(count = MAX_ROWS, filter: FilterSpec | null = null, hidden = createIntervalSet([]), order = createIdentityOrder()): NavigationContext {
  const projected = projectVisibility(hidden, filter, order, row => row % 2 === 0, count)
  const selection: GridSelection = Object.freeze({ active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 }, ranges: [] })
  return { viewport, metrics: createSheetMetrics({ rows: createAxisMetrics({ logicalCount: count, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility: projected.visibility }), view: { baseHidden: hidden, filter: projected.filter, visibility: projected.visibility, selection } }
}

describe('bounded Float64 scroll segments (AC-CAP-02, AC-VIEW-03)', () => {
  it('caps both axes through zoom, DPR and default-size changes', () => {
    const metrics = fullMetrics()
    for (const zoom of [0.25, 0.5, 1, 1.37, 4]) for (const deviceScale of [1, 1.5, 2, 3]) bounded(createScrollWindow(metrics, viewport, { zoom, deviceScale }))
    const huge = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 1000 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 1000 }) })
    bounded(refreshScrollWindow(createScrollWindow(metrics, viewport, { zoom: 4 }), huge, viewport))
  })
  it('wheel/touch deltas preserve local distances without whole-sheet amplification', () => {
    const metrics = fullMetrics()
    const start = createScrollWindow(metrics, viewport, { zoom: 0.5, offset: { x: 5000, y: 1000000 } })
    const wheel = scrollBy(start, metrics, viewport, { x: 30, y: 120 })
    expect(scrollOffset(wheel)).toEqual({ x: 5060, y: 1000240 })
    expect(scrollOffset(panScrollWindow(wheel, metrics, viewport, { x: 10, y: -20 }))).toEqual({ x: 5040, y: 1000280 })
    bounded(wheel)
  })
  it('recenters guard zones without changing world coordinates, anchors or selection identities', () => {
    const metrics = fullMetrics(), context = contextFor()
    const start = createScrollWindow(metrics, viewport, { offset: { x: 600000, y: 10000000 }, zoom: 1.37, deviceScale: 3 })
    const nearGuard = { ...start, physicalPosition: { x: 1, y: start.physicalExtent.height - viewport.height - 1 } }
    const anchor = { x: 73, y: 91 }
    const before = viewportToLogical(nearGuard, anchor)
    const centered = recenterScrollWindow(nearGuard, metrics, viewport)
    expect(centered.physicalPosition).not.toEqual(nearGuard.physicalPosition)
    expect(scrollOffset(centered)).toEqual(scrollOffset(nearGuard))
    expect(viewportToLogical(centered, anchor)).toEqual(before)
    expect(centered.zoom).toBe(nearGuard.zoom)
    expect(centered.deviceScale).toBe(nearGuard.deviceScale)
    expect(context.view.selection.active).toEqual({ row: 0, col: 0 })
    const physical = scrollPhysicalTo(start, metrics, viewport, { x: start.physicalPosition.x + 20, y: start.physicalPosition.y + 40 })
    expect(physical.logicalOrigin.x).toBeCloseTo(start.logicalOrigin.x + 20 / start.zoom, 8)
    expect(physical.logicalOrigin.y).toBeCloseTo(start.logicalOrigin.y + 40 / start.zoom, 8)
  })
  it('sets Go To offsets directly and clamps Home/End to full axes despite an A1 used range', () => {
    const metrics = fullMetrics()
    const start = createScrollWindow(metrics, viewport)
    const jumped = setScrollOffset(start, metrics, viewport, { x: 1048512, y: 25165800 })
    expect(jumped.logicalOrigin.x).toBe(metrics.cols.totalCss() - viewport.width)
    expect(jumped.logicalOrigin.y).toBe(metrics.rows.totalCss() - viewport.height)
    expect(scrollOffset(setScrollOffset(jumped, metrics, viewport, { x: -5, y: -100 }))).toEqual({ x: 0, y: 0 })
    bounded(jumped)
  })
  it('round-trips exact first/last cells across fractional zoom, DPR and recentering', () => {
    const metrics = fullMetrics()
    for (const zoom of [0.25, 0.5, 1, 1.37, 4]) for (const deviceScale of [1, 1.5, 2, 3]) for (const cell of [{ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }]) {
      const box = metrics.sourceCellToLogical(cell)!
      let window = setScrollOffset(createScrollWindow(metrics, viewport, { zoom, deviceScale }), metrics, viewport, box)
      for (let pass = 0; pass < 2; pass++) {
        const rendered = sourceToViewport(window, metrics, cell)!
        expect(hitTestViewport(window, metrics, viewport, rendered)).toEqual(cell)
        const world = { x: box.x + box.width / 2, y: box.y + box.height / 2 }
        const point = logicalToViewport(window, world)
        const recovered = viewportToLogical(window, point)
        expect(recovered.x).toBeCloseTo(world.x, 8)
        expect(recovered.y).toBeCloseTo(world.y, 8)
        const fromBitmap = bitmapToLogical(window, logicalToBitmap(window, world))
        expect(fromBitmap.x).toBeCloseTo(world.x, 8)
        expect(fromBitmap.y).toBeCloseTo(world.y, 8)
        const local = logicalToSegment(window, world)
        expect(local.x * zoom - window.physicalPosition.x).toBeCloseTo(point.x, 7)
        bounded(window)
        window = recenterScrollWindow({ ...window, physicalPosition: { x: 0, y: 0 } }, metrics, viewport)
      }
    }
  })
  it('preserves zoom anchors in the interior and clamps impossible edge anchors legally', () => {
    const metrics = fullMetrics(), anchor = { x: 187.5, y: 300 }
    const start = createScrollWindow(metrics, viewport, { offset: { x: 10000, y: 10000000 }, zoom: 1.37 })
    const world = viewportToLogical(start, anchor)
    for (const zoom of [0.25, 0.5, 1, 4]) {
      const changed = setScrollZoom(start, metrics, viewport, zoom, anchor)
      expect(viewportToLogical(changed, anchor).x).toBeCloseTo(world.x, 8)
      expect(viewportToLogical(changed, anchor).y).toBeCloseTo(world.y, 8)
      bounded(changed)
    }
    const edge = setScrollZoom(createScrollWindow(metrics, viewport), metrics, viewport, 0.25, anchor)
    expect(edge.logicalOrigin).toEqual({ x: 0, y: 0 })
    expect(visibleLogicalViewport(edge, viewport)).toEqual({ x: 0, y: 0, width: 1500, height: 2400 })
  })
  it('uses summaries for collapsed axes and rejects malformed scroll inputs', () => {
    const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 0 }) })
    const window = createScrollWindow(metrics, viewport)
    expect(hitTestViewport(window, metrics, viewport, { x: 0, y: 0 })).toBeNull()
    expect(scrollOffset(scrollBy(window, metrics, viewport, { x: 100, y: 100 }))).toEqual({ x: 0, y: 0 })
    expect(Array.from(metrics.rows.visibleBetween(0, 100).runs)).toEqual([])
    expect(() => createScrollWindow(fullMetrics(), viewport, { deviceScale: 0 })).toThrow()
    expect(() => scrollBy(window, metrics, viewport, { x: NaN, y: 0 })).toThrow()
    expect(() => createScrollWindow(fullMetrics(), { width: Infinity, height: 600 })).toThrow()
  })

  it.each([NaN, Infinity, -Infinity])('rejects non-finite zoom %s at scroll entry points', (zoom) => {
    const metrics = fullMetrics()
    const start = createScrollWindow(metrics, viewport)
    expect(() => createScrollWindow(metrics, viewport, { zoom })).toThrow(RangeError)
    expect(() => setScrollZoom(start, metrics, viewport, zoom, { x: 10, y: 10 })).toThrow(RangeError)
  })
})

describe('display/source navigator and hidden inspector (AC-VIEW-04)', () => {
  it('maps displayed A2 to original A3 and Original A2 to displayed A3', () => {
    const order = createRowOrder([{ r0: 0, r1: 2, sourceAtOffset: new Uint32Array([0, 2, 1]) }])
    const context = contextFor(3, null, createIntervalSet([]), order)
    const window = createScrollWindow(context.metrics, viewport)
    const displayed = navigateTo(window, context, navigatorRequestFromA1('A2'))
    expect(displayed.target).toEqual({ row: 2, col: 0 })
    expect(displayed.display).toEqual({ rowSlot: 1, col: 0 })
    const original = navigateTo(window, context, navigatorRequestFromA1('A2', 'source'))
    expect(original.target).toEqual({ row: 1, col: 0 })
    expect(original.display).toEqual({ rowSlot: 2, col: 0 })
  })
  it('returns fresh frozen active/target cells without aliasing caller objects', () => {
    const order = createRowOrder([{ r0: 0, r1: 2, sourceAtOffset: new Uint32Array([0, 2, 1]) }])
    const context = contextFor(3, null, createIntervalSet([]), order)
    const window = createScrollWindow(context.metrics, viewport)
    const displayed = navigateTo(window, context, navigatorRequestFromA1('A2'))
    expect(displayed.kind).toBe('visible')
    if (displayed.kind === 'visible') {
      expect(displayed.active).toEqual(displayed.target)
      expect(displayed.active).not.toBe(displayed.target)
      expect(Object.isFrozen(displayed.active)).toBe(true)
      expect(Object.isFrozen(displayed.target)).toBe(true)
    }
  })
  it('inspects hidden exact targets without unhide or activation and retains the valid active cell', () => {
    const context = contextFor(8, null, createIntervalSet([{ r0: 2, r1: 2 }]))
    const window = createScrollWindow(context.metrics, viewport), active = { row: 0, col: 0 }
    const result = navigateTo(window, context, navigatorRequestFromA1('A3', 'source'), active)
    expect(result.kind).toBe('hidden')
    expect(result.activate).toBe(false)
    if (result.kind === 'hidden') expect(result.hiddenReason).toBe('authored-hidden')
    expect(result.active).toEqual(active)
    expect(result.active).not.toBe(active)
    expect(Object.isFrozen(result.active)).toBe(true)
    expect(result.window).toBe(window)
    expect(context.view.visibility.visible(2)).toBe(false)
    expect(context.view.selection.active).toEqual(active)
  })
  it('retains a nearest visible body active cell, prefers forward ties and never substitutes headers', () => {
    const filter: FilterSpec = { range: { r0: 0, r1: 5, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const context = contextFor(6, filter)
    const window = createScrollWindow(context.metrics, viewport)
    const result = navigateTo(window, context, navigatorRequestFromA1('A4', 'source'), { row: 3, col: 0 })
    expect(result.kind).toBe('hidden')
    if (result.kind === 'hidden') expect(result.hiddenReason).toBe('filter-hidden')
    expect(result.active).toEqual({ row: 4, col: 0 })
    const none = createRowVisibility(6, row => row === 0 || row === 5)
    const noBody = { ...context, metrics: createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 6, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), visibility: none }), view: { ...context.view, visibility: none } }
    expect(navigateTo(window, noBody, navigatorRequestFromA1('A4'), { row: 3, col: 0 }).active).toBeNull()
  })
  it('keeps the active row when only its column is hidden, moving to the nearest visible column', () => {
    const filter: FilterSpec = { range: { r0: 0, r1: 5, c0: 0, c1: 3, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const hidden = createIntervalSet([])
    const order = createIdentityOrder()
    const projected = projectVisibility(hidden, filter, order, row => row !== 3, 12)
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 12, defaultSizeCss: 24 }),
      cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 64, collapsed: [{ start: 2, end: 2 }] }),
      order, visibility: projected.visibility })
    const context: NavigationContext = { viewport, metrics, view: { baseHidden: hidden, filter: projected.filter, visibility: projected.visibility, selection: Object.freeze({ active: null, anchor: null, ranges: [] }) } }
    const window = createScrollWindow(metrics, viewport)
    const header = navigateTo(window, context, navigatorRequestFromA1('A4', 'source'), { row: 0, col: 2 })
    expect(header.kind).toBe('hidden')
    expect(header.active).toEqual({ row: 0, col: 3 })
    const outside = navigateKey(window, context, 'ArrowDown', { row: 10, col: 2 })
    expect(outside.active).toEqual({ row: 11, col: 3 })
  })
  it('relocates authored-hidden rows outside the filter body to the nearest visible row of any kind', () => {
    const filter: FilterSpec = { range: { r0: 0, r1: 5, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const hidden = createIntervalSet([{ r0: 10, r1: 10 }])
    const order = createIdentityOrder()
    const projected = projectVisibility(hidden, filter, order, () => true, 12)
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 12, defaultSizeCss: 24 }),
      cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 64 }), order, visibility: projected.visibility })
    const context: NavigationContext = { viewport, metrics, view: { baseHidden: hidden, filter: projected.filter, visibility: projected.visibility, selection: Object.freeze({ active: null, anchor: null, ranges: [] }) } }
    const window = createScrollWindow(metrics, viewport)
    const result = navigateTo(window, context, navigatorRequestFromA1('C11', 'source'), { row: 10, col: 2 })
    expect(result.kind).toBe('hidden')
    expect(result.active).toEqual({ row: 11, col: 2 })
  })
  it('never substitutes a protected header for a filter-hidden body row', () => {
    const filter: FilterSpec = { range: { r0: 0, r1: 5, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const hidden = createIntervalSet([])
    const order = createIdentityOrder()
    const projected = projectVisibility(hidden, filter, order, row => row === 4, 6)
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 6, defaultSizeCss: 24 }),
      cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 64 }), order, visibility: projected.visibility })
    const context: NavigationContext = { viewport, metrics, view: { baseHidden: hidden, filter: projected.filter, visibility: projected.visibility, selection: Object.freeze({ active: null, anchor: null, ranges: [] }) } }
    const window = createScrollWindow(metrics, viewport)
    const result = navigateTo(window, context, navigatorRequestFromA1('A3', 'source'), { row: 1, col: 0 })
    expect(result.kind).toBe('hidden')
    expect(result.active).toEqual({ row: 4, col: 0 })
  })
  it('relocates a stale hidden body row under a permutation without touching headers', () => {
    const order = createRowOrder([{ r0: 5, r1: 7, sourceAtOffset: new Uint32Array([6, 5, 7]) }])
    const filter: FilterSpec = { range: { r0: 5, r1: 7, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const hidden = createIntervalSet([])
    const projected = projectVisibility(hidden, filter, order, row => row === 7, 10)
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 24 }),
      cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 64 }), order, visibility: projected.visibility })
    const context: NavigationContext = { viewport, metrics, view: { baseHidden: hidden, filter: projected.filter, visibility: projected.visibility, selection: Object.freeze({ active: null, anchor: null, ranges: [] }) } }
    const window = createScrollWindow(metrics, viewport)
    const result = navigateTo(window, context, navigatorRequestFromA1('A7', 'source'), { row: 6, col: 0 })
    expect(result.kind).toBe('hidden')
    expect(result.active).toEqual({ row: 7, col: 0 })
  })
  it('uses rank/select and axis summaries for arrows, Home/End, Ctrl-arrows and all-hidden cases', () => {
    const context = contextFor(8, null, createIntervalSet([{ r0: 1, r1: 2 }]))
    const window = createScrollWindow(context.metrics, viewport)
    const current = { row: 0, col: 0 }
    expect(navigateKey(window, context, 'ArrowDown', current).active).toEqual({ row: 3, col: 0 })
    expect(navigateKey(window, context, 'ArrowDown', current, { ctrl: true }).active).toEqual({ row: 7, col: 0 })
    expect(navigateKey(window, context, 'End', current).active).toEqual({ row: 0, col: MAX_COLS - 1 })
    expect(navigateKey(window, context, 'Home', { row: 3, col: 100 }, { ctrl: true }).active).toEqual({ row: 0, col: 0 })
    const hidden = contextFor(8, null, createIntervalSet([{ r0: 0, r1: 7 }]))
    expect(navigateKey(window, hidden, 'ArrowDown', null).kind).toBe('empty')
    expect(navigateKey(window, hidden, 'End', current, { ctrl: true }).active).toBeNull()
  })
  it('navigates End to the full worksheet corner independent of populated range', () => {
    const context = contextFor(), window = createScrollWindow(context.metrics, viewport)
    const result = navigateKey(window, context, 'End', { row: 0, col: 0 }, { ctrl: true })
    expect(result.active).toEqual({ row: MAX_ROWS - 1, col: MAX_COLS - 1 })
    const box = sourceToViewport(result.window, context.metrics, result.active!)!
    expect(hitTestViewport(result.window, context.metrics, viewport, { x: box.x + 1, y: box.y + 1 })).toEqual(result.active)
  })
  it('moves pages by viewport distance, skipping collapsed columns and retaining boundary cells', () => {
    const context = contextFor()
    const window = createScrollWindow(context.metrics, viewport, { zoom: 0.5 })
    const result = navigateKey(window, context, 'PageDown', { row: 0, col: 0 })
    expect(result.active?.row).toBe(50)
    expect(result.window.logicalOrigin.y).toBe(1200)
    expect(navigateKey(result.window, context, 'PageUp', result.active).active?.row).toBe(0)
    const columns = { ...context, metrics: createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: 5, defaultSizeCss: 64, collapsed: [{ start: 1, end: 2 }] }) }) }
    expect(navigateKey(window, columns, 'ArrowRight', { row: 0, col: 0 }).active?.col).toBe(3)
    expect(navigateKey(window, columns, 'ArrowLeft', { row: 0, col: 3 }).active?.col).toBe(0)
    expect(navigateKey(window, columns, 'ArrowLeft', { row: 0, col: 0 }).active?.col).toBe(0)
  })
})

describe('guard-zone and giant-cell page navigation regressions', () => {
  it('retains native segment motion until a guard zone is reached', () => {
    const metrics = fullMetrics()
    const start = createScrollWindow(metrics, viewport, { offset: { x: 600000, y: 10000000 }, zoom: 1.37 })
    const next = { x: start.physicalPosition.x + 20, y: start.physicalPosition.y + 40 }
    const moved = scrollPhysicalTo(start, metrics, viewport, next)
    expect(moved.physicalPosition).toEqual(next)
    expect(moved.logicalOrigin.y).toBeCloseTo(start.logicalOrigin.y + 40 / 1.37, 8)
    const guard = scrollPhysicalTo(moved, metrics, viewport, { x: 1, y: 1 })
    expect(guard.physicalPosition.x).toBeGreaterThan(100000)
    expect(guard.physicalPosition.y).toBeGreaterThan(100000)
  })
  it('pages within a giant cell by viewport distance and activates a represented cell after crossing its end', () => {
    const context = contextFor(10)
    context.metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 10, defaultSizeCss: 1000 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) })
    const start = createScrollWindow(context.metrics, viewport, { offset: { x: 0, y: 400 } })
    const down = navigateKey(start, context, 'PageDown', { row: 0, col: 0 })
    expect(down.window.logicalOrigin.y).toBe(1000)
    expect(down.active).toEqual({ row: 1, col: 0 })
    const up = navigateKey(down.window, context, 'PageUp', down.active)
    expect(up.window.logicalOrigin.y).toBe(400)
    expect(up.active).toEqual({ row: 0, col: 0 })
  })
  it('keeps the page target inside the moved viewport when the active cell is far outside it', () => {
    const context = contextFor(1000)
    const window = createScrollWindow(context.metrics, viewport, { offset: { x: 0, y: 12000 } })
    const down = navigateKey(window, context, 'PageDown', { row: 900, col: 0 })
    expect(down.window.logicalOrigin.y).toBe(12600)
    const box = context.metrics.sourceCellToLogical(down.active!)!
    expect(box.y).toBeGreaterThanOrEqual(down.window.logicalOrigin.y)
    expect(box.y).toBeLessThan(down.window.logicalOrigin.y + viewport.height)
    const up = navigateKey(window, context, 'PageUp', { row: 100, col: 0 })
    expect(up.window.logicalOrigin.y).toBe(11400)
    const upBox = context.metrics.sourceCellToLogical(up.active!)!
    expect(upBox.y).toBeGreaterThanOrEqual(up.window.logicalOrigin.y)
    expect(upBox.y).toBeLessThan(up.window.logicalOrigin.y + viewport.height)
  })
})

it('recenters only the guarded axis and keeps ordinary virtual wheel motion local', () => {
  const metrics = fullMetrics()
  const start = createScrollWindow(metrics, viewport, { offset: { x: 600000, y: 10000000 }, zoom: 1.37 })
  const ordinary = scrollBy(start, metrics, viewport, { x: 20, y: 40 })
  expect(ordinary.physicalPosition).toEqual({ x: start.physicalPosition.x + 20, y: start.physicalPosition.y + 40 })
  const oneGuard = { ...ordinary, physicalPosition: { x: ordinary.physicalPosition.x, y: 1 } }
  const centered = recenterScrollWindow(oneGuard, metrics, viewport)
  expect(centered.physicalPosition.x).toBe(oneGuard.physicalPosition.x)
  expect(centered.physicalPosition.y).not.toBe(1)
})

it('round-trips sparse giant custom intervals near the end at every fractional zoom/DPR', () => {
  const metrics = createSheetMetrics({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0.1, overrides: [{ start: 500000, end: 500004, sizeCss: 1e8 }] }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 0.1, overrides: [{ start: 8000, end: 8002, sizeCss: 1e7 }] }),
  })
  for (const zoom of [0.25, 0.5, 1, 1.37, 4]) for (const deviceScale of [1, 1.5, 2, 3]) {
    for (const cell of [{ row: 0, col: 0 }, { row: 500000, col: 8000 }, { row: 500004, col: 8002 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }]) {
      const box = metrics.sourceCellToLogical(cell)!
      const initial = setScrollOffset(createScrollWindow(metrics, viewport, { zoom, deviceScale }), metrics, viewport, box)
      const centered = recenterScrollWindow({ ...initial, physicalPosition: { x: 0, y: 0 } }, metrics, viewport)
      expect(centered.logicalOrigin).toEqual(initial.logicalOrigin)
      for (const window of [initial, centered]) {
        const painted = sourceToViewport(window, metrics, cell)!
        expect(hitTestViewport(window, metrics, viewport, painted)).toEqual(cell)
        const world = { x: box.x + Math.min(box.width / 2, 1), y: box.y + Math.min(box.height / 2, 1) }
        const recovered = bitmapToLogical(window, logicalToBitmap(window, world))
        expect(Math.abs(recovered.x - world.x)).toBeLessThanOrEqual(Math.max(1e-12, Math.abs(world.x) * Number.EPSILON * 4))
        expect(Math.abs(recovered.y - world.y)).toBeLessThanOrEqual(Math.max(1e-12, Math.abs(world.y) * Number.EPSILON * 4))
        bounded(window)
      }
    }
  }
})

it('answers fully collapsed navigation/hits without enumerating rows, columns or masks', () => {
  const zero = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 0 }) })
  const unexpected = () => { throw new Error('A collapsed axis must be answered by its summary') }
  const metrics = { ...zero, rows: { ...zero.rows, start: unexpected, size: unexpected }, cols: { ...zero.cols, start: unexpected, size: unexpected } }
  const context = contextFor()
  context.metrics = metrics
  context.view = { ...context.view, visibility: { visible: unexpected, rank: unexpected, select: unexpected, next: unexpected } }
  const window = createScrollWindow(metrics, viewport)
  expect(hitTestViewport(window, metrics, viewport, { x: 10, y: 10 })).toBeNull()
  expect(navigateKey(window, context, 'End', null, { ctrl: true }).kind).toBe('empty')
})

describe('seeded scroll segment properties (cap, bounds, coverage)', () => {
  let seed = 20261010
  const random = (n: number): number => {
    seed = (seed * 1103515245 + 12345) % 2147483648
    return seed % n
  }
  const checkWindow = (window: ScrollWindow, view: { width: number; height: number }, totals: { width: number; height: number }): void => {
    for (const axis of [{ extent: window.physicalExtent.width, position: window.physicalPosition.x, viewSize: view.width, origin: window.logicalOrigin.x, total: totals.width },
      { extent: window.physicalExtent.height, position: window.physicalPosition.y, viewSize: view.height, origin: window.logicalOrigin.y, total: totals.height }]) {
      expect(axis.extent).toBeLessThanOrEqual(MAX_PHYSICAL_SCROLL_CSS_PX)
      expect(axis.position).toBeGreaterThanOrEqual(0)
      expect(axis.position).toBeLessThanOrEqual(Math.max(0, axis.extent - axis.viewSize))
      const base = axis.origin - axis.position / window.zoom
      expect(base).toBeGreaterThanOrEqual(-1e-6)
      expect(base).toBeLessThanOrEqual(axis.origin)
      expect(base + axis.extent / window.zoom).toBeGreaterThanOrEqual(axis.origin + axis.viewSize / window.zoom - 1e-6)
      expect(base + axis.extent / window.zoom).toBeLessThanOrEqual(axis.total + 1e-6)
    }
  }
  it('holds across random metrics, zoom, DPR, sizes, giant intervals and operation sequences', () => {
    const zooms = [0.25, 0.5, 1, 1.37, 2, 4]
    const scales = [1, 1.5, 2, 3]
    const sizes = [8, 24, 64, 200, 1000]
    for (let trial = 0; trial < 25; trial++) {
      const rowSize = sizes[random(sizes.length)]
      const colSize = sizes[random(sizes.length)]
      const giant = random(3) === 0
      const metrics = createSheetMetrics({
        rows: createAxisMetrics({ logicalCount: 500 + random(2000), defaultSizeCss: rowSize,
          overrides: giant ? [{ start: 100, end: 104, sizeCss: 1e6 }] : [] }),
        cols: createAxisMetrics({ logicalCount: 100 + random(500), defaultSizeCss: colSize,
          overrides: giant ? [{ start: 10, end: 12, sizeCss: 1e6 }] : [] }),
      })
      const view = { width: 200 + random(800), height: 200 + random(800) }
      const zoom = zooms[random(zooms.length)]
      const deviceScale = scales[random(scales.length)]
      let window = createScrollWindow(metrics, view, {
        zoom, deviceScale,
        offset: { x: random(2_000_000), y: random(2_000_000) },
      })
      checkWindow(window, view, { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() })
      for (let step = 0; step < 12; step++) {
        const action = random(5)
        if (action === 0) window = scrollBy(window, metrics, view, { x: random(4000) - 2000, y: random(4000) - 2000 })
        else if (action === 1) window = setScrollZoom(window, metrics, view, zooms[random(zooms.length)], { x: random(view.width), y: random(view.height) })
        else if (action === 2) window = recenterScrollWindow({ ...window, physicalPosition: { x: 0, y: 0 } }, metrics, view)
        else if (action === 3) window = refreshScrollWindow(window, metrics, view)
        else window = setScrollOffset(window, metrics, view, { x: random(2_000_000), y: random(2_000_000) })
        checkWindow(window, view, { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() })
      }
    }
  })
})

it('rejects scroll origins whose Float64 spacing loses device pixels instead of silently mis-scrolling', () => {
  const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e20 }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e20 }) })
  expect(() => createScrollWindow(metrics, viewport, { offset: { x: 1e18, y: 1e18 } })).toThrow(/precision|represent/i)
  const origin = createScrollWindow(metrics, viewport)
  expect(logicalToViewport(origin, viewportToLogical(origin, { x: 100, y: 100 }))).toEqual({ x: 100, y: 100 })
  expect(() => setScrollOffset(origin, metrics, viewport, { x: 1e20, y: 1e20 })).toThrow(/precision|represent/i)
  const fine = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e9 }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e9 }) })
  expect(() => createScrollWindow(fine, viewport, { offset: { x: 1e8, y: 1e8 }, deviceScale: 1e12 })).toThrow(/precision|represent/i)
})

it.each([{ origin: 2 ** 52 - 1, zoom: 1, deviceScale: 2 }, { origin: 2 ** 49 - 1, zoom: 4, deviceScale: 3 }])(
  'rejects a viewport crossing a Float64 spacing boundary: %j', ({ origin, zoom, deviceScale }) => {
    const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e20 }), cols: createAxisMetrics({ logicalCount: 1, defaultSizeCss: 1e20 }) })
    expect(() => createScrollWindow(metrics, viewport, { zoom, deviceScale, offset: { x: origin, y: origin } })).toThrow(/precision|represent/i)
  },
)

describe('review: source body membership under arbitrary display order', () => {
  function filtered(values: number[], r0: number, range: FilterSpec['range'], match: (row: number) => boolean, count = 12): NavigationContext {
    const order = createRowOrder([{ r0, r1: r0 + values.length - 1, sourceAtOffset: new Uint32Array(values) }])
    const filter: FilterSpec = { range, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const baseHidden = createIntervalSet([])
    const applied = projectVisibility(baseHidden, filter, order, match, count)
    return { viewport, metrics: createSheetMetrics({ rows: createAxisMetrics({ logicalCount: count, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility: applied.visibility }),
      view: { baseHidden, filter: applied.filter, visibility: applied.visibility, selection: { active: null, anchor: null, ranges: [] } } }
  }
  it('reproduces the relocated protected-header counterexample and checks source membership', () => {
    const context = filtered([6, 5, 7], 5, { r0: 5, r1: 7, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, row => row !== 6)
    const window = createScrollWindow(context.metrics, viewport)
    const result = navigateTo(window, context, navigatorRequestFromA1('A7', 'source'), { row: 6, col: 0 })
    expect(result.kind).toBe('hidden')
    expect(result.active).toEqual({ row: 7, col: 0 })
    expect(result.activate).toBe(false)
  })
  it.each(['Home', 'End', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Enter', 'PageUp', 'PageDown'] as const)(
    'seeds %s from null with a visible source body record, not a moved header/totals row', key => {
      const context = filtered([6, 5, 8, 7], 5, { r0: 5, r1: 8, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, row => row !== 6)
      const window = createScrollWindow(context.metrics, viewport)
      const result = navigateKey(window, context, key, null)
      expect(result.active?.row).toBe(7)
      expect(context.metrics.sourceCellToLogical(result.active!)).not.toBeNull()
    },
  )
  it.each(['Home', 'End', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'] as const)(
    'keeps Ctrl+%s null initialization inside the visible source body', key => {
      const context = filtered([6, 5, 8, 7], 5, { r0: 5, r1: 8, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, row => row !== 6)
      expect(navigateKey(createScrollWindow(context.metrics, viewport), context, key, null, { ctrl: true }).active?.row).toBe(7)
    },
  )
  it('prefers the forward display candidate on ties while excluding a moved protected header and totals', () => {
    const context = filtered([7, 5, 8, 9, 6], 5, { r0: 5, r1: 9, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, row => row !== 8)
    const result = navigateTo(createScrollWindow(context.metrics, viewport), context, navigatorRequestFromA1('A9', 'source'), { row: 8, col: 0 })
    expect(result.active).toEqual({ row: 6, col: 0 }) // slot 9; slot 5 is equally far backward; slots 6/8 are protected.
  })
  it('returns null active without substituting protected or outside rows when no body candidate remains', () => {
    const context = filtered([6, 5, 8, 7], 5, { r0: 5, r1: 8, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, () => false)
    const window = createScrollWindow(context.metrics, viewport)
    expect(navigateTo(window, context, navigatorRequestFromA1('A7', 'source'), { row: 6, col: 0 }).active).toBeNull()
    for (const key of ['Home', 'End', 'ArrowUp', 'ArrowDown', 'PageUp', 'PageDown'] as const) {
      const result = navigateKey(window, context, key, null)
      expect(result.kind).toBe('empty')
      expect(result.active).toBeNull()
      expect(result.activate).toBe(false)
    }
  })
  it('matches a source-membership oracle across all five-row permutations and body masks', () => {
    function* permutations(rows: number[]): Generator<number[]> {
      if (!rows.length) { yield []; return }
      for (let i = 0; i < rows.length; i++) {
        for (const tail of permutations(rows.filter((_, at) => at !== i))) yield [rows[i], ...tail]
      }
    }
    for (const values of permutations([5, 6, 7, 8, 9])) for (let mask = 0; mask < 7; mask++) {
      const matches = (row: number) => Math.floor(mask / 2 ** (row - 6)) % 2 === 1
      const context = filtered(values, 5, { r0: 5, r1: 9, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, matches)
      const window = createScrollWindow(context.metrics, viewport)
      const candidates = values.map((source, slot) => ({ source, slot: slot + 5 })).filter(({ source }) => source >= 6 && source <= 8 && matches(source))
      expect(navigateKey(window, context, 'Home', null).active?.row ?? null).toBe(candidates[0]?.source ?? null)
      expect(navigateKey(window, context, 'End', null, { ctrl: true }).active?.row ?? null).toBe(candidates.at(-1)?.source ?? null)
      for (let hidden = 6; hidden <= 8; hidden++) if (!matches(hidden)) {
        const slot = values.indexOf(hidden) + 5
        const nearest = [...candidates].sort((a, b) => Math.abs(a.slot - slot) - Math.abs(b.slot - slot) || b.slot - a.slot)[0]
        const result = navigateTo(window, context, { addressSpace: 'source', at: { row: hidden, col: 0 } }, { row: hidden, col: 0 })
        expect(result.active?.row ?? null).toBe(nearest?.source ?? null)
      }
    }
  })
  it('skips visible-mask rows with authored zero height in both display directions', () => {
    const original = filtered([8, 5, 7, 6, 9], 5, { r0: 5, r1: 9, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, row => row !== 7)
    for (const collapsed of [[{ start: 6, end: 6 }], [{ start: 8, end: 8 }], [{ start: 6, end: 8 }]]) {
      const context = { ...original, metrics: createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 12, defaultSizeCss: 24, collapsed }), cols: original.metrics.cols, order: original.metrics.order, visibility: original.view.visibility }) }
      const window = createScrollWindow(context.metrics, viewport)
      const expected = collapsed[0].end === 8 ? (collapsed[0].start === 6 ? null : 6) : 8
      expect(navigateTo(window, context, { addressSpace: 'source', at: { row: 7, col: 0 } }, { row: 7, col: 0 }).active?.row ?? null).toBe(expected)
      expect(navigateKey(window, context, 'Home', null).active?.row ?? null).toBe(expected)
    }
  })
  it('bounds membership probes for an empty one-record body permuted across the full row domain', () => {
    const sourceAtOffset = new Uint32Array(MAX_ROWS)
    for (let slot = 0; slot < MAX_ROWS - 1; slot++) sourceAtOffset[slot] = slot + 1
    sourceAtOffset[MAX_ROWS - 1] = 0
    const order = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1, sourceAtOffset }])
    const filter: FilterSpec = { range: { r0: 5, r1: 7, c0: 0, c1: 0, headerRows: 1, totalsRows: 1 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }
    const baseHidden = createIntervalSet([])
    const applied = projectVisibility(baseHidden, filter, order, () => false)
    const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility: applied.visibility })
    let probes = 0
    const measuredOrder = { ...order, displayToSource(slot: number) { probes++; return order.displayToSource(slot) }, sourceToDisplay(row: number) { probes++; return order.sourceToDisplay(row) } }
    const context: NavigationContext = { viewport, metrics: { ...metrics, order: measuredOrder }, view: { baseHidden, filter, visibility: applied.visibility, selection: { active: null, anchor: null, ranges: [] } } }
    const window = createScrollWindow(context.metrics, viewport)
    expect(navigateKey(window, context, 'Home', null).kind).toBe('empty')
    expect(navigateTo(window, context, { addressSpace: 'source', at: { row: 6, col: 0 } }, { row: 6, col: 0 }).active).toBeNull()
    expect(probes).toBeLessThanOrEqual(5)
  })
  it('returns null for invalidated out-of-domain active coordinates instead of throwing', () => {
    const context = filtered([6, 5, 7], 5, { r0: 5, r1: 7, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, row => row !== 6)
    const window = createScrollWindow(context.metrics, viewport)
    for (const active of [{ row: -1, col: 0 }, { row: NaN, col: 0 }, { row: Infinity, col: 0 }, { row: 0, col: MAX_COLS }, { row: MAX_ROWS, col: 0 }, { row: -0, col: 0 }]) {
      const result = navigateTo(window, context, navigatorRequestFromA1('A7', 'source'), active)
      expect(result.active).toBeNull()
      expect(result.activate).toBe(false)
    }
  })
  it('guards malformed filter maps with RangeError consistently in normalization and navigation', () => {
    const context = contextFor(4, null, createIntervalSet([{ r0: 1, r1: 1 }]))
    const window = createScrollWindow(context.metrics, viewport)
    for (const malformed of [{}, { columns: null }, { columns: {} }]) {
      const filter = malformed as FilterSpec
      expect(() => normalizeFilter(filter)).toThrow(RangeError)
      expect(() => navigateTo(window, { ...context, view: { ...context.view, filter } }, navigatorRequestFromA1('A2', 'source'), { row: 0, col: 0 })).toThrow(RangeError)
      expect(() => navigateKey(window, { ...context, view: { ...context.view, filter } }, 'Home', null)).toThrow(RangeError)
    }
    expect(normalizeFilter(null)).toBeNull()
    expect(normalizeFilter({ range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map() })).toBeNull()
  })
})
