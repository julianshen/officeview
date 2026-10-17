import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { ResizeHandles } from '../src/components/xlsx/sheetResize'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { gridFixture } from './helpers/xlsx-grid-fixture'

afterEach(cleanup)
const visibility = visibilityFromHidden(createIntervalSet([]))
const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }), order: createIdentityOrder(), visibility })
const viewport = { width: 400, height: 240 }
const at = (x = 840) => createScrollWindow(projection, viewport, { offset: { x, y: 2050 }, frozen: { rows: 1, cols: 1 } })
const xOf = (el: HTMLElement) => Number(/^translate\(([-\d.]+)px,/.exec(el.style.transform)![1])
function mount(touch = false, col = 11, x = 840) {
  const onResize = vi.fn()
  const ui = render(<ResizeHandles projection={projection} window={at(x)} viewport={viewport} headerHeight={touch ? 88 : 24} gutter={44} touch={touch} columns={[0, 11]} slots={[0, 103]} active={{ row: 103, col }} onResize={onResize} />)
  return { ui, onResize }
}

it('a hidden transient active body cell supplies semantics but no resize handles under frozen strips', () => {
  const active = { row: 102, col: 10 }
  const props = gridFixture({ projection, viewport, window: at(), body: null, view: { baseHidden: createIntervalSet([]), visibility, filter: null, sort: null }, selection: { ranges: [], active, anchor: active }, onResize: vi.fn() })
  const ui = render(<GridSemantics {...props} />)
  expect(ui.container.querySelector('[data-transient-active="true"] [role=gridcell]')).not.toBeNull()
  expect(ui.container.querySelector('[data-resize-column="10"]')).toBeNull()
  expect(ui.container.querySelector('[data-resize-row="102"]')).toBeNull()
})

it('a body-column drag guide disappears while its resized boundary lies under the frozen column', () => {
  const { ui, onResize } = mount()
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="11"]')!
  expect(handle).not.toBeNull()
  // Logical column11 starts880; viewport40..120 is partly visible beyond80.
  fireEvent.pointerDown(handle, { button: 0, pointerId: 1, clientX: 120 })
  fireEvent.pointerMove(window, { pointerId: 1, clientX: 60 })
  // New width20 puts its edge at viewport60, hidden under frozen A0..80.
  expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')).toBeNull()
  fireEvent.pointerUp(window, { pointerId: 1, clientX: 60 })
  expect(onResize).toHaveBeenCalledWith({ kind: 'column', index: 11, size: 20 })
})

it('a body-row drag guide disappears while its resized boundary lies under the frozen row', () => {
  const { ui, onResize } = mount()
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-row="103"]')!
  expect(handle).not.toBeNull()
  // Row103 starts2060: viewport10..30; shrinking20→8 leaves edge18 under strip20.
  fireEvent.pointerDown(handle, { button: 0, pointerId: 2, clientY: 30 })
  fireEvent.pointerMove(window, { pointerId: 2, clientY: 10 })
  expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')).toBeNull()
  fireEvent.pointerUp(window, { pointerId: 2, clientY: 10 })
  expect(onResize).toHaveBeenCalledWith({ kind: 'row', index: 103, size: 8 })
})

it('visible body guide remains correctly positioned and commits once (pointer-event control)', () => {
  const { ui, onResize } = mount()
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="11"]')!
  fireEvent.pointerDown(handle, { button: 0, pointerId: 3, clientX: 120 })
  fireEvent.pointerMove(window, { pointerId: 3, clientX: 140 })
  const guide = ui.container.querySelector<HTMLElement>('[data-officeview-ui="resize-guide"]')!
  expect(guide).not.toBeNull()
  expect(xOf(guide)).toBe(44 + 140 - 1)
  expect(onResize).not.toHaveBeenCalled()
  fireEvent.pointerUp(window, { pointerId: 3, clientX: 140 })
  expect(onResize).toHaveBeenCalledTimes(1)
  expect(onResize).toHaveBeenCalledWith({ kind: 'column', index: 11, size: 100 })
})

it('a coarse body handle retains44px within its own body header region', () => {
  const { ui } = mount(true, 11, 878)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="11"]')!
  expect(handle).not.toBeNull()
  expect(Number.parseFloat(handle.style.width)).toBe(44)
  // The body column has only2px beyond frozen A; slide its target into the body.
  expect(xOf(handle)).toBeGreaterThanOrEqual(44 + 80)
  expect(xOf(handle) + 44).toBeLessThanOrEqual(44 + 400)
})

it('a coarse frozen-column handle retains44px without stealing the body header', () => {
  const { ui } = mount(true, 0)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!
  expect(handle).not.toBeNull()
  expect(Number.parseFloat(handle.style.width)).toBe(44)
  expect(xOf(handle)).toBeGreaterThanOrEqual(44)
  expect(xOf(handle) + 44).toBeLessThanOrEqual(44 + 80)
})
