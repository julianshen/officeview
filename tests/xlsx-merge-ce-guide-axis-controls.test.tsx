import { afterEach, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { ResizeHandles } from '../src/components/xlsx/sheetResize'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'

afterEach(cleanup)
const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }), order: createIdentityOrder() })
const viewport = { width: 400, height: 240 }
function mount(rows: number, cols: number, zoom = 1) {
  const scroll = createScrollWindow(projection, viewport, { offset: { x: 840, y: 2050 }, zoom, frozen: { rows, cols } })
  return render(<ResizeHandles projection={projection} window={scroll} viewport={viewport} headerHeight={24} gutter={44} touch={false} columns={[0, 11]} slots={[0, 103]} active={{ row: 103, col: 11 }} onResize={vi.fn()} />)
}
const guideXY = (ui: ReturnType<typeof render>) => {
  const guide = ui.container.querySelector<HTMLElement>('[data-officeview-ui="resize-guide"]')!
  expect(guide).not.toBeNull()
  return /^translate\(([-\d.]+)px,\s*([-\d.]+)px\)$/.exec(guide.style.transform)!.slice(1).map(Number)
}

it('growing pinned A beyond the old divider projects its proposed edge with fixed origin', () => {
  const ui = mount(0, 1)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="0"]')!
  fireEvent.pointerDown(handle, { button: 0, pointerId: 1, clientX: 124 })
  fireEvent.pointerMove(window, { pointerId: 1, clientX: 164 })
  expect(guideXY(ui)).toEqual([163, 0]) // gutter44 + new fixed edge120 -1
})

it('growing pinned row1 beyond the old divider projects its proposed edge with fixed origin', () => {
  const ui = mount(1, 0)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-row="0"]')!
  fireEvent.pointerDown(handle, { button: 0, pointerId: 2, clientY: 44 })
  fireEvent.pointerMove(window, { pointerId: 2, clientY: 84 })
  expect(guideXY(ui)).toEqual([0, 83]) // header24 + new fixed edge60 -1
})

it('column-only freezing at1.25zoom hides a body proposal at75px under the100px strip', () => {
  const ui = mount(0, 1, 1.25)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-column="11"]')!
  expect(handle).not.toBeNull()
  fireEvent.pointerDown(handle, { button: 0, pointerId: 3, clientX: 194 })
  fireEvent.pointerMove(window, { pointerId: 3, clientX: 119 })
  expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')).toBeNull()
})

it('row-only freezing at1.25zoom hides a body proposal at22.5px under the25px strip', () => {
  const ui = mount(1, 0, 1.25)
  const handle = ui.container.querySelector<HTMLElement>('[data-resize-row="103"]')!
  expect(handle).not.toBeNull()
  fireEvent.pointerDown(handle, { button: 0, pointerId: 4, clientY: 61.5 })
  fireEvent.pointerMove(window, { pointerId: 4, clientY: 46.5 })
  expect(ui.container.querySelector('[data-officeview-ui="resize-guide"]')).toBeNull()
})
