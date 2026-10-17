import { afterEach, expect, it } from 'vitest'
import { cleanup, render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { gridFixture } from './helpers/xlsx-grid-fixture'

afterEach(cleanup)
const rowX = (cell: HTMLElement) => Number(cell.parentElement!.style.transform.match(/^translate\(([-\d.]+)px,/)![1])
const ownX = (element: HTMLElement) => Number(element.style.transform.match(/^translateX\(([-\d.]+)px\)$/)?.[1] ?? 0)
const pageX = (cell: HTMLElement) => rowX(cell) + Number.parseFloat(cell.style.left || '0') + ownX(cell)

it.each([1, 1.25])('moves only the body row frame at zoom %s while frozen cells and gutter stay fixed', zoom => {
  const props = gridFixture()
  const at = (x: number) => createScrollWindow(props.projection, props.viewport, { zoom, offset: { x, y: 0 }, frozen: { rows: 1, cols: 1 } })
  const ui = render(<GridSemantics {...props} window={at(0)} />)
  const pinned = ui.container.querySelector<HTMLElement>('[role=row] [role=gridcell][aria-colindex="2"]')!
  const body = ui.container.querySelector<HTMLElement>('[role=row] [role=gridcell][aria-colindex="5"]')!
  const header = pinned.parentElement!.querySelector<HTMLElement>('[role=rowheader]')!
  const old = { pinned: pageX(pinned), body: pageX(body), frame: rowX(body), bodyLeft: body.style.left, pinnedLeft: pinned.style.left, header: rowX(pinned) + ownX(header) }
  ui.rerender(<GridSemantics {...props} window={at(10)} />)
  expect(body.style.left).toBe(old.bodyLeft)
  // A pinned cell keeps its own `left` too: only its counter-transform moves.
  expect(pinned.style.left).toBe(old.pinnedLeft)
  expect(rowX(body)).toBeCloseTo(old.frame - 10 * zoom)
  expect(pageX(body)).toBeCloseTo(old.body - 10 * zoom)
  expect(pageX(pinned)).toBeCloseTo(old.pinned)
  expect(rowX(pinned) + ownX(header)).toBe(old.header)
  expect(pinned.parentElement!.getAttribute('role')).toBe('row')
  expect(body.parentElement).toBe(pinned.parentElement)
})
