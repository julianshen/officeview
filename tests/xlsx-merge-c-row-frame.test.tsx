import { afterEach, expect, it } from 'vitest'
import { cleanup, render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createScrollWindow, logicalToViewport } from '../src/xlsx/view/scrollWindow'
import { gridFixture } from './helpers/xlsx-grid-fixture'

afterEach(cleanup)

it('keeps an unclipped body-cell left stable during horizontal scroll with frozen column A', () => {
  const props = gridFixture()
  const at = (x: number) => createScrollWindow(props.projection, props.viewport, { offset: { x, y: 0 }, frozen: { rows: 1, cols: 1 } })
  // Fixture: A80px, hidden B0px, C80px, D80px. D starts160 and stays entirely
  // beyond pinned A80 after37px scroll (viewport123..203 inside80..240).
  expect(props.projection.cols.start(3)).toBe(160)
  expect(props.projection.cols.size(3)).toBe(80)
  expect(logicalToViewport(at(37), { x: 160, y: 0 }).x).toBe(123)
  const ui = render(<GridSemantics {...props} window={at(0)} />)
  const cell = ui.container.querySelector<HTMLElement>('[role=row] [role=gridcell][aria-colindex="5"]')!
  expect(cell).not.toBeNull()
  const before = cell.style.left
  const id = cell.id
  ui.rerender(<GridSemantics {...props} window={at(37)} />)
  const after = document.getElementById(id)!
  expect(after).not.toBeNull()
  // The established integration horizontal-scroll invariant, now with a frozen strip.
  expect(after.style.left).toBe(before)
})
