import { afterEach, expect, it } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'
import type { GridTransition } from '../src/components/xlsx/GridSemantics'
import { workbookOf } from './helpers/xlsx-find'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(cleanup)

const rowY = (row: HTMLElement) => Number(row.style.transform.match(/^translate\([-\d.]+px, ([-\d.]+)px\)$/)![1])

it('a frozen keyboard step that widens the row gutter keeps the active row where it was', FIRST_PAINT_TEST, async () => {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'head' }, { row: 10010, col: 5, value: 'tail' }] }])
  doc.sheets[0].frozen = { rows: 1, cols: 0 }
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  const grid = ui.getByRole('grid')
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'F9980' } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  await waitFor(() => expect(grid.getAttribute('aria-activedescendant')).toBeTruthy(), FIRST_PAINT_WAIT)
  act(() => grid.focus())
  const gutterWidth = () => ui.getAllByRole('button', { name: /^Select row \d+/ })[0].style.width
  const activeRow = () => {
    const id = grid.getAttribute('aria-activedescendant')!
    return document.getElementById(id)!.parentElement!
  }
  const narrow = gutterWidth()
  let crossed = false
  // Walk down until the last visible row label reaches 10000 (gutter 44 -> 46px).
  for (let step = 0; step < 40 && !crossed; step++) {
    const before = activeRow()
    const beforeY = rowY(before), beforeIndex = Number(before.getAttribute('aria-rowindex'))
    fireEvent.keyDown(grid, { key: 'ArrowDown' })
    const after = activeRow()
    expect(Number(after.getAttribute('aria-rowindex'))).toBe(beforeIndex + 1)
    if (gutterWidth() === narrow) continue
    crossed = true
    // Only the gutter changed width: the active row stays at the bottom of the
    // body instead of jumping to the first body row (frozen Go To policy).
    expect(rowY(after)).toBe(beforeY)
  }
  expect(crossed).toBe(true)
})

const visibility = visibilityFromHidden(createIntervalSet([]))
const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 100, defaultSizeCss: 64 }) })
const viewport = { width: 640, height: 400 }
const view = { baseHidden: createIntervalSet([]), visibility, filter: null, sort: null }

it.each([
  ['frozen', { rows: 1, cols: 1 }],
  ['unfrozen', undefined],
] as const)('a column header click on a %s sheet keeps a visible active row in place', (_kind, frozen) => {
  const window = createScrollWindow(projection, viewport, { offset: { x: 0, y: 400 }, frozen })
  const selection = selectCell(emptyGridSelection(), projection, { row: 30, col: 0 })
  let next: GridTransition | undefined
  const ui = render(<GridSemantics {...gridFixture({ projection, viewport, window, selection, view, onChange: t => { next = t } })} />)
  fireEvent.click(ui.getByRole('button', { name: /^Select Column D\b/ }))
  expect(next!.selection.active).toEqual({ row: 30, col: 3 })
  expect(next!.window.logicalOrigin).toEqual({ x: 0, y: 400 })
})
