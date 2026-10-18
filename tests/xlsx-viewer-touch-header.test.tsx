/**
 * Stream F item 2 (iOS Simulator run): on a coarse pointer each column header
 * stacked its label, sort, filter and more-sort targets vertically, so the
 * column-header row was 4 x 44 = 176 px tall on iPhone and iPad. Spec §9.3
 * keeps 44 px targets with deterministic ownership; the row now holds the
 * label row plus ONE actions row (2 x 44 px): a column wide enough for its
 * actions side by side shows them inline, a narrower one shows a disclosure
 * ("Actions for column B", aria-expanded/controls) that opens them as a
 * 44 x 44 group under its header. Columns narrower than 44 px keep the
 * corner lane (S13.2 item 5).
 */
import { describe, expect, it, vi } from 'vitest'
import { act, fireEvent, render } from '@testing-library/react'
import { GRID_HEADER_SIZE, GridHeaders, columnHeaderHeight } from '../src/components/xlsx/GridHeaders'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'

const projectionOf = (width: number) => createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: width }) })
function headers(width: number, options: { more?: boolean; controlMode?: boolean } = {}) {
  const projection = projectionOf(width), viewport = { width: 600, height: 400 }
  const handlers = { onSortCycle: vi.fn(), onFilter: vi.fn(), onSelect: vi.fn(), ...(options.more === false ? {} : { onSortDialog: vi.fn() }) }
  const ui = render(<GridHeaders projection={projection} window={createScrollWindow(projection, viewport)} viewport={viewport} columns={[0, 1, 2]} controlMode={options.controlMode}
    selection={selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })} sort={null} touchControls {...handlers} />)
  return { ui, handlers }
}
const size = (el: HTMLElement) => ({
  width: Math.max(parseFloat(el.style.width) || 0, parseFloat(el.style.minWidth) || 0),
  height: Math.max(parseFloat(el.style.height) || 0, parseFloat(el.style.minHeight) || 0),
})

describe('touch column headers: one label row plus one actions row', () => {
  it('the header row is at most 2 x 44 px with sort, filter and more-sort actions', () => {
    expect(columnHeaderHeight(true, true, true)).toBeLessThanOrEqual(2 * GRID_HEADER_SIZE)
    expect(columnHeaderHeight(true, true, false)).toBeLessThanOrEqual(2 * GRID_HEADER_SIZE)
    expect(columnHeaderHeight(false, true, true)).toBe(GRID_HEADER_SIZE)
    const { ui } = headers(64)
    for (const letter of ['A', 'B', 'C']) expect(parseFloat(ui.getByRole('columnheader', { name: `Column ${letter}` }).style.height)).toBeLessThanOrEqual(2 * GRID_HEADER_SIZE)
  })

  it('a column wide enough lays its actions side by side in the actions row, each a 44 px target inside its column', () => {
    const { ui, handlers } = headers(160)
    const header = ui.getByRole('columnheader', { name: 'Column B' })
    expect(ui.queryByRole('button', { name: 'Actions for column B' })).toBeNull()
    const actions = [ui.getByRole('button', { name: 'Cycle sort column B' }), ui.getByRole('button', { name: /^Filter column B/ }), ui.getByRole('button', { name: 'More sort options for B' })]
    const row = actions[0].parentElement!
    for (const action of actions) {
      expect(header.contains(action)).toBe(true)
      expect(action.parentElement).toBe(row)
      expect(size(action).width).toBeGreaterThanOrEqual(44)
      expect(size(action).height).toBeGreaterThanOrEqual(44)
    }
    expect(row.style.display).toBe('flex')
    expect(actions.reduce((sum, action) => sum + size(action).width, 0)).toBeLessThanOrEqual(parseFloat(header.style.width) + 1e-6)
    fireEvent.click(actions[0])
    expect(handlers.onSortCycle).toHaveBeenCalledWith(1, false, actions[0])
  })

  it('a default 64 px column gets a 44 px disclosure that opens its actions as 44 x 44 targets', () => {
    const { ui, handlers } = headers(64)
    const header = ui.getByRole('columnheader', { name: 'Column B' })
    const disclosure = ui.getByRole('button', { name: 'Actions for column B' })
    expect(header.contains(disclosure)).toBe(true)
    expect(size(disclosure).width).toBeGreaterThanOrEqual(44)
    expect(size(disclosure).height).toBeGreaterThanOrEqual(44)
    expect(size(disclosure).width).toBeLessThanOrEqual(64)
    expect(disclosure.getAttribute('aria-expanded')).toBe('false')
    expect(ui.queryByRole('button', { name: 'Cycle sort column B' })).toBeNull()
    fireEvent.click(disclosure)
    expect(disclosure.getAttribute('aria-expanded')).toBe('true')
    const group = ui.getByRole('group', { name: 'Column B actions' })
    expect(disclosure.getAttribute('aria-controls')).toBe(group.id)
    const actions = [ui.getByRole('button', { name: 'Cycle sort column B' }), ui.getByRole('button', { name: /^Filter column B/ }), ui.getByRole('button', { name: 'More sort options for B' })]
    for (const action of actions) {
      expect(group.contains(action)).toBe(true)
      expect(size(action).width).toBeGreaterThanOrEqual(44)
      expect(size(action).height).toBeGreaterThanOrEqual(44)
    }
    // An action closes the group.
    fireEvent.click(actions[1])
    expect(handlers.onFilter).toHaveBeenCalledWith(1, actions[1])
    expect(ui.queryByRole('group', { name: 'Column B actions' })).toBeNull()
    // A tap outside closes it too.
    fireEvent.click(disclosure)
    fireEvent.pointerDown(document.body)
    expect(ui.queryByRole('group', { name: 'Column B actions' })).toBeNull()
  })

  it('in control mode the group takes focus and Escape returns it to the disclosure', () => {
    const { ui } = headers(64, { controlMode: true })
    const disclosure = ui.getByRole('button', { name: 'Actions for column C' })
    expect(disclosure.tabIndex).toBe(0)
    act(() => disclosure.focus())
    fireEvent.click(disclosure)
    const sort = ui.getByRole('button', { name: 'Cycle sort column C' })
    expect(document.activeElement).toBe(sort)
    fireEvent.keyDown(sort, { key: 'Escape' })
    expect(ui.queryByRole('group', { name: 'Column C actions' })).toBeNull()
    expect(document.activeElement).toBe(disclosure)
  })
})
