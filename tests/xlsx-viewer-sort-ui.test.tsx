import { describe, expect, it, vi } from 'vitest'
import { act, fireEvent, render, waitFor } from '@testing-library/react'
import { SortDialog } from '../src/components/xlsx/SortDialog'
import { controlsFixture } from './helpers/xlsx-controls-fixture'
import { createControlsClient } from '../src/components/xlsx/controlsClient'

const setup = (table = false) => {
  const fixture = controlsFixture(8, { kind: 'number', hasNumeric: true }, table), close = vi.fn(), notice = vi.fn(), applied = vi.fn()
  const props = { owner: fixture.snapshot.owner, client: fixture.client, applied: null, onClose: close, onNotice: notice, onApplied: applied }
  return { fixture, props, close, notice, applied }
}
describe('S9 atomic sort modal', () => {
  it('adds/removes/reorders unique keys up to eight, keeps a private case/header draft and applies only on OK', async () => {
    const f = setup(), ui = render(<SortDialog {...f.props} />)
    await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
    expect(ui.getByText('View-only sort; formulas unchanged.')).toBeTruthy()
    for (let i = 1; i < 8; i++) fireEvent.click(ui.getByRole('button', { name: 'Add level' }))
    expect(ui.container.querySelectorAll('[data-sort-level]').length).toBe(8)
    expect(ui.getByRole('button', { name: 'Add level' }).hasAttribute('disabled')).toBe(true)
    fireEvent.click(ui.getByRole('button', { name: 'Move level 2 up' }))
    expect((ui.getByRole('textbox', { name: 'Column level 1' }) as HTMLInputElement).value).toBe('B')
    fireEvent.click(ui.getByRole('button', { name: 'Remove level 8' }))
    fireEvent.click(ui.getByRole('checkbox', { name: 'Case sensitive' }))
    expect(f.fixture.publications).toBe(0)
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(f.applied).toHaveBeenCalledOnce())
    expect(f.fixture.snapshot.sort?.keys).toHaveLength(7)
    expect(f.fixture.snapshot.sort?.keys[0]).toEqual({ col: 1, order: 'asc' })
    expect(f.fixture.snapshot.sort?.caseSensitive).toBe(true)
    expect(f.fixture.publications).toBe(1)
  })
  it('traps modal Tab, returns Escape focus and publishes no cancelled changes', async () => {
    const f = setup(), invoker = document.createElement('button'); document.body.append(invoker)
    const ui = render(<SortDialog {...f.props} invoker={invoker} />)
    await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
    const first = ui.getByRole('textbox', { name: 'Column level 1' }), cancel = ui.getByRole('button', { name: 'Cancel' })
    act(() => cancel.focus()); fireEvent.keyDown(cancel, { key: 'Tab' })
    expect(document.activeElement).toBe(first)
    fireEvent.keyDown(first, { key: 'Tab', shiftKey: true })
    expect(document.activeElement).toBe(cancel)
    fireEvent.keyDown(cancel, { key: 'Escape' })
    expect(document.activeElement).toBe(invoker)
    expect(f.fixture.publications).toBe(0)
    expect(f.close).toHaveBeenCalledOnce()
    invoker.remove()
  })
  it('respects no-header table metadata and allows a non-table header override', async () => {
    const table = setup(true), ui = render(<SortDialog {...table.props} />)
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'Range has a header row' })).toBeTruthy())
    const header = ui.getByRole('checkbox', { name: 'Range has a header row' }) as HTMLInputElement
    expect(header.disabled).toBe(true)
    expect(header.checked).toBe(false)
    expect(ui.getByText(/Table metadata/)).toBeTruthy()
    ui.unmount()
    const region = setup(), next = render(<SortDialog {...region.props} />)
    await waitFor(() => expect((next.getByRole('checkbox', { name: 'Range has a header row' }) as HTMLInputElement).checked).toBe(true))
    fireEvent.click(next.getByRole('checkbox', { name: 'Range has a header row' }))
    await waitFor(() => expect(next.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.click(next.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(region.applied).toHaveBeenCalledOnce())
    expect(region.fixture.snapshot.sort?.range.headerRows).toBe(0)
  })
  it('shows a guard refusal and never publishes a stale or cancelled staged sort', async () => {
    const f = setup(), client = createControlsClient({ ...f.fixture.options, stageSort: async () => ({ kind: 'refused', refusal: { reason: 'merge', range: f.fixture.range } }) })
    const ui = render(<SortDialog {...f.props} client={client} />)
    await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(f.notice).toHaveBeenCalledWith('Operation refused: merge.'))
    expect(f.fixture.publications).toBe(0)
    expect(f.applied).not.toHaveBeenCalled()
  })
})

it('sort Cancel/reopen resets private levels and never leaves an empty initial draft', async () => {
  const f = setup(), ui = render(<SortDialog {...f.props} />)
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
  fireEvent.click(ui.getByRole('button', { name: 'Add level' }))
  ui.rerender(<SortDialog {...f.props} open={false} />)
  ui.rerender(<SortDialog {...f.props} open />)
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
  expect(ui.queryByRole('textbox', { name: 'Column level 2' })).toBeNull()
})

import { GridHeaders } from '../src/components/xlsx/GridHeaders'
import { gridFixture } from './helpers/xlsx-grid-fixture'
it('provides separate named select/sort/funnel targets and reports committed rule state on tap', () => {
  const p = gridFixture(), sort = vi.fn(), filter = vi.fn(), select = vi.fn(), notice = vi.fn()
  const ui = render(<div role="grid"><GridHeaders projection={p.projection} window={p.window} viewport={p.viewport} columns={[0]} selection={p.selection} sort={null} controlMode onSelect={select}
    filter={{ range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }} onSortCycle={sort} onFilter={filter} onNotice={notice} /></div>)
  fireEvent.click(ui.getByRole('button', { name: 'Select Column A' }))
  expect(select).toHaveBeenCalledOnce()
  expect(sort).not.toHaveBeenCalled()
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }), { shiftKey: true })
  expect(sort).toHaveBeenCalledWith(0, true, expect.any(HTMLButtonElement))
  fireEvent.click(ui.getByRole('button', { name: /Filter column A.*Greater than 0/ }))
  expect(filter).toHaveBeenCalledWith(0, expect.any(HTMLButtonElement))
  expect(notice).toHaveBeenCalledWith(expect.stringContaining('Greater than 0'))
  expect(ui.getAllByRole('button').every(button => button.style.minHeight === '44px' || button.style.height === '44px')).toBe(true)
})

import { cycleHeaderSort } from '../src/components/xlsx/controlsClient'
it('cycles a guarded committed header sort through ascending, descending and unsorted', async () => {
  const f = controlsFixture(8)
  for (const order of ['asc', 'desc', null]) {
    const result = await cycleHeaderSort(f.client, f.snapshot, 0, false, new AbortController().signal)
    expect(result.kind).toBe('committed')
    expect(f.snapshot.sort?.keys[0].order ?? null).toBe(order)
  }
  expect(f.publications).toBe(3)
})

it('aborts a pending sort on Cancel even if the worker ignores cancellation until it returns', async () => {
  const f = setup()
  let finish!: (stage: Awaited<ReturnType<typeof f.fixture.options.stageSort>>) => void
  const client = createControlsClient({ ...f.fixture.options, stageSort: () => new Promise(resolve => { finish = resolve }) })
  const ui = render(<SortDialog {...f.props} client={client} />)
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(finish).toBeTypeOf('function'))
  fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  finish({ kind: 'ready', publish: () => { throw new Error('Cancelled sort published') } })
  await act(async () => { await Promise.resolve() })
  expect(f.applied).not.toHaveBeenCalled()
  expect(f.fixture.publications).toBe(0)
})

import { GridSemantics } from '../src/components/xlsx/GridSemantics'
it('keeps sort and funnel actions visible and pointer-usable in navigation mode without entering control mode', () => {
  const p = gridFixture(), sort = vi.fn(), filter = vi.fn()
  const ui = render(<GridSemantics {...p} onSortCycle={sort} onFilter={filter} />)
  const grid = ui.getByRole('grid')
  act(() => grid.focus())
  const arrow = ui.getByRole('button', { name: 'Cycle sort column A' })
  const funnel = ui.getByRole('button', { name: /Filter column A/ })
  expect(arrow.tabIndex).toBe(-1)
  expect(funnel.tabIndex).toBe(-1)
  expect(arrow.closest('[hidden]')).toBeNull()
  expect(funnel.closest('[hidden]')).toBeNull()
  expect(fireEvent.mouseDown(arrow)).toBe(false)
  act(() => arrow.focus())
  expect(document.activeElement).toBe(grid)
  fireEvent.click(arrow)
  fireEvent.click(funnel)
  expect(sort).toHaveBeenCalledWith(0, false, arrow)
  expect(filter).toHaveBeenCalledWith(0, funnel)
  expect(grid.tabIndex).toBe(0)
  expect(grid.querySelectorAll('[role="gridcell"][tabindex]')).toHaveLength(0)
  fireEvent.keyDown(grid, { key: 'F2' })
  expect(ui.getByRole('button', { name: 'Select Column A' }).tabIndex).toBe(0)
  expect(arrow.tabIndex).toBe(0)
  expect(funnel.tabIndex).toBe(0)
})

it('shows the committed sort direction and pressed arrow/funnel state without focused headers', () => {
  const p = gridFixture(), ui = render(<GridHeaders projection={p.projection} window={p.window} viewport={p.viewport} columns={[0]} selection={p.selection}
    controlMode={false} onSortCycle={() => {}} onFilter={() => {}}
    sort={{ range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, keys: [{ col: 0, order: 'desc' }], caseSensitive: false }}
    filter={{ range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }} />)
  const arrow = ui.getByRole('button', { name: 'Cycle sort column A' }), funnel = ui.getByRole('button', { name: /Filter column A/ })
  expect(arrow.getAttribute('aria-pressed')).toBe('true')
  expect(arrow.textContent).toContain('↓')
  expect(funnel.getAttribute('aria-pressed')).toBe('true')
  expect(funnel.querySelector('svg')).toBeTruthy()
  expect(ui.getByRole('columnheader', { name: /Column A/ }).getAttribute('aria-sort')).toBe('descending')
})

it.each(['ctrlKey', 'metaKey'])('restores the invoker before %s+F reaches workbook find', async modifier => {
  const f = setup(), invoker = document.createElement('button')
  document.body.append(invoker)
  let focusAtFind: Element | null = null
  const ui = render(<div onKeyDown={event => { if (event.key === 'f') focusAtFind = document.activeElement }}><SortDialog {...f.props} invoker={invoker} /></div>)
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
  const first = ui.getByRole('textbox', { name: 'Column level 1' })
  act(() => first.focus())
  fireEvent.keyDown(first, { key: 'f', [modifier]: true })
  expect(f.close).toHaveBeenCalledOnce()
  expect(focusAtFind).toBe(invoker)
  expect(document.activeElement).toBe(invoker)
  expect(f.fixture.publications).toBe(0)
  invoker.remove()
})

it('returns workbook-find focus to the grid when the modal invoker is detached', async () => {
  const f = setup(), grid = document.createElement('div'), fallback = vi.fn(() => grid.focus())
  grid.tabIndex = 0; document.body.append(grid)
  const ui = render(<SortDialog {...f.props} invoker={document.createElement('button')} fallbackFocus={fallback} />)
  await waitFor(() => expect(ui.getByRole('textbox', { name: 'Column level 1' })).toBeTruthy())
  fireEvent.keyDown(ui.getByRole('textbox', { name: 'Column level 1' }), { key: 'f', ctrlKey: true })
  expect(fallback).toHaveBeenCalledOnce()
  expect(document.activeElement).toBe(grid)
  grid.remove()
})

import { columnRuleDescription } from '../src/components/xlsx/GridHeaders'
it('describes committed date conditions with zero-padded ISO dates', () => {
  expect(columnRuleDescription({ kind: 'date', op: 'between', first: { year: 2026, month: 1, day: 2 }, second: { year: 2026, month: 10, day: 9 } })).toBe('Between 2026-01-02 through 2026-10-09')
})

it('reserves touch header chrome for visible 44px controls and retains source-space hit testing', () => {
  const previous = window.matchMedia
  window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null,
    addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true }))
  try {
    const p = gridFixture(), change = vi.fn()
    const ui = render(<GridSemantics {...p} onChange={change} onSortCycle={vi.fn()} onFilter={vi.fn()} />)
    const grid = ui.getByRole('grid')
    const header = ui.getByRole('button', { name: 'Select Column A' }).closest<HTMLElement>('[role="columnheader"]')!
    expect(header.style.height).toBe('132px')
    for (const button of header.querySelectorAll('button')) {
      expect(button.style.minWidth).toBe('44px')
      expect(button.style.minHeight).toBe('44px')
      expect(button.closest('[hidden]')).toBeNull()
    }
    // Rows follow the scroll window by transform (no layout shift).
    expect((grid.querySelector('[aria-rowindex="2"]') as HTMLElement).style.transform).toBe('translateY(132px)')
    act(() => grid.focus())
    fireEvent.click(grid, { detail: 1, clientX: 54, clientY: 142 })
    expect(change).toHaveBeenCalledOnce()
    expect(change.mock.calls[0][0].selection.active).toEqual({ row: 0, col: 0 })
    expect(document.activeElement).toBe(grid)
    ui.unmount()
  } finally { window.matchMedia = previous }
})

import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { MAX_ROWS, MAX_COLS } from '../src/xlsx/view/coordinates'
it.each([16, 40, 64])('keeps compact desktop header targets inside a %ipx projected column', width => {
  const p = gridFixture(), projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order: p.projection.order, visibility: p.view.visibility })
  const window = createScrollWindow(projection, p.viewport, { zoom: width / 64 })
  const ui = render(<GridHeaders projection={projection} window={window} viewport={p.viewport} columns={[0]} selection={p.selection}
    sort={null} onSortCycle={vi.fn()} onFilter={vi.fn()} touchControls={false} />)
  const header = ui.getByRole('columnheader', { name: 'Column A' })
  const buttons = [...header.querySelectorAll('button')]
  expect(buttons).toHaveLength(3)
  // Explicit CSS width/minimum constraints cannot demand more than the projected box.
  // The label flexes into what is left (no floor); the action targets have explicit
  // widths (visually hidden 1px ones when the column is too narrow for its glyphs).
  expect(buttons[0].style.flex).toBe('1 1 auto')
  const floors = buttons.map(button => Math.max(parseFloat(button.style.width) || 0, parseFloat(button.style.minWidth) || 0))
  expect(floors.slice(1).every(value => value > 0)).toBe(true)
  expect(floors.reduce((sum, value) => sum + value, 0)).toBeLessThanOrEqual(parseFloat(header.style.width) + 1e-6)
})
it.each([16, 40])('keeps touch header targets of a %ipx projected column out of the next column', width => {
  const p = gridFixture(), projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order: p.projection.order, visibility: p.view.visibility })
  const window = createScrollWindow(projection, p.viewport, { zoom: width / 64 })
  const ui = render(<GridHeaders projection={projection} window={window} viewport={p.viewport} columns={[0, 1]} selection={p.selection}
    sort={null} onSortCycle={vi.fn()} onFilter={vi.fn()} onSortDialog={vi.fn()} touchControls />)
  const a = ui.getByRole('columnheader', { name: 'Column A' }), b = ui.getByRole('columnheader', { name: 'Column B' })
  // Header boxes are placed by translateX (no layout shift).
  const x = (el: HTMLElement) => parseFloat(/translateX\(([-\d.]+)px\)/.exec(el.style.transform)![1])
  const right = x(a) + parseFloat(a.style.width)
  expect(right).toBeLessThanOrEqual(x(b) + 1e-6)
  // Deterministic hit ownership: no control of A may demand space past A's box.
  for (const button of a.querySelectorAll('button')) {
    expect(Math.max(parseFloat(button.style.width) || 0, parseFloat(button.style.minWidth) || 0)).toBeLessThanOrEqual(parseFloat(a.style.width) + 1e-6)
  }
  expect(a.style.overflow).toBe('hidden')
})
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { GRID_HEADER_SIZE } from '../src/components/xlsx/GridHeaders'
describe('S13.2 item 5: 44px touch targets for legal narrow columns (spec §9.3, A11Y-02)', () => {
  const narrowProjection = () => createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 20 }) })
  const narrow = (active: number, controlMode = false) => {
    const projection = narrowProjection()
    const viewport = { width: 200, height: 100 }
    const handlers = { onSortCycle: vi.fn(), onFilter: vi.fn(), onSortDialog: vi.fn(), onSelect: vi.fn() }
    const ui = render(<GridHeaders projection={projection} window={createScrollWindow(projection, viewport)} viewport={viewport} columns={[0, 1, 2]} controlMode={controlMode}
      selection={selectCell(emptyGridSelection(), projection, { row: 0, col: active })} sort={null} touchControls {...handlers} />)
    return { ui, handlers }
  }
  const box = (element: HTMLElement) => {
    const left = parseFloat(element.style.left), top = parseFloat(element.style.top)
    const width = Math.max(parseFloat(element.style.width) || 0, parseFloat(element.style.minWidth) || 0)
    const height = Math.max(parseFloat(element.style.height) || 0, parseFloat(element.style.minHeight) || 0)
    return { left, top, width, height }
  }
  const columnBoxes = (ui: ReturnType<typeof render>) => ['A', 'B', 'C'].map(letter => ui.getByRole('columnheader', { name: `Column ${letter}` }))
    .map(header => {
      const left = parseFloat(header.style.left) + parseFloat(/translateX\(([-\d.]+)px\)/.exec(header.style.transform)![1])
      return { left, right: left + parseFloat(header.style.width) }
    })
  const outsideColumns = (ui: ReturnType<typeof render>, target: { left: number; width: number }) => {
    // The corner lane: no column header box (all start at GRID_HEADER_SIZE) is overlapped.
    expect(target.left + target.width).toBeLessThanOrEqual(GRID_HEADER_SIZE)
    for (const column of columnBoxes(ui)) expect(target.left + target.width <= column.left || target.left >= column.right).toBe(true)
  }

  it('a 20px active column gets a ≥44x44 corner disclosure opening ≥44x44 select/sort/filter/more targets that no neighbour target overlaps', () => {
    const { ui, handlers } = narrow(1)
    const disclosure = ui.getByRole('button', { name: 'Actions for column B' })
    expect(box(disclosure).width).toBeGreaterThanOrEqual(44)
    expect(box(disclosure).height).toBeGreaterThanOrEqual(44)
    outsideColumns(ui, box(disclosure))
    expect(disclosure.getAttribute('aria-expanded')).toBe('false')
    fireEvent.click(disclosure)
    expect(disclosure.getAttribute('aria-expanded')).toBe('true')
    const group = ui.getByRole('group', { name: 'Column B actions' })
    expect(disclosure.getAttribute('aria-controls')).toBe(group.id)
    const targets = [ui.getByRole('button', { name: 'Select whole column B' }), ui.getByRole('button', { name: 'Cycle sort column B' }), ui.getByRole('button', { name: /^Filter column B/ }), ui.getByRole('button', { name: 'More sort options for B' })]
    for (const target of targets) {
      expect(group.contains(target)).toBe(true)
      const b = box(target)
      expect(b.width).toBeGreaterThanOrEqual(44)
      expect(b.height).toBeGreaterThanOrEqual(44)
      outsideColumns(ui, { left: parseFloat(group.style.left) + b.left, width: b.width })
    }
    const stacked = targets.map(box).sort((a, b) => a.top - b.top)
    for (let i = 1; i < stacked.length; i++) expect(stacked[i].top).toBeGreaterThanOrEqual(stacked[i - 1].top + stacked[i - 1].height)
    fireEvent.click(targets[1])
    expect(handlers.onSortCycle).toHaveBeenCalledWith(1, false, targets[1])
    // An action closes the group.
    expect(ui.queryByRole('group', { name: 'Column B actions' })).toBeNull()
    fireEvent.click(disclosure)
    fireEvent.click(ui.getByRole('button', { name: 'Select whole column B' }))
    expect(handlers.onSelect).toHaveBeenCalledWith('column', 1, { toggle: false, extend: false }, expect.anything())
    // No sub-44px action target is left inside the narrow column headers.
    for (const letter of ['A', 'B', 'C']) {
      const header = ui.getByRole('columnheader', { name: `Column ${letter}` })
      expect(header.querySelectorAll('[data-action-column]')).toHaveLength(0)
    }
  })

  it('follows the active column, closes on Escape with focus back on the disclosure in control mode', () => {
    const { ui } = narrow(2, true)
    expect(ui.queryByRole('button', { name: 'Actions for column B' })).toBeNull()
    const disclosure = ui.getByRole('button', { name: 'Actions for column C' })
    expect(disclosure.tabIndex).toBe(0)
    act(() => disclosure.focus())
    fireEvent.click(disclosure)
    const select = ui.getByRole('button', { name: 'Select whole column C' })
    expect(document.activeElement).toBe(select)
    fireEvent.keyDown(select, { key: 'Escape' })
    expect(ui.queryByRole('group', { name: 'Column C actions' })).toBeNull()
    expect(document.activeElement).toBe(disclosure)
  })

  it('the group closes when the active column changes and stays closed when it returns', () => {
    const projection = narrowProjection(), viewport = { width: 200, height: 100 }
    const props = (col: number) => ({ projection, window: createScrollWindow(projection, viewport), viewport, columns: [0, 1, 2], sort: null, touchControls: true,
      onSortCycle: vi.fn(), onFilter: vi.fn(), onSelect: vi.fn(), selection: selectCell(emptyGridSelection(), projection, { row: 0, col }) })
    const ui = render(<GridHeaders {...props(1)} />)
    fireEvent.click(ui.getByRole('button', { name: 'Actions for column B' }))
    expect(ui.getByRole('group', { name: 'Column B actions' })).toBeTruthy()
    ui.rerender(<GridHeaders {...props(2)} />)
    expect(ui.queryByRole('group', { name: 'Column C actions' })).toBeNull()
    ui.rerender(<GridHeaders {...props(1)} />)
    expect(ui.queryByRole('group', { name: 'Column B actions' })).toBeNull()
    expect(ui.getByRole('button', { name: 'Actions for column B' }).getAttribute('aria-expanded')).toBe('false')
  })

  it('a tap outside closes the group', () => {
    const { ui } = narrow(1)
    fireEvent.click(ui.getByRole('button', { name: 'Actions for column B' }))
    fireEvent.pointerDown(document.body)
    expect(ui.queryByRole('group', { name: 'Column B actions' })).toBeNull()
  })

  it('selecting a narrow column from the lane selects that column in the grid (no re-hit-test of the lane position)', () => {
    const previous = window.matchMedia
    window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null,
      addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true }))
    try {
      const projection = narrowProjection(), viewport = { width: 200, height: 100 }
      const change = vi.fn()
      const p = gridFixture({ projection, viewport, window: createScrollWindow(projection, viewport), view: { baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), filter: null, sort: null }, selection: selectCell(emptyGridSelection(), projection, { row: 0, col: 1 }) })
      const ui = render(<GridSemantics {...p} onChange={change} onSortCycle={vi.fn()} onFilter={vi.fn()} />)
      fireEvent.click(ui.getByRole('button', { name: 'Actions for column B' }))
      fireEvent.click(ui.getByRole('button', { name: 'Select whole column B' }), { detail: 1, clientX: 10, clientY: 100 })
      expect(change).toHaveBeenCalled()
      const selection = change.mock.calls.at(-1)![0].selection
      expect(selection.ranges[0].cols).toEqual([{ c0: 1, c1: 1 }])
      ui.unmount()
    } finally { window.matchMedia = previous }
  })
})
