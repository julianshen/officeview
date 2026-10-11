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
    expect((grid.querySelector('[aria-rowindex="2"]') as HTMLElement).style.top).toBe('132px')
    act(() => grid.focus())
    fireEvent.click(grid, { detail: 1, clientX: 54, clientY: 142 })
    expect(change).toHaveBeenCalledOnce()
    expect(change.mock.calls[0][0].selection.active).toEqual({ row: 0, col: 0 })
    expect(document.activeElement).toBe(grid)
    ui.unmount()
  } finally { window.matchMedia = previous }
})
