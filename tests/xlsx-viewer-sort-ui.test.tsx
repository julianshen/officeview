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
