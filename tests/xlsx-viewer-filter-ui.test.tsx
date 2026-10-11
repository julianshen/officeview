import { describe, expect, it } from 'vitest'
import { controlsFixture } from './helpers/xlsx-controls-fixture'
import { createControlsClient, filterWithColumn } from '../src/components/xlsx/controlsClient'
import { compileValuesMembership } from '../src/xlsx/view/valueIndex'

const signal = () => new AbortController().signal

describe('S9 atomic client and complete values universe', () => {
  it('resolves through S7.1 and filters an exact excluded value beyond the display prefix', async () => {
    const f = controlsFixture(), owner = f.snapshot.owner, resolved = await f.client.resolve(owner, 'filter', signal(), undefined, 0)
    expect(resolved.kind).toBe('complete')
    const draft = await f.client.values(owner, f.range, 0, null, signal())
    const page = await draft.read({ text: 'beyond', offset: 0 }, signal())
    expect(page.entries.map(entry => entry.label)).toEqual(['beyond-A', 'beyond-B'])
    expect(page.count.complete).toBe(true)
    await draft.change({ kind: 'all', selected: true, includeBlanks: false }, { text: 'beyond', offset: 0 }, signal())
    await draft.change({ kind: 'toggle', ordinal: page.entries[1].ordinal, selected: false }, { text: 'beyond', offset: 0 }, signal())
    expect(f.publications).toBe(0)
    const rule = await draft.candidate(signal())
    expect(rule?.kind).toBe('values')
    const outcome = await f.client.filter(owner, filterWithColumn(null, f.range, 0, rule), signal(), undefined, draft)
    expect(outcome.kind).toBe('committed')
    expect(f.publications).toBe(1)
    const members = await compileValuesMembership(rule!, draft.identity, f.runs)
    expect(members.matches({ kind: 'string', value: 'beyond-A' })).toBe(true)
    expect(members.matches({ kind: 'string', value: 'beyond-B' })).toBe(false)
    expect(members.matches({ kind: 'string', value: 'item-10001' })).toBe(true)
    draft.cancel()
    expect(f.runs.refCount).toBe(1) // transferred to the applied view
  })
  it('drops late staged results on owner changes or cancellation, without publication', async () => {
    const f = controlsFixture(5)
    let finish!: (value: Awaited<ReturnType<typeof f.options.stageFilter>>) => void
    const client = createControlsClient({ ...f.options, stageFilter: () => new Promise(resolve => { finish = resolve }) })
    const old = f.snapshot.owner, pending = client.filter(old, { range: f.range, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }, signal())
    f.snapshot = { ...f.snapshot, owner: { ...old, valueGen: old.valueGen + 1 } }
    if (finish) finish(await f.options.stageFilter(null, old, signal()))
    expect((await pending).kind).toBe('stale')
    expect(f.publications).toBe(0)
  })
  it('keeps untouched confirmation unchanged and normalizes final-column Clear to null', async () => {
    const f = controlsFixture(5)
    expect((await f.client.filter(f.snapshot.owner, null, signal())).kind).toBe('unchanged')
    expect(f.publications).toBe(0)
    const applied = { range: f.range, columns: new Map([[0, { kind: 'number' as const, op: 'gt' as const, first: 1 }]]) }
    expect(filterWithColumn(applied, f.range, 0, null)).toBeNull()
  })
})

import { useState } from 'react'
import { act, fireEvent, render, waitFor } from '@testing-library/react'
import { vi } from 'vitest'
import { FilterDropdown } from '../src/components/xlsx/FilterDropdown'
import { ViewerStatus } from '../src/components/xlsx/ViewerStatus'
import type { ControlsClient, ControlSnapshot } from '../src/components/xlsx/controlsClient'

function FilterHarness({ fixture, client = fixture.client }: { fixture: ReturnType<typeof controlsFixture>; client?: ControlsClient }) {
  const [open, setOpen] = useState(true), [snapshot, setSnapshot] = useState<ControlSnapshot>(fixture.snapshot), [notice, setNotice] = useState('')
  return <><button onClick={() => setOpen(true)}>Open filter</button><FilterDropdown open={open} owner={snapshot.owner} client={client} col={0} applied={snapshot.filter}
    onClose={() => setOpen(false)} onNotice={setNotice} onApplied={setSnapshot} /><ViewerStatus records={snapshot.records} announcement={notice} /></>
}

describe('S9 filter dropdown draft UI', () => {
  it('searches beyond 10k, selects all matches, excludes one and commits only on OK with bounded DOM', async () => {
    const f = controlsFixture(), ui = render(<FilterHarness fixture={f} />)
    await waitFor(() => expect(ui.getByRole('textbox', { name: 'Search all values' })).toBeTruthy())
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    expect(ui.getByText(/Showing first 10,000/)).toBeTruthy()
    expect(ui.container.querySelectorAll('[data-value-ordinal]').length).toBeLessThanOrEqual(200)
    expect(ui.container.querySelector('img')).toBeNull()
    fireEvent.change(ui.getByRole('textbox', { name: 'Search all values' }), { target: { value: 'beyond' } })
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beyond-B' })).toBeTruthy())
    fireEvent.click(ui.getByRole('button', { name: 'Select All' }))
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.click(ui.getByRole('checkbox', { name: 'beyond-B' }))
    expect(f.publications).toBe(0)
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(ui.queryByRole('dialog')).toBeNull())
    expect(f.publications).toBe(1)
    expect(f.snapshot.filter?.columns.get(0)?.kind).toBe('values')
    expect(ui.getByRole('status').textContent).toContain('Filter applied')
  })
  it('Cancel/Escape discard edits and reopening shows the applied selection', async () => {
    const f = controlsFixture(10), ui = render(<FilterHarness fixture={f} />)
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
    fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
    fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
    expect(f.publications).toBe(0)
    fireEvent.click(ui.getByRole('button', { name: 'Open filter' }))
    await waitFor(() => expect((ui.getByRole('checkbox', { name: 'item-3' }) as HTMLInputElement).checked).toBe(true))
    fireEvent.keyDown(ui.getByRole('textbox', { name: 'Search all values' }), { key: 'Escape' })
    expect(ui.queryByRole('dialog')).toBeNull()
    expect(f.publications).toBe(0)
    expect(f.runs.refCount).toBe(0)
  })
  it('drops old generation results and returns Escape focus without trapping Tab', async () => {
    const f = controlsFixture(10), close = vi.fn(), notice = vi.fn(), applied = vi.fn()
    const invoker = document.createElement('button'); document.body.append(invoker)
    const ui = render(<FilterDropdown owner={f.snapshot.owner} client={f.client} col={0} applied={null} onClose={close} onNotice={notice} onApplied={applied} invoker={invoker} />)
    await waitFor(() => expect(ui.getByRole('textbox', { name: 'Search all values' })).toBeTruthy())
    const tab = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true })
    act(() => { ui.getByRole('textbox', { name: 'Search all values' }).dispatchEvent(tab) })
    expect(tab.defaultPrevented).toBe(false)
    fireEvent.keyDown(ui.getByRole('textbox', { name: 'Search all values' }), { key: 'Escape' })
    expect(document.activeElement).toBe(invoker)
    expect(applied).not.toHaveBeenCalled()
    invoker.remove()
  })
})

describe('S9 typed builders and guarded transactions', () => {
  it('discards custom builder edits on Cancel/reopen and supports AND/OR without publishing drafts', async () => {
    const f = controlsFixture(6), ui = render(<FilterHarness fixture={f} />)
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
    fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
    fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'contains' } })
    fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: 'item' } })
    fireEvent.click(ui.getByRole('checkbox', { name: 'Use two conditions' }))
    fireEvent.change(ui.getByRole('combobox', { name: 'Combine conditions' }), { target: { value: 'or' } })
    fireEvent.change(ui.getByRole('combobox', { name: 'Condition 2 type' }), { target: { value: 'number' } })
    fireEvent.change(ui.getByRole('textbox', { name: 'Condition 2 first' }), { target: { value: '0' } })
    expect(f.publications).toBe(0)
    fireEvent.click(ui.getByRole('button', { name: 'Cancel' })); fireEvent.click(ui.getByRole('button', { name: 'Open filter' }))
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
    expect((ui.getByRole('combobox', { name: 'Filter method' }) as HTMLSelectElement).value).toBe('values')
  })
  it('keeps date presets symbolic and previews the resolved calendar period/zone', async () => {
    const f = controlsFixture(4, { kind: 'date', hasNumeric: true }), ui = render(<FilterHarness fixture={f} />)
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
    fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 type' }), { target: { value: 'datePreset' } })
    expect(ui.getByText(/2026-10-11.*UTC/)).toBeTruthy()
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(f.publications).toBe(1))
    expect(f.snapshot.filter?.columns.get(0)).toEqual({ kind: 'datePreset', preset: 'today' })
    expect(f.snapshot.records.visible).toBe(2)
  })
  it('reports invalid reversed bounds and does not apply wrong-type/error/blank conditions', async () => {
    const f = controlsFixture(5, { kind: 'number', hasNumeric: true }), ui = render(<FilterHarness fixture={f} />)
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
    fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'between' } })
    fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: '0' } })
    fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 second' }), { target: { value: '-1' } })
    expect(ui.getByText(/reversed numeric bounds/)).toBeTruthy()
    expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(true)
    expect(f.publications).toBe(0)
  })
  it('never commits late operation results after an owner change and does not overwrite the retry notice', async () => {
    const f = controlsFixture(6), applied = vi.fn(), notice = vi.fn(), close = vi.fn()
    let finish!: (stage: Awaited<ReturnType<typeof f.options.stageFilter>>) => void
    const client = createControlsClient({ ...f.options, stageFilter: () => new Promise(resolve => { finish = resolve }) })
    const props = { owner: f.snapshot.owner, client, col: 0, applied: null, onApplied: applied, onNotice: notice, onClose: close }
    const ui = render(<FilterDropdown {...props} />)
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
    fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
    await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
    fireEvent.click(ui.getByRole('button', { name: 'OK' }))
    await waitFor(() => expect(finish).toBeTypeOf('function'))
    f.snapshot = { ...f.snapshot, owner: { ...f.snapshot.owner, sheetId: 'replacement' } }
    ui.rerender(<FilterDropdown {...props} owner={f.snapshot.owner} />)
    finish({ kind: 'ready', publish: () => { throw new Error('Stale operation published') } })
    await act(async () => { await Promise.resolve() })
    expect(applied).not.toHaveBeenCalled()
    expect(f.publications).toBe(0)
    expect(notice.mock.calls.at(-1)?.[0]).toBe('Source or view changed; reopen and retry.')
    expect(f.runs.refCount).toBe(0)
  })
})

it('surfaces approved load reports and authoritative body counts through the sole ViewerStatus', () => {
  const messages = [
    'This file is too large to open from the web in this browser. Download it and open it locally.',
    'Showing saved values — 12 formulas not recalculated (file too large)',
    'This sheet is too large to export as one image. Choose a range or print selected pages.',
  ]
  const ui = render(<ViewerStatus records={{ visible: 3, total: 10, completeness: 'complete' }} countMode="body" loadReport={{ status: 'ready', notices: messages }} />)
  expect(ui.getAllByRole('status')).toHaveLength(1)
  for (const message of messages) expect(ui.getByRole('status').textContent).toContain(message)
  expect(ui.getByRole('status').textContent).toContain('Visible body 3 / complete body 10')
})

it('applies a typed OR builder atomically and replaces the old range with a notice', async () => {
  const f = controlsFixture(6)
  f.snapshot = { ...f.snapshot, filter: { range: { ...f.range, r1: 2 }, columns: new Map([[1, { kind: 'number', op: 'gt', first: 0 }]]) } }
  const ui = render(<FilterHarness fixture={f} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'contains' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: 'item' } })
  fireEvent.click(ui.getByRole('checkbox', { name: 'Use two conditions' }))
  fireEvent.change(ui.getByRole('combobox', { name: 'Combine conditions' }), { target: { value: 'or' } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 2 type' }), { target: { value: 'number' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 2 first' }), { target: { value: '0' } })
  expect(f.publications).toBe(0)
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.filter?.range).toEqual(f.range)
  expect(f.snapshot.filter?.columns.size).toBe(1)
  expect(f.snapshot.filter?.columns.get(0)).toMatchObject({ kind: 'compound', connective: 'or', first: { kind: 'text', op: 'contains' }, second: { kind: 'number', op: 'eq', first: 0 } })
  expect(ui.getByRole('status').textContent).toContain('Previous filter range replaced')
})

it('Clear removes the final column rule, restores authoritative counts and notes snapshot-hidden rows', async () => {
  const f = controlsFixture(6)
  f.snapshot = { ...f.snapshot, filter: { range: f.range, columns: new Map([[0, { kind: 'number', op: 'gt', first: 0 }]]) }, records: { visible: 0, total: 6, completeness: 'complete' } }
  const notice = vi.fn(), applied = vi.fn()
  const ui = render(<FilterDropdown owner={f.snapshot.owner} client={f.client} col={0} applied={f.snapshot.filter} snapshotHidden onClose={() => {}} onNotice={notice} onApplied={applied} />)
  await waitFor(() => expect(ui.getByRole('button', { name: 'Clear column' }).hasAttribute('disabled')).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'Clear column' }))
  await waitFor(() => expect(applied).toHaveBeenCalledOnce())
  expect(f.snapshot.filter).toBeNull()
  expect(f.snapshot.records).toEqual({ visible: 6, total: 6, completeness: 'complete' })
  expect(notice).toHaveBeenCalledWith("Filter cleared. File's saved hidden rows remain hidden.")
})

it('shows incomplete progress, disables Apply, permits cancel and releases a refused candidate', async () => {
  const f = controlsFixture(6), client = createControlsClient({ ...f.options, stageFilter: async () => ({ kind: 'refused', refusal: { reason: 'merge', range: f.range } }) })
  const ui = render(<FilterHarness fixture={f} client={client} />)
  expect(ui.getByRole('progressbar')).toBeTruthy()
  expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(true)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
  await waitFor(() => expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Operation refused: merge'))
  expect(f.publications).toBe(0)
  expect(f.snapshot.filter).toBeNull()
  fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  expect(f.runs.refCount).toBe(0)
})

it('search Enter searches without applying and native checkbox/negative typing stays isolated', async () => {
  const f = controlsFixture(6), ui = render(<FilterHarness fixture={f} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  const input = ui.getByRole('textbox', { name: 'Search all values' })
  fireEvent.change(input, { target: { value: 'item' } }); fireEvent.keyDown(input, { key: 'Enter' })
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  expect(f.publications).toBe(0)
})

it.each(['and', 'or'] as const)('commits %s compound predicates and excludes errors/blanks through typed conditions', async connective => {
  const f = controlsFixture(6), ui = render(<FilterHarness fixture={f} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'contains' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: 'item' } })
  fireEvent.click(ui.getByRole('checkbox', { name: 'Use two conditions' }))
  fireEvent.change(ui.getByRole('combobox', { name: 'Combine conditions' }), { target: { value: connective } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 2 operator' }), { target: { value: 'neq' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 2 first' }), { target: { value: 'never' } })
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.filter?.columns.get(0)).toMatchObject({ kind: 'compound', connective })
})

it('renders pending membership as indeterminate and never enables Apply for incomplete selections', async () => {
  const f = controlsFixture(6), base = f.client
  const client: ControlsClient = { ...base, async values(...args) {
    const real = await base.values(...args)
    return { ...real, async read(...input) { const page = await real.read(...input); return { ...page, entries: page.entries.map((entry, index) => ({ ...entry, selected: index === 0 ? 'pending' as const : entry.selected })), count: { count: 0, pending: 1, complete: false }, complete: false } } }
  } }
  const ui = render(<FilterHarness fixture={f} client={client} />)
  await waitFor(() => expect(ui.getByText(/selection pending/)).toBeTruthy())
  const pending = ui.container.querySelector<HTMLInputElement>('[data-value-ordinal] input')!
  expect(pending.indeterminate).toBe(true)
  expect(pending.disabled).toBe(true)
  expect(ui.getByRole('button', { name: 'OK' }).hasAttribute('disabled')).toBe(true)
  fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
  expect(f.publications).toBe(0)
})

it('treats a canonical-empty search as no search for whole-domain Select All, including blanks', async () => {
  const f = controlsFixture(6), owner = f.snapshot.owner, draft = await f.client.values(owner, f.range, 0, null, signal())
  await draft.read({ text: '', offset: 0 }, signal())
  await draft.change({ kind: 'all', selected: false, includeBlanks: false }, { text: '', offset: 0 }, signal())
  const page = await draft.change({ kind: 'all', selected: true, includeBlanks: false }, { text: '\u00AD', offset: 0 }, signal())
  expect(page.blanks).toBe(true)
  expect(await draft.candidate(signal())).toBeNull()
  draft.cancel()
})
