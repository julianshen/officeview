import { describe, expect, it } from 'vitest'
import { controlsFixture } from './helpers/xlsx-controls-fixture'
import { createControlsClient, filterWithColumn } from '../src/components/xlsx/controlsClient'
import { compileValuesMembership } from '../src/xlsx/view/valueIndex'
import { ValueIndexError } from '../src/xlsx/view/valueRuns'

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
    await waitFor(() => expect(ui.getByRole('checkbox', { name: 'beyond-B' })).toBeTruthy(), { timeout: 10_000 })
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
  }, 20_000)
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

function holdDraftCalls(base: ControlsClient, gate: { hold: boolean; release: (() => void)[] }): ControlsClient {
  return { ...base, async values(...args) {
    const port = await base.values(...args)
    const wait = () => gate.hold ? new Promise<void>(resolve => gate.release.push(resolve)) : Promise.resolve()
    return { ...port,
      async read(...input) { await wait(); return port.read(...input) },
      async change(...input) { await wait(); return port.change(...input) },
    }
  } }
}
async function releaseDraft(gate: { hold: boolean; release: (() => void)[] }) {
  gate.hold = false
  await act(async () => { for (const release of gate.release.splice(0)) release() })
}

it('disables Inspect during a read so inspection cannot abort the read or leave the dropdown busy', async () => {
  const f = controlsFixture(10), gate = { hold: false, release: [] as (() => void)[] }
  const ui = render(<FilterHarness fixture={f} client={holdDraftCalls(f.client, gate)} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  gate.hold = true
  fireEvent.change(ui.getByRole('textbox', { name: 'Search all values' }), { target: { value: 'item' } })
  const inspect = ui.getByRole('button', { name: 'Inspect value 4' }) as HTMLButtonElement
  expect(inspect.disabled).toBe(true)
  fireEvent.click(inspect)
  await releaseDraft(gate)
  await waitFor(() => expect(ui.queryByRole('progressbar')).toBeNull())
  expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false)
})

it('disables Inspect during OK without cancelling the atomic apply', async () => {
  const f = controlsFixture(10)
  let finish!: (stage: Awaited<ReturnType<typeof f.options.stageFilter>>) => void
  const client = createControlsClient({ ...f.options, stageFilter: () => new Promise(resolve => { finish = resolve }) })
  const ui = render(<FilterHarness fixture={f} client={client} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
  await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(finish).toBeTypeOf('function'))
  const inspect = ui.getByRole('button', { name: 'Inspect value 4' }) as HTMLButtonElement
  expect(inspect.disabled).toBe(true)
  fireEvent.click(inspect)
  await act(async () => finish(await f.options.stageFilter({ range: f.range, columns: new Map([[0, { kind: 'text', op: 'contains', first: 'item', wildcards: false }]]) }, f.snapshot.owner, new AbortController().signal)))
  await waitFor(() => expect(f.publications).toBe(1))
})

it.each(['search', 'wildcards', 'include blanks'])('protects an in-flight checkbox mutation from %s edits', async input => {
  const f = controlsFixture(10), gate = { hold: false, release: [] as (() => void)[] }
  const ui = render(<FilterHarness fixture={f} client={holdDraftCalls(f.client, gate)} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  const search = ui.getByRole('textbox', { name: 'Search all values' }) as HTMLInputElement
  fireEvent.change(search, { target: { value: 'item' } })
  await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  gate.hold = true
  fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
  const control = input === 'search' ? search : ui.getByRole('checkbox', { name: input === 'wildcards' ? 'Search with wildcards' : 'Include blanks in searched Select All' }) as HTMLInputElement
  expect(control.disabled).toBe(true)
  if (input === 'search') fireEvent.change(control, { target: { value: 'beyond' } })
  else fireEvent.click(control)
  await releaseDraft(gate)
  await waitFor(() => expect(ui.queryByRole('progressbar')).toBeNull())
  expect((ui.getByRole('checkbox', { name: 'item-3' }) as HTMLInputElement).checked).toBe(false)
  expect(search.value).toBe('item')
  expect(f.publications).toBe(0)
})

it('permits a valid custom predicate when the values universe is incomplete', async () => {
  const f = controlsFixture(10), base = f.client
  const client: ControlsClient = { ...base, async values(...args) {
    const port = await base.values(...args)
    return { ...port, async read(...input) { const page = await port.read(...input); return { ...page, complete: false, count: { ...page.count, complete: false } } } }
  } }
  const ui = render(<FilterHarness fixture={f} client={client} />)
  await waitFor(() => expect(ui.getByText(/universe incomplete/)).toBeTruthy())
  expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true)
  fireEvent.change(ui.getByRole('combobox', { name: 'Filter method' }), { target: { value: 'custom' } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'contains' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: 'item' } })
  expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false)
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.filter?.columns.get(0)).toEqual({ kind: 'text', op: 'contains', first: 'item', wildcards: false })
})

it('keeps applied sort case sensitivity when using the dropdown sort shortcut', async () => {
  const f = controlsFixture(10)
  f.snapshot = { ...f.snapshot, sort: { range: f.range, keys: [{ col: 1, order: 'desc' }], caseSensitive: true } }
  const props = { owner: f.snapshot.owner, client: f.client, col: 0, applied: null, appliedSort: f.snapshot.sort, onClose: vi.fn(), onNotice: vi.fn(), onApplied: vi.fn() }
  const ui = render(<FilterDropdown {...props} />)
  await waitFor(() => expect((ui.getByRole('button', { name: 'Sort ascending' }) as HTMLButtonElement).disabled).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'Sort ascending' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.sort?.caseSensitive).toBe(true)
  expect(f.snapshot.sort?.keys).toEqual([{ col: 0, order: 'asc' }])
})

it.each(['coarse pointer', 'narrow viewport'])('opens a full-width bottom sheet on a %s and returns focus on Cancel', async mode => {
  const previous = window.matchMedia
  window.matchMedia = vi.fn(query => ({ matches: query.includes(mode === 'coarse pointer' ? 'pointer: coarse' : 'max-width: 640px'), media: query, onchange: null,
    addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true }))
  const f = controlsFixture(10), invoker = document.createElement('button'), close = vi.fn()
  document.body.append(invoker)
  try {
    const ui = render(<FilterDropdown owner={f.snapshot.owner} client={f.client} col={0} applied={null} invoker={invoker} onClose={close} onApplied={vi.fn()} onNotice={vi.fn()} />)
    const popup = ui.getByRole('dialog', { name: 'Filter column A' })
    expect(popup.style.position).toBe('fixed')
    expect(popup.style.bottom).toBe('0px')
    expect(popup.style.left).toBe('0px')
    expect(popup.style.width).toBe('100%')
    expect(popup.style.getPropertyValue('--officeview-safe-bottom')).toContain('safe-area-inset-bottom')
    expect(popup.style.paddingBottom).toContain('--officeview-safe-bottom')
    fireEvent.click(ui.getByRole('button', { name: 'Cancel' }))
    expect(close).toHaveBeenCalledOnce()
    expect(document.activeElement).toBe(invoker)
    ui.unmount()
  } finally { window.matchMedia = previous; invoker.remove() }
})

it('Deselect All then searched Select All and one exclusion commits only the remaining search match', async () => {
  const f = controlsFixture(), ui = render(<FilterHarness fixture={f} />)
  const ready = () => waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  await ready()
  fireEvent.click(ui.getByRole('button', { name: 'Deselect All' })); await ready()
  fireEvent.change(ui.getByRole('textbox', { name: 'Search all values' }), { target: { value: 'beyond' } }); await ready()
  fireEvent.click(ui.getByRole('button', { name: 'Select All' })); await ready()
  fireEvent.click(ui.getByRole('checkbox', { name: 'beyond-B' })); await ready()
  expect(f.publications).toBe(0)
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  const rule = f.snapshot.filter!.columns.get(0)!
  expect(rule.kind).toBe('values')
  if (rule.kind !== 'values') throw new Error('Expected values membership')
  const membership = await compileValuesMembership(rule, f.options.domain(f.snapshot.owner, f.range, 0).identity, f.runs)
  expect(membership.matches({ kind: 'string', value: 'beyond-A' })).toBe(true)
  for (const value of ['beyond-B', 'item-10001', 'item-5']) expect(membership.matches({ kind: 'string', value })).toBe(false)
  expect(membership.matches({ kind: 'blank' })).toBe(false)
  expect(membership.matches({ kind: 'error', value: '#N/A' })).toBe(false)
  expect(f.snapshot.records).toEqual({ visible: 1, total: 11020, completeness: 'complete' })
})

it('drops a delayed values page after valueGen changes and retains the retry notice', async () => {
  const f = controlsFixture(10), base = f.client, notice = vi.fn(), close = vi.fn(), applied = vi.fn()
  let hold = false, release!: () => void
  const client: ControlsClient = { ...base, async values(...args) {
    const port = await base.values(...args)
    let last!: Awaited<ReturnType<typeof port.read>>
    return { ...port, async read(...input) {
      if (!hold) return last = await port.read(...input)
      return new Promise(resolve => { release = () => resolve({ ...last, entries: [{ ...last.entries[0], label: 'OLD-PAGE' }] }) })
    } }
  } }
  const props = { owner: f.snapshot.owner, client, col: 0, applied: null, onNotice: notice, onClose: close, onApplied: applied }
  const ui = render(<FilterDropdown {...props} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  hold = true
  fireEvent.change(ui.getByRole('textbox', { name: 'Search all values' }), { target: { value: 'item' } })
  expect(release).toBeTypeOf('function')
  f.snapshot = { ...f.snapshot, owner: { ...f.snapshot.owner, valueGen: f.snapshot.owner.valueGen + 1 } }
  ui.rerender(<FilterDropdown {...props} owner={f.snapshot.owner} />)
  await act(async () => release())
  expect(ui.queryByText('OLD-PAGE')).toBeNull()
  expect(close).toHaveBeenCalledOnce()
  expect(applied).not.toHaveBeenCalled()
  expect(notice.mock.calls.at(-1)?.[0]).toBe('Source or view changed; reopen and retry.')
  expect(f.publications).toBe(0)
})

it('respects no-header table metadata in the filter dropdown and includes the first row in body counts', async () => {
  const f = controlsFixture(6, { kind: 'text', hasNumeric: true }, true), ui = render(<FilterHarness fixture={f} />)
  await waitFor(() => expect(ui.getByRole('checkbox', { name: 'item-3' })).toBeTruthy())
  expect(ui.getByText(/Table metadata: 0 header row/)).toBeTruthy()
  fireEvent.click(ui.getByRole('checkbox', { name: 'item-3' }))
  await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.filter?.range).toEqual(f.range)
  expect(f.snapshot.filter?.range.headerRows).toBe(0)
  expect(f.snapshot.records).toEqual({ visible: 5, total: 6, completeness: 'complete' })
})

it('keeps custom Apply disabled while classification is pending instead of clearing the applied rule', async () => {
  const f = controlsFixture(10), base = f.client
  f.snapshot = { ...f.snapshot, filter: { range: f.range, columns: new Map([[0, { kind: 'text', op: 'contains', first: 'item', wildcards: false }]]) } }
  const client: ControlsClient = { ...base, async values(...args) {
    const port = await base.values(...args)
    return { ...port, identity: { ...port.identity, classification: { kind: 'loading', hasNumeric: false } },
      async read(...input) { const page = await port.read(...input); return { ...page, complete: false } },
    }
  } }
  const ui = render(<FilterHarness fixture={f} client={client} />)
  await waitFor(() => expect(ui.queryByRole('progressbar')).toBeNull())
  expect((ui.getByRole('textbox', { name: 'Condition 1 first' }) as HTMLInputElement).value).toBe('item')
  expect((ui.getByRole('combobox', { name: 'Filter method' }) as HTMLSelectElement).value).toBe('custom')
  expect(ui.getByText('Column classification is pending; wait and retry.')).toBeTruthy()
  expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(true)
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  expect(f.publications).toBe(0)
  expect(f.snapshot.filter?.columns.get(0)).toEqual({ kind: 'text', op: 'contains', first: 'item', wildcards: false })
})

it('offers typed conditions from the resolved domain when values() fails on quota, and applies a typed rule', async () => {
  const f = controlsFixture(10), base = f.client
  const client: ControlsClient = { ...base, async values() { throw new ValueIndexError('quota', 'Value index quota exceeded', { resource: 'memory' }) } }
  const ui = render(<FilterHarness fixture={f} client={client} />)
  await waitFor(() => expect(ui.queryByRole('progressbar')).toBeNull())
  expect(ui.getAllByText(/quota/i).length).toBeGreaterThan(0)
  const method = ui.getByRole('combobox', { name: 'Filter method' }) as HTMLSelectElement
  expect([...method.options].map(option => option.textContent)).toContain('Text / mixed filters')
  fireEvent.change(method, { target: { value: 'custom' } })
  fireEvent.change(ui.getByRole('combobox', { name: 'Condition 1 operator' }), { target: { value: 'contains' } })
  fireEvent.change(ui.getByRole('textbox', { name: 'Condition 1 first' }), { target: { value: 'item' } })
  await waitFor(() => expect((ui.getByRole('button', { name: 'OK' }) as HTMLButtonElement).disabled).toBe(false))
  fireEvent.click(ui.getByRole('button', { name: 'OK' }))
  await waitFor(() => expect(f.publications).toBe(1))
  expect(f.snapshot.filter?.columns.get(0)).toMatchObject({ kind: 'text', op: 'contains', first: 'item' })
})
