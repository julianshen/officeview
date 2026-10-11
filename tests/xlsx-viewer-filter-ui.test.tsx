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
