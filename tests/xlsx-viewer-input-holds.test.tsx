/**
 * Spec §12.4 retained views (codex ownership review r2, frozen): an external
 * `views` input stays charged while any sheet in the committed props still
 * provides it, whether or not it is displayed. Collection callbacks never run
 * in this file, so only explicit releases end a charge.
 */
import { afterEach, beforeEach, expect, it } from 'vitest'
import { Suspense, startTransition, useState, type JSX } from 'react'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { workbookOf } from './helpers/xlsx-find'

const RealRegistry = globalThis.FinalizationRegistry
beforeEach(() => {
  globalThis.FinalizationRegistry = class { register() {} unregister() { return false } } as unknown as typeof FinalizationRegistry
})
afterEach(() => { cleanup(); globalThis.FinalizationRegistry = RealRegistry })

const docOf = () => workbookOf([{ sheetId: 'A', cells: [
  { row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' },
  { row: 2, col: 0, value: 'alpha' }, { row: 3, col: 0, value: 'gamma' },
] }])
const loaded = (ui: { getByRole: ReturnType<typeof render>['getByRole'] }) => waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))


it('external view prop remains provided and readable after local header sort overrides its revision', async () => {
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const doc = docOf()
  const externalOrder = createRowOrder([{ r0: 1, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 3) }])
  const externalViews = { A: { revision: 1, order: externalOrder } }
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} views={externalViews} onMainLedger={value => { ledger = value }} />)
  await loaded(ui)
  expect(ledger.hasLifetime(externalOrder), 'initial provided order is charged while displayed').toBe(true)
  expect(externalViews.A.order.displayToSource(1)).toBe(2)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  const before = header().getAttribute('aria-sort')
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).not.toBe(before))
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  // The exact same externalViews prop remains supplied throughout this local state update.
  expect(externalViews.A.revision).toBe(1)
  expect(externalViews.A.order).toBe(externalOrder)
  expect(externalViews.A.order.displayToSource(1)).toBe(2)
  expect(ledger.hasLifetime(externalOrder), 'provided-but-locally-overridden external view charge stays').toBe(true)
  ui.unmount()
})


it('same-revision new input wrapper retains the same external allocation through subsequent local sort', async () => {
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const doc = docOf(), order = createRowOrder([{ r0: 1, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 3) }])
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger
  const observe = (value: typeof ledger) => { ledger = value }
  const a = { revision: 1, order }, b = { revision: 1, order }
  const ui = render(<InteractiveSheetDoc doc={doc} views={{ A: a }} onMainLedger={observe} />)
  await loaded(ui)
  expect(ledger.hasLifetime(order)).toBe(true)
  ui.rerender(<InteractiveSheetDoc doc={doc} views={{ A: b }} onMainLedger={observe} />)
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'same allocation is still displayed').toBe(true)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'))
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'new wrapper still in props must retain its order').toBe(true)
  ui.unmount()
})

it('removing then readding an input below local revision reacquires its retained allocation', async () => {
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const doc = docOf(), order = createRowOrder([{ r0: 1, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 3) }])
  const input = { revision: 1, order }
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger
  const observe = (value: typeof ledger) => { ledger = value }
  const ui = render(<InteractiveSheetDoc doc={doc} views={{ A: input }} onMainLedger={observe} />)
  await loaded(ui)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'))
  expect(ledger.hasLifetime(order)).toBe(true)
  ui.rerender(<InteractiveSheetDoc doc={doc} onMainLedger={observe} />)
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'removed input releases').toBe(false)
  ui.rerender(<InteractiveSheetDoc doc={doc} views={{ A: input }} onMainLedger={observe} />)
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'readded input is retained even though local view still overrides it').toBe(true)
  ui.unmount()
})


it('shared input across two sheets stays retained when only one sheet prop is removed', async () => {
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const cells = [{ row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' }, { row: 2, col: 0, value: 'alpha' }, { row: 3, col: 0, value: 'gamma' }]
  const doc = workbookOf([{ sheetId: 'A', cells }, { sheetId: 'B', cells }])
  const order = createRowOrder([{ r0: 1, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 3) }]), input = { revision: 1, order }
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger
  const observe = (value: typeof ledger) => { ledger = value }
  const ui = render(<InteractiveSheetDoc doc={doc} views={{ A: input, B: input }} onMainLedger={observe} />)
  await loaded(ui)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'))
  fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: 'B' } })
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'))
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'shared external input still retained before removing A').toBe(true)
  ui.rerender(<InteractiveSheetDoc doc={doc} views={{ B: input }} onMainLedger={observe} />)
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'B still provides same external input').toBe(true)
  ui.unmount()
})

it('abandoned Suspense prop removal preserves original input ownership after it returns', async () => {
  const { createRowOrder } = await import('../src/xlsx/view/order')
  const doc = docOf(), order = createRowOrder([{ r0: 1, r1: 3, sourceAtOffset: Uint32Array.of(2, 1, 3) }])
  const a = { revision: 1, order }
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger, change!: (value: boolean) => void
  const observe = (value: typeof ledger) => { ledger = value }, never = new Promise<never>(() => {})
  function Gate({ on }: { on: boolean }): JSX.Element | null { if (on) throw never; return null }
  function Host(): JSX.Element {
    const [on, setOn] = useState(false); change = setOn
    return <Suspense fallback={<p>blocked</p>}><InteractiveSheetDoc doc={doc} views={on ? undefined : { A: a }} onMainLedger={observe} /><Gate on={on} /></Suspense>
  }
  const ui = render(<Host />)
  await loaded(ui)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
  await waitFor(() => expect(header().getAttribute('aria-sort')).toBe('ascending'))
  expect(ledger.hasLifetime(order)).toBe(true)
  await act(async () => { startTransition(() => change(true)) })
  expect(ui.queryByText('blocked')).toBeNull()
  expect(ledger.hasLifetime(order), 'old committed input still held during suspension').toBe(true)
  await act(async () => { change(false); await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'abandoning transition alone still preserves old input').toBe(true)
  // An unrelated real committed viewer render drains the speculative retirement.
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A2' } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
  expect(ledger.hasLifetime(order), 'original external prop still provided after suspended replacement abandoned').toBe(true)
  ui.unmount()
})
