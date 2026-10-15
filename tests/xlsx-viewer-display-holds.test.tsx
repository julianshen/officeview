/**
 * Spec §§10.1, 12.4 in the mounted viewer: a worker view stays charged while
 * any real owner still holds it — a committed display (this viewer's or a
 * sibling's), a copy continuation that ignores cancellation, or a display that
 * a newer, not-yet-committed render would replace. Collection callbacks never
 * run in this file, so only explicit releases end a charge.
 */
import { afterEach, beforeEach, expect, it } from 'vitest'
import { Suspense, StrictMode, startTransition, useState, type JSX } from 'react'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import type { SortSpec } from '../src/xlsx/view/contracts'
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
async function opened(docId: string) {
  const doc = docOf(), source = createObjectWorkbookSource(doc, docId)
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId, source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId, valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session(docId)!.layout(sheet.sheetId)])) })
  const sort = async (order: 'asc' | 'desc') => {
    const spec: SortSpec = { range: { r0: 0, r1: 3, c0: 0, c1: 0, headerRows: 1, totalsRows: 0 }, keys: [{ col: 0, order }], caseSensitive: false }
    let out!: Awaited<ReturnType<RemoteWorkbook['sort']>>
    await act(async () => { out = await remote.sort(remote.identity('A'), spec, new AbortController().signal) })
    expect(out.kind).toBe('committed')
    return remote.view('A')
  }
  return { doc, remote, sort }
}
const loaded = (ui: { getByRole: ReturnType<typeof render>['getByRole'] }) => waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))

it('a pending copy and the committed display keep their view charged after the workbook moves on; both end with no double release', async () => {
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: async () => {} } })
  const { doc, remote, sort } = await opened('copy-hold')
  try {
    const ui = render(<InteractiveSheetDoc doc={doc} source={remote} />)
    await loaded(ui)
    const first = await sort('asc')
    fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A2' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    const one = remote.mainLedger.snapshot().byPool.viewBuffers
    const store = remote.store('A'), read = store.readCell.bind(store), pending: (() => void)[] = []
    store.readCell = async (...args) => { await new Promise<void>(resolve => pending.push(resolve)); return read(...args) }
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(pending.length).toBe(1))
    // The workbook adopts revision 2 before this viewer renders it: the copy and
    // the committed display still read revision 1's order.
    const notify = remote.notify.bind(remote)
    remote.notify = () => {}
    const second = await sort('desc')
    expect(remote.mainLedger.hasLifetime(first.order), 'held by the display and the pending copy').toBe(true)
    expect(remote.mainLedger.snapshot().byPool.viewBuffers).toBe(2 * one)
    // Revision 2 commits: the display moves on and the copy (cancelled by the
    // owner change) drains; each releases its own hold exactly once.
    remote.notify = notify
    await act(async () => { notify() })
    await act(async () => { pending.splice(0).forEach(resume => resume()); await new Promise(resolve => setTimeout(resolve, 0)) })
    await waitFor(() => expect(remote.mainLedger.hasLifetime(first.order)).toBe(false))
    expect(remote.mainLedger.hasLifetime(second.order)).toBe(true)
    expect(remote.mainLedger.snapshot().byPool.viewBuffers).toBe(one)
    ui.unmount()
    expect(remote.mainLedger.snapshot().byPool.viewBuffers, 'the workbook still holds its current view').toBe(one)
  } finally { await remote.close() }
})

it('a render that is never committed refunds nothing; the committed display keeps its view until a newer one commits', async () => {
  const { doc, remote, sort } = await opened('uncommitted')
  let suspend!: (value: boolean) => void
  const never = new Promise<never>(() => {})
  function Gate({ on }: { on: boolean }): JSX.Element | null { if (on) throw never; return null }
  function Host(): JSX.Element {
    const [on, setOn] = useState(false)
    suspend = setOn
    return <Suspense fallback={<p>fallback</p>}><InteractiveSheetDoc doc={doc} source={remote} /><Gate on={on} /></Suspense>
  }
  try {
    const ui = render(<Host />)
    await loaded(ui)
    const first = await sort('asc')
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
    expect(remote.mainLedger.hasLifetime(first.order)).toBe(true)
    // A newer view arrives without this viewer hearing of it yet.
    const notify = remote.notify.bind(remote)
    remote.notify = () => {}
    const second = await sort('desc')
    expect(remote.mainLedger.hasLifetime(first.order), 'the committed display still shows revision 1').toBe(true)
    // A transition re-renders the viewer (building revision 2's entry) but suspends: nothing commits.
    await act(async () => { startTransition(() => suspend(true)) })
    expect(ui.queryByText('fallback'), 'the old display stays on screen').toBeNull()
    expect(remote.mainLedger.hasLifetime(first.order), 'an uncommitted render must not refund the displayed view').toBe(true)
    // The transition is abandoned and a real display of revision 2 commits.
    remote.notify = notify
    await act(async () => { suspend(false); notify() })
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
    expect(remote.mainLedger.hasLifetime(first.order), 'released once revision 2 is committed').toBe(false)
    expect(remote.mainLedger.hasLifetime(second.order)).toBe(true)
    ui.unmount()
  } finally { await remote.close() }
})

it('sibling viewers: the older view stays charged until every viewer has committed past it; StrictMode remounts hold again', async () => {
  const { doc, remote, sort } = await opened('siblings')
  try {
    const a = render(<StrictMode><InteractiveSheetDoc doc={doc} source={remote} /></StrictMode>)
    const holder = document.body.appendChild(document.createElement('div'))
    const b = render(<InteractiveSheetDoc doc={doc} source={remote} />, { container: holder })
    await loaded(within(a.container)); await loaded(within(holder))
    const first = await sort('asc')
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
    // Revision 2 arrives; only A re-renders (an unrelated update) and commits it.
    const notify = remote.notify.bind(remote)
    remote.notify = () => {}
    const second = await sort('desc')
    await act(async () => {
      fireEvent.change(within(a.container).getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A3' } })
      fireEvent.click(within(a.container).getByRole('button', { name: 'Go To' }))
    })
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
    expect(remote.mainLedger.hasLifetime(first.order), 'viewer B still displays revision 1').toBe(true)
    expect(remote.mainLedger.hasLifetime(second.order)).toBe(true)
    b.unmount()
    expect(remote.mainLedger.hasLifetime(first.order), 'released once its last display is gone').toBe(false)
    expect(remote.mainLedger.hasLifetime(second.order), 'A (after StrictMode effect cycles) and the workbook hold revision 2').toBe(true)
    remote.notify = notify
    a.unmount()
    await remote.close()
    expect(remote.mainLedger.snapshot().byPool.viewBuffers).toBe(0)
  } finally { await remote.close() }
})

it('C3 geometry: replaced paged projections end with their last display, so repeated sorts keep the geometry pool flat', async () => {
  const { doc, remote, sort } = await opened('paged-geometry')
  try {
    const ui = render(<InteractiveSheetDoc doc={doc} source={remote} />)
    await loaded(ui)
    const pools: number[] = []
    for (let i = 0; i < 6; i++) {
      await sort(i % 2 ? 'desc' : 'asc')
      await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
      pools.push(remote.mainLedger.snapshot().byPool.geometry)
    }
    expect(pools[0]).toBeGreaterThan(0)
    for (const bytes of pools) expect(bytes, pools.join(' ')).toBe(pools[0])
    ui.unmount()
  } finally { await remote.close() }
})

it('C3 resident: repeated header sorts keep resident geometry, view buffers and the sorted views\' masks from accumulating', async () => {
  const doc = docOf()
  let ledger!: import('../src/xlsx/view/mainLedger').MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await loaded(ui)
  const header = () => ui.getByRole('columnheader', { name: /^Column A/ })
  const pools: { geometry: number; viewBuffers: number; controls: number }[] = []
  for (let i = 0; i < 6; i++) {
    const before = header().getAttribute('aria-sort')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column A' }))
    await waitFor(() => expect(header().getAttribute('aria-sort')).not.toBe(before))
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)) })
    const { geometry, viewBuffers, controls } = ledger.snapshot().byPool
    pools.push({ geometry, viewBuffers, controls })
  }
  // ascending, descending, clear: the same point of each cycle costs the same.
  for (let i = 3; i < pools.length; i++) {
    expect({ geometry: pools[i].geometry, viewBuffers: pools[i].viewBuffers }, JSON.stringify(pools)).toEqual({ geometry: pools[i - 3].geometry, viewBuffers: pools[i - 3].viewBuffers })
    // Each resident sort publishes a full-axis display mask (~132 KiB): replaced
    // ones end with their display. Only small per-operation records (kept for a
    // provider that may ignore cancellation) remain until collection.
    expect(pools[i].controls - pools[i - 3].controls, JSON.stringify(pools)).toBeLessThan(64 * 1024)
  }
  ui.unmount()
})
