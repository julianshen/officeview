/**
 * AC-LIFE-02 (spec §12.4): repeated multi-sheet switching keeps main-thread
 * charges bounded by explicit owners, not by garbage collection. Collection
 * callbacks never run in this file, so only explicit releases end a charge.
 */
import { afterEach, beforeEach, expect, it, onTestFinished, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { createViewerControlResources, viewerControls } from '../src/components/xlsx/viewerControlsClient'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { StructuralIndex } from '../src/xlsx/view/structure'
import { emptyGridSelection } from '../src/xlsx/view/selection'
import { workbookOf } from './helpers/xlsx-find'

const RealRegistry = globalThis.FinalizationRegistry
beforeEach(() => {
  globalThis.FinalizationRegistry = class { register() {} unregister() { return false } } as unknown as typeof FinalizationRegistry
})
afterEach(() => { cleanup(); globalThis.FinalizationRegistry = RealRegistry })

const sheetOf = (sheetId: string, rows: number) => ({ sheetId, cells: [
  { row: 0, col: 0, value: `Head ${sheetId}` }, { row: 0, col: 1, value: 'Amount' },
  ...Array.from({ length: rows }, (_, i) => [{ row: i + 1, col: 0, value: `${sheetId}-${i}` }, { row: i + 1, col: 1, value: i * 3 }]).flat(),
] })
const settled = () => act(async () => { for (let i = 0; i < 4; i++) await new Promise(resolve => setTimeout(resolve, 0)) })

it('sheet switching keeps semantic presentation charges bounded by the active sheet cache', async () => {
  const doc = workbookOf([sheetOf('A', 12), sheetOf('B', 9), sheetOf('C', 6)])
  // Lifetime charges moved from a format workspace into the document pool:
  // exactly the semantic presentations (each cell label the grid shows).
  const presentations = new Set<string>()
  let ledger!: MainMemoryLedger
  const observe = (value: MainMemoryLedger) => { ledger = value }
  const adopt = MainMemoryLedger.prototype.adoptWhileReachable
  MainMemoryLedger.prototype.adoptWhileReachable = function (fromPool, fromOwner, pool, ...rest) {
    const token = adopt.call(this, fromPool, fromOwner, pool, ...rest)
    if (fromPool === 'layout' && fromOwner.startsWith('format-') && pool === 'document') presentations.add(token)
    return token
  }
  onTestFinished(() => { MainMemoryLedger.prototype.adoptWhileReachable = adopt })
  const live = () => [...presentations].reduce((sum, token) => sum + ledger.ownerBytes('document', token), 0)
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={observe} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  await settled()
  const choose = async (sheetId: string) => {
    fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: sheetId } })
    await waitFor(() => expect(ui.getByRole('grid').getAttribute('aria-label') ?? ui.getByRole('grid').textContent).toContain(sheetId))
    await settled()
  }
  const first = live()
  expect(first, 'the active sheet formats its visible cells through the cache').toBeGreaterThan(0)
  const perSheet: Record<string, number> = { A: first }
  for (const id of ['B', 'C']) { await choose(id); perSheet[id] = live() }
  // The largest sheet's own cache bounds every later moment of the cycle.
  const bound = Math.max(...Object.values(perSheet))
  const observed: number[] = []
  for (let cycle = 0; cycle < 4; cycle++) for (const id of ['A', 'B', 'C']) {
    await choose(id)
    observed.push(live())
  }
  expect(Math.max(...observed), `presentation charges per switch: ${observed.join(', ')}`).toBeLessThanOrEqual(bound)
  // Returning to A charges exactly what A's own fresh cache holds.
  expect(observed[observed.length - 3]).toBe(first)
  ui.unmount()
  expect(live(), 'unmount releases the active sheet cache').toBe(0)
}, 30000)

it('disposing a control port returns the controls pool to its pre-port value', async () => {
  const doc = workbookOf([sheetOf('A', 12), sheetOf('B', 6)])
  doc.sheets[0].rows[1].cells[1].style = { numFmtId: 164, formatCode: '0.00' }
  doc.sheets[0].rows[1].cells[1].styleIndex = 1
  const source = createObjectWorkbookSource(doc, 'd')
  const ledger = new MainMemoryLedger({ register() {} })
  const resources = createViewerControlResources('D', undefined, ledger)
  const options = (sheetId: string) => ({ doc, source, resources, sheetId, ledger,
    snapshot: () => ({ owner: { ...source.store(sheetId).stamp, fontGen: 0, viewGen: 0 }, filter: null, sort: null, records: { visible: 13, total: 13, completeness: 'complete' as const } }),
    view: () => { throw new Error('unused') }, publish: () => { throw new Error('unused') } })
  const keep = await viewerControls(options('B') as never, new AbortController().signal)
  const before = ledger.snapshot().byPool.controls
  for (let i = 0; i < 3; i++) {
    const port = await viewerControls(options('A') as never, new AbortController().signal)
    expect(ledger.snapshot().byPool.controls, 'a live port charges its lookup metadata').toBeGreaterThan(before)
    port.dispose()
    expect(ledger.snapshot().byPool.controls, `after dispose ${i}`).toBe(before)
  }
  keep.dispose()
  resources.dispose()
})

/** A port over one sheet whose structural indexes (in creation order) are observed. */
async function portWithIndexes(structural?: (call: number, result: { completeness: string }) => Promise<void> | void) {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'Header' }, { row: 1, col: 0, value: 2 }, { row: 2, col: 0, value: 1 }] }])
  const source = createObjectWorkbookSource(doc, 'd'), store = source.store('s')
  const ledger = new MainMemoryLedger({ register() {} }), resources = createViewerControlResources('D', undefined, ledger)
  const indexes: StructuralIndex[] = []
  const adopt = ledger.adoptWhileReachable.bind(ledger)
  ledger.adoptWhileReachable = (...args) => { if (args[3][0] instanceof StructuralIndex) indexes.push(args[3][0]); return adopt(...args) }
  if (structural) {
    const query = store.queryStructural.bind(store)
    let call = 0
    store.queryStructural = async (...args) => { const result = await query(...args); await structural(call++, result as { completeness: string }); return result }
  }
  let release!: () => void, hit!: () => void
  const reading = new Promise<void>(resolve => hit = resolve), held = new Promise<void>(resolve => release = resolve)
  const readRows = store.readRows.bind(store)
  // Every header probe waits until released: a resolver stays in flight.
  store.readRows = async (...args) => { hit(); await held; return readRows(...args) }
  const owner = { ...store.stamp, fontGen: 0, viewGen: 0 }
  const selection = { ...emptyGridSelection(), active: { row: 1, col: 0 }, ranges: [{ rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 2 }] }, cols: [{ c0: 0, c1: 0 }] }] }
  const port = await viewerControls({ doc, source, sheetId: 's', resources, ledger,
    snapshot: () => ({ owner, filter: null, sort: null, records: { visible: 3, total: 3, completeness: 'complete' as const } }),
    view: () => ({ selection }) as never, publish: () => { throw new Error('unused') } }, new AbortController().signal)
  const resolve = () => port.client.resolve(owner, 'sort', new AbortController().signal)
  return { ledger, resources, indexes, port, resolve, reading, release }
}

it('disposing a port defers its structural index release to an in-flight resolver, exactly once', async () => {
  const { ledger, resources, indexes, port, resolve, reading, release } = await portWithIndexes()
  const [index] = indexes
  const charged = () => ledger.snapshot().byPool.controls
  const pending = resolve()
  await reading
  const during = charged()
  port.dispose()
  expect(ledger.hasLifetime(index), 'the resolver still reads the index after its header probe').toBe(true)
  port.dispose()
  expect(ledger.hasLifetime(index)).toBe(true)
  release()
  await pending.catch(() => {})
  expect(ledger.hasLifetime(index), 'the last holder ends the charge').toBe(false)
  const drained = charged()
  expect(drained).toBeLessThan(during)
  port.dispose()
  expect(charged(), 'a later dispose releases nothing twice').toBe(drained)
  resources.dispose()
})

it('re-queried structure (loading, loading, complete) retires each index when its last reader drains', async () => {
  const gates: Array<{ open: () => void; wait: Promise<void> }> = []
  const gateAt = (call: number) => gates[call] ??= (() => { let open!: () => void; const wait = new Promise<void>(resolve => open = resolve); return { open, wait } })()
  const { ledger, resources, indexes, port, resolve, reading, release } = await portWithIndexes(async (call, result) => {
    // 0: the port opens on a loading sheet; 1: still loading; 2 and 3: complete.
    if (call <= 1) result.completeness = 'loading'
    if (call >= 2) await gateAt(call).wait
  })
  expect(indexes).toHaveLength(1)
  // A resolver on a loading structure re-queries, still loading: the replaced
  // index had no other reader and ends at once.
  await expect(resolve()).resolves.toMatchObject({ kind: 'incomplete' })
  expect(indexes).toHaveLength(2)
  expect(ledger.hasLifetime(indexes[0])).toBe(false)
  // Two resolvers re-query concurrently. The first completes and holds its
  // index in a header probe while the second replaces it.
  const first = resolve(), second = resolve()
  gateAt(2).open()
  await reading
  expect(indexes).toHaveLength(3)
  gateAt(3).open()
  await vi.waitFor(() => expect(indexes).toHaveLength(4))
  expect(ledger.hasLifetime(indexes[1])).toBe(false)
  expect(ledger.hasLifetime(indexes[2]), 'a replaced index stays charged while a resolver reads it').toBe(true)
  release()
  await Promise.all([first, second])
  expect(ledger.hasLifetime(indexes[2]), 'its reader drained').toBe(false)
  expect(ledger.hasLifetime(indexes[3]), 'the port still owns the current index').toBe(true)
  port.dispose()
  expect(ledger.hasLifetime(indexes[3])).toBe(false)
  resources.dispose()
})
