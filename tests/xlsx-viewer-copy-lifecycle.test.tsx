import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { workbookOf } from './helpers/xlsx-find'

const docOf = () => workbookOf([{ sheetId: 'A', cells: [
  { row: 0, col: 0, value: 'Name' }, { row: 1, col: 0, value: 'beta' },
  { row: 2, col: 0, value: 'a\t"b' }, { row: 3, col: 0, value: 'longer text' },
] }])
interface ClipboardJob { text: string; resolve(): void; reject(error: Error): void }
const outstanding = new Set<ClipboardJob>()
const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard')
const originalExecCommand = Object.getOwnPropertyDescriptor(document, 'execCommand')
afterEach(async () => {
  cleanup()
  await act(async () => { for (const job of outstanding) job.resolve() })
  outstanding.clear()
  vi.restoreAllMocks()
  if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard)
  else Reflect.deleteProperty(navigator, 'clipboard')
  if (originalExecCommand) Object.defineProperty(document, 'execCommand', originalExecCommand)
  else Reflect.deleteProperty(document, 'execCommand')
})
function clipboard() {
  const jobs: ClipboardJob[] = []
  const writeText = vi.fn((text: string) => new Promise<void>((resolve, reject) => {
    const job = { text, resolve: () => { outstanding.delete(job); resolve() }, reject: (error: Error) => { outstanding.delete(job); reject(error) } }
    jobs.push(job)
    outstanding.add(job)
  }))
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } })
  return { jobs, writeText }
}
function go(ui: ReturnType<typeof render>, at: string) {
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: at } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
}
async function copy(ui: ReturnType<typeof render>, writes: ReturnType<typeof clipboard>, count: number) {
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(writes.writeText).toHaveBeenCalledTimes(count))
}
async function mounted() {
  const doc = docOf()
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A2')
  return { ui, ledger, doc }
}

/**
 * Leaves exactly `capacity` bytes of the main budget for the copy pool from now
 * until stop(). Other ledger owners change asynchronously while a copy yields:
 * a FinalizationRegistry callback ends an unreachable page's lifetime charge, a
 * late render charges tile shaping caches and merges. A pressure owner absorbs
 * every such change, so a budget boundary test admits or refuses by its own
 * bytes alone. `onCopy` sees each copy-pool charge (new and previous pool bytes).
 */
function pinCopyCapacity(ledger: MainMemoryLedger, budget: number, capacity: number, onCopy?: (bytes: number, previous: number) => void) {
  const set = ledger.set.bind(ledger), outside = budget - capacity
  let active = true
  const outsideCopy = () => ledger.bytes - ledger.snapshot().byPool.copy
  const pin = () => set('controls', 'pressure', ledger.ownerBytes('controls', 'pressure') + outside - outsideCopy())
  vi.spyOn(ledger, 'set').mockImplementation((pool, owner, bytes) => {
    const previous = ledger.snapshot().byPool.copy
    set(pool, owner, bytes)
    if (!active) return
    if (pool === 'copy') onCopy?.(bytes, previous)
    else if (owner !== 'pressure') pin()
  })
  pin()
  return {
    /** Whether the pin held: only the copy pool moved since it was set. */
    held: () => outsideCopy() === outside,
    stop: () => { active = false; set('controls', 'pressure', 0) },
  }
}

it('older clipboard settlement cannot release the newer pending copy charge', async () => {
  const writes = clipboard(), { ui, ledger } = await mounted()
  await copy(ui, writes, 1)
  expect(writes.jobs[0].text).toBe('beta')
  expect(ledger.snapshot().byPool.copy).toBe(144)
  fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  await copy(ui, writes, 2)
  const both = ledger.snapshot().byPool.copy
  await act(async () => writes.jobs[0].resolve())
  expect(ui.getByRole('button', { name: 'Cancel copy' })).toBeTruthy()
  expect(ledger.snapshot().byPool.copy).toBe(144)
  expect(both).toBe(288)
  expect(ui.getByRole('status').textContent).not.toContain('Selection copied.')
  await act(async () => writes.jobs[1].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull()
  expect(ui.getByRole('status').textContent).toContain('Selection copied.')
})

it('cancelled native jobs keep distinct TSV charges through out-of-order completion', async () => {
  const writes = clipboard(), { ui, ledger } = await mounted()
  await copy(ui, writes, 1)
  fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  go(ui, 'A3')
  await copy(ui, writes, 2)
  fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  go(ui, 'A4')
  await copy(ui, writes, 3)
  expect(writes.jobs.map(job => job.text)).toEqual(['beta', '"a\t""b"', 'longer text'])
  const charges = writes.jobs.map(job => 128 + job.text.length * 4)
  expect(new Set(charges).size).toBe(3)
  expect(ledger.snapshot().byPool.copy).toBe(charges.reduce((a, b) => a + b))
  await act(async () => writes.jobs[2].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(charges[0] + charges[1])
  const notice = ui.getByRole('status').textContent
  await act(async () => writes.jobs[0].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(charges[1])
  expect(ui.getByRole('status').textContent).toBe(notice)
  await act(async () => writes.jobs[1].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ui.getByRole('status').textContent).toBe(notice)
})

it('a thrown clipboard boundary settles only its charge and reports failure without an unhandled rejection', async () => {
  const writes = clipboard(), { ui, ledger } = await mounted()
  await copy(ui, writes, 1)
  fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  go(ui, 'A3')
  let atThrow = 0
  Object.defineProperty(navigator, 'clipboard', { configurable: true, get: () => {
    atThrow = ledger.snapshot().byPool.copy
    throw new Error('Clipboard unavailable.')
  } })
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(ui.getByRole('status').textContent).toContain('Clipboard unavailable.'))
  expect(atThrow).toBe(300)
  expect(ledger.snapshot().byPool.copy).toBe(144)
  expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull()
  const notice = ui.getByRole('status').textContent
  await act(async () => writes.jobs[0].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ui.getByRole('status').textContent).toBe(notice)
})

it.each(['Cancel copy', 'Escape', 'selection', 'permission'])('%s abort leaves a submitted clipboard charge until rejection and does not steal stale focus', async action => {
  const writes = clipboard(), { ui, ledger, doc } = await mounted()
  const fallback = vi.fn(() => true)
  Object.defineProperty(document, 'execCommand', { configurable: true, value: fallback })
  await copy(ui, writes, 1)
  if (action === 'Cancel copy') fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  else if (action === 'Escape') fireEvent.keyDown(ui.getByRole('region', { name: 'Spreadsheet viewer' }), { key: 'Escape' })
  else if (action === 'selection') go(ui, 'A3')
  else ui.rerender(<InteractiveSheetDoc doc={doc} allowCopy={false} />)
  const focus = ui.getByRole('textbox', { name: 'Go To cell' })
  act(() => focus.focus())
  expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull()
  expect(ledger.snapshot().byPool.copy).toBe(144)
  const notice = ui.getByRole('status').textContent
  await act(async () => writes.jobs[0].reject(new Error('Native clipboard rejected')))
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ui.getByRole('status').textContent).toBe(notice)
  expect(document.activeElement).toBe(focus)
  expect(fallback).not.toHaveBeenCalled()
})

it('a current rejected clipboard job refunds on refused fallback and preserves input focus', async () => {
  const writes = clipboard(), { ui, ledger } = await mounted()
  const fallback = vi.fn(() => false)
  Object.defineProperty(document, 'execCommand', { configurable: true, value: fallback })
  await copy(ui, writes, 1)
  const focus = ui.getByRole('textbox', { name: 'Go To cell' })
  act(() => focus.focus())
  await act(async () => writes.jobs[0].reject(new Error('Native clipboard rejected')))
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ui.getByRole('status').textContent).toContain('Copy failed. Try copying again.')
  expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull()
  expect(fallback).toHaveBeenCalledWith('copy')
  expect(document.activeElement).toBe(focus)
  expect(document.querySelector('textarea')).toBeNull()
})

it('accounts serialization bytes reported after synchronous cancellation before settling the job', async () => {
  const writes = clipboard(), doc = docOf(), source = createObjectWorkbookSource(doc, 'late-copy')
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A2')
  const store = source.store('A'), read = store.readCell.bind(store)
  let cancelled = false
  // The source value accessor cancels during the real formatter. The serializer
  // still owns its current batch and reports its bytes before checking abort.
  store.readCell = async (...args) => {
    const result = await read(...args)
    if (result.kind !== 'cell') return result
    return { ...result, cell: { ...result.cell, get value() {
      if (!cancelled) { cancelled = true; fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' })) }
      return result.cell.value
    } } }
  }
  const set = ledger.set.bind(ledger), held: number[] = []
  vi.spyOn(ledger, 'set').mockImplementation((pool, owner, bytes) => {
    set(pool, owner, bytes)
    if (pool === 'copy' && bytes > 0 && cancelled) held.push(ledger.snapshot().byPool.copy)
  })
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(held.some(bytes => bytes > 0)).toBe(true))
  await waitFor(() => expect(ledger.snapshot().byPool.copy).toBe(0))
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(writes.writeText).not.toHaveBeenCalled()
  expect(ui.queryByRole('button', { name: 'Cancel copy' })).toBeNull()
  expect(ui.getByRole('status').textContent).not.toContain('Selection copied.')
})

it('unmount and a real remote remount retain independent unsettled clipboard owners', async () => {
  const writes = clipboard(), doc = docOf(), source = createObjectWorkbookSource(doc, 'd')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])) })
  const ledger = remote.mainLedger
  try {
    const ui = render(<InteractiveSheetDoc doc={doc} source={remote} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
    go(ui, 'A2')
    await copy(ui, writes, 1)
    ui.unmount()
    expect(ledger.snapshot().byPool.copy).toBe(144)
    const again = render(<InteractiveSheetDoc doc={doc} source={remote} />)
    await waitFor(() => expect(again.getByRole('status').textContent).not.toContain('Loading sheet'))
    go(again, 'A3')
    await copy(again, writes, 2)
    const newer = 128 + writes.jobs[1].text.length * 4
    expect(ledger.snapshot().byPool.copy).toBe(144 + newer)
    await act(async () => writes.jobs[0].resolve())
    expect(ledger.snapshot().byPool.copy).toBe(newer)
    expect(again.getByRole('button', { name: 'Cancel copy' })).toBeTruthy()
    expect(again.getByRole('status').textContent).not.toContain('Selection copied.')
    again.unmount()
    await remote.close()
    expect(ledger.snapshot().byPool.copy).toBe(newer)
    await act(async () => writes.jobs[1].resolve())
    expect(ledger.snapshot().byPool.copy).toBe(0)
    await act(async () => { await Promise.resolve() })
    expect(ledger.snapshot().byPool.copy).toBe(0)
  } finally { await remote.close(); await client.dispose() }
})

it('a settled copy allocation callback cannot revive its closed ledger owner', async () => {
  const formatting = await import('../src/xlsx/render')
  const canonical = formatting.cellDisplayText
  let allocation: import('../src/xlsx/formula/format').FormatAllocation | undefined
  vi.spyOn(formatting, 'cellDisplayText').mockImplementation((sheet, cell, context) => {
    if (context) allocation = context
    return canonical(sheet, cell, context)
  })
  const writes = clipboard(), { ui, ledger } = await mounted()
  await copy(ui, writes, 1)
  expect(allocation).toBeDefined()
  await act(async () => writes.jobs[0].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(() => allocation!.allocate(512)).toThrow('Copy is too large')
  expect(ledger.snapshot().byPool.copy).toBe(0)
})

it('a diagnostic created by offscreen copy stays document-owned after native settlement and unmount', async () => {
  const writes = clipboard()
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 64, value: 'beta' }] }])
  doc.sheets[0].rows[0].cells[0].style = { numFmtId: 0, formatCode: '0;0;0;yyyy' }
  const source = createObjectWorkbookSource(doc, 'diagnostic-copy')
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} source={source} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  fireEvent.click(ui.getByRole('button', { name: 'Select row 1' }))
  const before = ledger.snapshot().byPool.document
  expect(doc.sheets[0].diagnostics).toBeUndefined()
  await copy(ui, writes, 1)
  expect(writes.jobs[0].text).toBe('\t'.repeat(64) + 'beta' + '\t'.repeat(16384 - 65))
  expect(doc.sheets[0].diagnostics).toHaveLength(1)
  const retained = ledger.snapshot().byPool.document
  expect(retained).toBeGreaterThan(before)
  await act(async () => writes.jobs[0].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(ledger.snapshot().byPool.document).toBe(retained)
  ui.unmount()
  expect(ledger.snapshot().byPool.document).toBe(retained)
  expect(doc.sheets[0].diagnostics).toHaveLength(1)
})

it.each([73868, 73867])('post-diagnostic growth uses the debited copy owner with %i bytes left', async capacity => {
  const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
  const writes = clipboard(), doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 64, value: 'beta' }] }])
  doc.sheets[0].rows[0].cells[0].style = { numFmtId: 0, formatCode: '0;0;0;yyyy' }
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  fireEvent.click(ui.getByRole('button', { name: 'Select row 1' }))
  const adopt = ledger.adoptWhileReachable.bind(ledger), growthTotals: number[] = []
  let pinned: ReturnType<typeof pinCopyCapacity> | undefined
  vi.spyOn(ledger, 'adoptWhileReachable').mockImplementation((...args) => {
    const token = adopt(...args)
    pinned ??= pinCopyCapacity(ledger, MAIN_THREAD_BUDGETS.D, capacity, (bytes, previous) => { if (bytes > previous) growthTotals.push(ledger.bytes) })
    return token
  })
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(writes.jobs.length === 1 || ui.getByRole('status').textContent?.includes('too large')).toBe(true), { timeout: 5000 })
  expect(pinned).toBeDefined()
  expect(growthTotals.every(bytes => bytes <= MAIN_THREAD_BUDGETS.D)).toBe(true)
  expect(doc.sheets[0].diagnostics).toHaveLength(1)
  if (capacity === 73868) {
    expect(writes.jobs[0].text).toBe('\t'.repeat(64) + 'beta' + '\t'.repeat(16384 - 65))
    await act(async () => writes.jobs[0].resolve())
  } else expect(writes.writeText).not.toHaveBeenCalled()
  expect(ledger.snapshot().byPool.copy).toBe(0)
  expect(pinned!.held()).toBe(true)
  pinned!.stop()
})

it('two real mounted serializers share live reservations after both source reads suspend', async () => {
  const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
  const writes = clipboard(), doc = docOf()
  doc.sheets[0].rows[1].cells[0].value = 'x'.repeat(100)
  const source = createObjectWorkbookSource(doc, 'concurrent-copy')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'concurrent-copy', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId: 'concurrent-copy', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('concurrent-copy')!.layout(sheet.sheetId)])) })
  try {
    const first = render(<InteractiveSheetDoc doc={doc} source={remote} />), second = render(<InteractiveSheetDoc doc={doc} source={remote} />)
    const a = { ...first, ...within(first.container) }, b = { ...second, ...within(second.container) }
    await waitFor(() => expect(a.getByRole('status').textContent).not.toContain('Loading sheet'))
    await waitFor(() => expect(b.getByRole('status').textContent).not.toContain('Loading sheet'))
    go(a, 'A2'); go(b, 'A2')
    const store = remote.store('A'), read = store.readCell.bind(store), pending: (() => void)[] = []
    store.readCell = async (...args) => { await new Promise<void>(resolve => pending.push(resolve)); return read(...args) }
    fireEvent.click(a.getByRole('button', { name: 'Copy selection' }))
    fireEvent.click(b.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(pending.length).toBe(2))
    const ledger = remote.mainLedger
    expect(ledger.snapshot().byPool.copy).toBe(320)
    const totals: number[] = [], pinned = pinCopyCapacity(ledger, MAIN_THREAD_BUDGETS.D, 720, () => totals.push(ledger.bytes))
    await act(async () => pending[0]())
    await waitFor(() => expect(writes.jobs.length).toBe(1))
    await act(async () => pending[1]())
    await waitFor(() => expect(b.getByRole('status').textContent).toContain('too large'))
    expect(writes.jobs.map(job => job.text)).toEqual(['x'.repeat(100)])
    expect(totals.every(bytes => bytes <= MAIN_THREAD_BUDGETS.D)).toBe(true)
    expect(ledger.snapshot().byPool.copy).toBe(528)
    expect(a.getByRole('button', { name: 'Cancel copy' })).toBeTruthy()
    expect(pinned.held()).toBe(true)
    pinned.stop()
    await act(async () => writes.jobs[0].resolve())
    expect(ledger.snapshot().byPool.copy).toBe(0)
    store.readCell = read
    first.unmount(); second.unmount()
  } finally { await remote.close(); await client.dispose() }
})

it('diagnostic adoption immediately debits the admission lease before its next scratch refund', async () => {
  const { MAIN_THREAD_BUDGETS } = await import('../src/xlsx/view/tileCache')
  const copying = await import('../src/xlsx/view/copy'), formatting = await import('../src/xlsx/render')
  const actualCopy = copying.copySelection, canonical = formatting.cellDisplayText
  let admission: import('../src/xlsx/view/copy').CopyInput['admitBytes']
  vi.spyOn(copying, 'copySelection').mockImplementation(input => { admission = input.admitBytes; return actualCopy(input) })
  const writes = clipboard(), doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 64, value: 'beta' }] }])
  doc.sheets[0].rows[0].cells[0].style = { numFmtId: 0, formatCode: '0;0;0;yyyy' }
  let ledger!: MainMemoryLedger, probed = false, pinHeld = false, accepted: boolean | undefined, beforeGrowth = 0, afterGrowth = 0
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  fireEvent.click(ui.getByRole('button', { name: 'Select row 1' }))
  vi.spyOn(formatting, 'cellDisplayText').mockImplementation((sheet, cell, context) => {
    const text = canonical(sheet, cell, context)
    if (context && sheet.diagnostics?.length && !probed) {
      probed = true
      const own = ledger.snapshot().byPool.copy
      const pinned = pinCopyCapacity(ledger, MAIN_THREAD_BUDGETS.D, own)
      beforeGrowth = ledger.bytes
      accepted = admission!(own + 1)
      afterGrowth = ledger.bytes
      pinHeld = pinned.held()
      pinned.stop()
    }
    return text
  })
  await copy(ui, writes, 1)
  expect(probed).toBe(true)
  expect(accepted).toBe(false)
  expect(pinHeld).toBe(true)
  expect(afterGrowth).toBe(beforeGrowth)
  expect(doc.sheets[0].diagnostics).toHaveLength(1)
  await act(async () => writes.jobs[0].resolve())
  expect(ledger.snapshot().byPool.copy).toBe(0)
})

it('an M viewer on a D remote gets copy admission at the D budget', async () => {
  const copying = await import('../src/xlsx/view/copy')
  const actualCopy = copying.copySelection
  let capturedSessionBudget: number | undefined
  let capturedAdmission: ((bytes: number) => boolean) | undefined
  vi.spyOn(copying, 'copySelection').mockImplementation(input => {
    capturedSessionBudget = input.sessionBudgetBytes
    capturedAdmission = input.admitBytes
    return actualCopy(input)
  })
  const writes = clipboard()
  const doc = docOf()
  const source = createObjectWorkbookSource(doc, 'd')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], profile: 'D', formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])) })
  try {
    const ui = render(<InteractiveSheetDoc doc={doc} source={remote} profile="M" />)
    await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
    go(ui, 'A2')
    await copy(ui, writes, 1)
    expect(capturedSessionBudget).toBeGreaterThan(copying.copySessionBudget('M'))
    expect(capturedSessionBudget).toBeLessThanOrEqual(copying.copySessionBudget('D'))
    // 15 MiB exceeds mobile budget (12 MiB), but fits desktop budget (24 MiB):
    expect(capturedAdmission!(15 * 1024 * 1024)).toBe(true)
    await act(async () => writes.jobs[0].resolve())
    expect(remote.mainLedger.snapshot().byPool.copy).toBe(0)
    ui.unmount()
  } finally {
    await remote.close(); await client.dispose()
  }
})
