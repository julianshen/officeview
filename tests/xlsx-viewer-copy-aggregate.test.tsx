import { afterEach, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { workbookOf } from './helpers/xlsx-find'

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
it.each(['D', 'M'] as const)('independent aggregate admission: repeated live clipboard writes stay within profile %s', async profile => {
  const writes = clipboard()
  const doc = workbookOf([{ sheetId: 'A', cells: Array.from({ length: 128 }, (_, col) => ({ row: 0, col: col + 64, value: `${col}:` + 'x'.repeat(profile === 'D' ? 32600 : 16200) })) }])
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} profile={profile} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  fireEvent.click(ui.getByRole('button', { name: 'Select row 1' }))
  const initial = ledger.snapshot(), budget = MAIN_THREAD_BUDGETS[profile]
  const sourceBefore = JSON.stringify(doc.sheets[0].rows), viewBefore = ui.getByRole('button', { name: 'Select row 1' }).closest('[role="rowheader"]')?.getAttribute('aria-selected')
  expect(initial.bytes).toBeLessThan(budget)
  expect(viewBefore).toBe('true')
  const fields = Array.from({ length: 16384 }, (_, col) => col >= 64 && col < 192 ? `${col - 64}:` + 'x'.repeat(profile === 'D' ? 32600 : 16200) : '')
  const expected = fields.join('\t'), totals: number[] = []
  const set = ledger.set.bind(ledger)
  vi.spyOn(ledger, 'set').mockImplementation((pool, owner, bytes) => { set(pool, owner, bytes); if (pool === 'copy') totals.push(ledger.bytes) })
  const fallback = vi.fn(() => true)
  Object.defineProperty(document, 'execCommand', { configurable: true, value: fallback })
  const readings: number[] = []
  for (let i = 0; i < 5; i++) {
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(writes.writeText.mock.calls.length >= i + 1 || ui.getByRole('status').textContent?.includes('too large')).toBe(true), { timeout: 5000 })
    readings.push(ledger.snapshot().byPool.copy)
    if (writes.jobs.length < i + 1) break
    expect(writes.jobs[i].text).toBe(expected)
    fireEvent.click(ui.getByRole('button', { name: 'Cancel copy' }))
  }
  const beforeSettle = ledger.snapshot()
  expect(totals.reduce((peak, bytes) => Math.max(peak, bytes), 0)).toBeLessThanOrEqual(budget)
  expect(beforeSettle.bytes).toBeLessThanOrEqual(budget)
  expect(writes.jobs.length).toBeLessThan(5)
  expect(writes.jobs.length).toBeGreaterThan(0)
  expect(ui.getByRole('status').textContent).toContain('too large')
  expect(fallback).not.toHaveBeenCalled()
  expect(document.querySelector('textarea')).toBeNull()
  expect(JSON.stringify(doc.sheets[0].rows)).toBe(sourceBefore)
  expect(ui.getByRole('button', { name: 'Select row 1' }).closest('[role="rowheader"]')?.getAttribute('aria-selected')).toBe(viewBefore)
  const accepted = writes.jobs.length
  await act(async () => writes.jobs[0].resolve())
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(writes.jobs.length).toBe(accepted + 1), { timeout: 5000 })
  expect(writes.jobs[accepted].text).toBe(expected)
  expect(totals.reduce((peak, bytes) => Math.max(peak, bytes), 0)).toBeLessThanOrEqual(budget)
  await act(async () => { for (const job of writes.jobs) job.resolve() })
  expect(ledger.snapshot().byPool.copy).toBe(0)
  console.log(JSON.stringify({ profile, initial: initial.bytes, accepted, readings, acceptedGrowths: totals.length, peak: totals.reduce((peak, bytes) => Math.max(peak, bytes), 0), budget }))
}, 60000)

it.each([5876, 5875])('other pools leave %i bytes of headroom for a tiny copy construction', async headroom => {
  const writes = clipboard(), doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'beta' }] }])
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  ledger.set('controls', 'pressure', MAIN_THREAD_BUDGETS.D - ledger.bytes - headroom)
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(writes.jobs.length === 1 || ui.getByRole('status').textContent?.includes('too large')).toBe(true))
  // One new pending-action DOM node (5700 B) plus 176 B serializer peak.
  if (headroom === 5876) {
    expect(writes.jobs.map(job => job.text)).toEqual(['beta'])
    expect(ledger.snapshot().byPool.copy).toBe(144)
    // Unrelated pressure may exceed the cap after admission; shrink/refund
    // must still work and must not erase the pressure owner's charge.
    const pressure = MAIN_THREAD_BUDGETS.D
    ledger.set('controls', 'pressure', pressure)
    await act(async () => writes.jobs[0].resolve())
    expect(ledger.snapshot().byPool.copy).toBe(0)
    expect(ledger.snapshot().byPool.controls).toBeGreaterThanOrEqual(pressure)
  } else {
    expect(writes.writeText).not.toHaveBeenCalled()
    expect(ui.getByRole('status').textContent).toContain('too large')
    expect(ledger.snapshot().byPool.copy).toBe(0)
  }
  ledger.release('controls', 'pressure')
})

it.each([7992, 7991])('caller-owned hidden-column buffers share the same %i-byte headroom', async capacity => {
  const writes = clipboard(), doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'beta' }] }])
  doc.sheets[0].cols = [{ min: 2, max: 2, hidden: true }]
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  ledger.set('controls', 'pressure', MAIN_THREAD_BUDGETS.D - ledger.bytes - capacity)
  fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
  await waitFor(() => expect(writes.jobs.length === 1 || ui.getByRole('status').textContent?.includes('too large')).toBe(true))
  if (capacity === 7992) {
    expect(writes.jobs.map(job => job.text)).toEqual(['beta'])
    expect(ledger.snapshot().byPool.copy).toBe(2260) // 2116 column mask + 144 output/projection
    await act(async () => writes.jobs[0].resolve())
  } else expect(writes.writeText).not.toHaveBeenCalled()
  expect(ledger.snapshot().byPool.copy).toBe(0)
  ledger.release('controls', 'pressure')
})

it('refuses hidden-column construction before cloning intervals or allocating visibility words', async () => {
  const writes = clipboard(), doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'beta' }] }])
  doc.sheets[0].cols = [{ min: 2, max: 2, hidden: true }]
  let ledger!: MainMemoryLedger
  const ui = render(<InteractiveSheetDoc doc={doc} onMainLedger={value => { ledger = value }} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  go(ui, 'A1')
  const Original = Uint32Array, constructed: number[] = []
  vi.stubGlobal('Uint32Array', class extends Original { constructor(length: number) { constructed.push(length); super(length) } })
  ledger.set('controls', 'pressure', MAIN_THREAD_BUDGETS.D - ledger.bytes - 4323)
  try {
    fireEvent.click(ui.getByRole('button', { name: 'Copy selection' }))
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('too large'))
    expect(constructed).not.toContain(512)
    expect(writes.writeText).not.toHaveBeenCalled()
    expect(ledger.snapshot().byPool.copy).toBe(0)
  } finally { vi.unstubAllGlobals(); ledger.release('controls', 'pressure') }
})
