/**
 * S13.3 review round 3 (#3, #5): the automatic refresh of a search that
 * finished over importing sheets publishes nothing after unmount, and keeps
 * arbitrary sheet IDs (a ',' inside one) intact. React's state setters are
 * counted so a publication after unmount is observable.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'

const setters = vi.hoisted(() => ({ calls: 0, counting: false }))
vi.mock('react', async original => {
  const actual = await original<typeof import('react')>()
  const wrapped = new WeakMap<object, unknown>()
  const useState = ((initial: unknown) => {
    const [value, set] = actual.useState(initial)
    let counted = wrapped.get(set)
    if (!counted) wrapped.set(set, counted = (next: unknown) => { if (setters.counting) setters.calls++; set(next) })
    return [value, counted]
  }) as typeof actual.useState
  return { ...actual, default: { ...actual, useState }, useState }
})
vi.mock('../src/xlsx/storage/importSession', async original => {
  const actual = await original<typeof import('../src/xlsx/storage/importSession')>()
  return { ...actual, importLimitsFor: (profile: 'D' | 'M', overrides?: Parameters<typeof actual.importLimitsFor>[1]) => actual.importLimitsFor(profile, { ...overrides, object: { ...overrides?.object, rows: 1 } }) }
})

import { OfficeFile } from '../src/components/OfficeFile'
import { ViewerKernel, type ViewerKernelEnv } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { WorkerCommand, WorkerReply } from '../src/xlsx/view/contracts'
import type { WorkbookSource } from '../src/worker/viewerSession'
import { xlsxBytes } from './helpers/xlsx-package'
import { workbookOf } from './helpers/xlsx-find'

afterEach(() => { cleanup(); vi.restoreAllMocks(); setters.counting = false; setters.calls = 0 })

const CORNER = '<row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c><c r="B1"><v>3</v></c></row>'
  + '<row r="2"><c r="A2" t="inlineStr"><is><t>beta</t></is></c></row>'
const workbook = () => new Blob([xlsxBytes({ sheetData: CORNER }) as BlobPart])

function workerHarness(env: Partial<ViewerKernelEnv> = {}) {
  const kernel = new ViewerKernel({ parse: async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes)), ...env }, { yield: async () => {} })
  const listeners = new Set<(event: MessageEvent) => void>()
  const posted: WorkerCommand[] = []
  const held: Array<() => void> = []
  let hold: ((command: WorkerCommand) => boolean) | null = null
  const port = {
    postMessage(message: unknown, transfer: Transferable[]) {
      const open = message as { kind: string; source?: { kind: string; blob?: Blob } }
      const blob = open.kind === 'open' && open.source?.kind === 'blob' ? open.source.blob : undefined
      const command = structuredClone(blob ? { ...open, source: { kind: 'blob' } } : message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      if (blob && command.kind === 'open') command.source = { kind: 'blob', blob }
      posted.push(command)
      void kernel.handle(command, (out: WorkerReply, replyTransfer) => {
        const deliver = () => { const data = structuredClone(out, { transfer: replyTransfer }); for (const listener of [...listeners]) listener({ data } as MessageEvent) }
        if (hold?.(command)) held.push(deliver)
        else deliver()
      })
    },
    addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.add(listener) },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') listeners.delete(listener) },
    terminate: vi.fn(),
  }
  return {
    port, kernel, posted,
    searches: () => posted.filter(command => command.kind === 'search').length,
    hold(predicate: (command: WorkerCommand) => boolean) { hold = predicate },
    release() { hold = null; for (const deliver of held.splice(0)) deliver() },
  }
}
/** The first search sees `sheetId`'s structure still importing. */
function importingOnce(h: ReturnType<typeof workerHarness>, sheetId: string) {
  const open = h.posted.find(command => command.kind === 'open')!
  if (open.kind !== 'open') throw new Error('Expected open command')
  const store = h.kernel.session(open.docId)!.source.store(sheetId)
  const structural = store.queryStructural.bind(store)
  return vi.spyOn(store, 'queryStructural').mockImplementationOnce(async (rect, options) => ({ ...(await structural(rect, options)), completeness: 'incomplete' }))
}

describe('S13.3 r3: automatic search refresh', () => {
  it('#3: a refresh still running at unmount publishes nothing afterwards', async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    importingOnce(h, '1')
    // Hold the automatic (second) search reply.
    h.hold(command => command.kind === 'search' && h.searches() >= 2)
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'alpha' } })
    await waitFor(() => expect(h.searches()).toBe(2))
    setters.counting = true
    ui.unmount()
    const afterUnmount = setters.calls
    act(() => h.release())
    await act(async () => { await new Promise(resolve => setTimeout(resolve, 20)) })
    expect(setters.calls - afterUnmount, 'state setters run after unmount').toBe(0)
  })

  it('r4 #3: a new query with the same incomplete sheets is refreshed with its own arguments', async () => {
    const h = workerHarness()
    const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
    // The sheet's structure is still importing for both searches...
    const structural = importingOnce(h, '1')
    const original = structural.getMockImplementation()!
    structural.mockImplementationOnce(original)
    let loaded = false, remote: RemoteWorkbook | undefined
    const loadState = RemoteWorkbook.prototype.loadState
    vi.spyOn(RemoteWorkbook.prototype, 'loadState').mockImplementation(function (this: RemoteWorkbook, sheetId: string) {
      remote = this
      return loaded ? loadState.call(this, sheetId) : { completeness: 'loading' }
    })
    const box = ui.getByRole('searchbox', { name: 'Search document' })
    fireEvent.change(box, { target: { value: 'alpha' } })
    await waitFor(() => expect(ui.container.textContent).toMatch(/some sheets still loading/))
    fireEvent.change(box, { target: { value: 'beta' } })
    await waitFor(() => expect(h.searches()).toBe(2))
    await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
    expect(ui.container.textContent).toMatch(/some sheets still loading/)
    // ...then the import completes: beta's page is refreshed, not alpha's.
    loaded = true
    act(() => remote!.notify())
    await waitFor(() => expect(h.searches()).toBe(3))
    const last = h.posted.filter(command => command.kind === 'search').at(-1)
    expect(last?.kind === 'search' && last.query.text).toBe('beta')
    await waitFor(() => expect(ui.container.textContent).not.toMatch(/some sheets still loading/))
    ui.unmount()
  })

  it('#5: a sheet ID containing "," refreshes as one sheet, with no phantom sheet lookups', async () => {
    const doc = workbookOf([{ sheetId: 'A,B', cells: [{ row: 0, col: 0, value: 'alpha' }, { row: 1, col: 0, value: 'beta' }] }])
    let source: WorkbookSource | undefined
    const h = workerHarness({ openSource: async (_handle, docId) => (source = createObjectWorkbookSource(doc, docId)) })
    const errors: unknown[] = []
    const onError = (event: PromiseRejectionEvent | Event) => { errors.push((event as PromiseRejectionEvent).reason ?? event) }
    process.on('unhandledRejection', onError as never)
    try {
      const ui = render(<OfficeFile data={workbook()} interactiveSheets viewerWorker={() => h.port} />)
      await waitFor(() => expect(ui.getByRole('gridcell', { name: /alpha/ })).toBeTruthy())
      expect(source?.sheets.map(sheet => sheet.sheetId)).toEqual(['A,B'])
      importingOnce(h, 'A,B')
      fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'alpha' } })
      await waitFor(() => expect(h.searches()).toBe(2))
      await waitFor(() => expect(ui.container.textContent).toMatch(/1 of 1/))
      expect(ui.container.textContent).not.toMatch(/some sheets still loading/)
      await act(async () => { await new Promise(resolve => setTimeout(resolve, 20)) })
      expect(errors.map(String).filter(text => /Unknown sheet/.test(text))).toEqual([])
      ui.unmount()
    } finally { process.off('unhandledRejection', onError as never) }
  })
})
