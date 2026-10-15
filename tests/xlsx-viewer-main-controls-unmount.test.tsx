import { afterEach, expect, test, vi } from 'vitest'
import { act, cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import * as factory from '../src/components/xlsx/viewerControlsClient'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { workbookOf } from './helpers/xlsx-find'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

test('actual mounted viewer unmount preserves held membership and run-copy owners through reader drain', async () => {
  const canonical = factory.viewerControls
  let options!: factory.ViewerControlsOptions, controls!: factory.ViewerControls
  vi.spyOn(factory, 'viewerControls').mockImplementation(async (...args) => {
    options = args[0]
    controls = await canonical(...args)
    return controls
  })
  let ledger!: MainMemoryLedger
  const doc = workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'Header' }, ...Array.from({ length: 16 }, (_, i) => ({ row: i + 1, col: 0, value: i + 1 }))] }])
  const ui = render(<InteractiveSheetDoc doc={doc} {...({ onMainLedger: (value: MainMemoryLedger) => { ledger = value } } as object)} />)
  await waitFor(() => expect(controls).toBeDefined())
  const resources = options.resources!, signal = () => new AbortController().signal
  const range = { r0: 0, r1: 16, c0: 0, c1: 0, headerRows: 1 as const, totalsRows: 0 as const }
  const draft = await controls.client.values(options.snapshot().owner, range, 0, null, signal())
  await draft.read({ text: '', offset: 0 }, signal())
  await draft.change({ kind: 'all', selected: false, includeBlanks: false }, { text: '', offset: 0 }, signal())
  await draft.change({ kind: 'toggle', ordinal: 0, selected: true }, { text: '', offset: 0 }, signal())
  const rule = await draft.candidate(signal())
  await act(async () => { expect((await controls.client.filter(options.snapshot().owner, { range, columns: new Map([[0, rule!]]) }, signal(), undefined, draft)).kind).toBe('committed') })
  let hit!: () => void, open!: () => void
  const reached = new Promise<void>(resolve => hit = resolve), wait = new Promise<void>(resolve => open = resolve)
  const read = resources.runs.read.bind(resources.runs)
  resources.runs.read = async (...args) => { const page = await read(...args); hit(); await wait; return page }
  const pending = controls.client.filter(options.snapshot().owner, { range, columns: new Map([[0, { ...rule!, showBlanks: !rule!.showBlanks }]]) }, signal())
  try {
    await reached
    const full = ledger.snapshot()
    expect(full.byPool.document).toBeGreaterThan(0)
    expect(full.byPool.geometry).toBeGreaterThan(0)
    expect(full.byPool.dom).toBeGreaterThan(0)
    expect(full.byPool.viewBuffers).toBeGreaterThan(0)
    ui.unmount()
    draft.cancel()
    expect(resources.cache.usedBytes).toBeGreaterThanOrEqual(1538)
    expect(ledger.snapshot().byPool.controls).toBeGreaterThanOrEqual(1538 + 9)
    expect(ledger.snapshot().byPool.dom).toBe(0)
    await expect(pending).resolves.toMatchObject({ kind: 'cancelled' })
    expect(resources.cache.usedBytes).toBeGreaterThanOrEqual(1538)
    expect(ledger.snapshot().byPool.controls).toBeGreaterThanOrEqual(1538 + 9)
    open()
    // Public cancellation already settled; wait for the actual held reader cleanup.
    await waitFor(() => expect(resources.cache.usedBytes).toBe(0))
    expect(resources.memory.pendingBytes).toBe(0)
  } finally { open(); await pending.catch(() => {}); draft.cancel(); ui.unmount() }
}, 30000)
