import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { ViewerClient } from '../src/worker/viewerClient'
import { connectInProcess, ViewerKernel } from '../src/worker/viewerKernel'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import { RemoteWorkbook } from '../src/worker/viewerRemoteSource'
import { workbookOf } from './helpers/xlsx-find'

afterEach(() => cleanup())

describe('round 6: actual paged viewer owner teardown', () => {
  it('unmount closes the physical cache even when ledger releaseAll already reports zero', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'long text' }, { row: 0, col: 3, value: 'block' }] }])
    const source = createObjectWorkbookSource(doc, 'd')
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel))
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile: 'D' })
    const remote = new RemoteWorkbook(client, { docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])) })
    const cache = remote.overflowFor()
    try {
      const ui = render(<InteractiveSheetDoc doc={doc} source={remote} />)
      await waitFor(() => expect(cache.bytes).toBeGreaterThan(4096))
      ui.unmount()
      expect(remote.mainLedger.snapshot().byPool.layout).toBe(0)
      expect(cache.closed).toBe(true)
      expect(cache.bytes).toBe(0)
      expect(cache.pool.totalBytes).toBe(0)
      const fresh = remote.overflowFor()
      expect(fresh).not.toBe(cache)
      const again = render(<InteractiveSheetDoc doc={doc} source={remote} />)
      await waitFor(() => expect(fresh.bytes).toBeGreaterThan(4096))
      expect(again.container.textContent).not.toContain('exceeds the document memory budget')
      again.unmount()
      expect(fresh.closed).toBe(true)
    } finally { await remote.close(); await client.dispose() }
  })

  it('keeps a shared remote overflow cache live until its last mounted viewer releases it', async () => {
    const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'long text' }, { row: 0, col: 3, value: 'block' }] }])
    const source = createObjectWorkbookSource(doc, 'shared-overflow-viewers')
    const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
    const client = new ViewerClient(connectInProcess(kernel))
    await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'shared-overflow-viewers', source: { kind: 'object', document: doc }, profile: 'D' })
    const remote = new RemoteWorkbook(client, { docId: 'shared-overflow-viewers', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('shared-overflow-viewers')!.layout(sheet.sheetId)])) })
    const cache = remote.overflowFor()
    try {
      const first = render(<InteractiveSheetDoc doc={doc} source={remote} />)
      const second = render(<InteractiveSheetDoc doc={doc} source={remote} />)
      await waitFor(() => expect(cache.bytes).toBeGreaterThan(4096))
      first.unmount()
      expect(cache.closed).toBe(false)
      expect(cache.bytes).toBeGreaterThan(4096)
      fireEvent.change(second.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'long' } })
      await waitFor(() => expect(second.getByTestId('officeview-search-status').textContent).toContain('1 of 1'))
      expect(second.container.textContent).not.toContain('Text overflow occupancy owner is closed')
      second.unmount()
      expect(cache.closed).toBe(true)
      expect(cache.bytes).toBe(0)
    } finally { await remote.close(); await client.dispose() }
  })
})
