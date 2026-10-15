import { afterEach, describe, expect, it } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
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

  it('a viewer re-rendered after its workbook closed refuses occupancy instead of throwing', async () => {
    const f = await opened()
    try {
      const cache = f.remote.overflowFor()
      const ui = render(<InteractiveSheetDoc doc={f.doc} source={f.remote} />)
      await waitFor(() => expect(cache.bytes).toBeGreaterThan(4096))
      await f.remote.close()
      expect(cache.closed).toBe(true)
      expect(() => ui.rerender(<InteractiveSheetDoc doc={f.doc} source={f.remote} />)).not.toThrow()
      // Still mounted (no error boundary unwind): the grid keeps rendering.
      expect(ui.container.querySelector('[role="grid"]')).not.toBeNull()
      expect(() => ui.rerender(<InteractiveSheetDoc doc={f.doc} source={f.remote} profile="M" />)).not.toThrow()
      expect(f.remote.mainLedger.snapshot().byPool.layout).toBe(0)
      ui.unmount()
    } finally { await f.remote.close(); await f.client.dispose() }
  })
})

async function opened() {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'long text' }, { row: 0, col: 3, value: 'block' }] }])
  const source = createObjectWorkbookSource(doc, 'd')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile: 'D' })
  const remote = new RemoteWorkbook(client, { docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])) })
  return { doc, client, remote }
}
