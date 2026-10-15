import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { boundOverflowCache } from '../src/xlsx/view/overflowCache'
import * as overflowGeometry from '../src/xlsx/view/viewport'
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
      // One shared closed sentinel: no per-render owner allocation churn.
      expect(f.remote.overflowFor()).toBe(f.remote.overflowFor('M'))
      ui.unmount()
    } finally { await f.remote.close(); await f.client.dispose() }
  })

  it('rebinds per frame when a sibling viewer unmount closes the shared remote cache', async () => {
    const f = await opened()
    try {
      // Siblings share the remote's cache under budgetProfile. When one viewer unmounts,
      // its teardown closes the shared cache. The surviving viewer's frame-time rebind
      // recovers with a fresh live cache rather than painting with a closed owner.
      function Harness({ showB }: { showB: boolean }) {
        return (
          <>
            <div data-testid="a"><InteractiveSheetDoc doc={f.doc} source={f.remote} /></div>
            {showB && <div data-testid="b"><InteractiveSheetDoc doc={f.doc} source={f.remote} /></div>}
          </>
        )
      }
      const ui = render(<Harness showB={true} />)
      await waitFor(() => expect(f.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(4096))
      await waitFor(() => expect(ui.getAllByRole('gridcell', { name: /value long text/ })).toHaveLength(2))
      const initialCache = f.remote.overflowFor()
      expect(initialCache.closed).toBe(false)

      // Unmount sibling B: closes its overflowContext.cache (the shared remote cache).
      ui.rerender(<Harness showB={false} />)
      await waitFor(() => expect(ui.getAllByRole('gridcell', { name: /value long text/ })).toHaveLength(1))
      expect(initialCache.closed).toBe(true)

      // Trigger a frame in surviving viewer A after B unmounted and closed the shared cache.
      f.remote.notify()
      await waitFor(() => expect(f.remote.overflowFor().closed).toBe(false))

      expect(f.remote.overflowFor()).not.toBe(initialCache)
      expect(ui.getByTestId('a').textContent).not.toContain('Text overflow occupancy owner is closed')
      ui.unmount()
    } finally { await f.remote.close(); await f.client.dispose() }
  })

  it('an overlay-only find update after a sibling unmount recovers with an open cache and shows no closed notice', async () => {
    const f = await opened()
    const seenCaches: { closed: boolean }[] = []
    const origCellTextOverflow = overflowGeometry.cellTextOverflow
    const overflowSpy = vi.spyOn(overflowGeometry, 'cellTextOverflow').mockImplementation((geom, src) => {
      const cache = geom?.kind === 'ready' ? boundOverflowCache(geom) : undefined
      if (cache) seenCaches.push({ closed: cache.closed })
      return origCellTextOverflow(geom, src)
    })
    try {
      function Harness({ showB }: { showB: boolean }) {
        return (
          <>
            <div data-testid="a"><InteractiveSheetDoc doc={f.doc} source={f.remote} /></div>
            {showB && <div data-testid="b"><InteractiveSheetDoc doc={f.doc} source={f.remote} /></div>}
          </>
        )
      }
      const ui = render(<Harness showB={true} />)
      await waitFor(() => expect(f.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(4096))
      await waitFor(() => expect(ui.getAllByRole('gridcell', { name: /value long text/ })).toHaveLength(2))
      const initialCache = f.remote.overflowFor()
      expect(initialCache.closed).toBe(false)

      // Unmount sibling B: closes its overflowContext.cache (the shared remote cache).
      ui.rerender(<Harness showB={false} />)
      await waitFor(() => expect(ui.getAllByRole('gridcell', { name: /value long text/ })).toHaveLength(1))
      expect(initialCache.closed).toBe(true)

      // Clear spy calls recorded before the unmount:
      seenCaches.length = 0

      // Type in search box to trigger an overlay-only update without a base frame.
      const searchBox = ui.getByRole('searchbox', { name: 'Search document' })
      fireEvent.change(searchBox, { target: { value: 'long' } })

      // Wait for search hits to appear:
      await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toMatch(/1 of 1/))

      // The find update must have consulted overflow geometry:
      expect(seenCaches.length).toBeGreaterThan(0)
      // Every call during the find update must see an open cache (never a closed one):
      expect(seenCaches.some(c => c.closed)).toBe(false)
      expect(ui.getByTestId('a').textContent).not.toContain('Text overflow occupancy owner is closed')
      expect(f.remote.overflowFor().closed).toBe(false)
      ui.unmount()
    } finally {
      overflowSpy.mockRestore()
      await f.remote.close()
      await f.client.dispose()
    }
  })

  it('an M viewer on a D remote and a D viewer on an M remote use the remote profile for shared ledger caches', async () => {
    const dFixture = await opened('D')
    const mFixture = await opened('M')
    try {
      const dCache = dFixture.remote.overflowFor()
      const mCache = mFixture.remote.overflowFor()
      expect(dCache.pool.limitBytes).toBe(64 * 1024 * 1024)
      expect(mCache.pool.limitBytes).toBe(32 * 1024 * 1024)

      // Mount an M viewer on the D remote:
      const uiM = render(<InteractiveSheetDoc doc={dFixture.doc} source={dFixture.remote} profile="M" />)
      await waitFor(() => expect(dFixture.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(4096))
      // Shared overflow cache on the D remote must retain D budget (64 MiB), NOT replace/close with M (32 MiB)
      expect(dFixture.remote.overflowFor().pool.limitBytes).toBe(64 * 1024 * 1024)
      expect(dCache.closed).toBe(false)
      uiM.unmount()

      // Mount a D viewer on the M remote:
      const uiD = render(<InteractiveSheetDoc doc={mFixture.doc} source={mFixture.remote} profile="D" />)
      await waitFor(() => expect(mFixture.remote.mainLedger.snapshot().byPool.layout).toBeGreaterThan(4096))
      // Shared overflow cache on the M remote must retain M budget (32 MiB), NOT replace/close with D (64 MiB)
      expect(mFixture.remote.overflowFor().pool.limitBytes).toBe(32 * 1024 * 1024)
      expect(mCache.closed).toBe(false)
      uiD.unmount()
    } finally {
      await dFixture.remote.close(); await dFixture.client.dispose()
      await mFixture.remote.close(); await mFixture.client.dispose()
    }
  })
})

async function opened(profile: 'D' | 'M' = 'D') {
  const doc = workbookOf([{ sheetId: 's', cells: [{ row: 0, col: 0, value: 'long text' }, { row: 0, col: 3, value: 'block' }] }])
  const source = createObjectWorkbookSource(doc, 'd')
  const kernel = new ViewerKernel({ parse: async () => doc, openSource: async () => source }, { yield: async () => {} })
  const client = new ViewerClient(connectInProcess(kernel))
  await client.send({ kind: 'open', requestId: client.nextRequestId(), docId: 'd', source: { kind: 'object', document: doc }, profile })
  const remote = new RemoteWorkbook(client, { docId: 'd', valuePolicy: source.valuePolicy, sheets: source.sheets, diagnostics: [], profile, formatting: source.formatting?.(), layouts: new Map(source.sheets.map(sheet => [sheet.sheetId, kernel.session('d')!.layout(sheet.sheetId)])) })
  return { doc, client, remote }
}
