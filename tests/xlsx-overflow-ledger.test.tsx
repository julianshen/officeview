import { StrictMode } from 'react'
import * as overflowGeometry from '../src/xlsx/view/viewport'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { createObjectWorkbookSource } from '../src/worker/viewerObjectSource'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { createDocumentOverflowCache } from '../src/xlsx/view/documentOverflow'
import { bindOverflowCache, boundOverflowCache, type OverflowCache, OverflowResourceError } from '../src/xlsx/view/overflowCache'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { buildSheetGeometry, cellTextOverflow, planViewport, rowOverflowInput } from '../src/xlsx/view/viewport'
import { workbookOf } from './helpers/xlsx-find'

const doc = () => workbookOf([{ sheetId: 'A', cells: [{ row: 0, col: 0, value: 'text' }, { row: 0, col: 10, value: 'block' }] }])
function geometry() {
  const g = buildSheetGeometry(doc().sheets[0])
  if (g.kind !== 'ready') throw Error('geometry')
  return g
}
afterEach(() => { cleanup(); vi.restoreAllMocks() })

describe('real main ledger owns overflow', () => {
  it('charges the distinct layout owner before allocations, with construction coexistence', () => {
    const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D'), g = geometry()
    const input = cache.get(g, 0)
    expect(input.blockers).toBeInstanceOf(Uint16Array)
    expect(ledger.snapshot().byPool.layout).toBe(cache.bytes)
    expect(ledger.snapshot().peakBytes).toBeGreaterThan(cache.bytes)
    cache.close()
    expect(ledger.bytes).toBe(0)
  })

  it.each(['D', 'M'] as const)('refuses occupancy beyond the unchanged %s shared main budget', profile => {
    const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, profile)
    ledger.set('copy', 'other-operation', MAIN_THREAD_BUDGETS[profile] - 1000)
    expect(() => cache.get(geometry(), 0)).toThrow(OverflowResourceError)
    expect(ledger.snapshot().byPool.layout).toBe(0)
    expect(ledger.bytes).toBe(MAIN_THREAD_BUDGETS[profile] - 1000)
  })

  it('keeps the construction lease charged through a synchronous close request and never revives it', () => {
    const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D'), built = geometry()
    let close = true, chargedAtClose = 0
    const g = { ...built, metrics: { ...built.metrics, cols: { ...built.metrics.cols, start: (col: number) => {
      if (close) { close = false; cache.close(); chargedAtClose = ledger.snapshot().byPool.layout }
      return built.metrics.cols.start(col)
    } } } }
    expect(() => cache.get(g, 0)).toThrow(OverflowResourceError)
    expect(chargedAtClose).toBeGreaterThan(4096)
    expect(ledger.snapshot().byPool.layout).toBe(0)
    expect(cache.bytes).toBe(0)
    expect(cache.pool.totalBytes).toBe(0)
    expect(() => cache.get(g, 0)).toThrow(OverflowResourceError)
  })

  it('distinguishes a closed owner from real quota refusal without reviving either cache', () => {
    const closedLedger = new MainMemoryLedger(), closed = createDocumentOverflowCache(closedLedger, 'D')
    closed.close()
    expect(() => closed.get(geometry(), 0)).toThrow('Text overflow occupancy owner is closed')
    expect(closed.bytes).toBe(0)
    const quotaLedger = new MainMemoryLedger(), quota = createDocumentOverflowCache(quotaLedger, 'D')
    quotaLedger.set('copy', 'pressure', MAIN_THREAD_BUDGETS.D - 1000)
    expect(() => quota.get(geometry(), 0)).toThrow('Text overflow occupancy exceeds the document memory budget')
    expect(quota.closed).toBe(false)
    expect(quota.bytes).toBe(0)
  })

  it('a closed owner is handled by the actual viewport resource path with the correct notice', () => {
    const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D'), g = geometry()
    bindOverflowCache(g, cache)
    cache.close()
    const plan = planViewport(g, { box: { x: 0, y: 0, width: 200, height: 80 }, zoom: 1, deviceScale: 1 }, 'D')
    expect(plan.kind).toBe('coarse')
    expect(plan.diagnostics.find(d => d.code === 'overflow-memory')?.message).toBe('Text overflow occupancy owner is closed')
    expect(cache.bytes).toBe(0)
  })

  it('find catches a real owner close at lookup and announces closure rather than quota', async () => {
    let closed: OverflowCache | undefined
    const solve = overflowGeometry.cellTextOverflow
    vi.spyOn(overflowGeometry, 'cellTextOverflow').mockImplementation((g, source) => {
      if (g?.kind === 'ready') { closed = boundOverflowCache(g)!; closed.close() }
      return solve(g, source)
    })
    const ui = render(<InteractiveSheetDoc doc={doc()} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value text/ })).toBeTruthy())
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'text' } })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toContain('1 of 1'))
    await waitFor(() => expect(ui.container.textContent).toContain('Text overflow occupancy owner is closed'))
    expect(ui.container.textContent).not.toContain('exceeds the document memory budget')
    expect(closed?.closed).toBe(true)
    expect(closed?.bytes).toBe(0)
  })

  it('an unrelated base present keeps a find refusal while its highlight is still missing', async () => {
    vi.spyOn(overflowGeometry, 'cellTextOverflow').mockImplementation(() => { throw new OverflowResourceError() })
    let frames = 0
    const workbook = doc(), onFrame = () => { frames++ }
    const ui = render(<InteractiveSheetDoc doc={workbook} onFrame={onFrame} />)
    await waitFor(() => expect(ui.getByRole('gridcell', { name: /value text/ })).toBeTruthy())
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'text' } })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toContain('1 of 1'))
    await waitFor(() => expect(ui.container.textContent).toContain('Text overflow occupancy exceeds the document memory budget'))
    // A watermark change re-presents the base layer only; the overlay (and its
    // refused highlight) is untouched, so the find notice must stay.
    const before = frames
    ui.rerender(<InteractiveSheetDoc doc={workbook} onFrame={onFrame} watermark={{ text: 'DRAFT' }} />)
    await waitFor(() => expect(frames).toBeGreaterThan(before))
    expect(ui.container.textContent).toContain('Text overflow occupancy exceeds the document memory budget')
  })

  it('uses bound ownership for viewport/find but never retains an unbound geometry memo', () => {
    const g = geometry(), ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D')
    const unbound = rowOverflowInput(g, 0)
    expect(rowOverflowInput(g, 0)).not.toBe(unbound)
    bindOverflowCache(g, cache)
    const input = rowOverflowInput(g, 0)
    expect(input.blockers).toBeInstanceOf(Uint16Array)
    expect(cellTextOverflow(g, { row: 0, col: 0 })).toEqual({ left: 0, right: 9 * 64 })
    expect(rowOverflowInput(g, 0)).toBe(input)
    expect(ledger.snapshot().byPool.layout).toBeGreaterThan(4096)
    cache.close()
  })

  it('returns an explicit coarse resource plan rather than throwing or asserting blank occupancy', () => {
    const g = geometry(), ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'M')
    bindOverflowCache(g, cache)
    ledger.set('copy', 'other-operation', MAIN_THREAD_BUDGETS.M - 1000)
    const plan = planViewport(g, { box: { x: 0, y: 0, width: 200, height: 80 }, zoom: 1, deviceScale: 1 }, 'M')
    expect(plan.kind).toBe('coarse')
    expect(plan.diagnostics.some(d => d.code === 'overflow-memory')).toBe(true)
    expect(ledger.snapshot().byPool.layout).toBe(0)
  })

  it('releases the old cache when a source is replaced while the resident document stays the same', async () => {
    const workbook = doc(), first = createObjectWorkbookSource(workbook, 'first'), second = createObjectWorkbookSource(workbook, 'second')
    let ledger: MainMemoryLedger | undefined
    const onMainLedger = (value: MainMemoryLedger) => { ledger = value }
    const ui = render(<InteractiveSheetDoc doc={workbook} source={first} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(ledger?.snapshot().byPool.layout ?? 0).toBeGreaterThan(4096))
    const old = ledger!
    ui.rerender(<InteractiveSheetDoc doc={workbook} source={second} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(ledger).not.toBe(old))
    expect(old.snapshot().byPool.layout).toBe(0)
    await waitFor(() => expect(ledger!.snapshot().byPool.layout).toBeGreaterThan(4096))
    ui.unmount()
    expect(ledger!.snapshot().byPool.layout).toBe(0)
  })

  it('the actual viewer displays resource refusal for paint/find and recovers after pressure is released', async () => {
    let ledger: MainMemoryLedger | undefined
    const workbook = doc(), onMainLedger = (value: MainMemoryLedger) => { ledger = value }
    const ui = render(<InteractiveSheetDoc doc={workbook} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(ledger?.snapshot().byPool.layout ?? 0).toBeGreaterThan(4096))
    ledger!.set('copy', 'pressure', MAIN_THREAD_BUDGETS.D)
    ui.rerender(<InteractiveSheetDoc doc={workbook} views={{ A: { revision: 1 } }} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(ui.container.textContent).toContain('Text overflow occupancy exceeds the document memory budget'))
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'text' } })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toContain('1 of 1'))
    expect(ui.getByRole('gridcell', { name: /value text/ })).toBeTruthy()
    ledger!.release('copy', 'pressure')
    ui.rerender(<InteractiveSheetDoc doc={workbook} views={{ A: { revision: 2 } }} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(ui.container.textContent).not.toContain('Text overflow occupancy exceeds the document memory budget'))
    expect(ledger!.snapshot().byPool.layout).toBeGreaterThan(4096)
  })

  it('rebinds a fresh cache after StrictMode cleanup rather than reviving the closed owner', async () => {
    let ledger: MainMemoryLedger | undefined
    const ui = render(<StrictMode><InteractiveSheetDoc doc={doc()} onMainLedger={value => { ledger = value }} /></StrictMode>)
    await waitFor(() => expect(ledger?.snapshot().byPool.layout ?? 0).toBeGreaterThan(4096))
    expect(ui.container.textContent).not.toContain('Text overflow occupancy exceeds the document memory budget')
    ui.unmount()
    expect(ledger!.snapshot().byPool.layout).toBe(0)
  })

  it('the actual resident viewer charges and releases overflow on profile and document replacement/unmount', async () => {
    const events: Array<{ ledger: MainMemoryLedger; owner: string; bytes: number }> = []
    let current: MainMemoryLedger | undefined
    const onMainLedger = (ledger: MainMemoryLedger) => {
      current = ledger
      const original = ledger.set.bind(ledger)
      ledger.set = (pool, owner, bytes) => { if (pool === 'layout' && owner.startsWith('overflow:')) events.push({ ledger, owner, bytes }); original(pool, owner, bytes) }
    }
    const workbook = doc()
    const ui = render(<InteractiveSheetDoc doc={workbook} onMainLedger={onMainLedger} />)
    await waitFor(() => expect(events.some(e => e.bytes > 0)).toBe(true))
    const oldOwner = events.find(e => e.bytes > 0)!.owner
    ui.rerender(<InteractiveSheetDoc doc={workbook} profile="M" onMainLedger={onMainLedger} />)
    await waitFor(() => expect(events.some(e => e.owner === oldOwner && e.bytes === 0)).toBe(true))
    await waitFor(() => expect(events.some(e => e.owner !== oldOwner && e.bytes > 0)).toBe(true))
    const firstLedger = current!
    ui.rerender(<InteractiveSheetDoc doc={doc()} profile="M" onMainLedger={onMainLedger} />)
    await waitFor(() => expect(current).not.toBe(firstLedger))
    expect(firstLedger.snapshot().byPool.layout).toBe(0)
    // Find's overlay uses the same bound geometry and cannot create another owner.
    await waitFor(() => expect(events.some(e => e.ledger === current && e.bytes > 0)).toBe(true))
    const owners = new Set(events.filter(e => e.ledger === current && e.bytes > 0).map(e => e.owner))
    fireEvent.change(ui.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'text' } })
    await waitFor(() => expect(ui.getByTestId('officeview-search-status').textContent).toContain('1 of 1'))
    expect(new Set(events.filter(e => e.ledger === current && e.bytes > 0).map(e => e.owner))).toEqual(owners)
    ui.unmount()
    expect(current!.snapshot().byPool.layout).toBe(0)
  })
})
