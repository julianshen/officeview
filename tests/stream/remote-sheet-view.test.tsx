/**
 * Stage 4a Task 14 (S4 §6.1, §6.4, §6.11): a tiles-mode stream opens in the interactive sheet viewer.
 * `RemoteDocView` routes a document whose `sheets` is set to the lazy `RemoteSheetView`, which mounts
 * `InteractiveSheetDoc` over `streamShellDoc` + `StreamWorkbook`:
 * - no selection, find, copy, or sort and filter menus; no local print (print is the server's static units);
 * - the a11y grid carries positions only, and the parent §15 notice appears once;
 * - a units-mode manifest (no `sheetMode`) still renders the Stage 3 unit canvases;
 * - each sheet activation fetches its `sheet-layout` once, and a parent re-render never remounts the viewer;
 * - a `layout-too-large` refusal shows the load-failure notice (spec §6.4, F21).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import RemoteDocView, { TEXT_UNAVAILABLE } from '../../src/components/remote/RemoteDocView'
import { StreamSessionError } from '../../src/stream/remote/errors'
import type { PrintedPage, RemoteOfficeDoc, RemoteSheets } from '../../src/stream/remote/remote-doc'
import type { TilesRequest, TilesResult } from '../../src/stream/remote/sheet-channel'
import type { StreamSheetLayout } from '../../src/stream/sheet-wire'

const pr = vi.hoisted(() => ({ printRasters: vi.fn(async (..._a: unknown[]) => null) }))
vi.mock('../../src/components/paged/printDocument', () => ({ printRasters: pr.printRasters }))

function wire(sheetId: string): StreamSheetLayout {
  return {
    sheetId, viewGen: 0, defaults: { rowPx: 20, colPx: 64 }, rows: [], cols: [], collapsedRows: [], collapsedCols: [],
    merges: [], filters: [], sort: null, filterSummary: [], authorHidden: false, counts: { visibleRows: 40, totalRows: 40 },
    used: { r0: 0, c0: 0, r1: 39, c1: 9 },
  } as StreamSheetLayout
}

type LayoutImpl = (sheetId: string, signal?: AbortSignal) => Promise<StreamSheetLayout>
function fakeDoc(o: { tiles?: boolean; layout?: LayoutImpl; print?: boolean } = {}) {
  const tilesMode = o.tiles ?? true
  const count = 2
  const list = Array.from({ length: count }, () => ({ w: 300, h: 200 }))
  const manifest = {
    v: 1, docRef: 'doc-ref:sheet-view', name: 'book.xlsx', kind: 'xlsx', units: { count, list },
    can: { search: true, copy: false, print: o.print ?? true, textA11y: false, present: false, notes: false, links: false },
    caps: { copy: null, a11y: 'off', print: { maxPages: null, maxSessionPages: null } },
    recorder: { scalesBuckets: [1, 2] }, notices: [],
    ...(tilesMode ? {
      sheetMode: 'tiles', sheets: [{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }],
      tiles: { px: 512, zoomSteps: [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4], dprs: [1, 2], palettes: ['light', 'dark'], maxBatch: 64 },
    } : {}),
  }
  const tileCalls: TilesRequest[] = []
  const sheets: RemoteSheets | undefined = tilesMode ? {
    list: [{ sheetId: '0', name: 'First' }, { sheetId: '1', name: 'Second' }],
    layout: vi.fn(o.layout ?? (async (sheetId: string) => wire(sheetId))),
    // Held open until cancelled: these tests are about the mount, not the tile loop (Task 13 covers it).
    tiles: vi.fn((req: TilesRequest, _onTile: unknown, signal?: AbortSignal): Promise<TilesResult> => {
      tileCalls.push(req)
      return new Promise(resolve => signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true }))
    }),
  } : undefined
  const pages = (from: number, to: number): PrintedPage[] =>
    Array.from({ length: to - from + 1 }, (_, i) => ({ unitIndex: from + i, widthPx: 300, heightPx: 200, scale: 2 as const, png: new Blob(['png']) }))
  const doc = {
    kind: 'remote', manifest,
    onManifestChange: vi.fn(() => () => {}),
    units: async () => list.map(u => ({ widthPx: u.w, heightPx: u.h })),
    paint: vi.fn(async () => {}),
    search: vi.fn(async () => ({ total: 0, truncated: false })),
    select: vi.fn(async () => ({ units: [], paragraphs: 0 })),
    copy: vi.fn(async () => ''),
    a11y: vi.fn(async () => []),
    print: vi.fn(async (r: { from: number; to: number }) => pages(r.from, r.to)),
    paintNotes: vi.fn(async () => {}),
    close: vi.fn(),
    ended: new Promise(() => {}),
    ...(sheets ? { sheets } : {}),
  }
  return { doc, sheets, tileCalls }
}
type Fake = ReturnType<typeof fakeDoc>['doc']
const view = (doc: Fake, props: Record<string, unknown> = {}) => render(<RemoteDocView document={doc as unknown as RemoteOfficeDoc} {...props} />)
const grid = (v: ReturnType<typeof render>) => v.findByRole('grid', {}, { timeout: 20_000 })
const occurrences = (text: string, needle: string) => text.split(needle).length - 1

let winPrint: ReturnType<typeof vi.fn>
beforeEach(() => {
  pr.printRasters.mockReset(); pr.printRasters.mockResolvedValue(null)
  winPrint = vi.fn()
  vi.spyOn(window, 'print').mockImplementation(winPrint as () => void)
})
afterEach(() => {
  // Neither the viewer's screen nor a local print container is ever printed on a stream.
  expect(winPrint).not.toHaveBeenCalled()
  cleanup(); vi.restoreAllMocks()
})

// The first mount loads the viewer chunk (lazy), which takes a while under jsdom.
describe('RemoteSheetView (tiles mode)', { timeout: 30_000 }, () => {
  it('tiles mode mounts the interactive viewer', async () => {
    const { doc } = fakeDoc()
    const v = view(doc)
    await grid(v)
    expect(v.container.querySelector('[data-unit]')).toBeNull()
    // No Find and no Copy control (S4 §6.11: 4a mounts neither).
    expect(v.container.querySelector('[data-officeview-ui="sheet-search"]')).toBeNull()
    expect(v.queryByRole('searchbox')).toBeNull()
    expect(v.queryByRole('button', { name: /copy/i })).toBeNull()
    expect(v.queryByRole('button', { name: /find/i })).toBeNull()
    // The parent §15 notice: exactly one status region says it, and it is said once.
    await waitFor(() => expect(v.getAllByRole('status').filter(s => s.textContent === TEXT_UNAVAILABLE)).toHaveLength(1))
    expect(occurrences(v.container.textContent ?? '', TEXT_UNAVAILABLE)).toBe(1)
    // Units-mode text and paint ops are never sent on a tiles session (R6).
    expect(doc.paint).not.toHaveBeenCalled()
    expect(doc.a11y).not.toHaveBeenCalled()
    expect(doc.search).not.toHaveBeenCalled()
    expect(doc.select).not.toHaveBeenCalled()
    expect(doc.copy).not.toHaveBeenCalled()
  })

  it('a11y grid has positions, no text', async () => {
    const { doc } = fakeDoc()
    const v = view(doc)
    await grid(v)
    await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
    const cells = [...v.container.querySelectorAll<HTMLElement>('[role="gridcell"]')]
    for (const cell of cells) {
      expect(Number(cell.getAttribute('aria-rowindex'))).toBeGreaterThanOrEqual(2)
      expect(Number(cell.getAttribute('aria-colindex'))).toBeGreaterThanOrEqual(2)
      expect(cell.textContent).toBe('')
      expect(cell.getAttribute('aria-description')).toBeNull()
    }
    // The row index a cell carries is its row's.
    const first = cells[0]!
    expect(first.getAttribute('aria-rowindex')).toBe(first.closest('[role="row"]')!.getAttribute('aria-rowindex'))
  })

  it('no selection: a click or an arrow key selects nothing', async () => {
    const { doc } = fakeDoc()
    const v = view(doc)
    const g = await grid(v)
    await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
    fireEvent.click(g, { clientX: 120, clientY: 60 })
    fireEvent.keyDown(g, { key: 'ArrowDown' })
    fireEvent.keyDown(g, { key: 'a', ctrlKey: true })
    expect(g.getAttribute('aria-activedescendant')).toBeNull()
    expect(v.container.querySelectorAll('[role="gridcell"][aria-selected="true"]')).toHaveLength(0)
    // No sort or filter menu in the column headers.
    expect(v.queryByRole('button', { name: /sort|filter/i })).toBeNull()
  })

  it('units mode unchanged', async () => {
    const { doc } = fakeDoc({ tiles: false })
    const v = view(doc)
    await waitFor(() => expect(v.container.querySelectorAll('canvas[data-unit]')).toHaveLength(2))
    expect(v.queryByRole('grid')).toBeNull()
  })

  it('sheet tabs fetch the layout once per activation', async () => {
    const { doc, sheets } = fakeDoc()
    const v = view(doc)
    await grid(v)
    await waitFor(() => expect(sheets!.layout).toHaveBeenCalledTimes(1))
    expect(vi.mocked(sheets!.layout).mock.calls[0]![0]).toBe('0')
    fireEvent.change(v.getByRole('combobox', { name: 'Worksheet' }), { target: { value: '1' } })
    await waitFor(() => expect(vi.mocked(sheets!.layout).mock.calls.filter(c => c[0] === '1')).toHaveLength(1))
    // Later renders of the same activation ask for nothing more.
    await act(async () => { await new Promise(r => setTimeout(r, 50)) })
    expect(vi.mocked(sheets!.layout).mock.calls.filter(c => c[0] === '1')).toHaveLength(1)
    expect(vi.mocked(sheets!.layout).mock.calls.filter(c => c[0] === '0')).toHaveLength(1)
  })

  it('a parent re-render keeps the viewer mounted (one shell document per stream)', async () => {
    const { doc, sheets } = fakeDoc()
    const v = view(doc, { style: { height: 400 } })
    const before = await grid(v)
    await waitFor(() => expect(sheets!.layout).toHaveBeenCalledTimes(1))
    v.rerender(<RemoteDocView document={doc as unknown as RemoteOfficeDoc} style={{ height: 401 }} title="Renamed" />)
    await act(async () => { await new Promise(r => setTimeout(r, 50)) })
    expect(v.getByRole('grid')).toBe(before)
    expect(sheets!.layout).toHaveBeenCalledTimes(1)
  })

  it('a layout-too-large refusal shows the load-failure notice and requests no tile', async () => {
    const { doc, tileCalls } = fakeDoc({ layout: async () => { throw new StreamSessionError('layout-too-large', 'the sheet layout is too large to send') } })
    const v = view(doc)
    await grid(v)
    await waitFor(() => expect(v.container.textContent).toContain('This sheet could not be fully loaded'))
    await act(async () => { await new Promise(r => setTimeout(r, 50)) })
    expect(tileCalls).toHaveLength(0)
  })

  it('print goes through RemoteOfficeDoc.print, the picker counting sheets', async () => {
    const { doc } = fakeDoc()
    const v = view(doc, { title: 'Book' })
    await grid(v)
    // No local print path: no static print container, no local Print button.
    expect(v.container.querySelector('[data-officeview-sheet-print]')).toBeNull()
    const buttons = v.container.querySelectorAll<HTMLButtonElement>('button[data-ov-print-button]')
    expect(buttons).toHaveLength(1)
    fireEvent.click(buttons[0]!)
    await waitFor(() => expect(pr.printRasters).toHaveBeenCalledTimes(1))
    expect(doc.print).toHaveBeenCalledTimes(1)
    expect(doc.print.mock.calls[0]![0]).toEqual({ from: 0, to: 1 })
    // Ctrl/Cmd+P inside the viewer starts the same server print.
    fireEvent.keyDown(v.getByRole('grid'), { key: 'p', ctrlKey: true })
    await waitFor(() => expect(doc.print).toHaveBeenCalledTimes(2))
  })

  it('the capped print picker speaks in sheets', async () => {
    const { doc } = fakeDoc()
    ;(doc.manifest.caps as { print: unknown }).print = { maxPages: 1, maxSessionPages: null }
    const v = view(doc)
    await grid(v)
    fireEvent.click(v.container.querySelector<HTMLButtonElement>('button[data-ov-print-button]')!)
    const dialog = await v.findByRole('dialog', { name: 'Print sheets' })
    expect(dialog.textContent).toContain('Up to 1 sheets per print.')
  })

  it('print off: no Print button, and Ctrl/Cmd+P prints nothing', async () => {
    const { doc } = fakeDoc({ print: false })
    const v = view(doc)
    await grid(v)
    expect(v.container.querySelector('button[data-ov-print-button]')).toBeNull()
    expect(v.queryByRole('button', { name: /print/i })).toBeNull()
    fireEvent.keyDown(v.getByRole('grid'), { key: 'p', ctrlKey: true })
    await act(async () => { await new Promise(r => setTimeout(r, 20)) })
    expect(doc.print).not.toHaveBeenCalled()
  })
})

// Stage 4b Task 10 (P1, S4 §6.11): with `sheetView: 1` the same fake mounts Find and Copy, and selection works.
describe('RemoteSheetView (tiles mode, 4b sheetView)', { timeout: 30_000 }, () => {
  it('4b stream mounts Find and Copy', async () => {
    const { doc, sheets } = fakeDoc()
    ;(doc.manifest as Record<string, unknown>).sheetView = 1
    doc.manifest.can.copy = true
    ;(sheets as { view?: unknown }).view = {
      find: vi.fn(async () => ({ hits: [], total: 0, truncated: false })), copy: vi.fn(async () => ''), a11y: vi.fn(async () => []),
      fit: vi.fn(async () => 64), resize: vi.fn(),
    }
    const v = view(doc)
    const g = await grid(v)
    expect(v.getByRole('searchbox')).toBeTruthy()
    expect(v.getByRole('button', { name: 'Copy selection' })).toBeTruthy()
    await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
    const area = g.firstElementChild as HTMLElement
    fireEvent.click(g, { clientX: parseFloat(area.style.left) + 64 + 10, clientY: parseFloat(area.style.top) + 20 + 5 })
    await waitFor(() => expect(g.getAttribute('aria-activedescendant')).toMatch(/-cell-16385$/))
    expect(g.getAttribute('aria-multiselectable')).toBe('true')
    expect(v.container.querySelectorAll('[role="gridcell"][aria-selected="true"]')).toHaveLength(1)
    // Units-mode text ops are still never sent on a tiles session.
    expect(doc.copy).not.toHaveBeenCalled()
    expect(doc.search).not.toHaveBeenCalled()
  })
})

// Stage 4c Task 10 (P16, S4 §6.11): with `sheetSort: 1` (and its `sortOps`) the column headers carry the sort and
// filter buttons; a 4b session (`sheetView: 1` alone) still mounts none.
describe('RemoteSheetView (tiles mode, 4c sheetSort)', { timeout: 30_000 }, () => {
  const textOps = () => ({
    find: vi.fn(async () => ({ hits: [], total: 0, truncated: false })), copy: vi.fn(async () => ''), a11y: vi.fn(async () => []),
    fit: vi.fn(async () => 64), resize: vi.fn(),
  })
  const sortOps = () => ({
    typedFilters: false, sort: vi.fn(), filter: vi.fn(), order: vi.fn(),
    values: vi.fn(async () => ({ list: 'L', items: [], blanks: { present: false, selected: true }, total: 0, truncated: false })),
    detect: vi.fn(async () => ({ range: { r0: 0, r1: 39, c0: 0, c1: 9, header: 1, totals: 0 }, source: 'region', header: 'heuristic', overrideIgnored: false })),
  })
  it('4c stream mounts the filter buttons', async () => {
    for (const sheetSort of [true, false]) {
      const { doc, sheets } = fakeDoc()
      ;(doc.manifest as Record<string, unknown>).sheetView = 1
      ;(sheets as { view?: unknown }).view = textOps()
      if (sheetSort) {
        ;(doc.manifest as Record<string, unknown>).sheetSort = 1
        ;(sheets as { sortOps?: unknown }).sortOps = sortOps()
      }
      const v = view(doc)
      await grid(v)
      await waitFor(() => expect(v.container.querySelectorAll('[role="gridcell"]').length).toBeGreaterThan(10))
      if (sheetSort) {
        expect(await v.findByRole('button', { name: /^Filter column A / })).toBeTruthy()
        expect(v.getByRole('button', { name: /^Cycle sort column A/ })).toBeTruthy()
      } else {
        await act(async () => { await new Promise(r => setTimeout(r, 50)) })
        expect(v.queryByRole('button', { name: /^Filter column/ })).toBeNull()
        expect(v.queryByRole('button', { name: /sort|filter/i })).toBeNull()
      }
      cleanup()
    }
  })
})
