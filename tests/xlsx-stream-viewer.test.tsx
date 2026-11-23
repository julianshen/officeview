/**
 * Stage 4a Task 13 (S4 §6.5), mounted: `InteractiveSheetDoc` over a `StreamWorkbook` paints server tiles only.
 * - F17: at a fractional dpr (1.5) the tiles are requested at bucket 2, the pane backing store is the CSS size × 2,
 *   and the pane carries exactly one CSS scale (its CSS size; no extra transform).
 * - Hidden rows (carried from Task 12): the axis already collapses the hidden rows, and `view.visibility` hides them
 *   again from `layout().hidden`; the rows and the requested tiles must sit exactly where the server's own geometry
 *   puts them (a double count would shift every row below the run by the hidden height).
 * - Wheel zoom on a stream settles to a step: nothing is requested at an off-step zoom.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { StreamWorkbook, streamShellDoc } from '../src/components/xlsx/streamWorkbook'
import type { RemoteSheets } from '../src/stream/remote/remote-doc'
import type { DecodedTile, TilesRequest, TilesResult } from '../src/stream/remote/sheet-channel'
import { decodeSheetLayout, encodeSheetLayout, type StreamSheetLayout } from '../src/stream/sheet-wire'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildSheetGeometry, type SheetGeometry } from '../src/xlsx/view/viewport'
import { SHEET_ZOOM_STEPS } from '../src/xlsx/view/sheetZoom'
import { tileRangeFor } from '../src/xlsx/view/tileCache'
import { streamSheetGeometry } from '../src/xlsx/view/streamGeometry'
import { buildStreamLayout } from '../server/src/sheet/layout'
import type { SheetSnapshot } from '../server/src/sheet/snapshot'
import { xlsxBytes, repeat } from './helpers/xlsx-package'

vi.mock('../src/xlsx/view/renderViewport', async original => {
  const actual = await original<typeof import('../src/xlsx/view/renderViewport')>()
  return { ...actual, renderSheetViewport: vi.fn(actual.renderSheetViewport) }
})
const { renderSheetViewport } = await import('../src/xlsx/view/renderViewport')

const savedDpr = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio')
afterEach(() => {
  cleanup()
  if (savedDpr) Object.defineProperty(window, 'devicePixelRatio', savedDpr)
  vi.clearAllMocks()
})
const setDpr = (dpr: number) => Object.defineProperty(window, 'devicePixelRatio', { configurable: true, get: () => dpr })

/** Rows 3–5 hidden, row 7 tall, row 12 hidden, column B hidden: the server's own geometry is the reference. */
async function serverSheet() {
  const sheetData = repeat(40, i => {
    const r = i + 1
    const attrs = r >= 3 && r <= 5 || r === 12 ? ' hidden="1"' : r === 7 ? ' ht="45" customHeight="1"' : ''
    return `<row r="${r}"${attrs}><c r="A${r}"><v>${r}</v></c></row>`
  })
  const bytes = xlsxBytes({ sheetData, beforeSheetData: '<cols><col min="2" max="2" width="10" hidden="1" customWidth="1"/></cols>' })
  const doc = await parseXlsx(await OfficePackage.load(bytes))
  const sheet = doc.sheets[0]!
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry is a diagnostic')
  // The wire layout exactly as the server builds and sends it (Task 6), viewGen as the session's.
  const built = buildStreamLayout({ sheet, geometry } as unknown as SheetSnapshot, '0')
  const wire: StreamSheetLayout = decodeSheetLayout(JSON.parse(JSON.stringify(encodeSheetLayout({ ...built, viewGen: 4 }))))
  return { server: geometry as SheetGeometry, wire }
}

function mount(wire: StreamSheetLayout, o: { serve?: boolean } = {}) {
  const calls: TilesRequest[] = []
  const sheets: RemoteSheets = {
    list: [{ sheetId: '0', name: 'S' }],
    layout: vi.fn(async () => wire),
    // `serve`: every tile is delivered and the batch settles; else held open until cancelled (the channel resolves an
    // aborted batch `cancelled`).
    tiles: vi.fn((req: TilesRequest, onTile: (t: DecodedTile) => void, signal?: AbortSignal): Promise<TilesResult> => {
      calls.push(req)
      if (o.serve) {
        return new Promise(resolve => setTimeout(() => {
          for (const t of req.tiles) onTile({ sheetIndex: 0, viewGen: req.viewGen, zoomStep: req.zoomStep, dpr: req.dpr, col: t.col, row: t.row, png: new Uint8Array([1]) })
          resolve({ refused: [] })
        }, 5))
      }
      return new Promise(resolve => signal?.addEventListener('abort', () => resolve({ refused: [], cancelled: true }), { once: true }))
    }),
  }
  const workbook = new StreamWorkbook(sheets, 'doc-ref')
  const frames: SheetFrameStats[] = []
  // One document object: a new one is a new document (the viewer remounts, §10.1).
  const doc = streamShellDoc(sheets.list)
  const element = () => <InteractiveSheetDoc doc={doc} source={workbook} onFrame={f => frames.push(f)} />
  const view = render(element())
  return { view, calls, frames, workbook, rerender: () => view.rerender(element()) }
}
const asked = (calls: TilesRequest[]) => calls.flatMap(c => c.tiles.map(t => `${t.col},${t.row}`))

describe('stream viewer (S4 §6.5)', () => {
  it('a fractional dpr (1.5) requests bucket 2; the pane is the CSS size × 2 with exactly one CSS scale', async () => {
    setDpr(1.5)
    const { wire } = await serverSheet()
    const { view, calls, frames } = mount(wire)
    await waitFor(() => expect(calls.length).toBeGreaterThan(0))
    expect(calls.every(c => c.dpr === 2)).toBe(true)
    expect(calls.every(c => c.zoomStep === SHEET_ZOOM_STEPS.indexOf(1))).toBe(true)
    const base = view.getByTestId('officeview-sheet-base') as HTMLCanvasElement
    const cssWidth = parseFloat(base.style.width), cssHeight = parseFloat(base.style.height)
    expect(base.width).toBe(Math.ceil(cssWidth * 2))
    expect(base.height).toBe(Math.ceil(cssHeight * 2))
    // One CSS scale: the CSS size alone (no transform on the pane surfaces).
    expect(base.style.transform).toBe('')
    expect((view.getByTestId('officeview-sheet-overlay') as HTMLCanvasElement).style.transform).toBe('')
    // The tiles asked for are the dpr-2 grid of the pane, not a dpr-1.5 one.
    const frame = frames.at(-1)!
    expect(frame.basePaints).toBe(0)
    expect(vi.mocked(renderSheetViewport)).toHaveBeenCalledTimes(0)
    const g = { origin: { x: 0, y: 0 }, viewport: { width: cssWidth, height: cssHeight }, zoom: 1, dpr: 2 }
    expect(asked(calls)).toEqual(expect.arrayContaining(expectedTiles(tileRangeFor({ ...g, extent: extentOf(wire) }))))
  })

  it('hidden rows: tile placement and row positions match the server layout exactly', async () => {
    setDpr(1)
    const { server, wire } = await serverSheet()
    expect(wire.collapsedRows).toEqual([{ start: 2, end: 4 }, { start: 11, end: 11 }])
    const { view, calls, frames } = mount(wire)
    await waitFor(() => expect(calls.length).toBeGreaterThan(0))
    await waitFor(() => expect(view.container.querySelectorAll('[role="row"][aria-rowindex]').length).toBeGreaterThan(5))
    // Every rendered row sits where the server's geometry puts it, relative to row 1.
    const rows = [...view.container.querySelectorAll<HTMLElement>('[role="row"][aria-rowindex]')]
      .map(el => ({ slot: Number(el.getAttribute('aria-rowindex')) - 2, y: Number(/translate\([^,]+,\s*([-\d.]+)px\)/.exec(el.style.transform)?.[1]), h: parseFloat(el.style.height) }))
      .filter(r => r.slot >= 0 && Number.isFinite(r.y))
    const y0 = rows.find(r => r.slot === 0)!.y
    expect(rows.some(r => r.slot === 5)).toBe(true) // row 6, right below the hidden run, is shown
    expect(rows.some(r => r.slot === 12)).toBe(true)
    // Positions for every row; heights for every row but the last, which the pane edge clips.
    const last = Math.max(...rows.map(r => r.slot))
    for (const r of rows) expect([r.slot, r.y - y0]).toEqual([r.slot, server.metrics.rows.start(r.slot) - server.metrics.rows.start(0)])
    for (const r of rows) if (r.slot !== last) expect([r.slot, r.h]).toEqual([r.slot, server.metrics.rows.size(r.slot)])
    // The tiles asked for are the server-geometry grid of the pane.
    const base = view.getByTestId('officeview-sheet-base') as HTMLCanvasElement
    const viewport = { width: parseFloat(base.style.width), height: parseFloat(base.style.height) }
    const extent = { width: server.metrics.cols.totalCss(), height: server.metrics.rows.totalCss() }
    expect(extentOf(wire)).toEqual(extent)
    expect(asked(calls).sort()).toEqual(expect.arrayContaining(expectedTiles(tileRangeFor({ origin: { x: 0, y: 0 }, viewport, zoom: 1, dpr: 1, extent })).sort()))
    expect(frames.at(-1)!.basePaints).toBe(0)
  })

  it('arrived tiles are drawn from the cache; nothing is requested twice', async () => {
    setDpr(1)
    // jsdom has no createImageBitmap: a canvas stands in for the decoded bitmap.
    vi.stubGlobal('createImageBitmap', vi.fn(async () => Object.assign(document.createElement('canvas'), { width: 512, height: 512, close: () => {} })))
    try {
      const { wire } = await serverSheet()
      const { calls, frames, rerender } = mount(wire, { serve: true })
      await waitFor(() => expect(frames.some(f => f.tileHits > 0 && f.pendingTiles === 0)).toBe(true), { timeout: 4000 })
      // Let the prefetch ring arrive and the local asset lease land, then force a render: still one request per tile.
      await act(async () => { await new Promise(resolve => setTimeout(resolve, 300)) })
      rerender()
      await act(async () => { await new Promise(resolve => setTimeout(resolve, 300)) })
      const last = frames.at(-1)!
      expect(last.pendingTiles).toBe(0)
      expect(last.tileHits).toBeGreaterThan(0)
      expect(frames.every(f => f.basePaints === 0)).toBe(true)
      const counts = new Map<string, number>()
      for (const k of asked(calls)) counts.set(k, (counts.get(k) ?? 0) + 1)
      expect([...counts].filter(([, n]) => n > 1)).toEqual([])
    } finally { vi.unstubAllGlobals() }
  })

  it('wheel zoom on a stream settles to a zoom step; nothing is requested off-step', async () => {
    setDpr(1)
    const { wire } = await serverSheet()
    const { view, calls, frames } = mount(wire)
    await waitFor(() => expect(calls.length).toBeGreaterThan(0))
    const scroller = view.getByTestId('officeview-sheet-scroller')
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] })
    try {
      // 0.002 per px: −100 px → ×1.22 then ×1.22 again ≈ 1.49 live, mid-gesture.
      act(() => { fireEvent.wheel(scroller, { deltaY: -100, ctrlKey: true, clientX: 50, clientY: 50 }) })
      act(() => { fireEvent.wheel(scroller, { deltaY: -100, ctrlKey: true, clientX: 50, clientY: 50 }) })
      const base = view.getByTestId('officeview-sheet-base') as HTMLCanvasElement
      expect(base.style.transform).toMatch(/^matrix\(1\.4/)
      expect(frames.at(-1)!.zoom).toBe(1)
      await act(async () => { vi.advanceTimersByTime(200) })
      expect(base.style.transform).toBe('')
    } finally { vi.useRealTimers() }
    await waitFor(() => expect(frames.at(-1)!.zoom).toBe(1.5))
    await waitFor(() => expect(calls.some(c => c.zoomStep === SHEET_ZOOM_STEPS.indexOf(1.5))).toBe(true))
    expect(calls.every(c => c.zoomStep >= 0 && c.zoomStep < SHEET_ZOOM_STEPS.length)).toBe(true)
    expect(frames.every(f => SHEET_ZOOM_STEPS.includes(f.zoom))).toBe(true)
  })
})

function extentOf(wire: StreamSheetLayout) {
  const { metrics } = streamSheetGeometry(wire)
  return { width: metrics.cols.totalCss(), height: metrics.rows.totalCss() }
}
function expectedTiles(r: { c0: number; c1: number; r0: number; r1: number }): string[] {
  const out: string[] = []
  for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) out.push(`${col},${row}`)
  return out
}
