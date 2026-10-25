import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { SheetNavigator } from '../src/components/xlsx/SheetNavigator'
import { ViewerStatus } from '../src/components/xlsx/ViewerStatus'
import { gridSelectionOverlay, paintGridSelectionOverlay } from '../src/core/overlay'
import { SourceSelection, activeAnnouncement } from '../src/xlsx/view/selection'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import type { CellRead, SourceStamp, ReadOptions } from '../src/xlsx/view/contracts'
import type { SourceCell } from '../src/xlsx/view/coordinates'

afterEach(cleanup)
const sheets = [{ id: 'a', name: 'Data <img src=x>' }, { id: 'b', name: 'Second' }]
describe('controlled S6.1 navigator/status shell', () => {
  it('does not split surrogate pairs in announcements, values or formulas (review 7)', async () => {
    const text = 'a'.repeat(255) + '😀', p = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 24 }), cols: createAxisMetrics({ logicalCount: 4, defaultSizeCss: 80 }) })
    expect(activeAnnouncement(p, { row: 0, col: 0 }, text)).toBe('A1, value ' + 'a'.repeat(255))
    const source = { stamp: { docId: 'doc', sheetId: 'a', valueGen: 1 }, async readCell(): Promise<CellRead> { return { kind: 'cell', cell: { id: 0, row: 0, col: 0, styleIndex: 0, value: text, formula: text } } } }
    const ui = render(<ViewerStatus records={{ visible: 1, total: 1 }} inspection={{ target: { row: 0, col: 0 }, hiddenReason: 'authored-hidden', source }} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('formula'))
    expect(ui.getByRole('status').textContent).not.toContain('\ud83d')
  })
  it('renders a safe notice for invalid counts instead of throwing (review 12)', () => {
    const ui = render(<ViewerStatus records={{ visible: 3, total: 2 }} />)
    expect(ui.getByRole('status').textContent).toContain('Record counts unavailable')
  })
  it('does not reread the same inspection when a parent recreates its source wrapper (review 12)', async () => {
    const stamp = { docId: 'doc', sheetId: 'a', valueGen: 1 }, readCell = vi.fn(async (): Promise<CellRead> => ({ kind: 'blank' }))
    const inspection = () => ({ target: { row: 0, col: 0 }, hiddenReason: 'authored-hidden', source: { stamp, readCell } })
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={inspection()} />)
    await waitFor(() => expect(readCell).toHaveBeenCalledOnce())
    ui.rerender(<ViewerStatus records={{ visible: 2, total: 2 }} inspection={inspection()} />)
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(readCell).toHaveBeenCalledOnce()
  })
  it('emits explicit display/source Go To requests through full capacity', () => {
    const navigate = vi.fn(), change = vi.fn()
    const ui = render(<SheetNavigator documentId="doc" sheets={sheets} activeSheetId="a" onSheetChange={change} onNavigate={navigate} />)
    fireEvent.change(ui.getByLabelText('Go To cell'), { target: { value: 'XFD1048576' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    expect(navigate).toHaveBeenLastCalledWith({ addressSpace: 'display', at: { rowSlot: 1048575, col: 16383 } })
    fireEvent.change(ui.getByLabelText('Address space'), { target: { value: 'source' } })
    fireEvent.change(ui.getByLabelText('Go To cell'), { target: { value: '$a$2' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    expect(navigate).toHaveBeenLastCalledWith({ addressSpace: 'source', at: { row: 1, col: 0 } })
    fireEvent.change(ui.getByLabelText('Worksheet'), { target: { value: 'b' } })
    expect(change).toHaveBeenCalledWith('b')
    expect(ui.container.querySelector('img')).toBeNull()
  })
  it('a disabled navigator is visibly inert, never silently accepting a Go To, and navigates once re-enabled', () => {
    const navigate = vi.fn(), notice = vi.fn()
    const props = { documentId: 'doc', sheets, activeSheetId: 'a', onSheetChange: () => {}, onNavigate: navigate, onNotice: notice }
    const ui = render(<SheetNavigator {...props} disabled />)
    for (const name of ['Worksheet', 'Address space', 'Go To cell']) expect((ui.getByLabelText(name) as HTMLInputElement).disabled).toBe(true)
    const button = ui.getByRole('button', { name: 'Go To' }) as HTMLButtonElement
    expect(button.disabled).toBe(true)
    // Even a submit that bypasses the disabled button (a host form submit) is a no-op, with no notice.
    fireEvent.submit(button.form!)
    expect(navigate).not.toHaveBeenCalled()
    expect(notice).not.toHaveBeenCalled()
    ui.rerender(<SheetNavigator {...props} disabled={false} />)
    expect(button.disabled).toBe(false)
    fireEvent.change(ui.getByLabelText('Go To cell'), { target: { value: 'B2' } })
    fireEvent.click(button)
    expect(navigate).toHaveBeenCalledWith({ addressSpace: 'display', at: { rowSlot: 1, col: 1 } })
  })
  it('diagnoses invalid addresses without issuing navigation and isolates viewer gestures', () => {
    const navigate = vi.fn(), notice = vi.fn(), parent = vi.fn()
    const ui = render(<div onPointerDown={parent} onDoubleClick={parent} onKeyDown={parent}><SheetNavigator documentId="doc" sheets={sheets} activeSheetId="a" onSheetChange={() => {}} onNavigate={navigate} onNotice={notice} /></div>)
    const input = ui.getByLabelText('Go To cell')
    fireEvent.change(input, { target: { value: 'XFE1048577' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    expect(navigate).not.toHaveBeenCalled()
    expect(notice).toHaveBeenCalledWith('Enter a valid cell address through XFD1048576.')
    expect(input.getAttribute('aria-invalid')).toBe('true')
    fireEvent.pointerDown(input)
    fireEvent.doubleClick(input)
    fireEvent.keyDown(input, { key: '0' })
    expect(parent).not.toHaveBeenCalled()
    fireEvent.keyDown(input, { key: 'f', ctrlKey: true })
    expect(parent).toHaveBeenCalledOnce()
    expect(ui.container.querySelector('[aria-live]')).toBeNull()
  })
  it('reports only resolved records with one polite status region', () => {
    const ui = render(<ViewerStatus records={{ visible: 2, total: 3 }} announcement="A2, value 10; original A3" />)
    expect(ui.getByRole('status').textContent).toContain('2 of 3 records')
    expect(ui.getByRole('status').textContent).toContain('A2, value 10; original A3')
    expect(ui.getByRole('status').getAttribute('aria-live')).toBe('polite')
    ui.rerender(<ViewerStatus records={{ visible: 0, total: 0 }} />)
    expect(ui.getByRole('status').textContent).toContain('0 records')
    expect(ui.getByRole('status').textContent).not.toContain('1048576')
  })
  it('keeps operational navigation failures distinct from invalid addresses', () => {
    const notice = vi.fn()
    const ui = render(<SheetNavigator documentId="doc" sheets={sheets} activeSheetId="a" onSheetChange={() => {}} onNavigate={() => { throw new Error('Navigation is loading; retry.') }} onNotice={notice} />)
    fireEvent.change(ui.getByLabelText('Go To cell'), { target: { value: 'A1' } })
    fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
    expect(notice).toHaveBeenCalledWith('Navigation is loading; retry.')
    expect(ui.getByLabelText('Go To cell').getAttribute('aria-invalid')).toBe('false')
  })
  it('reads an exact hidden source cell and keeps cell/formula strings as safe text', async () => {
    const stamp: SourceStamp = { docId: 'doc', sheetId: 'a', valueGen: 1 }
    const readCell = vi.fn(async (): Promise<CellRead> => ({ kind: 'cell', cell: { id: 32768, row: 2, col: 0, styleIndex: 0, value: '<img src=x>', formula: 'SUM(A1:A2)' } }))
    const source = { stamp, readCell }
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={{ target: { row: 2, col: 0 }, hiddenReason: 'Hidden row', source }} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('<img src=x>'))
    expect(readCell).toHaveBeenCalledWith({ row: 2, col: 0 }, expect.objectContaining({ stamp, signal: expect.any(AbortSignal) }))
    expect(ui.getByRole('status').textContent).toContain('Original cell A3')
    expect(ui.getByRole('status').textContent).toContain('SUM(A1:A2)')
    expect(ui.container.querySelector('img')).toBeNull()
  })
  it('cancels a hidden inspection on sheet changes and drops late replies', async () => {
    let resolve!: (read: CellRead) => void
    const stamp = { docId: 'doc', sheetId: 'a', valueGen: 1 }, readCell = vi.fn((_at: SourceCell, _options: ReadOptions) => new Promise<CellRead>(r => { resolve = r }))
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={{ target: { row: 2, col: 0 }, hiddenReason: 'Hidden row', source: { stamp, readCell } }} />)
    await waitFor(() => expect(readCell).toHaveBeenCalledOnce())
    const signal = readCell.mock.calls[0]?.[1]?.signal
    ui.rerender(<ViewerStatus records={{ visible: 0, total: 0 }} announcement="Second sheet" />)
    expect(signal?.aborted).toBe(true)
    resolve({ kind: 'cell', cell: { id: 32768, row: 2, col: 0, styleIndex: 0, value: 'old secret' } })
    await new Promise(r => setTimeout(r, 0))
    expect(ui.getByRole('status').textContent).not.toContain('old secret')
  })
  it('diagnoses a synchronous source read failure instead of crashing the shell', async () => {
    const source = { stamp: { docId: 'doc', sheetId: 'a', valueGen: 1 }, readCell(): Promise<CellRead> { throw new Error('Page unavailable') } }
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={{ target: { row: 2, col: 0 }, hiddenReason: 'Hidden row', source }} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('page-error'))
  })
  it('rejects a hidden-cell payload with another source identity', async () => {
    const source = { stamp: { docId: 'doc', sheetId: 'a', valueGen: 1 }, async readCell(): Promise<CellRead> { return { kind: 'cell', cell: { id: 49152, row: 3, col: 0, styleIndex: 0, value: 'wrong cell' } } } }
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={{ target: { row: 2, col: 0 }, hiddenReason: 'Hidden row', source }} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('page-error'))
    expect(ui.getByRole('status').textContent).not.toContain('wrong cell')
  })
  it('drops a stale value owner and rereads the hidden target under the new owner', async () => {
    let resolveOld!: (read: CellRead) => void
    const source = { stamp: { docId: 'doc', sheetId: 'a', valueGen: 1 }, readCell: vi.fn((_at: SourceCell, options: ReadOptions): Promise<CellRead> => options.stamp.valueGen === 1 ? new Promise(r => { resolveOld = r }) : Promise.resolve({ kind: 'cell', cell: { id: 32768, row: 2, col: 0, styleIndex: 0, value: 'fresh value' } })) }
    const ui = render(<ViewerStatus records={{ visible: 1, total: 2 }} inspection={{ target: { row: 2, col: 0 }, hiddenReason: 'Hidden row', source }} />)
    await waitFor(() => expect(source.readCell).toHaveBeenCalledOnce())
    source.stamp = { ...source.stamp, valueGen: 2 }
    resolveOld({ kind: 'cell', cell: { id: 32768, row: 2, col: 0, styleIndex: 0, value: 'old secret' } })
    await waitFor(() => expect(ui.getByRole('status').textContent).toContain('fresh value'))
    expect(ui.getByRole('status').textContent).not.toContain('old secret')
    expect(source.readCell).toHaveBeenCalledTimes(2)
  })
})

describe('viewport-only green grid overlay seam', () => {
  it('draws only outer selection borders, cancelling shared cell edges (review 9)', () => {
    const selection = new SourceSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 0 }] }, cols: [{ c0: 0, c1: 1 }] }])
    const records = [{ source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 10, height: 10 } }, { source: { row: 0, col: 1 }, bounds: { x: 10, y: 0, width: 10, height: 10 } }]
    const geometry = gridSelectionOverlay(selection, records, { x: -1, y: -1, width: 30, height: 30 }, 1) as unknown as { fills: unknown[]; borders: { x0: number; y0: number; x1: number; y1: number }[] }
    expect(geometry.borders).toEqual(expect.arrayContaining([{ x0: 0, y0: 0, x1: 20, y1: 0 }, { x0: 0, y0: 10, x1: 20, y1: 10 }, { x0: 0, y0: 0, x1: 0, y1: 10 }, { x0: 20, y0: 0, x1: 20, y1: 10 }]))
    expect(geometry.borders).toHaveLength(4)
    const clipped = gridSelectionOverlay(selection, records, { x: 5, y: 2, width: 10, height: 6 }, 1) as unknown as typeof geometry
    expect(clipped.borders).toEqual([])
    expect(clipped.fills).toHaveLength(2)
  })
  it('cancels floating-point shared edges on the device-pixel lattice (N4)', () => {
    const selected = { contains: () => true }
    const viewport = { x: -1, y: -1, width: 20, height: 20 }
    for (const pixelScale of [10, 20]) {
      const horizontal = gridSelectionOverlay(selected, [
        { source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 0.1 + 0.2, height: 10 } },
        { source: { row: 0, col: 1 }, bounds: { x: 0.3, y: 0, width: 10, height: 10 } },
      ], viewport, pixelScale)
      expect(horizontal.borders).toHaveLength(4)
      expect(horizontal.borders.some(edge => edge.x0 === 0.3 && edge.x1 === 0.3)).toBe(false)
      const vertical = gridSelectionOverlay(selected, [
        { source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 10, height: 0.1 + 0.2 } },
        { source: { row: 1, col: 0 }, bounds: { x: 0, y: 0.3, width: 10, height: 10 } },
      ], viewport, pixelScale)
      expect(vertical.borders).toHaveLength(4)
      expect(vertical.borders.some(edge => edge.y0 === 0.3 && edge.y1 === 0.3)).toBe(false)
      // Viewport cuts must still not acquire a false outline after snapping.
      expect(gridSelectionOverlay(selected, [
        { source: { row: 0, col: 0 }, bounds: { x: 0, y: 0, width: 10, height: 10 } },
      ], { x: 1, y: 1, width: 2, height: 2 }, pixelScale).borders).toEqual([])
    }
  })
  it('requires S5 callers to pass pixelScale (zoom × DPR) explicitly', () => {
    const selected = { contains: () => true }
    // @ts-expect-error pixelScale has no default; S5 must pass zoom × DPR.
    expect(() => gridSelectionOverlay(selected, [], { x: 0, y: 0, width: 1, height: 1 })).toThrow(RangeError)
  })
  it('paints at least one device pixel for selected cells smaller than a pixel at low zoom', () => {
    const selected = { contains: () => true }
    const tiny = gridSelectionOverlay(selected, [{ source: { row: 0, col: 0 }, bounds: { x: 1, y: 1, width: 0.04, height: 0.02 } }], { x: 0, y: 0, width: 10, height: 10 }, 10)
    expect(tiny.fills).toHaveLength(1)
    expect(tiny.fills[0]).toMatchObject({ x: 1, y: 1 })
    expect(tiny.fills[0].width * 10).toBeCloseTo(1)
    expect(tiny.fills[0].height * 10).toBeCloseTo(1)
    expect(tiny.borders).toHaveLength(4)
    // Sub-pixel neighbours coalesce into one device-pixel group, not stacked fills.
    const group = gridSelectionOverlay(selected, [
      { source: { row: 0, col: 0 }, bounds: { x: 1, y: 1, width: 0.02, height: 0.02 } },
      { source: { row: 0, col: 1 }, bounds: { x: 1.02, y: 1, width: 0.02, height: 0.02 } },
    ], { x: 0, y: 0, width: 10, height: 10 }, 10)
    expect(group.fills).toHaveLength(1)
    expect(group.borders).toHaveLength(4)
    // Zero-size (hidden) records still paint nothing.
    expect(gridSelectionOverlay(selected, [{ source: { row: 0, col: 0 }, bounds: { x: 1, y: 1, width: 0, height: 5 } }], { x: 0, y: 0, width: 10, height: 10 }, 10)).toEqual({ fills: [], borders: [] })
  })
  it('clips selected viewport records and keeps the blue text selection separate', () => {
    const selection = new SourceSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 3, r1: 3 }] }, cols: [{ c0: 0, c1: 0 }] }])
    const rects = gridSelectionOverlay(selection, [{ source: { row: 3, col: 0 }, bounds: { x: -5, y: 2, width: 20, height: 24 } }, { source: { row: 1, col: 0 }, bounds: { x: 15, y: 2, width: 20, height: 24 } }], { x: 0, y: 0, width: 10, height: 10 }, 1)
    expect(rects.fills).toEqual([{ x: 0, y: 2, width: 10, height: 8 }])
    expect(rects.borders).toEqual([{ x0: 0, y0: 2, x1: 10, y1: 2 }])
    const ctx = { save: vi.fn(), restore: vi.fn(), fillRect: vi.fn(), strokeRect: vi.fn(), beginPath: vi.fn(), moveTo: vi.fn(), lineTo: vi.fn(), stroke: vi.fn(), fillStyle: '', strokeStyle: '', lineWidth: 0 }
    paintGridSelectionOverlay(ctx as unknown as CanvasRenderingContext2D, rects, 2)
    expect(ctx.strokeStyle).toBe('#008000')
    expect(ctx.lineWidth).toBe(1)
    expect(ctx.strokeRect).not.toHaveBeenCalled()
    expect(ctx.moveTo).toHaveBeenCalledWith(0, 2)
    expect(ctx.lineTo).toHaveBeenCalledWith(10, 2)
    expect(ctx.restore).toHaveBeenCalledOnce()
  })
})
