import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { SheetNavigator } from '../src/components/xlsx/SheetNavigator'
import { ViewerStatus } from '../src/components/xlsx/ViewerStatus'
import { gridSelectionOverlay, paintGridSelectionOverlay } from '../src/core/overlay'
import { SourceSelection } from '../src/xlsx/view/selection'
import type { CellRead, SourceStamp, ReadOptions } from '../src/xlsx/view/contracts'
import type { SourceCell } from '../src/xlsx/view/coordinates'

afterEach(cleanup)
const sheets = [{ id: 'a', name: 'Data <img src=x>' }, { id: 'b', name: 'Second' }]
describe('controlled S6.1 navigator/status shell', () => {
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
  it('clips selected viewport records and keeps the blue text selection separate', () => {
    const selection = new SourceSelection([{ rows: { kind: 'intervals', intervals: [{ r0: 3, r1: 3 }] }, cols: [{ c0: 0, c1: 0 }] }])
    const rects = gridSelectionOverlay(selection, [{ source: { row: 3, col: 0 }, bounds: { x: -5, y: 2, width: 20, height: 24 } }, { source: { row: 1, col: 0 }, bounds: { x: 15, y: 2, width: 20, height: 24 } }], { x: 0, y: 0, width: 10, height: 10 })
    expect(rects).toEqual([{ x: 0, y: 2, width: 10, height: 8 }])
    const ctx = { save: vi.fn(), restore: vi.fn(), fillRect: vi.fn(), strokeRect: vi.fn(), fillStyle: '', strokeStyle: '', lineWidth: 0 }
    paintGridSelectionOverlay(ctx as unknown as CanvasRenderingContext2D, rects, 2)
    expect(ctx.strokeStyle).toBe('#008000')
    expect(ctx.lineWidth).toBe(1)
    expect(ctx.strokeRect).toHaveBeenCalledWith(0, 2, 10, 8)
    expect(ctx.restore).toHaveBeenCalledOnce()
  })
})
