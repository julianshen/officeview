/**
 * View-only resize admission (spec §5.6, §12.4), codex confirmation r3: a
 * resize builds its candidate geometry off to the side, measures it with the
 * entry's own accounting, admits that against the live ledger while the old
 * geometry is still held, and publishes only if it fits. The peak never
 * passes the hard limit. Also: a structure-only worker layout publication
 * refreshes authored-hidden rows even when the view revision is unchanged.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import type { XlsxDocument, XlsxRow } from '../src/xlsx/types'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import * as viewportModule from '../src/xlsx/view/viewport'
import { createRowOrder } from '../src/xlsx/view/order'
import { cell, docOf, frames, goTo, headerWidth, linesOver, mount, remoteOf, resizeKey, rowHeight, settle, shell, squeeze, status } from './helpers/resize-admission'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

describe('confirmation r3 C1: a resize admits its measured candidate before publishing it', () => {
  it('default resident view (no views prop), one cell, 100 KiB free: the view mask is reused, never allocated twice', FIRST_PAINT_TEST, async () => {
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf([{ index: 0, cells: [cell(0, 0, 'a')] }])} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const viewBuffers = ledger().snapshot().byPool.viewBuffers
    const check = squeeze(ledger(), 100 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle()
    check()
    expect(ledger().snapshot().byPool.viewBuffers).toBe(viewBuffers)
    expect(headerWidth(ui, 'A') === 72 || /Column A was not resized: .*memory/.test(status(ui))).toBe(true)
    ledger().set('copy', 'filler', 0)
  })

  it('worker-backed sheet with 2,999 authored row heights, 400 KiB free: the real rebuilt row axis is admitted', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 2999 }, (_, index) => ({ index, customHeight: true, heightPt: 15 + (index % 7), cells: [cell(index, 0, index)] }))
    const doc = docOf(rows)
    const remote = await remoteOf(doc)
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 400 * 1024)
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })

  it('auto-fitting a 400 px first row of a 100-row sheet (200 KiB free) admits the semantic DOM the fit reveals', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 100 }, (_, index) => ({ index, ...(index === 0 ? { customHeight: true, heightPt: 300 } : {}), cells: Array.from({ length: 12 }, (_, col) => cell(index, col, `${index}:${col}`)) }))
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    fireEvent.keyDown(grid, { key: 'F2' })
    const rowSelect = ui.getByRole('button', { name: 'Select row 1' })
    act(() => rowSelect.focus())
    const check = squeeze(ledger(), 200 * 1024)
    resizeKey(rowSelect, 'Enter'); await settle()
    check()
    expect(rowHeight(ui, 'Row 1') < 400 || /Row 1 was not (resized|fitted): .*memory/.test(status(ui))).toBe(true)
    ledger().set('copy', 'filler', 0)
  })

  it('an unrepresentable size (3,000 rows, first two 1e-15 px, 600 KiB free) is refused without publishing a fallback geometry', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 3000 }, (_, index) => ({ index, ...(index < 2 ? { customHeight: true, heightPt: 1e-15 * 72 / 96 } : {}), cells: [cell(index, 0, index)] }))
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const geometryBefore = ledger().snapshot().byPool.geometry
    const check = squeeze(ledger(), 600 * 1024)
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    expect(status(ui)).toMatch(/Row 1 was not resized/)
    expect(ledger().snapshot().byPool.geometry).toBe(geometryBefore)
    expect((ui.getByRole('button', { name: 'Reset sizes' }) as HTMLButtonElement).disabled).toBe(true)
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r4: admission happens before construction', () => {
  it('R1 a candidate whose construction cannot fit is refused before the geometry is built (100k sorted rows, 512 KiB free)', { timeout: 60_000 }, async () => {
    const rows = 100_000
    const doc = docOf(Array.from({ length: rows }, (_, index) => ({ index, cells: [cell(index, 0, index)] })))
    const order = createRowOrder([{ r0: 0, r1: rows - 1, sourceAtOffset: Uint32Array.from({ length: rows }, (_, i) => rows - 1 - i) }])
    const views = { s: { revision: 1, order } }
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={views} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 512 * 1024)
    const builds = vi.spyOn(viewportModule, 'buildSheetGeometry')
    resizeKey(grid, 'ArrowDown'); await settle()
    check()
    expect(status(ui)).toMatch(/Row 1 was not resized: .*memory/)
    expect(builds).not.toHaveBeenCalled()
    ledger().set('copy', 'filler', 0)
  })

  it('R3 repaint layout growth is admitted even when the visible counts do not change (2,048-character wrapped cell, 64 KiB free)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'word '.repeat(410).slice(0, 2048)), style: { numFmtId: 0, wrapText: true } }] }])
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), 64 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r5: repaint records and reservation transfer', () => {
  it('F2 a visible merge whose anchor is above the pane is admitted (A1:A200, 1,024-line wrapped anchor, scrolled to row 101, 200 KiB free)', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 200 }, (_, index) => ({ index, cells: index === 0 ? [{ ...cell(0, 0, 'a\n'.repeat(1024)), style: { numFmtId: 0, wrapText: true } }] : [] }))
    const doc: XlsxDocument = { sheets: [{ ...docOf(rows).sheets[0], merges: ['A1:A200'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: 199, maxCol: 0 }] }] }
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    goTo(ui, 'A1'); await settle()
    const scroller = ui.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 100 * 20; fireEvent.scroll(scroller) }); await settle()
    const grid = ui.getByRole('grid'); act(() => grid.focus())
    const check = squeeze(ledger(), 200 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })

  it('F4 the reservation transfer never holds the construction and repaint reservations together', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'z'.repeat(128))] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    const remote = await remoteOf(doc)
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    await act(async () => { await remote.filter(remote.identity('s'), null, [], new AbortController().signal) })
    await settle()
    const grid = goTo(ui, 'A1'); await settle()
    let overlapped = false
    const set = ledger().set.bind(ledger())
    vi.spyOn(ledger(), 'set').mockImplementation((pool, owner, bytes) => {
      set(pool, owner, bytes)
      if (ledger().ownerBytes('geometry', 'resize-candidate') > 0 && ledger().ownerBytes('dom', 'resize-reservation') > 0) overlapped = true
    })
    const check = squeeze(ledger(), 400_000)
    resizeKey(grid, 'ArrowRight'); await settle()
    check()
    expect(headerWidth(ui, 'A')).toBe(72)
    expect(overlapped).toBe(false)
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r6: repaint memory is admitted per record at construction', () => {
  const heavy = (row: number, col: number, lines: number) => ({ ...cell(row, col, 'a\n'.repeat(lines)), style: { numFmtId: 0, wrapText: true } })

  it('R2 detail tiles under a coarse aggregate are bounded too (A2 bounded at twice the free room among 400 dense 8 px rows, 1 MiB free, column A 8 → 16)', FIRST_PAINT_TEST, async () => {
    // A2's record bound is twice the free room.
    const free = 1024 * 1024
    const rows: XlsxRow[] = Array.from({ length: 400 }, (_, index) => ({ index, customHeight: true, heightPt: 6, cells: Array.from({ length: 64 }, (_, col) => index === 1 && col === 0 ? heavy(1, 0, linesOver(2 * free)) : cell(index, col, col)) }))
    const doc: XlsxDocument = { sheets: [{ ...docOf(rows).sheets[0], cols: [{ min: 0, max: 63, widthChars: 8 / 7 }] }] }
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A1'); await settle()
    const check = squeeze(live!, free)
    for (let i = 0; i < 1; i++) resizeKey(grid, 'ArrowRight')
    await settle(); await settle()
    check()
    live!.set('copy', 'filler', 0)
  })

  it('R3 formatting during painting is admitted through the allocation observer (2,048-unit text, 200-token @ format, 256 KiB free)', FIRST_PAINT_TEST, async () => {
    const style = { numFmtId: 164, formatCode: `0;0;0;${'@'.repeat(200)}` }
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'q'.repeat(2048)), style }] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A2'); await settle()
    const check = squeeze(live!, 256 * 1024)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    live!.set('copy', 'filler', 0)
  })

  it('scrolling into a heavy region without any resize stays within the hard limit and degrades to the coarse notice', FIRST_PAINT_TEST, async () => {
    // Each record's bound is twice the free room, so painting them is refused.
    const free = 4 * 1024 * 1024
    const rows: XlsxRow[] = [{ index: 0, cells: [cell(0, 0, 'top')] }, { index: 300, cells: [heavy(300, 0, linesOver(2 * free)), heavy(300, 1, linesOver(2 * free))] }]
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const check = squeeze(live!, free)
    const scroller = view.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 300 * 20; fireEvent.scroll(scroller) })
    await settle(); await settle()
    check()
    expect(status(view)).toMatch(/Not enough memory to lay out this text/)
    live!.set('copy', 'filler', 0)
  })
})

describe('confirmation r7 R2: tile context and text shaping are admitted before drawing', () => {
  it('scrolling ordinary text with 64 KiB free never passes the hard limit (shaping refused: coarse tile)', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 400 }, (_, index) => ({ index, cells: Array.from({ length: 12 }, (_, col) => cell(index, col, `t${index}-${col}`)) }))
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const check = squeeze(live!, 64 * 1024)
    const scroller = view.getByTestId('officeview-sheet-scroller')
    for (const top of [2000, 4000, 6000]) { act(() => { scroller.scrollTop = top; fireEvent.scroll(scroller) }); await settle() }
    await settle()
    check()
    live!.set('copy', 'filler', 0)
  })
})

describe('shared admitted measurement context', () => {
  it('wrapped text measured by the canonical layout is charged to one shared context, and tight memory never passes the limit', FIRST_PAINT_TEST, async () => {
    const words = (row: number, col: number) => Array.from({ length: 6 }, (_, i) => `r${row}c${col}w${i}`).join(' ')
    const rows: XlsxRow[] = Array.from({ length: 30 }, (_, index) => ({ index, customHeight: true, heightPt: 60, cells: Array.from({ length: 10 }, (_, col) => ({ ...cell(index, col, words(index, col)), style: { numFmtId: 0, wrapText: true } })) }))
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    expect(live!.ownerBytes('canvasText', 'measure')).toBeGreaterThan(0)
    const check = squeeze(live!, 96 * 1024)
    const scroller = view.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollLeft = 640; fireEvent.scroll(scroller) }); await settle()
    act(() => { scroller.scrollTop = 1600; fireEvent.scroll(scroller) }); await settle()
    check()
    live!.set('copy', 'filler', 0)
  })
})

describe('confirmation r8 F2: the resource notice label is admitted too', () => {
  it('with too little room for the label shaping, a refused tile paints the tone only (no unadmitted fillText)', FIRST_PAINT_TEST, async () => {
    const rows: XlsxRow[] = Array.from({ length: 200 }, (_, index) => ({ index, cells: Array.from({ length: 10 }, (_, col) => cell(index, col, `v${index}-${col}`)) }))
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={docOf(rows)} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const proto = Object.getPrototypeOf(document.createElement('canvas').getContext('2d')!) as CanvasRenderingContext2D
    const drawn: string[] = []
    const fillText = proto.fillText
    vi.spyOn(proto, 'fillText').mockImplementation(function (this: CanvasRenderingContext2D, text: string, x: number, y: number) { drawn.push(String(text)); return fillText.call(this, text, x, y) })
    const check = squeeze(live!, 8 * 1024)
    const scroller = view.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 2000; fireEvent.scroll(scroller) }); await settle()
    check()
    expect(drawn.filter(text => /Not enough memory to lay out this text/.test(text))).toEqual([])
    live!.set('copy', 'filler', 0)
  })
})

describe('confirmation r3 C2: structure-only worker layout publications refresh authored-hidden rows', () => {
  it('publishLayout hiding row 1 hides it in the viewer even though the view revision is unchanged', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'a')] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    const remote = await remoteOf(doc)
    const { ui } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    expect(rowHeight(ui, 'Row 1')).toBe(20)
    act(() => remote.publishLayout('s', { ...remote.layout('s'), hidden: [{ r0: 0, r1: 0 }] }))
    await settle()
    await waitFor(() => expect(ui.queryByRole('rowheader', { name: 'Row 1' })).toBeNull())
  })

  it('R4 after a committed worker filter view, a layout publication hiding row 1 still hides it (mask AND NOT authored-hidden)', FIRST_PAINT_TEST, async () => {
    const doc = docOf([{ index: 0, cells: [cell(0, 0, 'a')] }, { index: 1, cells: [cell(1, 0, 'b')] }, { index: 2, cells: [cell(2, 0, 'c')] }])
    const remote = await remoteOf(doc)
    const { ui } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={shell(doc)} source={remote as never} onMainLedger={onLedger} onFrame={onFrame} />)
    await act(async () => { await remote.filter(remote.identity('s'), null, [], new AbortController().signal) })
    await settle()
    expect(remote.view('s').visibility).toBeTruthy()
    expect(rowHeight(ui, 'Row 1')).toBe(20)
    act(() => remote.publishLayout('s', { ...remote.layout('s'), hidden: [{ r0: 0, r1: 0 }] }))
    await settle()
    await waitFor(() => expect(ui.queryByRole('rowheader', { name: 'Row 1' })).toBeNull())
    expect(rowHeight(ui, 'Row 2')).toBe(20)
  })
})
