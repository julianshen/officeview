/**
 * S10.3 browser-support regressions for the interactive viewer (spec §9.2,
 * §12.4; AC-A11Y-01/02, AC-BACKEND-01 fallback leg).
 *
 * - One grid Tab stop: Chromium makes an overflow:auto scroller keyboard-
 *   focusable when it has no tabindex. The pan scroller sits inside the grid,
 *   whose focus capture sends any inner focus back to the grid wrapper, so in
 *   Chromium Tab cycled grid → scroller → grid forever (a keyboard trap,
 *   reproduced in tests/browser/interactive-viewer-a11y.spec.ts). jsdom has no
 *   focusable-scroller rule, so this pins the structural fix: every scroll
 *   container inside the grid is out of the sequential tab order.
 * - Canvas context loss: a restored 2D context starts blank, and so may the
 *   cached tile canvases. The viewer must drop its tiles and repaint both
 *   layers on `contextrestored` instead of waiting for the next scroll.
 * - Font failure: a face that fails to register falls back to system faces;
 *   the viewer still paints and keeps its owned grid.
 */
import { describe, expect, it, vi } from 'vitest'
import { act, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc, { type SheetFrameStats } from '../src/components/xlsx/InteractiveSheetDoc'
import { prepareSheetAssets } from '../src/components/xlsx/sheetAssets'
import type { XlsxDocument } from '../src/xlsx/types'

function workbook(): XlsxDocument {
  const rows = Array.from({ length: 30 }, (_, r) => ({ index: r, cells: Array.from({ length: 6 }, (_, c) => ({ col: c, value: r * 10 + c })) }))
  return { sheets: [{ name: 'S', rows, cols: [], merges: [], mergeRanges: [] }] } as unknown as XlsxDocument
}
function frames() {
  const all: SheetFrameStats[] = []
  return { all, onFrame: (stats: SheetFrameStats) => { all.push(stats) }, last: () => all[all.length - 1] }
}

describe('one grid Tab stop', () => {
  it('no scroll container inside the grid is a sequential Tab stop', async () => {
    const ui = render(<InteractiveSheetDoc doc={workbook()} />)
    const grid = await waitFor(() => ui.getByRole('grid'))
    const scrollers = [...grid.querySelectorAll<HTMLElement>('*')].filter(el => /auto|scroll/.test(el.style.overflow + el.style.overflowY + el.style.overflowX))
    expect(scrollers.length).toBeGreaterThan(0)
    for (const el of scrollers) expect(el.getAttribute('tabindex')).toBe('-1')
  })
})

describe('canvas context loss', () => {
  it('contextrestored drops cached tiles and repaints the base and overlay layers', async () => {
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={workbook()} {...({ onFrame: f.onFrame } as object)} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0))
    await waitFor(() => expect(f.last()?.prefetched ?? 0).toBeGreaterThanOrEqual(0))
    const before = f.last()!
    const base = ui.getByTestId('officeview-sheet-base'), overlay = ui.getByTestId('officeview-sheet-overlay')
    act(() => { base.dispatchEvent(new Event('contextlost')); base.dispatchEvent(new Event('contextrestored')) })
    // Tiles were dropped: the frame is painted again from records, not blitted from a lost tile.
    await waitFor(() => expect(f.last()!.basePaints).toBeGreaterThan(before.basePaints))
    const afterBase = f.last()!
    act(() => { overlay.dispatchEvent(new Event('contextrestored')) })
    await waitFor(() => expect(f.last()!.overlayPaints).toBeGreaterThan(afterBase.overlayPaints))
  })
})

describe('font load failure', () => {
  it('a face that fails to register falls back; the viewer paints and keeps its owned grid', async () => {
    const doc = workbook()
    doc.embeddedFonts = [{ family: 'Brand Sans', variant: 'regular', partPath: 'xl/fonts/font1.odttf', relationshipId: 'rId1', bytes: new Uint8Array(16) }]
    const registerFont = vi.fn(async () => { throw new Error('FontFace load failed') })
    const leases: Array<Awaited<ReturnType<typeof prepareSheetAssets>>> = []
    const assetLoader: typeof prepareSheetAssets = (d, options) => prepareSheetAssets(d, { ...options, registerFont }).then(lease => { leases.push(lease); return lease })
    const f = frames()
    const ui = render(<InteractiveSheetDoc doc={doc} {...({ onFrame: f.onFrame, assetLoader } as object)} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0))
    expect(registerFont).toHaveBeenCalled()
    expect(leases[0].fontDiagnostics).toEqual([expect.objectContaining({ kind: 'font-load-failed', family: 'Brand Sans' })])
    expect(ui.getByRole('gridcell', { name: /^B2, value 11/ })).toBeTruthy()
    expect(ui.getByRole('status').textContent).not.toMatch(/Loading sheet/)
    // Retained, deduplicated font diagnostic on the status surface (§9.1).
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Font not available, shown with a fallback: Brand Sans/))
  })

  it('font notices are bounded: three families named, the rest counted, duplicates once', async () => {
    const doc = workbook()
    const face = (family: string, variant: 'regular' | 'bold') => ({ family, variant, partPath: `xl/fonts/${family}-${variant}.odttf`, relationshipId: 'rId1', bytes: new Uint8Array(16) })
    doc.embeddedFonts = [face('F1', 'regular'), face('F1', 'bold'), face('F2', 'regular'), face('F3', 'regular'), face('F4', 'regular'), face('F5', 'regular')]
    const assetLoader: typeof prepareSheetAssets = (d, options) => prepareSheetAssets(d, { ...options, registerFont: async () => { throw new Error('nope') } })
    const ui = render(<InteractiveSheetDoc doc={doc} {...({ assetLoader } as object)} />)
    await waitFor(() => expect(ui.getByRole('status').textContent).toMatch(/Fonts not available, shown with a fallback: F1, F2, F3 and 2 more/))
    expect(ui.getByRole('status').textContent!.match(/F1/g)).toHaveLength(1)
  })
})

describe('popup accessible names', () => {
  it('the filter popup is "Column filter" and the sort popup is "Sort worksheet"', async () => {
    const ui = render(<InteractiveSheetDoc doc={workbook()} />)
    fireEvent.click(await waitFor(() => ui.getByRole('button', { name: /^Filter column A/ })))
    expect((await waitFor(() => ui.getByRole('dialog'))).getAttribute('aria-label')).toBe('Column filter')
    fireEvent.keyDown(ui.getByRole('dialog'), { key: 'Escape' })
    await waitFor(() => expect(ui.queryByRole('dialog')).toBeNull())
    fireEvent.click(ui.getByRole('button', { name: 'More sort options for A' }))
    expect((await waitFor(() => ui.getByRole('dialog'))).getAttribute('aria-label')).toBe('Sort worksheet')
  })
})
