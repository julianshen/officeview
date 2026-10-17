/**
 * The heavy repaint-admission cases (split from xlsx-viewer-resize-admission):
 * each needs a wrapped record the layout cache cannot retain, so its size is
 * set by DEFAULT_LAYOUT_CACHE_BYTES (about 7,700 lines, derived below) and it
 * is laid out on every paint. They live in their own file so a slow runner
 * timing one of them out cannot starve the lighter suite's first paints.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor, type RenderResult } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'
import { DEFAULT_LAYOUT_CACHE_BYTES, layoutUpperBytes } from '../src/xlsx/view/textLayout'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'
import { BUDGET, cell, docOf, frames, goTo, linesBuiltOver, linesOver, mount, resizeKey, settle, squeeze, status } from './helpers/resize-admission'

afterEach(() => { cleanup(); vi.restoreAllMocks() })
const wrapped = (row: number, col: number, lines: number) => ({ ...cell(row, col, 'a\n'.repeat(lines)), style: { numFmtId: 0, wrapText: true } })
const layoutMemoryNotice = (ui: RenderResult) => /Not enough memory to lay out this text/.test(status(ui)) || ui.container.querySelectorAll('canvas').length > 0

describe('confirmation r5: repaint records and reservation transfer (heavy)', () => {
  it('F3 a single layout record larger than the cache budget is admitted at its own construction size (built record 10% above the free room, which exceeds the budget)', FIRST_PAINT_TEST, async () => {
    // The free room exceeds the cache budget, and the built record exceeds the
    // free room by 10% (more than the repaint's own releases can absorb):
    // admitting only the budget would let its construction pass the limit.
    const free = DEFAULT_LAYOUT_CACHE_BYTES + 140_000
    const lines = linesBuiltOver(1.1 * free)
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, 'a\n'.repeat(lines)), style: { numFmtId: 0, wrapText: true } }] }])
    const { ui, ledger } = await mount((onLedger, onFrame) => <InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={onLedger} onFrame={onFrame} />)
    const grid = goTo(ui, 'A1'); await settle()
    const check = squeeze(ledger(), free)
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    check()
    ledger().set('copy', 'filler', 0)
  })
})

describe('confirmation r6: repaint memory is admitted per record at construction (heavy)', () => {
  it('R1 three adjacent wrapped cells each bounded just above the cache budget (1.75 records free): the tile batch never holds more than admitted', { timeout: 120_000 }, async () => {
    // Each built record is too large for the cache (held by the batch,
    // uncached); one construction bound fits the free room, two do not.
    const lines = linesBuiltOver(DEFAULT_LAYOUT_CACHE_BYTES), free = Math.floor(1.75 * layoutUpperBytes(2 * lines, 0))
    const doc = docOf([{ index: 0, cells: [wrapped(0, 0, lines), wrapped(0, 1, lines), wrapped(0, 2, lines)] }])
    let peakOk = true
    const f = frames()
    let live: MainMemoryLedger | undefined
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value; value.set('copy', 'filler', Math.max(0, BUDGET - value.bytes - free)) }} onFrame={stats => { f.onFrame(stats); if (live && live.snapshot().peakBytes > BUDGET) peakOk = false }} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const grid = goTo(view, 'A1'); await settle()
    resizeKey(grid, 'ArrowRight'); await settle(); await settle()
    expect(peakOk).toBe(true)
    expect(live!.snapshot().peakBytes).toBeLessThanOrEqual(BUDGET)
    expect(live!.bytes).toBeLessThanOrEqual(BUDGET)
    expect(layoutMemoryNotice(view)).toBe(true)
  })
})

describe('find highlights are admitted like tiles', () => {
  it('a find hit on a heavy cell at low free memory stays within the hard limit (whole-cell highlight fallback)', FIRST_PAINT_TEST, async () => {
    // The hit's record exceeds the cache budget, so it is never a warm cache
    // hit (every highlight builds it), and the low free room: refused there,
    // built with room.
    const free = 2 * 1024 * 1024
    const doc = docOf([{ index: 0, cells: [{ ...cell(0, 0, `needle ${'a\n'.repeat(Math.max(linesOver(free), linesBuiltOver(DEFAULT_LAYOUT_CACHE_BYTES)))}`), style: { numFmtId: 0, wrapText: true } }] }, { index: 1, cells: [cell(1, 0, 'b')] }])
    let live: MainMemoryLedger | undefined
    const f = frames()
    const view = render(<InteractiveSheetDoc doc={doc} views={{ s: { revision: 0 } }} onMainLedger={value => { live = value }} onFrame={f.onFrame} />)
    await waitFor(() => expect(f.last()?.basePaints ?? 0).toBeGreaterThan(0), FIRST_PAINT_WAIT)
    await settle()
    const check = squeeze(live!, free)
    const paintRecords: number[] = []
    const set = live!.set.bind(live!)
    vi.spyOn(live!, 'set').mockImplementation((pool, owner, bytes) => { set(pool, owner, bytes); if (owner === 'paint-records') paintRecords.push(bytes) })
    fireEvent.change(view.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'needle' } })
    await waitFor(() => expect(status(view)).toMatch(/A1/))
    await settle(); await settle()
    check()
    // Not admitted at 2 MiB: no record was built for the overlay (whole-cell highlight).
    expect(paintRecords).toEqual([])
    // With room, the same highlight is built through the admitted batch.
    live!.set('copy', 'filler', Math.max(0, BUDGET - live!.bytes + live!.ownerBytes('copy', 'filler') - 40 * 1024 * 1024))
    fireEvent.change(view.getByRole('searchbox', { name: 'Search document' }), { target: { value: 'needl' } })
    await waitFor(() => expect(paintRecords.length).toBeGreaterThan(0))
    expect(live!.snapshot().peakBytes).toBeLessThanOrEqual(BUDGET)
    live!.set('copy', 'filler', 0)
  })
})
