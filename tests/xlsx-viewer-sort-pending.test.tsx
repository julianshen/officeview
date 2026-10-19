/**
 * S10.2 input acknowledgment (spec §12.2, ≤ 50 ms): a header sort renders a
 * pending state synchronously with the click — before range admission or the
 * worker sort runs — and clears it on completion, cancellation, failure, sheet
 * switch and document replacement. Reproduced in Chrome: on a 1M-row paged
 * sheet nothing changed for 12–13 s until the result arrived.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { GRID_HEADER_SIZE } from '../src/components/xlsx/GridHeaders'
import * as controlsModule from '../src/components/xlsx/viewerControlsClient'
import { workbookOf } from './helpers/xlsx-find'
import { domClassBytes, domClassCounts } from '../src/components/xlsx/domCost'
import type { MainMemoryLedger } from '../src/xlsx/view/mainLedger'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

const sheet = (sheetId: string) => ({ sheetId, cells: [{ row: 0, col: 0, value: 'Name' }, { row: 0, col: 1, value: 'Score' }, { row: 1, col: 0, value: 'beta' }, { row: 1, col: 1, value: 2 }, { row: 2, col: 0, value: 'alpha' }, { row: 2, col: 1, value: 1 }] })

/** Holds every range resolution until released, like a slow paged admission. */
function gateResolve(mode: 'hold' | 'fail' = 'hold') {
  let release: () => void = () => {}
  const gate = new Promise<void>(resolve => { release = resolve })
  const original = controlsModule.viewerControls
  vi.spyOn(controlsModule, 'viewerControls').mockImplementation(async (...args) => {
    const port = await original(...args)
    const resolve = port.client.resolve.bind(port.client)
    port.client.resolve = async (...call: Parameters<typeof resolve>) => {
      await gate
      if (mode === 'fail') throw new Error('Range admission failed (test).')
      return resolve(...call)
    }
    return port
  })
  return () => release()
}
const statusText = (ui: ReturnType<typeof render>) => ui.getAllByRole('status').map(s => s.textContent).join(' ')
async function ready(ui: ReturnType<typeof render>) {
  await waitFor(() => expect(statusText(ui)).not.toMatch(/Loading sheet/))
  await waitFor(() => expect(ui.getByRole('button', { name: 'Cycle sort column B' }).hasAttribute('disabled')).toBe(false))
}
const pending = (ui: ReturnType<typeof render>) => ({
  status: /Sorting column B/.test(statusText(ui)),
  busy: ui.getByRole('grid').getAttribute('aria-busy') === 'true',
  header: ui.getByRole('button', { name: 'Cycle sort column B' }).getAttribute('aria-busy') === 'true',
})

describe('header sort pending acknowledgment', () => {
  it('is visible in the same commit as the click, before admission finishes, and clears on completion', async () => {
    const release = gateResolve()
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} />)
    await ready(ui)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' })) })
    // No awaited work has run yet: the acknowledgment cannot depend on it.
    expect(pending(ui)).toEqual({ status: true, busy: true, header: true })
    await act(async () => { release() })
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
    expect(pending(ui)).toEqual({ status: false, busy: false, header: false })
    expect(statusText(ui)).toMatch(/View-only sort applied/)
  })

  it('acknowledges a pending narrow-column sort in the 44px coarse-pointer corner lane and restores committed feedback', async () => {
    const previous = window.matchMedia
    window.matchMedia = vi.fn(query => ({ matches: query.includes('pointer: coarse'), media: query, onchange: null,
      addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true }))
    try {
      const release = gateResolve()
      const doc = workbookOf([sheet('A')])
      doc.sheets[0].cols = [{ min: 1, max: 2, widthChars: 2 }]
      const ui = render(<InteractiveSheetDoc doc={doc} />)
      await waitFor(() => expect(statusText(ui)).not.toMatch(/Loading sheet/))
      fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'B1' } })
      fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
      const grid = ui.getByRole('grid')
      act(() => grid.focus())
      const disclosure = ui.getByRole('button', { name: 'Actions for column B' })
      fireEvent.click(disclosure)
      const sort = ui.getByRole('button', { name: 'Cycle sort column B' })
      expect(sort.tabIndex).toBe(-1)
      expect(fireEvent.mouseDown(sort)).toBe(false)
      expect(document.activeElement).toBe(grid)
      act(() => { fireEvent.click(sort) })
      // The action closes the lane; reopening must show the existing pending job.
      fireEvent.click(disclosure)
      const pendingSort = ui.getByRole('button', { name: 'Cycle sort column B' })
      expect(pending(ui)).toEqual({ status: true, busy: true, header: true })
      expect(pendingSort.getAttribute('aria-description')).toBe('sorting')
      expect(pendingSort.title).toBe('Column B; sorting')
      expect(pendingSort.querySelector('svg[data-icon="loader"]')).toBeTruthy()
      expect(pendingSort.getAttribute('aria-pressed')).toBe('false')
      expect(pendingSort.getAttribute('data-action-column')).toBe('1')
      expect(document.activeElement).toBe(grid)
      const group = ui.getByRole('group', { name: 'Column B actions' })
      const targets = [...group.querySelectorAll('button')]
      const boxes = targets.map(button => ({
        left: parseFloat(group.style.left) + parseFloat(button.style.left), top: parseFloat(button.style.top),
        width: Math.max(parseFloat(button.style.width), parseFloat(button.style.minWidth)),
        height: Math.max(parseFloat(button.style.height), parseFloat(button.style.minHeight)),
      }))
      for (const box of boxes) {
        expect(box.width).toBeGreaterThanOrEqual(44)
        expect(box.height).toBeGreaterThanOrEqual(44)
        expect(box.left + box.width).toBeLessThanOrEqual(GRID_HEADER_SIZE)
        for (const header of ui.getAllByRole('columnheader').filter(h => h.hasAttribute('aria-sort') || /^Column /.test(h.getAttribute('aria-label') ?? ''))) {
          const left = parseFloat(header.style.left) + parseFloat(/translateX\(([-\d.]+)px\)/.exec(header.style.transform)![1])
          expect(box.left + box.width).toBeLessThanOrEqual(left)
        }
      }
      const header = ui.getByRole('columnheader', { name: 'Column B' })
      expect(parseFloat(header.style.width)).toBeLessThan(44)
      expect(header.querySelectorAll('[data-action-column]')).toHaveLength(0)
      for (let i = 1; i < boxes.length; i++) expect(boxes[i].top).toBeGreaterThanOrEqual(boxes[i - 1].top + boxes[i - 1].height)
      await act(async () => { release() })
      await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
      const committed = ui.getByRole('button', { name: 'Cycle sort column B' })
      expect(pending(ui)).toEqual({ status: false, busy: false, header: false })
      expect(committed.getAttribute('aria-description')).toBe('ascending')
      expect(committed.title).toBe('Column B; sort priority 1, ascending; ascending')
      expect(committed.querySelector('svg[data-icon="arrow-up"]')).toBeTruthy()
      expect(committed.getAttribute('aria-pressed')).toBe('true')
      expect(committed.tabIndex).toBe(-1)
      expect(document.activeElement).toBe(grid)
      expect(statusText(ui)).toMatch(/View-only sort applied/)
    } finally { window.matchMedia = previous }
  })

  it('holds the DOM charge while pending: the smaller pending glyph frees no headroom that the restored glyph would overrun (§12.4)', async () => {
    // The pending spinner has fewer elements than the resting sort glyph. If
    // the charge followed that dip, admitted work (a ring-tile prefetch) could
    // take the freed bytes, and the unconditional re-measure when the glyph
    // returns would push the ledger over the hard budget.
    const release = gateResolve()
    let ledger: MainMemoryLedger | null = null
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} {...({ onMainLedger: (l: MainMemoryLedger) => { ledger = l } } as object)} />)
    await ready(ui)
    const sort = () => ui.getByRole('button', { name: 'Cycle sort column B' })
    const before = ledger!.ownerBytes('dom', 'grid')
    const whilePending: number[] = []
    const set = ledger!.set.bind(ledger!)
    ledger!.set = ((...args: Parameters<MainMemoryLedger['set']>) => {
      set(...args)
      if (args[0] === 'dom' && args[1] === 'grid' && ui.container.querySelector('[aria-label="Cycle sort column B"][aria-busy="true"]')) whilePending.push(args[2])
    }) as MainMemoryLedger['set']
    act(() => { fireEvent.click(sort()) })
    expect(pending(ui).header).toBe(true)
    await act(async () => { for (let i = 0; i < 4; i++) await new Promise(resolve => setTimeout(resolve, 0)) })
    const root = ui.container.querySelector('[data-officeview-root]')!
    expect(domClassBytes(domClassCounts(root)), 'the pending glyph is smaller than the resting one').toBeLessThan(before)
    expect(whilePending.length, 'the pending commit was re-measured').toBeGreaterThan(0)
    expect(Math.min(...whilePending)).toBeGreaterThanOrEqual(before)
    await act(async () => { release() })
    await waitFor(() => expect(ui.getByRole('columnheader', { name: /Column B; sort priority 1, ascending/ })).toBeTruthy())
    // The hold ends with the job: the charge is the measured DOM again.
    await waitFor(() => expect(ledger!.ownerBytes('dom', 'grid')).toBe(domClassBytes(domClassCounts(root))))
  })

  it.each([['refusal', 'fail'], ['completion', 'hold']] as const)('keeps the DOM hold until the commit that restores the resting glyph (%s)', async (_, mode) => {
    // Between the job's end and React committing the restored glyph, the DOM
    // still shows the smaller spinner. A re-measure in that window must not
    // publish the dip, or a charge admitted into it is overrun by the commit.
    const release = gateResolve(mode)
    let ledger: MainMemoryLedger | null = null
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} {...({ onMainLedger: (l: MainMemoryLedger) => { ledger = l } } as object)} />)
    await ready(ui)
    const before = ledger!.ownerBytes('dom', 'grid')
    fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' }))
    const root = ui.container.querySelector('[data-officeview-root]')!
    const budget = 64 * 1024 * 1024
    await act(async () => {
      // Inside act, React commits only when the scope ends: the job settles
      // (its finally runs) while the committed DOM still shows the spinner.
      release()
      for (let i = 0; i < 20; i++) await new Promise(resolve => setTimeout(resolve, 0))
      expect(statusText(ui), 'the restoring commit has not happened yet').toMatch(/Sorting column B/)
      expect(pending(ui).header).toBe(true)
      // A count-neutral mutation runs the real subtree observer's re-measure.
      const marker = document.createComment('observer tick'); root.append(marker); marker.remove()
      for (let i = 0; i < 5; i++) await Promise.resolve()
      expect(ledger!.ownerBytes('dom', 'grid')).toBeGreaterThanOrEqual(before)
      // Another consumer is admitted up to exactly the budget.
      ledger!.set('controls', 'test-pressure', budget - ledger!.bytes)
    })
    expect(pending(ui).header).toBe(false)
    expect(ledger!.bytes).toBeLessThanOrEqual(budget)
    // The hold ended with that commit: the charge is the measured DOM again.
    expect(ledger!.ownerBytes('dom', 'grid')).toBe(domClassBytes(domClassCounts(root)))
    ledger!.release('controls', 'test-pressure')
  })

  it('clears on failure with the failure announced', async () => {
    const release = gateResolve('fail')
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} />)
    await ready(ui)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' })) })
    expect(pending(ui).busy).toBe(true)
    await act(async () => { release() })
    await waitFor(() => expect(statusText(ui)).toMatch(/Range admission failed/))
    expect(pending(ui)).toEqual({ status: false, busy: false, header: false })
  })

  it('a sheet switch cancels the job and clears the pending state; the stale job publishes nothing', async () => {
    const release = gateResolve()
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A'), sheet('B')])} />)
    await ready(ui)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' })) })
    expect(pending(ui).busy).toBe(true)
    act(() => { fireEvent.change(ui.getByRole('combobox', { name: 'Worksheet' }), { target: { value: 'B' } }) })
    expect(ui.getByRole('grid').getAttribute('aria-busy')).not.toBe('true')
    expect(statusText(ui)).not.toMatch(/Sorting column/)
    await act(async () => { release() })
    await new Promise(resolve => setTimeout(resolve, 20))
    expect(statusText(ui)).not.toMatch(/Sorting column|View-only sort applied/)
  })

  it('a replaced document never inherits the pending state of the old one', async () => {
    const release = gateResolve()
    const ui = render(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} />)
    await ready(ui)
    act(() => { fireEvent.click(ui.getByRole('button', { name: 'Cycle sort column B' })) })
    expect(pending(ui).busy).toBe(true)
    ui.rerender(<InteractiveSheetDoc doc={workbookOf([sheet('A')])} />)
    expect(ui.getByRole('grid').getAttribute('aria-busy')).not.toBe('true')
    expect(statusText(ui)).not.toMatch(/Sorting column/)
    await act(async () => { release() })
    await new Promise(resolve => setTimeout(resolve, 20))
    expect(statusText(ui)).not.toMatch(/Sorting column|View-only sort applied/)
  })
})
