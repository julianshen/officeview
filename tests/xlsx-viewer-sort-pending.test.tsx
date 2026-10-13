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
import * as controlsModule from '../src/components/xlsx/viewerControlsClient'
import { workbookOf } from './helpers/xlsx-find'

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
