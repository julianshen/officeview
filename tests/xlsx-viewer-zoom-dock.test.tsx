/**
 * Stream F item 3 (iOS Simulator run): on iPhone the floating zoom cluster
 * (+ / − / 100%, bottom-right over the grid) covered cell content in columns
 * E-F. On a coarse pointer or a narrow pane the zoom controls dock into the
 * sheet toolbar instead, so they never overlap the grid; desktop keeps the
 * floating cluster.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { XlsxDocument } from '../src/xlsx/types'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

function workbook(): XlsxDocument {
  const rows = Array.from({ length: 30 }, (_, r) => ({ index: r, cells: Array.from({ length: 6 }, (_, c) => ({ col: c, value: r * 10 + c })) }))
  return { sheets: [{ name: 'S', rows, cols: [], merges: [], mergeRanges: [] }] } as unknown as XlsxDocument
}
function coarse(matches: boolean) {
  window.matchMedia = vi.fn(query => ({ matches: matches && query.includes('pointer: coarse'), media: query, onchange: null,
    addEventListener: vi.fn(), removeEventListener: vi.fn(), addListener: vi.fn(), removeListener: vi.fn(), dispatchEvent: () => true })) as unknown as typeof window.matchMedia
}

describe('zoom controls placement', () => {
  const previous = window.matchMedia
  afterEach(() => { cleanup(); window.matchMedia = previous })
  it('a coarse pointer docks them in the sheet toolbar, outside the grid pane, with 44px targets', async () => {
    coarse(true)
    const ui = render(<InteractiveSheetDoc doc={workbook()} />)
    await waitFor(() => ui.getByRole('grid'), FIRST_PAINT_WAIT)
    const zoom = ui.getByTestId('officeview-zoom-controls')
    expect(zoom.closest('[data-officeview-ui="sheet-toolbar"]')).not.toBeNull()
    expect(ui.getByTestId('officeview-sheet-pane').contains(zoom)).toBe(false)
    expect(zoom.style.position).not.toBe('absolute')
    for (const name of ['Zoom in', 'Zoom out', 'Reset zoom']) {
      const button = ui.getByRole('button', { name })
      expect(zoom.contains(button)).toBe(true)
      expect(parseFloat(button.style.minWidth)).toBeGreaterThanOrEqual(44)
      expect(parseFloat(button.style.minHeight)).toBeGreaterThanOrEqual(44)
    }
    expect(ui.getByRole('button', { name: 'Reset zoom' }).textContent).toBe('100%')
    fireEvent.click(ui.getByRole('button', { name: 'Zoom in' }))
    await waitFor(() => expect(ui.getByRole('button', { name: 'Reset zoom' }).textContent).not.toBe('100%'))
  }, FIRST_PAINT_TEST.timeout)
  it('a desktop pointer keeps the floating cluster', async () => {
    coarse(false)
    const ui = render(<InteractiveSheetDoc doc={workbook()} />)
    await waitFor(() => ui.getByRole('grid'), FIRST_PAINT_WAIT)
    const zoom = ui.getByTestId('officeview-zoom-controls')
    expect(zoom.closest('[data-officeview-ui="sheet-toolbar"]')).toBeNull()
    expect(zoom.style.position).toBe('absolute')
  }, FIRST_PAINT_TEST.timeout)
})
