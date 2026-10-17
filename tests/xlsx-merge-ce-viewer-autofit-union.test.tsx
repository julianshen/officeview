import { afterEach, expect, it } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { workbookOf } from './helpers/xlsx-find'
import { FIRST_PAINT_TEST, FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(cleanup)

it('actual viewer auto-fit samples pinned+body union before the cap and excludes its invisible enclosing hull', FIRST_PAINT_TEST, async () => {
  const title = 'W'.repeat(40), gap = 'G'.repeat(200)
  const doc = workbookOf([{ sheetId: 's', cells: Array.from({ length: 7001 }, (_, row) => ({ row, col: 0, value: row === 0 ? title : row === 500 ? gap : 'body' })) }])
  // Actual current pane sees over2048 tiny body rows. Row0 is pinned; row500
  // lies in the invisible gap, so a hull-based sampler would choose its huge text.
  for (const row of doc.sheets[0].rows) row.heightPt = row.index === 0 ? 15 : 0.0375
  doc.sheets[0].frozen = { rows: 1, cols: 0, topRow: 1000 }
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  const grid = ui.getByRole('grid')
  fireEvent.change(ui.getByRole('textbox', { name: 'Go To cell' }), { target: { value: 'A1' } })
  fireEvent.click(ui.getByRole('button', { name: 'Go To' }))
  await waitFor(() => expect(grid.getAttribute('aria-activedescendant')).toBeTruthy(), FIRST_PAINT_WAIT)
  act(() => grid.focus())
  fireEvent.keyDown(grid, { key: 'Enter', altKey: true, shiftKey: true })
  await waitFor(() => expect(ui.container.textContent).toContain('fitted to content'), FIRST_PAINT_WAIT)
  const fitStatus = ui.getByRole('status').textContent ?? ''
  const matched = /Column A width (\d+) pixels, fitted to content/.exec(fitStatus)
  expect(matched).not.toBeNull()
  const width = Number(matched![1])
  // Forty Ws must win over "body"; two hundred Gs from the hidden hull must not.
  expect(width).toBeGreaterThan(200)
  expect(width).toBeLessThan(800)
})
