import { afterEach, expect, it } from 'vitest'
import { cleanup, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import { workbookOf } from './helpers/xlsx-find'

afterEach(cleanup)

it('checks resident sparkline presence only once per document across component rerenders', async () => {
  const doc = workbookOf([{ sheetId: 'S', cells: [{ row: 0, col: 0, value: 1 }] }])
  let reads = 0
  const sheet = doc.sheets[0]
  const resolved = sheet.resolvedSparklines
  Object.defineProperty(sheet, 'resolvedSparklines', {
    configurable: true,
    get() { reads++; return resolved },
  })
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  await waitFor(() => expect(ui.getByRole('status').textContent).not.toContain('Loading sheet'))
  const afterInitialRender = reads

  ui.rerender(<InteractiveSheetDoc doc={doc} />)

  expect(reads).toBe(afterInitialRender)
})
