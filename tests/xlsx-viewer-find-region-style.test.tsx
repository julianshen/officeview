/**
 * Final review M2: resident Find formats a cell with the same region-resolved style
 * that paint and copy use (a styled table or a pivot `formats` dxf over its own s=).
 */
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { XlsxDocument } from '../src/xlsx/types'
import type { RegionStyleLayer } from '../src/xlsx/tableStyles/regionLayer'

afterEach(() => { cleanup() })

it('finds a pivot value by the percent format its region applies, not by its own General text', async () => {
  // A non-Excel writer: the 0.00% number format lives only in the pivot's <format> dxf, not in A1's own s=.
  const percent: RegionStyleLayer = {
    rects: [{ minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 }],
    retainedBytes: 0,
    styleAt: (row, col, own) => row === 0 && col === 0 ? { ...(own ?? { numFmtId: 0 }), numFmtId: 10, formatCode: '0.00%' } : own,
  }
  const doc: XlsxDocument = { sheets: [{
    name: 'S', rows: [{ index: 0, cells: [{ ref: 'A1', row: 0, col: 0, value: 0.125, styleIndex: 0, style: { numFmtId: 0 } }] }], cols: [], merges: [], mergeRanges: [],
    regionStyles: percent, regionStylesResolved: true,
  }] } as unknown as XlsxDocument
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.change(ui.getByLabelText('Search document'), { target: { value: '12.50%' } })
  await waitFor(() => expect(ui.getByLabelText('Next match').hasAttribute('disabled')).toBe(false), { timeout: 3_000 })
})
