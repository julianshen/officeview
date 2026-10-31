// Round 3, mounted: what the user sees when recomputeDependents throws its new R2-6 shape guard.
import { afterEach, expect, it } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { XlsxDocument } from '../src/xlsx/types'
import { patchedSeed } from './xlsx-pivots-dependent-r2-patched'

const value = (doc: XlsxDocument, ref: string) => doc.sheets.find(s => s.name === 'Formulas')!.rows.flatMap(r => r.cells).find(c => c.ref === ref)!.value
afterEach(cleanup)

it('R2-6 (Minor): after the cell shape changes post-evaluation, a commit still recomputes (or shows the persistent notice)', async () => {
  const doc = await patchedSeed([])
  // Any change in a sheet's row/cell count after the last full evaluation (here: a host-added blank cell).
  doc.sheets.find(s => s.name === 'Data')!.rows[0]!.cells.push({ ref: 'Z1', col: 25, row: 0, value: null, styleIndex: 0 } as never)
  const ui = render(<InteractiveSheetDoc doc={doc} />)
  fireEvent.click(await ui.findByRole('button', { name: 'Collapse East (Region)' }, { timeout: 10_000 }))
  await ui.findByRole('button', { name: 'Expand East (Region)' }, { timeout: 10_000 })
  await new Promise(resolve => setTimeout(resolve, 1500))
  const text = ui.container.textContent ?? ''
  console.log('B3 after commit:', value(doc, 'B3'), '| announcement present:', /may still show saved values/i.test(text))
  expect(value(doc, 'B3')).toBe('#REF!') // Excel gpd-collapse-item
}, 30_000)
