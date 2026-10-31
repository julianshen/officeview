/**
 * The faithfulness check compares every cell of the saved location and page fields; past a fixed area it returns
 * unverified instead of scanning (PR #56 review: a pivot with large axes and no data fields computed in ms but was
 * scanned cell by cell for seconds). A real oversized saved layout needs thousands of items, so the saved layout of
 * a pivot named BIG is replaced by one whose location spans 16,384 × 1,000 cells.
 */
import { describe, expect, it, vi } from 'vitest'
import { resolve } from 'node:path'

vi.mock('../src/xlsx/pivots/savedLayout', async importOriginal => {
  const real = await importOriginal<typeof import('../src/xlsx/pivots/savedLayout')>()
  return {
    ...real,
    savedPivotLayout: (def: Parameters<typeof real.savedPivotLayout>[0]) => {
      const layout = real.savedPivotLayout(def)
      if (!layout || def.name !== 'BIG') return layout
      const huge = { minRow: 2, minCol: 0, maxRow: 1001, maxCol: 16383 }
      return { ...layout, rect: huge, rects: [huge] }
    },
  }
})

const { loadAll } = await import('./helpers/xlsx-pivots-load')
const { checkSavedLayout, sheetCellReader } = await import('../src/xlsx/pivots/engine/faithful')

describe('faithfulness check area cap', () => {
  it('a saved area past the cap is unverified (cells-over-cap) without scanning it', async () => {
    const [p] = await loadAll(resolve(process.cwd(), 'tests/fixtures/xlsx-pivots/engine/baseline.xlsx'))
    expect(checkSavedLayout(p!.store, p!.cache, p!.def, sheetCellReader(p!.sheet), { date1904: false })).toEqual({ status: 'ok' })
    const started = performance.now()
    const result = checkSavedLayout(p!.store, p!.cache, { ...p!.def, name: 'BIG' }, sheetCellReader(p!.sheet), { date1904: false })
    expect(performance.now() - started).toBeLessThan(1000)
    expect(result).toMatchObject({ status: 'unverified', reason: 'cells-over-cap' })
  })
})
