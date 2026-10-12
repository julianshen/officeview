/**
 * S10.1 reproduced gap (Chromium, F7-shaped sheet): one authored 0.75pt row in
 * an ordinary 41-row viewport turned the whole owned grid coarse — no column
 * headers, no rows, no gridcells, only "Zoom in or inspect a cell for detail" —
 * while the painter kept every other row in detail (it summarizes sub-detail
 * runs as coarse groups, spec §11.1). A11Y-01 requires owned rows/cells for a
 * bounded viewport; only a sub-detail line itself may go without a cell.
 */
import { describe, expect, it } from 'vitest'
import { render } from '@testing-library/react'
import { GridSemantics } from '../src/components/xlsx/GridSemantics'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createScrollWindow } from '../src/xlsx/view/scrollWindow'
import { emptyGridSelection, selectCell } from '../src/xlsx/view/selection'
import { gridFixture } from './helpers/xlsx-grid-fixture'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'

const visibility = visibilityFromHidden(createIntervalSet([]))

describe('owned grid with one sub-detail row', () => {
  it('keeps headers, rows and cells for every detail line; the tiny row alone has no owned row', () => {
    const projection = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: [{ start: 2, end: 2, sizeCss: 1 }] }),
      cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }),
      visibility,
    })
    const viewport = { width: 400, height: 200 }
    const selection = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
    const props = gridFixture()
    const ui = render(<GridSemantics {...props} projection={projection} view={{ ...props.view, visibility }} viewport={viewport} window={createScrollWindow(projection, viewport)} selection={selection} body={{ r0: 0, r1: 20 }} />)
    const rows = ui.getAllByRole('row').map(row => row.getAttribute('aria-rowindex'))
    // aria-rowindex = display slot + 2: rows 1-2 and 4 are owned, the 1px row 3 is summarized.
    expect(rows).toEqual(expect.arrayContaining(['2', '3', '5']))
    expect(rows).not.toContain('4')
    expect(ui.getAllByRole('columnheader').length).toBeGreaterThan(1)
    expect(ui.getByRole('gridcell', { name: /^B4, value / })).toBeTruthy()
    expect(ui.getByRole('grid').getAttribute('aria-description')).not.toContain('Coarse')
  })

  it('a navigated cell on the sub-detail row stays exactly accessible', () => {
    const projection = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: [{ start: 2, end: 2, sizeCss: 1 }] }),
      cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }),
      visibility,
    })
    const viewport = { width: 400, height: 200 }
    const selection = selectCell(emptyGridSelection(), projection, { row: 2, col: 1 })
    const props = gridFixture()
    const ui = render(<GridSemantics {...props} projection={projection} view={{ ...props.view, visibility }} viewport={viewport} window={createScrollWindow(projection, viewport)} selection={selection} body={{ r0: 0, r1: 20 }} />)
    const active = ui.getByRole('gridcell', { name: /^B3, value / })
    expect(ui.getByRole('grid').getAttribute('aria-activedescendant')).toBe(active.id)
    expect(ui.getAllByRole('gridcell').length).toBeGreaterThan(10)
  })

  it('shares the painter\'s detail boundary: a 1pt row at 120% x DPR 2.5 (3.9999999999999996 px) is detail in both', () => {
    // viewport.ts paints a line in detail at >= 4 - 1e-9 device px; ownership must agree.
    const projection = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20, overrides: [{ start: 2, end: 2, sizeCss: 4 / 3 }] }),
      cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 80 }),
      visibility,
    })
    expect((4 / 3) * 1.2 * 2.5).toBeLessThan(4) // the old size*zoom*dpr product sits just under the nominal threshold
    const viewport = { width: 400, height: 200 }
    const selection = selectCell(emptyGridSelection(), projection, { row: 0, col: 0 })
    const props = gridFixture()
    const ui = render(<GridSemantics {...props} projection={projection} view={{ ...props.view, visibility }} viewport={viewport} window={createScrollWindow(projection, viewport, { zoom: 1.2, deviceScale: 2.5 })} selection={selection} body={{ r0: 0, r1: 20 }} />)
    expect(ui.getAllByRole('row').map(row => row.getAttribute('aria-rowindex'))).toContain('4')
  })
})
