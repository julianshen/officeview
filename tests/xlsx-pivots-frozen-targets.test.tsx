/**
 * PR #57 review (P2, PivotButtonLayer): a pivot button's transparent hit target is clipped to its cell's frozen-pane
 * region, as the canvas clips the body under the pinned strips. A body button scrolled under the frozen rows or
 * columns has no target (a click there reaches the pinned cell); one partly under the strip keeps only its visible
 * part. The keyboard route scrolls a hidden pivot cell into view, as the cell cursor does.
 */
import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createScrollWindow, logicalToViewport, setScrollOffset } from '../src/xlsx/view/scrollWindow'
import { PivotButtonLayer, PivotController } from '../src/components/xlsx/pivot/PivotButtonLayer'
import { pivotButtons, savedPivotLayout } from '../src/xlsx/pivots/index'
import InteractiveSheetDoc from '../src/components/xlsx/InteractiveSheetDoc'
import type { PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { XlsxDocument } from '../src/xlsx/types'
import { FIRST_PAINT_WAIT } from './helpers/viewer-timing'

afterEach(() => { cleanup(); vi.restoreAllMocks() })

const FIXTURE = 'tests/fixtures/xlsx-pivots/engine/baseline.xlsx'
async function load(): Promise<XlsxDocument> { return parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(FIXTURE)))) }
const projection = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: 200, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 60, defaultSizeCss: 64 }) })
const viewport = { width: 900, height: 500 }
const controller = () => new PivotController({ client: { open: async () => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'x', message: 'x' } }), apply: async () => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'x', message: 'x' } }), release: () => {} } as never, announce: () => {}, openMenu: () => {}, closeMenu: () => {} })
function targets(def: PivotTableDefinition, frozen: { rows: number; cols: number }, offset: { x: number; y: number }) {
  const window = setScrollOffset(createScrollWindow(projection, viewport, { frozen }), projection, viewport, offset)
  const ui = render(<PivotButtonLayer controller={controller()} sheet={0} defs={[def]} projection={projection} window={window} viewport={viewport} touch={false} />)
  const boxes = new Map([...ui.container.querySelectorAll<HTMLElement>('[data-pivot-key]')].map(el => [el.dataset.pivotKey!, {
    x: Number.parseFloat(el.style.left), y: Number.parseFloat(el.style.top), width: Number.parseFloat(el.style.width), height: Number.parseFloat(el.style.height) }]))
  const cellAt = (row: number, col: number) => { const box = projection.sourceCellToLogical({ row, col })!; const at = logicalToViewport(window, box); return { ...at, width: box.width, height: box.height } }
  ui.unmount()
  return { boxes, window, cellAt }
}
const placements = (def: PivotTableDefinition) => pivotButtons(def, 'screen', savedPivotLayout(def))

describe('pivot hit targets in frozen panes (PR #57 P2)', () => {
  it('fixture: the baseline pivot has body buttons in column A below row 3 and dropdowns in its header row', async () => {
    const def = (await load()).sheets.find(sheet => sheet.pivotTables?.length)!.pivotTables![0]!
    const buttons = placements(def)
    expect(buttons.some(button => button.col === 0 && button.row >= 4 && (button.kind === 'collapse' || button.kind === 'expand'))).toBe(true)
    expect(buttons.some(button => button.kind === 'filter' && button.col >= 1)).toBe(true)
  })

  it('a body button scrolled under the frozen rows has no target; one partly under the strip is cut at its edge', async () => {
    const def = (await load()).sheets.find(sheet => sheet.pivotTables?.length)!.pivotTables![0]!
    const drills = placements(def).filter(button => button.kind === 'collapse' || button.kind === 'expand')
    // Rows 0–2 frozen (60 px); the body scrolled by 3.5 rows: rows 3–5 and half of row 6 are under the strip.
    const { boxes } = targets(def, { rows: 3, cols: 0 }, { x: 0, y: 70 })
    const top = 60
    for (const button of drills) {
      const box = boxes.get(`0:${button.row}:${button.col}`)
      if (button.row <= 5) expect(box, `row ${button.row}`).toBeUndefined()
      else if (box) expect(box.y, `row ${button.row}`).toBeGreaterThanOrEqual(top)
    }
    // Row 6 shows its lower half only: a drill target there (centred in the row) is cut to the strip's edge.
    const partial = drills.find(button => button.row === 6)
    if (partial) {
      const box = boxes.get(`0:6:${partial.col}`)
      if (box) { expect(box.y).toBe(top); expect(box.height).toBeLessThan(20) }
    }
    // Frozen-row buttons (the header row 3 = index 2) stay where they are pinned.
    for (const button of placements(def).filter(button => button.row <= 2)) expect(boxes.has(`0:${button.row}:${button.col}`)).toBe(true)
  })

  it('a body button scrolled under the frozen columns has no target; the pinned column\'s buttons stay', async () => {
    const def = (await load()).sheets.find(sheet => sheet.pivotTables?.length)!.pivotTables![0]!
    // Column A frozen (64 px); the body scrolled by 1.5 columns: B and half of C are under the strip.
    const { boxes } = targets(def, { rows: 0, cols: 1 }, { x: 96, y: 0 })
    const all = placements(def)
    for (const button of all) {
      const box = boxes.get(`0:${button.row}:${button.col}`)
      if (button.col === 0) expect(box, `pinned ${button.row},${button.col}`).toBeDefined()
      else if (button.col === 1) expect(box, `hidden ${button.row},${button.col}`).toBeUndefined()
      else if (box) expect(box.x).toBeGreaterThanOrEqual(64)
    }
  })

  it('a dropdown partly under the frozen columns is cut at the strip; its visible part still opens it', async () => {
    const def = (await load()).sheets.find(sheet => sheet.pivotTables?.length)!.pivotTables![0]!
    const dropdown = placements(def).find(button => button.kind === 'filter' && button.col >= 1)!
    // Scroll so the dropdown (at the cell's right edge) straddles the strip's edge (column A frozen, 64 px).
    const unclipped = targets(def, { rows: 0, cols: 0 }, { x: 0, y: 0 }).boxes.get(`0:${dropdown.row}:${dropdown.col}`)!
    const shift = unclipped.x + unclipped.width / 2 - 64
    const { boxes } = targets(def, { rows: 0, cols: 1 }, { x: shift, y: 0 })
    const box = boxes.get(`0:${dropdown.row}:${dropdown.col}`)!
    expect(box.x).toBe(64)
    expect(box.width).toBeCloseTo(unclipped.width / 2, 5)
  })

  it('keyboard: Alt+Left on a pivot cell hidden under the frozen rows scrolls it into view, as the cursor does', async () => {
    const doc = await load()
    const sheet = doc.sheets.find(candidate => candidate.pivotTables?.length)!
    sheet.frozen = { rows: 3, cols: 0, topRow: 3, leftCol: 0 } as never
    const ui = render(<InteractiveSheetDoc doc={doc} />)
    await ui.findByRole('button', { name: /^Collapse East/ }, FIRST_PAINT_WAIT)
    const grid = ui.getByRole('grid')
    act(() => grid.focus())
    // Put the cursor on East's ± cell, then scroll it under the frozen strip.
    const east = () => ui.queryByRole('button', { name: /^Collapse East/ })
    const row = Number(east()!.dataset.row), col = Number(east()!.dataset.col)
    fireEvent.keyDown(grid, { key: 'Home', ctrlKey: true })
    for (let k = 0; k < row; k++) fireEvent.keyDown(grid, { key: 'ArrowDown' })
    for (let k = 0; k < col; k++) fireEvent.keyDown(grid, { key: 'ArrowRight' })
    expect(grid.getAttribute('aria-activedescendant')).toMatch(new RegExp(`-cell-${row * 16_384 + col}$`))
    const scroller = ui.getByTestId('officeview-sheet-scroller')
    act(() => { scroller.scrollTop = 400; fireEvent.scroll(scroller) })
    await waitFor(() => expect(east()).toBeNull())
    fireEvent.keyDown(grid, { key: 'ArrowLeft', altKey: true })
    // The cell comes back into view (its target with it) and the change applies.
    await waitFor(() => expect(ui.queryByRole('button', { name: /^Expand East/ })).not.toBeNull(), FIRST_PAINT_WAIT)
  }, 30_000)
})
