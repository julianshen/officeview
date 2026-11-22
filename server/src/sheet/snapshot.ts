/**
 * Stage 4a (S4 §6.2, R8): the servable snapshot of one sheet. Every tile and layout answer about a sheet derives from
 * it, and it is built from the projected package's saved-values parse (`OpenDoc.savedParse`, F1). Under
 * `'visible-only'` with known coverage, every base-covered cell is blank in the snapshot itself (step 5): its value is
 * removed and its style kept, and the assets (CF above all) are prepared over that blanked copy, so a covered cell
 * draws no data bar, colour scale or icon either. Under `'full'` the withholding steps are the identity.
 */
import type { SheetAssetsLease } from '../../../src/components/xlsx/sheetAssets'
import type { RenderSheetAssets } from '../../../src/xlsx/render'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../../src/xlsx/types'
import { cellSourceRow } from '../../../src/xlsx/view/sheetCells'
import { buildSheetGeometry, type SheetGeometryResult } from '../../../src/xlsx/view/viewport'
import type { SheetClock } from './clock'
import { baseCoverage, type BaseCoverage } from './coverage'

export interface SheetSnapshot {
  sheetIndex: number
  /** The sheet every answer reads: the projected sheet, or its blanked copy (visible-only, known coverage). */
  sheet: XlsxSheet
  /** Built over `sheet` (so `cellsOf` returns the blanked cells). */
  geometry: SheetGeometryResult
  assets: RenderSheetAssets
  /** `'off'` under `'full'`. */
  coverage: BaseCoverage | 'off'
  clock: SheetClock
  contentPolicy: 'visible-only' | 'full'
  /** Releases the asset lease (idempotent). The host releases cached snapshots on `drop`. */
  release(): void
}

/** A covered cell with its value removed: position and style only. */
function blankCell(cell: XlsxCell): XlsxCell {
  return { ref: cell.ref, row: cell.row, col: cell.col, value: null, styleIndex: cell.styleIndex, ...(cell.style ? { style: cell.style } : {}) }
}

/** A copy of `sheet` in which every covered cell is blank (`sheet` itself when none is). Rows without a covered cell are
 * shared with the original. */
export function blankCovered(sheet: XlsxSheet, coverage: Extract<BaseCoverage, { kind: 'known' }>): XlsxSheet {
  if (coverage.count === 0) return sheet
  let changed = false
  const rows = sheet.rows.map((row) => {
    let cells: XlsxCell[] | undefined
    row.cells.forEach((cell, i) => {
      if (!coverage.covered(cellSourceRow(row, cell), cell.col)) return
      if (cell.value === null && cell.formula === undefined) return
      cells ??= [...row.cells]
      cells[i] = blankCell(cell)
    })
    if (!cells) return row
    changed = true
    return { ...row, cells }
  })
  return changed ? { ...sheet, rows } : sheet
}

/**
 * The snapshot of `doc.sheets[sheetIndex]`. The order is fixed: coverage (author view, over the projected sheet), then
 * blanking, then `prepare` over the blanked sheet. `prepare` owns the lease it returns; a failure after it is released.
 */
export async function buildSheetSnapshot(doc: XlsxDocument, sheetIndex: number, contentPolicy: 'visible-only' | 'full', clock: SheetClock, prepare: (sheet: XlsxSheet) => Promise<SheetAssetsLease>): Promise<SheetSnapshot> {
  const projected = doc.sheets[sheetIndex]
  if (!projected) throw new RangeError(`no sheet ${sheetIndex}`)
  let sheet = projected
  let geometry: SheetGeometryResult | undefined
  let coverage: BaseCoverage | 'off' = 'off'
  if (contentPolicy === 'visible-only') {
    geometry = buildSheetGeometry(projected)
    coverage = baseCoverage(projected, geometry)
    if (coverage.kind === 'known') sheet = blankCovered(projected, coverage)
  }
  if (!geometry || sheet !== projected) geometry = buildSheetGeometry(sheet)
  const lease = await prepare(sheet)
  let released = false
  return {
    sheetIndex, sheet, geometry, assets: lease.assets, coverage, clock: { day: clock.day, timeZone: clock.timeZone }, contentPolicy,
    release() { if (!released) { released = true; lease.dispose() } },
  }
}
