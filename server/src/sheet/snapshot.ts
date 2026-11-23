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
import type { ResolvedSparkline, ResolvedSparklines } from '../../../src/xlsx/sparklines/resolve'
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

/** A copy of `sheet` in which every covered cell is blank and draws no sparkline located there (`sheet` itself when
 * nothing changes). Rows without a covered cell are shared with the original. */
export function blankCovered(sheet: XlsxSheet, coverage: Extract<BaseCoverage, { kind: 'known' }>): XlsxSheet {
  const sparklines = withoutCoveredSparklines(sheet.resolvedSparklines, coverage)
  if (coverage.count === 0) return sparklines === sheet.resolvedSparklines ? sheet : { ...sheet, resolvedSparklines: sparklines }
  let changed = sparklines !== sheet.resolvedSparklines
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
  return changed ? { ...sheet, rows, ...(sparklines !== sheet.resolvedSparklines ? { resolvedSparklines: sparklines } : {}) } : sheet
}

type Sparklines = ResolvedSparklines | undefined
/** The resolved sparklines minus those whose location cell is covered (the map itself when none is). Sparklines are
 * usually located in empty cells, so this does not depend on `count`. */
function withoutCoveredSparklines(map: Sparklines, coverage: Extract<BaseCoverage, { kind: 'known' }>): Sparklines {
  if (!map?.size) return map
  let out: Map<number, ReadonlyMap<number, ResolvedSparkline>> | undefined
  for (const [row, cols] of map) {
    let kept: Map<number, ResolvedSparkline> | undefined
    for (const col of cols.keys()) {
      if (!coverage.covered(row, col)) continue
      kept ??= new Map(cols)
      kept.delete(col)
    }
    if (!kept) continue
    out ??= new Map(map)
    if (kept.size) out.set(row, kept)
    else out.delete(row)
  }
  return out ?? map
}

/** One sheet as every answer may read it: the blanked copy and its base coverage (`'visible-only'`), or the projected
 * sheet itself (`'full'`, or unknown coverage). Clock-free, so a host computes it once per sheet and shares it between
 * snapshots and the cross-sheet views of other sheets' snapshots. */
export interface ServableSheet { sheet: XlsxSheet; coverage: BaseCoverage | 'off'; geometry: SheetGeometryResult }
export function servableSheet(projected: XlsxSheet, contentPolicy: 'visible-only' | 'full'): ServableSheet {
  if (contentPolicy === 'full') return { sheet: projected, coverage: 'off', geometry: buildSheetGeometry(projected) }
  const authored = buildSheetGeometry(projected)
  const coverage = baseCoverage(projected, authored)
  const sheet = coverage.kind === 'known' ? blankCovered(projected, coverage) : projected
  return { sheet, coverage, geometry: sheet === projected ? authored : buildSheetGeometry(sheet) }
}

/**
 * The snapshot of `doc.sheets[sheetIndex]`. The order is fixed: coverage (author view, over the projected sheet), then
 * blanking, then `prepare` over the blanked sheet. `servable` supplies a sheet's coverage and blanked copy (default:
 * computed here); a host passes its per-document cache, and its `prepare` reads every other sheet through the same
 * cache, so a cross-sheet rule never reads another sheet's covered cell. `prepare` owns the lease it returns.
 */
export async function buildSheetSnapshot(doc: XlsxDocument, sheetIndex: number, contentPolicy: 'visible-only' | 'full', clock: SheetClock, prepare: (sheet: XlsxSheet) => Promise<SheetAssetsLease>, servable: (sheetIndex: number) => ServableSheet = (i) => servableSheet(doc.sheets[i]!, contentPolicy)): Promise<SheetSnapshot> {
  if (!doc.sheets[sheetIndex]) throw new RangeError(`no sheet ${sheetIndex}`)
  const { sheet, coverage, geometry } = servable(sheetIndex)
  const lease = await prepare(sheet)
  let released = false
  return {
    sheetIndex, sheet, geometry, assets: lease.assets, coverage, clock: { day: clock.day, timeZone: clock.timeZone }, contentPolicy,
    release() { if (!released) { released = true; lease.dispose() } },
  }
}
