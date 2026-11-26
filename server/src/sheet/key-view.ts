/**
 * Stage 4c (plan P5, S4 §6.2 step 6, §6.12 item 7): the key view, the one sheet every key answer reads: sort keys,
 * filter evaluation, value lists and range detection. A list entry and a filter over that entry therefore agree.
 *
 * It is the text view at the identity view (`NO_SESSION_SIZES`: base coverage only, never the session's sizes or
 * order). Under `'visible-only'` it adds, on the text view's sheet:
 * - painted styles: each content cell's style becomes its painted, CF-effective style (the region style over its own,
 *   through the CF scope), under a fresh `styleIndex` that no other style in the sheet uses, so the engines'
 *   index-based style lookups format labels, display lanes and the date classification as the tile paints;
 * - every cell of a collapsed (zero-width) authored column, blank;
 * - masked strings (Open item 3): a string cell whose painted display text does not contain its raw string reads as
 *   that display text, with no number format. A formatting failure blanks the cell (fail closed).
 * Under `'full'` the key view is the projected sheet itself (local parity, own styles).
 */
import type { XlsxCell, XlsxCellStyle, XlsxDocument, XlsxSheet } from '../../../src/xlsx/types'
import { cellDisplayText } from '../../../src/xlsx/render'
import { NO_SESSION_SIZES } from '../../../src/xlsx/view/axisSizes'
import type { FormatBudget } from '../../../src/xlsx/view/formatWorkspace'
import { cellSourceRow } from '../../../src/xlsx/view/sheetCells'
import type { SheetGeometryResult } from '../../../src/xlsx/view/viewport'
import type { SheetClock } from './clock'
import { blankCell, holdsContent, type SheetSnapshot } from './snapshot'
import { sheetTextView, type TextView } from './text-view'

export interface KeyView {
  sheetIndex: number
  /** `projected` under `'full'`; else the text view's sheet with painted styles, blank zero-width columns and masked
   * strings read as their painted text. */
  sheet: XlsxSheet
  /** Unknown coverage: every cell is blank. */
  blankAll: boolean
  /** Cells holding content in the projected sheet that hold none here (0 under `'full'`). */
  withheld: number
  /** The authored geometry's collapsed (zero-width) column runs, ascending. */
  collapsedCols: ReadonlyArray<{ start: number; end: number }>
  /** Idempotent. The text view's scope is already released when the key view is returned. */
  release(): void
}

type Runs = ReadonlyArray<{ start: number; end: number }>

/** The zero-size column runs of `geometry`, as the layout reads them (`splitAxis`); `undefined` for a diagnostic. */
function collapsedColsOf(geometry: SheetGeometryResult): Runs | undefined {
  if (geometry.kind !== 'ready') return undefined
  return geometry.metrics.cols.overrides.filter((s) => s.sizeCss === 0).map((s) => ({ start: s.start, end: s.end }))
}

function inRuns(runs: Runs, col: number): boolean {
  let a = 0, b = runs.length
  while (a < b) { const mid = (a + b) >> 1; if (runs[mid]!.end < col) a = mid + 1; else b = mid }
  const run = runs[a]
  return run !== undefined && run.start <= col
}

function maxStyleIndex(sheet: XlsxSheet): number {
  let max = -1
  for (const row of sheet.rows) for (const cell of row.cells) if (cell.styleIndex > max) max = cell.styleIndex
  return max
}

/** `painted` without its number format: General. */
function generalOf(painted: XlsxCellStyle | undefined): XlsxCellStyle {
  const { formatCode: _formatCode, ...rest } = painted ?? { numFmtId: 0 }
  return { ...rest, numFmtId: 0 }
}

/** The text view's sheet under P5 (`'visible-only'`, known coverage). `collapsed` undefined (a diagnostic geometry:
 * no column is known to be drawn) blanks every cell. Returns the sheet and how many content cells it blanked. */
function paintKeys(view: TextView, collapsed: Runs | undefined): { sheet: XlsxSheet; blanked: number } {
  const source = view.sheet
  // Format diagnostics land on a copy, never on the cached sheet.
  const formatSheet = { ...source }
  let next = maxStyleIndex(source) + 1
  const fresh = new Map<string, number>()
  const indexOf = (style: XlsxCellStyle) => {
    const key = JSON.stringify(style)
    let index = fresh.get(key)
    if (index === undefined) { index = next++; fresh.set(key, index) }
    return index
  }
  let blanked = 0, changed = false
  const keyed = (row: number, cell: XlsxCell): XlsxCell => {
    if (!holdsContent(cell)) return cell
    if (!collapsed || inRuns(collapsed, cell.col)) { blanked++; return blankCell(cell) }
    const base = view.styleAt(row, cell.col, cell.style)
    const painted = view.scope?.effectiveStyle(row, cell.col, base) ?? base
    if (typeof cell.value === 'string' && cell.valueIsError !== true) {
      let text: string
      try { text = cellDisplayText(formatSheet, { value: cell.value, style: painted }) } catch { blanked++; return blankCell(cell) }
      if (!text.includes(cell.value)) {
        const style = generalOf(painted)
        return { ref: cell.ref, row: cell.row, col: cell.col, value: text, styleIndex: indexOf(style), style }
      }
    }
    return painted === cell.style || painted === undefined ? cell : { ...cell, style: painted, styleIndex: indexOf(painted) }
  }
  const rows = source.rows.map((row) => {
    let cells: XlsxCell[] | undefined
    row.cells.forEach((cell, i) => {
      const out = keyed(cellSourceRow(row, cell), cell)
      if (out === cell) return
      cells ??= [...row.cells]
      cells[i] = out
    })
    if (!cells) return row
    changed = true
    return { ...row, cells }
  })
  return { sheet: changed ? { ...source, rows } : source, blanked }
}

function keyView(sheetIndex: number, sheet: XlsxSheet, blankAll: boolean, withheld: number, collapsedCols: Runs): KeyView {
  return { sheetIndex, sheet, blankAll, withheld, collapsedCols, release() {} }
}

/**
 * The key view of `snap` (P5): `sheetTextView(snap, projected, NO_SESSION_SIZES, o)` plus, under `'visible-only'`,
 * painted styles, blank zero-width columns and masked strings. It takes no session sizes. The text view's CF scope is
 * released before this returns: every style it resolves is resolved by then.
 */
export async function sheetKeyView(snap: SheetSnapshot, projected: XlsxSheet, o: { workbook: XlsxDocument; clock: SheetClock; budget: FormatBudget }): Promise<KeyView> {
  const view = await sheetTextView(snap, projected, NO_SESSION_SIZES, o)
  try {
    const collapsed = collapsedColsOf(view.geometry)
    if (snap.contentPolicy === 'full') return keyView(snap.sheetIndex, projected, false, 0, collapsed ?? [])
    if (view.blankAll) return keyView(snap.sheetIndex, view.sheet, true, view.withheld, collapsed ?? [])
    const { sheet, blanked } = paintKeys(view, collapsed)
    return keyView(snap.sheetIndex, sheet, collapsed === undefined, view.withheld + blanked, collapsed ?? [])
  } finally {
    view.release()
  }
}

/** `base` with each sheet that has a key view replaced by its view's sheet; a sheet without one stays `base`'s (the
 * caller passes the servable workbook, so such a sheet is at least base-blanked). */
export function keyDocument(base: XlsxDocument, views: ReadonlyMap<number, KeyView>): XlsxDocument {
  return { ...base, sheets: base.sheets.map((s, i) => views.get(i)?.sheet ?? s) }
}
