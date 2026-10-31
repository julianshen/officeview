/**
 * Column autofit after a pivot change (plan Task 29, spec §7.8). Lazy chunk `pivots`.
 *
 * The rule, derived from Excel 16.114's saved `<col>` widths over every Task 24 interaction case
 * (docs/validation/xlsx-pivots-stage3-measurements.md (a); the derivation is in the Task 29 report):
 * - A column's fit is its widest pivot cell: the cell's text, plus the dropdown box on a field-header or page cell
 *   and the compact label indent (± gutter and level) on a row label, plus Excel's padding, in Excel for Mac's
 *   72-dpi pixels rounded up, as characters of the sheet's maximum digit width truncated to 1/256.
 * - After a change a column is resized only when its fit changed AND its width is still the old fit (one pixel of
 *   tolerance): Excel sized it for the old content. A column whose width is not its old fit (set by hand, or kept
 *   from another layout) keeps it, and so does one whose fit did not change, however far its width is from it
 *   (autofit-collapse-narrow: columns at 4 stay at 4). Fits grow and shrink.
 * - A column the pivot newly covers has no old fit: it takes its new fit (interaction-base-tabular-bottom's J).
 *   A column it vacates keeps its width. A column the user resized in this session keeps the user's width.
 * - The work is bounded: at most `AUTOFIT_MAX_CELLS` columns of the pivot and `AUTOFIT_MAX_ROWS` cells per column
 *   are measured (a longer column's fit is that of its first cells).
 */
import type { PivotTableDefinition } from '../model'
import { compactLabelIndents, pivotButtons } from '../buttons'
import { savedPivotLayout } from '../savedLayout'
import { AUTOFIT_MAX_CELLS, AUTOFIT_MAX_ROWS } from '../../view/axisSizes'

/** Excel's padding around a fitted cell's content, in 72-dpi pixels (points). */
export const PIVOT_FIT_PADDING_PT = 10
/** The dropdown box a field-header or page cell carries, in points (buttons.ts: 15 pt wide). */
export const PIVOT_FIT_BUTTON_PT = 15
/** One Excel-for-Mac pixel in characters of a 6-px digit, plus the 1/256 truncation: the tolerance of "the width is
 * still the old fit" (a measure a pixel off Excel's must not read a fitted column as a hand-sized one). */
export const PIVOT_FIT_TOLERANCE = 1 / 6 + 1 / 256

/** One alignment indent level, in points (Excel: three spaces of the default font, 9 px at 72 dpi). */
export const PIVOT_FIT_INDENT_PT = 9
/** The width of a cell's content in points: its shown text in the font it paints with, plus its own alignment indent
 * (`PIVOT_FIT_INDENT_PT` per level) unless the pivot's compact indent replaces it; undefined or 0 for an empty cell. */
export interface PivotFitMeasure { textPt(row: number, col: number): number | undefined }
export interface PivotFitInput { def: PivotTableDefinition; cells: Iterable<{ row: number; col: number }> }

/** The maximum digit width in 72-dpi pixels Excel for Mac uses for a sheet whose 96-dpi one is `mdwPx96` (Calibri 11:
 * 7 px at 96 dpi, 6 px at 72). */
export const macDigitWidth = (mdwPx96: number): number => Math.max(1, Math.ceil(mdwPx96 * 0.75))
/** A fitted width in characters: the content plus padding in whole 72-dpi pixels, over the digit width, to 1/256. */
export const fitChars = (contentPt: number, mdwPt: number): number => Math.trunc(Math.ceil(contentPt + PIVOT_FIT_PADDING_PT - 1e-9) / mdwPt * 256) / 256

/** Each pivot column's fit in characters (see the header). */
export function pivotColumnFits(input: PivotFitInput, measure: PivotFitMeasure, mdwPt: number): Map<number, number> {
  const layout = savedPivotLayout(input.def)
  const buttons = new Set(pivotButtons(input.def, 'screen', layout).filter(button => button.kind === 'filter' || button.kind === 'page').map(button => `${button.row},${button.col}`))
  const indent = layout ? compactLabelIndents(layout) : undefined
  const fits = new Map<number, number>(), counts = new Map<number, number>()
  for (const { row, col } of input.cells) {
    const seen = counts.get(col) ?? 0
    if (seen >= AUTOFIT_MAX_ROWS || (seen === 0 && counts.size >= AUTOFIT_MAX_CELLS)) continue
    counts.set(col, seen + 1)
    const text = measure.textPt(row, col)
    if (!text) continue
    const extra = (buttons.has(`${row},${col}`) ? PIVOT_FIT_BUTTON_PT : 0) + (indent?.(row, col, 'screen') ?? 0) * 0.75
    const chars = fitChars(text + extra, mdwPt)
    if (chars > (fits.get(col) ?? 0)) fits.set(col, chars)
  }
  return fits
}

/**
 * The columns to resize after a change, in characters (see the header's rule). `current` is a column's width now,
 * in characters (undefined: the sheet default); `userSet` names the columns the user resized in this session.
 */
export function pivotAutofit(before: ReadonlyMap<number, number>, after: ReadonlyMap<number, number>, current: (col: number) => number | undefined,
  userSet: (col: number) => boolean): Map<number, number> {
  const out = new Map<number, number>()
  for (const [col, fit] of after) {
    if (userSet(col)) continue
    const old = before.get(col), width = current(col)
    if (old === undefined) { if (width !== fit) out.set(col, fit); continue }
    if (old === fit) continue
    // A column at the sheet default width has no saved size to keep.
    if (width !== undefined && Math.abs(width - old) > PIVOT_FIT_TOLERANCE) continue
    if (width !== fit) out.set(col, fit)
  }
  return out
}
