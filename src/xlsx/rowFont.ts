/**
 * Ruling S4-ROW-UNIFORM (Task 30 fix round): Excel lifts a bottom-aligned icon onto its row's shared baseline, which a
 * larger font anywhere in the row moves (stage4-icon-geometry C3, M3..O3; stage4-icon-geometry-2 C2, G2, I2, J2: 2 px
 * above the measured frame in rows holding a 20 pt cell). The measured bottom frame stands only on rows like
 * stage4-icon-rows-bottom's, whose every cell is the default Calibri 11 regular (its resaved styles.xml: every cellXfs
 * entry uses fontId 0, `<sz val="11"/><name val="Calibri"/>`; no row carries a row style). A row's font uniformity is
 * computed by its data source (the resident sheet, or the worker's record reader) and travels with the row heights.
 * Kept tiny and outside `cf/` (a paged worker chunk may not import CF code): the worker reader, the layout and the
 * icon painter all import it.
 */
import type { XlsxCellStyle } from './types'

/** Whether a cell style is the measured default font: Calibri 11, neither bold nor italic. Unknown is not. */
export function cfKeyFontStyle(style: Pick<XlsxCellStyle, 'fontName' | 'fontSizePt' | 'bold' | 'italic'> | undefined): boolean {
  return style !== undefined && style.fontName === 'Calibri' && style.fontSizePt === 11 && style.bold !== true && style.italic !== true
}

/** Every worksheet row: the sheet-wide interval (Task 30 fix round 2). */
export const CF_ALL_ROWS: Readonly<{ start: number; end: number }> = Object.freeze({ start: 0, end: 1_048_575 })

/**
 * The rows (as inclusive intervals) whose font is not uniformly the default (fix round 2, codex P1). Every row is when
 * `blocked`: the sheet's default cell style (xf 0, the Normal style's font, which empty cells take) or a `<cols>` style
 * is not, or not known to be, Calibri 11 regular. Otherwise a row is when it carries a row style (`customFormat`, which
 * its empty cells take), a stored cell in another font, or a run-formatted string (`richText`: rich-text runs may name
 * any font; withheld whatever they name, the simplest fail-closed reading).
 */
export function cfNonUniformFontRows(rows: ReadonlyArray<{ index: number; styled?: boolean; richText?: boolean; cells: ReadonlyArray<{ style?: XlsxCellStyle }> }>, blocked = false): { start: number; end: number }[] {
  if (blocked) return [{ ...CF_ALL_ROWS }]
  const out: { start: number; end: number }[] = []
  for (const row of [...rows].sort((a, b) => a.index - b.index)) {
    if (row.styled !== true && row.richText !== true && row.cells.every(c => cfKeyFontStyle(c.style))) continue
    const last = out[out.length - 1]
    if (last && last.end + 1 === row.index) last.end = row.index
    else if (!last || last.end < row.index) out.push({ start: row.index, end: row.index })
  }
  return out
}

/** A row's uniformity from the layout's intervals: undefined (unknown: withheld) when the source sent none. */
export function cfRowUniformOfLayout(nonUniform: ReadonlyArray<{ start: number; end: number }> | undefined): (row: number) => boolean | undefined {
  if (!nonUniform) return () => undefined
  return row => {
    let lo = 0, hi = nonUniform.length - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1, r = nonUniform[mid]!
      if (row < r.start) hi = mid - 1
      else if (row > r.end) lo = mid + 1
      else return false
    }
    return true
  }
}

type RowRange = { start: number; end: number }
/** The sorted union of row intervals. */
function mergeRows(parts: readonly RowRange[]): RowRange[] {
  const out: RowRange[] = []
  for (const r of [...parts].sort((a, b) => a.start - b.start)) {
    const last = out[out.length - 1]
    if (last && r.start <= last.end + 1) last.end = Math.max(last.end, r.end)
    else out.push({ start: r.start, end: r.end })
  }
  return out
}

/** What `cfSheetNonUniformFontRows` reads of a sheet: the same fields on the resident sheet and the paged one. */
export interface CfRowFontSheet {
  rows?: ReadonlyArray<{ index: number; styled?: boolean; richText?: boolean; cells: ReadonlyArray<{ style?: XlsxCellStyle }> }>
  rowFontsBlocked?: boolean
  conditionalFormatting?: { unresolvedSheet?: unknown; rules: ReadonlyArray<{ ranges: ReadonlyArray<{ firstRow: number; lastRow: number }>; writes?: ReadonlySet<string>; dxf?: Readonly<Record<string, unknown>> }> }
  tables?: ReadonlyArray<{ extent: { firstRow: number; rows: number } }>
}

/**
 * Ruling S4-ROW-EFFECTIVE (final review I1): the rows whose effective font is not known to be uniformly Calibri 11
 * regular. Besides the stored formatting (`cfNonUniformFontRows`): every row a conditional-formatting rule that writes
 * a font (`font.*`, a DXF `<font>`) can reach, every row a table covers (its style or table DXF may carry a font), and
 * every row when the sheet's rules are refused as a whole or its tables are unknown. Computed from the sheet model,
 * identical on the resident and the paged paths (their CF rules and table inventories are the same).
 */
export function cfSheetNonUniformFontRows(sheet: CfRowFontSheet, nonUniform?: readonly RowRange[]): RowRange[] {
  const base = nonUniform ?? cfNonUniformFontRows(sheet.rows ?? [], sheet.rowFontsBlocked)
  const cf = sheet.conditionalFormatting
  if (cf?.unresolvedSheet !== undefined || sheet.tables === undefined) return [{ ...CF_ALL_ROWS }]
  const parts: RowRange[] = [...base]
  for (const rule of cf?.rules ?? []) {
    const font = [...(rule.writes ?? [])].some(k => k.startsWith('font')) || Object.keys(rule.dxf ?? {}).some(k => k.startsWith('font'))
    if (font) for (const r of rule.ranges) parts.push({ start: Math.max(0, r.firstRow), end: Math.min(CF_ALL_ROWS.end, r.lastRow) })
  }
  for (const t of sheet.tables) if (t.extent.rows > 0) parts.push({ start: t.extent.firstRow, end: Math.min(CF_ALL_ROWS.end, t.extent.firstRow + t.extent.rows - 1) })
  return mergeRows(parts)
}
