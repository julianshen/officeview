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

/** The rows (as inclusive intervals) whose font is not uniformly the default: a stored cell in another font, or a row
 * style (`customFormat`), which applies to the row's empty cells. */
export function cfNonUniformFontRows(rows: ReadonlyArray<{ index: number; styled?: boolean; cells: ReadonlyArray<{ style?: XlsxCellStyle }> }>): { start: number; end: number }[] {
  const out: { start: number; end: number }[] = []
  for (const row of [...rows].sort((a, b) => a.index - b.index)) {
    if (row.styled !== true && row.cells.every(c => cfKeyFontStyle(c.style))) continue
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
