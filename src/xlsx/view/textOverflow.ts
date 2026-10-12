/**
 * Excel text overflow (S13), shared by the static renderer (render.ts), the
 * interactive viewport planner (viewport.ts) and find highlight geometry: how
 * far a cell's text may run past its own box into empty neighbours. Pure; no
 * rendering dependencies.
 */
import type { XlsxCell } from '../types'

/** Room (px) a cell's text may use past its own box on each side. */
export interface OverflowAllowance { left: number; right: number }
/** One row's neighbour occupancy for textOverflowAllowance. */
export interface RowOverflowInput {
  /** The row's populated cells, ascending by column (formatting-only cells allowed). */
  cells: readonly Pick<XlsxCell, 'col' | 'value' | 'formula'>[]
  /** Merged ranges that intersect the row. */
  merges: readonly { minCol: number; maxCol: number }[]
  /** Left edge of a column in px; colStart(c + 1) - colStart(c) is its width. */
  colStart: (col: number) => number
  /** The worksheet's last column index. */
  lastCol: number
}
export const NO_OVERFLOW: OverflowAllowance = Object.freeze({ left: 0, right: 0 })

/** A cell blocks overflow when it holds a value (an empty-string constant
 * included) or a formula (one returning "" included). */
export function cellBlocksOverflow(cell: Pick<XlsxCell, 'value' | 'formula'>): boolean {
  return (cell.value !== null && cell.value !== undefined) || cell.formula !== undefined
}

/**
 * Text may run into EMPTY adjacent cells only and clips at the first
 * non-empty one (a value or formula cell, or any merged cell) or at the
 * worksheet edge. Zero-width (hidden) columns are skipped: they neither block
 * nor add room. A merged cell never overflows. The scan visits only the
 * cells between `col` and the blocker (formatting-only cells are skipped).
 */
export function textOverflowAllowance(input: RowOverflowInput, col: number): OverflowAllowance {
  const { cells, merges, colStart, lastCol } = input
  const visible = (c0: number, c1: number) => colStart(c1 + 1) - colStart(c0) > 0
  if (merges.some(m => m.minCol <= col && col <= m.maxCol)) return NO_OVERFLOW
  let low = 0, high = cells.length
  while (low < high) { const mid = (low + high) >> 1; if (cells[mid].col <= col) low = mid + 1; else high = mid }
  let blockRight: number | undefined, blockLeft: number | undefined
  for (let i = low; i < cells.length; i++) {
    const c = cells[i].col
    if (c > col && cellBlocksOverflow(cells[i]) && visible(c, c)) { blockRight = c; break }
  }
  for (let i = low - 1; i >= 0; i--) {
    const c = cells[i].col
    if (c < col && cellBlocksOverflow(cells[i]) && visible(c, c)) { blockLeft = c; break }
  }
  for (const m of merges) {
    if (!visible(m.minCol, m.maxCol)) continue
    if (m.minCol > col && (blockRight === undefined || m.minCol < blockRight)) blockRight = m.minCol
    if (m.maxCol < col && (blockLeft === undefined || m.maxCol > blockLeft)) blockLeft = m.maxCol
  }
  const right = colStart(blockRight ?? lastCol + 1) - colStart(col + 1)
  const left = colStart(col) - colStart(blockLeft === undefined ? 0 : blockLeft + 1)
  return { left: Math.max(0, left), right: Math.max(0, right) }
}

/** Horizontal alignments whose single unwrapped line overflows like Excel's
 * left/right/center text. Fill, justify, distributed and center-across-
 * selection keep their own cell semantics and never overflow. */
export function alignmentOverflows(horizontal: string | undefined): boolean {
  return horizontal === undefined || horizontal === 'general' || horizontal === 'left' || horizontal === 'right' || horizontal === 'center'
}
