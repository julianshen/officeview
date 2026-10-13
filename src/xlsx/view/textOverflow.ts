/**
 * Excel text overflow (S13), shared by the static renderer (render.ts), the
 * interactive viewport planner (viewport.ts) and find highlight geometry: how
 * far a cell's text may run past its own box into empty neighbours. Pure; no
 * rendering dependencies.
 */
import type { XlsxCell } from '../types'

/** Room (px) a cell's text may use past its own box on each side. */
export interface OverflowAllowance { left: number; right: number }
/** One row's neighbour occupancy for textOverflowAllowance, built once per
 * row (rowOverflowInput) so that each allowance is a binary search. */
export interface RowOverflowInput {
  /** Visible (non-zero-width) columns whose cell blocks overflow, ascending and unique. */
  blockers: readonly number[]
  /** Merged ranges that intersect the row and have visible width. */
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
 * The once-per-row input: the row's cells (any order; formatting-only cells
 * allowed) are reduced to the sorted visible blocking columns. Zero-width
 * (hidden) columns are dropped here, so they neither block nor add room.
 */
export function rowOverflowFromCells(cells: readonly Pick<XlsxCell, 'col' | 'value' | 'formula'>[], merges: readonly { minCol: number; maxCol: number }[], colStart: (col: number) => number, lastCol: number): RowOverflowInput {
  const visible = (c0: number, c1: number) => colStart(c1 + 1) - colStart(c0) > 0
  const blockers: number[] = []
  let sorted = true
  for (const cell of cells) {
    if (!cellBlocksOverflow(cell) || !visible(cell.col, cell.col)) continue
    if (blockers.length && cell.col <= blockers[blockers.length - 1]) sorted = false
    blockers.push(cell.col)
  }
  if (!sorted) blockers.sort((a, b) => a - b)
  let unique = 0
  for (const col of blockers) {
    if (unique === 0 || col !== blockers[unique - 1]) blockers[unique++] = col
  }
  blockers.length = unique
  return { blockers, merges: merges.filter(m => visible(m.minCol, m.maxCol)), colStart, lastCol }
}

/**
 * Text may run into EMPTY adjacent cells only and clips at the first
 * non-empty one (a value or formula cell, or any merged cell) or at the
 * worksheet edge. A merged cell never overflows. O(log blockers + merges in
 * the row) per call.
 */
export function textOverflowAllowance(input: RowOverflowInput, col: number): OverflowAllowance {
  const { blockers, merges, colStart, lastCol } = input
  if (merges.some(m => m.minCol <= col && col <= m.maxCol)) return NO_OVERFLOW
  let low = 0, high = blockers.length
  while (low < high) { const mid = (low + high) >> 1; if (blockers[mid] <= col) low = mid + 1; else high = mid }
  let blockRight = low < blockers.length ? blockers[low] : undefined
  let index = low - 1
  if (index >= 0 && blockers[index] === col) index--
  let blockLeft = index >= 0 ? blockers[index] : undefined
  for (const m of merges) {
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
