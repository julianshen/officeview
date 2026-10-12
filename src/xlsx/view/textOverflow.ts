/**
 * Excel text overflow (S13), shared by the static renderer (render.ts) and the
 * interactive viewport planner (viewport.ts): how far a cell's text may run
 * past its own box into empty neighbours. Pure; no rendering dependencies.
 */
import type { XlsxCell } from '../types'

/** Room (px) a cell's text may use past its own box on each side. */
export interface OverflowAllowance { left: number; right: number }
/** One row's neighbour occupancy for textOverflowAllowance. */
export interface RowOverflowInput {
  occupied: readonly number[]
  merges: readonly { minCol: number; maxCol: number }[]
  colStart: (col: number) => number
  lastCol: number
  knownFrom?: number
  knownTo?: number
}
/**
 * Excel text overflow (S13): text may run into EMPTY adjacent cells
 * only, and clips at the first non-empty one (a value, a formula, even one
 * returning "", or any merged cell) or at the sheet edge. Zero-width (hidden)
 * columns are skipped: they neither block nor add room. A merged cell never
 * overflows. Outside [knownFrom, knownTo] occupancy is unknown to the caller
 * (the viewer reads a bounded envelope), so the room there is unbounded; the
 * painted pane never reaches past that envelope.
 */
export function textOverflowAllowance(input: RowOverflowInput, col: number): OverflowAllowance {
  const { occupied, merges, colStart } = input
  const visible = (c0: number, c1: number) => colStart(c1 + 1) - colStart(c0) > 0
  if (merges.some(m => m.minCol <= col && col <= m.maxCol)) return { left: 0, right: 0 }
  let low = 0, high = occupied.length
  while (low < high) { const mid = (low + high) >> 1; if (occupied[mid] <= col) low = mid + 1; else high = mid }
  let blockRight: number | undefined, blockLeft: number | undefined
  for (let i = low; i < occupied.length; i++) if (visible(occupied[i], occupied[i])) { blockRight = occupied[i]; break }
  for (let i = low - 1; i >= 0; i--) if (occupied[i] !== col && visible(occupied[i], occupied[i])) { blockLeft = occupied[i]; break }
  for (const m of merges) {
    if (!visible(m.minCol, m.maxCol)) continue
    if (m.minCol > col && (blockRight === undefined || m.minCol < blockRight)) blockRight = m.minCol
    if (m.maxCol < col && (blockLeft === undefined || m.maxCol > blockLeft)) blockLeft = m.maxCol
  }
  const right = blockRight === undefined
    ? input.knownTo !== undefined ? Infinity : colStart(input.lastCol + 1) - colStart(col + 1)
    : input.knownTo !== undefined && blockRight > input.knownTo ? Infinity : colStart(blockRight) - colStart(col + 1)
  const left = blockLeft === undefined
    ? input.knownFrom !== undefined ? Infinity : colStart(col) - colStart(0)
    : input.knownFrom !== undefined && blockLeft < input.knownFrom ? Infinity : colStart(col) - colStart(blockLeft + 1)
  return { left: Math.max(0, left), right: Math.max(0, right) }
}
/** A cell blocks overflow when it holds a value or a formula. */
export function cellBlocksOverflow(cell: Pick<XlsxCell, 'value' | 'formula'>): boolean {
  return (cell.value !== null && cell.value !== undefined) || cell.formula !== undefined
}
