/** Full zero-based worksheet domain; spec §3 (CAP-01).
 * Negative zero is rejected so source identities always use canonical positive zero. */
export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384
export const MAX_CELL_ID = 2 ** 34 - 1
export type SourceRow = number
export type SourceCol = number
export type CellId = number
export interface SourceCell { row: SourceRow; col: SourceCol }
/** Inclusive source-space bounds, never a flattened multi-row ID interval. */
export interface SourceRect { r0: SourceRow; r1: SourceRow; c0: SourceCol; c1: SourceCol }

/** Exact cover of populated cells as disjoint source rects (Task 9): row runs
 * merged into identical-span slabs, so dense batches keep one hull rect while
 * sparse batches never pay a dense hull. Total area always equals the cell
 * count; callers, not this helper, enforce service caps per chunk. */
export function coverCellsToRects(cells: readonly { row: number; col: number }[]): SourceRect[] {
  const sorted = [...cells].sort((a, b) => a.row - b.row || a.col - b.col)
  type Run = { row: number; c0: number; c1: number }
  const runs: Run[] = []
  for (const cell of sorted) {
    if (!Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.col)) continue
    const last = runs[runs.length - 1]
    if (last && last.row === cell.row && cell.col <= last.c1 + 1) {
      if (cell.col > last.c1) last.c1 = cell.col
    } else runs.push({ row: cell.row, c0: cell.col, c1: cell.col })
  }
  const rects: SourceRect[] = []
  let open: { r0: number; r1: number; c0: number; c1: number } | undefined
  const flush = () => { if (open) { rects.push({ r0: open.r0, r1: open.r1, c0: open.c0, c1: open.c1 }); open = undefined } }
  for (const run of runs) {
    if (open && open.r1 + 1 === run.row && open.c0 === run.c0 && open.c1 === run.c1) open.r1 = run.row
    else { flush(); open = { r0: run.row, r1: run.row, c0: run.c0, c1: run.c1 } }
  }
  flush()
  return rects
}
export type DisplaySlot = number
export type NavigatorRequest =
  | { addressSpace: 'display'; at: { rowSlot: DisplaySlot; col: SourceCol } }
  | { addressSpace: 'source'; at: SourceCell }
export type CoordinateErrorCode = 'invalid-coordinate' | 'invalid-cell-id' | 'invalid-a1' | 'invalid-source-rect'
export class CoordinateError extends RangeError {
  constructor(readonly code: CoordinateErrorCode, message: string) {
    super(message)
    this.name = 'CoordinateError'
  }
}
export function isValidCoordinate(row: number, col: number): boolean {
  return Number.isSafeInteger(row) && !Object.is(row, -0) && row >= 0 && row < MAX_ROWS
    && Number.isSafeInteger(col) && !Object.is(col, -0) && col >= 0 && col < MAX_COLS
}

export function validateCoordinate(row: number, col: number): void {
  if (!isValidCoordinate(row, col)) {
    throw new CoordinateError('invalid-coordinate', 'Coordinate is outside the worksheet domain')
  }
}

export function cellId(row: SourceRow, col: SourceCol): CellId {
  validateCoordinate(row, col)
  // Arithmetic preserves all 34 bits; JavaScript bitwise packing would alias rows.
  return row * MAX_COLS + col
}

/** Typed caller boundary; untrusted values use cellIdToCoordinateInput. */
export function cellIdToCoordinate(id: CellId): SourceCell {
  return cellIdToCoordinateInput(id)
}

export function cellIdToCoordinateInput(id: unknown): SourceCell {
  if (typeof id !== 'number' || !Number.isSafeInteger(id) || Object.is(id, -0) || id < 0 || id > MAX_CELL_ID) {
    throw new CoordinateError('invalid-cell-id', 'Cell ID is outside the worksheet domain')
  }
  return { row: Math.floor(id / MAX_COLS), col: id % MAX_COLS }
}

/** Accept ASCII case and optional absolute markers; emit canonical relative A1.
 * Sheet names, ranges, whitespace and leading-zero row labels are not cell addresses. */
export function parseA1(address: string): SourceCell {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(address)
  if (!match) {
    throw new CoordinateError('invalid-a1', 'Invalid A1 cell address')
  }
  const col = lettersToColumn(match[1])
  const row = Number(match[2]) - 1
  if (!isValidCoordinate(row, col)) {
    throw new CoordinateError('invalid-a1', 'A1 address is outside the worksheet domain')
  }
  return { row, col }
}

/** An A1 cell or range reference (`A1`, `$A$1:$C$4`, either corner order) as a rectangle; undefined when either
 * corner is not a cell address. The shared parser behind every pivot location and region ref. */
export function parseA1Range(ref: string): SourceRect | undefined {
  const parts = ref.split(':')
  if (parts.length > 2) return undefined
  try {
    const a = parseA1(parts[0]!), b = parts.length === 2 ? parseA1(parts[1]!) : a
    return { r0: Math.min(a.row, b.row), r1: Math.max(a.row, b.row), c0: Math.min(a.col, b.col), c1: Math.max(a.col, b.col) }
  } catch { return undefined }
}

export function formatA1(at: SourceCell): string {
  validateCoordinate(at.row, at.col)
  return `${columnToLetters(at.col)}${at.row + 1}`
}

export function columnToLetters(col: SourceCol): string {
  validateCoordinate(0, col)
  let remaining = col + 1
  let letters = ''
  while (remaining > 0) {
    remaining--
    letters = String.fromCharCode(65 + remaining % 26) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

/** Column-only labels accept lowercase, but not `$` markers. */
export function lettersToColumn(letters: string): SourceCol {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) {
    throw new CoordinateError('invalid-a1', 'Invalid A1 column letters')
  }
  let ordinal = 0
  for (const letter of letters.toUpperCase()) {
    ordinal = ordinal * 26 + letter.charCodeAt(0) - 64
  }
  if (ordinal > MAX_COLS) {
    throw new CoordinateError('invalid-a1', 'Column is outside the worksheet domain')
  }
  return ordinal - 1
}

/** Already-typed rectangles remain SourceRect in both branches; geometry may be invalid. */
export function isValidSourceRect(rect: SourceRect): boolean {
  return isValidSourceRectInput(rect)
}

/** Untrusted-input predicate narrows only after validating every coordinate. */
export function isValidSourceRectInput(rect: unknown): rect is SourceRect {
  if (typeof rect !== 'object' || rect === null) return false
  const candidate = rect as Partial<SourceRect>
  if (typeof candidate.r0 !== 'number' || typeof candidate.r1 !== 'number'
    || typeof candidate.c0 !== 'number' || typeof candidate.c1 !== 'number') return false
  return isValidCoordinate(candidate.r0, candidate.c0) && isValidCoordinate(candidate.r1, candidate.c1)
    && candidate.r0 <= candidate.r1 && candidate.c0 <= candidate.c1
}

export function validateSourceRect(rect: SourceRect): void {
  validateSourceRectInput(rect)
}

/** Untrusted validators throw the same stable errors as the typed APIs. */
export function validateSourceRectInput(rect: unknown): asserts rect is SourceRect {
  if (!isValidSourceRectInput(rect)) {
    throw new CoordinateError('invalid-source-rect', 'Invalid inclusive source rectangle')
  }
}

export function sourceRectCellCount(rect: SourceRect): number {
  return sourceRectCellCountInput(rect)
}

export function sourceRectCellCountInput(rect: unknown): number {
  validateSourceRectInput(rect)
  return (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1)
}
