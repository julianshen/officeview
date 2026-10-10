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
