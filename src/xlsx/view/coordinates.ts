/** Full zero-based worksheet domain; spec §3 (CAP-01). */
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
  return Number.isSafeInteger(row) && row >= 0 && row < MAX_ROWS
    && Number.isSafeInteger(col) && col >= 0 && col < MAX_COLS
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

export function cellIdToCoordinate(id: CellId): SourceCell {
  if (!Number.isSafeInteger(id) || id < 0 || id > MAX_CELL_ID) {
    throw new CoordinateError('invalid-cell-id', 'Cell ID is outside the worksheet domain')
  }
  return { row: Math.floor(id / MAX_COLS), col: id % MAX_COLS }
}

/** Accept ASCII case and optional absolute markers; emit canonical relative A1.
 * Sheet names, ranges, whitespace and leading-zero row labels are not cell addresses. */
export function parseA1(address: string): SourceCell {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/.exec(address)
  // JS `$` also matches before a final newline, so require the entire input.
  if (!match || match[0] !== address) {
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

/** Column-only labels accept lowercase, but not `$` markers (spec §3). */
export function lettersToColumn(letters: string): SourceCol {
  if (!/^[A-Za-z]{1,3}$/.test(letters) || /[^A-Za-z]/.test(letters)) {
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

export function isValidSourceRect(rect: SourceRect): boolean {
  return isValidCoordinate(rect.r0, rect.c0) && isValidCoordinate(rect.r1, rect.c1)
    && rect.r0 <= rect.r1 && rect.c0 <= rect.c1
}

export function validateSourceRect(rect: SourceRect): void {
  if (!isValidSourceRect(rect)) {
    throw new CoordinateError('invalid-source-rect', 'Invalid inclusive source rectangle')
  }
}

export function sourceRectCellCount(rect: SourceRect): number {
  validateSourceRect(rect)
  return (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1)
}
