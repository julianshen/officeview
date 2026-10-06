/**
 * Pure domain types for Excel formula tokenization, AST, and evaluation.
 * Zero DOM, zero Canvas, zero XML dependencies.
 */

export type TokenType =
  | 'number'
  | 'string'
  | 'boolean'
  | 'cell'
  | 'range'
  | 'ident'
  | 'op'
  | 'comma'
  | 'lparen'
  | 'rparen'
  | 'colon'
  | 'error'
  | 'eof'

export interface CellRef {
  sheet?: string
  col: number // 0-based
  row: number // 0-based
  absCol: boolean
  absRow: boolean
}

export interface RangeRef {
  sheet?: string
  from: CellRef
  to: CellRef
}

export interface Token {
  type: TokenType
  value: string
  start: number
  numValue?: number
  sheet?: string
  cellRef?: CellRef
  rangeRef?: RangeRef
}

export type FormulaError =
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#NULL!'

export type FormulaValue = number | string | boolean | null | FormulaError

export interface EvaluationContext {
  getCellValue?(sheet: string | undefined, col: number, row: number): FormulaValue
  /**
   * Evaluates a range of cells into a row-major 2D matrix of values.
   * Contract:
   * - Coordinates are 0-based.
   * - Returns empty/blank cells as null.
   * - Bounded to populated sheet dimensions to avoid unbounded memory allocation.
   * - Must cooperate with cycle detection by delegating or tracking visited cells.
   */
  getRangeValues?(sheet: string | undefined, from: CellRef, to: CellRef): FormulaValue[][]
  currentSheet?: string
  currentCell?: CellRef
  /** In-flight evaluation path for cycle detection. Note: short-circuit-hidden cycles (e.g. IF-skipped branches) are dynamically avoided and not statically traversed. */
  visited?: Set<string>
  evalDepth?: number
  hasCycle?: boolean
}

export type BinaryOp =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '<='
  | '>'
  | '>='

export type AstNode =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'cell'; ref: CellRef }
  | { type: 'range'; ref: RangeRef }
  | { type: 'unary'; op: '+' | '-' | '%'; expr: AstNode }
  | { type: 'binary'; op: BinaryOp; left: AstNode; right: AstNode }
  | { type: 'call'; name: string; args: AstNode[] }
  | { type: 'empty' }
  | { type: 'error'; error: string }


