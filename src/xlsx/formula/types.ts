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

/** Tagged errors are used only during evaluation; public/model values stay primitive. */
export interface EvaluationError { readonly kind: 'formula-error'; readonly code: FormulaError }
export type EvaluationValue = FormulaValue | EvaluationError

export interface EvaluationContext {
  /** Typed callbacks return tagged errors and ordinary strings as text. Legacy callbacks
   * may return canonical error strings, which are interpreted as errors for compatibility.
   * Recursive typed callbacks should use evaluateFormulaInternal to preserve tags. */
  typedValues?: boolean
  /** Per-cell completed AST nodes, retained while an iterative dependency frame resumes. */
  nodeValues?: WeakMap<AstNode, EvaluationValue>
  /** Completed function arguments and next argument, owned by one suspended cell frame. */
  flatArgs?: WeakMap<AstNode[], {
    next: number
    items: Array<{ value: EvaluationValue; fromRef: boolean }>
  }>

  getCellValue?(sheet: string | undefined, col: number, row: number): EvaluationValue
  /**
   * Evaluates a range of cells into a row-major 2D matrix of values.
   * Contract:
   * - Coordinates are 0-based.
   * - Returns empty/blank cells as null.
   * - Bounded to populated sheet dimensions to avoid unbounded memory allocation.
   * - Must cooperate with cycle detection by delegating or tracking visited cells.
   */
  getRangeValues?(sheet: string | undefined, from: CellRef, to: CellRef): EvaluationValue[][]
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


