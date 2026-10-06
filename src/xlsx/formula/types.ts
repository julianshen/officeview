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
  col: number // 0-based
  row: number // 0-based
  absCol: boolean
  absRow: boolean
}

export interface Token {
  type: TokenType
  value: string
  numValue?: number
  cellRef?: CellRef
}
