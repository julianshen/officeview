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
  | 'eof'

export interface Token {
  type: TokenType
  value: string
  numValue?: number
}
