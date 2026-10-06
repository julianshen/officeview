import type { AstNode, BinaryOp, Token } from './types'
import { tokenize } from './lexer'

class Parser {
  private static readonly MAX_DEPTH = 256
  private tokens: Token[]
  private pos = 0
  private depth = 0

  constructor(tokens: Token[]) {
    this.tokens = tokens
  }

  private peek(): Token {
    return this.tokens[this.pos] || { type: 'eof', value: '', start: this.tokens[this.tokens.length - 1]?.start ?? 0 }
  }

  private advance(): Token {
    const tok = this.tokens[this.pos]
    if (this.pos < this.tokens.length) {
      this.pos++
    }
    return tok || { type: 'eof', value: '', start: this.tokens[this.tokens.length - 1]?.start ?? 0 }
  }

  private isEof(): boolean {
    return this.pos >= this.tokens.length || this.tokens[this.pos].type === 'eof'
  }

  private matchOp(...ops: string[]): boolean {
    const tok = this.peek()
    return tok.type === 'op' && ops.includes(tok.value)
  }

  parse(): AstNode {
    if (this.isEof()) {
      return { type: 'empty' }
    }

    const node = this.parseExpression()
    if (node.type === 'error') {
      return node
    }

    if (!this.isEof()) {
      const remaining = this.peek()
      if (remaining.type === 'error') {
        return { type: 'error', error: remaining.value }
      }
      return {
        type: 'error',
        error: `#NAME? Unexpected token "${remaining.value}" at position ${remaining.start}`,
      }
    }

    return node
  }

  // Precedence level 1: Comparisons (=, <>, <, <=, >, >=)
  private parseExpression(): AstNode {
    if (this.depth >= Parser.MAX_DEPTH) {
      return {
        type: 'error',
        error: `#NAME? Formula exceeds maximum nesting depth of ${Parser.MAX_DEPTH} at position ${this.peek().start}`,
      }
    }

    this.depth++
    try {
      let left = this.parseConcat()
      if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

      while (this.matchOp('=', '<>', '<', '<=', '>', '>=')) {
        const opTok = this.advance()
        const op = opTok.value as BinaryOp
        if (this.isEof()) {
          return {
            type: 'error',
            error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
          }
        }
        const right = this.parseConcat()
        if (right.type === 'empty') {
          return {
            type: 'error',
            error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
          }
        }
        if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
        left = { type: 'binary', op, left, right }
      }

      return left
    } finally {
      this.depth--
    }
  }

  // Precedence level 2: String concatenation (&)
  private parseConcat(): AstNode {
    let left = this.parseAdditive()
    if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

    while (this.matchOp('&')) {
      const opTok = this.advance()
      const op = opTok.value as BinaryOp
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      const right = this.parseAdditive()
      if (right.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 3: Addition and subtraction (+, -)
  private parseAdditive(): AstNode {
    let left = this.parseMultiplicative()
    if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

    while (this.matchOp('+', '-')) {
      const opTok = this.advance()
      const op = opTok.value as BinaryOp
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      const right = this.parseMultiplicative()
      if (right.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 4: Multiplication and division (*, /)
  private parseMultiplicative(): AstNode {
    let left = this.parseExponent()
    if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

    while (this.matchOp('*', '/')) {
      const opTok = this.advance()
      const op = opTok.value as BinaryOp
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      const right = this.parseExponent()
      if (right.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 5: Exponentiation (^)
  // Note: in Excel, left-associative, and unary binds tighter than ^ (-2^2 = 4)
  private parseExponent(): AstNode {
    let left = this.parseUnary()
    if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

    while (this.matchOp('^')) {
      const opTok = this.advance()
      const op = opTok.value as BinaryOp
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      const right = this.parseUnary()
      if (right.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after operator at position ${opTok.start}`,
        }
      }
      if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 6: Unary prefix (+, -)
  private parseUnary(): AstNode {
    if (this.matchOp('+', '-')) {
      const opTok = this.advance()
      const op = opTok.value as '+' | '-'
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after unary operator at position ${opTok.start}`,
        }
      }
      const expr = this.parseUnary()
      if (expr.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after unary operator at position ${opTok.start}`,
        }
      }
      if (expr.type === 'error' && expr.error.startsWith('#NAME?')) return expr
      return { type: 'unary', op, expr }
    }

    return this.parsePostfix()
  }

  // Precedence level 7: Postfix percent (%)
  private parsePostfix(): AstNode {
    let expr = this.parsePrimary()
    if (expr.type === 'error' && expr.error.startsWith('#NAME?')) return expr

    while (this.matchOp('%')) {
      this.advance()
      expr = { type: 'unary', op: '%', expr }
    }

    return expr
  }

  // Precedence level 8: Primary expressions
  private parsePrimary(): AstNode {
    if (this.isEof()) {
      return { type: 'empty' }
    }

    const tok = this.peek()

    if (tok.type === 'number') {
      this.advance()
      return { type: 'number', value: tok.numValue ?? parseFloat(tok.value) }
    }

    if (tok.type === 'string') {
      this.advance()
      return { type: 'string', value: tok.value }
    }

    if (tok.type === 'boolean') {
      this.advance()
      return { type: 'boolean', value: tok.value.toUpperCase() === 'TRUE' }
    }

    if (tok.type === 'cell') {
      this.advance()
      if (!tok.cellRef) {
        return { type: 'error', error: `#REF! Invalid cell reference at position ${tok.start}` }
      }
      return { type: 'cell', ref: tok.cellRef }
    }

    if (tok.type === 'range') {
      this.advance()
      if (!tok.rangeRef) {
        return { type: 'error', error: `#REF! Invalid range reference at position ${tok.start}` }
      }
      return { type: 'range', ref: tok.rangeRef }
    }

    if (tok.type === 'error') {
      this.advance()
      return { type: 'error', error: tok.value }
    }

    if (tok.type === 'lparen') {
      const lpTok = this.advance() // consume '('
      const expr = this.parseExpression()
      if (expr.type === 'error' && expr.error.startsWith('#NAME?')) {
        return expr
      }
      if (this.peek().type === 'rparen') {
        this.advance() // consume ')'
      } else {
        return {
          type: 'error',
          error: `#NAME? Missing closing parenthesis for "(" at position ${lpTok.start}`,
        }
      }
      return expr
    }

    if (tok.type === 'ident') {
      const identTok = this.advance()
      const fnName = identTok.value.toUpperCase()

      if (this.peek().type === 'lparen') {
        this.advance() // consume '('
        const args: AstNode[] = []

        if (this.peek().type !== 'rparen') {
          while (true) {
            if (this.peek().type === 'comma') {
              // empty argument, e.g. IF(A1,,1)
              args.push({ type: 'empty' })
              this.advance()
              if (this.peek().type === 'rparen') {
                args.push({ type: 'empty' })
                break
              }
              continue
            }

            args.push(this.parseExpression())

            if (this.peek().type === 'comma') {
              this.advance()
              if (this.peek().type === 'rparen') {
                // trailing empty arg
                args.push({ type: 'empty' })
                break
              }
            } else if (this.peek().type === 'rparen') {
              break
            } else if (this.isEof()) {
              return {
                type: 'error',
                error: `#NAME? Missing closing parenthesis for function ${fnName} at position ${this.peek().start}`,
              }
            } else {
              const unexpected = this.peek()
              return {
                type: 'error',
                error: `#NAME? Expected comma or closing parenthesis in function ${fnName} at position ${unexpected.start}`,
              }
            }
          }
        }

        if (this.peek().type === 'rparen') {
          this.advance()
        } else {
          return {
            type: 'error',
            error: `#NAME? Missing closing parenthesis for function ${fnName} at position ${this.peek().start}`,
          }
        }

        return { type: 'call', name: fnName, args }
      }

      return {
        type: 'error',
        error: `#NAME? Unknown identifier "${fnName}" at position ${identTok.start}`,
      }
    }

    this.advance()
    return {
      type: 'error',
      error: `#NAME? Unexpected token "${tok.value}" at position ${tok.start}`,
    }
  }
}

/**
 * Parses an Excel formula string or token stream into an AST node.
 */
export function parseFormula(input: string | Token[]): AstNode {
  const tokens = typeof input === 'string' ? tokenize(input) : input
  const parser = new Parser(tokens)
  return parser.parse()
}
