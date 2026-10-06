import type { AstNode, Token } from './types'
import { tokenize } from './lexer'

class Parser {
  private tokens: Token[]
  private pos = 0

  constructor(tokens: Token[]) {
    this.tokens = tokens
  }

  private peek(): Token {
    return this.tokens[this.pos] || { type: 'eof', value: '', start: 0 }
  }

  private advance(): Token {
    const tok = this.tokens[this.pos]
    if (this.pos < this.tokens.length) {
      this.pos++
    }
    return tok || { type: 'eof', value: '', start: 0 }
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

    const first = this.peek()
    if (first.type === 'error') {
      this.advance()
      return { type: 'error', error: first.value }
    }

    const node = this.parseExpression()

    if (!this.isEof()) {
      const remaining = this.peek()
      if (remaining.type === 'error') {
        return { type: 'error', error: remaining.value }
      }
      return { type: 'error', error: `#NAME? Unexpected token: ${remaining.value}` }
    }

    return node
  }

  // Precedence level 1: Comparisons (=, <>, <, <=, >, >=)
  private parseExpression(): AstNode {
    let left = this.parseConcat()

    while (this.matchOp('=', '<>', '<', '<=', '>', '>=')) {
      const op = this.advance().value
      const right = this.parseConcat()
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 2: String concatenation (&)
  private parseConcat(): AstNode {
    let left = this.parseAdditive()

    while (this.matchOp('&')) {
      const op = this.advance().value
      const right = this.parseAdditive()
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 3: Addition and subtraction (+, -)
  private parseAdditive(): AstNode {
    let left = this.parseMultiplicative()

    while (this.matchOp('+', '-')) {
      const op = this.advance().value
      const right = this.parseMultiplicative()
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 4: Multiplication and division (*, /)
  private parseMultiplicative(): AstNode {
    let left = this.parseExponent()

    while (this.matchOp('*', '/')) {
      const op = this.advance().value
      const right = this.parseExponent()
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 5: Exponentiation (^)
  // Note: in Excel, left-associative, and unary binds tighter than ^ (-2^2 = 4)
  private parseExponent(): AstNode {
    let left = this.parseUnary()

    while (this.matchOp('^')) {
      const op = this.advance().value
      const right = this.parseUnary()
      left = { type: 'binary', op, left, right }
    }

    return left
  }

  // Precedence level 6: Unary prefix (+, -)
  private parseUnary(): AstNode {
    if (this.matchOp('+', '-')) {
      const op = this.advance().value as '+' | '-'
      const expr = this.parseUnary()
      return { type: 'unary', op, expr }
    }

    return this.parsePostfix()
  }

  // Precedence level 7: Postfix percent (%)
  private parsePostfix(): AstNode {
    let expr = this.parsePrimary()

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
      return { type: 'cell', ref: tok.cellRef! }
    }

    if (tok.type === 'range') {
      this.advance()
      return { type: 'range', ref: tok.rangeRef! }
    }

    if (tok.type === 'error') {
      this.advance()
      return { type: 'error', error: tok.value }
    }

    if (tok.type === 'lparen') {
      this.advance() // consume '('
      const expr = this.parseExpression()
      if (this.peek().type === 'rparen') {
        this.advance() // consume ')'
      } else {
        return { type: 'error', error: '#NAME? Missing closing parenthesis' }
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
            } else {
              break
            }
          }
        }

        if (this.peek().type === 'rparen') {
          this.advance()
        } else {
          return { type: 'error', error: `#NAME? Missing closing parenthesis for function ${fnName}` }
        }

        return { type: 'call', name: fnName, args }
      }

      return { type: 'error', error: `#NAME? Unknown identifier: ${fnName}` }
    }

    this.advance()
    return { type: 'error', error: `#NAME? Unexpected token: ${tok.value}` }
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
