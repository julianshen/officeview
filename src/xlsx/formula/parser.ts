import type { ArrayLiteralElement, AstNode, AxisEndpoint, BinaryOp, FormulaError, GridReference, ReferenceNode, Token } from './types'
import { tokenize } from './lexer'
import { decodeTableSyntax } from './tables'

/** Column letters for 3D run names spelled as cell text (S1, Q4). */
function colLettersForRun(col: number): string {
  let name = ''
  let c = col
  while (c >= 0) {
    name = String.fromCharCode(65 + (c % 26)) + name
    c = Math.floor(c / 26) - 1
  }
  return name
}

class Parser {
  private static readonly MAX_DEPTH = 256
  private tokens: Token[]
  private pos = 0
  private depth = 0
  /** Filtered token indices immediately preceded by whitespace trivia. */
  private spaceBefore = new Set<number>()

  constructor(tokens: Token[]) {
    // Whitespace carries no tokens (callers may still pass legacy 'space'
    // tokens); trivia rides on the following token's spaceBefore flag.
    const filtered: Token[] = []
    for (const tok of tokens) {
      if ((tok.type as string) === 'space') continue
      if (tok.spaceBefore) this.spaceBefore.add(filtered.length)
      filtered.push(tok)
    }
    this.tokens = filtered
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
    const unaries: Array<{ op: '+' | '-'; start: number }> = []

    while (this.matchOp('+', '-')) {
      const opTok = this.advance()
      unaries.push({ op: opTok.value as '+' | '-', start: opTok.start })
      if (this.isEof()) {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after unary operator at position ${opTok.start}`,
        }
      }
    }

    if (unaries.length === 0) {
      return this.parsePostfix()
    }

    let expr = this.parsePostfix()
    if (expr.type === 'empty') {
      const lastOp = unaries[unaries.length - 1]
      return {
        type: 'error',
        error: `#NAME? Unexpected end of expression after unary operator at position ${lastOp.start}`,
      }
    }
    if (expr.type === 'error' && expr.error.startsWith('#NAME?')) return expr

    for (let i = unaries.length - 1; i >= 0; i--) {
      expr = { type: 'unary', op: unaries[i].op, expr }
    }

    return expr
  }

  // Precedence level 7: Postfix percent (%)
  private parsePostfix(): AstNode {
    let expr = this.parseReference()
    if (expr.type === 'error' && expr.error.startsWith('#NAME?')) return expr

    while (this.matchOp('%')) {
      this.advance()
      expr = { type: 'unary', op: '%', expr }
    }

    return expr
  }

  private isReferenceNode(n: AstNode): n is ReferenceNode {
    return (
      n.type === 'cell' ||
      n.type === 'range' ||
      n.type === 'wholeCol' ||
      n.type === 'wholeRow' ||
      n.type === 'name' ||
      n.type === 'table' ||
      n.type === 'ref3d' ||
      n.type === 'union' ||
      n.type === 'intersect' ||
      n.type === 'spill'
    )
  }

private startsReferenceOperand(): boolean {
    const t = this.peek()
    if (
      t.type === 'cell' ||
      t.type === 'range' ||
      t.type === 'wholeCol' ||
      t.type === 'wholeRow' ||
      t.type === 'table' ||
      t.type === 'lparen' ||
      t.type === 'at' ||
      t.type === 'spill'
    ) {
      return true
    }
    if (t.type === 'ident') {
      const n = this.tokens[this.pos + 1]
      return !(n && n.type === 'lparen')
    }
    return false
  }

  /** Quoted 3D sheet runs ('S1:S2'!A1) split into from/to sheets. */
  private splitSheetRun(sheet: string | undefined): { fromSheet: string; toSheet: string } | null {    if (!sheet) return null
    const idx = sheet.indexOf(':')
    if (idx < 0) return null
    return { fromSheet: sheet.slice(0, idx), toSheet: sheet.slice(idx + 1) }
  }

  private gridWithoutSheet(target: GridReference): GridReference {
    switch (target.type) {
      case 'cell':
        return { type: 'cell', ref: { ...target.ref, sheet: undefined } }
      case 'range':
        return {
          type: 'range',
          ref: {
            sheet: undefined,
            from: { ...target.ref.from, sheet: undefined },
            to: { ...target.ref.to, sheet: undefined },
          },
        }
      case 'wholeCol':
        return { type: 'wholeCol', ref: { ...target.ref, sheet: undefined } }
      case 'wholeRow':
        return { type: 'wholeRow', ref: { ...target.ref, sheet: undefined } }
    }
  }

  /**
   * Reference level: 3D sheet runs, spaced intersections and colon ranges.
   * Binds tighter than postfix percent; commas here belong to enclosing
   * parens (unions) or function calls, never to this level.
   */
  private parseReference(): AstNode {
    // 3D run: Sheet1:Sheet2!A1 (grammar-only; sheets never verified here).
    const head = this.peek()
    if (head.type === 'ident' && head.sheet === undefined) {
      const after = this.tokens[this.pos + 1]
      if (after && after.type === 'colon') {
        this.advance()
        this.advance()
        const target = this.parsePostfix()
        if (target.type === 'error' && target.error.startsWith('#NAME?')) return target
        // Group 8: whole-axis 3D targets (wholeCol/wholeRow) join the
        // approved GridReference target set with independent per-endpoint
        // anchors; the run name still comes from the head ident spelling.
        if (target.type === 'wholeCol' || target.type === 'wholeRow') {
          const sheet = (target.ref as { sheet?: string }).sheet
          if (!sheet || sheet.indexOf(':') >= 0) {
            return {
              type: 'error',
              error: `#REF! Invalid 3D reference target at position ${after.start}`,
            }
          }
          return {
            type: 'ref3d',
            sheets: { fromSheet: head.raw ?? head.value, toSheet: sheet },
            target: this.gridWithoutSheet(target) as { type: 'wholeCol' | 'wholeRow'; ref: { from: AxisEndpoint; to: AxisEndpoint; sheet?: string } },
          }
        }
        if (target.type === 'cell' || target.type === 'range') {
          const sheet = target.ref.sheet
          if (!sheet || sheet.indexOf(':') >= 0) {
            return {
              type: 'error',
              error: `#REF! Invalid 3D reference target at position ${after.start}`,
            }
          }
          return {
            type: 'ref3d',
            sheets: { fromSheet: head.raw ?? head.value, toSheet: sheet },
            target: this.gridWithoutSheet(target),
          }
        }
        return {
          type: 'error',
          error: `#NAME? Expected cell or range after 3D sheet run at position ${after.start}`,
        }
      }
    }

    let left = this.parsePrimary()
    if (left.type === 'error' && left.error.startsWith('#NAME?')) return left

    for (;;) {
      const next = this.peek()
      // Spaced intersection of reference operands (at least one range-like;
      // bare cell-cell adjacency keeps the Core parse error below).
      if (
        next.type !== 'eof' &&
        this.spaceBefore.has(this.pos) &&
        this.isReferenceNode(left) &&
        this.startsReferenceOperand()
      ) {
        // Parse tentatively: when the operands cannot combine (bare
        // cell-cell adjacency keeps the Core parse error), rewind so the
        // enclosing context reports at the unconsumed token.
        const savedPos = this.pos
        const right = this.parsePostfix()
        if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
        // Group 13 (root-authorized): legitimate single-cell reference
        // operands form intersections too — any reference node combines.
        if (!this.isReferenceNode(right)) {
          this.pos = savedPos
          const at = this.peek()
          return {
            type: 'error',
            error: `#NAME? Unexpected token "${at.value}" at position ${at.start}`,
          }
        }
        left = { type: 'intersect', left, right }
        continue
      }
      // Colon continuation: spaced ranges (A1 :B2) and out-of-range whole
      // rows (0:5). Lexer-greedy forms (A1:B2, 1:3, Sheet1:Sheet2!A1) never
      // reach here as separate tokens.
      if (next.type === 'colon') {
        const colon = this.advance()
        const right = this.parsePostfix()
        if (right.type === 'error' && right.error.startsWith('#NAME?')) return right
        // Cell-qualified 3D run (S1:S2!A1): a sheet-less left cell with a
        // sheet-qualified right side is a sheet run, never a cross-sheet
        // range (which Excel forbids). The run name reuses the source cell
        // spelling; unknown sheets fail closed at resolution.
        if (
          left.type === 'cell' &&
          left.ref.sheet === undefined &&
          (right.type === 'cell' || right.type === 'range' || right.type === 'wholeCol' || right.type === 'wholeRow') &&
          right.ref.sheet !== undefined &&
          right.ref.sheet.indexOf(':') < 0
        ) {
          return {
            type: 'ref3d',
            sheets: {
              fromSheet: `${colLettersForRun(left.ref.col)}${left.ref.row + 1}`,
              toSheet: right.ref.sheet,
            },
            target: this.gridWithoutSheet(right),
          }
        }
        if (left.type === 'cell' && right.type === 'cell') {
          const ls = left.ref.sheet
          const rs = right.ref.sheet
          if (ls !== undefined && rs !== undefined && ls.toLowerCase() !== rs.toLowerCase()) {
            return {
              type: 'error',
              error: `#REF! Mismatched sheets in range at position ${colon.start}`,
            }
          }
          const sheet = ls ?? rs
          left = {
            type: 'range',
            ref: {
              ...(sheet !== undefined ? { sheet } : {}),
              from: { ...left.ref, sheet },
              to: { ...right.ref, sheet },
            },
          }
          continue
        }
        if (left.type === 'number' || right.type === 'number') {
          return {
            type: 'error',
            error: `#REF! Invalid reference coordinates at position ${colon.start}`,
          }
        }
        return {
          type: 'error',
          error: `#NAME? Unexpected token "${next.value}" at position ${next.start}`,
        }
      }
      return left
    }
  }

  /** Canonical error literals are the only error atoms a literal may hold. */
  private static readonly LITERAL_ERRORS = new Set<string>([
    '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!',
  ])

  private parseArrayElement(): ArrayLiteralElement | null {
    const tok = this.peek()
    // MS-XLSX numerical-constant (45) permits an optional negative sign, so a
    // signed numeric literal is a legal array atom (65..69). A sign not
    // followed by a number is not an atom (no arbitrary literal formulas).
    if (tok.type === 'op' && tok.value === '-') {
      const next = this.tokens[this.pos + 1]
      if (next && next.type === 'number') {
        this.advance()
        const num = this.advance()
        return -(num.numValue ?? parseFloat(num.value))
      }
      return null
    }
    if (tok.type === 'number') {
      this.advance()
      return tok.numValue ?? parseFloat(tok.value)
    }
    if (tok.type === 'string') {
      this.advance()
      return tok.value
    }
    if (tok.type === 'boolean') {
      this.advance()
      return tok.value.toUpperCase() === 'TRUE'
    }
    if (tok.type === 'error' && Parser.LITERAL_ERRORS.has(tok.value)) {
      this.advance()
      return { kind: 'formula-error', code: tok.value as FormulaError }
    }
    return null
  }

  /**
   * Rectangular array literal `{1,2;3,4}`. Grammar only: rows are comma-
   * separated elements and semicolons separate rows. Ragged, nested, omitted
   * and empty literals are rejected without any invented zero fill.
   */
  private parseArrayConst(): AstNode {
    const open = this.advance()
    const rows: ArrayLiteralElement[][] = []
    for (;;) {
      if (this.peek().type === 'rbrace') {
        return { type: 'error', error: `#VALUE! Empty array literal at position ${open.start}` }
      }
      const row: ArrayLiteralElement[] = []
      for (;;) {
        const element = this.parseArrayElement()
        if (element === null) {
          const at = this.peek()
          return { type: 'error', error: `#NAME? Invalid array literal element "${at.value}" at position ${at.start}` }
        }
        row.push(element)
        if (this.peek().type === 'comma') { this.advance(); continue }
        break
      }
      rows.push(row)
      if (this.peek().type === 'semicolon') { this.advance(); continue }
      break
    }
    if (this.peek().type !== 'rbrace') {
      return { type: 'error', error: `#NAME? Unclosed array literal at position ${open.start}` }
    }
    this.advance()
    const width = rows[0].length
    if (rows.some((row) => row.length !== width)) {
      return { type: 'error', error: `#VALUE! Ragged array literal at position ${open.start}` }
    }
    return { type: 'arrayConst', rows }
  }

  // Precedence level 8: Primary expressions
  private parsePrimary(): AstNode {
    if (this.isEof()) {
      return { type: 'empty' }
    }

    const tok = this.peek()

    if (tok.type === 'lbrace') {
      return this.parseArrayConst()
    }

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
      if (this.peek().type === 'lparen') {
        this.advance() // consume '('
        if (this.peek().type === 'rparen') {
          this.advance() // consume ')'
        } else {
          return {
            type: 'error',
            error: `#NAME? Expected 0 arguments for ${tok.value}() at position ${this.peek().start}`,
          }
        }
      }
      return { type: 'boolean', value: tok.value.toUpperCase() === 'TRUE' }
    }

    if (tok.type === 'cell') {
      this.advance()
      if (!tok.cellRef) {
        return { type: 'error', error: `#REF! Invalid cell reference at position ${tok.start}` }
      }
      const run = this.splitSheetRun(tok.cellRef.sheet)
      if (run) {
        return {
          type: 'ref3d',
          sheets: run,
          target: { type: 'cell', ref: { ...tok.cellRef, sheet: undefined } },
        }
      }
      return { type: 'cell', ref: tok.cellRef }
    }

    if (tok.type === 'range') {
      this.advance()
      if (!tok.rangeRef) {
        return { type: 'error', error: `#REF! Invalid range reference at position ${tok.start}` }
      }
      const run = this.splitSheetRun(tok.rangeRef.sheet)
      if (run) {
        return {
          type: 'ref3d',
          sheets: run,
          target: {
            type: 'range',
            ref: {
              sheet: undefined,
              from: { ...tok.rangeRef.from, sheet: undefined },
              to: { ...tok.rangeRef.to, sheet: undefined },
            },
          },
        }
      }
      return { type: 'range', ref: tok.rangeRef }
    }

    if (tok.type === 'wholeCol') {
      this.advance()
      if (!tok.wholeColRef) {
        return { type: 'error', error: `#REF! Invalid column reference at position ${tok.start}` }
      }
      const run = this.splitSheetRun(tok.wholeColRef.sheet)
      if (run) {
        return {
          type: 'ref3d',
          sheets: run,
          target: { type: 'wholeCol', ref: { ...tok.wholeColRef, sheet: undefined } },
        }
      }
      return { type: 'wholeCol', ref: tok.wholeColRef }
    }

    if (tok.type === 'wholeRow') {
      this.advance()
      if (!tok.wholeRowRef) {
        return { type: 'error', error: `#REF! Invalid row reference at position ${tok.start}` }
      }
      const run = this.splitSheetRun(tok.wholeRowRef.sheet)
      if (run) {
        return {
          type: 'ref3d',
          sheets: run,
          target: { type: 'wholeRow', ref: { ...tok.wholeRowRef, sheet: undefined } },
        }
      }
      return { type: 'wholeRow', ref: tok.wholeRowRef }
    }

    if (tok.type === 'table') {
      this.advance()
      const decoded = decodeTableSyntax(tok.tableName, tok.value)
      if (decoded.kind === 'error') {
        return { type: 'error', error: decoded.message }
      }
      if (decoded.kind === 'union') {
        return { type: 'union', refs: decoded.items.map((syntax) => ({ type: 'table' as const, ref: syntax })) }
      }
      return { type: 'table', ref: decoded.syntax }
    }

    if (tok.type === 'spill') {
      this.advance()
      if (tok.cellRef) {
        return { type: 'spill', anchor: tok.cellRef }
      }
      return { type: 'error', error: `#REF! Spill anchor must be a single cell at position ${tok.start}` }
    }

    if (tok.type === 'at') {
      this.advance()
      if (this.peek().type === 'at') {
        return {
          type: 'error',
          error: `#NAME? Unexpected token "@" at position ${this.peek().start}`,
        }
      }
      const operand = this.parsePrimary()
      if (operand.type === 'error' && operand.error.startsWith('#NAME?')) return operand
      if (operand.type === 'empty') {
        return {
          type: 'error',
          error: `#NAME? Unexpected end of expression after "@" at position ${tok.start}`,
        }
      }
      return { type: 'implicitIntersect', expr: operand }
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
      // Parenthesized unions: (B2,B4) is an ordered union of occurrences.
      // Commas inside function calls remain argument separators (handled by
      // the call branch); only grouping parens form unions here.
      if (this.peek().type === 'comma') {
        const refs: ReferenceNode[] = []
        if (!this.isReferenceNode(expr)) {
          return {
            type: 'error',
            error: `#NAME? Unexpected token "," at position ${this.peek().start}`,
          }
        }
        refs.push(expr)
        while (this.peek().type === 'comma') {
          this.advance()
          const next = this.parseExpression()
          if (next.type === 'error' && next.error.startsWith('#NAME?')) {
            return next
          }
          if (!this.isReferenceNode(next)) {
            return {
              type: 'error',
              error: `#NAME? Unexpected token "," at position ${lpTok.start}`,
            }
          }
          refs.push(next)
        }
        if (this.peek().type === 'rparen') {
          this.advance() // consume ')'
        } else {
          return {
            type: 'error',
            error: `#NAME? Missing closing parenthesis for "(" at position ${lpTok.start}`,
          }
        }
        return { type: 'union', refs }
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

      // Sheet-qualified defined name: Data!MyName (qualification only).
      if (identTok.sheet !== undefined) {
        return { type: 'name', ref: { name: identTok.raw ?? identTok.value, sheet: identTok.sheet } }
      }

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

      // Bare identifiers are grammar-level names (B1 ABI): unknown names
      // evaluate to #NAME? only on the evaluated path, never at parse time.
      // Original source spelling is preserved; matching is case-insensitive.
      return {
        type: 'name',
        ref: { name: identTok.raw ?? identTok.value },
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
