import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { tokenize } from '../src/xlsx/formula/lexer'
import { parseFormula } from '../src/xlsx/formula/parser'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import { translateSharedFormula } from '../src/xlsx/formula/shared'
import { evaluateWorkbookFormulas } from '../src/xlsx/formula/workbook'
import { renderSheet } from '../src/xlsx/render'
import type { EvaluationContext, FormulaValue } from '../src/xlsx/formula/types'
import type { XlsxCell, XlsxDocument } from '../src/xlsx/types'

describe('xlsx formula fixture enablement', () => {
  test('XlsxCellSpec and buildXlsx support formula <f> and omitting <v>', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', v: 42 },
              { ref: 'B1', formula: 'A1+1' },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]
    expect(sheet.rows[0].cells).toHaveLength(2)

    const b1 = sheet.rows[0].cells[1]
    expect(b1.ref).toBe('B1')
    expect(b1.formula).toBe('A1+1')
    expect(b1.value).toBe(43)
  })

  test('XlsxCellSpec escapes XML entities in formula text (< and &)', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', formula: 'IF(B1<5,1,0)&"x"' },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]
    expect(sheet.rows[0].cells[0].formula).toBe('IF(B1<5,1,0)&"x"')
  })

  test('XlsxCellSpec supports shared formula attributes (t="shared", si, ref)', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', v: 10 },
              { ref: 'B1', formula: 'A1*2', sharedFormula: { si: 0, ref: 'B1:B2' } },
            ],
          },
          {
            r: 2,
            cells: [
              { ref: 'A2', v: 20 },
              { ref: 'B2', sharedFormula: { si: 0 } },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]

    const b1 = sheet.rows[0].cells[1]
    expect(b1.formula).toBe('A1*2')
    expect(b1.sharedFormula).toEqual({ si: 0, ref: 'B1:B2' })

    const b2 = sheet.rows[1].cells[1]
    expect(b2.formula).toBe('A2*2')
    expect(b2.value).toBe(40)
    expect(b2.sharedFormula).toEqual({ si: 0, ref: undefined })
  })

  test('XlsxCellSpec supports error cells (t="e")', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', t: 'e', v: '#DIV/0!' },
              { ref: 'B1', t: 'e', v: '#VALUE!' },
              { ref: 'C1', t: 'e' },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]

    expect(sheet.rows[0].cells[0].value).toBe('#DIV/0!')
    expect(sheet.rows[0].cells[1].value).toBe('#VALUE!')
    expect(sheet.rows[0].cells[2].value).toBeNull()
  })

  test('XlsxCellSpec escapes XML entities in <v> text (< and &)', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', t: 'str', v: 'alpha&beta<gamma' },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].rows[0].cells[0].value).toBe('alpha&beta<gamma')
  })
})

describe('formula lexer', () => {
  test('Tokenizer handles arithmetic operators, unary minus, and percent (+, -, *, /, ^, %)', () => {
    const tokens = tokenize('=10 + 20 * -3 / 4 ^ 2%')
    expect(tokens.map(t => ({ type: t.type, value: t.value }))).toEqual([
      { type: 'number', value: '10' },
      { type: 'op', value: '+' },
      { type: 'number', value: '20' },
      { type: 'op', value: '*' },
      { type: 'op', value: '-' },
      { type: 'number', value: '3' },
      { type: 'op', value: '/' },
      { type: 'number', value: '4' },
      { type: 'op', value: '^' },
      { type: 'number', value: '2' },
      { type: 'op', value: '%' },
      { type: 'eof', value: '' },
    ])
    expect(tokens[0].numValue).toBe(10)
  })

  test('Tokenizer handles string literals with escaped quotes and comparison operators (=, <>, <, <=, >, >=, &)', () => {
    const tokens = tokenize('="Hello ""World""" & "!" = "foo" <> "bar" <= 10 >= 5 < 20 > 1')
    expect(tokens.map(t => ({ type: t.type, value: t.value }))).toEqual([
      { type: 'string', value: 'Hello "World"' },
      { type: 'op', value: '&' },
      { type: 'string', value: '!' },
      { type: 'op', value: '=' },
      { type: 'string', value: 'foo' },
      { type: 'op', value: '<>' },
      { type: 'string', value: 'bar' },
      { type: 'op', value: '<=' },
      { type: 'number', value: '10' },
      { type: 'op', value: '>=' },
      { type: 'number', value: '5' },
      { type: 'op', value: '<' },
      { type: 'number', value: '20' },
      { type: 'op', value: '>' },
      { type: 'number', value: '1' },
      { type: 'eof', value: '' },
    ])
  })

  test('Tokenizer handles cell references (relative A1, absolute $A$1, mixed A$1, $A1)', () => {
    const tokens = tokenize('=A1 + $B$2 * C$3 / $D4 + aa10')
    expect(tokens.map(t => ({ type: t.type, value: t.value, cellRef: t.cellRef }))).toEqual([
      { type: 'cell', value: 'A1', cellRef: { col: 0, row: 0, absCol: false, absRow: false } },
      { type: 'op', value: '+', cellRef: undefined },
      { type: 'cell', value: '$B$2', cellRef: { col: 1, row: 1, absCol: true, absRow: true } },
      { type: 'op', value: '*', cellRef: undefined },
      { type: 'cell', value: 'C$3', cellRef: { col: 2, row: 2, absCol: false, absRow: true } },
      { type: 'op', value: '/', cellRef: undefined },
      { type: 'cell', value: '$D4', cellRef: { col: 3, row: 3, absCol: true, absRow: false } },
      { type: 'op', value: '+', cellRef: undefined },
      { type: 'cell', value: 'aa10', cellRef: { col: 26, row: 9, absCol: false, absRow: false } },
      { type: 'eof', value: '', cellRef: undefined },
    ])
  })

  test('Tokenizer distinguishes function calls ending in digits (LOG10) from cell references and handles error literals / fail-closed errors', () => {
    // 1. Disambiguation: LOG10( is ident followed by lparen, not cell LOG10
    const logTokens = tokenize('=LOG10(100)')
    expect(logTokens[0]).toMatchObject({ type: 'ident', value: 'LOG10' })
    expect(logTokens[1]).toMatchObject({ type: 'lparen', value: '(' })

    // 2. Error literals: #DIV/0!, #REF!, #N/A, #VALUE!, #NAME?, #NUM!, #NULL!
    const errTokens = tokenize('=#DIV/0! + #REF! + #N/A')
    expect(errTokens.map(t => ({ type: t.type, value: t.value }))).toEqual([
      { type: 'error', value: '#DIV/0!' },
      { type: 'op', value: '+' },
      { type: 'error', value: '#REF!' },
      { type: 'op', value: '+' },
      { type: 'error', value: '#N/A' },
      { type: 'eof', value: '' },
    ])

    // 3. Unterminated string fails closed with error token
    const unterminated = tokenize('="hello world')
    expect(unterminated.some(t => t.type === 'error')).toBe(true)

    // 4. Unknown character emits error token instead of op
    const unknown = tokenize('=@;')
    expect(unknown[0]).toMatchObject({ type: 'error', value: '@' })

    // 5. Overlong formula emits error token
    const overlong = tokenize('=' + 'A'.repeat(8200))
    expect(overlong[0]).toMatchObject({ type: 'error', value: 'Formula exceeds maximum length of 8192 characters' })
  })

  test('Tokenizer handles range references (A1:B10) and cross-sheet references (Sheet2!A1, \'My Sheet\'!A1:B2)', () => {
    // 1. Simple unquoted range
    const r1 = tokenize('=A1:B10')
    expect(r1).toMatchObject([
      {
        type: 'range',
        value: 'A1:B10',
        rangeRef: {
          from: { col: 0, row: 0, absCol: false, absRow: false },
          to: { col: 1, row: 9, absCol: false, absRow: false },
        },
      },
      { type: 'eof', value: '' },
    ])

    // 2. Absolute range
    const r2 = tokenize('=$A$1:$C$5')
    expect(r2).toMatchObject([
      {
        type: 'range',
        value: '$A$1:$C$5',
        rangeRef: {
          from: { col: 0, row: 0, absCol: true, absRow: true },
          to: { col: 2, row: 4, absCol: true, absRow: true },
        },
      },
      { type: 'eof', value: '' },
    ])

    // 3. Unquoted cross-sheet cell reference
    const c1 = tokenize('=Sheet2!A1')
    expect(c1).toMatchObject([
      {
        type: 'cell',
        value: 'Sheet2!A1',
        sheet: 'Sheet2',
        cellRef: {
          sheet: 'Sheet2',
          col: 0,
          row: 0,
          absCol: false,
          absRow: false,
        },
      },
      { type: 'eof', value: '' },
    ])

    // 4. Quoted cross-sheet range reference
    const r3 = tokenize("='My Sheet'!A1:B2")
    expect(r3).toMatchObject([
      {
        type: 'range',
        value: "'My Sheet'!A1:B2",
        sheet: 'My Sheet',
        rangeRef: {
          sheet: 'My Sheet',
          from: { sheet: 'My Sheet', col: 0, row: 0, absCol: false, absRow: false },
          to: { sheet: 'My Sheet', col: 1, row: 1, absCol: false, absRow: false },
        },
      },
      { type: 'eof', value: '' },
    ])

    // 5. Escaped quote in cross-sheet cell reference
    const c2 = tokenize("='Bob''s Data'!$D$10")
    expect(c2).toMatchObject([
      {
        type: 'cell',
        value: "'Bob''s Data'!$D$10",
        sheet: "Bob's Data",
        cellRef: {
          sheet: "Bob's Data",
          col: 3,
          row: 9,
          absCol: true,
          absRow: true,
        },
      },
      { type: 'eof', value: '' },
    ])
  })

  test('Tokenizer strips _xlfn. function prefix and normalizes function names case-insensitively', () => {
    const t1 = tokenize('=_xlfn.CONCAT(A1, "test")')
    expect(t1[0]).toMatchObject({ type: 'ident', value: 'CONCAT' })

    const t2 = tokenize('=_xlfn.concat(A1, B1)')
    expect(t2[0]).toMatchObject({ type: 'ident', value: 'CONCAT' })

    const t3 = tokenize('=_xlfn.STDEV.S(A1:B10)')
    expect(t3[0]).toMatchObject({ type: 'ident', value: 'STDEV.S' })

    const t4 = tokenize('=_xlfn._xlws.FILTER(A1:B10, A1:A10>0)')
    expect(t4[0]).toMatchObject({ type: 'ident', value: 'FILTER' })

    const t5 = tokenize('=sum(A1:B10) + Average(C1:C10)')
    expect(t5[0]).toMatchObject({ type: 'ident', value: 'SUM' })
    expect(t5[5]).toMatchObject({ type: 'ident', value: 'AVERAGE' })
  })

  test('Lexer tracks token start offsets, preserves authored cell case, and supports CJK sheet names', () => {
    // 1. Offsets
    const tOffsets = tokenize('=A1 + 20')
    expect(tOffsets.map(t => ({ type: t.type, value: t.value, start: t.start }))).toEqual([
      { type: 'cell', value: 'A1', start: 0 },
      { type: 'op', value: '+', start: 3 },
      { type: 'number', value: '20', start: 5 },
      { type: 'eof', value: '', start: 7 },
    ])

    // 2. Authored cell case preserved
    const tCase = tokenize('=aa10 + $b$2')
    expect(tCase[0].value).toBe('aa10')
    expect(tCase[0].cellRef).toEqual({ col: 26, row: 9, absCol: false, absRow: false })
    expect(tCase[2].value).toBe('$b$2')
    expect(tCase[2].cellRef).toEqual({ col: 1, row: 1, absCol: true, absRow: true })

    // 3. CJK unquoted and quoted sheet names
    const tCjk = tokenize('=工作表1!A1 + 销售!B2:C10')
    expect(tCjk[0]).toMatchObject({
      type: 'cell',
      value: '工作表1!A1',
      sheet: '工作表1',
      cellRef: { sheet: '工作表1', col: 0, row: 0 },
    })
    expect(tCjk[2]).toMatchObject({
      type: 'range',
      value: '销售!B2:C10',
      sheet: '销售',
      rangeRef: { sheet: '销售', from: { col: 1, row: 1 }, to: { col: 2, row: 9 } },
    })
  })
})

describe('xlsx formula parser (AST)', () => {
  test('Parser parses literals and respects Excel unary precedence (-2^2 parses as (-2)^2)', () => {
    // 1. Numbers, strings, booleans
    expect(parseFormula('42')).toEqual({ type: 'number', value: 42 })
    expect(parseFormula('3.14159')).toEqual({ type: 'number', value: 3.14159 })
    expect(parseFormula('"hello world"')).toEqual({ type: 'string', value: 'hello world' })
    expect(parseFormula('TRUE')).toEqual({ type: 'boolean', value: true })
    expect(parseFormula('FALSE')).toEqual({ type: 'boolean', value: false })

    // 2. Cell and range references
    expect(parseFormula('A1')).toEqual({
      type: 'cell',
      ref: { col: 0, row: 0, absCol: false, absRow: false },
    })
    expect(parseFormula('A1:B10')).toEqual({
      type: 'range',
      ref: {
        from: { col: 0, row: 0, absCol: false, absRow: false },
        to: { col: 1, row: 9, absCol: false, absRow: false },
      },
    })

    // 3. Unary minus and plus
    expect(parseFormula('-5')).toEqual({
      type: 'unary',
      op: '-',
      expr: { type: 'number', value: 5 },
    })
    expect(parseFormula('+5')).toEqual({
      type: 'unary',
      op: '+',
      expr: { type: 'number', value: 5 },
    })

    // 4. Postfix percent
    expect(parseFormula('50%')).toEqual({
      type: 'unary',
      op: '%',
      expr: { type: 'number', value: 50 },
    })

    // 5. Critical Excel Invariant: Unary minus binds tighter than exponentiation (-2^2 parses as (-2)^2)
    expect(parseFormula('=-2^2')).toEqual({
      type: 'binary',
      op: '^',
      left: {
        type: 'unary',
        op: '-',
        expr: { type: 'number', value: 2 },
      },
      right: {
        type: 'number',
        value: 2,
      },
    })

    // 6. Explicit parentheses override: -(2^2)
    expect(parseFormula('=-(2^2)')).toEqual({
      type: 'unary',
      op: '-',
      expr: {
        type: 'binary',
        op: '^',
        left: { type: 'number', value: 2 },
        right: { type: 'number', value: 2 },
      },
    })

    // 7. Right-hand unary minus in exponent: 2^-2 parses as 2 ^ (-2)
    expect(parseFormula('=2^-2')).toEqual({
      type: 'binary',
      op: '^',
      left: { type: 'number', value: 2 },
      right: {
        type: 'unary',
        op: '-',
        expr: { type: 'number', value: 2 },
      },
    })
  })

  test('Parser respects operator precedence (^ > *, / > +, - > & > comparisons)', () => {
    // 1. ^ over * and /
    expect(parseFormula('2 * 3 ^ 2')).toEqual({
      type: 'binary',
      op: '*',
      left: { type: 'number', value: 2 },
      right: {
        type: 'binary',
        op: '^',
        left: { type: 'number', value: 3 },
        right: { type: 'number', value: 2 },
      },
    })
    expect(parseFormula('16 / 2 ^ 3')).toEqual({
      type: 'binary',
      op: '/',
      left: { type: 'number', value: 16 },
      right: {
        type: 'binary',
        op: '^',
        left: { type: 'number', value: 2 },
        right: { type: 'number', value: 3 },
      },
    })

    // 2. * and / over + and -
    expect(parseFormula('1 + 2 * 3')).toEqual({
      type: 'binary',
      op: '+',
      left: { type: 'number', value: 1 },
      right: {
        type: 'binary',
        op: '*',
        left: { type: 'number', value: 2 },
        right: { type: 'number', value: 3 },
      },
    })
    expect(parseFormula('10 - 6 / 2')).toEqual({
      type: 'binary',
      op: '-',
      left: { type: 'number', value: 10 },
      right: {
        type: 'binary',
        op: '/',
        left: { type: 'number', value: 6 },
        right: { type: 'number', value: 2 },
      },
    })

    // 3. + and - over &
    expect(parseFormula('"A" & 1 + 2')).toEqual({
      type: 'binary',
      op: '&',
      left: { type: 'string', value: 'A' },
      right: {
        type: 'binary',
        op: '+',
        left: { type: 'number', value: 1 },
        right: { type: 'number', value: 2 },
      },
    })

    // 4. & over comparisons
    expect(parseFormula('"a" & "b" = "ab"')).toEqual({
      type: 'binary',
      op: '=',
      left: {
        type: 'binary',
        op: '&',
        left: { type: 'string', value: 'a' },
        right: { type: 'string', value: 'b' },
      },
      right: { type: 'string', value: 'ab' },
    })
    expect(parseFormula('A1 + B1 <> C1 & D1')).toEqual({
      type: 'binary',
      op: '<>',
      left: {
        type: 'binary',
        op: '+',
        left: { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
        right: { type: 'cell', ref: { col: 1, row: 0, absCol: false, absRow: false } },
      },
      right: {
        type: 'binary',
        op: '&',
        left: { type: 'cell', ref: { col: 2, row: 0, absCol: false, absRow: false } },
        right: { type: 'cell', ref: { col: 3, row: 0, absCol: false, absRow: false } },
      },
    })

    // 5. Left-associativity of same precedence
    expect(parseFormula('10 - 5 - 2')).toEqual({
      type: 'binary',
      op: '-',
      left: {
        type: 'binary',
        op: '-',
        left: { type: 'number', value: 10 },
        right: { type: 'number', value: 5 },
      },
      right: { type: 'number', value: 2 },
    })
  })

  test('Parser parses function calls with multiple arguments and empty arguments', () => {
    // 1. Zero arguments
    expect(parseFormula('NOW()')).toEqual({
      type: 'call',
      name: 'NOW',
      args: [],
    })

    // 2. Single argument
    expect(parseFormula('ABS(-10)')).toEqual({
      type: 'call',
      name: 'ABS',
      args: [
        {
          type: 'unary',
          op: '-',
          expr: { type: 'number', value: 10 },
        },
      ],
    })

    // 3. Multiple arguments
    expect(parseFormula('ROUND(3.14159, 2)')).toEqual({
      type: 'call',
      name: 'ROUND',
      args: [
        { type: 'number', value: 3.14159 },
        { type: 'number', value: 2 },
      ],
    })

    // 4. Empty arguments: middle, leading, trailing, and consecutive
    expect(parseFormula('IF(A1,,1)')).toEqual({
      type: 'call',
      name: 'IF',
      args: [
        { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
        { type: 'empty' },
        { type: 'number', value: 1 },
      ],
    })

    expect(parseFormula('IF(, 1, 2)')).toEqual({
      type: 'call',
      name: 'IF',
      args: [
        { type: 'empty' },
        { type: 'number', value: 1 },
        { type: 'number', value: 2 },
      ],
    })

    expect(parseFormula('IF(A1, 1, )')).toEqual({
      type: 'call',
      name: 'IF',
      args: [
        { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
        { type: 'number', value: 1 },
        { type: 'empty' },
      ],
    })

    expect(parseFormula('TEST(,,)')).toEqual({
      type: 'call',
      name: 'TEST',
      args: [
        { type: 'empty' },
        { type: 'empty' },
        { type: 'empty' },
      ],
    })

    // 5. Nested calls
    expect(parseFormula('SUM(A1, MAX(B1, C1))')).toEqual({
      type: 'call',
      name: 'SUM',
      args: [
        { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
        {
          type: 'call',
          name: 'MAX',
          args: [
            { type: 'cell', ref: { col: 1, row: 0, absCol: false, absRow: false } },
            { type: 'cell', ref: { col: 2, row: 0, absCol: false, absRow: false } },
          ],
        },
      ],
    })
  })

  test('Parser parses nested expressions and parenthesized sub-expressions', () => {
    // 1. Parenthesized arithmetic grouping overriding natural precedence
    expect(parseFormula('(1 + 2) * (3 + 4)')).toEqual({
      type: 'binary',
      op: '*',
      left: {
        type: 'binary',
        op: '+',
        left: { type: 'number', value: 1 },
        right: { type: 'number', value: 2 },
      },
      right: {
        type: 'binary',
        op: '+',
        left: { type: 'number', value: 3 },
        right: { type: 'number', value: 4 },
      },
    })

    // 2. Deep nesting
    expect(parseFormula('(((A1)))')).toEqual({
      type: 'cell',
      ref: { col: 0, row: 0, absCol: false, absRow: false },
    })

    // 3. Nested double unaries
    expect(parseFormula('-(-(A1 + 1))')).toEqual({
      type: 'unary',
      op: '-',
      expr: {
        type: 'unary',
        op: '-',
        expr: {
          type: 'binary',
          op: '+',
          left: { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
          right: { type: 'number', value: 1 },
        },
      },
    })

    // 4. Function containing expressions and parenthesized subexpressions
    expect(parseFormula('IF(A1 + B1 > 10, (C1 - D1) * 2, 0)')).toEqual({
      type: 'call',
      name: 'IF',
      args: [
        {
          type: 'binary',
          op: '>',
          left: {
            type: 'binary',
            op: '+',
            left: { type: 'cell', ref: { col: 0, row: 0, absCol: false, absRow: false } },
            right: { type: 'cell', ref: { col: 1, row: 0, absCol: false, absRow: false } },
          },
          right: { type: 'number', value: 10 },
        },
        {
          type: 'binary',
          op: '*',
          left: {
            type: 'binary',
            op: '-',
            left: { type: 'cell', ref: { col: 2, row: 0, absCol: false, absRow: false } },
            right: { type: 'cell', ref: { col: 3, row: 0, absCol: false, absRow: false } },
          },
          right: { type: 'number', value: 2 },
        },
        { type: 'number', value: 0 },
      ],
    })
  })

  test('Parser returns #NAME? or syntax error node on malformed input without throwing', () => {
    // 1. Literal error constants (#DIV/0!, #REF!, #VALUE!, #NAME?, #NUM!, #N/A, #NULL!)
    expect(parseFormula('#DIV/0!')).toEqual({ type: 'error', error: '#DIV/0!' })
    expect(parseFormula('=#REF!')).toEqual({ type: 'error', error: '#REF!' })
    expect(parseFormula('=#VALUE!')).toEqual({ type: 'error', error: '#VALUE!' })

    // 2. Empty formulas
    expect(parseFormula('')).toEqual({ type: 'empty' })
    expect(parseFormula('=')).toEqual({ type: 'empty' })

    // 3. Unknown identifiers without parentheses
    const errIdent = parseFormula('=UNKNOWN')
    expect(errIdent.type).toBe('error')
    expect((errIdent as any).error).toMatch(/#NAME\?/)

    // 4. Unclosed parentheses
    const errParen = parseFormula('=(1 + 2')
    expect(errParen.type).toBe('error')
    expect((errParen as any).error).toMatch(/#NAME\?/)

    const errFnParen = parseFormula('=SUM(A1, B1')
    expect(errFnParen.type).toBe('error')
    expect((errFnParen as any).error).toMatch(/#NAME\?/)

    // 5. Trailing junk / unexpected token
    const errJunk = parseFormula('=1 + 2 3')
    expect(errJunk.type).toBe('error')
    expect((errJunk as any).error).toMatch(/#NAME\?/)

    // 6. Incomplete operators
    const errOp = parseFormula('=1 + ')
    expect(errOp.type).toBe('error')
    expect((errOp as any).error).toMatch(/#NAME\?/)

    const errUnary = parseFormula('=+')
    expect(errUnary.type).toBe('error')
    expect((errUnary as any).error).toMatch(/#NAME\?/)

    // 7. Malformed inputs that must never throw unhandled exceptions
    expect(() => parseFormula(')(*&^%$#@!')).not.toThrow()
    expect(() => parseFormula('=,,')).not.toThrow()
    expect(() => parseFormula('====')).not.toThrow()
  })

  test('Parser addresses peer review hardening (recursion depth limit, error offsets, token validation, error operand preservation)', () => {
    // 1. Recursion depth guard against stack overflow (RangeError)
    const deepParens4000 = '=' + '('.repeat(4000) + '1' + ')'.repeat(4000)
    let deepErr: any
    expect(() => {
      deepErr = parseFormula(deepParens4000)
    }).not.toThrow()
    expect(deepErr.type).toBe('error')
    expect(deepErr.error).toMatch(/#NAME\? Formula exceeds maximum nesting depth/)

    // Deep legal nesting (50 levels) parses cleanly
    const legalParens50 = '=' + '('.repeat(50) + '42' + ')'.repeat(50)
    expect(parseFormula(legalParens50)).toEqual({ type: 'number', value: 42 })

    // 2. Token start offsets wired into error messages
    const errUnknown = parseFormula('=FOOBAR')
    expect((errUnknown as any).error).toBe('#NAME? Unknown identifier "FOOBAR" at position 0')

    const errMissingParen = parseFormula('=(1 + 2')
    expect((errMissingParen as any).error).toBe('#NAME? Missing closing parenthesis for "(" at position 0')

    const errFnMissingParen = parseFormula('=SUM(A1, B1')
    expect((errFnMissingParen as any).error).toBe('#NAME? Missing closing parenthesis for function SUM at position 10')

    const errFnMissingComma = parseFormula('=SUM(A1 B1)')
    expect((errFnMissingComma as any).error).toBe('#NAME? Expected comma or closing parenthesis in function SUM at position 7')

    const errUnexpected = parseFormula('=1 + 2 3')
    expect((errUnexpected as any).error).toBe('#NAME? Unexpected token "3" at position 6')

    // 3. Hand-made raw token validation for cellRef/rangeRef
    const strippedCell = parseFormula([{ type: 'cell', value: 'A1', start: 0 }])
    expect(strippedCell).toEqual({ type: 'error', error: '#REF! Invalid cell reference at position 0' })

    const strippedRange = parseFormula([{ type: 'range', value: 'A1:B2', start: 0 }])
    expect(strippedRange).toEqual({ type: 'error', error: '#REF! Invalid range reference at position 0' })

    // 4. Error literals as operands in binary expressions preserve AST tree
    expect(parseFormula('1 + #DIV/0!')).toEqual({
      type: 'binary',
      op: '+',
      left: { type: 'number', value: 1 },
      right: { type: 'error', error: '#DIV/0!' },
    })

    expect(parseFormula('=#DIV/0! + 1')).toEqual({
      type: 'binary',
      op: '+',
      left: { type: 'error', error: '#DIV/0!' },
      right: { type: 'number', value: 1 },
    })

    // 5. TRUE() and FALSE() zero-argument functions
    expect(parseFormula('=TRUE()')).toEqual({ type: 'boolean', value: true })
    expect(parseFormula('=FALSE()')).toEqual({ type: 'boolean', value: false })

    // 6. Long unary chain within 8192-char cap parses without throwing
    const unaryChain8000 = '-'.repeat(8000) + '1'
    expect(() => parseFormula(unaryChain8000)).not.toThrow()
  })
})

describe('xlsx formula evaluator', () => {
  test('Evaluates arithmetic operations with 15-digit rounding and percent (10 + 50% = 10.5)', () => {
    // 1. Basic arithmetic
    expect(evaluateFormula('1 + 2')).toBe(3)
    expect(evaluateFormula('10 - 4')).toBe(6)
    expect(evaluateFormula('6 * 7')).toBe(42)
    expect(evaluateFormula('20 / 4')).toBe(5)

    // 2. 15-digit floating-point precision rounding (avoiding 0.30000000000000004)
    expect(evaluateFormula('0.1 + 0.2')).toBe(0.3)
    expect(evaluateFormula('1 - 0.9')).toBe(0.1)
    expect(evaluateFormula('10 * 0.1')).toBe(1)

    // 3. Percent postfix
    expect(evaluateFormula('50%')).toBe(0.5)
    expect(evaluateFormula('10 + 50%')).toBe(10.5)
    expect(evaluateFormula('50% ^ 2')).toBe(0.25)
    expect(evaluateFormula('-50%')).toBe(-0.5)

    // 4. Excel unary precedence with exponentiation (-2^2 = 4)
    expect(evaluateFormula('-2^2')).toBe(4)
    expect(evaluateFormula('-(2^2)')).toBe(-4)
    expect(evaluateFormula('2^-2')).toBe(0.25)
    expect(evaluateFormula('-2^-2')).toBe(0.25)
    expect(evaluateFormula('2^3^2')).toBe(64) // left-associative in Excel: (2^3)^2 = 8^2 = 64

    // 5. Division by zero and invalid powers
    expect(evaluateFormula('10 / 0')).toBe('#DIV/0!')
    expect(evaluateFormula('0 / 0')).toBe('#DIV/0!')
    expect(evaluateFormula('0 ^ 0')).toBe(1) // in Excel, 0^0 evaluates to 1
  })

  test('Evaluates string concatenation (&) and Excel comparison ordering (number < text < FALSE < TRUE)', () => {
    // 1. String concatenation
    expect(evaluateFormula('"Hello " & "World"')).toBe('Hello World')
    expect(evaluateFormula('"Value: " & 42')).toBe('Value: 42')
    expect(evaluateFormula('TRUE & " is truth"')).toBe('TRUE is truth')
    expect(evaluateFormula('1 & 2 & 3')).toBe('123')

    // 2. Intra-type comparisons
    expect(evaluateFormula('1 < 2')).toBe(true)
    expect(evaluateFormula('2 <= 2')).toBe(true)
    expect(evaluateFormula('5 > 3')).toBe(true)
    expect(evaluateFormula('5 >= 5')).toBe(true)
    expect(evaluateFormula('5 = 5')).toBe(true)
    expect(evaluateFormula('5 <> 6')).toBe(true)

    // Case-insensitive string comparison in Excel
    expect(evaluateFormula('"apple" = "APPLE"')).toBe(true)
    expect(evaluateFormula('"apple" <> "APPLE"')).toBe(false)
    expect(evaluateFormula('"abc" < "def"')).toBe(true)

    // Boolean comparisons
    expect(evaluateFormula('FALSE < TRUE')).toBe(true)
    expect(evaluateFormula('TRUE = TRUE')).toBe(true)
    expect(evaluateFormula('FALSE = FALSE')).toBe(true)

    // 3. Cross-type Excel comparison hierarchy: Number < Text < FALSE < TRUE
    expect(evaluateFormula('1000000 < "a"')).toBe(true)
    expect(evaluateFormula('"z" > 999999')).toBe(true)
    expect(evaluateFormula('"text" < FALSE')).toBe(true)
    expect(evaluateFormula('"text" < TRUE')).toBe(true)
    expect(evaluateFormula('100 < FALSE')).toBe(true)
    expect(evaluateFormula('TRUE > "anything"')).toBe(true)
    expect(evaluateFormula('TRUE > 9999999')).toBe(true)
    expect(evaluateFormula('"10" = 10')).toBe(false)
    expect(evaluateFormula('"10" <> 10')).toBe(true)
  })

  test('Propagates error values through operations (1 + #DIV/0! = #DIV/0!)', () => {
    // 1. Unary error propagation
    expect(evaluateFormula('-#DIV/0!')).toBe('#DIV/0!')
    expect(evaluateFormula('+#VALUE!')).toBe('#VALUE!')
    expect(evaluateFormula('#REF!%')).toBe('#REF!')

    // 2. Binary arithmetic error propagation
    expect(evaluateFormula('1 + #DIV/0!')).toBe('#DIV/0!')
    expect(evaluateFormula('=#REF! * 5')).toBe('#REF!')
    expect(evaluateFormula('10 - #NUM!')).toBe('#NUM!')
    expect(evaluateFormula('#N/A / 2')).toBe('#N/A')
    expect(evaluateFormula('2 ^ #VALUE!')).toBe('#VALUE!')

    // 3. String concatenation and comparison error propagation
    expect(evaluateFormula('"Prefix: " & #REF!')).toBe('#REF!')
    expect(evaluateFormula('=#DIV/0! & " Suffix"')).toBe('#DIV/0!')
    expect(evaluateFormula('#NULL! = 0')).toBe('#NULL!')
    expect(evaluateFormula('100 < #NAME?')).toBe('#NAME?')

    // 4. First-error-wins in binary operations
    expect(evaluateFormula('#DIV/0! + #REF!')).toBe('#DIV/0!')
    expect(evaluateFormula('#REF! * #VALUE!')).toBe('#REF!')

    // 5. Deeply nested error propagation
    expect(evaluateFormula('((1 + 2) * #NUM!) / 4')).toBe('#NUM!')
    expect(evaluateFormula('1 + (2 * (3 + #DIV/0!))')).toBe('#DIV/0!')
  })

  test('Evaluates Math functions (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ABS, ROUND, INT, MOD, PRODUCT)', () => {
    // 1. SUM
    expect(evaluateFormula('SUM(1, 2, 3, 4)')).toBe(10)
    expect(evaluateFormula('SUM(10, -5, 2.5)')).toBe(7.5)

    // 2. AVERAGE
    expect(evaluateFormula('AVERAGE(10, 20, 30)')).toBe(20)
    expect(evaluateFormula('AVERAGE(1, 2)')).toBe(1.5)

    // 3. MIN and MAX
    expect(evaluateFormula('MIN(5, 2, 9, -1, 4)')).toBe(-1)
    expect(evaluateFormula('MAX(5, 2, 9, -1, 4)')).toBe(9)

    // 4. COUNT and COUNTA
    expect(evaluateFormula('COUNT(1, "hello", TRUE, 42)')).toBe(3) // direct args count numbers and booleans
    expect(evaluateFormula('COUNTA(1, "hello", TRUE, 42)')).toBe(4) // all non-empty

    // 5. ABS
    expect(evaluateFormula('ABS(-42.5)')).toBe(42.5)
    expect(evaluateFormula('ABS(42.5)')).toBe(42.5)

    // 6. ROUND
    expect(evaluateFormula('ROUND(3.14159, 2)')).toBe(3.14)
    expect(evaluateFormula('ROUND(3.14159, 0)')).toBe(3)
    expect(evaluateFormula('ROUND(125.4, -1)')).toBe(130)

    // 7. INT
    expect(evaluateFormula('INT(3.7)')).toBe(3)
    expect(evaluateFormula('INT(-3.7)')).toBe(-4)

    // 8. MOD (Excel floor-division modulo)
    expect(evaluateFormula('MOD(7, 3)')).toBe(1)
    expect(evaluateFormula('MOD(-7, 3)')).toBe(2)
    expect(evaluateFormula('MOD(7, -3)')).toBe(-2)
    expect(evaluateFormula('MOD(7, 0)')).toBe('#DIV/0!')

    // 9. PRODUCT
    expect(evaluateFormula('PRODUCT(2, 3, 4)')).toBe(24)
    expect(evaluateFormula('PRODUCT(2.5, 4)')).toBe(10)

    // 10. Evaluation with cell references in context
    const ctx = {
      getCellValue: (_s: any, col: number, row: number) => {
        if (col === 0 && row === 0) return 10 // A1
        if (col === 0 && row === 1) return 20 // A2
        if (col === 0 && row === 2) return 30 // A3
        return null
      },
    }
    expect(evaluateFormula('SUM(A1, A2, A3)', ctx)).toBe(60)
    expect(evaluateFormula('AVERAGE(A1, A2, A3)', ctx)).toBe(20)
  })

  test('Evaluates Logic functions with short-circuiting (IF, AND, OR, NOT, IFERROR)', () => {
    // 1. IF basic
    expect(evaluateFormula('IF(1 > 0, "yes", "no")')).toBe('yes')
    expect(evaluateFormula('IF(1 < 0, "yes", "no")')).toBe('no')
    expect(evaluateFormula('IF(FALSE, "yes")')).toBe(false) // omitted false branch returns false

    // 2. IF short-circuiting: unchosen branch with error is never evaluated
    expect(evaluateFormula('IF(TRUE, 42, 1/0)')).toBe(42)
    expect(evaluateFormula('IF(FALSE, 1/0, 99)')).toBe(99)

    // 3. IFERROR
    expect(evaluateFormula('IFERROR(10 / 2, "err")')).toBe(5)
    expect(evaluateFormula('IFERROR(10 / 0, "caught")')).toBe('caught')
    expect(evaluateFormula('IFERROR(#DIV/0!, "fallback")')).toBe('fallback')
    expect(evaluateFormula('IFERROR(#REF!, "fallback")')).toBe('fallback')
    expect(evaluateFormula('IFERROR(42, 1/0)')).toBe(42) // short-circuits error in fallback

    // 4. AND, OR, NOT
    expect(evaluateFormula('AND(TRUE, TRUE, TRUE)')).toBe(true)
    expect(evaluateFormula('AND(TRUE, FALSE, TRUE)')).toBe(false)
    expect(evaluateFormula('AND(1 > 0, 2 < 5)')).toBe(true)

    expect(evaluateFormula('OR(FALSE, FALSE, TRUE)')).toBe(true)
    expect(evaluateFormula('OR(FALSE, FALSE, FALSE)')).toBe(false)

    expect(evaluateFormula('NOT(TRUE)')).toBe(false)
    expect(evaluateFormula('NOT(FALSE)')).toBe(true)
    expect(evaluateFormula('NOT(0)')).toBe(true)
    expect(evaluateFormula('NOT(1)')).toBe(false)
  })

  test('Evaluates Text functions (CONCAT, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER)', () => {
    // 1. CONCAT
    expect(evaluateFormula('CONCAT("A", "B", "C")')).toBe('ABC')
    expect(evaluateFormula('CONCAT("Val: ", 42, " ", TRUE)')).toBe('Val: 42 TRUE')

    // 2. LEFT and RIGHT
    expect(evaluateFormula('LEFT("Spreadsheet", 6)')).toBe('Spread')
    expect(evaluateFormula('LEFT("Hello")')).toBe('H') // default num_chars = 1
    expect(evaluateFormula('RIGHT("Spreadsheet", 5)')).toBe('sheet')
    expect(evaluateFormula('RIGHT("Hello")')).toBe('o') // default num_chars = 1
    expect(evaluateFormula('LEFT("ABC", 10)')).toBe('ABC')
    expect(evaluateFormula('RIGHT("ABC", 10)')).toBe('ABC')

    // 3. MID (1-based index)
    expect(evaluateFormula('MID("Spreadsheet", 7, 5)')).toBe('sheet')
    expect(evaluateFormula('MID("Spreadsheet", 8, 4)')).toBe('heet')
    expect(evaluateFormula('MID("Hello", 2, 3)')).toBe('ell')
    expect(evaluateFormula('MID("Hello", 10, 2)')).toBe('')

    // 4. LEN
    expect(evaluateFormula('LEN("Hello World")')).toBe(11)
    expect(evaluateFormula('LEN(12345)')).toBe(5)
    expect(evaluateFormula('LEN("")')).toBe(0)

    // 5. TRIM (collapses internal consecutive spaces)
    expect(evaluateFormula('TRIM("  Hello   World  ")')).toBe('Hello World')

    // 6. UPPER and LOWER
    expect(evaluateFormula('UPPER("excel spreadsheet")')).toBe('EXCEL SPREADSHEET')
    expect(evaluateFormula('LOWER("EXCEL SPREADSHEET")')).toBe('excel spreadsheet')
  })

  test('Handles blank cells correctly (0 in math, "" in concat, ignored in SUM)', () => {
    // Context with blank cells (A1 has value 10, B1 is blank/null, C1 is 0, D1 is "")
    const ctx: EvaluationContext = {
      getCellValue: (_sheet, col, row) => {
        if (col === 0 && row === 0) return 10 // A1 = 10
        if (col === 1 && row === 0) return null // B1 = blank
        if (col === 2 && row === 0) return 0 // C1 = 0
        if (col === 3 && row === 0) return '' // D1 = ""
        return null // any other cell is blank
      },
      getRangeValues: (_sheet, from, to) => {
        const minRow = Math.min(from.row, to.row)
        const maxRow = Math.max(from.row, to.row)
        const minCol = Math.min(from.col, to.col)
        const maxCol = Math.max(from.col, to.col)
        const rows: FormulaValue[][] = []
        for (let r = minRow; r <= maxRow; r++) {
          const rowVals: FormulaValue[] = []
          for (let c = minCol; c <= maxCol; c++) {
            if (c === 0 && r === 0) rowVals.push(10)
            else if (c === 2 && r === 0) rowVals.push(0)
            else if (c === 3 && r === 0) rowVals.push('')
            else rowVals.push(null)
          }
          rows.push(rowVals)
        }
        return rows
      },
    }

    // 1. Math operations: blank cell treated as 0
    expect(evaluateFormula('B1 + 5', ctx)).toBe(5)
    expect(evaluateFormula('5 + B1', ctx)).toBe(5)
    expect(evaluateFormula('B1 - 5', ctx)).toBe(-5)
    expect(evaluateFormula('5 - B1', ctx)).toBe(5)
    expect(evaluateFormula('B1 * 10', ctx)).toBe(0)
    expect(evaluateFormula('B1 / 2', ctx)).toBe(0)
    expect(evaluateFormula('2 / B1', ctx)).toBe('#DIV/0!')
    expect(evaluateFormula('-B1', ctx)).toBe(0)
    expect(evaluateFormula('+B1', ctx)).toBe(0)

    // 2. String concatenation: blank cell treated as ""
    expect(evaluateFormula('B1 & "hello"', ctx)).toBe('hello')
    expect(evaluateFormula('"world" & B1', ctx)).toBe('world')
    expect(evaluateFormula('B1 & B1', ctx)).toBe('')
    expect(evaluateFormula('A1 & B1 & "!"', ctx)).toBe('10!')

    // 3. Comparison operations: blank cell equals 0 and equals ""
    expect(evaluateFormula('B1 = 0', ctx)).toBe(true)
    expect(evaluateFormula('B1 = ""', ctx)).toBe(true)
    expect(evaluateFormula('B1 = C1', ctx)).toBe(true) // B1 (blank) = C1 (0)
    expect(evaluateFormula('B1 = D1', ctx)).toBe(true) // B1 (blank) = D1 ("")
    expect(evaluateFormula('B1 <> 1', ctx)).toBe(true)
    expect(evaluateFormula('B1 = 1', ctx)).toBe(false)
    expect(evaluateFormula('B1 > -1', ctx)).toBe(true)
    expect(evaluateFormula('B1 < 1', ctx)).toBe(true)

    // 4. Function aggregation: blank cells are ignored
    // SUM: blank cells ignored, returns 0 if all blank
    expect(evaluateFormula('SUM(A1, B1)', ctx)).toBe(10)
    expect(evaluateFormula('SUM(B1)', ctx)).toBe(0)
    expect(evaluateFormula('SUM(B1:B5)', ctx)).toBe(0)
    expect(evaluateFormula('SUM(A1:B1)', ctx)).toBe(10)

    // AVERAGE: blank cells ignored from count
    expect(evaluateFormula('AVERAGE(A1, B1)', ctx)).toBe(10)
    expect(evaluateFormula('AVERAGE(B1)', ctx)).toBe('#DIV/0!')

    // COUNT: blank cells not counted
    expect(evaluateFormula('COUNT(A1, B1)', ctx)).toBe(1)
    expect(evaluateFormula('COUNT(B1)', ctx)).toBe(0)
    expect(evaluateFormula('COUNT(B1:B5)', ctx)).toBe(0)

    // COUNTA: blank cells not counted
    expect(evaluateFormula('COUNTA(A1, B1)', ctx)).toBe(1)
    expect(evaluateFormula('COUNTA(B1)', ctx)).toBe(0)

    // MIN / MAX: blank cells ignored
    expect(evaluateFormula('MIN(A1, B1)', ctx)).toBe(10)
    expect(evaluateFormula('MAX(A1, B1)', ctx)).toBe(10)
    expect(evaluateFormula('MIN(B1)', ctx)).toBe(0)

    // PRODUCT: blank cells ignored
    expect(evaluateFormula('PRODUCT(A1, B1)', ctx)).toBe(10)
    expect(evaluateFormula('PRODUCT(B1)', ctx)).toBe(0)

    // 5. Direct blank evaluation and unbound cells without context
    expect(evaluateFormula('B1', ctx)).toBeNull()
    expect(evaluateFormula('Z99')).toBeNull()
    expect(evaluateFormula('Z99 + 5')).toBe(5)
    expect(evaluateFormula('Z99 & "abc"')).toBe('abc')
  })

  test('Detects circular references and returns 0 without stack overflow', () => {
    // 1. Direct self-reference: A1 = A1, A2 = A2 + 10
    const cells: Record<string, string> = {
      'A1': 'A1',
      'A2': 'A2 + 10',
    }
    const ctx: EvaluationContext = {
      currentSheet: 'Sheet1',
      getCellValue: (_sheet, col, row) => {
        const colLetter = String.fromCharCode(65 + col)
        const ref = `${colLetter}${row + 1}`
        const formula = cells[ref]
        if (formula) {
          return evaluateFormula(formula, ctx)
        }
        return null
      },
    }

    // In Excel, any cell involved in a circular reference evaluates to 0
    expect(evaluateFormula('A1', ctx)).toBe(0)
    expect(evaluateFormula('A2', ctx)).toBe(0)

    // 2. Mutual circular reference: A1 = B1 + 1, B1 = A1 + 1
    const mutualCells: Record<string, string> = {
      'A1': 'B1 + 1',
      'B1': 'A1 + 1',
    }
    const mutualCtx: EvaluationContext = {
      currentSheet: 'Sheet1',
      getCellValue: (_sheet, col, row) => {
        const colLetter = String.fromCharCode(65 + col)
        const ref = `${colLetter}${row + 1}`
        const formula = mutualCells[ref]
        if (formula) {
          return evaluateFormula(formula, mutualCtx)
        }
        return null
      },
    }

    // In Excel, mutual circular references evaluate to 0
    expect(evaluateFormula('A1', mutualCtx)).toBe(0)
    expect(evaluateFormula('B1', mutualCtx)).toBe(0)

    // 3. 3-node cycle: A1 = B1, B1 = C1, C1 = A1
    const cycle3: Record<string, string> = {
      'A1': 'B1',
      'B1': 'C1',
      'C1': 'A1',
    }
    const cycle3Ctx: EvaluationContext = {
      currentSheet: 'Sheet1',
      getCellValue: (_sheet, col, row) => {
        const colLetter = String.fromCharCode(65 + col)
        const ref = `${colLetter}${row + 1}`
        const formula = cycle3[ref]
        if (formula) {
          return evaluateFormula(formula, cycle3Ctx)
        }
        return null
      },
    }
    expect(evaluateFormula('A1', cycle3Ctx)).toBe(0)

    // 4. Circular range reference: A1 = SUM(A1:A3), where A2=10, A3=20
    const rangeCycle: Record<string, string | number> = {
      'A1': 'SUM(A1:A3)',
      'A2': 10,
      'A3': 20,
    }
    const rangeCycleCtx: EvaluationContext = {
      currentSheet: 'Sheet1',
      getCellValue: (_sheet, col, row) => {
        const colLetter = String.fromCharCode(65 + col)
        const ref = `${colLetter}${row + 1}`
        const val = rangeCycle[ref]
        if (typeof val === 'string') {
          return evaluateFormula(val, rangeCycleCtx)
        }
        return val ?? null
      },
    }
    expect(evaluateFormula('A1', rangeCycleCtx)).toBe(0)
  })

  test('Evaluator addresses peer review hardening (poisoned cycle 0, direct arg coercion, half-away ROUND, 0^0, overflow #NUM!, space-32 TRIM)', () => {
    // 1. Direct argument coercion in aggregates (Excel parity)
    expect(evaluateFormula('SUM("5", 2)')).toBe(7)
    expect(evaluateFormula('SUM(TRUE, 2)')).toBe(3)
    expect(evaluateFormula('SUM("abc", 2)')).toBe('#VALUE!')
    expect(evaluateFormula('AVERAGE("3")')).toBe(3)
    // Per Microsoft Excel specification for COUNT:
    // "Arguments that are error values or text that cannot be translated into numbers are not counted."
    // Hence direct text that cannot be converted to number is not counted (0), whereas numeric strings ("10")
    // and booleans (TRUE) are counted. In contrast, SUM/AVERAGE attempt coercion and raise #VALUE!.
    expect(evaluateFormula('COUNT(TRUE, 42, "10", "text")')).toBe(3)
    expect(evaluateFormula('COUNT("abc")')).toBe(0)

    // In cell/range references, text and booleans are ignored in SUM
    const cellCtx: EvaluationContext = {
      getCellValue: (_sheet, col, row) => {
        if (col === 0 && row === 0) return '5' // A1 text "5"
        if (col === 0 && row === 1) return true // A2 boolean true
        if (col === 0 && row === 2) return 10 // A3 number 10
        return null
      },
    }
    expect(evaluateFormula('SUM(A1, A2, A3)', cellCtx)).toBe(10)
    expect(evaluateFormula('COUNT(A1, A2, A3)', cellCtx)).toBe(1)

    // 2. 0^0 = 1 in Excel
    expect(evaluateFormula('0^0')).toBe(1)
    expect(evaluateFormula('0^-1')).toBe('#NUM!')
    expect(evaluateFormula('(-1)^0.5')).toBe('#NUM!')

    // 3. Half-away-from-zero ROUND
    expect(evaluateFormula('ROUND(-1.5, 0)')).toBe(-2)
    expect(evaluateFormula('ROUND(-2.5, 0)')).toBe(-3)
    expect(evaluateFormula('ROUND(2.5, 0)')).toBe(3)
    expect(evaluateFormula('ROUND(1, 309)')).toBe('#NUM!')

    // 4. Arithmetic overflow guard returning #NUM!
    expect(evaluateFormula('1e308 * 10')).toBe('#NUM!')
    expect(evaluateFormula('1e308 + 1e308')).toBe('#NUM!')

    // 5. Space-32 TRIM (preserves tabs and NBSP per Excel spec)
    expect(evaluateFormula('TRIM("  Hello   World  ")')).toBe('Hello World')
    expect(evaluateFormula('TRIM("a\tb")')).toBe('a\tb')
    expect(evaluateFormula('TRIM("a\u00A0b")')).toBe('a\u00A0b')

    // 6. INT precision guard
    expect(evaluateFormula('INT(1.999999999999999)')).toBe(2)
    expect(evaluateFormula('INT(2.1)')).toBe(2)

    // 7. Direct empty/null arguments in aggregate functions (Excel parity)
    // Direct empty args (e.g. MIN(5,)) are ignored rather than coerced to 0
    expect(evaluateFormula('MIN(5, )')).toBe(5)
    expect(evaluateFormula('MAX(-5, )')).toBe(-5)
    expect(evaluateFormula('AVERAGE(, )')).toBe('#DIV/0!')
    expect(evaluateFormula('PRODUCT(5, )')).toBe(5)
    expect(evaluateFormula('COUNT(, )')).toBe(0)
    expect(evaluateFormula('SUM(, )')).toBe(0)
  })
})

describe('xlsx shared formula & workbook context', () => {
  test('Translates shared formula relative references by row/col offset (si master to dependent cells)', () => {
    // 1. Relative translation: A1 + B1 shifted by (dCol=0, dRow=1) -> A2 + B2
    const res1 = translateSharedFormula('A1 + B1', 0, 1)
    expect(res1.formula).toBe('A2+B2')

    // 2. Absolute and mixed references: $A$1 + A$1 + $A1 + A1 shifted by (dCol=2, dRow=3)
    const res2 = translateSharedFormula('$A$1 + A$1 + $A1 + A1', 2, 3)
    expect(res2.formula).toBe('$A$1+C$1+$A4+C4')

    // 3. Range references: SUM(A1:B2) shifted by (dCol=1, dRow=2) -> SUM(B3:C4)
    const res3 = translateSharedFormula('SUM(A1:B2)', 1, 2)
    expect(res3.formula).toBe('SUM(B3:C4)')

    // 4. Cross-sheet references: Sheet2!A1 shifted by (dCol=1, dRow=1) -> Sheet2!B2
    const res4 = translateSharedFormula('Sheet2!A1', 1, 1)
    expect(res4.formula).toBe('Sheet2!B2')

    // 5. Out of bounds translation produces #REF!
    const res5 = translateSharedFormula('A1', -1, 0)
    expect(res5.formula).toBe('#REF!')
  })

  test('Resolves cross-sheet references (Sheet2!A1) using workbook-level context', () => {
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Summary',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: null, formula: 'Sheet2!A1 + Data!B1', styleIndex: 0 },
                { ref: 'A2', col: 0, row: 1, value: null, formula: 'SUM(Data!A1:B2)', styleIndex: 0 },
              ],
            },
          ],
        },
        {
          name: 'Sheet2',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: 100, styleIndex: 0 },
              ],
            },
          ],
        },
        {
          name: 'Data',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: 10, styleIndex: 0 },
                { ref: 'B1', col: 1, row: 0, value: 20, styleIndex: 0 },
              ],
            },
            {
              index: 1,
              cells: [
                { ref: 'A2', col: 0, row: 1, value: 30, styleIndex: 0 },
                { ref: 'B2', col: 1, row: 1, value: 40, styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }

    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(120) // 100 + 20
    expect(doc.sheets[0].rows[0].cells[1].value).toBe(100) // 10 + 20 + 30 + 40
  })

  test('Evaluates multi-cell dependency chains across rows and sheets in correct order with per-cell memoization', () => {
    // 600-cell dependency chain: A1 = 1, A2 = A1 + 1, A3 = A2 + 1, ..., A600 = A599 + 1
    const cells: XlsxCell[] = [
      { ref: 'A1', col: 0, row: 0, value: 1, styleIndex: 0 },
    ]
    for (let i = 1; i < 600; i++) {
      cells.push({
        ref: `A${i + 1}`,
        col: 0,
        row: i,
        value: null,
        formula: `A${i} + 1`,
        styleIndex: 0,
      })
    }

    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Chain',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: cells.map((cell, idx) => ({ index: idx, cells: [cell] })),
        },
      ],
    }

    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(1)
    expect(doc.sheets[0].rows[1].cells[0].value).toBe(2)
    expect(doc.sheets[0].rows[599].cells[0].value).toBe(600)
  })

  test('Preserves cached <v> unless missing, ca="1", or fullCalcOnLoad="1"', () => {
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Cache',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                // Cell with cached value 999
                { ref: 'A1', col: 0, row: 0, value: 999, formula: '10 + 20', styleIndex: 0 },
                // Cell with missing value (null)
                { ref: 'B1', col: 1, row: 0, value: null, formula: '10 + 20', styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }

    // Default: preserves cached <v>, calculates missing <v>
    evaluateWorkbookFormulas(doc)
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(999) // preserved
    expect(doc.sheets[0].rows[0].cells[1].value).toBe(30) // calculated

    // With fullCalcOnLoad: recalculates all
    evaluateWorkbookFormulas(doc, { fullCalcOnLoad: true })
    expect(doc.sheets[0].rows[0].cells[0].value).toBe(30) // recalculated
    expect(doc.sheets[0].rows[0].cells[1].value).toBe(30)
  })

  test('Addresses Phase 4 peer review hardening (workbook cycle poisoning rollback, range clamp, ca flag, multi-sheet package roundtrip)', async () => {
    // 1. Workbook cycle poisoning: mutual cycle A1=B1+1, B1=A1+1 both evaluate to 0
    // Sibling C1=A1+100 evaluates to 100, D1=20 is completely uncontaminated
    const cycleDoc: XlsxDocument = {
      sheets: [
        {
          name: 'CycleSheet',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: null, formula: 'B1 + 1', styleIndex: 0 },
                { ref: 'B1', col: 1, row: 0, value: null, formula: 'A1 + 1', styleIndex: 0 },
                { ref: 'C1', col: 2, row: 0, value: null, formula: 'A1 + 100', styleIndex: 0 },
                { ref: 'D1', col: 3, row: 0, value: null, formula: '10 * 2', styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }

    evaluateWorkbookFormulas(cycleDoc)
    expect(cycleDoc.sheets[0].rows[0].cells[0].value).toBe(0) // A1 poisoned to 0
    expect(cycleDoc.sheets[0].rows[0].cells[1].value).toBe(0) // B1 poisoned to 0
    expect(cycleDoc.sheets[0].rows[0].cells[2].value).toBe(100) // C1 = 0 + 100
    expect(cycleDoc.sheets[0].rows[0].cells[3].value).toBe(20) // D1 unaffected

    // Evaluate-twice stability
    evaluateWorkbookFormulas(cycleDoc, { forceRecalc: true })
    expect(cycleDoc.sheets[0].rows[0].cells[0].value).toBe(0)
    expect(cycleDoc.sheets[0].rows[0].cells[1].value).toBe(0)
    expect(cycleDoc.sheets[0].rows[0].cells[2].value).toBe(100)

    // Dependent-first evaluation order: D1=B1+C1 evaluated before B1/A1
    // Suffix marking ensures only B1/A1 are zeroed, preserving D1=0+100=100
    const depFirstDoc: XlsxDocument = {
      sheets: [
        {
          name: 'DepFirst',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'D1', col: 3, row: 0, value: null, formula: 'B1 + C1', styleIndex: 0 },
                { ref: 'A1', col: 0, row: 0, value: null, formula: 'B1 + 1', styleIndex: 0 },
                { ref: 'B1', col: 1, row: 0, value: null, formula: 'A1 + 1', styleIndex: 0 },
                { ref: 'C1', col: 2, row: 0, value: 100, styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }
    evaluateWorkbookFormulas(depFirstDoc)
    expect(depFirstDoc.sheets[0].rows[0].cells[0].value).toBe(100) // D1 = 0 + 100 = 100
    expect(depFirstDoc.sheets[0].rows[0].cells[1].value).toBe(0) // A1 in cycle -> 0
    expect(depFirstDoc.sheets[0].rows[0].cells[2].value).toBe(0) // B1 in cycle -> 0
    expect(depFirstDoc.sheets[0].rows[0].cells[3].value).toBe(100) // C1 = 100

    // 2. Range clamping: SUM(A1:XFD1) clamps to used area and executes in milliseconds
    const largeRangeDoc: XlsxDocument = {
      sheets: [
        {
          name: 'BigRange',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: 5, styleIndex: 0 },
                { ref: 'B1', col: 1, row: 0, value: 10, styleIndex: 0 },
              ],
            },
            {
              index: 1,
              cells: [
                { ref: 'A2', col: 0, row: 1, value: null, formula: 'SUM(A1:XFD1)', styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }
    evaluateWorkbookFormulas(largeRangeDoc)
    expect(largeRangeDoc.sheets[0].rows[1].cells[0].value).toBe(15)

    // 3. ca="1" forces recalculation even when cached <v> is present
    const caDoc: XlsxDocument = {
      sheets: [
        {
          name: 'CaSheet',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                { ref: 'A1', col: 0, row: 0, value: 888, formula: '5 + 5', ca: true, styleIndex: 0 },
              ],
            },
          ],
        },
      ],
    }
    evaluateWorkbookFormulas(caDoc)
    expect(caDoc.sheets[0].rows[0].cells[0].value).toBe(10) // recalculated because ca=true

    // 4. Multi-sheet round-trip with buildXlsx and parseXlsx
    const multiBuf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', formula: 'Sheet2!A1 * 3', v: 999, ca: true },
            ],
          },
        ],
      },
      {
        name: 'Sheet2',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', v: 42 },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(multiBuf)
    const parsedDoc = await parseXlsx(pkg)
    expect(parsedDoc.sheets.length).toBe(2)
    expect(parsedDoc.sheets[0].rows[0].cells[0].ca).toBe(true)

    evaluateWorkbookFormulas(parsedDoc)
    expect(parsedDoc.sheets[0].rows[0].cells[0].value).toBe(126) // 42 * 3

    // 5. CJK unquoted sheet names in formatCellRef
    const cjkRes = translateSharedFormula('工作表1!A1', 1, 0)
    expect(cjkRes.formula).toBe('工作表1!B1')
  })
})

describe('xlsx formula integration & canvas rendering', () => {
  test('parseXlsx evaluates formula cells when <v> is absent or ca="1"', async () => {
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              // Formula with missing <v>
              { ref: 'A1', formula: '10 + 20' },
              // Dependent formula with missing <v>
              { ref: 'B1', formula: 'A1 * 2' },
              // Formula with cached <v> preserved
              { ref: 'C1', formula: '99 * 2', v: 5 },
              // Formula with cached <v> but ca="1" recalculated
              { ref: 'D1', formula: '100 * 2', v: 5, ca: true },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)

    expect(doc.sheets[0].rows[0].cells[0].value).toBe(30)
    expect(doc.sheets[0].rows[0].cells[1].value).toBe(60)
    expect(doc.sheets[0].rows[0].cells[2].value).toBe(5)
    expect(doc.sheets[0].rows[0].cells[3].value).toBe(200)
  })

  test('Viewport culling does not break off-screen formula dependencies', async () => {
    // A1 depends on Z100 which is far off-screen
    const buf = await buildXlsx([
      {
        name: 'Sheet1',
        rows: [
          {
            r: 1,
            cells: [
              { ref: 'A1', formula: 'Z100 * 2' },
            ],
          },
          {
            r: 100,
            cells: [
              { ref: 'Z100', v: 50 },
            ],
          },
        ],
      },
    ])

    const pkg = await OfficePackage.load(buf)
    const doc = await parseXlsx(pkg)
    const sheet = doc.sheets[0]

    // Verify parsed and calculated value
    expect(sheet.rows[0].cells[0].value).toBe(100)

    // Render with viewport covering only top-left region (A1 inside, Z100 far outside)
    const paintedTexts: Array<{ text: string; x: number; y: number }> = []
    const fakeCtx = {
      save: () => {},
      restore: () => {},
      beginPath: () => {},
      rect: () => {},
      clip: () => {},
      fillRect: () => {},
      stroke: () => {},
      moveTo: () => {},
      lineTo: () => {},
      translate: () => {},
      measureText: (text: string) => ({ width: text.length * 8 }),
      fillText: (text: string, x: number, y: number) => {
        paintedTexts.push({ text, x, y })
      },
      font: '',
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      textBaseline: 'alphabetic',
    }

    renderSheet(sheet, fakeCtx as unknown as CanvasRenderingContext2D, undefined, undefined, undefined, {
      x: 0,
      y: 0,
      width: 150,
      height: 50,
    })

    // On-screen cell A1 is painted with computed formula value "100"
    expect(paintedTexts.some(p => p.text === '100')).toBe(true)
    // Off-screen cell Z100 ("50") is culled and NOT painted
    expect(paintedTexts.some(p => p.text === '50')).toBe(false)
  })

  test('renderSheet renders calculated formula cell values onto canvas with correct alignment and styling', async () => {
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Sheet1',
          merges: [],
          mergeRanges: [],
          cols: [{ min: 0, max: 0, widthChars: 10 }], // col 0 width ~75px
          rows: [
            {
              index: 0,
              cells: [
                // A1: computed number, bold
                {
                  ref: 'A1',
                  col: 0,
                  row: 0,
                  value: 30,
                  formula: '10 + 20',
                  styleIndex: 0,
                  style: { bold: true, fontSizePt: 11, numFmtId: 0 },
                },
                // B1: computed text, italic
                {
                  ref: 'B1',
                  col: 1,
                  row: 0,
                  value: 'Hello World',
                  formula: '"Hello" & " World"',
                  styleIndex: 1,
                  style: { italic: true, fontSizePt: 11, numFmtId: 0 },
                },
              ],
            },
          ],
        },
      ],
    }

    const paintedCalls: Array<{ text: string; x: number; y: number; font: string }> = []
    const fakeCtx = {
      save: () => {},
      restore: () => {},
      beginPath: () => {},
      rect: () => {},
      clip: () => {},
      fillRect: () => {},
      stroke: () => {},
      moveTo: () => {},
      lineTo: () => {},
      translate: () => {},
      measureText: (text: string) => ({ width: text.length * 7 }),
      fillText: (text: string, x: number, y: number) => {
        paintedCalls.push({ text, x, y, font: fakeCtx.font })
      },
      font: '',
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      textBaseline: 'alphabetic',
    }

    renderSheet(doc.sheets[0], fakeCtx as unknown as CanvasRenderingContext2D)

    const a1Call = paintedCalls.find(c => c.text === '30')
    expect(a1Call).toBeDefined()
    expect(a1Call?.font).toContain('bold')
    // Numbers are right-aligned (x > 0)
    expect(a1Call?.x).toBeGreaterThan(40)

    const b1Call = paintedCalls.find(c => c.text === 'Hello World')
    expect(b1Call).toBeDefined()
    expect(b1Call?.font).toContain('italic')
    // Text strings are left-aligned (x starts at col 1 start + padding 3)
    const col0Width = Math.round(10 * 7 + 5)
    expect(b1Call?.x).toBe(col0Width + 3)
  })

  test('renderSheet applies number format (numFmtId) to formula results', async () => {
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Sheet1',
          merges: [],
          mergeRanges: [],
          cols: [],
          rows: [
            {
              index: 0,
              cells: [
                {
                  ref: 'A1',
                  col: 0,
                  row: 0,
                  value: 3.3333333333333335,
                  formula: '10 / 3',
                  styleIndex: 0,
                  style: { numFmtId: 2 }, // 0.00
                },
                {
                  ref: 'B1',
                  col: 1,
                  row: 0,
                  value: 0.42,
                  formula: '42 / 100',
                  styleIndex: 1,
                  style: { numFmtId: 9 }, // 0%
                },
                {
                  ref: 'C1',
                  col: 2,
                  row: 0,
                  value: 12345.678,
                  formula: '12345.678',
                  styleIndex: 2,
                  style: { numFmtId: 4 }, // #,##0.00
                },
              ],
            },
          ],
        },
      ],
    }

    const paintedTexts: string[] = []
    const fakeCtx = {
      save: () => {},
      restore: () => {},
      beginPath: () => {},
      rect: () => {},
      clip: () => {},
      fillRect: () => {},
      stroke: () => {},
      moveTo: () => {},
      lineTo: () => {},
      translate: () => {},
      measureText: (text: string) => ({ width: text.length * 7 }),
      fillText: (text: string) => {
        paintedTexts.push(text)
      },
      font: '',
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      textBaseline: 'alphabetic',
    }

    renderSheet(doc.sheets[0], fakeCtx as unknown as CanvasRenderingContext2D)

    expect(paintedTexts).toContain('3.33')
    expect(paintedTexts).toContain('42%')
    expect(paintedTexts).toContain('12,345.68')
  })

  test('renderSheet renders formula error strings (#DIV/0!) with alignment per native Excel center convention', async () => {
    const doc: XlsxDocument = {
      sheets: [
        {
          name: 'Sheet1',
          merges: [],
          mergeRanges: [],
          cols: [{ min: 0, max: 0, widthChars: 12 }], // ~89px wide
          rows: [
            {
              index: 0,
              cells: [
                {
                  ref: 'A1',
                  col: 0,
                  row: 0,
                  value: '#DIV/0!',
                  formula: '10 / 0',
                  styleIndex: 0,
                  style: { numFmtId: 0 },
                },
              ],
            },
          ],
        },
      ],
    }

    const paintedCalls: Array<{ text: string; x: number }> = []
    const textWidth = 7 * 7 // 49px for '#DIV/0!' with 7px/char
    const fakeCtx = {
      save: () => {},
      restore: () => {},
      beginPath: () => {},
      rect: () => {},
      clip: () => {},
      fillRect: () => {},
      stroke: () => {},
      moveTo: () => {},
      lineTo: () => {},
      translate: () => {},
      measureText: () => ({ width: textWidth }),
      fillText: (text: string, x: number) => {
        paintedCalls.push({ text, x })
      },
      font: '',
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      textBaseline: 'alphabetic',
    }

    renderSheet(doc.sheets[0], fakeCtx as unknown as CanvasRenderingContext2D)

    const call = paintedCalls.find(c => c.text === '#DIV/0!')
    expect(call).toBeDefined()
    // Col width is Math.round(12 * 7 + 5) = 89px.
    // Center alignment: x = (89 - 49) / 2 = 20px.
    // Left alignment would have been x = 3px.
    const colWidth = Math.round(12 * 7 + 5)
    const expectedCenterX = (colWidth - textWidth) / 2
    expect(call?.x).toBe(expectedCenterX)
  })
})






