import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { tokenize } from '../src/xlsx/formula/lexer'
import { parseFormula } from '../src/xlsx/formula/parser'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'

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
    expect(b1.value).toBeNull()
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
    expect(b2.formula).toBeUndefined()
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
    expect(evaluateFormula('0 ^ 0')).toBe('#NUM!')
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
    expect(evaluateFormula('COUNT(1, "hello", TRUE, 42)')).toBe(2) // only numbers
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
})




