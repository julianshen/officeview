import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { tokenize } from '../src/xlsx/formula/lexer'
import { parseFormula } from '../src/xlsx/formula/parser'

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
})



