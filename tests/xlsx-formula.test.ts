import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'
import { tokenize } from '../src/xlsx/formula/lexer'

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
    expect(sheet.rows[0].cells[2].value).toBe('#ERROR')
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
      { type: 'cell', value: 'AA10', cellRef: { col: 26, row: 9, absCol: false, absRow: false } },
      { type: 'eof', value: '', cellRef: undefined },
    ])
  })

  test('Tokenizer distinguishes function calls ending in digits (LOG10) from cell references and handles error literals / fail-closed errors', () => {
    // 1. Disambiguation: LOG10( is ident followed by lparen, not cell LOG10
    const logTokens = tokenize('=LOG10(100)')
    expect(logTokens[0]).toEqual({ type: 'ident', value: 'LOG10' })
    expect(logTokens[1]).toEqual({ type: 'lparen', value: '(' })

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
    expect(unknown[0]).toEqual({ type: 'error', value: '@' })

    // 5. Overlong formula emits error token
    const overlong = tokenize('=' + 'A'.repeat(8200))
    expect(overlong[0]).toEqual({ type: 'error', value: 'Formula exceeds maximum length of 8192 characters' })
  })

  test('Tokenizer handles range references (A1:B10) and cross-sheet references (Sheet2!A1, \'My Sheet\'!A1:B2)', () => {
    // 1. Simple unquoted range
    const r1 = tokenize('=A1:B10')
    expect(r1).toEqual([
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
    expect(r2).toEqual([
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
    expect(c1).toEqual([
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
    expect(r3).toEqual([
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
    expect(c2).toEqual([
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
    expect(t1[0]).toEqual({ type: 'ident', value: 'CONCAT' })

    const t2 = tokenize('=_xlfn.concat(A1, B1)')
    expect(t2[0]).toEqual({ type: 'ident', value: 'CONCAT' })

    const t3 = tokenize('=_xlfn.STDEV.S(A1:B10)')
    expect(t3[0]).toEqual({ type: 'ident', value: 'STDEV.S' })

    const t4 = tokenize('=_xlfn._xlws.FILTER(A1:B10, A1:A10>0)')
    expect(t4[0]).toEqual({ type: 'ident', value: 'FILTER' })

    const t5 = tokenize('=sum(A1:B10) + Average(C1:C10)')
    expect(t5[0]).toEqual({ type: 'ident', value: 'SUM' })
    expect(t5[5]).toEqual({ type: 'ident', value: 'AVERAGE' })
  })
})



