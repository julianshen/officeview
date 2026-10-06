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
})

