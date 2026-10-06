import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { buildXlsx } from '../src/testdata/ooxml-builders'

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
