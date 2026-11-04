/**
 * Review round 4 finding 1 (the reviewer's repro, kept as a regression):
 * implicit row/cell positions must stay inside the worksheet. More ref-less
 * cells than columns, or a ref-less row after the last row, is a typed
 * PackageError refusal (as the paged reader refuses ref-less input), never an
 * uncaught CoordinateError. One full column of ref-less cells still parses.
 * Excel itself reports such a file as needing repair; no clamp is attempted.
 */
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'
import { MAX_COLS } from '../src/xlsx/view/coordinates'

describe('implicit positioning stays inside the worksheet domain', () => {
  it('a row with more ref-less cells than the worksheet has columns is refused, not thrown', async () => {
    const cells = Array.from({ length: MAX_COLS + 16 }, (_, i) => `<c><v>${i}</v></c>`).join('')
    const pkg = await OfficePackage.load(xlsxBytes({ sheetData: `<row>${cells}</row>` }))
    await expect(parseXlsx(pkg)).rejects.toMatchObject({ name: 'PackageError', code: 'unsupported' })
  })
  it('a ref-less row after the last worksheet row is refused, not thrown', async () => {
    const rows = '<row r="1048576"><c r="A1048576"><v>1</v></c></row>' + Array.from({ length: 3 }, (_, i) => `<row><c><v>${i}</v></c></row>`).join('')
    const pkg = await OfficePackage.load(xlsxBytes({ sheetData: rows }))
    await expect(parseXlsx(pkg)).rejects.toMatchObject({ name: 'PackageError', code: 'unsupported' })
  })
  it('control: exactly one full row of ref-less cells still parses', async () => {
    const cells = Array.from({ length: MAX_COLS }, (_, i) => `<c><v>${i}</v></c>`).join('')
    const doc = await parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData: `<row>${cells}</row>` })))
    expect(doc.sheets[0].rows[0].cells.length).toBe(MAX_COLS)
    expect(doc.sheets[0].rows[0].cells[MAX_COLS - 1].col).toBe(MAX_COLS - 1)
  })
  it('a repeated explicit row number is merged into its first record with a diagnostic', async () => {
    const sheetData = '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="1"><c r="B1"><v>3</v></c><c r="C1"><v>4</v></c></row>'
    const doc = await parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData })))
    const sheet = doc.sheets[0]
    expect(sheet.rows).toHaveLength(1)
    expect(sheet.rows[0].cells.map(cell => [cell.ref, cell.value])).toEqual([['A1', 1], ['B1', 3], ['C1', 4]])
    expect(sheet.diagnostics).toEqual([expect.objectContaining({ feature: 'duplicate-row' })])
  })
})
