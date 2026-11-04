/**
 * Excel's implicit positions (found through the S13 paint-order change on
 * poi-59746_NoRowNums.xlsx): a <row> without r follows the previous row, and
 * a <c> without r follows the previous cell of its row. The resident parser
 * used to place every such cell at A1, stacking them on top of each other.
 */
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { xlsxBytes } from './helpers/xlsx-package'

describe('implicit row and cell references', () => {
  it('ref-less rows and cells take consecutive positions; explicit refs still win', async () => {
    const sheetData = '<row><c t="inlineStr"><is><t>a</t></is></c><c><v>1</v></c></row>'
      + '<row><c><v>2</v></c><c r="D2"><v>3</v></c><c><v>4</v></c></row>'
      + '<row r="5"><c><v>5</v></c></row><row><c><v>6</v></c></row>'
    const doc = await parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData })))
    const cells = doc.sheets[0].rows.flatMap(row => row.cells.map(cell => [row.index, cell.ref, cell.row, cell.col, cell.value]))
    expect(cells).toEqual([
      [0, 'A1', 0, 0, 'a'], [0, 'B1', 0, 1, 1],
      [1, 'A2', 1, 0, 2], [1, 'D2', 1, 3, 3], [1, 'E2', 1, 4, 4],
      [4, 'A5', 4, 0, 5], [5, 'A6', 5, 0, 6],
    ])
  })
})
