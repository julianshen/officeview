import { describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import * as tables from '../src/xlsx/formula/tables'

// Counts root-QName checks while keeping the real implementation.
vi.mock('../src/xlsx/formula/tables', async importOriginal => {
  const actual = await importOriginal<typeof import('../src/xlsx/formula/tables')>()
  return { ...actual, isTablePartSource: vi.fn(actual.isTablePartSource) }
})

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

async function parseFiles(files: Record<string, string>) {
  const zip = new JSZip()
  for (const [p, c] of Object.entries(files)) zip.file(p, c)
  return parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}

describe('table part source check', () => {
  it('runs once per part path for repeated tablePart refs, with identical table metadata', async () => {
    const parts = Array.from({ length: 40 }, (_, i) => `<tablePart r:id="rIdT${i % 2 === 0 ? 1 : 2}"/>`).join('')
    const table = (id: number) => `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="${id}" name="T${id}" displayName="T${id}" ref="A1:B3" headerRowCount="1" totalsRowShown="0"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`
    const doc = await parseFiles({
      'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts>${parts}</tableParts></worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/><Relationship Id="rIdT2" Type="${NS_R}/table" Target="../tables/table2.xml"/></Relationships>`,
      'xl/tables/table1.xml': table(1),
      'xl/tables/table2.xml': table(2),
    })
    expect(vi.mocked(tables.isTablePartSource)).toHaveBeenCalledTimes(2)
    // Inventory unchanged: one metadata entry per reference, in reference order.
    expect(doc.tables).toHaveLength(40)
    expect(doc.tables!.map(t => t.displayName)).toEqual(Array.from({ length: 40 }, (_, i) => (i % 2 === 0 ? 'T1' : 'T2')))
    expect(doc.tables![2]).toEqual(doc.tables![0])
  })
})
