import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import {
  classifyFilterRuleTag,
  hasDeclaredFiltering,
  parseDeclaredRef,
  sortByOf,
  summarizeDeclaredFiltering,
} from '../src/xlsx/view/declaredFilters'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

async function parseFiles(files: Record<string, string>): Promise<XlsxDocument> {
  const zip = new JSZip()
  for (const [p, c] of Object.entries(files)) zip.file(p, c)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  return parseXlsx(await OfficePackage.load(bytes))
}

function sheetFiles(sheetXml: string, extra: Record<string, string> = {}): Record<string, string> {
  return {
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheetXml,
    ...extra,
  }
}

const sheetWrap = (inner: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2"><c r="A2"><v>4</v></c></row><row r="3"><c r="A3"><v>5</v></c></row></sheetData>${inner}</worksheet>`

describe('declared worksheet autoFilter', () => {
  it('parses ref and values filterColumn kinds as diagnostics only', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:C3"><filterColumn colId="0"><filters blank="1"><filter val="a"/><filter val="b"/></filters></filterColumn><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>`,
    )))
    const declared = doc.declaredFilters
    expect(declared).toHaveLength(1)
    const auto = declared![0].worksheetAutoFilter
    expect(auto?.refVerbatim).toBe('A1:C3')
    expect(auto?.range).toEqual({ minRow: 0, minCol: 0, maxRow: 2, maxCol: 2 })
    expect(auto?.columns).toHaveLength(2)
    expect(auto?.columns[0]).toMatchObject({ colId: 0, rule: { kind: 'values', valueCount: 2, hasBlank: true } })
    expect(auto?.columns[1]).toMatchObject({ colId: 2, rule: { kind: 'values', valueCount: 1, hasBlank: false } })
    // Diagnostics only: stored row order untouched, no viewer rule state.
    expect(doc.sheets[0].rows.map(row => row.index)).toEqual([0, 1, 2])
    expect((doc.sheets[0] as unknown as Record<string, unknown>).filterSpec).toBeUndefined()
  })

  it('classifies unsupported color, top-10, dynamic and custom rule kinds', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:D4"><filterColumn colId="0"><colorFilter dxfId="1" cellColor="0"/></filterColumn><filterColumn colId="1"><top10 top="1" percent="0" val="10" filterVal="5"/></filterColumn><filterColumn colId="2"><dynamicFilter type="aboveAverage"/></filterColumn><filterColumn colId="3"><customFilters and="1"><customFilter operator="greaterThan" val="3"/><customFilter operator="lessThan" val="9"/></customFilters></filterColumn></autoFilter>`,
    )))
    const columns = doc.declaredFilters![0].worksheetAutoFilter!.columns
    expect(columns[0].rule).toMatchObject({ kind: 'color', dxfId: '1', cellColor: '0' })
    expect(columns[1].rule).toMatchObject({ kind: 'top10', top: true, percent: false, val: '10', filterVal: '5' })
    expect(columns[2].rule).toMatchObject({ kind: 'dynamic', filterType: 'aboveAverage' })
    expect(columns[3].rule).toMatchObject({
      kind: 'custom', operator: 'and',
      conditions: [{ operator: 'greaterThan', value: '3' }, { operator: 'lessThan', value: '9' }],
    })
  })

  it('records malformed declarations without throwing', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter><filterColumn><mysteryFilter/></filterColumn><filterColumn colId="7"/></autoFilter><sortState><sortCondition sortBy="fancy"/></sortState>`,
    )))
    const entry = doc.declaredFilters![0]
    expect(entry.worksheetAutoFilter?.refVerbatim).toBeUndefined()
    expect(entry.worksheetAutoFilter?.malformed).toBe('missing-ref')
    expect(entry.worksheetAutoFilter?.columns[0]).toMatchObject({ colId: undefined, malformed: 'invalid-colId' })
    expect(entry.worksheetAutoFilter?.columns[0].rule).toMatchObject({ kind: 'unknown' })
    expect(entry.worksheetAutoFilter?.columns[1].rule).toBeUndefined()
    expect(entry.worksheetSortState?.malformed).toBe('missing-ref')
    expect(entry.worksheetSortState?.conditions[0]).toMatchObject({ sortBy: 'unknown', sortByVerbatim: 'fancy' })
  })

  it('rejects out-of-grid and reversed refs as malformed verbatim', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A0:XFE3"/>`)))
    const auto = doc.declaredFilters![0].worksheetAutoFilter!
    expect(auto.refVerbatim).toBe('A0:XFE3')
    expect(auto.range).toBeUndefined()
    expect(auto.malformed).toBe('invalid-ref')
  })
})

describe('declared sortState', () => {
  it('is diagnostics-only and never reorders stored rows', async () => {
    const doc = await parseFiles(sheetFiles(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>30</v></c><c r="B1"><v>b</v></c></row><row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>a</v></c></row><row r="3"><c r="A3"><v>20</v></c><c r="B3"><v>c</v></c></row></sheetData><sortState ref="A1:B3"><sortCondition ref="A1:A3" descending="0"/><sortCondition ref="B1:B3" sortBy="cellColor" dxfId="2" descending="1"/><sortCondition ref="A1:A3" customList="b,a,c"/></sortState></worksheet>`,
    ))
    // Stored order preserved verbatim despite the ascending sortState.
    expect(doc.sheets[0].rows.map(row => row.cells[0].value)).toEqual([30, 10, 20])
    const sort = doc.declaredFilters![0].worksheetSortState!
    expect(sort.refVerbatim).toBe('A1:B3')
    expect(sort.conditions).toHaveLength(3)
    expect(sort.conditions[0]).toMatchObject({ sortBy: 'value', descending: false })
    expect(sort.conditions[1]).toMatchObject({ sortBy: 'cellColor', dxfId: '2', descending: true })
    expect(sort.conditions[2]).toMatchObject({ sortBy: 'value', customListCount: 3 })
  })
})

describe('declared table autoFilter', () => {
  const tableFiles = (tableXml: string, sheetExtra = ''): Record<string, string> => sheetFiles(
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${sheetExtra}<tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
    {
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
      'xl/tables/table1.xml': tableXml,
    },
  )
  const tableWrap = (inner: string): string =>
    `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="1" name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1" totalsRowShown="0">${inner}<tableColumns count="2"><tableColumn id="1" name="ColA"/><tableColumn id="2" name="ColB"/></tableColumns></table>`

  it('diagnoses table autoFilter with table attribution and keeps table metadata', async () => {
    const doc = await parseFiles(tableFiles(tableWrap(
      `<autoFilter ref="A1:B3"><filterColumn colId="1"><top10 top="0" percent="1" val="20"/></filterColumn></autoFilter><sortState ref="A1:B3"><sortCondition ref="A1:A3"/></sortState>`,
    )))
    expect(doc.tables?.[0].displayName).toBe('Table1')
    const tables = doc.declaredFilters![0].tableAutoFilters
    expect(tables).toHaveLength(1)
    expect(tables[0].table).toMatchObject({ displayName: 'Table1', partPath: 'xl/tables/table1.xml' })
    expect(tables[0].autoFilter.refVerbatim).toBe('A1:B3')
    expect(tables[0].autoFilter.columns[0].rule).toMatchObject({ kind: 'top10', top: false, percent: true, val: '20' })
    expect(tables[0].sortState?.conditions[0]).toMatchObject({ sortBy: 'value' })
  })

  it('leaves no declared table block when the table part is unreadable', async () => {
    const files = tableFiles(tableWrap(`<autoFilter ref="A1:B3"/>`))
    files['xl/tables/table1.xml'] = 'not xml at all <<'
    const doc = await parseFiles(files)
    expect(doc.tables).toBeUndefined()
    expect(doc.declaredFilters?.[0]?.tableAutoFilters ?? []).toEqual([])
  })
})

describe('file-hidden snapshot and Clear support', () => {
  it('preserves file-hidden rows verbatim and exposes the snapshot explanation', async () => {
    const doc = await parseFiles(sheetFiles(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>h</v></c></row><row r="2" hidden="1"><c r="A2"><v>gone</v></c></row><row r="3"><c r="A3"><v>kept</v></c></row></sheetData><autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="h"/><filter val="kept"/></filters></filterColumn></autoFilter></worksheet>`,
    ))
    expect(doc.sheets[0].rows.map(row => row.hidden ?? false)).toEqual([false, true, false])
    const entry = doc.declaredFilters![0]
    expect(hasDeclaredFiltering(entry)).toBe(true)
    const summary = summarizeDeclaredFiltering(entry)
    expect(summary).toContain('Data')
    expect(summary).toContain('A1:A3')
    expect(summary).toMatch(/hidden rows are part of the file snapshot/i)
  })

  it('reports no declared filtering for plain sheets', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap('')))
    expect(doc.declaredFilters).toBeUndefined()
    expect(hasDeclaredFiltering(undefined)).toBe(false)
  })
})

describe('adapter-neutral classifiers', () => {
  it('maps filter rule element names to kinds', () => {
    expect(classifyFilterRuleTag('filters')).toBe('filters')
    expect(classifyFilterRuleTag('top10')).toBe('top10')
    expect(classifyFilterRuleTag('customFilters')).toBe('customFilters')
    expect(classifyFilterRuleTag('dynamicFilter')).toBe('dynamicFilter')
    expect(classifyFilterRuleTag('colorFilter')).toBe('colorFilter')
    expect(classifyFilterRuleTag('mysteryFilter')).toBe('unknown')
  })

  it('parses declared refs strictly', () => {
    expect(parseDeclaredRef('A1:C10')).toEqual({ minRow: 0, minCol: 0, maxRow: 9, maxCol: 2 })
    expect(parseDeclaredRef('B2')).toEqual({ minRow: 1, minCol: 1, maxRow: 1, maxCol: 1 })
    expect(parseDeclaredRef('A1:XFD1048576')).toEqual({ minRow: 0, minCol: 0, maxRow: 1048575, maxCol: 16383 })
    expect(parseDeclaredRef('A0:B2')).toBeUndefined()
    expect(parseDeclaredRef('A1:XFE1')).toBeUndefined()
    expect(parseDeclaredRef('C3:A1')).toBeUndefined()
    expect(parseDeclaredRef('a1:c10')).toBeUndefined()
    expect(parseDeclaredRef('')).toBeUndefined()
  })

  it('maps sortBy spellings', () => {
    expect(sortByOf('value')).toBe('value')
    expect(sortByOf('cellColor')).toBe('cellColor')
    expect(sortByOf('fontColor')).toBe('fontColor')
    expect(sortByOf('icon')).toBe('icon')
    expect(sortByOf(undefined)).toBe('value')
    expect(sortByOf('fancy')).toBe('unknown')
  })
})
