import { describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDocument } from '../src/xlsx/types'
import {
  classifyFilterRuleTag,
  declaredAutoFilterBuilder,
  declaredSortStateBuilder,
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
    // Pin the clock so the volatile-function epoch matches across both parses.
    vi.useFakeTimers({ toFake: ['Date'], now: Date.UTC(2026, 0, 1) })
    let doc: XlsxDocument, plain: XlsxDocument
    try {
      doc = await parseFiles(sheetFiles(sheetWrap(
        `<autoFilter ref="A1:C3"><filterColumn colId="0"><filters blank="1"><filter val="a"/><filter val="b"/></filters></filterColumn><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>`,
      )))
      plain = await parseFiles(sheetFiles(sheetWrap('')))
    } finally {
      vi.useRealTimers()
    }
    const declared = doc.declaredFilters
    expect(declared).toHaveLength(1)
    const auto = declared![0].worksheetAutoFilter
    expect(auto?.refVerbatim).toBe('A1:C3')
    expect(auto?.range).toEqual({ minRow: 0, minCol: 0, maxRow: 2, maxCol: 2 })
    expect(auto?.columns).toHaveLength(2)
    expect(auto?.columns[0]).toMatchObject({ colId: 0, rule: { kind: 'values', valueCount: 2, hasBlank: true } })
    expect(auto?.columns[1]).toMatchObject({ colId: 2, rule: { kind: 'values', valueCount: 1, hasBlank: false } })
    // Diagnostics only: nothing outside declaredFilters differs from the same
    // sheet without the autoFilter (no viewer rule state, no row changes).
    const { declaredFilters: _declared, ...rest } = doc
    expect(rest).toEqual(plain)
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
    expect(tables[0].autoFilter?.refVerbatim).toBe('A1:B3')
    expect(tables[0].autoFilter?.columns[0].rule).toMatchObject({ kind: 'top10', top: false, percent: true, val: '20' })
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

describe('codex round-2 findings', () => {
  it('F1 keeps sortState nested inside autoFilter with attribution', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:C3"><filterColumn colId="0"><filters><filter val="a"/></filters></filterColumn><sortState ref="A1:C3"><sortCondition ref="B1:B3" descending="1"/></sortState></autoFilter>`,
    )))
    const sort = doc.declaredFilters![0].worksheetSortState!
    expect(sort.refVerbatim).toBe('A1:C3')
    expect(sort.placement).toBe('nested')
    expect(sort.conditions).toMatchObject([{ sortBy: 'value', descending: true }])
  })

  it('F1 direct sortState wins over nested with a duplicate flag', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:C3"><sortState ref="A1:C3"><sortCondition ref="B1:B3"/></sortState></autoFilter><sortState ref="A1:C3"><sortCondition ref="A1:A3"/></sortState>`,
    )))
    const sort = doc.declaredFilters![0].worksheetSortState!
    expect(sort.placement).toBe('direct')
    expect(sort.conditions).toHaveLength(1)
    expect(sort.conditions[0].refVerbatim).toBe('A1:A3')
    expect(sort.malformed).toBe('duplicate-sortState')
  })

  it('F2 retains sort-only tables in diagnostics', async () => {
    const files = sheetFiles(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
      {
        'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
        'xl/tables/table1.xml': `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="1" name="T" displayName="T" ref="A1:B2" headerRowCount="1" totalsRowShown="0"><sortState ref="A1:B2"><sortCondition ref="A1:A2"/></sortState><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`,
      },
    )
    const doc = await parseFiles(files)
    expect(doc.tables?.[0].displayName).toBe('T')
    const tables = doc.declaredFilters![0].tableAutoFilters
    expect(tables).toHaveLength(1)
    expect(tables[0].autoFilter).toBeUndefined()
    expect(tables[0].sortState?.refVerbatim).toBe('A1:B2')
  })

  it('F3 preserves worksheet diagnostics when table rels fail', async () => {
    const files = sheetFiles(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1"><c r="A2"><v>2</v></c></row></sheetData><autoFilter ref="A1:A2"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
      { 'xl/worksheets/_rels/sheet1.xml.rels': 'not xml at all <<' },
    )
    const doc = await parseFiles(files)
    expect(doc.tables).toBeUndefined()
    expect(doc.sheets[0].rows[1].hidden).toBe(true)
    expect(doc.declaredFilters![0].worksheetAutoFilter?.refVerbatim).toBe('A1:A2')
  })

  it('F4 bounds columns, conditions and summaries with omission counts', async () => {
    const cols = Array.from({ length: 3000 }, (_, i) => `<filterColumn colId="${i}"/>`).join('')
    const conds = Array.from({ length: 1500 }, () => `<sortCondition ref="A1:A3"/>`).join('')
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:C3">${cols}</autoFilter><sortState ref="A1:C3">${conds}</sortState>`)))
    const entry = doc.declaredFilters![0]
    expect(entry.worksheetAutoFilter!.columns).toHaveLength(1024)
    expect(entry.worksheetAutoFilter!.omittedColumns).toBe(1976)
    expect(entry.worksheetSortState!.conditions).toHaveLength(1024)
    expect(entry.worksheetSortState!.omittedConditions).toBe(476)
    expect(summarizeDeclaredFiltering(entry).length).toBeLessThan(4096)
    expect(summarizeDeclaredFiltering(entry)).toContain('more')
  })

  it('F5 ignores extLst metadata when collecting rule alternatives', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:B2"><filterColumn colId="0"><filters><filter val="a"/></filters><extLst><ext><x:y xmlns:x="urn:x"/></ext></extLst></filterColumn><filterColumn colId="1"><extLst><ext/></extLst></filterColumn></autoFilter>`,
    )))
    const columns = doc.declaredFilters![0].worksheetAutoFilter!.columns
    expect(columns[0].rule).toMatchObject({ kind: 'values', valueCount: 1 })
    expect(columns[0].malformed).toBeUndefined()
    expect(columns[1].rule).toBeUndefined()
    expect(columns[1].malformed).toBeUndefined()
  })

  it('keeps sortState nested inside a table autoFilter', async () => {
    const files = sheetFiles(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
      {
        'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
        'xl/tables/table1.xml': `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="1" name="T" displayName="T" ref="A1:B2" headerRowCount="1" totalsRowShown="0"><autoFilter ref="A1:B2"><sortState ref="A1:B2"><sortCondition ref="B1:B2" descending="1"/></sortState></autoFilter><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`,
      },
    )
    const doc = await parseFiles(files)
    const tables = doc.declaredFilters![0].tableAutoFilters
    expect(tables).toHaveLength(1)
    expect(tables[0].autoFilter?.refVerbatim).toBe('A1:B2')
    expect(tables[0].sortState?.placement).toBe('nested')
    expect(tables[0].sortState?.conditions).toMatchObject([{ sortBy: 'value', descending: true }])
  })

  it('F6 validates colId spelling, range and duplicates without repair', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:B3"><filterColumn colId="0junk"/><filterColumn colId="1.5"/><filterColumn colId="5"/><filterColumn colId="0"/><filterColumn colId="0"/></autoFilter>`,
    )))
    const columns = doc.declaredFilters![0].worksheetAutoFilter!.columns
    expect(columns[0]).toMatchObject({ colId: undefined, malformed: 'invalid-colId' })
    expect(columns[1]).toMatchObject({ colId: undefined, malformed: 'invalid-colId' })
    expect(columns[2]).toMatchObject({ colId: 5, malformed: 'colId-out-of-range' })
    expect(columns[3]).toMatchObject({ colId: 0 })
    expect(columns[3].malformed).toBeUndefined()
    expect(columns[4]).toMatchObject({ colId: 0, malformed: 'duplicate-colId' })
  })
})

describe('adapter-neutral builders', () => {
  it('produce the same diagnostics from per-element input as the tree adapter', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(
      `<autoFilter ref="A1:C3"><filterColumn colId="0" hiddenButton="1"><customFilters and="1"><customFilter operator="greaterThan" val="3"/><customFilter val="x"/></customFilters></filterColumn><filterColumn colId="0"><filters blank="1"><filter val="a"/><dateGroupItem year="2020" dateTimeGrouping="year"/></filters></filterColumn><filterColumn colId="x"><colorFilter dxfId="2"/></filterColumn></autoFilter><sortState ref="A1:C3" caseSensitive="1"><sortCondition ref="A1:A3" descending="1" customList="a,b"/><sortCondition sortBy="icon"/></sortState>`,
    )))
    const auto = declaredAutoFilterBuilder({ ref: 'A1:C3' })
    auto.addColumn({
      attrs: { colId: '0', hiddenButton: '1' }, ruleCount: 1,
      firstRule: { tag: 'customFilters', attrs: { and: '1' }, filterCount: 0, dateGroupItemCount: 0, customFilters: [{ operator: 'greaterThan', val: '3' }, { val: 'x' }], customFilterCount: 2 },
    })
    auto.addColumn({
      attrs: { colId: '0' }, ruleCount: 1,
      firstRule: { tag: 'filters', attrs: { blank: '1' }, filterCount: 1, dateGroupItemCount: 1, customFilters: [], customFilterCount: 0 },
    })
    auto.addColumn({
      attrs: { colId: 'x' }, ruleCount: 1,
      firstRule: { tag: 'colorFilter', attrs: { dxfId: '2' }, filterCount: 0, dateGroupItemCount: 0, customFilters: [], customFilterCount: 0 },
    })
    const sort = declaredSortStateBuilder({ ref: 'A1:C3', caseSensitive: '1' }, 'direct')
    sort.addCondition({ ref: 'A1:A3', descending: '1', customList: 'a,b' })
    sort.addCondition({ sortBy: 'icon' })
    const entry = doc.declaredFilters![0]
    expect(auto.finish()).toEqual(entry.worksheetAutoFilter)
    expect(sort.finish()).toEqual(entry.worksheetSortState)
  })
})

describe('claude round-2 findings', () => {
  const NS = `xmlns="${NS_MAIN}" xmlns:r="${NS_R}"`
  const relsFor = (ids: number[]): string =>
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}">${ids.map(i => `<Relationship Id="rIdT${i}" Type="${NS_R}/table" Target="../tables/table${i}.xml"/>`).join('')}</Relationships>`
  const tableXml = (i: number, inner: string, name = `T${i}`): string =>
    `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="${i}" name="T${i}" displayName="${name}" ref="A1:B3" headerRowCount="1" totalsRowShown="0">${inner}<tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`
  const withTables = (parts: string, rels: string, tables: Record<string, string>, sheetInner = ''): Record<string, string> => sheetFiles(
    `<?xml version="1.0" encoding="UTF-8"?><worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${sheetInner}<tableParts>${parts}</tableParts></worksheet>`,
    { 'xl/worksheets/_rels/sheet1.xml.rels': rels, ...tables },
  )
  const cols = (n: number): string => Array.from({ length: n }, (_, i) => `<filterColumn colId="${i}"><filters><filter val="a"/></filters></filterColumn>`).join('')

  it('1A de-duplicates repeated tablePart refs by part path without changing table metadata', async () => {
    const parts = Array.from({ length: 40 }, () => '<tablePart r:id="rIdT1"/>').join('')
    const doc = await parseFiles(withTables(parts, relsFor([1]), {
      'xl/tables/table1.xml': tableXml(1, `<autoFilter ref="A1:B3">${cols(2)}</autoFilter>`),
    }))
    const entry = doc.declaredFilters![0]
    expect(entry.tableAutoFilters).toHaveLength(1)
    expect(entry.duplicateTableParts).toBe(39)
    // Table inventory behavior is not the diagnostics' business.
    expect(doc.tables).toHaveLength(40)
  })

  it('1A caps declared tables per sheet with an omission count', async () => {
    const ids = Array.from({ length: 67 }, (_, i) => i + 1)
    const tables: Record<string, string> = {}
    for (const i of ids) tables[`xl/tables/table${i}.xml`] = tableXml(i, '<autoFilter ref="A1:B3"/>')
    const doc = await parseFiles(withTables(ids.map(i => `<tablePart r:id="rIdT${i}"/>`).join(''), relsFor(ids), tables))
    const entry = doc.declaredFilters![0]
    expect(entry.tableAutoFilters).toHaveLength(64)
    expect(entry.omittedTables).toBe(3)
    const summary = summarizeDeclaredFiltering(entry)
    expect(summary).toContain('and 59 more table declarations')
  })

  it('1A bounds retained column and condition records per document', async () => {
    const ids = Array.from({ length: 9 }, (_, i) => i + 1)
    const tables: Record<string, string> = {}
    for (const i of ids) tables[`xl/tables/table${i}.xml`] = tableXml(i, `<autoFilter ref="A1:AMJ3">${cols(1024)}</autoFilter>`)
    const doc = await parseFiles(withTables(ids.map(i => `<tablePart r:id="rIdT${i}"/>`).join(''), relsFor(ids), tables))
    const entry = doc.declaredFilters![0]
    const retained = entry.tableAutoFilters.reduce((sum, t) => sum + t.autoFilter!.columns.length, 0)
    const omitted = entry.tableAutoFilters.reduce((sum, t) => sum + (t.autoFilter!.omittedColumns ?? 0), 0)
    expect(retained).toBe(8192)
    expect(retained + omitted).toBe(9 * 1024)
  })

  it('1B truncates retained text with a flag and bounds the summary', async () => {
    const big = 'A'.repeat(200000)
    const doc = await parseFiles(withTables('<tablePart r:id="rIdT1"/>', relsFor([1]), {
      'xl/tables/table1.xml': tableXml(1, `<autoFilter ref="${big}"/>`, `N${'x'.repeat(200000)}`),
    }, `<autoFilter ref="${big}"><filterColumn colId="0"><customFilters><customFilter operator="${'o'.repeat(5000)}" val="${'v'.repeat(5000)}"/></customFilters></filterColumn><filterColumn colId="1"><top10 val="${'9'.repeat(5000)}"/></filterColumn></autoFilter><sortState ref="${big}" sortMethod="${'s'.repeat(5000)}"><sortCondition ref="${big}" sortBy="${'q'.repeat(5000)}" dxfId="${'7'.repeat(5000)}"/></sortState>`))
    const entry = doc.declaredFilters![0]
    const auto = entry.worksheetAutoFilter!
    expect(auto.malformed).toBe('invalid-ref')
    expect(auto.refVerbatim!.length).toBeLessThanOrEqual(256)
    expect(auto.truncated).toBe(true)
    const custom = auto.columns[0].rule as { conditions: { operator: string; value: string }[]; truncated?: boolean }
    expect(custom.conditions[0].value.length).toBeLessThanOrEqual(256)
    expect(custom.conditions[0].operator.length).toBeLessThanOrEqual(256)
    expect(custom.truncated).toBe(true)
    expect((auto.columns[1].rule as { val?: string }).val!.length).toBeLessThanOrEqual(256)
    const sort = entry.worksheetSortState!
    expect(sort.refVerbatim!.length).toBeLessThanOrEqual(256)
    expect(sort.sortMethod!.length).toBeLessThanOrEqual(256)
    expect(sort.truncated).toBe(true)
    const condition = sort.conditions[0]
    expect(condition.sortByVerbatim!.length).toBeLessThanOrEqual(256)
    expect(condition.refVerbatim!.length).toBeLessThanOrEqual(256)
    expect(condition.dxfId!.length).toBeLessThanOrEqual(256)
    expect(condition.truncated).toBe(true)
    const table = entry.tableAutoFilters[0]
    expect(table.table.displayName!.length).toBeLessThanOrEqual(256)
    expect(table.table.truncated).toBe(true)
    expect(table.autoFilter!.refVerbatim!.length).toBeLessThanOrEqual(256)
    expect(summarizeDeclaredFiltering(entry).length).toBeLessThan(2048)
    expect(JSON.stringify(entry).length).toBeLessThan(8192)
  })

  it('1B parses the ref from the full spelling, never the clipped one', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:B3${'9'.repeat(400)}"/>`)))
    const auto = doc.declaredFilters![0].worksheetAutoFilter!
    expect(auto.range).toBeUndefined()
    expect(auto.malformed).toBe('invalid-ref')
  })

  it('2 keeps the first column flag and records every flag', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:B3"><filterColumn colId="x"><filters/><top10/></filterColumn></autoFilter>`)))
    const column = doc.declaredFilters![0].worksheetAutoFilter!.columns[0]
    expect(column.malformed).toBe('invalid-colId')
    expect(column.flags).toEqual(['invalid-colId', 'multiple-rule-elements'])
  })

  it('3 describes sort-only tables and unreadable filters distinctly', async () => {
    const doc = await parseFiles(withTables('<tablePart r:id="rIdT1"/><tablePart r:id="rIdT2"/>', relsFor([1, 2]), {
      'xl/tables/table1.xml': tableXml(1, '<sortState ref="A2:B3"><sortCondition ref="A2:A3"/></sortState>'),
      'xl/tables/table2.xml': tableXml(2, '<autoFilter ref="bad"/>'),
    }))
    const summary = summarizeDeclaredFiltering(doc.declaredFilters![0])
    expect(summary).toContain('table "T1" has a stored sort on A2:B3')
    expect(summary).not.toMatch(/T1" declares a filter/)
    expect(summary).toContain('table "T2" has a file filter on an unreadable range "bad"')
  })

  it('3 describes nested worksheet sorts and sorts without a ref', async () => {
    const nested = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:B3"><sortState ref="A2:B3"><sortCondition ref="B2:B3"/></sortState></autoFilter>`)))
    expect(summarizeDeclaredFiltering(nested.declaredFilters![0])).toContain('a stored sort on A2:B3')
    const noRef = await parseFiles(sheetFiles(sheetWrap(`<sortState><sortCondition ref="A1:A3"/></sortState>`)))
    expect(summarizeDeclaredFiltering(noRef.declaredFilters![0])).toContain('a stored sort the viewer could not read')
    const badRef = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="$A$1:$B$3"/>`)))
    expect(summarizeDeclaredFiltering(badRef.declaredFilters![0])).toContain('a file filter on an unreadable range "$A$1:$B$3"')
  })

  it('4 flags a second direct autoFilter or sortState', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:A3"/><autoFilter ref="B1:B3"><filterColumn colId="0"><colorFilter/></filterColumn></autoFilter><sortState ref="A1:A3"/><sortState ref="B1:B3"/>`)))
    const entry = doc.declaredFilters![0]
    expect(entry.worksheetAutoFilter!.refVerbatim).toBe('A1:A3')
    expect(entry.worksheetAutoFilter!.malformed).toBe('duplicate-autoFilter')
    expect(entry.worksheetSortState!.refVerbatim).toBe('A1:A3')
    expect(entry.worksheetSortState!.malformed).toBe('duplicate-sortState')
  })

  it('4 never loses the duplicate flag when the kept declaration is malformed', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter/><autoFilter ref="B1:B3"/><sortState ref="bad"/><sortState ref="B1:B3"/>`)))
    const entry = doc.declaredFilters![0]
    expect(entry.worksheetAutoFilter!.malformed).toBe('missing-ref')
    expect(entry.worksheetAutoFilter!.flags).toEqual(['missing-ref', 'duplicate-autoFilter'])
    expect(entry.worksheetSortState!.malformed).toBe('invalid-ref')
    expect(entry.worksheetSortState!.flags).toEqual(['invalid-ref', 'duplicate-sortState'])
  })

  it('4 flags duplicate table autoFilter and sortState', async () => {
    const doc = await parseFiles(withTables('<tablePart r:id="rIdT1"/>', relsFor([1]), {
      'xl/tables/table1.xml': tableXml(1, '<autoFilter ref="A1:B3"/><autoFilter ref="A1:A3"/><sortState ref="A1:B3"/><sortState ref="A1:A3"/>'),
    }))
    const table = doc.declaredFilters![0].tableAutoFilters[0]
    expect(table.autoFilter!.flags).toEqual(['duplicate-autoFilter'])
    expect(table.sortState!.flags).toEqual(['duplicate-sortState'])
  })

  it('5 records columnSort as an unsupported kind and sortMethod verbatim', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<sortState ref="A1:C3" columnSort="1" sortMethod="stroke"><sortCondition ref="A1:C1"/></sortState>`)))
    const sort = doc.declaredFilters![0].worksheetSortState!
    expect(sort).toMatchObject({ columnSort: true, sortMethod: 'stroke', unsupported: 'left-to-right' })
    expect(summarizeDeclaredFiltering(doc.declaredFilters![0])).toContain('a stored left-to-right sort on A1:C3 (unsupported)')
    const rows = await parseFiles(sheetFiles(sheetWrap(`<sortState ref="A1:C3" columnSort="0"><sortCondition ref="A1:A3"/></sortState>`)))
    const rowSort = rows.declaredFilters![0].worksheetSortState!
    expect(rowSort.columnSort).toBe(false)
    expect(rowSort.unsupported).toBeUndefined()
  })

  it('7 counts only columns with a rule as filtered', async () => {
    const doc = await parseFiles(sheetFiles(sheetWrap(`<autoFilter ref="A1:B3"><filterColumn colId="0" hiddenButton="1"/><filterColumn colId="1"><top10 val="3"/></filterColumn></autoFilter>`)))
    const summary = summarizeDeclaredFiltering(doc.declaredFilters![0])
    expect(summary).toContain('(1 filtered column: Top 10)')
    expect(summary).not.toContain('unfiltered')
  })
})
