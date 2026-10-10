/**
 * B0 workbook metadata extraction regression tests (implemented r1–r6).
 *
 * Scope: workbook order + defined names, worksheet tableParts discovery,
 * calc/date/unicode settings, ordinary value preservation, and table
 * inventory completeness. Uses public parseXlsx / OfficePackage paths with
 * inline synthetic OOXML ZIP fixtures and repository-relative imports.
 * Synthetic extraction controls only; no native calculation claims.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXml } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'
import { parseTableMetadata } from '../src/xlsx/formula/tables'
import { evaluateFormula } from '../src/xlsx/formula/evaluator'
import type { XlsxDocument } from '../src/xlsx/types'

// Exact ABI inner shapes (reference-contract-final/ABI.md), local inspection only.
interface LocalSheetIdentity {
  sheetId: string
  workbookIndex: number
  name: string
  kind: 'worksheet' | 'other'
}
interface LocalDefinedName {
  name: string
  localSheetIndex?: number
  source: string
  relativeBase?: { sheetId: string; col: number; row: number }
  baseProvenance: 'explicit' | 'verified-file' | 'unknown'
}
interface LocalTableColumn {
  id: string
  name: string
  index: number
}
interface LocalTable {
  id: string
  name: string
  displayName: string
  sheetId: string
  partPath: string
  extent: { sheetId: string; firstCol: number; firstRow: number; cols: number; rows: number }
  headerRowCount: 0 | 1
  totalsRowCount: 0 | 1
  columns: readonly LocalTableColumn[]
}
interface LocalCalc {
  calcMode: 'auto' | 'manual' | 'autoNoTable'
  fullCalcOnLoad: boolean
  iterate: boolean
  iterateCount: number
  iterateDelta: number
}

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

async function parseFiles(files: Record<string, string>): Promise<XlsxDocument> {
  const zip = new JSZip()
  for (const [p, c] of Object.entries(files)) zip.file(p, c)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const pkg = await OfficePackage.load(bytes)
  return parseXlsx(pkg)
}

function asRecord(doc: XlsxDocument): Record<string, unknown> {
  return doc as unknown as Record<string, unknown>
}

function sheetIdentities(doc: XlsxDocument): LocalSheetIdentity[] | undefined {
  const r = asRecord(doc)
  if (Array.isArray(r.workbookSheets)) return r.workbookSheets as LocalSheetIdentity[]
  const sheets = (doc.sheets ?? []) as unknown as Array<Record<string, unknown>>
  const hasAny = sheets.some(
    (s) => 'workbookIndex' in s || 'sheetId' in s || 'kind' in s,
  )
  if (hasAny) return sheets as unknown as LocalSheetIdentity[]
  return undefined
}

function definedNamesOf(doc: XlsxDocument): LocalDefinedName[] | undefined {
  const r = asRecord(doc)
  if (Array.isArray(r.definedNames)) return r.definedNames as LocalDefinedName[]
  if (Array.isArray(r.names)) return r.names as LocalDefinedName[]
  return undefined
}

function tablesOf(doc: XlsxDocument): LocalTable[] | undefined {
  const r = asRecord(doc)
  if (Array.isArray(r.tables)) return r.tables as LocalTable[]
  if (Array.isArray(r.workbookTables)) return r.workbookTables as LocalTable[]
  return undefined
}

function calcOf(doc: XlsxDocument): LocalCalc | undefined {
  const r = asRecord(doc)
  for (const k of ['calc', 'calcSettings', 'calculation', 'calcPr']) {
    const v = r[k]
    if (v && typeof v === 'object') return v as LocalCalc
  }
  return undefined
}

function semanticsOf(
  doc: XlsxDocument,
): { dateSystem?: unknown; unicode?: unknown } | undefined {
  const r = asRecord(doc)
  for (const k of ['semantics', 'resolvedSemantics', 'workbookSemantics']) {
    const v = r[k]
    if (v && typeof v === 'object') return v as { dateSystem?: unknown; unicode?: unknown }
  }
  if ('dateSystem' in r || 'date1904' in r || 'unicode' in r) {
    return {
      dateSystem: r.dateSystem,
      unicode: (r as Record<string, unknown>).unicode,
    }
  }
  return undefined
}

function namesWorkbookFiles(opts: {
  calcPr: string
  workbookPr: string
  withProducer?: boolean
}): Record<string, string> {
  const files: Record<string, string> = {
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${opts.workbookPr}<sheets><sheet name="Chart1" sheetId="1" r:id="rIdChart"/><sheet name="Data" sheetId="2" r:id="rIdData"/></sheets><definedNames><definedName name="MyName" localSheetId="1">Data!$A$1</definedName><definedName name="MyName">Data!$B$2</definedName><definedName name="Plain">1+1</definedName></definedNames>${opts.calcPr}</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdChart" Type="${NS_R}/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rIdData" Type="${NS_R}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/chartsheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><chartsheet xmlns="${NS_MAIN}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>`,
    'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="B2"><v>7</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
    'xl/worksheets/_rels/sheet2.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
    'xl/tables/table1.xml': `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="1" name="Table1" displayName="Table1" ref="A1:B3" headerRowCount="1" totalsRowShown="0"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="ColA"/><tableColumn id="2" name="ColB"/></tableColumns><tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/></table>`,
  }
  if (opts.withProducer) {
    files['docProps/app.xml'] =
      `<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application><AppVersion>16.0000</AppVersion></Properties>`
  }
  return files
}

function tableOnlyFiles(opts: {
  ref: string
  headerRowCount: string
  totalsRowShown: string
  totalsRowCount?: string
  id?: string
}): Record<string, string> {
  const id = opts.id ?? '1'
  const totalsCountAttr =
    opts.totalsRowCount !== undefined ? ` totalsRowCount="${opts.totalsRowCount}"` : ''
  return {
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rIdT1" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
    'xl/tables/table1.xml': `<?xml version="1.0" encoding="UTF-8"?><table xmlns="${NS_MAIN}" id="${id}" name="Table1" displayName="Table1" ref="${opts.ref}" headerRowCount="${opts.headerRowCount}" totalsRowShown="${opts.totalsRowShown}"${totalsCountAttr}><autoFilter ref="${opts.ref}"/><tableColumns count="2"><tableColumn id="1" name="ColA"/><tableColumn id="2" name="ColB"/></tableColumns></table>`,
  }
}

describe('B0 prep: workbook order with non-rendered sheet + defined-name scope', () => {
  it('renders the worksheet Data without requiring the non-rendered sheet', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    // Original Chart1/Data order belongs to WorkbookSheetIdentity metadata,
    // not to rendered doc.sheets; do not require Chart1 here.
    expect(doc.sheets.map((s) => s.name)).toContain('Data')
  })

  it('retains sheet identities in original order with stable IDs and kind', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    const ids = sheetIdentities(doc)
    expect(ids, 'workbook sheet identities in original order').toBeDefined()
    if (!ids) return
    expect(ids.map((s) => s.name)).toEqual(['Chart1', 'Data'])
    expect(ids.map((s) => s.workbookIndex)).toEqual([0, 1])
    expect(ids.map((s) => String(s.sheetId))).toEqual(['1', '2'])
    expect(ids[0]?.kind).toBe('other')
    expect(ids[1]?.kind).toBe('worksheet')
    // Rendered Data keeps its stable identity (sheetId 2 / workbookIndex 1).
    const dataMeta = ids.find((s) => s.name === 'Data')
    expect(dataMeta?.workbookIndex).toBe(1)
    expect(String(dataMeta?.sheetId)).toBe('2')
    const renderedData = doc.sheets.find((s) => s.name === 'Data')
    expect(renderedData).toBeDefined()
    const renderedId = (renderedData as unknown as Record<string, unknown>).sheetId
    if (renderedId !== undefined) expect(String(renderedId)).toBe('2')
  })

  it('retains local/global defined names distinctly with source formula and provenance', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    const names = definedNamesOf(doc)
    expect(names, 'defined-name metadata retained').toBeDefined()
    if (!names) return
    const local = names.filter(
      (n) => n.name === 'MyName' && n.localSheetIndex === 1,
    )
    const global = names.filter(
      (n) => n.name === 'MyName' && n.localSheetIndex === undefined,
    )
    expect(local.length).toBe(1)
    expect(global.length).toBe(1)
    expect(local[0]?.source).toBe('Data!$A$1')
    expect(global[0]?.source).toBe('Data!$B$2')
    // Same spelling stays distinct; binding uses original order/stable IDs.
    expect(local[0]?.name).toBe(global[0]?.name)
    // No relative base is invented when the definition carries none.
    const plain = names.find((n) => n.name === 'Plain')
    expect(plain?.source).toBe('1+1')
    expect(plain?.relativeBase).toBeUndefined()
    expect(plain?.baseProvenance).toBe('unknown')
  })
})

describe('B0 prep: table discovery through worksheet tableParts + worksheet rels', () => {
  it('discovers the table through the worksheet rel target with source part and columns', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    const tables = tablesOf(doc)
    expect(tables, 'table metadata discovered via worksheet tableParts').toBeDefined()
    if (!tables) return
    expect(tables.length).toBe(1)
    const t = tables[0] as LocalTable
    expect(t.name).toBe('Table1')
    expect(t.displayName).toBe('Table1')
    expect(t.partPath).toBe('xl/tables/table1.xml')
    expect(t.columns.map((c) => c.name)).toEqual(['ColA', 'ColB'])
    expect(t.columns.map((c) => String(c.id))).toEqual(['1', '2'])
    expect(t.columns.map((c) => c.index)).toEqual([0, 1])
    expect(t.headerRowCount).toBe(1)
    expect(t.totalsRowCount).toBe(0)
    expect(t.extent.firstCol).toBe(0)
    expect(t.extent.firstRow).toBe(0)
    expect(t.extent.cols).toBe(2)
    expect(t.extent.rows).toBe(3)
  })

  it('retains a header0 table distinctly', async () => {
    const doc = await parseFiles(
      tableOnlyFiles({ ref: 'A1:B2', headerRowCount: '0', totalsRowShown: '0' }),
    )
    const tables = tablesOf(doc)
    expect(tables, 'header0 table metadata').toBeDefined()
    if (!tables) return
    expect(tables.length).toBe(1)
    expect(tables[0]?.headerRowCount).toBe(0)
    expect(tables[0]?.totalsRowCount).toBe(0)
    expect(tables[0]?.extent.cols).toBe(2)
    expect(tables[0]?.extent.rows).toBe(2)
  })

  it('retains a totals0 table distinctly', async () => {
    const doc = await parseFiles(
      tableOnlyFiles({ ref: 'A1:B3', headerRowCount: '1', totalsRowShown: '0' }),
    )
    const tables = tablesOf(doc)
    expect(tables, 'totals0 table metadata').toBeDefined()
    if (!tables) return
    expect(tables.length).toBe(1)
    expect(tables[0]?.headerRowCount).toBe(1)
    expect(tables[0]?.totalsRowCount).toBe(0)
  })

  it('retains explicit totalsRowCount independently of totalsRowShown', async () => {
    const doc = await parseFiles(
      tableOnlyFiles({
        ref: 'A1:B4',
        headerRowCount: '1',
        totalsRowShown: '0',
        totalsRowCount: '1',
      }),
    )
    const tables = tablesOf(doc)
    expect(tables, 'explicit totalsRowCount=1 with totalsRowShown=0').toBeDefined()
    if (!tables) return
    expect(tables.length).toBe(1)
    expect(tables[0]?.headerRowCount).toBe(1)
    expect(tables[0]?.totalsRowCount).toBe(1)
    expect(tables[0]?.extent.cols).toBe(2)
    expect(tables[0]?.extent.rows).toBe(4)
  })

  it('retains a data-empty table extent without native formula claims', async () => {
    const doc = await parseFiles(
      tableOnlyFiles({ ref: 'A1:B1', headerRowCount: '1', totalsRowShown: '0' }),
    )
    const tables = tablesOf(doc)
    expect(tables, 'data-empty table metadata').toBeDefined()
    if (!tables) return
    expect(tables.length).toBe(1)
    expect(tables[0]?.headerRowCount).toBe(1)
    expect(tables[0]?.extent.cols).toBe(2)
    expect(tables[0]?.extent.rows).toBe(1)
    expect(tables[0]?.columns.map((c) => c.name)).toEqual(['ColA', 'ColB'])
  })
})

describe('B0 prep: calc/date settings and absent unicode without inference', () => {
  it('reads workbookPr date1904 and calcPr manual/autoNoTable/fullCalcOnLoad/iterate/count/delta', async () => {
    const manual = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    const autoNoTableFiles = namesWorkbookFiles({
      workbookPr: `<workbookPr/>`,
      calcPr: `<calcPr calcMode="autoNoTable" fullCalcOnLoad="0" iterate="0" iterateCount="100" iterateDelta="0.001"/>`,
    })
    const autoNoTable = await parseFiles(autoNoTableFiles)
    const mCalc = calcOf(manual)
    expect(mCalc, 'manual calc metadata').toBeDefined()
    if (!mCalc) return
    expect(mCalc.calcMode).toBe('manual')
    expect(mCalc.fullCalcOnLoad).toBe(true)
    expect(mCalc.iterate).toBe(true)
    expect(mCalc.iterateCount).toBe(50)
    expect(mCalc.iterateDelta).toBeCloseTo(0.002, 6)
    const mSem = semanticsOf(manual)
    expect(mSem?.dateSystem).toBe('1904')
    const aCalc = calcOf(autoNoTable)
    expect(aCalc, 'autoNoTable calc metadata').toBeDefined()
    if (!aCalc) return
    expect(aCalc.calcMode).toBe('autoNoTable')
    expect(aCalc.fullCalcOnLoad).toBe(false)
    expect(aCalc.iterate).toBe(false)
    expect(aCalc.iterateCount).toBe(100)
    expect(aCalc.iterateDelta).toBeCloseTo(0.001, 6)
  })

  it('leaves missing unicode metadata absent/unknown without producer inference', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1"/>`,
        withProducer: true,
      }),
    )
    const sem = semanticsOf(doc)
    expect(sem, 'resolved semantics/unicode metadata present').toBeDefined()
    if (!sem) return
    const u = sem.unicode as
      | { version?: unknown; source?: unknown }
      | undefined
    // Parsed workbook has no verified unicode part; when resolved settings are
    // present they must be version 1 from workbook-default, never
    // standalone-default, and never inferred from the producer name.
    expect(u, 'unicode settings resolved for parsed absent metadata').toBeDefined()
    if (u === undefined) return
    expect(u.version).toBe(1)
    expect(u.source).toBe('workbook-default')
    expect(u.source).not.toBe('standalone-default')
    expect(u.source).not.toBe('explicit')
  })
})

describe('B0 prep: ordinary values preserved while metadata is captured', () => {
  it('keeps existing sheet values and source references', async () => {
    const doc = await parseFiles(
      namesWorkbookFiles({
        workbookPr: `<workbookPr date1904="1"/>`,
        calcPr: `<calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="50" iterateDelta="0.002"/>`,
      }),
    )
    const data = doc.sheets.find((s) => s.name === 'Data')
    expect(data).toBeDefined()
    const byRef = new Map(
      (data?.rows ?? []).flatMap((r) => r.cells.map((c) => [c.ref, c.value] as const)),
    )
    expect(byRef.get('A1')).toBe(42)
    expect(byRef.get('B2')).toBe(7)
  })

  it('keeps scalar API primitive and typed error flags', async () => {
    expect(evaluateFormula('1+1')).toBe(2)
    expect(typeof evaluateFormula('"abc"')).toBe('string')
    const doc = await parseFiles({
      'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1"><v>5</v></c><c r="C3" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData></worksheet>`,
    })
    const cell = doc.sheets[0]?.rows[0]?.cells.find((c) => c.ref === 'C3')
    expect(cell?.value).toBe('#DIV/0!')
    expect(cell?.valueIsError).toBe(true)
    expect(cell?.hasCachedValue).toBe(true)
    expect(doc.sheets[0]?.rows[0]?.cells.find((c) => c.ref === 'A1')?.value).toBe(5)
  })
})

// B0 r2: ported independent extraction + inventory completeness controls
// (refined expectations — synthetic fixtures, never native calculation proof).
// Declared-but-unresolvable table targets leave `tables` undefined
// (unavailable); only known no-tableParts declarations yield complete [].
const REFINED_TABLE = (extra = '', name = 'name="T"') =>
  `<table xmlns="${NS_MAIN}" id="17" ${name} displayName="T" ref="C5:D8" ${extra}><tableColumns count="2"><tableColumn id="29" name="A&apos;[]"/><tableColumn id="3" name=" B "/></tableColumns></table>`

function refinedFiles(opts: {
  extra?: string
  type?: string
  external?: boolean
  missingRels?: boolean
  missingPart?: boolean
  orphan?: boolean
  secondMissing?: boolean
  tableXml?: string
  tablePartDecl?: string
} = {}): Record<string, string> {
  const f: Record<string, string> = {
    'xl/workbook.xml': `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name=" Chart " sheetId="19" r:id="rChart"/><sheet name=" Data " sheetId="44" r:id="rData"/></sheets><definedNames><definedName name="Local" localSheetId="1"> ' Data '!$C$5 </definedName></definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${NS_REL}"><Relationship Id="rChart" Type="${NS_R}/chartsheet" Target="chartsheets/chart.xml"/><Relationship Id="rData" Type="${NS_R}/worksheet" Target="worksheets/data.xml"/><Relationship Id="table1" Type="${NS_R}/table" Target="tables/wrong.xml"/></Relationships>`,
    'xl/chartsheets/chart.xml': `<chartsheet xmlns="${NS_MAIN}"/>`,
    'xl/worksheets/data.xml': `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="5"><c r="C5"><v>11</v></c></row></sheetData><tableParts count="1">${opts.tablePartDecl ?? '<tablePart r:id="table1"/>'}</tableParts></worksheet>`,
    'xl/worksheets/_rels/data.xml.rels': `<Relationships xmlns="${NS_REL}"><Relationship Id="table1" Type="${opts.type ?? `${NS_R}/table`}" Target="../tables/./data.xml"${opts.external ? ' TargetMode="External"' : ''}/></Relationships>`,
    'xl/tables/data.xml': opts.tableXml ?? REFINED_TABLE(opts.extra),
    'xl/tables/wrong.xml': REFINED_TABLE('totalsRowCount="1"'),
  }
  if (opts.missingRels) delete f['xl/worksheets/_rels/data.xml.rels']
  if (opts.missingPart) delete f['xl/tables/data.xml']
  if (opts.orphan) f['xl/tables/orphan.xml'] = REFINED_TABLE('totalsRowCount="1"')
  if (opts.secondMissing) {
    f['xl/worksheets/data.xml'] = f['xl/worksheets/data.xml'].replace(
      'count="1"><tablePart r:id="table1"/>',
      'count="2"><tablePart r:id="table1"/><tablePart r:id="table2"/>',
    )
    f['xl/worksheets/_rels/data.xml.rels'] = f['xl/worksheets/_rels/data.xml.rels'].replace(
      '</Relationships>',
      `<Relationship Id="table2" Type="${NS_R}/table" Target="../tables/missing.xml"/></Relationships>`,
    )
  }
  return f
}

describe('B0 r2: refined extraction controls', () => {
  it('does not turn historical totalsRowShown into actual totals rows', async () => {
    expect((await parseFiles(refinedFiles({ extra: 'totalsRowShown="1"' }))).tables?.[0]?.totalsRowCount).toBe(0)
    expect((await parseFiles(refinedFiles({ extra: 'totalsRowShown="true"' }))).tables?.[0]?.totalsRowCount).toBe(0)
  })

  it('keeps explicit totals counts independent of historical shown', async () => {
    const zero = await parseFiles(refinedFiles({ extra: 'totalsRowShown="true" totalsRowCount="0"' }))
    expect(zero.tables?.[0]?.totalsRowCount).toBe(0)
    const one = await parseFiles(refinedFiles({ extra: 'totalsRowShown="true" totalsRowCount="1"' }))
    expect(one.tables?.[0]).toMatchObject({
      totalsRowCount: 1,
      extent: { sheetId: '44', firstCol: 2, firstRow: 4, cols: 2, rows: 4 },
    })
  })

  it('accepts a valid optional-name table through required displayName', () => {
    expect(parseTableMetadata(parseXml(REFINED_TABLE('', '')), '44', 'xl/tables/data.xml')).toMatchObject({
      name: 'T',
      displayName: 'T',
    })
  })

  it('retains order/identities/names and ignores orphan parts', async () => {
    const d = await parseFiles(refinedFiles({ orphan: true }))
    expect(d.workbookSheets).toEqual([
      { sheetId: '19', workbookIndex: 0, name: ' Chart ', kind: 'other' },
      { sheetId: '44', workbookIndex: 1, name: ' Data ', kind: 'worksheet' },
    ])
    expect(d.definedNames?.[0]).toEqual({
      name: 'Local',
      localSheetIndex: 1,
      source: " ' Data '!$C$5 ",
      baseProvenance: 'unknown',
    })
    expect(d.tables).toHaveLength(1)
    expect(d.tables?.[0]).toMatchObject({
      sheetId: '44',
      partPath: 'xl/tables/data.xml',
      columns: [
        { id: '29', name: "A'[]", index: 0 },
        { id: '3', name: ' B ', index: 1 },
      ],
    })
    expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
  })

  it('keeps declared external target inventory unavailable without publication', async () => {
    expect((await parseFiles(refinedFiles({ external: true }))).tables).toBeUndefined()
  })

  it('keeps wrong-type table ownership unavailable without publication', async () => {
    const d = await parseFiles(refinedFiles({ type: `${NS_R}/image` }))
    expect(d.tables).toBeUndefined()
    expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
  })

  it.each(['missingRels', 'missingPart'] as const)(
    'distinguishes unavailable %s table metadata from known empty',
    async (key) => {
      const d = await parseFiles(refinedFiles({ [key]: true }))
      expect(d.tables).toBeUndefined()
      expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
    },
  )

  it('uses workbook defaults with no Unicode provenance guess', async () => {
    const d = await parseFiles(refinedFiles())
    expect(d.calc).toEqual({
      calcMode: 'auto',
      fullCalcOnLoad: false,
      iterate: false,
      iterateCount: 100,
      iterateDelta: 0.001,
    })
    expect(d.semantics?.dateSystem).toBe('1900')
    expect(d.semantics?.unicode).toEqual({ version: 1, source: 'workbook-default' })
  })
})

describe('B0 r2: refined inventory completeness', () => {
  it('publishes known empty only when worksheet declares no tableParts', async () => {
    const f = refinedFiles()
    f['xl/worksheets/data.xml'] = f['xl/worksheets/data.xml'].replace(/<tableParts[^>]*>.*?<\/tableParts>/, '')
    expect((await parseFiles(f)).tables).toEqual([])
  })

  it('keeps a declared target with wrong XML root unavailable', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE().replace('<table ', '<notTable ').replace('</table>', '</notTable>')
    const d = await parseFiles(f)
    expect(d.tables).toBeUndefined()
    expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
  })

  it('does not publish a partial valid+missing inventory as complete', async () => {
    const d = await parseFiles(refinedFiles({ secondMissing: true }))
    expect(d.tables).toBeUndefined()
    expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
  })
})

// B0 r3: root identity, required metadata, declaration counting, unsigned counts.
// Synthetic extraction controls; official ECMA schema provenance (see REPORT).
describe('B0 r3: table root and required metadata controls', () => {
  it('accepts a namespace-prefixed table root with identical metadata', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE()
      .replace('<table ', `<s:table xmlns:s="${NS_MAIN}" `)
      .replace('</table>', '</s:table>')
    expect((await parseFiles(f)).tables?.[0]).toMatchObject({
      id: '17',
      name: 'T',
      extent: { firstCol: 2, firstRow: 4, cols: 2, rows: 4 },
    })
  })

  it('accepts a leading XML comment before a valid table root', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = '<?xml version="1.0"?><!-- source note -->' + REFINED_TABLE()
    expect((await parseFiles(f)).tables?.[0]?.id).toBe('17')
  })

  it('rejects prefix table with non-table local root name', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE()
      .replace('<table ', `<table:wrong xmlns:table="${NS_MAIN}" `)
      .replace('</table>', '</table:wrong>')
    expect((await parseFiles(f)).tables).toBeUndefined()
  })

  it.each(['id', 'displayName', 'ref'])(
    'leaves declared table with required %s missing unavailable',
    async (attr) => {
      const f = refinedFiles()
      f['xl/tables/data.xml'] = REFINED_TABLE().replace(new RegExp(` ${attr}="[^"]*"`), '')
      const d = await parseFiles(f)
      expect(d.tables).toBeUndefined()
      expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
    },
  )

  it('does not invent A1 geometry for invalid source ref', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE().replace('ref="C5:D8"', 'ref="not-a-range"')
    expect((await parseFiles(f)).tables).toBeUndefined()
  })

  it('does not call a declared tablePart without required relationship id known empty', async () => {
    const f = refinedFiles({ tablePartDecl: '<tablePart/>' })
    expect((await parseFiles(f)).tables).toBeUndefined()
  })

  it('retains valid unsigned count lexical values instead of equality guessing', async () => {
    const d = await parseFiles(refinedFiles({ extra: 'headerRowCount="00" totalsRowCount="01"' }))
    expect(d.tables?.[0]).toMatchObject({ headerRowCount: 0, totalsRowCount: 1 })
  })

  it('does not accept boolean true as an unsigned totals count', async () => {
    expect((await parseFiles(refinedFiles({ extra: 'totalsRowCount="true"' }))).tables).toBeUndefined()
  })

  it('does not fabricate a missing tableColumn id from its order', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE().replace('tableColumn id="29"', 'tableColumn')
    expect((await parseFiles(f)).tables).toBeUndefined()
  })
})

// B0 r6: column caption and worksheet discovery completeness (QUALITY).
// Missing essential column metadata keeps the inventory unavailable — never
// a shortened, reindexed list. An uninspectable referenced worksheet keeps
// the inventory unavailable even with zero observed declarations; a readable
// worksheet with no tableParts is complete [].
function completenessFiles(
  columns = '<tableColumn id="1" name="Left"/><tableColumn id="2" name="Right"/>',
): Record<string, string> {
  return {
    'xl/workbook.xml': `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="Data" sheetId="17" r:id="ws"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${NS_REL}"><Relationship Id="ws" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData><row r="5"><c r="C5"><v>11</v></c></row></sheetData><tableParts count="1"><tablePart r:id="t"/></tableParts></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${NS_REL}"><Relationship Id="t" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
    'xl/tables/table1.xml': `<table xmlns="${NS_MAIN}" id="1" displayName="T" ref="A1:B3"><tableColumns count="2">${columns}</tableColumns></table>`,
  }
}

async function parseEntries(entries: Record<string, string>): Promise<XlsxDocument> {
  const zip = new JSZip()
  for (const [p, v] of Object.entries(entries)) zip.file(p, v)
  return parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}

describe('B0 r6: column and worksheet discovery completeness', () => {
  it('preserves valid column caption IDs and positions', async () => {
    const doc = await parseEntries(completenessFiles())
    expect(doc.sheets[0].rows.flatMap((r) => r.cells).find((c) => c.ref === 'C5')?.value).toBe(11)
    expect(doc.tables?.[0]?.columns).toEqual([
      { id: '1', name: 'Left', index: 0 },
      { id: '2', name: 'Right', index: 1 },
    ])
  })

  it('missing required first column name keeps inventory unavailable instead of shifting', async () => {
    const doc = await parseEntries(completenessFiles('<tableColumn id="1"/><tableColumn id="2" name="Right"/>'))
    expect(doc.sheets[0].rows.flatMap((r) => r.cells).find((c) => c.ref === 'C5')?.value).toBe(11)
    expect(doc.tables).toBeUndefined()
  })

  it('worksheet discovery read failure cannot become known-empty inventory', async () => {
    // Error-path control only: injected second-read failure proves the
    // catch/completeness path; it does not claim cached reads normally throw.
    const zip = new JSZip()
    for (const [p, v] of Object.entries(completenessFiles())) zip.file(p, v)
    const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
    const original = pkg.xml.bind(pkg)
    let calls = 0
    pkg.xml = async (path: string) => {
      if (path === 'xl/worksheets/sheet1.xml' && ++calls === 2) throw new Error('injected discovery read failure')
      return original(path)
    }
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].rows.flatMap((r) => r.cells).find((c) => c.ref === 'C5')?.value).toBe(11)
    expect(doc.tables).toBeUndefined()
  })

  it('unreadable/missing referenced worksheet cannot establish known-empty inventory', async () => {
    const entries = completenessFiles()
    delete entries['xl/worksheets/sheet1.xml']
    const doc = await parseEntries(entries)
    expect(doc.workbookSheets).toEqual([{ sheetId: '17', workbookIndex: 0, name: 'Data', kind: 'worksheet' }])
    expect(doc.tables).toBeUndefined()
  })

  it('readable worksheet with no table declarations is known empty', async () => {
    const entries = completenessFiles()
    entries['xl/worksheets/sheet1.xml'] =
      `<worksheet xmlns="${NS_MAIN}"><sheetData><row r="5"><c r="C5"><v>11</v></c></row></sheetData></worksheet>`
    const doc = await parseEntries(entries)
    expect(doc.sheets[0].rows.flatMap((r) => r.cells).find((c) => c.ref === 'C5')?.value).toBe(11)
    expect(doc.tables).toEqual([])
  })
})

// B0 r4: parsed document-element root and required unsigned identity.
// Synthetic extraction controls; official ECMA schema provenance (see REPORT).
describe('B0 r4: parsed root and unsigned identity boundaries', () => {
  it('accepts whitespace before the XML document element without prolog', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = ' \n\t' + REFINED_TABLE()
    expect((await parseFiles(f)).tables?.[0]?.id).toBe('17')
  })

  it('rejects a non-table Unicode document element without prefix truncation', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE().replace('<table ', '<tableé ').replace('</table>', '</tableé>')
    expect((await parseFiles(f)).tables).toBeUndefined()
  })

  it('accepts a legal Unicode namespace prefix on table', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE()
      .replace('<table ', `<é:table xmlns:é="${NS_MAIN}" `)
      .replace('</table>', '</é:table>')
    expect((await parseFiles(f)).tables?.[0]?.id).toBe('17')
  })

  it.each(['table', 'column'])(
    'does not publish invalid required unsigned %s ID as complete',
    async (which) => {
      const f = refinedFiles()
      f['xl/tables/data.xml'] =
        which === 'table'
          ? REFINED_TABLE().replace('id="17"', 'id="bad"')
          : REFINED_TABLE().replace('id="29"', 'id="bad"')
      const d = await parseFiles(f)
      expect(d.tables).toBeUndefined()
      expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
    },
  )
})

// B0 r5: required unsignedInt identity domain (W3C xmlschema-2#unsignedInt,
// max 4294967295). Synthetic extraction controls; no native claim.
describe('B0 r5: required unsignedInt identity domain', () => {
  it.each(['table', 'column'])('rejects unsignedInt overflow in required %s ID', async (which) => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] =
      which === 'table'
        ? REFINED_TABLE().replace('id="17"', 'id="4294967296"')
        : REFINED_TABLE().replace('id="29"', 'id="4294967296"')
    const d = await parseFiles(f)
    expect(d.tables).toBeUndefined()
    expect(d.sheets[1].rows[0].cells[0].value).toBe(11)
  })

  it('retains valid source ID spelling', async () => {
    const f = refinedFiles()
    f['xl/tables/data.xml'] = REFINED_TABLE().replace('id="17"', 'id="00017"').replace('id="29"', 'id="00029"')
    const d = await parseFiles(f)
    expect(d.tables?.[0]?.id).toBe('00017')
    expect(d.tables?.[0]?.columns[0]?.id).toBe('00029')
  })
})
