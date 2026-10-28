/** Deterministic raw SpreadsheetML cases for pivot/table/control reference captures. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { REF_CASES, type RefCase, type SlicerCase, type TableCase, type TimelineCase } from './cases'

const NS = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pkg: 'http://schemas.openxmlformats.org/package/2006/relationships',
  content: 'http://schemas.openxmlformats.org/package/2006/content-types',
  x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
  x15: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
  xm: 'http://schemas.microsoft.com/office/excel/2006/main',
  sle: 'http://schemas.microsoft.com/office/drawing/2010/slicer',
  tsle: 'http://schemas.microsoft.com/office/drawing/2010/timeslicer',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pivot: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
  cache: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
  pivotTableRel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
  pivotCacheRel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
  cacheRecordsRel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
  slicerRel: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
  slicerCacheRel: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
  timelineRel: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
  timelineCacheRel: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
}
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const ZIP_DATE = new Date(Date.UTC(2026, 0, 1))
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export interface SourceRow { region: string; product: string; year: number; date: number; sales: number; units: number }
/** A seeded source table with stable categories and non-tied numeric values. */
export function pivotSourceRows(seed: number, n: number): SourceRow[] {
  if (!Number.isSafeInteger(seed) || !Number.isSafeInteger(n) || n < 0 || n > 1_000_000) throw new RangeError('seed and row count must be safe integers; n must be 0..1000000')
  let state = seed >>> 0
  const next = () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state }
  const regions = ['East', 'West', 'North', 'South']
  const products = ['Apples', 'Bananas', 'Cherries', 'Dates', 'Elderberries', 'Figs', 'Grapes', 'Honeydew']
  return Array.from({ length: n }, (_, i) => ({
    region: regions[(i + (seed >>> 0)) % regions.length], product: products[(i + (seed >>> 0)) % products.length],
    year: 2022 + (next() % 3), date: 44_927 + (next() % 1_095), sales: 100 + (next() % 9000) + i / 100, units: 1 + (next() % 90),
  }))
}

const colName = (index: number): string => { let n = index + 1, out = ''; while (n) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26) } return out }
const cell = (row: number, col: number) => `${colName(col)}${row}`
const headers = ['Region', 'Product', 'Year', 'Sales', 'Units']
const rows = pivotSourceRows(20261029, 5)
const data = [headers, ...rows.map(r => [r.region, r.product, r.year, r.sales, r.units] as const)]
const controlHeaders = ['Region', 'Product', 'Date', 'Sales', 'Units']
const controlRows = pivotSourceRows(20261029, 16)
const controlData = [controlHeaders, ...controlRows.map(r => [r.region, r.product, r.date, r.sales, r.units] as const)]

function worksheetXml(tableCase?: TableCase, controlKind?: 'slicer' | 'timeline', controlSource = false): string {
  const control = controlKind !== undefined || controlSource
  const baseRows = controlSource ? controlData : data
  const tableRows = tableCase?.totalsRowCount ? [...baseRows, ['Total', '', '', rows.reduce((n, r) => n + r.sales, 0), rows.reduce((n, r) => n + r.units, 0)] as const] : baseRows
  const startCol = controlKind ? 6 : 0
  const rowXml = tableRows.map((values, ri) => `<row r="${ri + 1}">${values.map((v, ci) => {
    const s = ''
    if (v === '') return `<c r="${cell(ri + 1, ci + startCol)}"${s}/>`
    return typeof v === 'number' ? `<c r="${cell(ri + 1, ci + startCol)}"${s}><v>${v}</v></c>` : `<c r="${cell(ri + 1, ci + startCol)}"${s} t="inlineStr"><is><t>${esc(String(v))}</t></is></c>`
  }).join('')}</row>`).join('')
  const tablePart = tableCase ? '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' : ''
  const drawing = control ? '<drawing r:id="rId2"/>' : ''
  const pivotPart = control ? '<pivotTableParts count="1"><pivotTablePart r:id="rId4"/></pivotTableParts>' : ''
  const ext = controlKind === 'slicer'
    ? `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${NS.x14}"><x14:slicerList><x14:slicer r:id="rId3"/></x14:slicerList></ext></extLst>`
    : controlKind === 'timeline' ? `<extLst><ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" xmlns:x15="${NS.x15}"><x15:timelineRefs><x15:timelineRef r:id="rId3"/></x15:timelineRefs></ext></extLst>` : ''
  const dimension = `${cell(1, startCol)}:${cell(tableRows.length, startCol + 4)}`
  return `${XML}<worksheet xmlns="${NS.main}" xmlns:r="${NS.rel}" xmlns:x14="${NS.x14}" xmlns:x15="${NS.x15}"><dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rowXml}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${tablePart}${drawing}${pivotPart}${ext}</worksheet>`
}

function tableXml(spec: TableCase): string {
  const columns = headers.map((name, i) => `<tableColumn id="${i + 1}" name="${name}"/>`).join('')
  const lastRow = spec.totalsRowCount ? 7 : 6
  return `${XML}<table xmlns="${NS.main}" id="1" name="Table1" displayName="Table1" ref="A1:E${lastRow}" headerRowCount="${spec.headerRowCount}" totalsRowCount="${spec.totalsRowCount}"><autoFilter ref="A1:E${lastRow - spec.totalsRowCount}"/><tableColumns count="5">${columns}</tableColumns><tableStyleInfo name="${esc(spec.styleName)}" showFirstColumn="${+spec.flags.showFirstColumn}" showLastColumn="${+spec.flags.showLastColumn}" showRowStripes="${+spec.flags.showRowStripes}" showColumnStripes="${+spec.flags.showColumnStripes}"/></table>`
}

function drawingXml(c: RefCase): string {
  const isTimeline = c.kind === 'timeline'
  const graphic = isTimeline
    ? `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Timeline 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${NS.tsle}"><tsle:timeslicer name="Timeline1"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    : `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Slicer 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${NS.sle}"><sle:slicer name="Slicer1"/></a:graphicData></a:graphic></xdr:graphicFrame>`
  const fallback = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Control fallback"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Control</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
  return `${XML}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:mc="${NS.mc}" xmlns:sle="${NS.sle}" xmlns:tsle="${NS.tsle}"><xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent><mc:Choice Requires="${isTimeline ? 'tsle' : 'sle'}">${graphic}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
}

function controlParts(c: RefCase): Record<string, string> {
  const slicer = c.kind === 'slicer'
  const spec = c.spec as SlicerCase | TimelineCase
  if (slicer) {
    const s = spec as SlicerCase
    const items = Array.from({ length: s.itemCount }, (_, i) => `<i x="${i}" s="${i >= s.startItem && i < s.startItem + 2 ? 1 : 0}"${s.noData && i === s.itemCount - 1 ? ' nd="1"' : ''}/>`).join('')
    return {
      'xl/slicers/slicer1.xml': `${XML}<slicers xmlns="${NS.x14}" xmlns:r="${NS.rel}"><slicer name="Slicer1" cache="Slicer_Cache1" caption="Product" rowHeight="285750" style="${s.styleName}" showCaption="1" level="0" columnCount="${s.columnCount}" startItem="${s.startItem}"/></slicers>`,
      'xl/slicerCaches/slicerCache1.xml': `${XML}<slicerCacheDefinition xmlns="${NS.x14}" xmlns:r="${NS.rel}" name="Slicer_Cache1" sourceName="Product"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="1"><items count="${s.itemCount}">${items}</items></tabular></data></slicerCacheDefinition>`,
    }
  }
  const t = spec as TimelineCase
  const level = ({ years: 0, quarters: 1, months: 2, days: 3 } as const)[t.level]
  return {
    'xl/timelines/timeline1.xml': `${XML}<timelines xmlns="${NS.x15}" xmlns:r="${NS.rel}"><timeline name="Timeline1" cache="Timeline_Cache1" caption="Date" style="${t.styleName}" level="${level}"/></timelines>`,
    'xl/timelineCaches/timelineCache1.xml': `${XML}<timelineCacheDefinition xmlns="${NS.x15}" xmlns:r="${NS.rel}" name="Timeline_Cache1"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><timelineState field="Date" start="2022-01-01T00:00:00" end="2024-12-31T00:00:00"/></timelineCacheDefinition>`,
  }
}

function pivotParts(): Record<string, string> {
  const values = controlRows.map(row => [row.region, row.product, row.date, row.sales, row.units] as const)
  const cacheFields = controlHeaders.map((name, field) => {
    const unique = [...new Set(values.map(row => row[field]))]
    const items = unique.map(value => typeof value === 'number' ? `<n v="${value}"/>` : `<s v="${esc(String(value))}"/>`).join('')
    return `<cacheField name="${name}"><sharedItems count="${unique.length}">${items}</sharedItems></cacheField>`
  }).join('')
  const records = values.map(row => `<r>${row.map((value, i) => `<x v="${[...new Set(values.map(r => r[i]))].indexOf(value)}"/>`).join('')}</r>`).join('')
  const fields = controlHeaders.map((_, i) => `<pivotField${i === 0 ? ' axis="axisRow"' : ''} showAll="0"><items count="${i === 0 ? 4 : 1}">${i === 0 ? Array.from({ length: 4 }, (_, j) => `<item x="${j}"/>`).join('') : '<item t="default"/>'}</items></pivotField>`).join('')
  return {
    'xl/pivotCache/pivotCacheDefinition1.xml': `${XML}<pivotCacheDefinition xmlns="${NS.main}" xmlns:r="${NS.rel}" r:id="rId1" recordCount="${values.length}" refreshOnLoad="1" saveData="1"><cacheSource type="worksheet"><worksheetSource ref="A1:E${values.length + 1}" sheet="Data"/></cacheSource><cacheFields count="5">${cacheFields}</cacheFields></pivotCacheDefinition>`,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.cacheRecordsRel}" Target="pivotCacheRecords1.xml"/></Relationships>`,
    'xl/pivotCache/pivotCacheRecords1.xml': `${XML}<pivotCacheRecords xmlns="${NS.main}" count="${values.length}">${records}</pivotCacheRecords>`,
    'xl/pivotTables/pivotTable1.xml': `${XML}<pivotTableDefinition xmlns="${NS.main}" xmlns:r="${NS.rel}" name="PivotTable1" cacheId="1" dataCaption="Values" compact="1" outline="1" subtotalTop="1"><location ref="G1:K6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="5">${fields}</pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i><x v="0"/></i></rowItems><colFields count="0"/><colItems count="0"/><dataFields count="1"><dataField name="Sum of Sales" fld="3" subtotal="sum"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0"/></pivotTableDefinition>`,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.pivotCacheRel}" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
  }
}

function workbookFiles(c: RefCase): Record<string, string> {
  const tableCase = c.kind === 'table' ? c.spec as TableCase : undefined
  const control = c.kind === 'slicer' || c.kind === 'timeline'
  const controlXml = control ? controlParts(c) : {}
  const contentOverrides = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ...Array.from({ length: control ? 2 : 1 }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`),
    ...(tableCase ? ['<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'] : []),
    ...(control ? ['<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'] : []),
    ...(control ? ['<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'] : []),
    ...(control && c.kind === 'slicer' ? ['<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>'] : []),
    ...(control && c.kind === 'timeline' ? ['<Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>', '<Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>'] : []),
  ].join('')
  const mainSheet = c.kind === 'table' ? 'Data' : 'Report'
  const sourceSheets = control ? ['Data', 'Report'] : [mainSheet]
  const sheetFiles: Record<string, string> = {}
  sourceSheets.forEach((name, i) => {
    const isReport = name === 'Report'
    sheetFiles[`xl/worksheets/sheet${i + 1}.xml`] = worksheetXml(!control && i === 0 ? tableCase : undefined, control && isReport ? c.kind as 'slicer' | 'timeline' : undefined, control && name === 'Data')
    const rels = [
      ...(tableCase && i === 0 ? ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'] : []),
      ...(control && isReport ? [
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>',
        `<Relationship Id="rId3" Type="${c.kind === 'slicer' ? NS.slicerRel : NS.timelineRel}" Target="../${c.kind === 'slicer' ? 'slicers/slicer1.xml' : 'timelines/timeline1.xml'}"/>`,
        `<Relationship Id="rId4" Type="${NS.pivotTableRel}" Target="../pivotTables/pivotTable1.xml"/>`,
      ] : []),
    ].join('')
    if (rels) sheetFiles[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = `${XML}<Relationships xmlns="${NS.pkg}">${rels}</Relationships>`
  })
  const workbookRels = sourceSheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${NS.rel}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
  const workbookExt = control ? c.kind === 'slicer'
    ? `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${NS.x14}"><x14:slicerCaches><x14:slicerCache r:id="rId${sourceSheets.length + 2}" xmlns:r="${NS.rel}"/></x14:slicerCaches></ext></extLst>`
    : `<extLst><ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" xmlns:x15="${NS.x15}"><x15:timelineCacheRefs><x15:timelineCacheRef r:id="rId${sourceSheets.length + 2}" xmlns:r="${NS.rel}"/></x15:timelineCacheRefs></ext></extLst>` : ''
  const wbControlRel = control ? `<Relationship Id="rId${sourceSheets.length + 1}" Type="${NS.pivotCacheRel}" Target="pivotCache/pivotCacheDefinition1.xml"/><Relationship Id="rId${sourceSheets.length + 2}" Type="${c.kind === 'slicer' ? NS.slicerCacheRel : NS.timelineCacheRel}" Target="${c.kind === 'slicer' ? 'slicerCaches/slicerCache1.xml' : 'timelineCaches/timelineCache1.xml'}"/>` : ''
  return {
    '[Content_Types].xml': `${XML}<Types xmlns="${NS.content}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${contentOverrides}</Types>`,
    '_rels/.rels': `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${XML}<workbook xmlns="${NS.main}" xmlns:r="${NS.rel}"><workbookPr/><sheets>${sourceSheets.map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${control ? `<pivotCaches><pivotCache cacheId="1" r:id="rId${sourceSheets.length + 1}"/></pivotCaches>` : ''}${workbookExt}</workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}">${workbookRels}${wbControlRel}</Relationships>`,
    'xl/styles.xml': `${XML}<styleSheet xmlns="${NS.main}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${tableCase?.custom ? '<dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="OfficeViewCustom" pivot="0" table="1" count="2"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="0"/></tableStyle></tableStyles>' : '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'}</styleSheet>`,
    ...sheetFiles,
    ...(tableCase ? { 'xl/tables/table1.xml': tableXml(tableCase) } : {}),
    ...(control ? { 'xl/drawings/drawing1.xml': drawingXml(c), ...controlXml, ...pivotParts() } : {}),
  }
}

/** Makes a package for an XML case. Excel-driven cases are created by Task 3. */
export async function buildCaseWorkbook(c: RefCase): Promise<Uint8Array> {
  if (c.build !== 'xml') throw new Error(`case ${c.id} is Excel-built and must be generated by the Excel reference runner`)
  const zip = new JSZip()
  for (const [path, content] of Object.entries(workbookFiles(c))) zip.file(path, content, { date: ZIP_DATE, createFolders: false })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Files written by this tool when invoked directly; output is deterministic. */
export const FIXTURE_DIR = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots')
export async function writeFixtures(): Promise<number> {
  mkdirSync(FIXTURE_DIR, { recursive: true })
  let count = 0
  for (const c of REF_CASES) if (c.build === 'xml') {
    writeFileSync(resolve(FIXTURE_DIR, `${c.id}.xlsx`), await buildCaseWorkbook(c))
    count++
  }
  return count
}

if (process.argv[1] && resolve(process.argv[1]) === resolve(process.cwd(), 'scripts/xlsx-refs/generate.ts')) {
  console.log(`wrote ${await writeFixtures()} pivot reference workbooks to ${FIXTURE_DIR}`)
}
