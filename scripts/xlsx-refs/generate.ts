/** Deterministic raw SpreadsheetML cases for pivot/table/control reference captures. */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import JSZip from 'jszip'
import { REF_CASES, type RefCase, type SlicerCase, type TableCase, type TimelineCase } from './cases'

const NS = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pkg: 'http://schemas.openxmlformats.org/package/2006/relationships',
  content: 'http://schemas.openxmlformats.org/package/2006/content-types',
  x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
  x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
  x15: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
  xr: 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  xr2: 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
  xr6: 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
  xr10: 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
  x15ac: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac',
  xm: 'http://schemas.microsoft.com/office/excel/2006/main',
  sle: 'http://schemas.microsoft.com/office/drawing/2010/slicer',
  tsle: 'http://schemas.microsoft.com/office/drawing/2012/timeslicer',
  sle15: 'http://schemas.microsoft.com/office/drawing/2012/slicer',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
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
  const makeRandom = (salt: number) => {
    let state = (seed ^ salt) >>> 0
    return () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state }
  }
  const regions = ['East', 'West', 'North', 'South']
  const products = ['Apples', 'Bananas', 'Cherries', 'Dates', 'Elderberries', 'Figs', 'Grapes', 'Honeydew']
  const nextRegion = makeRandom(0x9e3779b9), nextProduct = makeRandom(0x85ebca6b), nextYear = makeRandom(0xc2b2ae35), nextDate = makeRandom(0x27d4eb2f), nextSales = makeRandom(0x165667b1), nextUnits = makeRandom(0xd3a2646c)
  return Array.from({ length: n }, (_, i) => ({
    region: regions[nextRegion() % regions.length], product: products[i < products.length ? (i + (seed >>> 0)) % products.length : nextProduct() % products.length],
    year: 2022 + (nextYear() % 3), date: 44_562 + (nextDate() % 1_096), sales: 100 + (nextSales() % 9000) + i / 100, units: 1 + (nextUnits() % 90),
  }))
}

const colName = (index: number): string => { let n = index + 1, out = ''; while (n) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26) } return out }
const cell = (row: number, col: number) => `${colName(col)}${row}`
const headers = ['Region', 'Product', 'Year', 'Sales', 'Units']
const rows = pivotSourceRows(20261029, 5)
const data = [headers, ...rows.map(r => [r.region, r.product, r.year, r.sales, r.units] as const)]
const controlHeaders = ['Region', 'Product', 'Date', 'Sales', 'Units']
const controlRows = pivotSourceRows(20261029, 40).map((row, i) => ({
  ...row,
  product: ['Apples', 'Bananas', 'Cherries', 'Dates'][[1, 3, 0, 2][i % 4]],
  // Match the reference's 39 distinct Date members across 40 pivot-cache records.
  date: 45_303 + (i % 39),
}))
const controlData = [controlHeaders, ...controlRows.map(r => [r.region, r.product, r.date, r.sales, r.units] as const)]
const reportRegions = ['East', 'West', 'North', 'South']
const reportRows = [['Region', 'Sum of Sales'], ...reportRegions.map(region => [region, controlRows.filter(row => row.region === region).reduce((total, row) => total + row.sales, 0)] as const), ['Grand Total', controlRows.reduce((total, row) => total + row.sales, 0)] as const]
const toDateTime = (serial: number) => new Date(Date.UTC(1899, 11, 30 + serial)).toISOString().replace('.000Z', 'Z')

function worksheetXml(tableCase?: TableCase, controlKind?: 'slicer' | 'timeline', controlSource = false): string {
  const control = controlKind !== undefined || controlSource
  const baseRows = controlSource ? controlData : controlKind ? reportRows : data
  const tableRows = tableCase?.totalsRowCount ? [...baseRows, ['Total', '', '', rows.reduce((n, r) => n + r.sales, 0), rows.reduce((n, r) => n + r.units, 0)] as const] : baseRows
  // Keep the Report's saved values to the left of the pivot output (G:K).
  const startCol = controlKind ? 6 : 0
  const rowXml = tableRows.map((values, ri) => `<row r="${ri + 1}">${values.map((v, ci) => {
    const s = (controlSource || controlKind) && ci === 2 && ri > 0 ? ' s="2"' : ''
    if (v === '') return `<c r="${cell(ri + 1, ci + startCol)}"${s}/>`
    return typeof v === 'number' ? `<c r="${cell(ri + 1, ci + startCol)}"${s}><v>${v}</v></c>` : `<c r="${cell(ri + 1, ci + startCol)}"${s} t="inlineStr"><is><t>${esc(String(v))}</t></is></c>`
  }).join('')}</row>`).join('')
  const tablePart = tableCase ? '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' : ''
  const drawing = controlKind ? '<drawing r:id="rId2"/>' : ''
  const ext = controlKind
    ? `<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${NS.x14}"><x14:slicerList><x14:slicer r:id="rId3"/></x14:slicerList></ext><ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${NS.x15}"><x15:timelineRefs><x15:timelineRef r:id="rId4"/></x15:timelineRefs></ext></extLst>` : ''
  const dimension = `${cell(1, startCol)}:${cell(tableRows.length, startCol + 4)}`
  const controlNamespaces = controlKind ? ` xmlns:mc="${NS.mc}" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="${NS.x14ac}" xmlns:xr="${NS.xr}" xmlns:xr2="${NS.xr2}" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"` : ''
  const sheetViews = controlKind ? '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>' : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
  const cols = controlKind ? '<cols><col min="3" max="3" width="12.796875" customWidth="1"/><col min="1" max="2" width="9" customWidth="1"/></cols>' : ''
  const phonetic = controlKind ? '<phoneticPr fontId="1" type="noConversion"/>' : ''
  const otherNamespaces = controlKind ? '' : ` xmlns:x14="${NS.x14}" xmlns:x15="${NS.x15}"`
  return `${XML}<worksheet xmlns="${NS.main}" xmlns:r="${NS.rel}"${controlNamespaces}${otherNamespaces}><dimension ref="${dimension}"/>${sheetViews}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowXml}</sheetData>${phonetic}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${tablePart}${drawing}${ext}</worksheet>`
}

function tableXml(spec: TableCase): string {
  const columns = headers.map((name, i) => `<tableColumn id="${i + 1}" name="${name}"/>`).join('')
  const lastRow = spec.totalsRowCount ? 7 : 6
  return `${XML}<table xmlns="${NS.main}" id="1" name="Table1" displayName="Table1" ref="A1:E${lastRow}" headerRowCount="${spec.headerRowCount}" totalsRowCount="${spec.totalsRowCount}"><autoFilter ref="A1:E${lastRow - spec.totalsRowCount}"/><tableColumns count="5">${columns}</tableColumns><tableStyleInfo name="${esc(spec.styleName)}" showFirstColumn="${+spec.flags.showFirstColumn}" showLastColumn="${+spec.flags.showLastColumn}" showRowStripes="${+spec.flags.showRowStripes}" showColumnStripes="${+spec.flags.showColumnStripes}"/></table>`
}

function drawingXml(_c: RefCase): string {
  // These are Excel's saved anchor positions; the timeline ends at column G's left edge,
  // where the saved pivot output begins, so the controls do not cover pivot cells.
  const fallback = (text: string, x: string, y: string, cx: string, cy: string) => `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" altLang="en-US" sz="1100"/><a:t>${esc(text)}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`
  const anchor = (kind: 'slicer' | 'timeline') => {
    const timeline = kind === 'timeline'
    const requires = timeline ? 'tsle' : 'a14'
    const choiceNamespace = timeline ? `xmlns:tsle="${NS.tsle}"` : `xmlns:a14="${NS.a14}"`
    const name = timeline ? 'Date' : 'Product'
    const uri = timeline ? NS.tsle : NS.sle
    const control = timeline ? `<tsle:timeslicer xmlns:tsle="${NS.tsle}" name="Date"/>` : `<sle:slicer xmlns:sle="${NS.sle}" name="Product"/>`
    const id = timeline ? 3 : 2
    const creationId = timeline ? '{BFDAAD64-EAE5-0E16-B7BA-B58DA87D5805}' : '{2D6B128A-1A5B-FE90-F130-93BB57E1F64A}'
    const choice = `<mc:Choice ${choiceNamespace} Requires="${requires}"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${creationId}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${control}</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>`
    const text = timeline ? 'Timeline supported in Excel 2013 or later.' : 'This shape represents a slicer. Excel 2010 or later supports slicers.'
    const fb = timeline ? fallback(text, '1536700', '749300', '3340100', '1346200') : fallback(text, '2286000', '101600', '1828800', '2238375')
    const colFrom = timeline ? 1 : 3, colTo = timeline ? 6 : 5
    const rowFrom = timeline ? 4 : 0, rowTo = timeline ? 11 : 13
    const colOffFrom = timeline ? 787400 : 12700, rowOffFrom = timeline ? 38100 : 101600
    const colOffTo = timeline ? 508000 : 444500, rowOffTo = timeline ? 139700 : 28575
    return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${colFrom}</xdr:col><xdr:colOff>${colOffFrom}</xdr:colOff><xdr:row>${rowFrom}</xdr:row><xdr:rowOff>${rowOffFrom}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${colTo}</xdr:col><xdr:colOff>${colOffTo}</xdr:colOff><xdr:row>${rowTo}</xdr:row><xdr:rowOff>${rowOffTo}</xdr:rowOff></xdr:to><mc:AlternateContent xmlns:mc="${NS.mc}">${choice}<mc:Fallback>${fb}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
  }
  return `${XML}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${anchor('slicer')}${anchor('timeline')}</xdr:wsDr>`
}
function controlParts(c: RefCase): Record<string, string> {
  const slicer = c.kind === 'slicer' ? c.spec as SlicerCase : { styleName: 'SlicerStyleLight1', columnCount: 1, noData: false, startItem: 0, itemCount: 4, scroll: false }
  const timeline = c.kind === 'timeline' ? c.spec as TimelineCase : { styleName: 'TimeSlicerStyleLight1', level: 'months' as const }
  const slicerItems = [2, 1, 3, 0].slice(0, slicer.itemCount).map((index, order) => `<i x="${index}"${index === 2 || index === 3 ? ' s="1"' : ''}${slicer.noData && order === slicer.itemCount - 1 ? ' nd="1"' : ''}/>`).join('')
  const level = ({ years: 0, quarters: 1, months: 2, days: 3 } as const)[timeline.level]
  const nsCache = `xmlns:mc="${NS.mc}" mc:Ignorable="x xr10" xmlns:x="${NS.main}" xmlns:xr10="${NS.xr10}"`
  const nsTimeline = `xmlns:mc="${NS.mc}" mc:Ignorable="x xr10" xmlns:x="${NS.main}" xmlns:xr10="${NS.xr10}"`
  const nsTimelineCache = `xmlns:mc="${NS.mc}" mc:Ignorable="xr10" xmlns:xr10="${NS.xr10}"`
  // Excel omits style for the default TimeSlicerStyleLight1; every other style is written explicitly.
  const timelineStyle = timeline.styleName === 'TimeSlicerStyleLight1' ? '' : ` style="${esc(timeline.styleName)}"`
  // Derived from the Data sheet dates: scroll to the first of the min date's month (months/days) or Jan 1 of
  // its year (years/quarters); bounds run from Jan 1 of the min year to Jan 1 of the year after the max year.
  const serials = controlRows.map(row => row.date)
  const minDate = new Date(toDateTime(Math.min(...serials))), maxDate = new Date(toDateTime(Math.max(...serials)))
  const firstOf = (year: number, month: number) => `${year}-${String(month).padStart(2, '0')}-01T00:00:00`
  const scrollPosition = timeline.level === 'months' || timeline.level === 'days'
    ? firstOf(minDate.getUTCFullYear(), minDate.getUTCMonth() + 1)
    : firstOf(minDate.getUTCFullYear(), 1)
  const boundsStart = firstOf(minDate.getUTCFullYear(), 1), boundsEnd = firstOf(maxDate.getUTCFullYear() + 1, 1)
  return {
    'xl/slicers/slicer1.xml': `${XML}<slicers xmlns="${NS.x14}" ${nsCache}><slicer name="Product" xr10:uid="{9FD6D7D5-54C3-D440-9F1C-DE221ED3D12A}" cache="Slicer_Product" caption="Product" rowHeight="209550" style="${slicer.styleName}" columnCount="${slicer.columnCount}" startItem="${slicer.startItem}" showNoDataItems="${+slicer.noData}"/></slicers>`,
    'xl/slicerCaches/slicerCache1.xml': `${XML}<slicerCacheDefinition xmlns="${NS.x14}" ${nsCache} name="Slicer_Product" xr10:uid="{51DE0EAB-D8C0-484E-8A22-4B5730F7CD76}" sourceName="Product"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><data><tabular pivotCacheId="1"><items count="${slicer.itemCount}">${slicerItems}</items></tabular></data></slicerCacheDefinition>`,
    'xl/timelines/timeline1.xml': `${XML}<timelines xmlns="${NS.x15}" ${nsTimeline}><timeline name="Date" xr10:uid="{BCDD19F5-7DFF-6745-8336-A62734B2B7F9}" cache="NativeTimeline_Date" caption="Date"${timelineStyle} level="${level}" selectionLevel="${level}" scrollPosition="${scrollPosition}"/></timelines>`,
    'xl/timelineCaches/timelineCache1.xml': `${XML}<timelineCacheDefinition xmlns="${NS.x15}" xmlns:x15="${NS.x15}" ${nsTimelineCache} name="NativeTimeline_Date" xr10:uid="{FE35D545-E153-F14A-8900-EDB8C368CE04}" sourceName="Date"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables><state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="1" filterType="unknown"><bounds startDate="${boundsStart}" endDate="${boundsEnd}"/></state></timelineCacheDefinition>`,
  }
}
function pivotParts(): Record<string, string> {
  const values = controlRows.map(row => [row.region, row.product, row.date, row.sales, row.units] as const)
  const cacheFields = controlHeaders.map((name, field) => {
    const unique = [...new Set(values.map(row => row[field]))]
    const dateField = name === 'Date'
    const numericField = !dateField && typeof unique[0] === 'number'
    if (numericField) {
      const numericValues = unique as number[]
      const integerFlag = numericValues.every(Number.isInteger) ? ' containsInteger="1"' : ''
      return `<cacheField name="${name}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${integerFlag} minValue="${Math.min(...numericValues)}" maxValue="${Math.max(...numericValues)}"/></cacheField>`
    }
    const items = unique.map(value => dateField ? `<d v="${toDateTime(Number(value)).replace(/Z$/, '')}"/>` : `<s v="${esc(String(value))}"/>`).join('')
    if (dateField) {
      const dates = unique.map(value => toDateTime(Number(value)).replace(/Z$/, '')).sort()
      return `<cacheField name="${name}" numFmtId="176"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="${dates[0]}" maxDate="${dates.at(-1)}" count="${unique.length}">${items}</sharedItems></cacheField>`
    }
    return `<cacheField name="${name}" numFmtId="0"><sharedItems count="${unique.length}">${items}</sharedItems></cacheField>`
  }).join('')
  const records = values.map(row => `<r>${row.map((value, i) => i >= 3
    ? `<n v="${value}"/>`
    : `<x v="${[...new Set(values.map(r => r[i]))].indexOf(value)}"/>`).join('')}</r>`).join('')
  const fields = [
    `<pivotField axis="axisRow" showAll="0"><items count="5">${Array.from({ length: 4 }, (_, j) => `<item x="${j}"/>`).join('')}<item t="default"/></items></pivotField>`,
    '<pivotField showAll="0"><items count="5"><item x="2"/><item h="1" x="1"/><item x="3"/><item h="1" x="0"/><item t="default"/></items></pivotField>',
    `<pivotField numFmtId="176" showAll="0"><items count="40">${Array.from({ length: 39 }, (_, j) => `<item x="${j}"/>`).join('')}<item t="default"/></items></pivotField>`,
    '<pivotField dataField="1" showAll="0"/>',
    '<pivotField showAll="0"/>',
  ].join('')
  const rowItems = `<i><x/></i>${Array.from({ length: 3 }, (_, i) => `<i><x v="${i + 1}"/></i>`).join('')}<i t="grand"><x/></i>`
  const x14Cache = `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS.x14}"><x14:pivotCacheDefinition pivotCacheId="1"/></ext></extLst>`
  const pivotTableExt = `<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="${NS.x14}"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="${NS.xm}"/></ext><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst>`
  return {
    'xl/pivotCache/pivotCacheDefinition1.xml': `${XML}<pivotCacheDefinition xmlns="${NS.main}" xmlns:r="${NS.rel}" r:id="rId1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${values.length}" saveData="1"><cacheSource type="worksheet"><worksheetSource ref="A1:E${values.length + 1}" sheet="Data"/></cacheSource><cacheFields count="5">${cacheFields}</cacheFields>${x14Cache}</pivotCacheDefinition>`,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.cacheRecordsRel}" Target="pivotCacheRecords1.xml"/></Relationships>`,
    'xl/pivotCache/pivotCacheRecords1.xml': `${XML}<pivotCacheRecords xmlns="${NS.main}" count="${values.length}">${records}</pivotCacheRecords>`,
    'xl/pivotTables/pivotTable1.xml': `${XML}<pivotTableDefinition xmlns="${NS.main}" xmlns:r="${NS.rel}" name="PivotTable1" cacheId="1" dataCaption="Values" compact="1" outline="1" subtotalTop="1" updatedVersion="8" minRefreshableVersion="5" createdVersion="8" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" useAutoFormatting="1" itemPrintTitles="1" indent="0" outlineData="1" multipleFieldFilters="0"><location ref="G1:H6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="5">${fields}</pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="5">${rowItems}</rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Sales" fld="3" subtotal="sum"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0"/>${pivotTableExt}</pivotTableDefinition>`,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.pivotCacheRel}" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`,
  }
}

function workbookFiles(c: RefCase): Record<string, string> {
  const tableCase = c.kind === 'table' ? c.spec as TableCase : undefined
  const control = c.kind === 'slicer' || c.kind === 'timeline'
  const controlXml = control ? controlParts(c) : {}
  const contentOverrides = control ? [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
    '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>',
    '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>',
    '<Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>',
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
    '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>',
    '<Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ].join('') : [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    ...(tableCase ? ['<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'] : []),
  ].join('')
  const mainSheet = c.kind === 'table' ? 'Data' : 'Report'
  const sourceSheets = control ? ['Report', 'Data'] : [mainSheet]
  const sheetFiles: Record<string, string> = {}
  sourceSheets.forEach((name, i) => {
    const isReport = name === 'Report'
    sheetFiles[`xl/worksheets/sheet${i + 1}.xml`] = worksheetXml(!control && i === 0 ? tableCase : undefined, control && isReport ? c.kind as 'slicer' | 'timeline' : undefined, control && name === 'Data')
    const rels = [
      ...(tableCase && i === 0 ? ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'] : []),
      ...(control && isReport ? [
        `<Relationship Id="rId3" Type="${NS.slicerRel}" Target="../slicers/slicer1.xml"/>`,
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>',
        `<Relationship Id="rId1" Type="${NS.pivotTableRel}" Target="../pivotTables/pivotTable1.xml"/>`,
        `<Relationship Id="rId4" Type="${NS.timelineRel}" Target="../timelines/timeline1.xml"/>`,
      ] : []),
    ].join('')
    if (rels) sheetFiles[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = `${XML}<Relationships xmlns="${NS.pkg}">${rels}</Relationships>`
  })
  const workbookRels = control
    ? `<Relationship Id="rId8" Type="${NS.rel}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="${NS.pivotCacheRel}" Target="pivotCache/pivotCacheDefinition1.xml"/><Relationship Id="rId7" Type="${NS.rel}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${NS.rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId1" Type="${NS.rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId6" Type="${NS.rel}/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="${NS.timelineCacheRel}" Target="timelineCaches/timelineCache1.xml"/><Relationship Id="rId4" Type="${NS.slicerCacheRel}" Target="slicerCaches/slicerCache1.xml"/>`
    : sourceSheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${NS.rel}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
  const workbookExt = control ? `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${NS.x14}"><x14:slicerCaches><x14:slicerCache r:id="rId4"/></x14:slicerCaches></ext><ext uri="{79F54976-1DA5-4618-B147-4CDE4B953A38}" xmlns:x14="${NS.x14}"><x14:workbookPr/></ext><ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${NS.x15}"><x15:timelineCacheRefs><x15:timelineCacheRef r:id="rId5"/></x15:timelineCacheRefs></ext></extLst>` : ''
  return {
    '[Content_Types].xml': `${XML}<Types xmlns="${NS.content}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${contentOverrides}</Types>`,
    '_rels/.rels': control
      ? `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${NS.rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
      : `${XML}<Relationships xmlns="${NS.pkg}"><Relationship Id="rId1" Type="${NS.rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': control
      ? `${XML}<workbook xmlns="${NS.main}" xmlns:r="${NS.rel}" xmlns:mc="${NS.mc}" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="${NS.x15}" xmlns:xr="${NS.xr}" xmlns:xr6="${NS.xr6}" xmlns:xr10="${NS.xr10}" xmlns:xr2="${NS.xr2}"><fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="11011"/><workbookPr defaultThemeVersion="124226"/><mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="x15"><x15ac:absPath url="/" xmlns:x15ac="${NS.x15ac}"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="8_{00000000-0000-0000-0000-000000000000}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="240" yWindow="600" windowWidth="16100" windowHeight="9660" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets><sheet name="Report" sheetId="2" r:id="rId1"/><sheet name="Data" sheetId="1" r:id="rId2"/></sheets><definedNames><definedName name="NativeTimeline_Date">#N/A</definedName><definedName name="Slicer_Product">#N/A</definedName></definedNames><calcPr calcId="124519"/><pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches>${workbookExt}</workbook>`
      : `${XML}<workbook xmlns="${NS.main}" xmlns:r="${NS.rel}"><workbookPr/><sheets>${sourceSheets.map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML}<Relationships xmlns="${NS.pkg}">${workbookRels}</Relationships>`,
    ...(control ? {
      'xl/sharedStrings.xml': `${XML}<sst xmlns="${NS.main}" count="0" uniqueCount="0"/>`,
      'xl/theme/theme1.xml': `${XML}<a:theme xmlns:a="${NS.a}" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>${Array.from({ length: 6 }, (_, i) => `<a:accent${i + 1}><a:srgbClr val="4472C4"/></a:accent${i + 1}>`).join('')}<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
      'docProps/core.xml': `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>OfficeView</dc:creator><cp:lastModifiedBy>OfficeView</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-29T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-29T00:00:00Z</dcterms:modified></cp:coreProperties>`,
      'docProps/app.xml': `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Report</vt:lpstr><vt:lpstr>Data</vt:lpstr></vt:vector></TitlesOfParts><Company/><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.114</AppVersion></Properties>`,
    } : {}),
    'xl/styles.xml': `${XML}<styleSheet xmlns="${NS.main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${tableCase?.custom ? '<dxfs count="1"><dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="OfficeViewCustom" pivot="0" table="1" count="2"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="0"/></tableStyle></tableStyles>' : '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'}</styleSheet>`,
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
const invokedScript = process.argv[1] && /scripts[\\/]xlsx-refs[\\/]generate\.(?:ts|js)$/.test(process.argv[1])
  ? resolve(process.argv[1])
  : resolve(process.cwd(), 'scripts/xlsx-refs/generate.ts')
export const FIXTURE_DIR = resolve(dirname(invokedScript), '../../tests/fixtures/xlsx-pivots')
export async function writeFixtures(): Promise<number> {
  mkdirSync(FIXTURE_DIR, { recursive: true })
  let count = 0
  for (const c of REF_CASES) if (c.build === 'xml') {
    writeFileSync(resolve(FIXTURE_DIR, `${c.id}.xlsx`), await buildCaseWorkbook(c))
    count++
  }
  return count
}

if (process.argv[1] && /scripts[\\/]xlsx-refs[\\/]generate\.(?:ts|js)$/.test(process.argv[1])) {
  console.log(`wrote ${await writeFixtures()} pivot reference workbooks to ${FIXTURE_DIR}`)
}
