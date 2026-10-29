import { readFileSync, readdirSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXml, parseXmlOrdered } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'
import type { PivotRegionRef } from '../src/xlsx/pivots/stub'
import { loadPivotWorkbook, type PackageReader } from '../src/xlsx/pivots/load'
import { MAX_UNSUPPORTED_ENTRIES, UNSUPPORTED_TRUNCATED } from '../src/xlsx/pivots/parseCacheDefinition'
import { PIVOT_BASE_ITEM_DEFAULT, PIVOT_BASE_ITEM_NEXT, PIVOT_BASE_ITEM_PREVIOUS, PIVOT_VALUES_FIELD, type PivotWorkbook } from '../src/xlsx/pivots/model'

const root = process.cwd()
const corpusFiles = readdirSync(resolve(root, 'corpus')).filter(f => /pivot/i.test(f) && f.endsWith('.xlsx')).map(f => `corpus/${f}`)
const fixtureFiles = readdirSync(resolve(root, 'tests/fixtures/xlsx-pivots')).filter(f => /^(pivot|slicer|timeline)-.*\.xlsx$/.test(f)).map(f => `tests/fixtures/xlsx-pivots/${f}`)

async function loadFile(path: string): Promise<{ book: PivotWorkbook; regions: Map<string, PivotRegionRef[]> }> {
  const pkg = await OfficePackage.load(new Uint8Array(readFileSync(resolve(root, path))))
  const doc = await parseXlsx(pkg)
  const regions = new Map(doc.sheets.filter(s => s.pivotRegions?.length).map(s => [s.name, s.pivotRegions!]))
  return { book: await loadPivotWorkbook(pkg, regions), regions }
}

/** The structural PackageReader over in-memory parts, as the worker's RangePackage would serve them. */
function memoryPackage(parts: Record<string, string>): PackageReader {
  return {
    has: path => path in parts,
    xml: async path => path in parts ? parseXml(parts[path]) : undefined,
    xmlOrdered: async path => path in parts ? parseXmlOrdered(parts[path]) : undefined,
  }
}

async function zipParts(path: string): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(readFileSync(resolve(root, path)))
  const out: Record<string, string> = {}
  for (const name of Object.keys(zip.files)) if (name.endsWith('.xml') || name.endsWith('.rels')) out[name] = await zip.files[name].async('string')
  return out
}

const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (...items: Array<[string, string, string]>) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`
const region: PivotRegionRef = { partPath: 'xl/pivotTables/pivotTable1.xml', ref: 'A3:B6' }

/** A minimal one-cache, one-table package; `cache` and `table` are the inner XML of their roots. */
function synthetic(opts: { cacheAttrs?: string; cache: string; tableAttrs?: string; table: string; tableRaw?: string; cacheRaw?: string }): PackageReader {
  return memoryPackage({
    'xl/workbook.xml': `<workbook ${MAIN}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches></workbook>`,
    'xl/_rels/workbook.xml.rels': rels(['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml']),
    'xl/pivotCache/pivotCacheDefinition1.xml': opts.cacheRaw ?? `<pivotCacheDefinition ${MAIN} r:id="rId1" ${opts.cacheAttrs ?? ''}>${opts.cache}</pivotCacheDefinition>`,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': rels(['rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml']),
    'xl/pivotTables/pivotTable1.xml': opts.tableRaw ?? `<pivotTableDefinition ${MAIN} name="P" cacheId="7" dataCaption="Values" ${opts.tableAttrs ?? ''}>${opts.table}</pivotTableDefinition>`,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': rels(['rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml']),
  })
}
const loadSynthetic = (opts: Parameters<typeof synthetic>[0]) => loadPivotWorkbook(synthetic(opts), new Map([['S', [region]]]))
const SOURCE = '<cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="Data"/></cacheSource>'
const TWO_FIELDS = `${SOURCE}<cacheFields count="2"><cacheField name="K" numFmtId="0"><sharedItems count="2"><s v="a"/><s v="b"/></sharedItems></cacheField><cacheField name="V" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1" maxValue="2"/></cacheField></cacheFields>`
const LOCATION = '<location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'

describe('pivot definition parsers', () => {
  it.each([...corpusFiles, ...fixtureFiles])('parses %s', async (path) => {
    const { book, regions } = await loadFile(path)
    const listed = [...regions.values()].flat()
    expect(book.diagnostics.filter(d => d.kind === 'parse')).toEqual([])
    expect(book.tables.map(t => t.partPath).sort()).toEqual(listed.map(r => r.partPath).sort())
    for (const table of book.tables) {
      expect(table.location.ref).toBe(table.region.ref)
      const cache = book.caches.get(table.cacheId)
      expect(cache, `cache ${table.cacheId}`).toBeDefined()
      expect(table.fields.length).toBe(cache!.fields.length)
      // Excel, LibreOffice and POI write nothing outside the model; only the hand-written slicer/timeline XML does (its table-level subtotalTop).
      if (!/(slicer|timeline)-/.test(path)) expect(table.unsupported).toEqual([])
    }
  })

  it('reads the Excel-built compact pivot exactly', async () => {
    const { book } = await loadFile('tests/fixtures/xlsx-pivots/pivot-layout-compact.xlsx')
    expect(book.tables).toHaveLength(1)
    const t = book.tables[0]
    expect(t).toMatchObject({ sheet: 'Pivot', partPath: 'xl/pivotTables/pivotTable1.xml', cacheId: 2045, cachePartPath: 'xl/pivotCache/pivotCacheDefinition1.xml', name: '樞紐分析表2', dataCaption: 'Values', grandTotalCaption: 'Grand Total', rowHeaderCaption: 'Row Labels', colHeaderCaption: 'Column Labels' })
    expect(t.location).toEqual({ ref: 'A3:I22', firstHeaderRow: 1, firstDataRow: 3, firstDataCol: 1, rowPageCount: 0, colPageCount: 0 })
    // Excel omits schema defaults: compact form writes outline="1" but no compact, and indent="0".
    expect([t.compact, t.outline, t.outlineData, t.compactData, t.indent, t.multipleFieldFilters, t.useAutoFormatting]).toEqual([true, true, true, true, 0, false, true])
    expect(t.rowFields).toEqual([0, 1])
    expect(t.colFields).toEqual([2, PIVOT_VALUES_FIELD])
    expect(t.rowItems).toHaveLength(17)
    expect(t.rowItems![0]).toEqual({ type: 'data', r: 0, i: 0, x: [0] })
    expect(t.rowItems![2]).toEqual({ type: 'data', r: 1, i: 0, x: [1] })
    expect(t.rowItems![16]).toEqual({ type: 'grand', r: 0, i: 0, x: [0] })
    expect(t.colItems![1]).toEqual({ type: 'data', r: 1, i: 1, x: [1] })
    expect(t.colItems![7]).toEqual({ type: 'grand', r: 0, i: 1, x: [0] })
    expect(t.fields[0]).toMatchObject({ axis: 'axisRow', showAll: false, compact: true, outline: true, subtotalTop: true, defaultSubtotal: true })
    expect(t.fields[0].items!.map(i => i.type === 'data' ? i.x : i.type)).toEqual([3, 1, 0, 2, 'default'])
    expect(t.fields[3]).toMatchObject({ dataField: true })
    expect(t.fields[3].axis).toBeUndefined()
    expect(t.dataFields).toEqual([
      { name: 'Sum of Sales', fld: 3, subtotal: 'sum', showDataAs: 'normal', baseField: 0, baseItem: 0 },
      { name: 'Sum of Units', fld: 4, subtotal: 'sum', showDataAs: 'normal', baseField: 0, baseItem: 0 },
    ])
    expect(t.styleInfo).toEqual({ name: 'PivotStyleLight16', showRowHeaders: true, showColHeaders: true, showRowStripes: false, showColStripes: false, showLastColumn: true })
    expect(t.x14?.hideValuesRow).toBe(true)
    expect(t.unsupported).toEqual([])

    const cache = book.caches.get(2045)!
    expect(cache).toMatchObject({ partPath: 'xl/pivotCache/pivotCacheDefinition1.xml', recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml', recordCount: 36, saveData: true, refreshOnLoad: false, olap: false, createdVersion: 8, refreshedVersion: 8, minRefreshableVersion: 3 })
    expect(cache.source).toEqual({ type: 'worksheet', connectionId: 0, worksheet: { ref: 'A1:E37', sheet: 'Data' } })
    expect(cache.fields.map(f => f.name)).toEqual(['Region', 'Product', 'Year', 'Sales', 'Units'])
    expect(cache.unsupported).toEqual([])
  })

  it('tells shared item lists from inline record values', async () => {
    const { book } = await loadFile('tests/fixtures/xlsx-pivots/pivot-layout-compact.xlsx')
    const [region, , year, sales] = book.caches.get(2045)!.fields
    expect(region.sharedItems).toMatchObject({ inline: false, containsString: true, containsNumber: false })
    expect(region.sharedItems).not.toHaveProperty('count')
    expect(region.sharedItems!.items).toEqual([{ type: 'string', value: 'South' }, { type: 'string', value: 'North' }, { type: 'string', value: 'West' }, { type: 'string', value: 'East' }])
    expect(year.sharedItems).toMatchObject({ inline: false, containsInteger: true, minValue: 2022, maxValue: 2024 })
    expect(year.sharedItems!.items).toEqual([{ type: 'number', value: 2022 }, { type: 'number', value: 2024 }, { type: 'number', value: 2023 }])
    expect(sales.sharedItems).toMatchObject({ inline: true, items: [], containsSemiMixedTypes: false, containsString: false, containsNumber: true, containsInteger: false, minValue: 766.05, maxValue: 8954.09 })
  })

  it('keeps mixed shared items in document order (records index them by position)', async () => {
    const book = await loadSynthetic({
      cache: `${SOURCE}<cacheFields count="1"><cacheField name="K"><sharedItems containsBlank="1" containsMixedTypes="1" count="6"><s v="a"/><m/><n v="1.5"/><b v="1"/><e v="#N/A"/><d v="2020-01-02T00:00:00" u="1"/></sharedItems></cacheField></cacheFields>`,
      table: `${LOCATION}<pivotFields count="1"><pivotField/></pivotFields>`,
    })
    expect(book.caches.get(7)!.fields[0].sharedItems!.items).toEqual([
      { type: 'string', value: 'a' }, { type: 'missing' }, { type: 'number', value: 1.5 }, { type: 'boolean', value: true },
      { type: 'error', value: '#N/A' }, { type: 'date', value: '2020-01-02T00:00:00', unused: true },
    ])
  })

  it('pins ECMA-376 attribute defaults, not writer habits', async () => {
    const book = await loadSynthetic({
      cache: `${SOURCE}<cacheFields count="2"><cacheField name="K"><sharedItems><s v="a"/></sharedItems><fieldGroup base="0"><rangePr/></fieldGroup></cacheField><cacheField name="V"/></cacheFields>`,
      table: `<location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow"><items count="1"><item x="0"/></items></pivotField><pivotField dataField="1"/></pivotFields>`
        + '<rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField fld="1"/></dataFields>'
        + '<formats count="1"><format dxfId="0"><pivotArea><references count="1"><reference field="0"><x v="0"/></reference></references></pivotArea></format></formats><pivotTableStyleInfo name="PivotStyleLight16"/>',
    })
    const t = book.tables[0], cache = book.caches.get(7)!
    // Table level: compact true, outline false (§18.10.1.73), indent 1, both grand totals on.
    expect(t).toMatchObject({ compact: true, outline: false, compactData: true, outlineData: false, indent: 1, rowGrandTotals: true, colGrandTotals: true, dataOnRows: false, showDrill: true, printDrill: false, showHeaders: true, multipleFieldFilters: true, useAutoFormatting: false, mergeItem: false, showMissing: true, showError: false, gridDropZones: false, customListSort: true, chartFormat: 0, applyNumberFormats: false })
    // Field level: compact, outline and subtotalTop all default true.
    expect(t.fields[0]).toMatchObject({ compact: true, outline: true, subtotalTop: true, showAll: true, defaultSubtotal: true, subtotals: [], sortType: 'manual', itemPageCount: 10, insertBlankRow: false, dataField: false, topAutoShow: true, autoShow: false, multipleItemSelectionAllowed: false })
    expect(t.fields[0].items![0]).toEqual({ type: 'data', x: 0, hidden: false, char: false, showDetails: true, calculated: false, missing: false, childItems: false, expanded: false, drillAcross: true })
    expect(t.rowItems![0]).toEqual({ type: 'data', r: 0, i: 0, x: [0] })
    expect(t.colItems![0]).toEqual({ type: 'data', r: 0, i: 0, x: [] })
    expect(t.dataFields[0]).toEqual({ fld: 1, subtotal: 'sum', showDataAs: 'normal', baseField: -1, baseItem: PIVOT_BASE_ITEM_DEFAULT })
    expect(t.formats[0]).toEqual({ action: 'formatting', dxfId: 0, pivotArea: { type: 'normal', dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [{ field: 0, selected: true, byPosition: false, relative: false, defaultSubtotal: false, subtotals: [], x: [0] }] } })
    expect(t.styleInfo).toEqual({ name: 'PivotStyleLight16', showRowHeaders: false, showColHeaders: false, showRowStripes: false, showColStripes: false, showLastColumn: false })
    expect(cache).toMatchObject({ saveData: true, refreshOnLoad: false, enableRefresh: true, invalid: false, createdVersion: 0, refreshedVersion: 0, minRefreshableVersion: 0, upgradeOnRefresh: false, olap: false })
    expect(cache.fields[0]).toMatchObject({ databaseField: true, uniqueList: true, serverField: false, sqlType: 0, hierarchy: 0, level: 0, memberPropertyField: false })
    expect(cache.fields[0].sharedItems).toMatchObject({ containsSemiMixedTypes: true, containsNonDate: true, containsString: true, containsDate: false, containsBlank: false, containsMixedTypes: false, containsNumber: false, containsInteger: false, longText: false, inline: false })
    expect(cache.fields[0].fieldGroup).toEqual({ base: 0, rangePr: { autoStart: true, autoEnd: true, groupBy: 'range', groupInterval: 1 } })
    expect(cache.fields[1].sharedItems).toBeUndefined()
    expect(t.unsupported).toEqual([])
  })

  it('reads table-level overrides LibreOffice writes', async () => {
    const { book } = await loadFile('corpus/lo-calc-pivottable_date_field_filter.xlsx')
    const t = book.tables[0]
    expect(t).toMatchObject({ compact: false, compactData: false, outline: false, dataOnRows: true, rowGrandTotals: false, colGrandTotals: false, gridDropZones: true, enableDrill: false, showItems: false })
    expect(t.fields[3]).toMatchObject({ compact: false, outline: false, subtotalTop: false, defaultSubtotal: false, sortType: 'ascending', itemPageCount: 0, includeNewItemsInFilter: true })
    expect(t.fields[5].items![1]).toMatchObject({ x: 1, hidden: true })
    expect(t.colItems).toEqual([{ type: 'data', r: 0, i: 0, x: [] }])
    expect(t.styleInfo?.name).toBeUndefined()
  })

  it('reads page fields, data pivot areas and grouping', async () => {
    const poi = (await loadFile('corpus/poi-ExcelPivotTableSample.xlsx')).book
    const paged = poi.tables.find(t => t.cacheId === 2)!
    expect(paged.location).toMatchObject({ rowPageCount: 1, colPageCount: 1 })
    expect(paged.pageFields).toEqual([{ fld: 2, hier: -1 }])
    expect(paged.fields[2]).toMatchObject({ axis: 'axisPage', multipleItemSelectionAllowed: true, numFmtId: 16 })
    expect(poi.caches.get(2)!.source.worksheet).toEqual({ name: 'Tabelle1' })
    expect(poi.caches.get(0)!.fields[2].sharedItems).toMatchObject({ containsDate: true, containsNonDate: false, minDate: '2022-01-01T00:00:00', maxDate: '2022-01-04T00:00:00' })

    const formats = (await loadFile('tests/fixtures/xlsx-pivots/pivot-formats-override.xlsx')).book.tables[0].formats
    expect(formats[0].pivotArea).toMatchObject({ collapsedLevelsAreSubtotals: true, fieldPosition: 0 })
    expect(formats[0].pivotArea.references.map(r => [r.field, r.selected, r.x])).toEqual([[PIVOT_VALUES_FIELD, false, [0]], [0, false, [0]], [1, true, [0]], [2, false, [0]]])

    const lo = (await loadFile('corpus/lo-calc-PivotTableWithNoSourceData.xlsx')).book
    const [day, month] = lo.caches.get(0)!.fields
    expect(day.fieldGroup).toMatchObject({ par: 1, base: 0, rangePr: { autoStart: true, autoEnd: false, groupBy: 'days', startDate: '2015-05-12T00:00:00', endDate: '2015-07-08T00:00:00', groupInterval: 1 } })
    expect(day.fieldGroup!.groupItems).toHaveLength(368)
    expect(day.fieldGroup!.groupItems![0]).toEqual({ type: 'string', value: '<5/12/2015' })
    expect(month.fieldGroup).toMatchObject({ base: 0, rangePr: { groupBy: 'months' } })
    expect(month.sharedItems).toBeUndefined()
    expect(lo.tables[0].chartFormats[0]).toMatchObject({ chart: 0, format: 1, series: true, pivotArea: { type: 'data', outline: false, references: [{ field: PIVOT_VALUES_FIELD, selected: false, x: [0] }] } })

    const repeat = (await loadFile('tests/fixtures/xlsx-pivots/pivot-repeat-labels.xlsx')).book.tables[0]
    expect(repeat.fields.some(f => f.x14?.fillDownLabels)).toBe(true)
    expect(repeat.unsupported).toEqual([])
  })

  it('keeps the previous/next base-item sentinels and x14 show-as', async () => {
    const book = await loadSynthetic({
      cache: TWO_FIELDS,
      table: `${LOCATION}<pivotFields count="2"><pivotField axis="axisRow"><items count="2"><item x="0"/><item x="1" sd="0" h="1" n="Bee"/></items></pivotField><pivotField dataField="1"/></pivotFields>`
        + '<dataFields count="3"><dataField name="Prev" fld="1" showDataAs="difference" baseField="0" baseItem="1048828"/><dataField name="Next" fld="1" showDataAs="percentDiff" baseField="0" baseItem="1048829"/>'
        + '<dataField name="Parent" fld="1" subtotal="average" numFmtId="10"><extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataField pivotShowAs="percentOfParentRow"/></ext></extLst></dataField></dataFields>',
    })
    const t = book.tables[0]
    expect(t.dataFields.map(d => [d.showDataAs, d.baseItem])).toEqual([['difference', PIVOT_BASE_ITEM_PREVIOUS], ['percentDiff', PIVOT_BASE_ITEM_NEXT], ['normal', PIVOT_BASE_ITEM_DEFAULT]])
    expect(t.dataFields[2]).toMatchObject({ subtotal: 'average', numFmtId: 10, x14: { pivotShowAs: 'percentOfParentRow' } })
    expect(t.fields[0].items![1]).toMatchObject({ x: 1, showDetails: false, hidden: true, name: 'Bee' })
    expect(t.unsupported).toEqual([])
  })

  it('reads filters, conditional formats, sort scope and discrete grouping', async () => {
    const book = await loadSynthetic({
      cache: `${SOURCE}<cacheFields count="3"><cacheField name="K"><sharedItems count="2"><s v="a"/><s v="b"/></sharedItems></cacheField>`
        + '<cacheField name="V"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1" maxValue="2"/></cacheField>'
        + '<cacheField name="K2" databaseField="0"><fieldGroup base="0"><discretePr count="2"><x v="1"/><x/></discretePr><groupItems count="2"><s v="Group1"/><s v="Group2"/></groupItems></fieldGroup></cacheField></cacheFields>',
      table: `${LOCATION}<pivotFields count="3"><pivotField axis="axisRow" sortType="descending"><items count="2"><item x="0"/><item x="1"/></items>`
        + '<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></autoSortScope></pivotField>'
        + '<pivotField dataField="1"/><pivotField/></pivotFields>'
        + '<conditionalFormats count="1"><conditionalFormat scope="data" type="row" priority="2"><pivotAreas count="1"><pivotArea type="data" outline="0" collapsedLevelsAreSubtotals="1" fieldPosition="0"><references count="1"><reference field="0" count="1" selected="0"><x v="1"/></reference></references></pivotArea></pivotAreas></conditionalFormat></conditionalFormats>'
        + '<filters count="3">'
        + '<filter fld="0" type="captionBeginsWith" evalOrder="-1" id="1" stringValue1="a"><autoFilter ref="A1"><filterColumn colId="0"><customFilters><customFilter val="a*"/></customFilters></filterColumn></autoFilter></filter>'
        + '<filter fld="0" type="count" id="2" iMeasureFld="0"><autoFilter ref="A1"><filterColumn colId="0"><top10 val="3" filterVal="3"/></filterColumn></autoFilter></filter>'
        + '<filter fld="0" type="thisMonth" evalOrder="1" id="3"><autoFilter ref="A1"><filterColumn colId="0"><dynamicFilter type="thisMonth"/></filterColumn></autoFilter></filter>'
        + '</filters>',
    })
    const t = book.tables[0], cache = book.caches.get(7)!
    expect(cache.fields[2]).toMatchObject({ databaseField: false, fieldGroup: { base: 0, discretePr: [1, 0], groupItems: [{ type: 'string', value: 'Group1' }, { type: 'string', value: 'Group2' }] } })
    expect(t.fields[0].autoSortScope).toMatchObject({ dataOnly: false, outline: false, fieldPosition: 0, references: [{ field: PIVOT_VALUES_FIELD, selected: false, x: [0] }] })
    expect(t.conditionalFormats).toEqual([{ scope: 'data', type: 'row', priority: 2, pivotAreas: [expect.objectContaining({ type: 'data', outline: false, collapsedLevelsAreSubtotals: true, references: [expect.objectContaining({ field: 0, x: [1] })] })] }])
    expect(t.filters).toEqual([
      { fld: 0, type: 'captionBeginsWith', evalOrder: -1, id: 1, stringValue1: 'a', autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, customFilters: { and: false, filters: [{ operator: 'equal', val: 'a*' }] } }] } },
      { fld: 0, type: 'count', evalOrder: 0, id: 2, iMeasureFld: 0, autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: false, val: 3, filterVal: 3 } }] } },
      { fld: 0, type: 'thisMonth', evalOrder: 1, id: 3, autoFilter: { ref: 'A1', columns: [{ colId: 0, hiddenButton: false, showButton: true, dynamicFilter: { type: 'thisMonth' } }] } },
    ])
    expect(t.unsupported).toEqual([])
  })

  it('bounds the unsupported list and does not mistake prototype names for known attributes', async () => {
    const bad = Array.from({ length: MAX_UNSUPPORTED_ENTRIES + 50 }, (_, i) => `<n v="bad${i}"/>`).join('')
    const book = await loadSynthetic({
      cache: `${SOURCE}<cacheFields count="1"><cacheField name="K" valueOf="1"><sharedItems count="${MAX_UNSUPPORTED_ENTRIES + 50}">${bad}</sharedItems></cacheField></cacheFields>`,
      table: `${LOCATION}<pivotFields count="1"><pivotField toString="1"/></pivotFields>`,
    })
    const cache = book.caches.get(7)!
    expect(cache.fields[0].sharedItems!.items).toHaveLength(MAX_UNSUPPORTED_ENTRIES + 50)
    expect(cache.unsupported).toHaveLength(MAX_UNSUPPORTED_ENTRIES + 1)
    // fast-xml-parser may rename prototype names (valueOf → __valueOf); either way the attribute is recorded, never taken as known.
    expect(cache.unsupported.filter(e => /^cacheField@(__)?valueOf$/.test(e))).toHaveLength(1)
    expect(cache.unsupported.at(-1)).toBe(UNSUPPORTED_TRUNCATED)
    expect(book.tables[0].unsupported.filter(e => /^pivotField@(__)?toString$/.test(e))).toHaveLength(1)
  })

  it('parses a part without rowItems as rowItems undefined (Review Focus 5)', async () => {
    const parts = await zipParts('tests/fixtures/xlsx-pivots/pivot-layout-compact.xlsx')
    const path = 'xl/pivotTables/pivotTable1.xml'
    parts[path] = parts[path].replace(/<rowItems\b.*?<\/rowItems>/s, '')
    const book = await loadPivotWorkbook(memoryPackage(parts), new Map([['Pivot', [{ partPath: path, ref: 'A3:I22' }]]]))
    expect(book.tables).toHaveLength(1)
    expect(book.tables[0].rowItems).toBeUndefined()
    expect(book.tables[0].colItems).toHaveLength(8)
    expect(book.diagnostics).toEqual([])
  })

  it('flags an OLAP cache', async () => {
    const book = await loadSynthetic({
      cacheAttrs: 'refreshOnLoad="1" saveData="0"',
      cache: '<cacheSource type="external" connectionId="3"/><cacheFields count="1"><cacheField name="[Measures].[X]" numFmtId="0" hierarchy="1" level="32767"/></cacheFields><cacheHierarchies count="1"><cacheHierarchy uniqueName="[Measures].[X]" caption="X" measure="1" count="0"/></cacheHierarchies>',
      table: `${LOCATION}<pivotFields count="1"><pivotField dataField="1"/></pivotFields>`,
    })
    const cache = book.caches.get(7)!
    expect(cache).toMatchObject({ olap: true, refreshOnLoad: true, saveData: false, source: { type: 'external', connectionId: 3 } })
    expect(cache.fields[0]).toMatchObject({ hierarchy: 1, level: 32767 })
    expect(cache.unsupported).toEqual([])
    expect(book.tables).toHaveLength(1)
  })

  it('records unknown constructs instead of throwing', async () => {
    const book = await loadSynthetic({
      cacheAttrs: 'futureFlag="1"',
      cache: `${TWO_FIELDS}<futureThing/><extLst><ext uri="{00000000-AAAA-BBBB-CCCC-000000000000}"><x99:thing xmlns:x99="urn:x"/></ext></extLst>`,
      tableAttrs: 'subtotalTop="1" outline="maybe"',
      table: `${LOCATION}<pivotFields count="2"><pivotField axis="axisSideways" sortType="random"><items count="1"><item x="0" t="bogus" zz="1"/></items></pivotField><pivotField dataField="1"/></pivotFields>`
        + '<dataFields count="1"><dataField fld="1" subtotal="median" baseItem="-3"/></dataFields><pivotHierarchies count="0"/><mystery/>',
    })
    const t = book.tables[0]
    expect(t.outline).toBe(false)
    expect(t.fields[0].axis).toBeUndefined()
    expect(t.fields[0].sortType).toBe('manual')
    expect(t.dataFields[0].subtotal).toBe('sum')
    expect(t.unsupported).toEqual(expect.arrayContaining([
      'pivotTableDefinition@subtotalTop', 'pivotTableDefinition@outline=maybe', 'pivotField@axis=axisSideways', 'pivotField@sortType=random',
      'item@t=bogus', 'item@zz', 'dataField@subtotal=median', 'dataField@baseItem=-3', 'pivotTableDefinition/pivotHierarchies', 'pivotTableDefinition/mystery',
      'cache:pivotCacheDefinition@futureFlag', 'cache:pivotCacheDefinition/futureThing', 'cache:ext{00000000-AAAA-BBBB-CCCC-000000000000}',
    ]))
    expect(new Set(t.unsupported).size).toBe(t.unsupported.length)
    expect(book.diagnostics.filter(d => d.kind === 'parse')).toEqual([])
  })

  it('skips an unreadable definition and keeps a table whose cache is unreadable', async () => {
    const broken = await loadSynthetic({ cache: TWO_FIELDS, table: '', tableRaw: '<pivotTableDefinition><location ref="A1"' })
    expect(broken.tables).toEqual([])
    expect(broken.diagnostics).toEqual([expect.objectContaining({ kind: 'parse', feature: 'pivot-definition-unreadable', partPath: 'xl/pivotTables/pivotTable1.xml' })])

    const noLocation = await loadSynthetic({ cache: TWO_FIELDS, table: '<pivotFields count="0"/>' })
    expect(noLocation.tables).toEqual([])
    expect(noLocation.diagnostics.map(d => d.feature)).toEqual(['pivot-definition-unreadable'])

    const badCache = await loadSynthetic({ cache: '', cacheRaw: '<pivotCacheDefinition><cacheFields>', table: `${LOCATION}<pivotFields count="0"/>` })
    expect(badCache.tables).toHaveLength(1)
    expect(badCache.tables[0]).toMatchObject({ cacheId: 7, cachePartPath: 'xl/pivotCache/pivotCacheDefinition1.xml' })
    expect(badCache.caches.size).toBe(0)
    expect(badCache.diagnostics).toEqual([expect.objectContaining({ kind: 'parse', feature: 'pivot-cache-unreadable', partPath: 'xl/pivotCache/pivotCacheDefinition1.xml' })])
  })

  it('falls back to the definition relationship when the workbook does not list the cache', async () => {
    const pkg = synthetic({ cache: TWO_FIELDS, table: `${LOCATION}<pivotFields count="2"><pivotField/><pivotField/></pivotFields>` })
    const parts: PackageReader = { ...pkg, xml: async path => path === 'xl/workbook.xml' ? parseXml(`<workbook ${MAIN}><sheets/></workbook>`) : pkg.xml(path) }
    const book = await loadPivotWorkbook(parts, new Map([['S', [region]]]))
    expect(book.tables[0].cachePartPath).toBe('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(book.caches.get(7)?.fields).toHaveLength(2)
    expect(book.tables[0].unsupported).toEqual([])
    expect(book.diagnostics).toEqual([expect.objectContaining({ kind: 'parse', feature: 'pivot-cache-mapping-recovered', partPath: 'xl/pivotTables/pivotTable1.xml' })])
  })

  describe('index validation (ruling AC)', () => {
    const FIELDS = `${LOCATION}<pivotFields count="3"><pivotField axis="axisRow"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>`
      + '<pivotField dataField="1"/><pivotField axis="axisPage"><items count="2"><item x="0"/><item x="1"/></items></pivotField></pivotFields>'
    const CACHE3 = `${SOURCE}<cacheFields count="3"><cacheField name="K"><sharedItems count="2"><s v="a"/><s v="b"/></sharedItems></cacheField>`
      + '<cacheField name="V"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"/></cacheField>'
      + '<cacheField name="P"><sharedItems count="2"><s v="p"/><s v="q"/></sharedItems></cacheField></cacheFields>'
    const DATA2 = '<dataFields count="2"><dataField fld="1"/><dataField name="B" fld="1"/></dataFields>'
    const run = (table: string, cache = CACHE3) => loadSynthetic({ cache, table })
    const invalidOf = (book: PivotWorkbook) => book.tables[0].unsupported.filter(e => e.startsWith('invalid-index:'))

    it('accepts a well-formed table with Values on columns', async () => {
      const book = await run(`${FIELDS}<rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>`
        + '<colFields count="1"><field x="-2"/></colFields><colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>'
        + `<pageFields count="1"><pageField fld="2" item="1" hier="-1"/></pageFields>${DATA2}`)
      expect(invalidOf(book)).toEqual([])
      expect(book.diagnostics).toEqual([])
    })

    it.each([
      ['pivot item x beyond the shared items', FIELDS.replace('<item x="1"/>', '<item x="2"/>'), 'pivotField[0].item[1].x'],
      ['row field beyond the fields', `${FIELDS}<rowFields count="1"><field x="3"/></rowFields>`, 'rowFields[0]'],
      ['negative row field', `${FIELDS}<rowFields count="1"><field x="-1"/></rowFields>`, 'rowFields[0]'],
      ['Values on both axes', `${FIELDS}<rowFields count="1"><field x="-2"/></rowFields><colFields count="1"><field x="-2"/></colFields>${DATA2}`, 'colFields[0]'],
      ['Values without data fields', `${FIELDS}<colFields count="1"><field x="-2"/></colFields>`, 'colFields[0]'],
      ['axis item x beyond the field items', `${FIELDS}<rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i><x v="3"/></i></rowItems>`, 'rowItems[0].x[0]'],
      ['negative axis item x', `${FIELDS}<rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i><x v="-1"/></i></rowItems>`, 'rowItems[0].x[0]'],
      ['axis item x beyond the axis levels', `${FIELDS}<rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i r="1"><x/></i></rowItems>`, 'rowItems[0].x[0]'],
      ['Values axis item beyond the data fields', `${FIELDS}<colFields count="1"><field x="-2"/></colFields><colItems count="1"><i><x v="2"/></i></colItems>${DATA2}`, 'colItems[0].x[0]'],
      ['axis item i beyond the data fields', `${FIELDS}<rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i i="2"><x/></i></rowItems>${DATA2}`, 'rowItems[0].i'],
      ['page field item beyond its items', `${FIELDS}<pageFields count="1"><pageField fld="2" item="2"/></pageFields>`, 'pageField[0].item'],
      ['page field on Values', `${FIELDS}<pageFields count="1"><pageField fld="-2"/></pageFields>`, 'pageField[0].fld'],
      ['data field beyond the fields', `${FIELDS}<dataFields count="1"><dataField fld="3"/></dataFields>`, 'dataField[0].fld'],
      ['base field beyond the fields', `${FIELDS}<dataFields count="1"><dataField fld="1" showDataAs="runTotal" baseField="5"/></dataFields>`, 'dataField[0].baseField'],
      ['base item beyond the base field items', `${FIELDS}<dataFields count="1"><dataField fld="1" showDataAs="difference" baseField="0" baseItem="3"/></dataFields>`, 'dataField[0].baseItem'],
    ])('fails closed on %s', async (_, table, where) => {
      const book = await run(table)
      expect(book.tables).toHaveLength(1)
      expect(invalidOf(book)).toContain(`invalid-index:${where}`)
      expect(book.diagnostics).toContainEqual(expect.objectContaining({ kind: 'parse', feature: 'pivot-definition-invalid', partPath: 'xl/pivotTables/pivotTable1.xml' }))
    })

    it('keeps the previous/next sentinels and ignores inert base fields', async () => {
      const book = await run(`${FIELDS}<dataFields count="3"><dataField fld="1" showDataAs="difference" baseField="0" baseItem="1048828"/>`
        + '<dataField fld="1" showDataAs="percent" baseField="0" baseItem="1048829"/><dataField fld="1" baseField="1" baseItem="0"/></dataFields>')
      expect(invalidOf(book)).toEqual([])
    })

    it('checks cache grouping indices and the field count', async () => {
      const grouped = await run(`${LOCATION}<pivotFields count="2"><pivotField/><pivotField/></pivotFields>`,
        `${SOURCE}<cacheFields count="2"><cacheField name="K"><sharedItems count="1"><s v="a"/></sharedItems></cacheField>`
        + '<cacheField name="G" databaseField="0"><fieldGroup base="4"><discretePr count="1"><x v="1"/></discretePr><groupItems count="1"><s v="g"/></groupItems></fieldGroup></cacheField></cacheFields>')
      expect(invalidOf(grouped)).toEqual(expect.arrayContaining(['invalid-index:cache.fieldGroup[1].base', 'invalid-index:cache.fieldGroup[1].discretePr[0]']))
      const short = await run(`${LOCATION}<pivotFields count="1"><pivotField/></pivotFields>`)
      expect(invalidOf(short)).toEqual(['invalid-index:pivotFields'])
    })
  })

  it('reports declared counts that disagree with the children, which win', async () => {
    const book = await loadSynthetic({
      cache: `${SOURCE}<cacheFields count="1"><cacheField name="K"><sharedItems count="999999999"><s v="a"/></sharedItems></cacheField></cacheFields>`,
      table: `${LOCATION}<pivotFields count="3"><pivotField><items count="7"><item x="0"/></items></pivotField></pivotFields><rowItems count="0"><i/></rowItems>`,
    })
    expect(book.caches.get(7)!.fields[0].sharedItems!.items).toHaveLength(1)
    expect(book.caches.get(7)!.fields[0].sharedItems).not.toHaveProperty('count')
    expect(book.tables[0].fields).toHaveLength(1)
    expect(book.tables[0].rowItems).toHaveLength(1)
    expect(book.diagnostics.filter(d => d.feature === 'pivot-count-mismatch').map(d => d.message).sort()).toEqual([
      'items in xl/pivotTables/pivotTable1.xml declares count 7 but has 1',
      'pivotFields in xl/pivotTables/pivotTable1.xml declares count 3 but has 1',
      'rowItems in xl/pivotTables/pivotTable1.xml declares count 0 but has 1',
      'sharedItems in xl/pivotCache/pivotCacheDefinition1.xml declares count 999999999 but has 1',
    ])
  })
})
