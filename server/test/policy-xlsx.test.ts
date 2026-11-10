import { expect, test } from 'bun:test'
import JSZip from 'jszip'
import { OfficePackage } from '../../src/core/zip'
import { projectPackage } from '../src/policy'

async function projected(parts: Record<string, string>) {
  const zip = new JSZip()
  for (const [p, xml] of Object.entries(parts)) zip.file(p, xml)
  return projectPackage(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), 'xlsx', 'visible-only')
}
const workbook = '<workbook xmlns:r="rel"><sheets><sheet name="Visible" sheetId="1" r:id="s1"/><sheet name="Secret" state="veryHidden" sheetId="2" r:id="s2"/><sheet name="Last" sheetId="3" r:id="s3"/></sheets><definedNames><definedName name="secret" localSheetId="1">SECRET</definedName><definedName name="last" localSheetId="2">Last!A1</definedName></definedNames></workbook>'
const rels = '<Relationships><Relationship Id="s1" Target="worksheets/sheet1.xml"/><Relationship Id="s2" Target="worksheets/sheet2.xml"/><Relationship Id="s3" Target="worksheets/sheet3.xml"/></Relationships>'

test('XLSX removes hidden sheets, renumbers local names, purges hidden strings but preserves collapsed geometry', async () => {
  const { pkg } = await projected({
    'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': '<worksheet><cols><col min="2" max="2" width="18" hidden="true"/></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2" hidden="1" ht="22"><c r="A2" t="s"><v>2</v></c></row></sheetData></worksheet>',
    'xl/worksheets/sheet2.xml': '<worksheet><sheetData><row><c t="s"><v>3</v></c></row></sheetData></worksheet>',
    'xl/worksheets/sheet3.xml': '<worksheet/>',
    'xl/sharedStrings.xml': '<sst count="4" uniqueCount="4"><si><t>VISIBLE</t></si><si><t>SECRET_COLUMN</t></si><si><t>SECRET_ROW</t></si><si><t>SECRET_SHEET</t></si></sst>',
  })
  expect(await pkg.text('xl/workbook.xml')).not.toContain('SECRET')
  expect(await pkg.text('xl/workbook.xml')).toContain('localSheetId="1"')
  expect(pkg.has('xl/worksheets/sheet2.xml')).toBe(false)
  const sheet = await pkg.text('xl/worksheets/sheet1.xml')
  expect(sheet).toContain('ht="22"'); expect(sheet).toContain('width="18"')
  expect(sheet).not.toContain('r="B1"'); expect(sheet).not.toContain('r="A2"')
  expect(await pkg.text('xl/sharedStrings.xml')).not.toContain('SECRET')
  expect(await pkg.text('xl/sharedStrings.xml')).toContain('VISIBLE')
})

test('XLSX removes entire hidden drawing anchors and comment/cache records parts', async () => {
  const { pkg } = await projected({
    'xl/workbook.xml': '<workbook/>',
    'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="xdr"><xdr:twoCellAnchor><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr hidden="1"/></xdr:nvGrpSpPr><xdr:sp><t>SECRET</t></xdr:sp></xdr:grpSp></xdr:twoCellAnchor><xdr:oneCellAnchor><xdr:sp><t>VISIBLE</t></xdr:sp></xdr:oneCellAnchor></xdr:wsDr>',
    'xl/comments1.xml': '<comments>SECRET</comments>',
    'xl/threadedComments/threadedComment1.xml': '<ThreadedComments>SECRET</ThreadedComments>',
    'xl/pivotCache/pivotCacheRecords1.xml': '<pivotCacheRecords>SECRET</pivotCacheRecords>',
  })
  expect(await pkg.text('xl/drawings/drawing1.xml')).not.toContain('SECRET')
  expect(await pkg.text('xl/drawings/drawing1.xml')).toContain('VISIBLE')
  expect(pkg.has('xl/comments1.xml')).toBe(false)
  expect(pkg.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
  expect(pkg.has('xl/pivotCache/pivotCacheRecords1.xml')).toBe(false)
})

test('XLSX hidden-source and unresolvable caches empty items; visible-source cache stays intact', async () => {
  const cache = (sheet: string, ref: string) => `<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="${sheet}" ref="${ref}"/></cacheSource><cacheFields><cacheField name="Label"><sharedItems count="1"><s v="SECRET"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`
  const { pkg } = await projected({
    'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="2" hidden="1"/></sheetData></worksheet>',
    'xl/worksheets/sheet2.xml': '<worksheet/>', 'xl/worksheets/sheet3.xml': '<worksheet/>',
    'xl/pivotCache/pivotCacheDefinition1.xml': cache('Visible', 'A1:B3'),
    'xl/pivotCache/pivotCacheDefinition2.xml': cache('Visible', 'A4:B8'),
    'xl/pivotCache/pivotCacheDefinition3.xml': cache('Secret', 'A1:B3'),
    'xl/slicerCaches/slicerCache1.xml': '<slicerCacheDefinition><data><tabular pivotCacheId="unknown"><items count="1"><i n="SECRET"/></items></tabular></data></slicerCacheDefinition>',
  })
  expect(await pkg.text('xl/pivotCache/pivotCacheDefinition1.xml')).not.toContain('SECRET')
  expect(await pkg.text('xl/pivotCache/pivotCacheDefinition2.xml')).toContain('SECRET')
  expect(await pkg.text('xl/pivotCache/pivotCacheDefinition3.xml')).not.toContain('SECRET')
  expect(await pkg.text('xl/slicerCaches/slicerCache1.xml')).not.toContain('SECRET')
})

test('XLSX controls resolve x14 cache ids and purge hidden source timeline dates', async () => {
  const { pkg } = await projected({
    'xl/workbook.xml': workbook, 'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': '<worksheet/>', 'xl/worksheets/sheet2.xml': '<worksheet/>', 'xl/worksheets/sheet3.xml': '<worksheet/>',
    'xl/pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition xmlns:x14="x14"><cacheSource type="worksheet"><worksheetSource sheet="Visible" ref="A1:B2"/></cacheSource><extLst><ext><x14:pivotCacheDefinition pivotCacheId="42"/></ext></extLst></pivotCacheDefinition>',
    'xl/slicerCaches/slicerCache1.xml': '<slicerCacheDefinition><data><tabular pivotCacheId="42"><items count="1"><i n="VISIBLE"/></items></tabular></data></slicerCacheDefinition>',
    'xl/timelineCaches/timelineCache1.xml': '<timelineCacheDefinition><state pivotCacheId="missing"><bounds startDate="1999-01-01" endDate="SECRET_DATE"/><selection startDate="SECRET_DATE"/></state></timelineCacheDefinition>',
  })
  expect(await pkg.text('xl/slicerCaches/slicerCache1.xml')).toContain('VISIBLE')
  expect(await pkg.text('xl/timelineCaches/timelineCache1.xml')).not.toContain('SECRET_DATE')
})
