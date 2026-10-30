import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { parsePivotRegion, sheetHasRegionStyles } from '../src/xlsx/pivots/stub'

const fixture = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`)))
const resident = async (bytes: Uint8Array) => parseXlsx(await OfficePackage.load(bytes))

async function paged(bytes: Uint8Array, docId: string) {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId, profile: 'D' })
  await session.done
  return { session, sheet: session.workbook.sheets[0], workbook: session.workbook }
}

describe('eager pivot/control detection stub', () => {
  it('decodes XML entities in pivot names and rejects reversed regions', () => {
    expect(parsePivotRegion('<pivotTableDefinition name="A&amp;B &#x1F600;"><location ref="$A$1:$C$3"/></pivotTableDefinition>', 'xl/pivotTables/p.xml').region?.name).toBe('A&B 😀')
    expect(parsePivotRegion('<pivotTableDefinition><location ref="C3:A1"/></pivotTableDefinition>', 'xl/pivotTables/p.xml').unreadable).toBe(true)
  })

  it('detects pivot locations on resident and paged paths', async () => {
    const bytes = fixture('pivot-page-field')
    const objectDoc = await resident(bytes)
    expect(objectDoc.sheets[0].pivotRegions).toEqual([{ partPath: 'xl/pivotTables/pivotTable1.xml', ref: 'A3:I10', name: '樞紐分析表2', styleName: 'PivotStyleLight16' }])

    const result = await paged(bytes, 'pivot-stub-resident-paged')
    try {
      expect(result.sheet.metadata.pivotRegions).toEqual(objectDoc.sheets[0].pivotRegions)
    } finally { await result.session.release() }
  })

  it('resolves worksheet slicer and timeline extension references on both paths', async () => {
    const bytes = fixture('slicer-light1')
    const objectSheet = (await resident(bytes)).sheets[0]
    expect(objectSheet.slicerRefs).toEqual([{ partPath: 'xl/slicers/slicer1.xml' }])
    expect(objectSheet.timelineRefs).toEqual([{ partPath: 'xl/timelines/timeline1.xml' }])

    const result = await paged(bytes, 'control-stub-resident-paged')
    try {
      expect(result.sheet.metadata.slicerRefs).toEqual(objectSheet.slicerRefs)
      expect(result.sheet.metadata.timelineRefs).toEqual(objectSheet.timelineRefs)
    } finally { await result.session.release() }
  })

  it('reads Excel table style info and exposes the per-sheet region-style predicate', async () => {
    const bytes = fixture('table-light-1')
    const doc = await resident(bytes)
    expect(doc.tables?.[0].styleInfo).toEqual({ name: 'TableStyleLight1', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false })
    expect(doc.sheets[0].tables?.[0].styleInfo?.name).toBe('TableStyleLight1')
    expect(sheetHasRegionStyles(doc.sheets[0])).toBe(true)

    const result = await paged(bytes, 'table-stub-style-info')
    try {
      expect(result.sheet.tables?.[0].styleInfo).toEqual(doc.tables?.[0].styleInfo)
    } finally { await result.session.release() }
  })

  it('leaves detection fields undefined for a workbook without pivots, controls or styled tables', async () => {
    const zip = new JSZip()
    zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Plain" sheetId="1" r:id="rId1"/></sheets></workbook>')
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const doc = await resident(bytes)
    expect(doc.sheets[0].pivotRegions).toBeUndefined()
    expect(doc.sheets[0].slicerRefs).toBeUndefined()
    expect(doc.sheets[0].timelineRefs).toBeUndefined()
    expect(sheetHasRegionStyles(doc.sheets[0])).toBe(false)
    const result = await paged(bytes, 'pivot-stub-plain-paged')
    try {
      expect(result.sheet.metadata.pivotRegions).toBeUndefined()
      expect(result.sheet.metadata.slicerRefs).toBeUndefined()
      expect(result.sheet.metadata.timelineRefs).toBeUndefined()
    } finally { await result.session.release() }
  })

  it('diagnoses an unreadable pivot location and omits its region on both paths', async () => {
    const zip = await JSZip.loadAsync(fixture('pivot-page-field'))
    const path = 'xl/pivotTables/pivotTable1.xml'
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace(/<location\b([^>]*?)ref="[^"]*"/, '<location$1ref="not-a-range"'))
    const bytes = await zip.generateAsync({ type: 'uint8array' })

    const objectSheet = (await resident(bytes)).sheets[0]
    expect(objectSheet.pivotRegions).toBeUndefined()
    expect(objectSheet.diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported', feature: 'pivot-location-unreadable' }))

    const result = await paged(bytes, 'pivot-stub-malformed-location')
    try {
      expect(result.sheet.metadata.pivotRegions).toBeUndefined()
      expect(result.workbook.diagnostics).toContainEqual(expect.objectContaining({ code: 'pivot-location-unreadable' }))
    } finally { await result.session.release() }
  })

  it('keeps table completeness when a table-less sheet has malformed relationships', async () => {
    const zip = new JSZip()
    zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Plain" sheetId="1" r:id="rId1"/></sheets></workbook>')
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship></Wrong>')
    const doc = await resident(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.tables).toEqual([])
    expect(doc.sheets[0].diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported', feature: 'pivot-control-refs-unreadable' }))
  })

  it('scans only the first 64 KiB of a large pivot part on both paths', async () => {
    const make = async (locationAfterPrefix: boolean) => {
      const zip = await JSZip.loadAsync(fixture('pivot-page-field'))
      const path = 'xl/pivotTables/pivotTable1.xml'
      const xml = await zip.file(path)!.async('string')
      const location = xml.match(/<location\b[^>]*\/>/)![0]
      const withoutLocation = xml.replace(location, '')
      const padding = `<!--${'x'.repeat(70 * 1024)}-->`
      const enlarged = locationAfterPrefix
        ? withoutLocation.replace('<pivotTableDefinition', `<pivotTableDefinition`).replace(/(<pivotTableDefinition\b[^>]*>)/, `$1${padding}` + location)
        : withoutLocation.replace(/(<pivotTableDefinition\b[^>]*>)/, `$1${location}${padding}`)
      zip.file(path, enlarged, { compression: 'DEFLATE' })
      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
    }
    const earlyBytes = await make(false)
    expect((await resident(earlyBytes)).sheets[0].pivotRegions?.[0]?.ref).toBe('A3:I10')
    const earlyPaged = await paged(earlyBytes, 'pivot-stub-large-early')
    try { expect(earlyPaged.sheet.metadata.pivotRegions?.[0]?.ref).toBe('A3:I10') } finally { await earlyPaged.session.release() }

    const lateBytes = await make(true)
    const lateDoc = await resident(lateBytes)
    expect(lateDoc.sheets[0].pivotRegions).toBeUndefined()
    expect(lateDoc.sheets[0].diagnostics).toContainEqual(expect.objectContaining({ feature: 'pivot-location-unreadable' }))
    const latePaged = await paged(lateBytes, 'pivot-stub-large-late')
    try {
      expect(latePaged.sheet.metadata.pivotRegions).toBeUndefined()
      expect(latePaged.workbook.diagnostics).toContainEqual(expect.objectContaining({ code: 'pivot-location-unreadable' }))
    } finally { await latePaged.session.release() }
  })
})
