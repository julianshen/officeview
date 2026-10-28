import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { REF_CASES, type TableCase } from '../scripts/xlsx-refs/cases'
import { buildCaseWorkbook } from '../scripts/xlsx-refs/generate'

function elementShape(xml: string): string[] {
  return [...xml.matchAll(/<\/?([A-Za-z_][\w:.-]*)\b[^>]*>/g)].map(match => {
    const name = match[1]
    if (match[0].startsWith('</')) return `/${name}`
    return match[0].endsWith('/>') ? `${name}/` : name
  })
}

function subtree(xml: string, name: string): string {
  const start = xml.indexOf(`<${name}`)
  if (start < 0) throw new Error(`missing <${name}>`)
  const end = xml.indexOf(`</${name}>`, start)
  if (end < 0) throw new Error(`missing </${name}>`)
  return xml.slice(start, end + name.length + 3)
}

function namespaceDeclarations(xml: string): string[] {
  return [...xml.matchAll(/\bxmlns(?::([\w.-]+))?="([^"]+)"/g)].map(match => `${match[1] ?? ''}=${match[2]}`)
}

function tagAttributeValues(xml: string, tag: string, attribute: string): string[] {
  return [...xml.matchAll(new RegExp(`<${tag}\\b[^>]*>`, 'g'))].map(match => {
    const value = new RegExp(`\\b${attribute}="([^"]+)"`).exec(match[0])?.[1]
    return value ?? ''
  })
}

describe('pivot reference workbook generator', () => {
  it('includes the full Stage 1 style and control matrix', () => {
    expect(REF_CASES.filter(c => c.kind === 'table' && c.id.startsWith('table-light-')).length
      + REF_CASES.filter(c => c.kind === 'table' && c.id.startsWith('table-medium-')).length
      + REF_CASES.filter(c => c.kind === 'table' && c.id.startsWith('table-dark-')).length).toBe(60)
    expect(REF_CASES.filter(c => c.kind === 'pivot' && c.id.startsWith('pivot-') && c.id.match(/^pivot-(light|medium|dark)-\d+$/)).length).toBe(84)
    expect(REF_CASES.filter(c => c.kind === 'slicer' && !c.id.startsWith('slicer-two') && !c.id.startsWith('slicer-no') && !c.id.startsWith('slicer-scrolled')).length).toBe(14)
    expect(REF_CASES.filter(c => c.kind === 'timeline').length).toBe(56)
  })

  it('round-trips every XML case through the resident parser without diagnostics', async () => {
    for (const c of REF_CASES.filter(item => item.build === 'xml')) {
      const workbook = await buildCaseWorkbook(c)
      const parsed = await parseXlsx(await OfficePackage.load(workbook))
      expect(parsed.diagnostics ?? [], c.id).toEqual([])
    }
  })

  it('writes each table case declared tableStyleInfo into its table part', async () => {
    const tables = REF_CASES.filter(item => item.kind === 'table')
    expect(tables.length).toBeGreaterThan(60)
    for (const c of tables) {
      const workbook = await buildCaseWorkbook(c)
      const zip = await JSZip.loadAsync(workbook)
      const tableParts = Object.keys(zip.files).filter(path => /^xl\/tables\/table\d+\.xml$/.test(path))
      expect(tableParts, c.id).toHaveLength(1)
      const xml = await zip.file(tableParts[0])!.async('text')
      const spec = c.spec as TableCase
      expect(xml, c.id).toContain(`<tableStyleInfo name="${spec.styleName}"`)
      for (const [name, value] of Object.entries(spec.flags)) {
        expect(xml, `${c.id} ${name}`).toContain(`${name}="${value ? 1 : 0}"`)
      }
    }
  })

  it('provides deterministic source rows for Excel-built pivot cases', async () => {
    const { pivotSourceRows } = await import('../scripts/xlsx-refs/generate')
    expect(pivotSourceRows(17, 20)).toEqual(pivotSourceRows(17, 20))
    expect(pivotSourceRows(17, 20)).toHaveLength(20)
    expect(pivotSourceRows(17, 20)).not.toEqual(pivotSourceRows(18, 20))
    const sample = pivotSourceRows(17, 100)
    expect(new Set(sample.map(row => row.product)).size).toBe(8)
    expect(new Set(sample.filter(row => row.product === 'Apples').map(row => row.region)).size).toBeGreaterThan(1)
    expect(new Set(sample.filter(row => row.region === 'East').map(row => row.product)).size).toBeGreaterThan(1)
  })

  it('backs every XML slicer and timeline with saved pivot cache records', async () => {
    for (const c of REF_CASES.filter(item => item.kind === 'slicer' || item.kind === 'timeline')) {
      const zip = await JSZip.loadAsync(await buildCaseWorkbook(c))
      expect(zip.file('xl/pivotCache/pivotCacheDefinition1.xml'), c.id).toBeTruthy()
      expect(zip.file('xl/pivotCache/pivotCacheRecords1.xml'), c.id).toBeTruthy()
      expect(zip.file('xl/pivotTables/pivotTable1.xml'), c.id).toBeTruthy()
      const records = await zip.file('xl/pivotCache/pivotCacheRecords1.xml')!.async('text')
      expect(records, c.id).toContain('count="16"')
    }
  })

  it('writes conforming control extensions, cache names, dates and non-overlapping drawings', async () => {
    for (const c of REF_CASES.filter(item => item.kind === 'slicer' || item.kind === 'timeline')) {
      const zip = await JSZip.loadAsync(await buildCaseWorkbook(c))
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
      const workbook = await zip.file('xl/workbook.xml')!.async('text')
      const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('text')
      const pivot = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('text')
      expect(drawing, c.id).toContain('<mc:AlternateContent')
      expect(drawing, c.id).toContain('<mc:Choice')
      expect(drawing, c.id).toContain('Requires=')
      expect(drawing, c.id).toContain('<mc:Fallback>')
      expect(sheet, c.id).not.toContain('<pivotTableParts')
      expect(drawing, c.id).toContain('<a:off')
      expect(drawing, c.id).toContain('<a:ext')
      expect(drawing, c.id).toContain('<xdr:col>3</xdr:col>')
      expect(drawing, c.id).toContain('<xdr:col>6</xdr:col>')
      expect(pivot, c.id).toContain('ref="G1:H6"')
      expect(pivot, c.id).toContain('<pivotField axis="axisRow"')
      expect(pivot, c.id).toContain('<pivotField dataField="1"')
      expect(pivot, c.id).toContain('<items count="4">')
      expect(sheet, c.id).toContain('r="G1"')
      expect(sheet, c.id).toContain('r="H6"')
      expect(sheet, c.id).not.toContain('r="A1"')
      expect(pivot, c.id).not.toContain('<colFields count="0"')
      expect((pivot.match(/<i>/g) ?? []).length, c.id).toBe(Number(/rowItems count="(\d+)"/.exec(pivot)?.[1]))
      expect(workbook, c.id).toContain('name="NativeTimeline_Date">#N/A</definedName>')
      expect(workbook, c.id).toContain('name="Slicer_Product">#N/A</definedName>')
      expect(sheet, c.id).toContain('uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"')
      expect(sheet, c.id).toContain('uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}"')
      expect(workbook, c.id).toContain('uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}"')
      expect(workbook, c.id).toContain('uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"')
      const sheetRels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')
      const workbookRels = await zip.file('xl/_rels/workbook.xml.rels')!.async('text')
      expect(tagAttributeValues(sheetRels, 'Relationship', 'Type')).toEqual([
        'http://schemas.microsoft.com/office/2007/relationships/slicer',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
        'http://schemas.microsoft.com/office/2011/relationships/timeline',
      ])
      expect(workbookRels).toContain('http://schemas.microsoft.com/office/2011/relationships/timelineCache')
      expect(workbookRels).toContain('http://schemas.microsoft.com/office/2007/relationships/slicerCache')
    }
  })

  it('stores a date-formatted Date source field for timeline cases', async () => {
    for (const c of REF_CASES.filter(item => item.kind === 'timeline')) {
      const zip = await JSZip.loadAsync(await buildCaseWorkbook(c))
      const source = await zip.file('xl/worksheets/sheet2.xml')!.async('text')
      const styles = await zip.file('xl/styles.xml')!.async('text')
      const cache = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml')!.async('text')
      expect(source, c.id).toContain('r="C2" s="2"')
      expect(styles, c.id).toContain('formatCode="yyyy-mm-dd"')
      expect(cache, c.id).toContain('<sharedItems containsDate="1" containsNonDate="0" count="16">')
      expect(cache, c.id).toContain('<d v="')
      expect(cache, c.id).toContain('<x14:pivotCacheDefinition pivotCacheId="1"')
      const timelineCache = await zip.file('xl/timelineCaches/timelineCache1.xml')!.async('text')
      expect(timelineCache, c.id).toContain('sourceName="Date"')
      expect(timelineCache, c.id).toContain('<state ')
      expect(timelineCache, c.id).toContain('<bounds startDate=')
      expect(timelineCache, c.id).not.toContain('<timelineState')
      expect(timelineCache, c.id).not.toContain('<selection')
      expect(timelineCache, c.id).toContain('<bounds startDate="2024-01-01T00:00:00" endDate="2026-01-01T00:00:00"/>')
    }
  })

  it('builds byte-identical workbooks for the same case', async () => {
    const c = REF_CASES.find(item => item.id === 'slicer-light1')!
    expect(await buildCaseWorkbook(c)).toEqual(await buildCaseWorkbook(c))
  })

  it('matches the Excel-written control-part element structure and namespaces', async () => {
    const reference = await JSZip.loadAsync(readFileSync('tests/fixtures/xlsx-pivots/recipes/r1-slicer-timeline.xlsx'))
    const cases = [REF_CASES.find(item => item.id === 'slicer-light1')!, ...REF_CASES.filter(item => item.kind === 'timeline')]
    const paths = [
      '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
      'xl/worksheets/_rels/sheet1.xml.rels', 'xl/slicerCaches/slicerCache1.xml',
      'xl/timelineCaches/timelineCache1.xml', 'xl/slicers/slicer1.xml', 'xl/timelines/timeline1.xml',
      'xl/drawings/drawing1.xml',
    ]
    for (const c of cases) {
      const generated = await JSZip.loadAsync(await buildCaseWorkbook(c))
      for (const path of paths) {
        const expected = await reference.file(path)!.async('text')
        const actual = await generated.file(path)!.async('text')
        expect(elementShape(actual), `${c.id} ${path}`).toEqual(elementShape(expected))
        expect(namespaceDeclarations(actual), `${c.id} namespaces ${path}`).toEqual(namespaceDeclarations(expected))
        expect([...actual.matchAll(/mc:Ignorable="([^"]+)"/g)].map(match => match[1]), `${c.id} ignorable ${path}`)
          .toEqual([...expected.matchAll(/mc:Ignorable="([^"]+)"/g)].map(match => match[1]))
      }
      const expectedTypes = await reference.file('[Content_Types].xml')!.async('text')
      const actualTypes = await generated.file('[Content_Types].xml')!.async('text')
      for (const attribute of ['PartName', 'ContentType']) {
        expect(tagAttributeValues(actualTypes, 'Override', attribute), `${c.id} content type ${attribute}`)
          .toEqual(tagAttributeValues(expectedTypes, 'Override', attribute))
      }
      const expectedSheetRels = await reference.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')
      const actualSheetRels = await generated.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('text')
      const expectedWorkbookRels = await reference.file('xl/_rels/workbook.xml.rels')!.async('text')
      const actualWorkbookRels = await generated.file('xl/_rels/workbook.xml.rels')!.async('text')
      const expectedRootRels = await reference.file('_rels/.rels')!.async('text')
      const actualRootRels = await generated.file('_rels/.rels')!.async('text')
      for (const [expected, actual, label] of [
        [expectedSheetRels, actualSheetRels, 'sheet relationships'],
        [expectedWorkbookRels, actualWorkbookRels, 'workbook relationships'],
        [expectedRootRels, actualRootRels, 'package relationships'],
      ]) {
        for (const attribute of ['Type', 'Target']) {
          expect(tagAttributeValues(actual, 'Relationship', attribute), `${c.id} ${label} ${attribute}`)
            .toEqual(tagAttributeValues(expected, 'Relationship', attribute))
        }
      }
      const expectedDrawing = await reference.file('xl/drawings/drawing1.xml')!.async('text')
      const actualDrawing = await generated.file('xl/drawings/drawing1.xml')!.async('text')
      expect(tagAttributeValues(actualDrawing, 'mc:Choice', 'Requires'), `${c.id} Choice Requires`)
        .toEqual(tagAttributeValues(expectedDrawing, 'mc:Choice', 'Requires'))
      const expectedSlicerCache = await reference.file('xl/slicerCaches/slicerCache1.xml')!.async('text')
      const actualSlicerCache = await generated.file('xl/slicerCaches/slicerCache1.xml')!.async('text')
      expect(actualSlicerCache, `${c.id} selected Product items`).toContain('<items count="4"><i x="2" s="1"/><i x="1"/><i x="3" s="1"/><i x="0"/></items>')
      expect(actualSlicerCache, `${c.id} slicer cache structure`).toContain('<tabular pivotCacheId="1"><items count="4">')
      expect(elementShape(actualSlicerCache), `${c.id} reference slicer cache structure`).toEqual(elementShape(expectedSlicerCache))
      const expectedSheet = await reference.file('xl/worksheets/sheet1.xml')!.async('text')
      const actualSheet = await generated.file('xl/worksheets/sheet1.xml')!.async('text')
      expect(namespaceDeclarations(actualSheet), `${c.id} sheet namespaces`).toEqual(namespaceDeclarations(expectedSheet))
      expect(elementShape(subtree(actualSheet, 'extLst')), `${c.id} sheet ext`).toEqual(elementShape(subtree(expectedSheet, 'extLst')))
      const expectedBook = await reference.file('xl/workbook.xml')!.async('text')
      const actualBook = await generated.file('xl/workbook.xml')!.async('text')
      expect(elementShape(subtree(actualBook, 'extLst')), `${c.id} workbook ext`).toEqual(elementShape(subtree(expectedBook, 'extLst')))
      expect(elementShape(subtree(actualBook, 'definedNames')), `${c.id} defined names`).toEqual(elementShape(subtree(expectedBook, 'definedNames')))
    }
  })
})
