import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { REF_CASES, type TableCase } from '../scripts/xlsx-refs/cases'
import { buildCaseWorkbook } from '../scripts/xlsx-refs/generate'

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
})
