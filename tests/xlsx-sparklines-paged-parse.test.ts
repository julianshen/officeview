import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { MAX_SPARKLINE_GROUPS_PER_SHEET, SPARKLINE_EXTENSION_URI, X14_SPARKLINE_NS, XM_SPARKLINE_NS } from '../src/xlsx/sparklines/parseGroups'
import { SPARKLINE_CASES } from '../scripts/sparkline-refs/cases'
import { xlsxEntries } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'

const fixtureBytes = (name: string) => new Uint8Array(readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-sparklines/${name}.xlsx`)))

async function readBoth(bytes: Uint8Array, docId: string) {
  const resident = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer]) }, { docId, profile: 'D' })
  try {
    const outcome = await session.done
    expect(outcome.kind).toBe('complete')
    return { resident, paged: session.workbook }
  } finally {
    await session.release()
  }
}

describe('paged XLSX sparkline parsing', () => {
  for (const sparkCase of SPARKLINE_CASES) for (const suffix of ['', '-bordered']) {
    it(`${sparkCase.id}${suffix} matches resident sparkline groups and diagnostics`, async () => {
      const { resident, paged } = await readBoth(fixtureBytes(`${sparkCase.id}${suffix}`), `spark-paged-${sparkCase.id}${suffix}`)
      expect(paged.sheets.map(sheet => sheet.metadata.sparklineGroups ?? [])).toEqual(resident.sheets.map(sheet => sheet.sparklineGroups ?? []))
      expect(paged.sheets.map(sheet => (sheet.metadata.sparklineDiagnostics ?? []).map(({ kind, feature, message }) => ({ kind, feature, message }))))
        .toEqual(resident.sheets.map(sheet => (sheet.diagnostics ?? []).filter(item => item.feature.startsWith('sparkline-')).map(({ kind, feature, message }) => ({ kind, feature, message }))))
      expect(paged.diagnostics.filter(item => item.code.startsWith('sparkline-')).map(({ code, message }) => ({ feature: code, message })))
        .toEqual(resident.sheets.flatMap(sheet => (sheet.diagnostics ?? []).filter(item => item.feature.startsWith('sparkline-')).map(({ feature, message }) => ({ feature, message }))))
    })
  }

  it('tracks alternate prefixes and ignores the wrong extension URI and foreign sparkline namespace', async () => {
    const good = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:a14="${X14_SPARKLINE_NS}" xmlns:xmm="${XM_SPARKLINE_NS}" xmlns:bad="urn:wrong"><extLst><ext uri="${SPARKLINE_EXTENSION_URI}"><a14:sparklineGroups><a14:sparklineGroup type="column"><a14:sparklines><a14:sparkline><xmm:f>A1:A2</xmm:f><xmm:sqref>J1</xmm:sqref></a14:sparkline></a14:sparklines></a14:sparklineGroup><bad:sparklineGroup type="stacked"/></a14:sparklineGroups></ext></extLst></worksheet>`
    const wrong = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:a14="${X14_SPARKLINE_NS}"><extLst><ext uri="urn:wrong"><a14:sparklineGroups><a14:sparklineGroup type="column"/></a14:sparklineGroups></ext><ext uri="${SPARKLINE_EXTENSION_URI}"><sparklineGroups xmlns="urn:wrong"><sparklineGroup type="stacked"/></sparklineGroups></ext></extLst></worksheet>`
    for (const [xml, expected] of [[good, 'column'], [wrong, undefined]] as const) {
      const bytes = buildZip(xlsxEntries({ afterSheetData: '' }).map(entry => entry.name === 'xl/worksheets/sheet1.xml' ? { ...entry, data: xml } : entry)).bytes
      const { resident, paged } = await readBoth(bytes, `spark-ns-${expected ?? 'ignored'}`)
      const residentGroups = resident.sheets[0].sparklineGroups ?? []
      const pagedGroups = paged.sheets[0].metadata.sparklineGroups ?? []
      expect(pagedGroups).toEqual(residentGroups)
      expect(pagedGroups[0]?.type).toBe(expected)
    }
  })

  it('bounds streamed groups before building and reports one group-limit diagnostic without failing import', async () => {
    const groups = Array.from({ length: MAX_SPARKLINE_GROUPS_PER_SHEET + 1 }, () => `<x14:sparklineGroup/>`).join('')
    const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="${X14_SPARKLINE_NS}"><extLst><ext uri="${SPARKLINE_EXTENSION_URI}"><x14:sparklineGroups>${groups}</x14:sparklineGroups></ext></extLst></worksheet>`
    const bytes = buildZip(xlsxEntries({ afterSheetData: '' }).map(entry => entry.name === 'xl/worksheets/sheet1.xml' ? { ...entry, data: xml } : entry)).bytes
    const { resident, paged } = await readBoth(bytes, 'spark-over-cap')
    const diagnostics = paged.sheets[0].metadata.sparklineDiagnostics ?? []
    const residentDiagnostics = (resident.sheets[0].diagnostics ?? []).filter(item => item.feature.startsWith('sparkline-')).map(({ kind, feature, message }) => ({ kind, feature, message }))
    expect(paged.sheets[0].metadata.sparklineGroups).toHaveLength(MAX_SPARKLINE_GROUPS_PER_SHEET)
    expect(diagnostics.filter(item => item.feature === 'sparkline-group-limit')).toHaveLength(1)
    expect(diagnostics.map(({ kind, feature, message }) => ({ kind, feature, message }))).toEqual(residentDiagnostics)
    expect(paged.diagnostics.filter(item => item.code === 'sparkline-group-limit')).toHaveLength(1)
    expect(paged.sheets[0].metadata.sparklineGroups).toEqual(resident.sheets[0].sparklineGroups)
  })
})
