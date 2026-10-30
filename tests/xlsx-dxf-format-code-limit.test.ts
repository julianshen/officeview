// @vitest-environment node
/**
 * PR #55 round 5 P2: a table-style or pivot-format dxf number format longer than
 * MAX_FORMAT_CODE_UNITS shows as General, as a cell xf's does, on the resident and
 * paged paths alike; a code within the limit still applies.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { MAX_FORMAT_CODE_UNITS, parseXlsx } from '../src/xlsx/parse'
import { ensureTableStyleEngine } from '../src/xlsx/tableStyleLoader'
import { openImportSession } from '../src/xlsx/storage/pagedImport'

const fixture = (name: string) => readFileSync(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/${name}.xlsx`))
const LONG = `0.00${'"x"'.repeat(Math.ceil(MAX_FORMAT_CODE_UNITS / 3))}`
const SHORT = '0.000'

/** table-medium-2 restyled by a custom table style whose wholeTable dxf sets `code`. */
async function tableWithDxfFormat(code: string): Promise<Uint8Array<ArrayBuffer>> {
  const archive = await JSZip.loadAsync(fixture('table-medium-2'))
  const edit = async (path: string, change: (xml: string) => string) => {
    const xml = await archive.file(path)!.async('string'), next = change(xml)
    expect(next).not.toBe(xml)
    archive.file(path, next)
  }
  await edit('xl/styles.xml', xml => xml
    .replace('<dxfs count="0"/>', `<dxfs count="1"><dxf><numFmt numFmtId="170" formatCode="${code.replace(/"/g, '&quot;')}"/></dxf></dxfs>`)
    .replace(/<tableStyles count="0"([^/]*)\/>/, '<tableStyles count="1"$1><tableStyle name="OVFormat" pivot="0" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>'))
  await edit('xl/tables/table1.xml', xml => xml.replace('name="TableStyleMedium2"', 'name="OVFormat"'))
  return new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
}

/** pivot-formats-override with its `formats` dxf also setting `code`. */
async function pivotWithDxfFormat(code: string): Promise<Uint8Array<ArrayBuffer>> {
  const archive = await JSZip.loadAsync(fixture('pivot-formats-override'))
  const styles = await archive.file('xl/styles.xml')!.async('string')
  const edited = styles.replace('<dxf><fill>', `<dxf><numFmt numFmtId="170" formatCode="${code.replace(/"/g, '&quot;')}"/><fill>`)
  expect(edited).not.toBe(styles)
  archive.file('xl/styles.xml', edited)
  return new Uint8Array(await archive.generateAsync({ type: 'uint8array' }))
}

async function resolvedFormats(input: Uint8Array): Promise<Array<{ numFmtId: number; formatCode?: string }>> {
  const doc = await parseXlsx(await OfficePackage.load(input))
  await ensureTableStyleEngine(doc.sheets)
  const sheet = doc.sheets.find(item => item.regionStyles)!
  const out: Array<{ numFmtId: number; formatCode?: string }> = []
  for (const row of sheet.rows) for (const cell of row.cells) {
    if (typeof cell.value !== 'number') continue
    const style = sheet.regionStyles!.styleAt(cell.row, cell.col, cell.style, 'screen')
    if (style && style !== cell.style) out.push({ numFmtId: style.numFmtId, ...(style.formatCode !== undefined ? { formatCode: style.formatCode } : {}) })
  }
  return out
}

describe('dxf number formats obey the authored format-code limit', () => {
  it('a table style dxf: an overlong code is General, a short one applies (resident)', async () => {
    expect(LONG.length).toBeGreaterThan(MAX_FORMAT_CODE_UNITS)
    const short = await resolvedFormats(await tableWithDxfFormat(SHORT))
    expect(short.length).toBeGreaterThan(0)
    expect(short.every(style => style.formatCode === SHORT)).toBe(true)
    const long = await resolvedFormats(await tableWithDxfFormat(LONG))
    expect(long.length).toBe(short.length)
    // 79ff888f: the whole code reached every General number in the table.
    expect(long.every(style => style.numFmtId === 0 && style.formatCode === undefined)).toBe(true)
  })

  it('a pivot format dxf: an overlong code is General (resident)', async () => {
    const short = await resolvedFormats(await pivotWithDxfFormat(SHORT))
    expect(short.some(style => style.formatCode === SHORT)).toBe(true)
    const long = await resolvedFormats(await pivotWithDxfFormat(LONG))
    expect(long.some(style => style.formatCode === LONG)).toBe(false)
  })

  it('the paged worker parses the same dxf as General', async () => {
    const session = await openImportSession({ kind: 'blob', blob: new Blob([await tableWithDxfFormat(LONG)]) }, { docId: 'dxf-format-limit', profile: 'D' })
    try {
      await session.done
      expect(session.workbook.tableStyleExtras?.dxfs[0]?.numFmt).toEqual({ id: 0 })
    } finally { await session.release() }
  })
})
