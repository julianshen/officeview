import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { SEED } from './xlsx-pivots-dependent-r2-helpers'
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
/** The seed with extra Formulas!K<row> cells (formula + cached value) and workbook defined names, as Excel would save them. */
export async function patchedSeedBytes(cells: Array<{ row: number; formula: string; cached: number | string; col?: string }>, names: Array<[string, string]> = [], pivotCells: Array<{ row: number; formula: string; cached: number | string }> = []): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(readFileSync(SEED))
  let wb = await zip.file('xl/workbook.xml')!.async('string')
  if (names.length) wb = wb.replace('</sheets>', `</sheets><definedNames>${names.map(([n, s]) => `<definedName name="${n}">${esc(s)}</definedName>`).join('')}</definedNames>`)
  zip.file('xl/workbook.xml', wb)
  const patch = async (path: string, list: typeof cells) => {
    let xml = await zip.file(path)!.async('string')
    for (const c of list) {
      const v = typeof c.cached === 'number' ? `<v>${c.cached}</v>` : `<v>${esc(c.cached)}</v>`
      const t = typeof c.cached === 'string' ? (c.cached.startsWith('#') ? ' t="e"' : ' t="str"') : ''
      const cell = `<c r="${(c as { col?: string }).col ?? 'K'}${c.row}"${t}><f>${esc(c.formula)}</f>${v}</c>`
      const re = new RegExp(`(<row r="${c.row}"[^>]*>)([\\s\\S]*?)(</row>)`)
      if (re.test(xml)) xml = xml.replace(re, (_m, a, b, z) => `${a}${b}${cell}${z}`)
      else xml = xml.replace('</sheetData>', `<row r="${c.row}">${cell}</row></sheetData>`)
    }
    xml = xml.replace(/<dimension ref="[^"]*"\/>/, '')
    zip.file(path, xml)
  }
  await patch('xl/worksheets/sheet3.xml', cells)
  if (pivotCells.length) await patch('xl/worksheets/sheet1.xml', pivotCells)
  return zip.generateAsync({ type: 'uint8array' })
}
export async function patchedSeed(...args: Parameters<typeof patchedSeedBytes>) {
  return parseXlsx(await OfficePackage.load(await patchedSeedBytes(...args)))
}
