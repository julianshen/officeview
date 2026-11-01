// The GETPIVOTDATA seed with a second pivot on the Pivot sheet (T31 r4): a copy of the first at K3:S24 (same cache), plus
// Formulas!<ref> cells (formula + cached value) as Excel would save them.
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { SEED } from './xlsx-pivots-dependent-fixtures'
const SHIFT = 10
const letters = (n: number) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s }
const colOf = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
export async function twoPivotSeed(formulas: Array<{ ref: string; row: number; formula: string; cached: number }>) {
  const zip = await JSZip.loadAsync(readFileSync(SEED))
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet = sheet.replace(/(<row r="(\d+)"[^>]*>)([\s\S]*?)(<\/row>)/g, (_m, open: string, _r: string, body: string, close: string) => {
    const copies = [...body.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(\/>|>[\s\S]*?<\/c>)/g)].map(([, col, row, attrs, rest]) =>
      `<c r="${letters(colOf(col!) + SHIFT)}${row}"${attrs}${rest}`)
    return `${open.replace(/ spans="[^"]*"/, '')}${body}${copies.join('')}${close}`
  }).replace(/<dimension ref="[^"]*"\/>/, '')
  zip.file('xl/worksheets/sheet1.xml', sheet)
  const pt = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string')
  zip.file('xl/pivotTables/pivotTable2.xml', pt.replace(/name="[^"]*"/, 'name="SecondPivot"').replace(/xr:uid="\{[^}]*\}"/, 'xr:uid="{00000000-0000-0000-0000-000000000002}"').replace(/<location ref="A3:I24"/, '<location ref="K3:S24"'))
  zip.file('xl/pivotTables/_rels/pivotTable2.xml.rels', await zip.file('xl/pivotTables/_rels/pivotTable1.xml.rels')!.async('string'))
  const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdP2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/></Relationships>'))
  const ct = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('[Content_Types].xml', ct.replace('</Types>', '<Override PartName="/xl/pivotTables/pivotTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>'))
  let f = await zip.file('xl/worksheets/sheet3.xml')!.async('string')
  for (const c of formulas) {
    const cell = `<c r="${c.ref}"><f>${c.formula}</f><v>${c.cached}</v></c>`
    f = f.replace(new RegExp(`(<row r="${c.row}"[^>]*>)([\\s\\S]*?)(</row>)`), (_m, a, b, z) => `${a.replace(/ spans="[^"]*"/, '')}${b}${cell}${z}`)
  }
  zip.file('xl/worksheets/sheet3.xml', f.replace(/<dimension ref="[^"]*"\/>/, ''))
  return parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}
