// @vitest-environment node
/**
 * The XLSX sparkline reference workbooks (spec 8; scripts/sparkline-refs). They are generated, committed, and opened
 * in Microsoft Excel to capture reference images, so this checks they exist, carry the x14 sparkline extension, parse
 * as workbooks, and match what the generator produces now (a stale fixture would silently measure the wrong thing).
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXml } from '../src/core/xml'
import { parseXlsx } from '../src/xlsx/parse'
import { SPARKLINE_CASES, caseHostRanges } from '../scripts/sparkline-refs/cases'
import { buildCaseWorkbook } from '../scripts/sparkline-refs/generate'

const EXT_URI = 'ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"'
const fixturePath = (name: string) => fileURLToPath(new URL(`./fixtures/xlsx-sparklines/${name}.xlsx`, import.meta.url))
const load = (name: string) => JSZip.loadAsync(readFileSync(fixturePath(name)))
const sheetPaths = (zip: JSZip) => Object.keys(zip.files).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).sort()

describe('xlsx sparkline reference workbooks', () => {
  it('every case parses as a workbook with its sparkline XML', async () => {
    for (const c of SPARKLINE_CASES) {
      for (const name of [c.id, `${c.id}-bordered`]) {
        const zip = await load(name)
        const sheets = await Promise.all(sheetPaths(zip).map(p => zip.file(p)!.async('string')))
        expect(sheets.length, name).toBe(c.sheets.length)
        expect(sheets.some(x => x.includes(EXT_URI)), name).toBe(true)
        for (const x of sheets) expect(() => parseXml(x), name).not.toThrow()
        // The package is a workbook the viewer itself opens.
        const doc = await parseXlsx(await OfficePackage.load(readFileSync(fixturePath(name))))
        expect(doc.sheets.length, name).toBe(c.sheets.length)
      }
    }
  })

  it('covers every Stage 1 measure item', () => {
    const ids = new Set(SPARKLINE_CASES.map(c => c.id))
    const required = [
      'line-basic', 'line-defaults-omitted', 'line-markers', 'line-points-hl-fl', 'line-precedence', 'line-ties', 'line-weight-2_25',
      'line-scale-individual', 'line-scale-group', 'line-scale-custom', 'line-scale-custom-clip',
      'line-empty-gap', 'line-empty-zero', 'line-empty-span', 'line-text-bool-error',
      'line-hidden-rows', 'line-hidden-shown', 'line-filter-hidden',
      'line-axis', 'line-flat', 'line-single', 'line-all-negative', 'line-merged-host', 'line-later-sheet',
      'line-colours-theme-tint', 'line-colours-missing', 'line-tiny-cell',
      'layering',
    ]
    for (const id of required) expect(ids.has(id), id).toBe(true)
    expect(ids.size).toBe(SPARKLINE_CASES.length)
    for (const c of SPARKLINE_CASES) expect(c.stage).toBe(1)
  })

  it('the committed fixtures match the generator output', async () => {
    for (const c of SPARKLINE_CASES) {
      for (const bordered of [false, true]) {
        const name = bordered ? `${c.id}-bordered` : c.id
        const committed = await load(name)
        const fresh = await JSZip.loadAsync(await buildCaseWorkbook(c, bordered))
        const parts = (z: JSZip) => Object.keys(z.files).filter(p => !z.files[p].dir).sort()
        expect(parts(committed), name).toEqual(parts(fresh))
        for (const p of parts(fresh)) expect(await committed.file(p)!.async('string'), `${name} ${p}`).toBe(await fresh.file(p)!.async('string'))
      }
    }
  })

  it('every case has the same host geometry, bordered twins add only borders, and every host is addressable', async () => {
    for (const c of SPARKLINE_CASES) {
      const ranges = caseHostRanges(c)
      expect(ranges.length, c.id).toBeGreaterThan(0)
      const plain = await load(c.id), bordered = await load(`${c.id}-bordered`)
      for (const p of sheetPaths(plain)) {
        const a = await plain.file(p)!.async('string'), b = await bordered.file(p)!.async('string')
        // Identical but for the style indices of host cells (and the empty styled host cells that carry them).
        const bare = (x: string) => x.replace(/ s="\d+"/g, '').replace(/<c r="[A-Z]+\d+"\/>/g, '')
        expect(bare(b), `${c.id} ${p}`).toBe(bare(a))
      }
      const styles = await bordered.file('xl/styles.xml')!.async('string')
      expect(styles, c.id).toContain('<left style="thin">')
      expect(styles, c.id).not.toContain('indexedColors')
      const plainSheet = await plain.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(plainSheet, c.id).toContain('showGridLines="' + (c.id === 'layering' ? '1' : '0') + '"')
      // Host geometry: 40 characters wide (OOXML 40.7109375) and 60 pt high, except the tiny-cell case.
      const [width, height] = c.id === 'line-tiny-cell' ? ['2.7109375', '5'] : ['40.7109375', '60']
      c.sheets.forEach((sheet, i) => {
        const xml = i === 0 ? plainSheet : ''
        for (const g of sheet.groups) for (const sp of g.sparklines) {
          const m = /^([A-Z]+)(\d+)$/.exec(sp.sqref)!
          expect(xml, `${c.id} row of ${sp.sqref}`).toContain(`<row r="${m[2]}" ht="${height}" customHeight="1"`)
          const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
          expect(xml, `${c.id} column of ${sp.sqref}`).toContain(`<col min="${col}" max="${col}" width="${width}" customWidth="1"/>`)
        }
      })
    }
  })
})
