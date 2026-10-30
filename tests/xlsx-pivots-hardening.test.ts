// @vitest-environment node
/**
 * Final-review hardening (final-review.md C1, I5, I6): a hostile pivot
 * definition fails closed per sheet, the region factory and the button pass
 * never throw through paint, `<format>` count does not scale `styleAt`, and
 * resident control ids are de-duplicated and capped.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import JSZip from 'jszip'
import { afterEach, describe, expect, test } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { ensureTableStyleEngine, registerTableStyleEngine, regionStylesFor } from '../src/xlsx/tableStyleLoader'
import { regionStylesFromSource } from '../src/xlsx/tableStyles/regionLayer'
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import { paintPivotButtons } from '../src/xlsx/pivots/buttons'
import { pivotStyleRegions } from '../src/xlsx/pivots/regions'
import { parseThemeContext } from '../src/drawing/style'
import type { XlsxSheet } from '../src/xlsx/types'
import type { PivotTableDefinition } from '../src/xlsx/pivots/model'
import { controlIds, MAX_CONTROL_IDS } from '../src/xlsx/pivots/stub'
import { parseXml } from '../src/core/xml'

const fixture = (name: string) => readFileSync(fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}.xlsx`, import.meta.url)))
const PIVOT_PART = 'xl/pivotTables/pivotTable1.xml'

/** pivot-layout-compact with its pivot part rewritten by `edit`. */
async function withPivotPart(edit: (xml: string) => string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(fixture('pivot-layout-compact'))
  zip.file(PIVOT_PART, edit(await zip.file(PIVOT_PART)!.async('string')))
  return zip.generateAsync({ type: 'uint8array' })
}
const pivotSheet = (sheets: readonly XlsxSheet[]) => sheets.find(sheet => sheet.pivotRegions?.length)!

describe('C1: a hostile pivot definition fails closed', () => {
  const N = 150_000
  // The reviewer's case: one field repeated N times on the row axis, and a first row line N levels deep.
  const repeated = (xml: string) => xml
    .replace(/<rowFields count="2">[\s\S]*?<\/rowFields>/, `<rowFields count="${N}">${'<field x="0"/>'.repeat(N)}</rowFields>`)
    .replace('<rowItems count="17"><i><x/></i>', `<rowItems count="17"><i>${'<x/>'.repeat(N)}</i>`)
  // A grand line's placeholder list N long: the axis is two deep.
  const deepGrand = (xml: string) => xml.replace('<i t="grand"><x/></i></rowItems>', `<i t="grand">${'<x/>'.repeat(N)}</i></rowItems>`)

  for (const [name, edit] of [['a field repeated on the row axis', repeated], ['a grand line past the axis depth', deepGrand]] as const) {
    test(`${name}: invalid-index, no layout, and every sheet still paints`, async () => {
      const doc = await parseXlsx(await OfficePackage.load(await withPivotPart(edit)))
      const sheet = pivotSheet(doc.sheets)
      const def = sheet.pivotTables![0]!
      const pages = await getPaintables(doc)
      try {
        expect(pages).toHaveLength(doc.sheets.length)
        for (const page of pages) {
          const canvas = createCanvas(Math.min(page.spec.widthPx, 400), Math.min(page.spec.heightPx, 400))
          expect(() => page.paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)).not.toThrow()
        }
      } finally { pages.dispose() }
      expect(def.unsupported.some(entry => entry.startsWith('invalid-index:rowFields') || entry.startsWith('invalid-index:rowItems'))).toBe(true)
      expect(savedPivotLayout(def)).toBeUndefined()
      expect(sheet.regionStylesResolved).toBe(true)
      expect(sheet.diagnostics?.map(item => item.feature)).toContain('pivot-layout-unavailable')
    }, 60_000)
  }

  test('a real compact pivot still validates and lays out', async () => {
    const doc = await parseXlsx(await OfficePackage.load(fixture('pivot-layout-compact')))
    const def = pivotSheet(doc.sheets).pivotTables![0]!
    expect(def.unsupported.filter(entry => entry.startsWith('invalid-index:'))).toEqual([])
    expect(savedPivotLayout(def)).toBeDefined()
  })

  test('a deep, long tabular axis of distinct fields is not laid out (parent bands stay linear)', async () => {
    // D distinct tabular row fields, one data line through all of them, then M lines one level above the
    // leaf: every level's parent band would cover all M lines (D × M entries).
    const doc = await parseXlsx(await OfficePackage.load(fixture('pivot-layout-compact')))
    const real = pivotSheet(doc.sheets).pivotTables![0]!
    const D = 2_000, M = 20_000
    const def: PivotTableDefinition = {
      ...real,
      fields: Array.from({ length: D }, () => ({ ...real.fields[0]!, compact: false, outline: false })),
      rowFields: Array.from({ length: D }, (_, f) => f),
      rowItems: [{ type: 'data', r: 0, i: 0, x: Array(D).fill(0) }, ...Array.from({ length: M }, () => ({ type: 'data' as const, r: D - 1, i: 0, x: [0] }))],
      colFields: [],
      colItems: [{ type: 'data', r: 0, i: 0, x: [] }],
      location: { ...real.location, ref: `A3:B${4 + M}`, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
      unsupported: [],
    }
    const started = performance.now()
    const why: { tooDeep?: boolean } = {}
    expect(savedPivotLayout(def, why)).toBeUndefined()
    expect(performance.now() - started).toBeLessThan(2_000)
    expect(why.tooDeep).toBe(true)
    // A reachable, fail-closed styling limit (34 or more nested tabular row fields over long tails): it says so.
    const reports: string[] = []
    pivotStyleRegions({ name: 'Deep', rows: [], pivotTables: [def] } as unknown as XlsxSheet, { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }, parseThemeContext(), d => reports.push(`${d.feature}: ${d.message}`))
    expect(reports.some(line => line.startsWith('pivot-layout-unavailable:') && line.includes('styling limit'))).toBe(true)
    // The same axis with a short tail is an ordinary tabular pivot and lays out.
    const short = { ...def, rowItems: def.rowItems!.slice(0, 3), location: { ...def.location, ref: 'A3:B6' } }
    expect(savedPivotLayout(short)?.parentBands.get(2)?.length).toBe(D - 1)
  })
})

describe('C1: the region factory and the button pass fail closed', () => {
  afterEach(() => registerTableStyleEngine(regionStylesFromSource))

  test('a throwing factory leaves its sheet unstyled with one diagnostic; other sheets still paint', async () => {
    const doc = await parseXlsx(await OfficePackage.load(fixture('pivot-layout-compact')))
    const sheet = pivotSheet(doc.sheets)
    await ensureTableStyleEngine([])
    registerTableStyleEngine(() => { throw new Error('boom') })
    await expect(ensureTableStyleEngine(doc.sheets)).resolves.toBeUndefined()
    expect(sheet.regionStyles).toBeUndefined()
    expect(sheet.regionStylesResolved).toBe(true)
    expect(sheet.diagnostics?.filter(item => item.feature.startsWith('pivot-display-failed'))).toHaveLength(1)
    // The synchronous accessor fails closed too, once.
    sheet.regionStylesResolved = false
    expect(() => regionStylesFor(sheet)).not.toThrow()
    expect(sheet.regionStylesResolved).toBe(true)
    expect(sheet.diagnostics?.filter(item => item.feature === 'pivot-display-failed')).toHaveLength(1)
    const pages = await getPaintables(doc)
    try {
      for (const page of pages) {
        const canvas = createCanvas(Math.min(page.spec.widthPx, 400), Math.min(page.spec.heightPx, 400))
        expect(() => page.paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)).not.toThrow()
      }
    } finally { pages.dispose() }
  })

  test('a definition whose button placement throws paints no buttons, once, with one diagnostic', async () => {
    const doc = await parseXlsx(await OfficePackage.load(fixture('pivot-layout-compact')))
    const sheet = pivotSheet(doc.sheets)
    const def = sheet.pivotTables![0]!
    let reads = 0
    Object.defineProperty(def, 'location', { get() { reads++; throw new Error('boom') } })
    const ctx = createCanvas(200, 200).getContext('2d') as unknown as CanvasRenderingContext2D
    const rect = () => ({ x: 0, y: 0, width: 20, height: 20 })
    for (let tile = 0; tile < 3; tile++) expect(() => paintPivotButtons(ctx, sheet, rect, { r0: 0, c0: 0, r1: 40, c1: 20 }, 'screen')).not.toThrow()
    expect(reads).toBe(1)
    expect(sheet.diagnostics?.filter(item => item.feature === 'pivot-buttons-failed')).toHaveLength(1)
  })
})

describe('I5: styleAt cost does not grow with the <format> count', () => {
  /** pivot-formats-override with `count` formats of dxf 0, each over one (Values, Region, Product, Year) combination. */
  async function formatted(count: number, elsewhere = false) {
    const doc = await parseXlsx(await OfficePackage.load(fixture('pivot-formats-override')))
    const sheet = pivotSheet(doc.sheets)
    await ensureTableStyleEngine(doc.sheets)
    const def = sheet.pivotTables![0]!
    const base = def.formats[0]!
    const reference = (field: number, x: number) => ({ ...base.pivotArea.references[0]!, field, x: [x] })
    def.formats = Array.from({ length: count }, (_, k) => ({
      ...base,
      pivotArea: { ...base.pivotArea, references: elsewhere && k > 0
        // Formats over items this window does not show (a long pivot's other rows).
        ? [...base.pivotArea.references.slice(0, 1), reference(0, 4 + k)]
        : [reference(-2, k % 2), reference(0, (k >> 1) % 4), reference(1, (k >> 3) % 3), reference(2, Math.floor(k / 24) % 3)] },
    }))
    sheet.regionStyles = undefined
    sheet.regionStylesResolved = false
    const layer = regionStylesFor(sheet)!
    const own = new Map(sheet.rows.flatMap(row => row.cells.map(cell => [`${cell.row}:${cell.col}`, cell.style] as const)))
    return { layer, own }
  }
  const pass = ({ layer, own }: Awaited<ReturnType<typeof formatted>>, passes: number) => {
    const started = performance.now()
    for (let k = 0; k < passes; k++) for (let row = 2; row <= 21; row++) for (let col = 0; col <= 8; col++) layer.styleAt(row, col, own.get(`${row}:${col}`), 'screen')
    return performance.now() - started
  }

  test('1 vs 200 formats, 199 of them over cells elsewhere', async () => {
    const one = await formatted(1), many = await formatted(200, true)
    expect(many.layer.styleAt(6, 1, undefined, 'screen')?.fillColor).toBe('FFFFE699')
    pass(one, 20); pass(many, 20)
    const small = pass(one, 200), large = pass(many, 200)
    expect(large).toBeLessThan(2 * small + 50)
  })

  test('1 vs 200 formats over the pivot', async () => {
    const one = await formatted(1), many = await formatted(200)
    // The 200 formats cover all 48 data cells of the four-field grid; the first format still paints B7.
    expect(many.layer.styleAt(6, 1, undefined, 'screen')?.fillColor).toBe('FFFFE699')
    pass(one, 20); pass(many, 20) // warm both
    const small = pass(one, 200), large = pass(many, 200)
    expect(large).toBeLessThan(2 * small + 50)
  })
})

describe('I6: resident slicer and timeline ids are de-duplicated and capped', () => {
  test('20,000 copies of one slicer id read its part once', async () => {
    const zip = await JSZip.loadAsync(fixture('slicer-light1'))
    const path = 'xl/worksheets/sheet1.xml'
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace('<x14:slicer r:id="rId3"/>', '<x14:slicer r:id="rId3"/>'.repeat(20_000)))
    const input = await zip.generateAsync({ type: 'uint8array' })
    expect(controlIds(parseXml(await (await JSZip.loadAsync(input)).file(path)!.async('string'))).slicerIds).toEqual(['rId3'])
    const reference = await parseXlsx(await OfficePackage.load(fixture('slicer-light1')))
    const started = performance.now()
    const doc = await parseXlsx(await OfficePackage.load(input))
    expect(performance.now() - started).toBeLessThan(2_000)
    const frames = (sheet: XlsxSheet) => (sheet.drawings ?? []).filter(drawing => JSON.stringify(drawing).includes('"slicer"')).length
    expect(frames(doc.sheets[0]!)).toBe(frames(reference.sheets[0]!))
    expect(doc.sheets[0]!.slicerRefs).toHaveLength(1)
  })

  test('distinct ids are capped', () => {
    const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    const slicers = Array.from({ length: MAX_CONTROL_IDS + 10 }, (_, k) => `<x14:slicer r:id="rId${k}"/>`).join('')
    const root = parseXml(`<worksheet ${ns}><extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"><x14:slicerList>${slicers}</x14:slicerList></ext></extLst></worksheet>`)
    expect(controlIds(root).slicerIds).toHaveLength(MAX_CONTROL_IDS)
  })
})
