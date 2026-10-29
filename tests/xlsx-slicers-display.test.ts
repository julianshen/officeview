// @vitest-environment node
/**
 * Slicer and timeline display (spec §3, §7.10, §10; Task 12). Every pixel
 * expectation cites the Excel 16.114 reference it is read from
 * (tests/fixtures/xlsx-pivots/refs/<case>.png, 1 px per pt, screen
 * appearance) and the measurement in docs/validation/xlsx-pivots.md
 * ("Slicer geometry", "Timeline geometry").
 *
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { createCanvas, loadImage } from 'canvas'
import { describe, expect, test, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { getPaintables } from '../src/render/paint'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { parseThemeContext } from '../src/drawing/style'
import { XLSX_CONTROL_REQUIREMENTS, XLSX_REQUIREMENTS, supportedChoiceRequirements, xlsxRequirements } from '../src/drawing/coverage'
import { getChildren, parseXml } from '../src/core/xml'
import { REF_CASES, type SlicerCase, type TimelineCase } from '../scripts/xlsx-refs/cases'
import { slicerItemState, slicerLayout, timelineLayout, TIMELINE_PERIOD_PT } from '../src/xlsx/pivots/slicers/paint'
import { slicerAppearance, timelineAppearance } from '../src/xlsx/pivots/slicers/style'
import type { SlicerModel, TimelineModel, XlsxControlContent } from '../src/xlsx/pivots/slicers/model'
import type { StylesExtras } from '../src/xlsx/tableStyles/model'
import type { XlsxDocument, XlsxDrawing, XlsxSheet } from '../src/xlsx/types'

const PX = 96 / 72
const fixturePath = (name: string) => fileURLToPath(new URL(`./fixtures/xlsx-pivots/${name}.xlsx`, import.meta.url))
const refPath = (name: string) => fileURLToPath(new URL(`./fixtures/xlsx-pivots/refs/${name}.png`, import.meta.url))
const fixture = (name: string) => readFileSync(fixturePath(name))
const r1 = () => readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/recipes/r1-slicer-timeline.xlsx', import.meta.url)))

/** A case's bytes with edits applied (path → new text, or null to delete the part). */
async function edited(bytes: Uint8Array, edits: Record<string, string | null | ((text: string) => string)>): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  for (const [path, edit] of Object.entries(edits)) {
    if (edit === null) zip.remove(path)
    else zip.file(path, typeof edit === 'function' ? edit(await zip.file(path)!.async('string')) : edit)
  }
  return zip.generateAsync({ type: 'uint8array' })
}
async function parse(bytes: Uint8Array): Promise<XlsxDocument> { return parseXlsx(await OfficePackage.load(bytes)) }
function controlOf<K extends XlsxControlContent['kind']>(drawing: XlsxDrawing | undefined, kind: K): Extract<XlsxControlContent, { kind: K }> | undefined {
  const content = drawing?.content
  return content?.kind === kind ? content as Extract<XlsxControlContent, { kind: K }> : undefined
}
function controls(sheet: XlsxSheet): { slicer?: XlsxDrawing; timeline?: XlsxDrawing } {
  const drawings = sheet.drawings ?? []
  return { slicer: drawings.find(d => d.content?.kind === 'slicer'), timeline: drawings.find(d => d.content?.kind === 'timeline') }
}
const slicerModel = (drawing: XlsxDrawing | undefined): SlicerModel => controlOf(drawing, 'slicer')!.slicer
const timelineModel = (drawing: XlsxDrawing | undefined): TimelineModel => controlOf(drawing, 'timeline')!.timeline
const boxOf = (drawing: XlsxDrawing) => ({ x: drawing.xEmu / 9525, y: drawing.yEmu / 9525, width: drawing.widthEmu / 9525, height: drawing.heightEmu / 9525 })

type Pixels = { width: number; height: number; data: Uint8ClampedArray | Uint8Array }
async function refPixels(name: string): Promise<Pixels> {
  const image = await loadImage(refPath(name))
  const canvas = createCanvas(image.width, image.height), ctx = canvas.getContext('2d')
  ctx.drawImage(image, 0, 0)
  return { width: image.width, height: image.height, data: ctx.getImageData(0, 0, image.width, image.height).data }
}
async function renderPixels(doc: XlsxDocument, sheet: XlsxSheet): Promise<Pixels> {
  const pages = await getPaintables(doc)
  try {
    const page = pages[doc.sheets.indexOf(sheet)]!
    const canvas = createCanvas(page.spec.widthPx, page.spec.heightPx), ctx = canvas.getContext('2d')
    page.paint(ctx as unknown as CanvasRenderingContext2D)
    return { width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data }
  } finally { pages.dispose() }
}
const rgb = (p: Pixels, x: number, y: number): [number, number, number] => {
  const i = (Math.floor(y) * p.width + Math.floor(x)) * 4
  return [p.data[i]!, p.data[i + 1]!, p.data[i + 2]!]
}
const hex = (c: readonly number[]) => c.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase()
const parseHex = (h: string): [number, number, number] => [0, 2, 4].map(i => parseInt(h.replace(/^#/, '').slice(i, i + 2), 16)) as [number, number, number]
const within = (a: readonly number[], b: readonly number[], tolerance: number) => a.every((v, i) => Math.abs(v - b[i]!) <= tolerance)
/** The pixel of a region farthest from its background: the text colour where a stem covers a whole pixel. */
function ink(p: Pixels, x0: number, y0: number, x1: number, y1: number, background: readonly number[]): [number, number, number] {
  let best: [number, number, number] = [background[0]!, background[1]!, background[2]!], bestDistance = -1
  for (let y = Math.floor(y0); y <= Math.floor(y1); y++) for (let x = Math.floor(x0); x <= Math.floor(x1); x++) {
    const c = rgb(p, x, y), d = Math.abs(c[0] - background[0]!) + Math.abs(c[1] - background[1]!) + Math.abs(c[2] - background[2]!)
    if (d > bestDistance) { bestDistance = d; best = c }
  }
  return best
}
function expectNear(label: string, actual: readonly number[], expected: readonly number[], tolerance = 3) {
  expect(within(actual, expected, tolerance), `${label}: rendered ${hex(actual)} vs reference ${hex(expected)} (±${tolerance})`).toBe(true)
}

const slicerCases = REF_CASES.filter(c => c.kind === 'slicer')
const timelineCases = REF_CASES.filter(c => c.kind === 'timeline')
const LEVELS = ['years', 'quarters', 'months', 'days'] as const

describe('every slicer and timeline reference case parses into a control drawing', () => {
  test.each(slicerCases.map(c => [c.id, c.spec as SlicerCase] as const))('%s', async (id, spec) => {
    const doc = await parse(fixture(id))
    const sheet = doc.sheets.find(s => s.name === 'Report')!
    const { slicer, timeline } = controls(sheet)
    const model = slicerModel(slicer)
    expect(model).toMatchObject({ name: 'Product', caption: 'Product', showCaption: true, columnCount: spec.columnCount, rowHeightEmu: 209550, startItem: spec.startItem })
    expect(model.style).toBe(spec.styleName)
    expect(model.items).toHaveLength(spec.itemCount)
    expect(model.drives).toEqual([{ pivotTable: 'PivotTable1' }])
    // refs/slicer-light1.png: the slicer cache's order (x = 2, 1, 3, 0) is the display order; Apples and Cherries selected.
    expect(model.items.slice(0, 4)).toEqual([
      { label: 'Apples', selected: true, noData: false }, { label: 'Dates', selected: false, noData: false },
      { label: 'Cherries', selected: true, noData: false }, { label: 'Bananas', selected: false, noData: false },
    ])
    if (spec.noData) expect(model.items[4]).toEqual({ label: 'Elderberries', selected: false, noData: true })
    expect(slicer!.source.representation).toBe('choice')
    // Every slicer case carries the default timeline (no style attribute: TimeSlicerStyleLight1, months).
    expect(timelineModel(timeline)).toMatchObject({ name: 'Date', caption: 'Date', style: 'TimeSlicerStyleLight1', level: 'months' })
    expect(sheet.diagnostics?.filter(d => /^(slicer|timeline)-/.test(d.feature)) ?? []).toEqual([])
  })

  test.each(timelineCases.map(c => [c.id, c.spec as TimelineCase] as const))('%s', async (id, spec) => {
    const doc = await parse(fixture(id))
    const sheet = doc.sheets.find(s => s.name === 'Report')!
    const model = timelineModel(controls(sheet).timeline)
    expect(model).toMatchObject({ name: 'Date', caption: 'Date', style: spec.styleName, level: spec.level })
    // timelineCache state bounds 2024-01-01 … 2025-01-01 (end exclusive), no selection: All Periods.
    expect(model.range).toEqual({ start: Date.UTC(2024, 0, 1), end: Date.UTC(2025, 0, 1) })
    expect(model.selection).toBeUndefined()
    expect(LEVELS).toContain(model.level)
  })

  test('recipe R1 (Excel-authored): the slicer cache joins the pivot cache by its x14 pivotCacheId, not the workbook cacheId', async () => {
    // r1: workbook <pivotCache cacheId="26">, x14:pivotCacheDefinition pivotCacheId="200334102" = tabular@pivotCacheId.
    const doc = await parse(r1())
    const sheet = doc.sheets.find(s => controls(s).slicer)!
    const model = slicerModel(controls(sheet).slicer)
    expect(model.items).toEqual([
      { label: 'Apples', selected: true, noData: false }, { label: 'Bananas', selected: false, noData: false },
      { label: 'Cherries', selected: true, noData: false }, { label: 'Dates', selected: false, noData: false },
    ])
    expect(model.style).toBe('SlicerStyleLight1')
    const timeline = timelineModel(controls(sheet).timeline)
    expect(timeline).toMatchObject({ style: 'TimeSlicerStyleLight1', level: 'months', range: { start: Date.UTC(2024, 0, 1), end: Date.UTC(2026, 0, 1) } })
  })
})

describe('item states map to style elements (Rulings A, Q, R)', () => {
  const theme = async () => parseThemeContext(await (await JSZip.loadAsync(r1())).file('xl/theme/theme1.xml')!.async('string'))
  const noExtras: StylesExtras = { dxfs: [], tableStyles: new Map(), slicerStyles: new Map() }

  test('selected/unselected × with/without data → the four SlicerStyleElement states (hover is never painted)', () => {
    expect(slicerItemState({ selected: true, noData: false })).toBe('selectedItemWithData')
    expect(slicerItemState({ selected: false, noData: false })).toBe('unselectedItemWithData')
    expect(slicerItemState({ selected: true, noData: true })).toBe('selectedItemWithNoData')
    expect(slicerItemState({ selected: false, noData: true })).toBe('unselectedItemWithNoData')
  })

  test('a built-in slicer style is two lookups: frame from builtinTableStyle, item states from builtinSlicerStyle', async () => {
    const t = await theme()
    const light = slicerAppearance('SlicerStyleLight1', noExtras, t).appearance
    // refs/slicer-light1.png: frame and separator #4F81BD, selected #B8CCE4 / #999999, unselected white / #CCCCCC.
    expect(light.border).toBe('#4F81BD')
    expect(light.header.separator).toBe('#4F81BD')
    // Theme tints resolve within a few RGB levels of Excel's (accent1 tint 0.6 → #B9CDE5 for the captured #B8CCE4).
    expectNear('Light1 selected fill', parseHex(light.items.selectedItemWithData.fill), parseHex('#B8CCE4'))
    expect(light.items.selectedItemWithData).toMatchObject({ border: '#999999', text: '#000000', bold: false })
    expect(light.items.unselectedItemWithData).toMatchObject({ fill: '#FFFFFF', border: '#CCCCCC' })
    // refs/slicer-no-data.png: no-data items white, #E0E0E0 border, #828282 text.
    expect(light.items.unselectedItemWithNoData).toMatchObject({ fill: '#FFFFFF', border: '#E0E0E0', text: '#828282' })
    // refs/slicer-dark1.png: selected accent fill with bold white text, unselected #C0C0C0.
    const dark = slicerAppearance('SlicerStyleDark1', noExtras, t).appearance
    expect(dark.items.selectedItemWithData).toMatchObject({ fill: '#4F81BD', text: '#FFFFFF', bold: true })
    expect(dark.items.unselectedItemWithData).toMatchObject({ fill: '#C0C0C0', border: '#C0C0C0' })
  })

  test('a custom slicer style is the same two lookups: its same-named tableStyle frame, its x14 slicerStyle items', async () => {
    const t = await theme()
    const extras: StylesExtras = {
      dxfs: [],
      tableStyles: new Map([['OfficeViewSlicer', { name: 'OfficeViewSlicer', pivot: false, table: false, elements: {
        wholeTable: { dxf: { border: { left: { style: 'thin', color: { rgb: 'FF00B050' } }, right: { style: 'thin', color: { rgb: 'FF00B050' } }, top: { style: 'thin', color: { rgb: 'FF00B050' } }, bottom: { style: 'thin', color: { rgb: 'FF00B050' } } } } },
        headerRow: { dxf: { font: { bold: true, color: { rgb: 'FF7030A0' } }, border: { bottom: { style: 'thin', color: { rgb: 'FFFF0000' } } } } },
      } }]]),
      slicerStyles: new Map([['OfficeViewSlicer', { name: 'OfficeViewSlicer', elements: {
        selectedItemWithData: { dxf: { fill: { pattern: 'solid', bg: { rgb: 'FFFFC000' } }, font: { color: { rgb: 'FF0000FF' } } } },
      } }]]),
    }
    const { style, appearance } = slicerAppearance('OfficeViewSlicer', extras, t)
    expect(style).toBe('OfficeViewSlicer')
    expect(appearance.border).toBe('#00B050')
    expect(appearance.header).toMatchObject({ separator: '#FF0000', text: '#7030A0', bold: true })
    expect(appearance.items.selectedItemWithData).toMatchObject({ fill: '#FFC000', text: '#0000FF' })
    // An element the custom style leaves out keeps the default built-in's (SlicerStyleLight1) value.
    expect(appearance.items.unselectedItemWithData).toMatchObject({ fill: '#FFFFFF', border: '#CCCCCC' })
  })

  test('TimeSlicerStyleOther1–2 are not Excel timeline styles: the default style with a diagnostic (Ruling K/Q)', async () => {
    const t = await theme()
    const diagnostics: Array<{ feature: string }> = []
    const result = timelineAppearance('TimeSlicerStyleOther1', noExtras, t, d => diagnostics.push(d))
    expect(result.style).toBe('TimeSlicerStyleLight1')
    expect(diagnostics.map(d => d.feature)).toEqual(['timeline-style-unknown'])
    expect(result.appearance).toEqual(timelineAppearance('TimeSlicerStyleLight1', noExtras, t).appearance)
  })
})

describe('slicer paint matches the Excel references at the measured points (±3 per channel)', () => {
  // Reference frame origins (outer top-left of the 1 px #4F81BD frame), read from the captures:
  // anchors at D1 + 1 pt / 8 pt → (186, 8); slicer-scrolled is anchored at row 1 + 0 pt → (186, 0).
  const REF_FRAME: Record<string, { x: number; y: number }> = { 'slicer-scrolled': { x: 186, y: 0 } }
  const refFrame = (id: string) => REF_FRAME[id] ?? { x: 186, y: 8 }
  // Reference item rects (1 px per pt): the item area is 192–319 (6 pt inside the frame), 31 + 19k high rows of 17,
  // two columns split it 63 | 2 | 63; slicer-scrolled starts at y 23 and narrows to 192–305 beside its scrollbar.
  const refItems = (id: string): Array<{ x: number; y: number; width: number; height: number }> => {
    const f = refFrame(id), top = f.y + 23
    if (id === 'slicer-two-column') return [0, 1, 2, 3].map(k => ({ x: 192 + (k % 2) * 65, y: top + Math.floor(k / 2) * 19, width: 63, height: 17 }))
    if (id === 'slicer-scrolled') return [0, 1, 2].map(k => ({ x: 192, y: top + k * 19, width: 114, height: 17 }))
    return Array.from({ length: id === 'slicer-no-data' ? 5 : 4 }, (_, k) => ({ x: 192, y: top + k * 19, width: 128, height: 17 }))
  }

  test('geometry: 23 pt from the top border to the first item, a 19 pt pitch of 17 pt items, 6 pt insets (refs/slicer-light1.png)', async () => {
    const doc = await parse(fixture('slicer-light1'))
    const drawing = controls(doc.sheets[0]!).slicer!
    const box = boxOf(drawing), layout = slicerLayout(slicerModel(drawing), box, PX)
    expect(layout.separatorY! - box.y).toBeCloseTo(19 * PX, 5)
    expect(layout.items.map(item => item.rect.y - box.y)).toEqual([0, 1, 2, 3].map(k => expect.closeTo((23 + 19 * k) * PX, 5)))
    for (const item of layout.items) {
      expect(item.rect.height).toBeCloseTo(17 * PX, 5)
      expect(item.rect.x - box.x).toBeCloseTo(6 * PX, 5)
      expect(box.x + box.width - (item.rect.x + item.rect.width)).toBeCloseTo(6 * PX, 5)
      expect(item.text.x - item.rect.x).toBeCloseTo(8 * PX, 5)
    }
    expect(layout.thumb).toBeUndefined()
    // On the reference's own box at 1 px per pt (x 186–325, y 8–197) the items land on the capture's pixels.
    const excel = slicerLayout(slicerModel(drawing), { x: 186, y: 8, width: 140, height: 190 }, 1)
    expect(excel.separatorY).toBe(27)
    expect(excel.items.map(item => item.rect)).toEqual([31, 50, 69, 88].map(y => ({ x: 192, y, width: 128, height: 17 })))
  })

  test('geometry: slicer-scrolled clamps startItem 2 to 1 so the last item fills the view, and narrows the items beside a scrollbar', async () => {
    const doc = await parse(fixture('slicer-scrolled'))
    const drawing = controls(doc.sheets[0]!).slicer!
    const model = slicerModel(drawing), box = boxOf(drawing), layout = slicerLayout(model, box, PX)
    // refs/slicer-scrolled.png: 3 visible rows of 4 items show Dates, Cherries, Bananas.
    expect(layout.items.map(item => model.items[item.index]!.label)).toEqual(['Dates', 'Cherries', 'Bananas'])
    // Item right edge 305 (20 pt inside the frame's right edge 326); thumb x 310–317, y 39–83 (track 23–83).
    for (const item of layout.items) expect(box.x + box.width - (item.rect.x + item.rect.width)).toBeCloseTo(20 * PX, 5)
    expect(box.x + box.width - (layout.thumb!.x + layout.thumb!.width)).toBeCloseTo(8 * PX, 5)
    expect(layout.thumb!.width).toBeCloseTo(8 * PX, 5)
    expect(layout.thumb!.y + layout.thumb!.height - box.y).toBeCloseTo(84 * PX, 5)
    expect(layout.thumb!.height).toBeCloseTo(61 * 3 / 4 * PX, 5)
    // On the reference's own box (x 186–325, y 0–89): items 192–305 at y 23/42/61, thumb 310–317 from 38.25 to 84.
    const excel = slicerLayout(model, { x: 186, y: 0, width: 140, height: 90 }, 1)
    expect(excel.items.map(item => item.rect)).toEqual([23, 42, 61].map(y => ({ x: 192, y, width: 114, height: 17 })))
    expect(excel.thumb).toEqual({ x: 310, y: 38.25, width: 8, height: 45.75 })
  })

  test.each(['slicer-light1', 'slicer-dark1', 'slicer-light2', 'slicer-dark6', 'slicer-other1', 'slicer-no-data', 'slicer-two-column', 'slicer-scrolled'])('%s: frame, header, item fills, borders and text colour', async (id) => {
    const doc = await parse(fixture(id))
    const sheet = doc.sheets[0]!
    const drawing = controls(sheet).slicer!
    const model = slicerModel(drawing), box = boxOf(drawing), layout = slicerLayout(model, box, PX)
    const [actual, ref] = await Promise.all([renderPixels(doc, sheet), refPixels(id)])
    const f = refFrame(id)
    const sample = (label: string, at: [number, number], refAt: [number, number], tolerance = 3) => expectNear(`${id} ${label}`, rgb(actual, ...at), rgb(ref, ...refAt), tolerance)
    sample('frame left border', [box.x + 0.5, box.y + 60 * PX], [f.x, f.y + 60])
    sample('frame top border', [box.x + 100 * PX, box.y + 0.5], [f.x + 100, f.y])
    sample('header fill', [box.x + 75 * PX, box.y + 6 * PX], [f.x + 75, f.y + 6])
    sample('header separator', [box.x + 100 * PX, layout.separatorY! + 0.5], [f.x + 100, f.y + 19])
    const refCaption = ink(ref, f.x + 6, f.y + 3, f.x + 60, f.y + 16, [255, 255, 255])
    expectNear(`${id} caption text`, ink(actual, box.x + 6 * PX, box.y + 3 * PX, box.x + 60 * PX, box.y + 16 * PX, [255, 255, 255]), refCaption)
    const refRects = refItems(id)
    expect(layout.items).toHaveLength(refRects.length)
    layout.items.forEach((item, k) => {
      const r = item.rect, rr = refRects[k]!
      const label = `item ${k} (${model.items[item.index]!.label}, ${item.state})`
      // Fill 3 pt inside the right edge (text is clipped 4 pt before it), top border 6 pt in (clear of the
      // 3 pt corner), left border at mid-height.
      sample(`${label} fill`, [r.x + r.width - 3 * PX, r.y + r.height / 2], [rr.x + rr.width - 3, rr.y + 8])
      sample(`${label} top border`, [r.x + r.width - 6 * PX, r.y + 0.5], [rr.x + rr.width - 6, rr.y])
      sample(`${label} left border`, [r.x + 0.5, r.y + r.height / 2], [rr.x, rr.y + 8])
      const fill = rgb(ref, rr.x + rr.width - 3, rr.y + 8)
      const refText = ink(ref, rr.x + 7, rr.y + 2, rr.x + Math.min(60, rr.width - 8), rr.y + rr.height - 3, fill)
      expectNear(`${label} text`, ink(actual, r.x + 7 * PX, r.y + 2 * PX, r.x + Math.min(60, rr.width - 8) * PX, r.y + r.height - 3 * PX, fill), refText)
    })
    if (id === 'slicer-scrolled') {
      // refs/slicer-scrolled.png: the grey #BCBCBC scrollbar thumb at x 310–317.
      sample('scrollbar thumb', [layout.thumb!.x + layout.thumb!.width / 2, layout.thumb!.y + layout.thumb!.height / 2], [313, 60])
    }
  }, 30_000)
})

describe('timeline paint matches the Excel references at the measured points', () => {
  // Reference frame: the 120 pt anchor at H8 + 62 pt … L16 + 40 pt → outer box x 450–649, y 105–225.
  const F = { x: 450, y: 105 }

  test('geometry: separator 22 pt down, the 13 pt bar 78 pt down and 6 pt in, the scrollbar 5 pt under it (refs/timeline-light1-months.png)', async () => {
    const doc = await parse(fixture('timeline-light1-months'))
    const drawing = controls(doc.sheets[0]!).timeline!
    const box = boxOf(drawing), layout = timelineLayout(timelineModel(drawing), box, PX)
    expect(layout.separatorY! - box.y).toBeCloseTo(22 * PX, 5)
    expect(layout.bar.y - box.y).toBeCloseTo(78 * PX, 5)
    expect(layout.bar.height).toBeCloseTo(13 * PX, 5)
    expect(layout.bar.x - box.x).toBeCloseTo(6 * PX, 5)
    expect(box.x + box.width - (layout.bar.x + layout.bar.width)).toBeCloseTo(7 * PX, 5)
    expect(layout.track!.y - (layout.bar.y + layout.bar.height)).toBeCloseTo(6 * PX, 5)
    expect(layout.track!.height).toBeCloseTo(5 * PX, 5)
    expect(layout.thumb!.height).toBeCloseTo(7 * PX, 5)
    expect(layout.thumb!.x).toBeCloseTo(layout.bar.x, 5)
    // Months: dividers every 39 pt (495, 534, 573, 612); twelve periods of 2024, Jan first.
    expect(layout.pitch).toBeCloseTo(TIMELINE_PERIOD_PT.months * PX, 5)
    expect(TIMELINE_PERIOD_PT.months).toBe(39)
    expect(layout.total).toBe(12)
    expect(layout.first).toBe(0)
    expect(layout.periods[0]).toMatchObject({ start: Date.UTC(2024, 0, 1), label: 'JAN', group: '2024' })
    // The periods in view: 187 / 39 = 4.8 of them, and the next one.
    expect(layout.periods).toHaveLength(6)
    expect(layout.periods.every(period => period.selected)).toBe(true)
    // Laid out on the reference's own box at 1 px per pt (x 450–649, y 105–225), the layout lands on the capture's
    // pixels: bar 456–642 and 183–195, handle 451–455, dividers 495/534/573/612, track 202–206, thumb 201–207 and
    // 456–528 (the visible share of the bar: 187 / 39 of 12 periods).
    const excel = timelineLayout(timelineModel(drawing), { x: 450, y: 105, width: 200, height: 121 }, 1)
    expect(excel.bar).toEqual({ x: 456, y: 183, width: 187, height: 13 })
    expect(excel.handles[0]).toEqual({ x: 451, y: 183, width: 5, height: 13 })
    expect(excel.periods.slice(0, 4).map(period => period.x + period.width)).toEqual([495, 534, 573, 612])
    expect(excel.track).toEqual({ x: 456, y: 202, width: 187, height: 5 })
    expect(excel.thumb).toMatchObject({ x: 456, y: 201, height: 7 })
    expect(Math.abs(excel.thumb!.width - 73)).toBeLessThanOrEqual(3)
    // refs/timeline-dark2-days.png: day dividers every 21 pt (476, 497, 518, …).
    expect(TIMELINE_PERIOD_PT.days).toBe(21)
  })

  test.each(['timeline-light1-months', 'timeline-dark2-days', 'timeline-light3-years'])('%s: frame, header, bar, handle, scrollbar and label colours', async (id) => {
    const doc = await parse(fixture(id))
    const sheet = doc.sheets[0]!
    const drawing = controls(sheet).timeline!
    const model = timelineModel(drawing), box = boxOf(drawing), layout = timelineLayout(model, box, PX)
    const [actual, ref] = await Promise.all([renderPixels(doc, sheet), refPixels(id)])
    const at = (x: number, y: number) => rgb(actual, x, y)
    // The 1 px frame and separator are anti-aliased in the timeline captures (#5887BF / #7095C2 beside the
    // crisp #4F81BD of the slicer frame): the painter draws the style's accent, compared within ±12 here.
    expectNear(`${id} frame left border`, at(box.x + 0.5, box.y + 60 * PX), rgb(ref, F.x, F.y + 60), 12)
    expectNear(`${id} header separator`, at(box.x + 100 * PX, layout.separatorY! + 0.5), rgb(ref, F.x + 100, F.y + 22), 12)
    expectNear(`${id} header fill`, at(box.x + 100 * PX, box.y + 4 * PX), rgb(ref, F.x + 100, F.y + 4))
    // The bar is a vertical gradient in Excel, a flat [approx] fill here (Ruling R): inside the gradient's range.
    const top = rgb(ref, F.x + 20, F.y + 78), bottom = rgb(ref, F.x + 20, F.y + 90), bar = at(layout.bar.x + 14 * PX, layout.bar.y + layout.bar.height / 2)
    for (let c = 0; c < 3; c++) expect(bar[c]!, `${id} bar channel ${c}: ${hex(bar)} within ${hex(top)}…${hex(bottom)}`).toBeGreaterThanOrEqual(Math.min(top[c]!, bottom[c]!) - 3)
    for (let c = 0; c < 3; c++) expect(bar[c]!).toBeLessThanOrEqual(Math.max(top[c]!, bottom[c]!) + 3)
    // Selection handle at x 451–455: Excel draws a textured grip (#C1C1C1 → #6C6C6C with light lines), painted flat
    // [approx] in its centre-line grey #9E9E9E, compared on that line (y 190). Scrollbar track (pale accent) at y 202–206.
    const handle = layout.handles[0]!
    expectNear(`${id} selection handle`, at(handle.x + handle.width / 2, handle.y + handle.height / 2), rgb(ref, F.x + 3, F.y + 85))
    expectNear(`${id} scrollbar track`, at(layout.track!.x + layout.track!.width - 10 * PX, layout.track!.y + layout.track!.height / 2), rgb(ref, F.x + 182, F.y + 99))
    if (layout.thumb && model.level === 'months') {
      // Thumb #AFB3B7 (x 456–528 at the months level).
      expectNear(`${id} scrollbar thumb`, at(layout.thumb.x + 10 * PX, layout.thumb.y + layout.thumb.height / 2), rgb(ref, F.x + 16, F.y + 99))
    }
    if (model.level !== 'years') {
      // A period divider: accent tint 0.4 (#95B3D7 light1, #DA9694 dark2) at the end of the first period.
      const x = layout.periods[0]!.x + layout.periods[0]!.width + 0.5
      const refX = F.x + 6 + TIMELINE_PERIOD_PT[model.level] - (model.level === 'days' ? 1 : 0)
      expectNear(`${id} period divider`, at(x, layout.bar.y + layout.bar.height / 2), rgb(ref, refX, F.y + 85))
    }
    // Label text is Excel's UI language (zh-TW captures): only its colour is compared, never its glyphs.
    const periodRow = (p: Pixels, x0: number, y0: number, x1: number, y1: number) => ink(p, x0, y0, x1, y1, [255, 255, 255])
    expectNear(`${id} period labels`, periodRow(actual, box.x + 6 * PX, box.y + 62 * PX, box.x + box.width - 8 * PX, box.y + 76 * PX), periodRow(ref, F.x + 6, F.y + 62, F.x + 198, F.y + 76))
    if (model.level !== 'years') expectNear(`${id} year label`, periodRow(actual, box.x + 6 * PX, box.y + 45 * PX, box.x + 60 * PX, box.y + 59 * PX), periodRow(ref, F.x + 6, F.y + 45, F.x + 60, F.y + 59))
  }, 30_000)
})

describe('slicer straddling tiles paints identically at 175%/DPR2 (Review Focus 4)', () => {
  test('left and right tiles split through the slicer equal one wide surface', async () => {
    const doc = await parse(fixture('slicer-light1'))
    const sheet = doc.sheets[0]!
    const drawing = controls(sheet).slicer!
    const box = boxOf(drawing)
    const geometry = buildSheetGeometry(sheet)
    if (geometry.kind !== 'ready') throw new Error('expected ready geometry')
    const zoom = 1.75, dpr = 2, scale = zoom * dpr
    // Tiles are 512 device px wide on the global device grid: the seam is a whole device pixel through the slicer.
    const seam = Math.round((box.x + box.width / 2) * scale) / scale, tileCss = 512 / scale, start = seam - tileCss
    const logicalHeight = Math.ceil(box.y + box.height + 4), height = Math.ceil(logicalHeight * scale)
    const render = (x: number, widthCss: number, widthPx: number) => {
      const canvas = createCanvas(widthPx, height), ctx = canvas.getContext('2d')
      ctx.scale(scale, scale)
      renderSheetViewport(sheet, ctx as unknown as CanvasRenderingContext2D, geometry, { box: { x, y: 0, width: widthCss, height: logicalHeight }, zoom, deviceScale: dpr })
      return ctx.getImageData(0, 0, widthPx, height).data
    }
    const left = render(start, tileCss, 512), right = render(seam, tileCss, 512), whole = render(start, tileCss * 2, 1024)
    let differing = 0, worst = 0, where = ''
    for (let y = 0; y < height; y++) for (let x = 0; x < 512; x++) for (let c = 0; c < 4; c++) {
      for (const [tile, offset, side] of [[left, 0, 'L'], [right, 512, 'R']] as const) {
        const d = Math.abs(tile[(y * 512 + x) * 4 + c]! - whole[(y * 1024 + offset + x) * 4 + c]!)
        if (d) differing++
        if (d > worst) { worst = d; where = `${side} ${x},${y},${c}` }
      }
    }
    // The slicer reaches both tiles (its frame colour on each side of the seam).
    const seen = (data: Uint8ClampedArray | Uint8Array) => { for (let i = 0; i < data.length; i += 4) if (within([data[i]!, data[i + 1]!, data[i + 2]!], [0x4F, 0x81, 0xBD], 2)) return true; return false }
    expect(seen(left) && seen(right)).toBe(true)
    expect(worst, `max channel difference ${worst} at ${where} (${differing} differing samples)`).toBe(0)
  }, 60_000)
})

describe('mc:AlternateContent: the Choice needs the pivots chunk, otherwise the Fallback is drawn', () => {
  test('the slicer and timeline requirements are a per-call input (a14 / sle / sle15 / tsle)', () => {
    expect(XLSX_REQUIREMENTS).not.toContain('tsle')
    expect(xlsxRequirements(false)).toEqual(XLSX_REQUIREMENTS)
    expect(xlsxRequirements(true)).toEqual([...XLSX_REQUIREMENTS, ...XLSX_CONTROL_REQUIREMENTS])
    const choice = (requires: string, ns: string) => {
      const root = parseXml(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:${requires}="${ns}" Requires="${requires}"/></mc:AlternateContent>`)
      return getChildren(getChildren(root, 'AlternateContent')[0] ?? root, 'Choice')[0]!
    }
    const slicer = choice('a14', 'http://schemas.microsoft.com/office/drawing/2010/main'), timeline = choice('tsle', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer')
    expect(supportedChoiceRequirements(slicer, XLSX_REQUIREMENTS).supported).toBe(false)
    expect(supportedChoiceRequirements(timeline, XLSX_REQUIREMENTS).supported).toBe(false)
    expect(supportedChoiceRequirements(slicer, xlsxRequirements(true)).supported).toBe(true)
    expect(supportedChoiceRequirements(timeline, xlsxRequirements(true)).supported).toBe(true)
  })

  test('with the chunk loaded the Choice is taken: coverage reports the slicer and timeline as choices', async () => {
    const doc = await parse(fixture('slicer-light1'))
    const sheet = doc.sheets[0]!
    const { slicer, timeline } = controls(sheet)
    expect(slicer?.source).toMatchObject({ element: 'graphicFrame', representation: 'choice', name: 'Product' })
    expect(timeline?.source).toMatchObject({ element: 'graphicFrame', representation: 'choice', name: 'Date' })
    expect(sheet.drawingCoverage!.filter(entry => entry.scope === 'original').map(entry => [entry.feature, entry.status, entry.selectedRepresentation, entry.representation]))
      .toEqual([['slicer', 'native', 'slicer', 'choice'], ['timeline', 'native', 'timeline', 'choice']])
  })

  test('an a14 Choice that is not a slicer frame keeps its Fallback on a slicer sheet', async () => {
    const shape = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="9" name="Equation"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="9" name="EquationFallback"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent>'
    const anchor = `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${shape}<xdr:clientData/></xdr:twoCellAnchor>`
    const doc = await parse(await edited(fixture('slicer-light1'), { 'xl/drawings/drawing1.xml': text => text.replace('</xdr:wsDr>', `${anchor}</xdr:wsDr>`) }))
    const drawings = doc.sheets[0]!.drawings!
    expect(drawings.map(d => d.source.name)).toEqual(['Product', 'Date', 'EquationFallback'])
    expect(drawings[2]!.source.representation).toBe('fallback')
  })

  test('without the pivots chunk the Fallback shape is drawn (exports and renders that never loaded it)', async () => {
    vi.resetModules()
    vi.doMock('../src/xlsx/pivots/index', () => { throw new Error('pivots chunk unavailable') })
    try {
      const fresh = await import('../src/xlsx/parse')
      const { OfficePackage: FreshPackage } = await import('../src/core/zip')
      const doc = await fresh.parseXlsx(await FreshPackage.load(fixture('slicer-light1')))
      const sheet = doc.sheets[0]!
      const drawings = sheet.drawings!
      expect(drawings.map(d => [d.source.element, d.source.representation, d.content?.kind])).toEqual([['sp', 'fallback', undefined], ['sp', 'fallback', undefined]])
      expect(drawings[0]!.textBody?.paragraphs.flatMap(p => p.runs.map(r => r.text)).join('')).toContain('This shape represents a slicer')
      expect(sheet.drawingCoverage!.filter(entry => entry.scope === 'original').map(entry => [entry.feature, entry.status, entry.representation]))
        .toEqual([['slicer', 'fallback', 'fallback'], ['timeline', 'fallback', 'fallback']])
    } finally {
      vi.doUnmock('../src/xlsx/pivots/index')
      vi.resetModules()
    }
  })

  test('a slicer whose cache part is missing paints its Fallback and adds slicer-cache-missing', async () => {
    const doc = await parse(await edited(fixture('slicer-light1'), { 'xl/slicerCaches/slicerCache1.xml': null }))
    const sheet = doc.sheets[0]!
    const [slicer, timeline] = sheet.drawings!
    expect(slicer!.source).toMatchObject({ element: 'sp', representation: 'fallback' })
    expect(slicer!.content).toBeUndefined()
    expect(controlOf(timeline, 'timeline')).toBeDefined()
    expect(sheet.diagnostics?.filter(d => d.feature === 'slicer-cache-missing')).toHaveLength(1)
    // The Fallback paints (its green-bordered white rectangle), never blank space.
    const pixels = await renderPixels(doc, sheet)
    const box = boxOf(slicer!)
    expect(hex(rgb(pixels, box.x + box.width / 2, box.y + box.height - 4))).toBe('FFFFFF')
  })

  test('an OLAP slicer cache fails closed: Fallback plus slicer-cache-unsupported', async () => {
    const olap = (text: string) => text.replace(/<data>.*<\/data>/s, '<data><olap pivotCacheId="1"><levels count="1"><level uniqueName="[Product].[Product]" sourceCaption="Product" count="0"/></levels><selections count="1"><selection n="[Product].[All]"/></selections></olap></data>')
    const doc = await parse(await edited(fixture('slicer-light1'), { 'xl/slicerCaches/slicerCache1.xml': olap }))
    const sheet = doc.sheets[0]!
    expect(sheet.drawings![0]!.source.representation).toBe('fallback')
    expect(sheet.diagnostics?.map(d => d.feature)).toContain('slicer-cache-unsupported')
  })

  test('a slicer item index outside the pivot cache field fails closed with slicer-item-invalid', async () => {
    const doc = await parse(await edited(fixture('slicer-light1'), { 'xl/slicerCaches/slicerCache1.xml': text => text.replace('<i x="0"/>', '<i x="9"/>') }))
    const sheet = doc.sheets[0]!
    expect(sheet.drawings![0]!.source.representation).toBe('fallback')
    expect(sheet.diagnostics?.map(d => d.feature)).toContain('slicer-item-invalid')
  })

  test('ensurePivotEngine also loads the chunk for a sheet with only slicer or timeline refs (Ruling AH)', async () => {
    vi.resetModules()
    try {
      const loader = await import('../src/xlsx/pivotLoader')
      expect(loader.pivotEngine()).toBeUndefined()
      await loader.ensurePivotEngine([{ name: 'S', rows: [], cols: [], merges: [], mergeRanges: [], timelineRefs: [{ partPath: 'xl/timelines/timeline1.xml' }] }])
      expect(loader.pivotEngine()).toBeDefined()
    } finally { vi.resetModules() }
  })
})
