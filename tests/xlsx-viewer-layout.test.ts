/**
 * S5.1 canonical text/layout records and viewport queries (spec §§5.3–5.5,
 * 11.1–11.2; plan S5.1). Object-sheet frames go through `renderSheetViewport`;
 * source-store frames (F1/F3/F6) through `queryViewport`.
 */
import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { computeMetrics, createCellLayoutHost, renderSheet, renderSheetViewport } from '../src/xlsx/render'
import {
  buildSheetGeometry, planViewport, queryViewport, sourceSetsForSlots, VIEWPORT_DETAIL_LIMITS,
  type SheetGeometry, type StoreViewportOutcome,
} from '../src/xlsx/view/viewport'
import { buildCellRecords, LayoutCache, placedFromViewportCells, type CellLayoutHost } from '../src/xlsx/view/textLayout'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics, type ProjectedSheetMetrics } from '../src/xlsx/view/projection'
import { createRowOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { MAX_COLS, MAX_ROWS } from '../src/xlsx/view/coordinates'
import type { CacheIdentity, LogicalViewport, OverviewQuery, SourceStore } from '../src/xlsx/view/contracts'
import type { XlsxCell, XlsxCellStyle, XlsxColumnSpec, XlsxRow, XlsxSheet } from '../src/xlsx/types'
import { createFixture, type CellReader } from './helpers/xlsx-viewer-fixtures'

const OWNER: CacheIdentity = { docId: 'doc', sheetId: 'sheet', valueGen: 0, fontGen: 0, viewGen: 0 }

function cell(row: number, col: number, value: XlsxCell['value'], style?: XlsxCellStyle): XlsxCell {
  return { ref: `R${row}C${col}`, row, col, value, styleIndex: 0, ...(style ? { style } : {}) }
}
function sheetOf(rows: XlsxRow[], cols: XlsxColumnSpec[] = [], extra: Partial<XlsxSheet> = {}): XlsxSheet {
  return { name: 'Layout', rows, cols, merges: [], mergeRanges: [], ...extra }
}
function ready(sheet: XlsxSheet, options?: Parameters<typeof buildSheetGeometry>[1]): SheetGeometry {
  const geometry = buildSheetGeometry(sheet, options)
  if (geometry.kind !== 'ready') throw new Error(`geometry not ready: ${JSON.stringify(geometry.diagnostics)}`)
  return geometry
}
function view(x: number, y: number, width = 320, height = 200, zoom = 1, deviceScale = 1): LogicalViewport {
  return { box: { x, y, width, height }, zoom, deviceScale }
}

interface Painted { text: string; x: number; y: number; rawX: number; rawY: number }
/** Records fillText (absolute and raw arguments) and every translate magnitude. */
function recorder(width = 320, height = 200) {
  const canvas = createCanvas(width, height)
  const raw = canvas.getContext('2d')
  const texts: Painted[] = []
  const translations: number[] = []
  let t = { tx: 0, ty: 0, sx: 1, sy: 1 }
  const stack: typeof t[] = []
  const ctx = new Proxy(raw, {
    get(target, property) {
      if (property === 'save') return () => { stack.push({ ...t }); target.save() }
      if (property === 'restore') return () => { t = stack.pop() ?? t; target.restore() }
      if (property === 'translate') return (x: number, y: number) => { translations.push(Math.abs(x), Math.abs(y)); t.tx += x * t.sx; t.ty += y * t.sy; target.translate(x, y) }
      if (property === 'scale') return (x: number, y: number) => { t.sx *= x; t.sy *= y; target.scale(x, y) }
      if (property === 'fillText') return (text: string, x: number, y: number) => { texts.push({ text, x: t.tx + x * t.sx, y: t.ty + y * t.sy, rawX: x, rawY: y }); target.fillText(text, x, y) }
      const value = Reflect.get(target, property, target)
      return typeof value === 'function' ? value.bind(target) : value
    },
    set(target, property, value) { return Reflect.set(target, property, value, target) },
  }) as unknown as CanvasRenderingContext2D
  return { ctx, texts, translations, canvas, find: (text: string) => texts.filter(item => item.text === text) }
}
function diffPixels(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
  let diff = 0
  for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2]) diff++
  return diff
}
/** The static full-sheet crop of `box` against the interactive pane. */
function cropParity(sheet: XlsxSheet, box: { x: number; y: number; width: number; height: number }, cache?: LayoutCache) {
  const m = computeMetrics(sheet)
  const full = createCanvas(Math.max(m.widthPx, box.x + box.width), Math.max(m.heightPx, box.y + box.height))
  renderSheet(sheet, full.getContext('2d') as never, { ...m, widthPx: full.width, heightPx: full.height })
  const pane = createCanvas(box.width, box.height)
  const result = renderSheetViewport(sheet, pane.getContext('2d') as never, ready(sheet), { box, zoom: 1, deviceScale: 1 }, {}, { owner: OWNER, cache })
  const a = full.getContext('2d').getImageData(box.x, box.y, box.width, box.height).data
  const b = pane.getContext('2d').getImageData(0, 0, box.width, box.height).data
  return { diff: diffPixels(a, b), result, ink: Array.from(b).some(v => v < 200) }
}

describe('canonical records (object sheet)', () => {
  test('each painted cell has one record with source identity, Float64 bounds and the painted text', () => {
    const sheet = sheetOf([
      { index: 0, cells: [cell(0, 0, 'Alpha'), cell(0, 2, 42)] },
      { index: 3, cells: [cell(3, 1, true)] },
    ])
    const geometry = ready(sheet)
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, geometry, view(0, 0), {}, { owner: OWNER })
    expect(result.kind).toBe('detail')
    const ids = result.records.map(record => record.id)
    expect(ids).toEqual([0 * MAX_COLS + 0, 0 * MAX_COLS + 2, 3 * MAX_COLS + 1])
    const b = result.records[1]
    expect(b.source).toEqual({ row: 0, col: 2 })
    expect(b.owner).toEqual(OWNER)
    expect(b.logicalBounds).toEqual({ x: 128, y: 0, width: 64, height: 20 })
    // Canonical extraction and Canvas2D text agree.
    expect(result.records.map(record => record.text[0]?.formattedText)).toEqual(['Alpha', '42', 'TRUE'])
    expect(out.texts.map(item => item.text).sort()).toEqual(['42', 'Alpha', 'TRUE'])
    for (const record of result.records) {
      const text = record.text[0]
      expect(text.cellId).toBe(record.id)
      expect(text.valueGen).toBe(OWNER.valueGen)
      expect(text.fontGen).toBe(OWNER.fontGen)
      expect(text.runs[0].graphemeBoundaries[0]).toBe(0)
    }
  })

  test('lone surrogates survive formatting, grapheme segmentation and painting', () => {
    const value = 'a\uD800b\uDC00'
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, value)] }])
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, ready(sheet), view(0, 0), {}, { owner: OWNER })
    const text = result.records[0].text[0]
    expect(text.rawText).toBe(value)
    expect(text.formattedText).toBe(value)
    expect(text.formattedText.includes('�')).toBe(false)
    expect(text.runs[0].graphemeBoundaries).toEqual([0, 1, 2, 3, 4])
    expect(out.find(value)).toHaveLength(1)
  })

  test('wrap, rotation and stacked text: records carry the layout and replay the static pixels', () => {
    const long = 'Wrapped words across several lines'
    const sheet = sheetOf([
      { index: 0, heightPt: 60, customHeight: true, cells: [
        cell(0, 0, long, { numFmtId: 0, wrapText: true }),
        cell(0, 1, 'Tilted', { numFmtId: 0, textRotation: 45 }),
        cell(0, 2, 'Stack', { numFmtId: 0, textRotation: 255 }),
        cell(0, 3, 'Down', { numFmtId: 0, textRotation: 180 }),
      ] },
    ], [{ min: 0, max: 0, widthChars: 12 }])
    // The static bitmap ends at the used grid (89 + 3 x 64 px, one 80px row).
    const { diff, result, ink } = cropParity(sheet, { x: 0, y: 0, width: 281, height: 80 })
    expect(diff).toBe(0)
    expect(ink).toBe(true)
    const [wrapped, tilted, stacked, down] = result.records.map(record => record.text[0].runs[0])
    expect(wrapped.wrap).toBe(true)
    expect(wrapped.lines).toBeGreaterThan(1)
    expect(tilted.rotation).toBeCloseTo(-Math.PI / 4, 12)
    expect(stacked.stacked).toBe(true)
    expect(down.rotation).toBeCloseTo(Math.PI / 2, 12)
    // Clipped layouts clip to the cell box.
    expect(result.records[0].text[0].clip).toEqual(result.records[0].logicalBounds)
  })

  test('General formatting fits the measured cell width; width is a layout key', () => {
    const value = 123456789.123456
    const wide = sheetOf([{ index: 0, cells: [cell(0, 0, value)] }], [{ min: 0, max: 0, widthChars: 30 }])
    const narrow = sheetOf([{ index: 0, cells: [cell(0, 0, value)] }], [{ min: 0, max: 0, widthChars: 6 }])
    const cache = new LayoutCache()
    const first = renderSheetViewport(wide, recorder().ctx, ready(wide), view(0, 0), {}, { owner: OWNER, cache })
    const second = renderSheetViewport(narrow, recorder().ctx, ready(narrow), view(0, 0), {}, { owner: { ...OWNER, viewGen: 1 }, cache })
    const wideText = first.records[0].text[0].formattedText
    const narrowText = second.records[0].text[0].formattedText
    expect(wideText).toBe('123456789.1')
    expect(narrowText).not.toBe(wideText)
    expect(narrowText.length).toBeLessThan(wideText.length)
    expect(second.work.layoutsBuilt).toBe(1)
    // Static and interactive paint the same formatted string.
    const out = recorder()
    renderSheet(narrow, out.ctx)
    expect(out.texts.map(item => item.text)).toEqual([narrowText])
  })

  test('a merge whose anchor is offscreen paints from one record without visiting its area', () => {
    const sheet = sheetOf([{ index: 0, cells: [cell(0, 0, 'Giant anchor')] }], [], {
      merges: ['A1:XFD1048576'], mergeRanges: [{ minRow: 0, minCol: 0, maxRow: MAX_ROWS - 1, maxCol: MAX_COLS - 1 }],
    })
    const geometry = ready(sheet)
    const y = geometry.metrics.rows.start(500_000), x = geometry.metrics.cols.start(8000)
    const result = renderSheetViewport(sheet, recorder().ctx, geometry, view(x, y), {}, { owner: OWNER })
    expect(result.kind).toBe('detail')
    expect(result.records).toHaveLength(1)
    const record = result.records[0]
    expect(record.source).toEqual({ row: 0, col: 0 })
    expect(record.logicalBounds).toEqual({ x: 0, y: 0, width: MAX_COLS * 64, height: MAX_ROWS * 20 })
    expect(result.work.visitedRecords).toBeLessThanOrEqual(2)
    expect(result.work.records).toBe(1)
  })

  test('hidden rows produce no record, text or hit target', () => {
    const sheet = sheetOf([
      { index: 0, cells: [cell(0, 0, 'Shown')] },
      { index: 1, hidden: true, cells: [cell(1, 0, 'Hidden')] },
      { index: 2, cells: [cell(2, 0, 'After')] },
    ])
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, ready(sheet), view(0, 0), {}, { owner: OWNER })
    expect(result.records.map(record => record.text[0].formattedText)).toEqual(['Shown', 'After'])
    expect(result.records[1].logicalBounds.y).toBe(20)
    expect(out.find('Hidden')).toEqual([])
  })
})

describe('layout cache across frames', () => {
  const sheet = () => sheetOf(Array.from({ length: 6 }, (_, r) => ({ index: r, cells: Array.from({ length: 4 }, (_, c) => cell(r, c, `r${r}c${c}`, { numFmtId: 0, wrapText: c === 0 })) })))
  test('a second frame with the complete identity rebuilds nothing; a changed cell rebuilds alone', () => {
    const cache = new LayoutCache()
    const base = sheet()
    const first = renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: OWNER, cache })
    expect(first.work.layoutsBuilt).toBe(24)
    const again = renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: OWNER, cache })
    expect(again.work.layoutsBuilt).toBe(0)
    expect(again.work.layoutsReused).toBe(24)
    // A scrolled frame reuses every layout that stays in view.
    const scrolled = renderSheetViewport(base, recorder().ctx, ready(base), view(0, 20), {}, { owner: OWNER, cache })
    expect(scrolled.work.layoutsBuilt).toBe(0)
    const changed = sheet()
    changed.rows[2].cells[1] = cell(2, 1, 'changed')
    const next = renderSheetViewport(changed, recorder().ctx, ready(changed), view(0, 0), {}, { owner: { ...OWNER, valueGen: 1 }, cache })
    expect(next.work.layoutsBuilt).toBe(1)
    expect(next.records.find(record => record.id === 2 * MAX_COLS + 1)!.text[0].formattedText).toBe('changed')
    // Records carry the publication owner they were served under.
    expect(next.records.every(record => record.owner.valueGen === 1)).toBe(true)
  })
  test('fontGen invalidates dependent layouts; new measurements take effect', () => {
    const cache = new LayoutCache()
    const base = sheet()
    renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: OWNER, cache })
    const fontsChanged = renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: { ...OWNER, fontGen: 1 }, cache })
    expect(fontsChanged.work.layoutsBuilt).toBe(24)
    expect(cache.stats.invalidations).toBeGreaterThanOrEqual(1)
    // A host whose metrics changed (late font) feeds the next generation only.
    let scale = 1
    const host: CellLayoutHost = { layoutBody: (c, w, h) => ({ x: 0, y: 0, w, h, formattedText: String(c.value), text: { kind: 'plain', text: String(c.value), font: 'f', color: '#000', tx: 3, ty: 15, width: 10 * scale } }) }
    const placed = { cell: cell(0, 0, 'x'), source: { row: 0, col: 0 }, x: 0, y: 0, width: 64, height: 20 }
    expect(cache.layout(OWNER, placed, host).text).toMatchObject({ width: 10 })
    scale = 2
    expect(cache.layout(OWNER, placed, host).text).toMatchObject({ width: 10 })
    expect(cache.layout({ ...OWNER, fontGen: 7 }, placed, host).text).toMatchObject({ width: 20 })
  })
  test('another document or sheet never reuses a layout', () => {
    const cache = new LayoutCache()
    const base = sheet()
    renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: OWNER, cache })
    expect(renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: { ...OWNER, sheetId: 'other' }, cache }).work.layoutsBuilt).toBe(24)
    expect(renderSheetViewport(base, recorder().ctx, ready(base), view(0, 0), {}, { owner: { ...OWNER, sheetId: 'other', docId: 'doc2' }, cache }).work.layoutsBuilt).toBe(24)
  })
  test('cached replay is pixel-identical to the static crop', () => {
    const cache = new LayoutCache()
    const base = sheet()
    const box = { x: 0, y: 0, width: 256, height: 120 }
    expect(cropParity(base, box, cache).diff).toBe(0)
    const replay = cropParity(base, box, cache)
    expect(replay.result.work.layoutsBuilt).toBe(0)
    expect(replay.diff).toBe(0)
  })
})

describe('Float64 logical positions, tile-local paint coordinates', () => {
  test('the last row at a fractional origin paints from local coordinates without Float32 drift', () => {
    const last = MAX_ROWS - 1
    const sheet = sheetOf([{ index: last, cells: [cell(last, 0, 'Last row')] }])
    const geometry = ready(sheet)
    const top = geometry.metrics.rows.start(last)
    expect(top).toBe(last * 20)
    // A fractional origin 37.3 px above the row at zoom 0.37.
    const originY = top - 37.3
    expect(Math.abs(Math.fround(originY) - originY)).toBeGreaterThan(0.1)
    const out = recorder(400, 200)
    const result = renderSheetViewport(sheet, out.ctx, geometry, { box: { x: 0, y: originY, width: 400, height: 200 }, zoom: 0.37, deviceScale: 2 }, {}, { owner: OWNER })
    expect(result.kind).toBe('detail')
    expect(result.records[0].logicalBounds.y).toBe(top)
    // No global coordinate reaches the canvas: every translate and text argument is tile-local.
    expect(Math.max(0, ...out.translations)).toBeLessThan(1000)
    const painted = out.find('Last row')
    expect(painted).toHaveLength(1)
    expect(Math.abs(painted[0].rawY)).toBeLessThan(1000)
    // Exact Float64 local offset: (top - originY) + baseline, no 2px quantization.
    const baseline = 20 - (20 - 10 * 96 / 72) / 2
    expect(painted[0].y).toBeCloseTo((top - originY) + baseline, 9)
  })
})

describe('detail admission (SheetZoom x DPR, no font-size floor)', () => {
  test('a small sheet at SheetZoom 0.25 still paints all of its text', () => {
    const sheet = sheetOf(Array.from({ length: 5 }, (_, r) => ({ index: r, cells: [cell(r, 0, `t${r}`), cell(r, 1, r)] })))
    const out = recorder()
    const result = renderSheetViewport(sheet, out.ctx, ready(sheet), view(0, 0, 1280, 800, 0.25, 1), {}, { owner: OWNER })
    expect(result.kind).toBe('detail')
    expect(result.coarse).toEqual([])
    expect(out.texts.map(item => item.text).sort()).toEqual(['0', '1', '2', '3', '4', 't0', 't1', 't2', 't3', 't4'])
  })
  test('the legacy fit-width path at an effective 0.375 scale is unchanged and paints every cell', () => {
    const sheet = sheetOf(Array.from({ length: 5 }, (_, r) => ({ index: r, cells: [cell(r, 0, `t${r}`)] })))
    const scaled = recorder(48, 40)
    scaled.ctx.scale(0.375, 0.375)
    renderSheet(sheet, scaled.ctx)
    expect(scaled.texts.map(item => item.text)).toEqual(['t0', 't1', 't2', 't3', 't4'])
    expect(scaled.texts.some(item => /Zoom in/.test(item.text))).toBe(false)
  })
  test('sub-detail runs coarsen alone; the rest of the pane keeps detail', () => {
    // Ten 1px rows (0.25 device px at 25%) above ordinary 20px rows (5 device px).
    const rows: XlsxRow[] = Array.from({ length: 10 }, (_, index) => ({ index, heightPt: 0.75, customHeight: true, cells: [cell(index, 0, `tiny${index}`)] }))
    rows.push({ index: 10, cells: [cell(10, 0, 'Normal')] })
    const sheet = sheetOf(rows)
    const geometry = ready(sheet)
    const plan = planViewport(geometry, view(0, 0, 1280, 800, 0.25, 1))
    if (plan.kind !== 'detail') throw new Error(`expected detail, got ${plan.kind}`)
    expect(plan.cells.map(placed => placed.source.row)).toEqual([10])
    expect(plan.coarse.length).toBeGreaterThan(0)
    for (const group of plan.coarse) {
      expect(group.slots.end).toBeLessThan(10)
      // A group is at least four device pixels unless its whole span is smaller
      // (ten 1px rows = 2.5 device px: one group row).
      expect(group.bounds.height).toBe(10)
      expect(group.bounds.width * 0.25).toBeGreaterThanOrEqual(4 - 1e-9)
    }
    // The groups tile the ten tiny rows across the pane width, nothing else.
    expect(new Set(plan.coarse.map(group => `${group.slots.start}-${group.slots.end}`))).toEqual(new Set(['0-9']))
    expect(plan.coarse.reduce((sum, group) => sum + group.bounds.width, 0)).toBe(1280)
    const out = recorder(1280, 800)
    const result = renderSheetViewport(sheet, out.ctx, geometry, view(0, 0, 1280, 800, 0.25, 1), {}, { owner: OWNER })
    expect(out.find('Normal')).toHaveLength(1)
    expect(out.texts.some(item => item.text.startsWith('tiny'))).toBe(false)
    expect(out.find('Zoom in or inspect a cell for detail')).toHaveLength(1)
    expect(result.kind).toBe('detail')
  })
  test('the active cell keeps one exact detail record in a coarse frame', () => {
    const rows: XlsxRow[] = Array.from({ length: 50 }, (_, index) => ({ index, heightPt: 0.75, customHeight: true, cells: [cell(index, 0, `v${index}`)] }))
    const sheet = sheetOf(rows)
    const geometry = ready(sheet)
    const plan = planViewport(geometry, view(0, 0, 320, 60, 1, 1), 'D', { activeCell: { row: 7, col: 0 } })
    expect(plan.kind === 'coarse' || (plan.kind === 'detail' && plan.cells.length === 0)).toBe(true)
    if (plan.kind !== 'coarse' && plan.kind !== 'detail') throw new Error('unexpected')
    expect(plan.active).toMatchObject({ source: { row: 7, col: 0 }, y: 7, height: 1 })
    expect(plan.active!.cell.value).toBe('v7')
  })
  test('a tile holding more than the per-tile record cap is coarse even under the frame cap', () => {
    // 5px cells: one 512px tile holds 10,000 records (> 5,000) though the frame cap is 20,000.
    const rows: XlsxRow[] = Array.from({ length: 100 }, (_, r) => ({ index: r, heightPt: 3.75, customHeight: true, cells: Array.from({ length: 100 }, (_, c) => cell(r, c, 1)) }))
    const sheet = sheetOf(rows, [{ min: 0, max: 99, widthChars: 0 }])
    const geometry = ready(sheet)
    expect(geometry.metrics.cols.size(0)).toBe(5)
    expect(geometry.metrics.rows.size(0)).toBe(5)
    const plan = planViewport(geometry, view(0, 0, 500, 500))
    expect(plan.kind).toBe('coarse')
    expect(plan.diagnostics.at(-1)!.message).toMatch(/tile/)
    // Below the tile cap the same cells are detail.
    expect(planViewport(geometry, view(0, 0, 300, 300)).kind).toBe('detail')
  })
  test('profiles carry the §11.1 per-frame and per-tile caps', () => {
    expect(VIEWPORT_DETAIL_LIMITS.D).toMatchObject({ records: 20_000, pages: 32, gridlines: 4096, tile: { records: 5000, pages: 8 } })
    expect(VIEWPORT_DETAIL_LIMITS.M).toMatchObject({ records: 10_000, pages: 16, gridlines: 2048, tile: { records: 2500, pages: 4 } })
  })
})

// ---------------------------------------------------------------------------
// Source-store frames
// ---------------------------------------------------------------------------

function storeMetrics(options: { rowCss?: number; colCss?: number; order?: Parameters<typeof createSheetMetrics>[0]['order']; visibility?: Parameters<typeof createSheetMetrics>[0]['visibility'] } = {}): ProjectedSheetMetrics {
  return createSheetMetrics({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: options.rowCss ?? 20 }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: options.colCss ?? 64 }),
    order: options.order, visibility: options.visibility,
  })
}
/** Counts every backend cell read; `cellAt` is the content access. */
function probe(reader: CellReader) {
  const counts = { cellAt: 0, isPopulated: 0 }
  return {
    counts,
    read: {
      cellAt: (row: number, col: number) => { counts.cellAt++; return reader.cellAt(row, col) },
      isPopulated: (row: number, col: number) => { counts.isPopulated++; return reader.isPopulated(row, col) },
      styleOf: (row: number, col: number) => reader.styleOf(row, col),
    },
  }
}
function storeOwner(store: SourceStore): CacheIdentity {
  return { docId: store.stamp.docId, sheetId: store.stamp.sheetId, valueGen: store.stamp.valueGen, fontGen: 0, viewGen: 0 }
}
function signal() { return new AbortController().signal }
function viewportOf(outcome: StoreViewportOutcome) {
  if (outcome.kind !== 'viewport') throw new Error(`expected viewport, got ${outcome.kind}: ${JSON.stringify(outcome.diagnostics)}`)
  return outcome
}

describe('source-store viewport at the edges of F1/F3/F6', () => {
  test('F1: the XFD1048576 corner is detail at its exact logical box, bounded pages and reads', async () => {
    const fixture = createFixture('F1', { nnz: 2000 })
    const p = probe(fixture)
    const store = fixture.createStore({ read: p.read })
    const metrics = storeMetrics()
    const box = { x: metrics.cols.totalCss() - 1440, y: metrics.rows.totalCss() - 900, width: 1440, height: 900 }
    const outcome = viewportOf(await queryViewport(store, metrics, { box, zoom: 1, deviceScale: 2 }, storeOwner(store), { signal: signal() }))
    expect(outcome.completeness).toBe('complete')
    const corner = outcome.cells.find(item => item.kind === 'cell' && item.source.row === MAX_ROWS - 1 && item.source.col === MAX_COLS - 1)
    expect(corner).toMatchObject({ kind: 'cell', bounds: { x: (MAX_COLS - 1) * 64, y: (MAX_ROWS - 1) * 20, width: 64, height: 20 } })
    expect(outcome.work.pages).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.D.pages)
    // Reads are bounded by the window (45 rows x 23 columns), never the sheet.
    expect(p.counts.isPopulated).toBeLessThanOrEqual(46 * 24)
    expect(outcome.work.visitedRecords).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.D.records)
  })
  test('F3: the XFD column edge is complete detail for every visible dense cell', async () => {
    const fixture = createFixture('F3')
    const store = fixture.createStore()
    const metrics = storeMetrics()
    const box = { x: metrics.cols.totalCss() - 640, y: 0, width: 640, height: 400 }
    const outcome = viewportOf(await queryViewport(store, metrics, { box, zoom: 1, deviceScale: 1 }, storeOwner(store), { signal: signal() }))
    expect(outcome.completeness).toBe('complete')
    const cells = outcome.cells.filter(item => item.kind === 'cell')
    expect(cells).toHaveLength(20 * 10)
    expect(cells.every(item => item.source.col >= MAX_COLS - 10)).toBe(true)
    expect(outcome.work.records).toBe(200)
    expect(outcome.work.gridlines).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.D.gridlines)
  })
  test('F6: the full-axis provider is detail at XFD1048576 and only the window is read', async () => {
    const fixture = createFixture('F6')
    const p = probe(fixture)
    const store = fixture.createStore({ read: p.read })
    const metrics = storeMetrics()
    const box = { x: metrics.cols.totalCss() - 320, y: metrics.rows.totalCss() - 200, width: 320, height: 200 }
    const outcome = viewportOf(await queryViewport(store, metrics, { box, zoom: 1, deviceScale: 1 }, storeOwner(store), { signal: signal() }))
    expect(outcome.completeness).toBe('complete')
    expect(outcome.cells.filter(item => item.kind === 'cell')).toHaveLength(50)
    expect(p.counts.cellAt).toBe(50)
  })
  test('F6: tiny positive dimensions at 25% zoom resolve analytically to complete coarse summaries', async () => {
    const fixture = createFixture('F6')
    const p = probe(fixture)
    const store = fixture.createStore({ read: p.read })
    // 0.5 x 1 css cells: 1,800 x 1,440 logical cells fit a 900 x 1440 css pane at 25%.
    const metrics = storeMetrics({ rowCss: 0.5, colCss: 1 })
    const outcome = viewportOf(await queryViewport(store, metrics, view(0, 0, 1440 * 4, 900 * 4, 0.25, 2), storeOwner(store), { signal: signal() }))
    expect(outcome.completeness).toBe('complete')
    expect(outcome.cells).toEqual([])
    expect(outcome.groups.length).toBeGreaterThan(0)
    expect(outcome.groups.length).toBeLessThanOrEqual(VIEWPORT_DETAIL_LIMITS.D.overviewGroups)
    expect(outcome.groups.every(group => group.content === 'populated')).toBe(true)
    // Exact source membership covers the whole window once.
    const cellsCovered = outcome.groups.reduce((sum, group) => sum + (group.content === 'unknown' ? 0 : group.populatedCount), 0)
    expect(cellsCovered).toBe(7200 * 5760)
    expect(p.counts.cellAt).toBe(0)
    expect(p.counts.isPopulated).toBe(0)
    expect(outcome.work.records).toBe(0)
    expect(outcome.work.runsVisited).toBeLessThan(64)
    expect(outcome.diagnostics).toEqual(expect.arrayContaining([expect.objectContaining({ code: 'detail-limit' })]))
  })
  test('unknown overview coverage is loading, never blank; failures and aborts are explicit', async () => {
    const fixture = createFixture('F6')
    const metrics = storeMetrics({ rowCss: 0.5, colCss: 1 })
    const pane = view(0, 0, 1440, 900, 0.25, 1)
    const pending = fixture.createStore({ pending: [{ r0: 0, r1: 100, c0: 0, c1: 100 }] })
    const loading = viewportOf(await queryViewport(pending, metrics, pane, storeOwner(pending), { signal: signal() }))
    expect(loading.completeness).toBe('loading')
    expect(loading.groups.some(group => group.content === 'unknown')).toBe(true)
    expect(loading.groups.some(group => group.content === 'blank')).toBe(false)
    const failing = fixture.createStore({ errors: [{ range: { r0: 0, r1: 100, c0: 0, c1: 100 }, code: 'page-error', retryable: true }] })
    const failed = viewportOf(await queryViewport(failing, metrics, pane, storeOwner(failing), { signal: signal() }))
    expect(failed.completeness).toBe('incomplete')
    expect(failed.error).toMatchObject({ code: 'page-error', retryable: true })
    const store = fixture.createStore()
    const controller = new AbortController()
    controller.abort()
    expect((await queryViewport(store, metrics, pane, storeOwner(store), { signal: controller.signal })).kind).toBe('cancelled')
  })
  test('a window needing more pages than the frame budget declares the remainder loading', async () => {
    // Display slot s shows source 2s (then the odd rows): consecutive display rows
    // are disjoint source rows, so the window needs one page request per row.
    const fixture = createFixture('F3', { rowCount: 4096, colCount: 8 })
    const store = fixture.createStore()
    const n = 4096
    const permutation = new Uint32Array(n).map((_, slot) => slot < n / 2 ? 2 * slot : 2 * (slot - n / 2) + 1)
    const metrics = storeMetrics({ order: createRowOrder([{ r0: 0, r1: n - 1, sourceAtOffset: permutation }]) })
    const outcome = viewportOf(await queryViewport(store, metrics, view(0, 0, 512, 1600), storeOwner(store), { signal: signal(), profile: 'M' }))
    expect(outcome.work.pages).toBe(VIEWPORT_DETAIL_LIMITS.M.pages)
    expect(outcome.completeness).toBe('loading')
    expect(outcome.cells.some(item => item.kind === 'unknown')).toBe(true)
    // Rows that were read sit at their display slot.
    const second = outcome.cells.find(item => item.kind === 'cell' && item.source.row === 2)
    expect(second).toMatchObject({ bounds: { y: 20 } })
    const unknown = outcome.cells.filter(item => item.kind === 'unknown')
    expect(unknown).toHaveLength(80 - VIEWPORT_DETAIL_LIMITS.M.pages)
    expect(unknown.every(item => item.kind === 'unknown' && item.pageId === 'page-budget')).toBe(true)
  })
})

describe('overview groups map display pixels to exact source sets', () => {
  const orderMetrics = (hidden: number[] = []) => storeMetrics({
    order: createRowOrder([{ r0: 0, r1: 3, sourceAtOffset: new Uint32Array([0, 3, 1, 2]) }]),
    visibility: visibilityFromHidden(createIntervalSet(hidden.map(r => ({ r0: r, r1: r }))), MAX_ROWS),
  })
  test('displayed [0,3,1,2]: grouping display rows 1-2 is source {1,3}, never the bounding interval', () => {
    const sets = sourceSetsForSlots(orderMetrics(), 1, 2)
    expect(sets).toEqual([{ kind: 'list', r0: 1, r1: 3, rows: new Uint32Array([1, 3]) }])
  })
  test('filtering changes exact membership', () => {
    // Display slot 2 (source 1) hidden by the canonical mask.
    expect(sourceSetsForSlots(orderMetrics([2]), 1, 2)).toEqual([{ kind: 'list', r0: 3, r1: 3, rows: new Uint32Array([3]) }])
    // Identity rows outside the band stay intervals, split around a hidden slot.
    expect(sourceSetsForSlots(orderMetrics([10]), 8, 12)).toEqual([{ kind: 'intervals', intervals: [{ r0: 8, r1: 9 }, { r0: 11, r1: 12 }] }])
    expect(sourceSetsForSlots(orderMetrics(), 2, 6)).toEqual([
      { kind: 'intervals', intervals: [{ r0: 4, r1: 6 }] },
      { kind: 'list', r0: 1, r1: 2, rows: new Uint32Array([1, 2]) },
    ])
  })
  test('every overview domain sent to the store is an exact display-group membership', async () => {
    const fixture = createFixture('F6')
    const inner = fixture.createStore()
    const domains: OverviewQuery['groups'][number][] = []
    const store: SourceStore = { ...inner, queryOverview: (request, options) => { domains.push(...request.groups); return inner.queryOverview(request, options) } }
    const n = 8
    const metrics = createSheetMetrics({
      rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 0.5 }),
      cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }),
      order: createRowOrder([{ r0: 0, r1: n - 1, sourceAtOffset: new Uint32Array([0, 7, 1, 6, 2, 5, 3, 4]) }]),
    })
    const outcome = viewportOf(await queryViewport(store, metrics, view(0, 0, 128, 4), storeOwner(store), { signal: signal() }))
    expect(outcome.groups.length).toBeGreaterThan(0)
    for (const domain of domains) {
      if (domain.rows.kind !== 'list') continue
      // Display slots [0..7] in 4px (8 slot) groups at DPR 1: one group holds all 8 sources.
      expect([...domain.rows.rows]).toEqual([...domain.rows.rows].sort((a, b) => a - b))
    }
    const listed = domains.filter(domain => domain.rows.kind === 'list')
    expect(listed.length).toBeGreaterThan(0)
    expect(listed.every(domain => domain.rows.kind === 'list' && domain.rows.rows.length === 8)).toBe(true)
  })
  test('late aggregates from an older viewGen never publish', async () => {
    const fixture = createFixture('F6')
    const inner = fixture.createStore()
    let release!: () => void
    const gate = new Promise<void>(resolve => { release = resolve })
    const store: SourceStore = { ...inner, queryOverview: async (request, options) => { await gate; return inner.queryOverview(request, options) } }
    let current = storeOwner(store)
    const owner = current
    const metrics = storeMetrics({ rowCss: 0.5, colCss: 1 })
    const pending = queryViewport(store, metrics, view(0, 0, 1440, 900, 0.25, 1), owner, { signal: signal(), isCurrent: candidate => candidate.viewGen === current.viewGen })
    current = { ...current, viewGen: 1 }
    release()
    const outcome = await pending
    expect(outcome.kind).toBe('stale')
  })
})

describe('store frames feed the same canonical records', () => {
  test('F6 corner cells become records with host-formatted text, cached across frames', async () => {
    const fixture = createFixture('F6')
    const store = fixture.createStore()
    const metrics = storeMetrics()
    const box = { x: metrics.cols.totalCss() - 320, y: metrics.rows.totalCss() - 200, width: 320, height: 200 }
    const outcome = viewportOf(await queryViewport(store, metrics, { box, zoom: 1, deviceScale: 1 }, storeOwner(store), { signal: signal() }))
    const ctx = createCanvas(8, 8).getContext('2d') as unknown as CanvasRenderingContext2D
    const host = createCellLayoutHost(ctx)
    const cache = new LayoutCache()
    const placed = placedFromViewportCells(outcome.cells)
    expect(placed).toHaveLength(50)
    const records = buildCellRecords(placed, outcome.owner, host, cache)
    const corner = records.find(record => record.source.row === MAX_ROWS - 1 && record.source.col === MAX_COLS - 1)!
    expect(corner.id).toBe(2 ** 34 - 1)
    expect(corner.logicalBounds).toEqual({ x: (MAX_COLS - 1) * 64, y: (MAX_ROWS - 1) * 20, width: 64, height: 20 })
    const value = fixture.cellAt(MAX_ROWS - 1, MAX_COLS - 1)!.value
    expect(corner.text[0].formattedText).toBe(String(value))
    expect(cache.stats.built).toBe(50)
    buildCellRecords(placedFromViewportCells(outcome.cells), outcome.owner, host, cache)
    expect(cache.stats.built).toBe(50)
    expect(cache.stats.reused).toBe(50)
  })
  test('F5: a merge anchor offscreen paints across the merged box; covered cells disappear', async () => {
    const fixture = createFixture('F5')
    const store = fixture.createStore()
    const metrics = storeMetrics()
    const origin = viewportOf(await queryViewport(store, metrics, view(0, 0, 320, 200), storeOwner(store), { signal: signal() }))
    const anchor = origin.cells.find(item => item.source.row === 0 && item.source.col === 0)
    expect(anchor).toMatchObject({ kind: 'cell', bounds: { x: 0, y: 0, width: 128, height: 40 } })
    expect(origin.cells.some(item => item.source.row <= 1 && item.source.col <= 1 && item !== anchor)).toBe(false)
    // Inside B100:CW200000 far from its anchor: nothing but the merge.
    const y = metrics.rows.start(150_000), x = metrics.cols.start(50)
    const inside = viewportOf(await queryViewport(store, metrics, view(x, y, 320, 200), storeOwner(store), { signal: signal() }))
    expect(inside.completeness).toBe('complete')
    expect(inside.cells).toEqual([])
    expect(inside.merges).toEqual([{ range: { minRow: 99, maxRow: 199_999, minCol: 1, maxCol: 100 }, bounds: { x: 64, y: 99 * 20, width: 100 * 64, height: (199_999 - 99 + 1) * 20 } }])
    expect(inside.work.pages).toBeLessThanOrEqual(2)
  })
})
