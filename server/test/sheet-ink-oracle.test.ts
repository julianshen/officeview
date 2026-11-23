/**
 * Stage 4b Task 14 — the ink oracle (S4 §6.7; plan P16). Pixel-based, and independent of `coverage.ts`: a cell's ink
 * is the set of pixels that differ between two renders of the projected sheet with its drawings removed, one as is and
 * one with that cell blank. The cell is ink-withheld when the authored render (drawings on) differs from the
 * drawings-removed render at any of those pixels. Every ink-withheld cell must already be blank in the text view the
 * server's text ops read (`DocumentHost`'s own view of the sheet, at the session sizes). The converse is not asserted:
 * the analytic rule fails closed (an unfilled outline strictly inside a cell covers it and may touch no ink).
 *
 * Renders: the projected (unblanked) saved-values sheet, painted as a tile is (profile 'D', its frame budget, the light
 * palette), at zoom 1, dpr 1, with one set of assets prepared over the projected workbook, so blanking one cell moves
 * no other cell's CF. Each render rebuilds its geometry over its own sheet copy (the plan's drawings and cells come from
 * it). Where the authored and drawings-removed renders agree everywhere, no cell can be ink-withheld and the per-cell
 * renders are skipped (exact). Fixtures render a band per text row (the used width + 512 px, the row ± 256 px);
 * the corpus renders tile (0,0) at zoomStep 5, dpr 1 (512² CSS px) only.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { prepareSheetAssets } from '../../src/components/xlsx/sheetAssets'
import { TILE_PX } from '../../src/stream/sheet-wire'
import type { RenderSheetAssets } from '../../src/xlsx/render'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import type { LogicalViewport } from '../../src/xlsx/view/contracts'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { cellSourceRow } from '../../src/xlsx/view/sheetCells'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { buildSheetGeometry, frameBudgetFor, planViewport, type SheetGeometryResult } from '../../src/xlsx/view/viewport'
import { DocumentHost, type RenderRefusal } from '../src/render'
import { clockEpochMs, type SheetClock } from '../src/sheet/clock'
import type { TextView } from '../src/sheet/text-view'
import { TILE_PALETTE_COLORS } from '../src/sheet/tile'
import { sessionSizesOf, type WireSizes } from '../src/sheet/view-state'
import {
  absoluteOverB2, BOLD_REACH_LABEL, cfHiddenFormat, cfHiddenValue, cfLongColumn, hiddenFormat, labelBesideShape, LONG_LABEL, mergeTail, offViewportCovered,
  rotatedBox, slidUnderAnchor, spillUnderShape, textGrid, TEXT_GRID, unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

const UTC: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const hosts: DocumentHost[] = []
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })
const registerFont = () => { throw new Error('ink oracle: embedded fonts are not registered') }

type Box = { x: number; y: number; width: number; height: number }
const holdsContent = (c: XlsxCell) => (c.value !== null && c.value !== undefined) || c.formula !== undefined
const blankCell = (c: XlsxCell): XlsxCell => ({ ref: c.ref, row: c.row, col: c.col, value: null, styleIndex: c.styleIndex, ...(c.style ? { style: c.style } : {}) })
/** `sheet` with the cell at source (row, col) blank. */
function blanked(sheet: XlsxSheet, row: number, col: number): XlsxSheet {
  return { ...sheet, rows: sheet.rows.map((r) => r.cells.some((c) => cellSourceRow(r, c) === row && c.col === col) ? { ...r, cells: r.cells.map((c) => cellSourceRow(r, c) === row && c.col === col ? blankCell(c) : c) } : r) }
}

/** One render of `box` (CSS px, zoom 1, dpr 1) as a tile paints: RGBA, or undefined when the plan is not 'detail'
 * (a degraded frame paints no cell text, on a tile either). */
function paint(host: DocumentHost, sheet: XlsxSheet, geometry: SheetGeometryResult, assets: RenderSheetAssets, box: Box): Uint8ClampedArray | undefined {
  const w = Math.ceil(box.width), h = Math.ceil(box.height)
  const canvas = host.backend.createCanvas(w, h)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const viewport: LogicalViewport = { box, zoom: 1, deviceScale: 1 }
  const frameBudget = frameBudgetFor('D')
  const plan = planViewport(geometry, viewport, 'D', { frameBudget })
  if (plan.kind !== 'detail') return undefined
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  renderSheetViewport(sheet, ctx, geometry, viewport, assets, { colors: TILE_PALETTE_COLORS.light, preparedPlan: plan, profile: 'D', cache: new LayoutCache(), frameBudget })
  return ctx.getImageData(0, 0, w, h).data
}
/** The pixel indexes where `a` and `b` differ. */
function differ(a: Uint8ClampedArray, b: Uint8ClampedArray): number[] {
  const out: number[] = []
  for (let i = 0; i < a.length; i += 4) if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) out.push(i)
  return out
}

interface Opened { host: DocumentHost; docKey: string; saved: XlsxDocument }
async function opened(name: string, bytes: Uint8Array): Promise<Opened> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open(name, bytes, 'visible-only', { sheetMode: 'tiles' })
  return { host, docKey: o.docKey, saved: (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse }
}
/** The text view the server's text ops read for (sheet, sizes): the host's own cached view. */
function hostView(o: Opened, sheetIndex: number, sizes: WireSizes): Promise<TextView> {
  return (o.host as unknown as { textViewEntry(docKey: string, i: number, clock: SheetClock, sizes: WireSizes): { view: Promise<TextView> } }).textViewEntry(o.docKey, sheetIndex, UTC, sizes).view
}
/** Whether (row, col) holds no content in the view. */
function blankIn(view: TextView): (row: number, col: number) => boolean {
  if (view.blankAll) return () => true
  const content = new Set<string>()
  for (const r of view.sheet.rows) for (const c of r.cells) if (holdsContent(c)) content.add(`${cellSourceRow(r, c)}:${c.col}`)
  return (row, col) => !content.has(`${row}:${col}`)
}

interface SheetInk { sheetIndex: number; textCells: number; checked: number; withheld: Array<[number, number]>; violations: Array<[number, number, string]>; skipped: number }
/**
 * The oracle over sheet `sheetIndex` of `o` at `sizes`. `regions(cells, geometry)` groups the text cells to check into
 * render boxes. Returns the ink-withheld cells and every one of them not blank in the view (a violation).
 */
async function inkOf(o: Opened, sheetIndex: number, sizes: WireSizes, regions: (cells: Array<[number, number, XlsxCell]>, g: SheetGeometryResult) => Array<{ box: Box; cells: Array<[number, number, XlsxCell]> }>): Promise<SheetInk> {
  const projected = o.saved.sheets[sheetIndex]!
  const view = await hostView(o, sheetIndex, sizes)
  const blank = blankIn(view)
  const opts = sizes.length ? { sizes: sessionSizesOf(sizes) } : {}
  const authored = buildSheetGeometry(projected, opts)
  const out: SheetInk = { sheetIndex, textCells: 0, checked: 0, withheld: [], violations: [], skipped: 0 }
  if (authored.kind !== 'ready') return out
  const cells: Array<[number, number, XlsxCell]> = []
  for (const r of projected.rows) for (const c of r.cells) {
    const row = cellSourceRow(r, c)
    if (holdsContent(c) && authored.metrics.rows.size(row) > 0 && authored.metrics.cols.size(c.col) > 0) cells.push([row, c.col, c])
  }
  out.textCells = cells.length
  // No drawing on the sheet: the authored and drawings-removed renders are the same sheet (exact).
  if (!projected.drawings?.length) return out
  const bare: XlsxSheet = { ...projected, drawings: [] }
  const bareGeometry = buildSheetGeometry(bare, opts)
  const lease = await prepareSheetAssets(o.saved, { sheetIndex, registerFont, cfClock: { epochNowMs: clockEpochMs(UTC), timeZone: UTC.timeZone }, chartOwner: o.saved })
  try {
    for (const { box, cells: group } of regions(cells, authored)) {
      const a = paint(o.host, projected, authored, lease.assets, box), r0 = paint(o.host, bare, bareGeometry, lease.assets, box)
      if (!a || !r0) { out.skipped += group.length; continue }
      const drawn = new Set(differ(a, r0))
      if (!drawn.size) { out.checked += group.length; continue }
      for (const [row, col, cell] of group) {
        const one = blanked(bare, row, col)
        const rc = paint(o.host, one, buildSheetGeometry(one, opts), lease.assets, box)
        if (!rc) { out.skipped++; continue }
        out.checked++
        if (!differ(r0, rc).some((i) => drawn.has(i))) continue
        out.withheld.push([row, col])
        if (!blank(row, col)) out.violations.push([row, col, String(cell.value)])
      }
    }
  } finally { lease.dispose() }
  return out
}

/** Fixture regions: one band per text row, the used width + 512 px wide, the row ± 256 px tall. */
function bands(cells: Array<[number, number, XlsxCell]>, g: SheetGeometryResult): Array<{ box: Box; cells: Array<[number, number, XlsxCell]> }> {
  if (g.kind !== 'ready') return []
  const m = g.metrics
  const right = Math.max(...cells.map(([, c]) => m.cols.start(c) + m.cols.size(c))) + 512
  const byRow = new Map<number, Array<[number, number, XlsxCell]>>()
  for (const x of cells) byRow.set(x[0], [...(byRow.get(x[0]) ?? []), x])
  return [...byRow].map(([row, group]) => {
    const y = Math.max(0, m.rows.start(row) - 256)
    return { box: { x: 0, y, width: right, height: m.rows.start(row) + m.rows.size(row) + 256 - y }, cells: group }
  })
}
/** Corpus regions: tile (0,0) at zoomStep 5, dpr 1, and at most `max` of the text cells that meet it. */
const tile00 = (max: number) => (cells: Array<[number, number, XlsxCell]>, g: SheetGeometryResult): Array<{ box: Box; cells: Array<[number, number, XlsxCell]> }> => {
  if (g.kind !== 'ready') return []
  const m = g.metrics
  const inside = cells.filter(([r, c]) => m.cols.start(c) < TILE_PX && m.rows.start(r) < TILE_PX).slice(0, max)
  return inside.length ? [{ box: { x: 0, y: 0, width: TILE_PX, height: TILE_PX }, cells: inside }] : []
}

const FIXTURES: ReadonlyArray<{ name: string; doc: () => Promise<Uint8Array>; sizes?: WireSizes }> = [
  { name: 'offViewportCovered', doc: offViewportCovered },
  { name: 'slidUnderAnchor', doc: slidUnderAnchor },
  { name: 'slidUnderAnchor col B 8 px', doc: slidUnderAnchor, sizes: [{ axis: 'col', index: 1, px: 8 }] },
  { name: 'mergeTail', doc: mergeTail },
  { name: 'hiddenFormat', doc: hiddenFormat },
  { name: 'cfHiddenValue', doc: cfHiddenValue },
  { name: 'cfHiddenFormat', doc: cfHiddenFormat },
  { name: 'absoluteOverB2', doc: absoluteOverB2 },
  { name: 'absoluteOverB2 col A 300 px', doc: absoluteOverB2, sizes: [{ axis: 'col', index: 0, px: 300 }] },
  { name: 'textGrid', doc: textGrid },
  { name: 'cfLongColumn top1-hidden', doc: () => cfLongColumn('top1-hidden') },
  { name: 'unfilledShapeOverText', doc: unfilledShapeOverText },
  { name: 'rotatedBox', doc: rotatedBox },
  { name: 'unresolvedAnchor', doc: unresolvedAnchor },
  { name: 'labelBesideShape long', doc: () => labelBesideShape(LONG_LABEL) },
  { name: 'labelBesideShape short', doc: () => labelBesideShape('Short') },
  { name: 'labelBesideShape long rotated', doc: () => labelBesideShape(LONG_LABEL, { rotated: true }) },
  { name: 'labelBesideShape bold reach', doc: () => labelBesideShape(BOLD_REACH_LABEL, { cf: 'bold' }) },
  { name: 'spillUnderShape left', doc: () => spillUnderShape('left') },
  { name: 'spillUnderShape right', doc: () => spillUnderShape('right') },
]

describe('ink oracle (S4 §6.7, P16)', () => {
  test('ink-withheld cells are blank in the text view', async () => {
    const found = new Map<string, SheetInk[]>()
    for (const f of FIXTURES) {
      const o = await opened(f.name, await f.doc())
      const sheets: SheetInk[] = []
      for (let i = 0; i < o.saved.sheets.length; i++) sheets.push(await inkOf(o, i, f.sizes ?? [], bands))
      found.set(f.name, sheets)
      for (const s of sheets) {
        expect([f.name, s.sheetIndex, s.violations]).toEqual([f.name, s.sheetIndex, []])
        expect([f.name, s.sheetIndex, s.skipped]).toEqual([f.name, s.sheetIndex, 0])
      }
    }
    const withheld = (name: string, sheet = 0) => found.get(name)![sheet]!.withheld
    // Non-vacuity: the oracle sees the occluded cells of the fixtures built for it, and only those.
    const [cr, cc] = TEXT_GRID.covered
    expect([withheld('textGrid', 0), withheld('textGrid', 1)]).toEqual([[[cr, cc]], [[cr, cc]]])
    expect(withheld('labelBesideShape long')).toEqual([[1, 0]])
    expect(withheld('labelBesideShape short')).toEqual([])
    expect(withheld('spillUnderShape left')).toEqual([[2, 0]])
    expect(withheld('spillUnderShape right')).toEqual([])
    expect(withheld('rotatedBox')).toEqual([[0, 0], [1, 0], [2, 0]])
    expect(withheld('slidUnderAnchor')).toEqual([])
    expect(withheld('slidUnderAnchor col B 8 px')).toEqual([[0, 4]])
    expect(found.get('offViewportCovered')![0]!.checked).toBe(2)
    console.log(`ink oracle (fixtures): ${[...found].map(([n, s]) => `${n} ${s.map((x) => `${x.checked}/${x.textCells} checked, ${x.withheld.length} withheld`).join(' | ')}`).join('; ')}`)
  }, 30_000)

  test('ink-withheld cells are blank in the text view on the corpus', async () => {
    const dir = resolve(import.meta.dir, '../../corpus')
    const names = readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.xlsx')).sort()
    let files = 0, sheets = 0, checked = 0, withheld = 0, drawn = 0, skipped = 0
    const refused: string[] = []
    for (const name of names) {
      let o: Opened
      try { o = await opened(name, new Uint8Array(readFileSync(resolve(dir, name)))) } catch (e) {
        // sheet-corpus.test.ts pins that such a file is refused identically in units mode.
        expect([name, (e as RenderRefusal).code]).toEqual([name, 'unsupported-format'])
        refused.push(name)
        continue
      }
      files++
      for (let i = 0; i < Math.min(o.saved.sheets.length, 20); i++) {
        const s = await inkOf(o, i, [], tile00(64))
        expect([name, i, s.violations]).toEqual([name, i, []])
        sheets++; checked += s.checked; withheld += s.withheld.length; skipped += s.skipped
        if (o.saved.sheets[i]!.drawings?.length) drawn++
      }
      for (const k of [...(o.host as unknown as { docs: Map<string, unknown> }).docs.keys()]) o.host.drop(k)
    }
    expect(files).toBeGreaterThanOrEqual(1)
    console.log(`ink oracle (corpus): ${files} files, ${sheets} sheets (${drawn} with drawings), ${checked} cells checked, ${withheld} ink-withheld, ${skipped} on degraded frames; refused: ${refused.join(', ') || 'none'}`)
  }, 120_000)
})
