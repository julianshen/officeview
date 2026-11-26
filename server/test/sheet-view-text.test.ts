/**
 * Stage 4c Task 6 — tiles and text ops under the session's sort and filter view (plan P9, P10; Open item 2).
 *
 * The oracle is the local engine (`viewerControls`) over `sortGrid` parsed here apart from the server inputs, with the
 * cells the key view withholds under `visible-only` physically deleted (column E, row 12, hidden row 9). Its order and
 * visibility drive local `renderSheetViewport`, `serializeSelection` and `autoFitSize`, which the server's answers over
 * the same view must equal.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import { autoFitSize, createFitMeasure } from '../../src/components/xlsx/sheetResize'
import type { CacheIdentity, FilterSpec, RowOrder, RowVisibility, SheetView, SortSpec } from '../../src/xlsx/view/contracts'
import { serializeSelection } from '../../src/xlsx/view/copy'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { tileLogicalBox } from '../../src/xlsx/view/tileCache'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { buildSheetGeometry, frameBudgetFor, planViewport, type SheetGeometry } from '../../src/xlsx/view/viewport'
import { cellDisplayText } from '../../src/xlsx/render'
import type { WireRange, WireSelection, WireSortSpec } from '../../src/stream/sheet-wire'
import type { XlsxDocument } from '../../src/xlsx/types'
import { serverBackend } from '../src/backend'
import { decodePng } from '../src/mark/png'
import { Pool } from '../src/pool'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetTextJob, SheetTextReply } from '../src/sheet/text-ops'
import { TILE_PALETTE_COLORS, type TileJob } from '../src/sheet/tile'
import type { CanonicalView, ResolvedFilter } from '../src/sheet/view-spec'
import { absoluteAnchorXml, addDrawing, patchSheetXml, patchXlsx, rectShapeXml, SORT_GRID, sortGrid, sortGridScore } from './sheet-fixtures'

const CLOCK: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const T = 512
const signal = () => new AbortController().signal

const hosts: DocumentHost[] = []
const pools: Pool[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await Promise.all(pools.map((p) => p.dispose()))
})

interface HostDoc { derivedViews?: Map<string, unknown>; sizedGeometries?: Map<string, unknown>; textViews?: Map<string, unknown> }
async function serve(bytes: Uint8Array) {
  const host = new DocumentHost()
  hosts.push(host)
  const { docKey } = await host.open('doc', bytes, 'visible-only', { sheetMode: 'tiles' })
  const doc = () => (host as unknown as { docs: Map<string, HostDoc> }).docs.get(docKey)!
  const text = (job: SheetTextJob) => host.sheetText(docKey, CLOCK, job)
  return { host, docKey, doc, text }
}

// ---- the oracle ----------------------------------------------------------------------------------------------

/** `bytes` parsed here (saved values) with the cells `visible-only` withholds from `sortGrid` deleted. */
async function gridOracleDoc(bytes: Uint8Array): Promise<XlsxDocument> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const sheet = doc.sheets[0]!
  const gone = (r: number, c: number) => c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow
  return { ...doc, sheets: [{ ...sheet, rows: sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !gone(c.row, c.col)) })) }, ...doc.sheets.slice(1)] }
}

async function oracle(doc: XlsxDocument) {
  const source = createObjectWorkbookSource(doc, 'oracle')
  const sheetId = source.sheets[0]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  let view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  const owner = (): CacheIdentity => ({ docId: 'oracle', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; viewGen++; return snapshot() }
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish, ledger: new MainMemoryLedger() }, signal())
  const client = controls.client
  return {
    source, sheetId,
    view: () => view,
    async sort(spec: SortSpec) { return (await client.sort(owner(), spec, signal())).kind },
    async filter(spec: FilterSpec | null) { return (await client.filter(owner(), spec, signal())).kind },
    dispose: async () => { controls.dispose(); await source.release() },
  }
}

/** Local `serializeSelection` over `doc` with the oracle's order and visibility (no CF, no regions in `sortGrid`). */
async function localCopy(doc: XlsxDocument, selection: WireSelection, order: RowOrder, visibility: RowVisibility): Promise<string> {
  const source = createObjectWorkbookSource(doc, 'local')
  const sheetId = source.sheets[0]!.sheetId
  const sheet = doc.sheets[0]!
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry')
  const store = source.store(sheetId)
  const identity = { ...store.stamp, fontGen: 0, viewGen: 0 }
  const r = await serializeSelection({
    store, owner: identity, currentOwner: () => identity, allowed: () => true, signal: signal(),
    selection: { active: null, anchor: null, ranges: selection.ranges.map((x) => ({ rows: { kind: 'intervals' as const, intervals: x.rows.map(([r0, r1]) => ({ r0, r1 })) }, cols: x.cols.map(([c0, c1]) => ({ c0, c1 })) })) },
    order, visibility,
    columns: visibilityFromHidden(createIntervalSet([{ r0: SORT_GRID.zeroWidthCol, r1: SORT_GRID.zeroWidthCol }]), geometry.metrics.cols.logicalCount),
    format: (cell, allocation) => {
      const styled = geometry.cellsOf(cell.row)?.find((c) => c.col === cell.col)
      return cellDisplayText(sheet, { value: cell.value, style: styled?.style }, allocation)
    },
  })
  await source.release()
  if (r.kind !== 'ready') throw new Error(`local copy ${r.kind}`)
  return r.text
}

// ---- spec helpers ----------------------------------------------------------------------------------------------

const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const GRID = wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
const localRange = (r: WireRange) => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const sortOn = (col: number, order: 'asc' | 'desc'): WireSortSpec => ({ range: GRID, keys: [{ col, order }], caseSensitive: false })
const localSort = (s: WireSortSpec): SortSpec => ({ range: localRange(s.range), keys: s.keys.map((k) => ({ ...k })), caseSensitive: s.caseSensitive })
const containsA: ResolvedFilter = { range: GRID, columns: [[3, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]] }
const localContainsA: FilterSpec = { range: localRange(GRID), columns: new Map([[3, { kind: 'text', op: 'contains', first: 'a', wildcards: false }]]) }
const sorted = (s: WireSortSpec): CanonicalView => ({ sort: s, sortFilter: null, filter: null })
const filtered = (f: ResolvedFilter): CanonicalView => ({ sort: null, sortFilter: null, filter: f })
/** The grid's whole range (the worksheet select-all is refused by the local engine before any cell is read, 4b). */
const ALL: WireSelection = { active: [0, 0], anchor: [0, 0], ranges: [{ rows: [[0, SORT_GRID.rows - 1]], cols: [[0, SORT_GRID.cols - 1]] }] }
/** 1-based sheet rows whose D word holds an "a" (apple, banana, grape), among the shown body rows. */
const wordOf = (r: number) => `${SORT_GRID.words[r % 8]}-${r}`

function tileJob(view: CanonicalView | undefined, extra: Partial<TileJob> = {}): TileJob {
  return { sheetIndex: 0, viewGen: 1, zoomStep: 5, dpr: 1, palette: 'light', col: 0, row: 0, mark: { kind: 'off' }, clock: CLOCK, ...(view ? { view } : {}), ...extra }
}
/** Tile (0,0) at zoom 1, dpr 1 of `sheet` painted locally over `geometry`. */
function localTile(sheet: Parameters<typeof renderSheetViewport>[0], geometry: SheetGeometry, assets: Parameters<typeof renderSheetViewport>[4]): Uint8ClampedArray {
  const canvas = serverBackend().createCanvas(T, T)
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const viewport = { box: tileLogicalBox(0, 0, 1, 1), zoom: 1, deviceScale: 1 }
  const frameBudget = frameBudgetFor('D')
  const plan = planViewport(geometry, viewport, 'D', { frameBudget })
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  renderSheetViewport(sheet, ctx, geometry, viewport, assets, { colors: TILE_PALETTE_COLORS.light, preparedPlan: plan, profile: 'D', cache: new LayoutCache(), frameBudget })
  return ctx.getImageData(0, 0, T, T).data
}
const differing = (a: ArrayLike<number>, b: ArrayLike<number>) => { let n = 0; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) n++; return n }

const findOf = (r: SheetTextReply) => { if (r.op !== 'find') throw new Error(JSON.stringify(r)); return r }

describe('tiles and text ops under the view (plan P9)', () => {
  test('tiles follow the view', async () => {
    const bytes = await sortGrid()
    const s = await serve(bytes)
    const o = await oracle(await gridOracleDoc(bytes))
    try {
      const desc = sortOn(1, 'desc')
      expect(await o.sort(localSort(desc))).toBe('committed')
      // F5: a sorted view without sizes takes any viewGen.
      const got = decodePng((await s.host.renderTile(s.docKey, tileJob(sorted(desc), { viewGen: 3 }))).png).rgba
      const snap = await s.host.sheetSnapshot(s.docKey, 0, CLOCK)
      const g = buildSheetGeometry(snap.sheet, { order: o.view().order, visibility: o.view().visibility })
      if (g.kind !== 'ready') throw new Error('view geometry is a diagnostic')
      expect(differing(got, localTile(snap.sheet, g, snap.assets))).toBe(0)
      const identity = decodePng((await s.host.renderTile(s.docKey, tileJob(undefined, { viewGen: 0 }))).png).rgba
      expect(differing(got, identity)).toBeGreaterThan(1000)

      // Equal views (another object, keys in another order) share the derived view and its geometry; a different
      // filter builds a second of each.
      const again: CanonicalView = { filter: null, sortFilter: null, sort: { caseSensitive: false, keys: [{ order: 'desc', col: 1 }], range: { ...GRID } } }
      const same = decodePng((await s.host.renderTile(s.docKey, tileJob(again))).png).rgba
      expect(differing(same, got)).toBe(0)
      expect([s.doc().derivedViews?.size, s.doc().sizedGeometries?.size]).toEqual([1, 1])
      const both = decodePng((await s.host.renderTile(s.docKey, tileJob({ ...sorted(desc), filter: containsA }))).png).rgba
      expect([s.doc().derivedViews?.size, s.doc().sizedGeometries?.size]).toEqual([2, 2])
      expect(differing(both, got)).toBeGreaterThan(1000)

      // Through the pool: its worker receives the view and paints the same tile.
      const pool = new Pool({ workers: 1 }); pools.push(pool)
      await pool.open('grid', bytes, 'visible-only', { sheetMode: 'tiles' })
      const pooled = await pool.tile('s1', { docRef: 'grid', contentPolicy: 'visible-only', ...tileJob(sorted(desc)) }, 4 * T * T * 2)
      expect(differing(decodePng(pooled.png).rgba, got)).toBe(0)
    } finally { await o.dispose() }
  }, 30_000)

  test('copy, find and a11y follow the view', async () => {
    const bytes = await sortGrid()
    const s = await serve(bytes)
    const deleted = await gridOracleDoc(bytes)
    const o = await oracle(deleted)
    try {
      expect(await o.filter(localContainsA)).toBe('committed')
      const view = filtered(containsA)
      const copy = await s.text({ op: 'copy', sheetIndex: 0, selection: ALL, sizes: [], maxBytes: 4 * 1024 * 1024, view })
      if (copy.op !== 'copy' || !('text' in copy)) throw new Error(JSON.stringify(copy))
      expect(copy.text).toBe(await localCopy(deleted, ALL, o.view().order, o.view().visibility))
      expect(copy.text).not.toContain('cherry')

      // Open item 2: a term only in rows the filter hides never matches; a shown row's still does.
      const plum = { op: 'find' as const, query: 'plum-3', matchCase: false, sheetIndex: 0, sizes: [], limit: 50 }
      expect(findOf(await s.text(plum)).hits.length).toBeGreaterThan(0)
      expect(findOf(await s.text({ ...plum, view }))).toEqual({ op: 'find', hits: [], more: false })
      expect(findOf(await s.text({ ...plum, query: wordOf(8), view })).hits).toEqual([[7, 3]])

      const a11y = await s.text({ op: 'a11y', sheetIndex: 0, rows: [[0, SORT_GRID.rows - 1]], cols: [0, SORT_GRID.cols - 1], sizes: [], view })
      if (a11y.op !== 'a11y') throw new Error(JSON.stringify(a11y))
      const rows = new Set(a11y.cells.map(([r]) => r))
      for (let r = 2; r <= SORT_GRID.rows; r++) {
        const shown = wordOf(r).includes('a') && r - 1 !== SORT_GRID.hiddenRow && r - 1 !== SORT_GRID.zeroHeightRow
        expect({ r, listed: rows.has(r - 1) }).toEqual({ r, listed: shown })
      }
    } finally { await o.dispose() }
  }, 30_000)

  test('fit follows the view', async () => {
    const bytes = await sortGrid()
    const s = await serve(bytes)
    const deleted = await gridOracleDoc(bytes)
    const o = await oracle(deleted)
    try {
      // D `contains "plum"` shows only plum-3 .. plum-35; the widest D text (cherry-…, banana-…) is filtered out.
      const plum: ResolvedFilter = { range: GRID, columns: [[3, { kind: 'text', op: 'contains', first: 'plum', wildcards: false }]] }
      expect(await o.filter({ range: localRange(GRID), columns: new Map([[3, { kind: 'text', op: 'contains', first: 'plum', wildcards: false }]]) })).toBe('committed')
      const pane = { x: 0, y: 0, width: 640, height: 400 }
      const r = await s.text({ op: 'fit', sheetIndex: 0, axis: 'col', index: 3, pane, sizes: [], view: filtered(plum) })
      if (r.op !== 'fit') throw new Error(JSON.stringify(r))
      const sheet = deleted.sheets[0]!
      const geometry = buildSheetGeometry(sheet, { order: o.view().order, visibility: o.view().visibility })
      if (geometry.kind !== 'ready') throw new Error('oracle geometry')
      const ctx = serverBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
      const fit = createFitMeasure({ admit: () => true, measureText: (font, text) => { ctx.font = font; return ctx.measureText(text).width }, sheet })
      const local = autoFitSize('column', 3, { geometry, sheet, pane, knownRows: (limit) => sheet.rows.slice(0, limit).map((x) => x.index), measure: fit.measure, allocation: fit.allocation })
      expect(r.size).toBe(local)
      const unfiltered = await s.text({ op: 'fit', sheetIndex: 0, axis: 'col', index: 3, pane, sizes: [] })
      if (unfiltered.op !== 'fit') throw new Error(JSON.stringify(unfiltered))
      console.log('[4c task6] fit D', JSON.stringify({ filtered: r.size, unfiltered: unfiltered.size }))
      expect(r.size).toBeLessThan(unfiltered.size)
    } finally { await o.dispose() }
  }, 30_000)

  test('a sort that slides a cell under a fixed drawing blanks its text', async () => {
    // `UNDER` in A30 (source row 30); B30 just above the smallest shown score, so an ascending sort on B makes row 30
    // the second body row (display row 3). An absolute box over A3:A4 (display rows 3–4) stays put.
    const shownScores = Array.from({ length: SORT_GRID.rows - 1 }, (_, i) => i + 2).filter((r) => r !== 30 && r - 1 !== SORT_GRID.hiddenRow && r - 1 !== SORT_GRID.zeroHeightRow).map(sortGridScore)
    const b30 = Math.min(...shownScores) + 0.5
    const bytes = await patchXlsx(await sortGrid(), async (zip) => {
      await patchSheetXml(zip, 0, (xml) => {
        const out = xml.replace(/<c r="A30"[^>]*>.*?<\/c>/, '<c r="A30" t="inlineStr"><is><t>UNDER</t></is></c>').replace(/<c r="B30"[^>]*>.*?<\/c>/, `<c r="B30"><v>${b30}</v></c>`)
        if (out === xml) throw new Error('A30/B30 not found')
        return out
      })
      await addDrawing(zip, 0, absoluteAnchorXml([8, 44], [48, 32], rectShapeXml({ id: 2, name: 'Fixed', offPx: [8, 44], extPx: [48, 32], fill: null, line: '404040' })))
    })
    const s = await serve(bytes)
    const under = { op: 'find' as const, query: 'UNDER', matchCase: true, sheetIndex: 0, sizes: [], limit: 10 }
    expect(findOf(await s.text(under)).hits).toEqual([[29, 0]])
    const view = sorted(sortOn(1, 'asc'))
    const order = await s.host.sheetView(s.docKey, CLOCK, 'visible-only', { op: 'derive', sheetIndex: 0, view })
    if (order.op !== 'view' || !('order' in order)) throw new Error(JSON.stringify(order))
    const band = order.order.bands[0]!
    expect(band.sourceAtOffset[2 - band.r0]).toBe(29)
    expect(findOf(await s.text({ ...under, view }))).toEqual({ op: 'find', hits: [], more: false })
    // Positive control: row 30's D cell, outside the box, still reads under the same view (coverage is known, not blank-all).
    expect(findOf(await s.text({ ...under, query: wordOf(30), view })).hits).toEqual([[29, 3]])
    // The zero-height row's cells stay blank under the sort (they moved display slot, not size).
    const hidden = { ...under, query: wordOf(SORT_GRID.zeroHeightRow + 1), matchCase: false }
    expect(findOf(await s.text({ ...hidden, view })).hits).toEqual([])
  }, 30_000)
})
