/**
 * Stage 4b Task 5 — the `sheet-text` pool job and `sheet-fit` (plan P9, P12, P12a). A fit is local `autoFitSize` over
 * the text view (ruling F4: the view's blanked sheet, the base style), so it equals the local fit of the fixture with
 * its withheld cells physically deleted. Fit is a `'sheet'` op: 30 a minute, one in flight, `stale-view` and
 * `hidden-target` before the job; it is answered with copy off and never changes view state.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { autoFitSize, createFitMeasure } from '../../src/components/xlsx/sheetResize'
import { decodeSheetLayout, type StreamSheetLayout } from '../../src/stream/sheet-wire'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import { NO_SESSION_SIZES } from '../../src/xlsx/view/axisSizes'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { serverBackend } from '../src/backend'
import { DocumentHost } from '../src/render'
import { servableSheet } from '../src/sheet/snapshot'
import type { SheetClock } from '../src/sheet/clock'
import { textInputs } from '../src/sheet/text-ops'
import { sheetTextView, textDocument } from '../src/sheet/text-view'
import { Client, disposePool, start } from './harness'
import { rejectionMessage } from './fixtures/rejection'
import { cfHiddenFormat, slidUnderAnchor, tallColumn, TALL_COLUMN, textGrid, TEXT_GRID, unfilledShapeOverText, unresolvedAnchor } from './sheet-fixtures'

const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})

const VIEW = ['sheet-view@1']
const PANE = { x: 0, y: 0, width: 640, height: 400 }
let rq = 500

/** The local reference measure: the server backend's canvas (the pinned pack), independent of the code under test. */
const refCtx = serverBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
const measureText = (font: string, text: string) => { refCtx.font = font; return refCtx.measureText(text).width }

/** Local `autoFitSize` over `sheet`, as the resident viewer calls it (no scope), with the pane. */
function localFit(sheet: XlsxSheet, kind: 'row' | 'column', index: number, pane = PANE): number {
  const geometry = buildSheetGeometry(sheet)
  if (geometry.kind !== 'ready') throw new Error('fixture geometry')
  const fit = createFitMeasure({ admit: () => true, measureText, sheet })
  return autoFitSize(kind, index, { geometry, sheet, pane, knownRows: (limit) => sheet.rows.slice(0, limit).map((r) => r.index), measure: fit.measure, allocation: fit.allocation })
}
/** The fixture's projected saved-values sheets (what the host's snapshots are built from). */
async function projectedOf(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only'): Promise<{ host: DocumentHost; docKey: string; saved: XlsxDocument }> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('ref', bytes, policy, { sheetMode: 'tiles' })
  return { host, docKey: o.docKey, saved: (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse }
}
/** `sheet` with every cell `drop` names physically deleted. */
const without = (sheet: XlsxSheet, drop: (cell: XlsxCell) => boolean): XlsxSheet => ({ ...sheet, rows: sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c)) })) })

async function open(url: string) {
  const c = new Client(url)
  await c.hello({ features: VIEW })
  await c.ready()
  return c
}
const fit = async (c: Client, body: Record<string, unknown>) => (await c.text({ reqId: ++rq, type: 'sheet-fit', sheetId: '0', viewGen: 0, pane: PANE, ...body })).end

describe('sheet-fit (P9, P12)', () => {
  test('fit equals local autoFitSize on the deleted fixture', async () => {
    const bytes = await textGrid()
    const { saved } = await projectedOf(bytes)
    const [cr, cc] = TEXT_GRID.covered
    // Withheld in textGrid: the covered D20 and the tails of the merge A5:C6.
    const deleted = without(saved.sheets[0]!, (c) => (c.row === cr && c.col === cc) || (c.row >= 4 && c.row <= 5 && c.col <= 2 && !(c.row === 4 && c.col === 0)))
    // The comparison can fail: the covered D20 changes column D's fit. (`autoFitSize` skips every cell of a merge that
    // spans the fitted direction, so the tails of A5:C6 never change a fit; they are deleted for fidelity only.)
    expect(localFit(saved.sheets[0]!, 'column', cc)).not.toBe(localFit(deleted, 'column', cc))

    let t = Date.now()
    const h = await start({ doc: bytes, now: () => t })
    const c = await open(h.url)
    const answered: Array<[string, number, unknown]> = [], expected: Array<[string, number, unknown]> = []
    const one = async (axis: 'row' | 'col', index: number) => {
      t += 2_100 // the 'sheet' class allows 30 a minute: 38 fits are spread over 80 s of the injected clock
      const r = await fit(c, { axis, index })
      const hidden = axis === 'col' ? index === TEXT_GRID.hiddenCol : index === TEXT_GRID.hiddenRow
      answered.push([axis, index, r.type === 'sheet-fit-result' ? r.size : r.code])
      expected.push([axis, index, hidden ? 'hidden-target' : localFit(deleted, axis === 'col' ? 'column' : 'row', index)])
      if (r.type === 'sheet-fit-result') expect(Object.keys(r).sort()).toEqual(['reqId', 'size', 'type'])
    }
    for (let col = 0; col < TEXT_GRID.cols; col++) await one('col', col)
    for (let row = 0; row < TEXT_GRID.rows; row++) await one('row', row)
    expect(answered).toEqual(expected)
    // A pane past the sheet's extent is clamped: it answers as the extent-sized pane does.
    const g = buildSheetGeometry(deleted)
    if (g.kind !== 'ready') throw new Error('geometry')
    const extent = { x: 0, y: 0, width: g.metrics.cols.totalCss(), height: g.metrics.rows.totalCss() }
    t += 2_100
    const huge = await fit(c, { axis: 'col', index: 0, pane: { x: 0, y: 0, width: 1e300, height: 1e300 } })
    t += 2_100
    expect(huge).toEqual({ ...(await fit(c, { axis: 'col', index: 0, pane: extent })), reqId: huge.reqId })
    expect(huge).toMatchObject({ type: 'sheet-fit-result', size: localFit(deleted, 'column', 0, extent) })
    // Fit never changes view state.
    const layout = await c.sheetLayout(++rq, '0')
    expect((decodeSheetLayout(layout.parts.map((p) => p.text)) as StreamSheetLayout).viewGen).toBe(0)
    c.close(); await h.stop()
  }, 30_000)

  test('fit ignores withheld text', async () => {
    const shape = await unfilledShapeOverText()
    const { saved } = await projectedOf(shape)
    // COVERED-B2 (under the unfilled shape) is column B's longest text: the fit is the fixture's with B2 deleted.
    const deleted = without(saved.sheets[0]!, (c) => c.row === 1 && c.col === 1)
    expect(localFit(saved.sheets[0]!, 'column', 1)).toBeGreaterThan(localFit(deleted, 'column', 1))
    const h = await start({ doc: shape })
    const c = await open(h.url)
    for (let col = 0; col < 4; col++) expect(await fit(c, { axis: 'col', index: col })).toMatchObject({ type: 'sheet-fit-result', size: localFit(deleted, 'column', col) })
    c.close(); await h.stop()

    // Unknown coverage: every cell is blank, so every column fits to the blank-column size.
    const unres = await unresolvedAnchor()
    const u = await projectedOf(unres)
    const blank = localFit(without(u.saved.sheets[0]!, () => true), 'column', 0)
    expect(localFit(u.saved.sheets[0]!, 'column', 0)).not.toBe(blank)
    const g = await start({ doc: unres })
    const d = await open(g.url)
    for (let col = 0; col < 3; col++) expect(await fit(d, { axis: 'col', index: col })).toMatchObject({ type: 'sheet-fit-result', size: blank })
    d.close(); await g.stop()
  }, 30_000)

  test('fit without regions puts the pane rows first (fix round 1)', async () => {
    // More rows than AUTOFIT_MAX_ROWS: a pane scrolled to row 2901 must read it, as local autoFitSize with no regions does.
    const bytes = await tallColumn()
    const { saved } = await projectedOf(bytes)
    const sheet = saved.sheets[0]!
    const pane = { x: 0, y: TALL_COLUMN.longRow * 20 - 100, width: 640, height: 400 }
    const scrolled = localFit(sheet, 'column', 0, pane), top = localFit(sheet, 'column', 0)
    // The fixture can tell them apart: at the top the long text is past the row bound.
    expect(scrolled).toBeGreaterThan(top)
    const h = await start({ doc: bytes })
    const c = await open(h.url)
    expect(await fit(c, { axis: 'col', index: 0, pane })).toMatchObject({ type: 'sheet-fit-result', size: scrolled })
    expect(await fit(c, { axis: 'col', index: 0 })).toMatchObject({ type: 'sheet-fit-result', size: top })
    c.close(); await h.stop()
  }, 30_000)

  test('fit is on without copy', async () => {
    const bytes = await textGrid()
    const { saved } = await projectedOf(bytes)
    const h = await start({ doc: bytes, policy: { copy: false, textA11y: false } })
    const c = await open(h.url)
    expect(await fit(c, { axis: 'col', index: 0 })).toMatchObject({ type: 'sheet-fit-result', size: localFit(saved.sheets[0]!, 'column', 0) })
    c.close(); await h.stop()
  }, 30_000)

  test('fit is a sheet op', async () => {
    let t = Date.now()
    const h = await start({ doc: await textGrid(), now: () => t })
    const c = await open(h.url)
    // A collapsed column (F) and row (7) are refused hidden-target.
    t += 60
    expect(await fit(c, { axis: 'col', index: TEXT_GRID.hiddenCol })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    t += 60
    expect(await fit(c, { axis: 'row', index: TEXT_GRID.hiddenRow })).toMatchObject({ type: 'refused', code: 'hidden-target' })
    // An old viewGen is refused stale-view.
    t += 60
    expect(await fit(c, { viewGen: 3, axis: 'col', index: 0 })).toMatchObject({ type: 'refused', code: 'stale-view' })
    // Bad arguments are refused policy-refused (and charge nothing of the 'sheet' class).
    expect(await fit(c, { axis: 'col', index: 0, pane: { x: -1, y: 0, width: 1, height: 1 } })).toMatchObject({ type: 'refused', code: 'policy-refused' })
    // Fits and resizes share the class: 30 a minute, the 31st refused rate-limited.
    for (let i = 3; i < 30; i++) {
      t += 60
      expect(await fit(c, { axis: 'col', index: 0 })).toMatchObject({ type: 'sheet-fit-result' })
    }
    t += 60
    expect(await fit(c, { axis: 'col', index: 0 })).toMatchObject({ type: 'refused', code: 'rate-limited' })
    t += 60
    expect((await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '0', viewGen: 0, axis: 'col', index: 0, size: 90 })).end).toMatchObject({ type: 'refused', code: 'rate-limited' })
    // A minute later the class has room again.
    t += 61_000
    expect(await fit(c, { axis: 'col', index: 0 })).toMatchObject({ type: 'sheet-fit-result' })
    c.close(); await h.stop()
  }, 30_000)

  test('fit reads the session sizes', async () => {
    // slidUnderAnchor: E1 is clear at authored sizes; column B at 8 px slides it under the shape (current-view coverage).
    const bytes = await slidUnderAnchor()
    const { saved } = await projectedOf(bytes)
    const sheet = saved.sheets[0]!
    const withE1 = localFit(sheet, 'column', 4), withoutE1 = localFit(without(sheet, (c) => c.row === 0 && c.col === 4), 'column', 4)
    expect(withE1).not.toBe(withoutE1)
    const h = await start({ doc: bytes })
    const c = await open(h.url)
    expect(await fit(c, { axis: 'col', index: 4 })).toMatchObject({ type: 'sheet-fit-result', size: withE1 })
    expect((await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '0', viewGen: 0, axis: 'col', index: 1, size: 8 })).end).toMatchObject({ type: 'sheet-view', viewGen: 1, change: { size: 8 } })
    // The old viewGen is stale; at viewGen 1 the fit reads the sized view, where E1 is covered.
    expect(await fit(c, { axis: 'col', index: 4 })).toMatchObject({ type: 'refused', code: 'stale-view' })
    expect(await fit(c, { viewGen: 1, axis: 'col', index: 4 })).toMatchObject({ type: 'sheet-fit-result', size: withoutE1 })
    // Reset: back to the identity view (a cached view is reused, not rebuilt wrong).
    expect((await c.text({ reqId: ++rq, type: 'sheet-resize', sheetId: '0', viewGen: 1, reset: true })).end).toMatchObject({ type: 'sheet-view', viewGen: 2 })
    expect(await fit(c, { viewGen: 2, axis: 'col', index: 4 })).toMatchObject({ type: 'sheet-fit-result', size: withE1 })
    c.close(); await h.stop()
  }, 30_000)

  test('textInputs: the local viewer inputs over a text view', async () => {
    const { host, docKey, saved } = await projectedOf(await textGrid())
    const clock: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
    const snap = await host.sheetSnapshot(docKey, 0, clock)
    const workbook: XlsxDocument = { ...saved, sheets: saved.sheets.map((s) => servableSheet(s, 'visible-only').sheet) }
    const view = await sheetTextView(snap, saved.sheets[0]!, NO_SESSION_SIZES, { workbook, clock, budget: { reserve: () => {}, release: () => {} } })
    try {
      const source = createObjectWorkbookSource(textDocument(workbook, new Map([[0, view]])), 'tg')
      const t = textInputs(view, source)
      expect(t.sheetId).toBe(source.sheets[0]!.sheetId)
      // Authored hidden row 7 (display 6) and column F are hidden; their neighbours are not.
      expect(t.visibility.visible(TEXT_GRID.hiddenRow)).toBe(false)
      expect(t.visibility.visible(TEXT_GRID.hiddenRow - 1)).toBe(true)
      expect(t.columns.visible(TEXT_GRID.hiddenCol)).toBe(false)
      expect(t.columns.visible(TEXT_GRID.hiddenCol - 1)).toBe(true)
      expect(t.order.sourceToDisplay(12)).toBe(12)
      // Display text is the view's: B2 is `needle`, B1 formats under its number format, the covered D20 is blank.
      expect(t.textOf({ row: 1, col: 1, value: 'needle' })).toBe('needle')
      expect(t.textOf({ row: 0, col: 1, value: 1.5 })).toBe('1.5')
      const [cr, cc] = TEXT_GRID.covered
      expect(t.store().stamp.sheetId).toBe(t.sheetId)
      expect(view.sheet.rows.flatMap((r) => r.cells).find((x) => x.row === cr && x.col === cc)?.value ?? null).toBeNull()
    } finally { view.release() }
  }, 30_000)

  test('DocumentHost.sheetText caches 16 text views per document and releases them', async () => {
    const { host, docKey } = await projectedOf(await cfHiddenFormat())
    const doc = (host as unknown as { docs: Map<string, { textViews?: Map<string, unknown>; textBudget?: unknown }> }).docs.get(docKey)!
    // A tracking budget in place of the host's: every CF text scope a view holds reserves from it.
    let held = 0
    doc.textBudget = { reserve: (b: number) => { held += b }, release: (b: number) => { held -= b } }
    const clock: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
    const fitJob = (px: number) => ({ op: 'fit' as const, sheetIndex: 0, axis: 'col' as const, index: 1, pane: PANE, sizes: px ? [{ axis: 'col' as const, index: 3, px }] : [] })
    const first = await host.sheetText(docKey, clock, fitJob(0))
    expect(first).toEqual({ op: 'fit', size: expect.any(Number) })
    const one = held
    expect(one).toBeGreaterThan(0)
    // The same key is a hit (no new scope).
    expect(await host.sheetText(docKey, clock, fitJob(0))).toEqual(first)
    expect(held).toBe(one)
    for (let px = 20; px < 20 + 16; px++) await host.sheetText(docKey, clock, fitJob(px))
    await Bun.sleep(0)
    // 17 keys: the least recently used (the identity view) was evicted and released.
    expect(doc.textViews!.size).toBe(16)
    expect(held).toBe(16 * one)
    // The other ops are Tasks 6–8's: a clear error, nothing cached for them.
    expect(await rejectionMessage(host.sheetText(docKey, clock, { op: 'a11y', sheetIndex: 0, rows: [[0, 0]], cols: [0, 0], sizes: [] }))).toContain("op 'a11y' is not implemented")
    host.drop(docKey)
    await Bun.sleep(0)
    expect(held).toBe(0)
    expect(doc.textViews!.size).toBe(0)
  }, 30_000)
})
