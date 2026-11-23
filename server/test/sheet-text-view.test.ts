/**
 * Stage 4b Task 1 — `sheetTextView` (plan P2–P5): the one sheet every server text answer (find, copy, a11y, fit) reads.
 * Covered (base and current view), merge-tail and undrawn cells are blank in it; an unknown-coverage sheet is blank
 * throughout; under `'full'` it is the projected sheet itself. Tiles keep reading the snapshot, which is never changed.
 */
import { afterAll, describe, expect, spyOn, test } from 'bun:test'
import { DocumentHost } from '../src/render'
import { clockEpochMs, type SheetClock } from '../src/sheet/clock'
import { servableSheet, type SheetSnapshot } from '../src/sheet/snapshot'
import { sheetTextView, textDocument, type TextView } from '../src/sheet/text-view'
import * as cfText from '../../src/xlsx/cf/text'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { WorkbookFind } from '../../src/xlsx/view/find'
import { serializeSelection } from '../../src/xlsx/view/copy'
import { formatCellPresentation } from '../../src/components/xlsx/formatPresentation'
import { autoFitSize } from '../../src/components/xlsx/sheetResize'
import { cellDisplayText } from '../../src/xlsx/render'
import { initialWorkbookOwner } from '../../src/worker/viewerProtocol'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { MAX_COLS } from '../../src/xlsx/view/coordinates'
import { NO_SESSION_SIZES, type SessionSizes } from '../../src/xlsx/view/axisSizes'
import type { AuthoritativeCell } from '../../src/xlsx/view/contracts'
import type { FormatAllocation } from '../../src/xlsx/formula/format'
import type { EffectiveFormatScope, FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { drain, findInputOf } from '../../tests/helpers/xlsx-find'
import {
  absoluteOverB2, cfHiddenFormat, cfHiddenValue, cfLongColumn, CF_HIDDEN_FORMAT_VALUES, CF_HIDDEN_VALUES, CF_LONG_ROWS, hiddenFormat, mergeTail,
  offViewportCovered, rotatedBox, slidUnderAnchor, textGrid, TEXT_GRID, unfilledShapeOverText, unresolvedAnchor,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
const views: TextView[] = []
afterAll(() => {
  for (const v of views) v.release()
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
})

const UTC_23: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
/** UTC+14: its 12:00 on 2026-11-23 is not UTC's, so a clock that is not threaded is caught. */
const KIRITIMATI_23: SheetClock = { day: '2026-11-23', timeZone: 'Pacific/Kiritimati' }
type Policy = 'visible-only' | 'full'
const cols = (entries: Array<[number, number]>): SessionSizes => ({ rows: new Map(), cols: new Map(entries) })

/** A budget that tracks what is held, so a released view is seen to return every byte. */
function trackingBudget(): FormatBudget & { held: () => number } {
  let held = 0
  return { reserve: (b) => { held += b }, release: (b) => { held -= b }, held: () => held }
}

interface Opened { host: DocumentHost; docKey: string; saved: XlsxDocument; workbook: XlsxDocument; policy: Policy }
async function open(bytes: Uint8Array, policy: Policy = 'visible-only'): Promise<Opened> {
  const host = new DocumentHost()
  hosts.push(host)
  const o = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  const saved = (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
  // The servable workbook (every slot its blanked sheet), as the host hands CF's cross-sheet reads.
  const workbook: XlsxDocument = { ...saved, sheets: saved.sheets.map((s) => servableSheet(s, policy).sheet) }
  return { host, docKey: o.docKey, saved, workbook, policy }
}
async function viewIn(d: Opened, o: { sheetIndex?: number; sizes?: SessionSizes; clock?: SheetClock; budget?: FormatBudget; snap?: (s: SheetSnapshot) => SheetSnapshot } = {}): Promise<{ view: TextView; snap: SheetSnapshot; projected: XlsxSheet }> {
  const sheetIndex = o.sheetIndex ?? 0, clock = o.clock ?? UTC_23
  const snap = await d.host.sheetSnapshot(d.docKey, sheetIndex, clock)
  const projected = d.saved.sheets[sheetIndex]!
  const view = await sheetTextView(o.snap ? o.snap(snap) : snap, projected, o.sizes ?? NO_SESSION_SIZES, { workbook: d.workbook, clock, budget: o.budget ?? trackingBudget() })
  views.push(view)
  return { view, snap, projected }
}
async function viewOf(bytes: Uint8Array, policy: Policy = 'visible-only', o: Parameters<typeof viewIn>[1] = {}) {
  return viewIn(await open(bytes, policy), o)
}
const cellAt = (sheet: XlsxSheet, row: number, col: number): XlsxCell | undefined => sheet.rows.flatMap((r) => r.cells).find((c) => c.row === row && c.col === col)
const valueAt = (sheet: XlsxSheet, row: number, col: number) => cellAt(sheet, row, col)?.value ?? null
const allCells = (sheet: XlsxSheet) => sheet.rows.flatMap((r) => r.cells)
/** Display text as the local viewer forms it: the cell's base style through the scope's CF-effective style. */
const textOfView = (sheet: XlsxSheet) => (cell: AuthoritativeCell, _sheetId: string, allocation?: FormatAllocation, scope?: EffectiveFormatScope): string => {
  const base = cellAt(sheet, cell.row, cell.col)?.style
  return cellDisplayText(sheet, { value: cell.value, style: scope?.effectiveStyle(cell.row, cell.col, base) ?? base }, allocation)
}

describe('sheetTextView (S4 §6.2, plan P2–P5)', () => {
  test('engines import under Bun', async () => {
    const d = await open(await textGrid())
    const budget = trackingBudget()
    const byIndex = new Map<number, TextView>()
    for (let i = 0; i < TEXT_GRID.sheets.length; i++) byIndex.set(i, (await viewIn(d, { sheetIndex: i, budget })).view)
    const doc = textDocument(d.workbook, byIndex)
    expect(doc.sheets.map((s) => s.name)).toEqual([...TEXT_GRID.sheets])
    for (const [i, v] of byIndex) expect(doc.sheets[i]).toBe(v.sheet)

    // WorkbookFind over the object source of the text document: real hits, none in the covered D20.
    const { input } = findInputOf(doc, { docId: 'tg', source: createObjectWorkbookSource(doc, 'tg') })
    const sheetIdOf = new Map(input.sheets.map((s, i) => [s.sheetId, i]))
    const find = new WorkbookFind(input, { text: 'needle', caseSensitive: false }, {
      textOf: (cell, sheetId, allocation, scope) => textOfView(doc.sheets[sheetIdOf.get(sheetId)!]!)(cell, sheetId, allocation, scope),
      formatBudget: budget,
    })
    const hits = await drain(find)
    find.close()
    expect(hits.length).toBeGreaterThan(0)
    expect(hits.some((h) => sheetIdOf.get(h.sheetId) === 2 && h.source.row === 1 && h.source.col === 1)).toBe(true)
    const [cr, cc] = TEXT_GRID.covered
    expect(hits.some((h) => sheetIdOf.get(h.sheetId)! < 2 && h.source.row === cr && h.source.col === cc)).toBe(false)

    // serializeSelection over the `Plain` sheet: a non-empty TSV holding its needle.
    const source = createObjectWorkbookSource(doc, 'tg-copy')
    const owner = initialWorkbookOwner('tg-copy', source.sheets)
    const plainId = source.sheets[2]!.sheetId
    const store = source.store(plainId)
    const entry = owner.sheetVector.find((e) => e.sheetId === plainId)!
    const plain = doc.sheets[2]!
    const copied = await serializeSelection({
      store,
      owner: { ...store.stamp, fontGen: 0, viewGen: entry.revision },
      currentOwner: () => ({ ...store.stamp, fontGen: 0, viewGen: entry.revision }),
      selection: { active: null, anchor: null, ranges: [{ rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 4 }] }, cols: [{ c0: 0, c1: 2 }] }] },
      order: createIdentityOrder(),
      visibility: visibilityFromHidden(createIntervalSet([])),
      columns: visibilityFromHidden(createIntervalSet([]), MAX_COLS),
      allowed: () => true,
      signal: new AbortController().signal,
      format: (cell, allocation, scope) => textOfView(plain)(cell, plainId, allocation, scope),
    })
    expect(copied.kind).toBe('ready')
    const tsv = copied.kind === 'ready' ? copied.text : ''
    expect(tsv.length).toBeGreaterThan(0)
    expect(tsv).toContain(TEXT_GRID.plainNeedle[2])

    // prepareCfTextScope runs on the grid (no CF: absence) and on a CF sheet (a real scope).
    const grid = byIndex.get(0)!
    expect(await cfText.prepareCfTextScope(grid.sheet, { rects: [{ r0: 0, r1: 29, c0: 0, c1: 7 }], signal: new AbortController().signal, budget })).toBeUndefined()
    const cf = await viewOf(await cfHiddenFormat())
    expect(cf.view.scope).toBeDefined()
    expect(cf.view.scope!.effectiveStyle(1, 1, undefined)?.formatCode).toBe(';;;')

    // formatCellPresentation: the a11y text of a formatted number and of the needle.
    expect(formatCellPresentation(grid.sheet, cellAt(grid.sheet, 0, 1)!, budget).text).toBe('1.5')
    expect(formatCellPresentation(grid.sheet, cellAt(grid.sheet, 1, 1)!, budget).text).toBe('needle')

    // autoFitSize (reachable under Bun from sheetResize.tsx): a finite column width over the view's geometry.
    if (grid.geometry.kind !== 'ready') throw new Error('geometry not ready')
    const size = autoFitSize('column', 0, { geometry: grid.geometry, sheet: grid.sheet, pane: { x: 0, y: 0, width: 800, height: 600 }, knownRows: () => [], measure: (_font, text) => text.length * 7 })
    expect(Number.isFinite(size)).toBe(true)
    expect(size).toBeGreaterThan(0)
  }, 30_000)

  test('covered cells are blank in the view', async () => {
    const shape = (await viewOf(await unfilledShapeOverText())).view
    expect(valueAt(shape.sheet, 1, 1)).toBeNull()
    expect(valueAt(shape.sheet, 1, 0)).toBe('CELL-2-1')
    expect(shape.withheld).toBe(1)

    // F9: the rotated box spans x 40..60, y -40..60: A1, A2 and A3 are blank; B1:C6 and A4:A6 keep their text.
    const rot = (await viewOf(await rotatedBox())).view
    for (let r = 0; r < 6; r++) for (let c = 0; c < 3; c++) {
      if (c === 0 && r < 3) expect([r, c, valueAt(rot.sheet, r, c)]).toEqual([r, c, null])
      else expect([r, c, valueAt(rot.sheet, r, c)]).toEqual([r, c, `ROT-${r + 1}-${c + 1}`])
    }

    const off = (await viewOf(await offViewportCovered())).view
    expect(valueAt(off.sheet, 4999, 25)).toBeNull()
    expect(valueAt(off.sheet, 0, 0)).toBe('OFFVIEW-A1')
  }, 30_000)

  test('unknown coverage blanks every cell', async () => {
    const { view, snap, projected } = await viewOf(await unresolvedAnchor())
    expect(snap.coverage).toEqual({ kind: 'unknown', reason: 'unresolved-anchors' })
    expect(view.blankAll).toBe(true)
    expect(view.scope).toBeUndefined()
    expect(allCells(view.sheet).length).toBe(allCells(projected).length)
    for (const cell of allCells(view.sheet)) {
      expect(cell.value).toBeNull()
      expect(cell.formula).toBeUndefined()
      const authored = cellAt(projected, cell.row, cell.col)!
      expect(cell.styleIndex).toBe(authored.styleIndex)
      expect(cell.style).toEqual(authored.style!)
    }
    expect(view.withheld).toBe(4)
    // Tiles paint as authored: the snapshot's own sheet still holds the values.
    expect(valueAt(snap.sheet, 0, 0)).toBe('UNRES-1-1')
  }, 30_000)

  test('current-view coverage', async () => {
    const d = await open(await slidUnderAnchor())
    const authored = await viewIn(d)
    expect(valueAt(authored.view.sheet, 0, 4)).toBe('SLID-E1')
    const shrunk = await viewIn(d, { sizes: cols([[1, 8]]) })
    expect(valueAt(shrunk.view.sheet, 0, 4)).toBeNull()
    expect(shrunk.view.withheld).toBe(1)
    // The view's geometry is built with the session sizes.
    if (shrunk.view.geometry.kind !== 'ready') throw new Error('geometry not ready')
    expect(shrunk.view.geometry.metrics.cols.size(1)).toBe(8)
    // Tiles stay base-blanked only: the snapshot's tile sheet still holds E1.
    expect(valueAt(shrunk.snap.sheet, 0, 4)).toBe('SLID-E1')
  }, 30_000)

  test('merge tails and undrawn cells are blank', async () => {
    const merge = (await viewOf(await mergeTail())).view
    expect([valueAt(merge.sheet, 0, 0), valueAt(merge.sheet, 0, 1)]).toEqual(['MERGE-A1', null])

    const fmt = (await viewOf(await hiddenFormat())).view
    expect([valueAt(fmt.sheet, 1, 1), valueAt(fmt.sheet, 0, 0)]).toEqual([null, 'SHOWN-A1'])

    // P4 precision: the showValue="0" bars blank B1:B3; the text cell under the same rule draws no bar and keeps its text.
    const hv = await open(await cfHiddenValue())
    const bars = (await viewIn(hv)).view
    expect([0, 1, 2, 3].map((r) => valueAt(bars.sheet, r, 1))).toEqual([null, null, null, 'CFTEXT-B4'])
    // The CF batch unavailable: every cell of the showValue="0" rule's sqref is blank.
    const noBatch = (await viewIn(hv, { snap: (s) => ({ ...s, assets: { ...s.assets, conditionalFormatting: undefined } }) })).view
    expect([0, 1, 2, 3].map((r) => valueAt(noBatch.sheet, r, 1))).toEqual([null, null, null, null])

    // A CF rule writing `;;;`: B2 and B3 paint nothing, B1 paints 4.
    const hf = await open(await cfHiddenFormat())
    const cf = (await viewIn(hf)).view
    expect([0, 1, 2].map((r) => valueAt(cf.sheet, r, 1))).toEqual([CF_HIDDEN_FORMAT_VALUES[0]!, null, null])
    // The CF scope unavailable (absent or refused): every cell of every CF sqref is blank.
    for (const stub of [() => Promise.resolve(undefined), () => Promise.reject(new Error('refused'))]) {
      const spy = spyOn(cfText, 'prepareCfTextScope').mockImplementation(stub)
      try {
        const blind = (await viewIn(hf)).view
        expect(spy).toHaveBeenCalled()
        expect(blind.scope).toBeUndefined()
        expect([0, 1, 2].map((r) => valueAt(blind.sheet, r, 1))).toEqual([null, null, null])
      } finally { spy.mockRestore() }
    }

    // The scope's clock is the session's: the same instant and zone as the tile's cfClock.
    const seen: Array<Parameters<typeof cfText.prepareCfTextScope>[1]> = []
    const real = cfText.prepareCfTextScope
    const spy = spyOn(cfText, 'prepareCfTextScope').mockImplementation((sheet, request) => { seen.push(request); return real(sheet, request) })
    try {
      const clocked = (await viewIn(hf, { clock: KIRITIMATI_23 })).view
      expect(clocked.scope).toBeDefined()
    } finally { spy.mockRestore() }
    expect(seen.length).toBe(1)
    expect(seen[0]!.dateContext?.epochNowMs).toBe(clockEpochMs(KIRITIMATI_23))
    expect(seen[0]!.dateContext?.timeZone).toBe(KIRITIMATI_23.timeZone)
    expect(seen[0]!.now?.()).toBe(clockEpochMs(KIRITIMATI_23))
    expect(seen[0]!.evaluation?.timeZone).toBe(KIRITIMATI_23.timeZone)
  }, 30_000)

  test('a release returns every scope byte', async () => {
    const budget = trackingBudget()
    const { view } = await viewOf(await cfHiddenFormat(), 'visible-only', { budget })
    expect(view.scope).toBeDefined()
    expect(budget.held()).toBeGreaterThan(0)
    view.release()
    view.release()
    expect(budget.held()).toBe(0)
  }, 30_000)

  test('CF over a long column (F3): the scope serves every row, and rank rules keep their whole-range rank', async () => {
    // A fill-only `cellIs > 0` rule over A1:A10000: no cell is undrawn, so every value keeps its text.
    const fill = (await viewOf(await cfLongColumn('fill'))).view
    expect(fill.scope).toBeDefined()
    expect(fill.withheld).toBe(0)
    for (const r of [0, 4999, 5000, CF_LONG_ROWS - 1]) expect(valueAt(fill.sheet, r, 0)).toBe(r + 1)
    // A `top10 rank=1` rule writing `;;;`: only the column's maximum (A10000) is undrawn. Preparing the scope in
    // 5,000-cell slices never re-ranks per slice (A5000, the first slice's maximum, keeps its text).
    const top = (await viewOf(await cfLongColumn('top1-hidden'))).view
    expect(top.scope).toBeDefined()
    expect(top.withheld).toBe(1)
    expect(valueAt(top.sheet, CF_LONG_ROWS - 1, 0)).toBeNull()
    for (const r of [0, 4999, 5000, CF_LONG_ROWS - 2]) expect(valueAt(top.sheet, r, 0)).toBe(r + 1)
  }, 30_000)

  test('covered stays covered in answers', async () => {
    const { view, snap } = await viewOf(await absoluteOverB2(), 'visible-only', { sizes: cols([[0, 300]]) })
    expect(valueAt(view.sheet, 1, 1)).toBeNull()
    expect(valueAt(snap.sheet, 1, 1)).toBeNull()
  }, 30_000)

  test('full is the identity', async () => {
    const fixtures: Array<[string, () => Promise<Uint8Array>, SessionSizes, Array<[number, number, XlsxCell['value']]>]> = [
      ['unfilledShapeOverText', unfilledShapeOverText, NO_SESSION_SIZES, [[1, 1, 'COVERED-B2']]],
      ['rotatedBox', rotatedBox, NO_SESSION_SIZES, [[0, 0, 'ROT-1-1']]],
      ['offViewportCovered', offViewportCovered, NO_SESSION_SIZES, [[4999, 25, 'OFFVIEW-Z5000']]],
      ['unresolvedAnchor', unresolvedAnchor, NO_SESSION_SIZES, [[0, 0, 'UNRES-1-1']]],
      ['slidUnderAnchor', slidUnderAnchor, cols([[1, 8]]), [[0, 4, 'SLID-E1']]],
      ['mergeTail', mergeTail, NO_SESSION_SIZES, [[0, 1, 'TAIL-B1']]],
      ['hiddenFormat', hiddenFormat, NO_SESSION_SIZES, [[1, 1, 'HIDDENFMT']]],
      ['cfHiddenValue', cfHiddenValue, NO_SESSION_SIZES, CF_HIDDEN_VALUES.map((v, r) => [r, 1, v] as [number, number, number])],
      ['cfHiddenFormat', cfHiddenFormat, NO_SESSION_SIZES, [[1, 1, CF_HIDDEN_FORMAT_VALUES[1]!], [2, 1, CF_HIDDEN_FORMAT_VALUES[2]!]]],
      ['absoluteOverB2', absoluteOverB2, cols([[0, 300]]), [[1, 1, 'ABS-B2']]],
    ]
    for (const [name, make, sizes, present] of fixtures) {
      const { view, projected } = await viewOf(await make(), 'full', { sizes })
      expect([name, view.sheet === projected, view.withheld, view.blankAll]).toEqual([name, true, 0, false])
      // Not vacuous: every value the visible-only view withholds is present.
      for (const [r, c, v] of present) expect([name, r, c, valueAt(view.sheet, r, c)]).toEqual([name, r, c, v])
    }
  }, 30_000)
})
