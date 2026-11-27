/**
 * Stage 4c Task 1 — `sheetKeyView` (plan P5): the one sheet every key answer (sort keys, filter evaluation, value
 * lists, range detection) reads. It is the text view at the identity view (`NO_SESSION_SIZES`) plus, under
 * `visible-only`, painted CF-effective styles, blank zero-width columns and masked strings read as their painted text.
 * Under `'full'` it is the projected sheet itself. These tests also prove the local engines (`viewerControls` over an
 * object workbook source) run under Bun.
 */
import { afterAll, describe, expect, spyOn, test } from 'bun:test'
import { DocumentHost } from '../src/render'
import { clockEpochMs, type SheetClock } from '../src/sheet/clock'
import { keyDocument, sheetKeyView, type KeyView } from '../src/sheet/key-view'
import { servableSheet, type SheetSnapshot } from '../src/sheet/snapshot'
import * as textView from '../src/sheet/text-view'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import { cellDisplayText } from '../../src/xlsx/render'
import { NO_SESSION_SIZES } from '../../src/xlsx/view/axisSizes'
import type { CacheIdentity, FilterRange, FilterSpec, SheetView, TypedCondition } from '../../src/xlsx/view/contracts'
import type { FilterContext } from '../../src/xlsx/view/filter'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { createIntervalSet, visibilityFromHidden } from '../../src/xlsx/view/visibility'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import {
  filterAutoHidden, hiddenPivotCache, oneMillionSort, paintedFormats, PAINTED_FORMATS, SORT_GRID, SORT_WITHHELD, sortGrid, sortGridScore, sortMerges,
  sortUnderShape, sortWithheld, unresolvedAnchor,
} from './sheet-fixtures'

const hosts: DocumentHost[] = []
const keyViews: KeyView[] = []
afterAll(() => {
  for (const v of keyViews) v.release()
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
})

const UTC_23: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
type Policy = 'visible-only' | 'full'

function trackingBudget(): FormatBudget & { held: () => number } {
  let held = 0
  return { reserve: (b) => { held += b }, release: (b) => { held -= b }, held: () => held }
}

interface Opened { host: DocumentHost; docKey: string; saved: XlsxDocument; workbook: XlsxDocument }
async function open(bytes: Uint8Array, policy: Policy = 'visible-only'): Promise<Opened> {
  const host = new DocumentHost()
  hosts.push(host)
  const o = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  const saved = (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
  const workbook: XlsxDocument = { ...saved, sheets: saved.sheets.map((s) => servableSheet(s, policy).sheet) }
  return { host, docKey: o.docKey, saved, workbook }
}
async function keyIn(d: Opened, o: { sheetIndex?: number; clock?: SheetClock; budget?: FormatBudget } = {}): Promise<{ view: KeyView; snap: SheetSnapshot; projected: XlsxSheet }> {
  const sheetIndex = o.sheetIndex ?? 0, clock = o.clock ?? UTC_23
  const snap = await d.host.sheetSnapshot(d.docKey, sheetIndex, clock)
  const projected = d.saved.sheets[sheetIndex]!
  const view = await sheetKeyView(snap, projected, { workbook: d.workbook, clock, budget: o.budget ?? trackingBudget() })
  keyViews.push(view)
  return { view, snap, projected }
}
const cellAt = (sheet: XlsxSheet, row: number, col: number): XlsxCell | undefined => sheet.rows.flatMap((r) => r.cells).find((c) => c.row === row && c.col === col)
const valueAt = (sheet: XlsxSheet, row: number, col: number) => cellAt(sheet, row, col)?.value ?? null
const allCells = (sheet: XlsxSheet) => sheet.rows.flatMap((r) => r.cells)
/** 1-based A1 ref → 0-based (row, col). */
const rc = (ref: string): [number, number] => [Number(ref.slice(1)) - 1, ref.charCodeAt(0) - 65]
const signal = () => new AbortController().signal

/** `viewerControls` over the object source of `doc`, sheet `sheetIndex`, with a publishing view as the viewer keeps. */
async function controlsOver(doc: XlsxDocument, sheetIndex: number, o: { context?: () => FilterContext } = {}) {
  const source = createObjectWorkbookSource(doc, 'key-doc')
  const sheetId = source.sheets[sheetIndex]!.sheetId
  let view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden: createIntervalSet([]), visibility: visibilityFromHidden(createIntervalSet([])), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  let filter: FilterSpec | null = null
  const owner = (): CacheIdentity => ({ docId: 'key-doc', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter, sort: null, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => {
    view = { ...view, ...next }
    filter = next.filter
    viewGen++
    return snapshot()
  }
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish, ...(o.context ? { context: o.context } : {}) }, signal())
  return {
    controls, owner, view: () => view,
    /** Source rows of the body of `range`, in display order. */
    order: (range: FilterRange) => Array.from({ length: range.r1 - range.totalsRows - range.r0 - range.headerRows + 1 }, (_, i) => view.order.displayToSource(range.r0 + range.headerRows + i)),
    /** Source rows of the body of `range`, in display order, that the view shows. */
    shown: (range: FilterRange) => {
      const out: number[] = []
      for (let d = range.r0 + range.headerRows; d <= range.r1 - range.totalsRows; d++) {
        const row = view.order.displayToSource(d)
        if (view.visibility.visible(d)) out.push(row)
      }
      return out
    },
    /** The labels of the first values page of `col` over `range`. */
    labels: async (range: FilterRange, col: number) => {
      const port = await controls.client.values(owner(), range, col, null, signal())
      try { return (await port.read({ text: '', offset: 0 }, signal())).entries.map((e) => e.label) } finally { port.cancel() }
    },
    async close() { controls.dispose(); await source.release() },
  }
}

const GRID_RANGE: FilterRange = { r0: 0, r1: SORT_GRID.rows - 1, c0: 0, c1: SORT_GRID.cols - 1, headerRows: 1, totalsRows: 0 }
const contains = (first: string): TypedCondition => ({ kind: 'text', op: 'contains', first, wildcards: false } as TypedCondition)

describe('sheetKeyView (S4 §6.2, §6.12 item 7; plan P5)', () => {
  test('engines import under Bun', async () => {
    const d = await open(await sortGrid())
    const { view, projected } = await keyIn(d)
    const doc = keyDocument(d.workbook, new Map([[0, view]]))
    expect(doc.sheets[0]).toBe(view.sheet)
    const h = await controlsOver(doc, 0)
    try {
      // range-detect reads the key view's occupancy: zero-width column E and hidden row 9 hold blank keys, so the
      // region around B1 is A1:D8 with its header row.
      const detected = await h.controls.client.resolve(h.owner(), 'filter', signal(), undefined, 1)
      expect(detected).toMatchObject({ kind: 'complete', range: { r0: 0, r1: SORT_GRID.hiddenRow - 1, c0: 0, c1: SORT_GRID.zeroWidthCol - 1, headerRows: 1 } })

      // Sort by B ascending: the order computed from the projected sheet, the withheld rows (hidden row 9, `ht="0"` row 12)
      // blank keys, which sort last in source order.
      const sorted = await h.controls.client.sort(h.owner(), { range: GRID_RANGE, keys: [{ col: 1, order: 'asc' }], caseSensitive: false }, signal())
      expect(sorted.kind).toBe('committed')
      const body = Array.from({ length: SORT_GRID.rows - 1 }, (_, i) => i + 1)
      const withheld = new Set<number>([SORT_GRID.hiddenRow, SORT_GRID.zeroHeightRow])
      const keyed = body.filter((r) => !withheld.has(r))
      for (const r of keyed) expect(valueAt(projected, r, 1)).toBe(sortGridScore(r + 1))
      const expected = [...keyed.sort((a, b) => sortGridScore(a + 1) - sortGridScore(b + 1)), ...body.filter((r) => withheld.has(r))]
      expect(h.order(GRID_RANGE)).toEqual(expected)

      // Filter D contains "a": exactly the drawn rows whose word holds an "a".
      const filtered = await h.controls.client.filter(h.owner(), { range: GRID_RANGE, columns: new Map([[3, contains('a')]]) }, signal())
      expect(filtered.kind).toBe('committed')
      const words = keyed.filter((r) => SORT_GRID.words[(r + 1) % 8]!.includes('a'))
      expect(new Set(h.shown(GRID_RANGE))).toEqual(new Set(words))

      // One values page: column A's labels, the drawn names only.
      const labels = await h.labels(GRID_RANGE, 0)
      expect(new Set(labels)).toEqual(new Set(keyed.map((r) => `N${r + 1}`)))
    } finally { await h.close() }
  }, 30_000)

  test('withheld cells are blank keys', async () => {
    const d = await open(await sortWithheld())
    const { view, projected } = await keyIn(d)
    const grid = await keyIn(await open(await sortGrid()))
    expect(view.blankAll).toBe(false)
    const blank = new Set<string>()
    for (const ref of [SORT_WITHHELD.covered, SORT_WITHHELD.hiddenFormat, SORT_WITHHELD.cfHidden, 'C30', 'C31', 'C32', 'C33', 'C34']) {
      const [r, c] = rc(ref)
      expect(valueAt(projected, r, c)).not.toBeNull()
      expect(valueAt(view.sheet, r, c)).toBeNull()
      blank.add(`${r}:${c}`)
    }
    for (let r = 0; r < SORT_GRID.rows; r++) {
      expect(valueAt(view.sheet, r, SORT_GRID.zeroWidthCol)).toBeNull()
      blank.add(`${r}:${SORT_GRID.zeroWidthCol}`)
    }
    for (let c = 0; c < SORT_GRID.cols; c++) {
      expect(valueAt(view.sheet, SORT_GRID.zeroHeightRow, c)).toBeNull()
      blank.add(`${SORT_GRID.zeroHeightRow}:${c}`)
    }
    expect(view.collapsedCols).toEqual([{ start: SORT_GRID.zeroWidthCol, end: SORT_GRID.zeroWidthCol }])

    // D23: read as its painted text, with no number format.
    const [mr, mc] = rc(SORT_WITHHELD.masked)
    const masked = cellAt(view.sheet, mr, mc)!
    expect(valueAt(projected, mr, mc)).toBe(SORT_WITHHELD.maskedText)
    expect(masked.value).toBe(SORT_WITHHELD.maskedDisplay)
    expect(masked.style?.numFmtId ?? 0).toBe(0)
    expect(masked.style?.formatCode).toBeUndefined()
    blank.add(`${mr}:${mc}`)

    // Every other cell holds its projected value, which is sortGrid's.
    for (const cell of allCells(projected)) {
      if (blank.has(`${cell.row}:${cell.col}`)) continue
      expect(valueAt(view.sheet, cell.row, cell.col)).toEqual(cell.value)
      expect(valueAt(grid.projected, cell.row, cell.col)).toEqual(cell.value)
    }
    expect(view.withheld).toBe(allCells(projected).filter((c) => c.value !== null).length - allCells(view.sheet).filter((c) => c.value !== null).length)
  }, 30_000)

  test('painted styles', async () => {
    const d = await open(await paintedFormats())
    const budget = trackingBudget()
    const { view, snap } = await keyIn(d, { budget })
    // The text view's scope is released once the key view is built: its styles are resolved.
    expect(budget.held()).toBe(0)
    const b2 = cellAt(view.sheet, 1, 1)!, c2 = cellAt(view.sheet, 1, 2)!
    expect(cellDisplayText(view.sheet, b2)).toBe(PAINTED_FORMATS.tablePaints)
    expect(cellDisplayText(view.sheet, c2)).toBe(PAINTED_FORMATS.cfPaints)
    // Fresh style indexes: no index means two different styles in the key view.
    const byIndex = new Map<number, string>()
    for (const cell of allCells(view.sheet)) {
      const key = JSON.stringify(cell.style ?? null)
      expect(byIndex.get(cell.styleIndex) ?? key).toBe(key)
      byIndex.set(cell.styleIndex, key)
    }

    const doc = keyDocument(d.workbook, new Map([[0, view]]))
    const h = await controlsOver(doc, 0)
    try {
      const b = await h.labels({ r0: 1, r1: 3, c0: 1, c1: 1, headerRows: 0, totalsRows: 0 }, 1)
      expect(b).toEqual([PAINTED_FORMATS.tablePaints])
      expect(b).not.toContain(String(PAINTED_FORMATS.tableValue))
      expect(await h.labels({ r0: 1, r1: 1, c0: 2, c1: 2, headerRows: 0, totalsRows: 0 }, 2)).toEqual([PAINTED_FORMATS.cfPaints])
    } finally { await h.close() }

    // Column D's and row 6's default styles (`;;;`): the parser keeps no column style and only a row's `styled` flag,
    // so a cell with no own `s` paints under style 0 and the tile shows its text. The key and text views agree with
    // the tile: those cells keep their values.
    const text = await textView.sheetTextView(snap, d.saved.sheets[0]!, NO_SESSION_SIZES, { workbook: d.workbook, clock: UTC_23, budget: trackingBudget() })
    try {
      for (const [r, c] of [[1, 3], [2, 3], [3, 3], [5, 0], [5, 2]] as const) {
        const tile = cellAt(snap.sheet, r, c)!
        const painted = cellDisplayText(snap.sheet, tile)
        expect(painted).toBe(String(tile.value))
        expect(valueAt(text.sheet, r, c)).toBe(tile.value)
        expect(valueAt(view.sheet, r, c)).toBe(tile.value)
      }
    } finally { text.release() }
  }, 30_000)

  test('the key view ignores session sizes', async () => {
    const d = await open(await sortWithheld())
    const spy = spyOn(textView, 'sheetTextView')
    try {
      expect(sheetKeyView.length).toBe(3)
      const first = await keyIn(d)
      expect(spy).toHaveBeenCalledTimes(1)
      expect(spy.mock.calls[0]![2]).toBe(NO_SESSION_SIZES)
      // The build takes no sizes, so whatever sizes a caller holds, a second build is the same key view.
      const second = await keyIn(d)
      expect(spy.mock.calls[1]![2]).toBe(NO_SESSION_SIZES)
      expect(second.view.sheet).toEqual(first.view.sheet)
      expect(second.view.collapsedCols).toEqual(first.view.collapsedCols)
      expect(second.view.withheld).toBe(first.view.withheld)
    } finally { spy.mockRestore() }
  }, 30_000)

  test('unknown coverage', async () => {
    const { view } = await keyIn(await open(await unresolvedAnchor()))
    expect(view.blankAll).toBe(true)
    expect(allCells(view.sheet).every((c) => c.value === null && c.formula === undefined)).toBe(true)
    expect(view.withheld).toBe(4)
  }, 30_000)

  describe('full is the identity', () => {
    const fixtures: Array<[string, () => Promise<Uint8Array>]> = [
      ['sortGrid', sortGrid], ['sortWithheld', sortWithheld], ['sortUnderShape', sortUnderShape], ['sortMerges', sortMerges],
      ['oneMillionSort', oneMillionSort], ['hiddenPivotCache', hiddenPivotCache], ['filterAutoHidden', filterAutoHidden], ['paintedFormats', paintedFormats],
    ]
    for (const [name, build] of fixtures) {
      test(name, async () => {
        const d = await open(await build(), 'full')
        for (let i = 0; i < d.saved.sheets.length; i++) {
          const { view, projected } = await keyIn(d, { sheetIndex: i })
          expect(view.sheet).toBe(projected)
          expect(view.blankAll).toBe(false)
          expect(view.withheld).toBe(0)
        }
      }, 30_000)
    }
  })

  test('the injected clock', async () => {
    const d = await open(await sortGrid())
    const clock: SheetClock = { day: '2026-11-19', timeZone: 'UTC' }
    const { view } = await keyIn(d, { clock })
    const doc = keyDocument(d.workbook, new Map([[0, view]]))
    const context = (): FilterContext => ({ locale: 'en-US', dateSystem: '1900', timeZone: clock.timeZone, nowMs: clockEpochMs(clock) })
    // F{r} holds dateBase + r: r 5 is 2026-11-19 (the clock's day), r 11 is 2026-11-25 (the stubbed wall clock's).
    const today = 5 - 1, stubbedDay = 11 - 1
    const run = async () => {
      const h = await controlsOver(doc, 0, { context })
      try {
        const outcome = await h.controls.client.filter(h.owner(), { range: GRID_RANGE, columns: new Map([[5, { kind: 'datePreset', preset: 'today' }]]) }, signal())
        expect(outcome.kind).toBe('committed')
        return h.shown(GRID_RANGE)
      } finally { await h.close() }
    }
    expect(await run()).toEqual([today])
    const now = spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 10, 25, 12))
    try { expect(await run()).toEqual([today]) } finally { now.mockRestore() }
    // The control: without `context` the local default reads the (stubbed) wall clock.
    const stubbed = spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 10, 25, 12))
    try {
      const h = await controlsOver(doc, 0)
      try {
        await h.controls.client.filter(h.owner(), { range: GRID_RANGE, columns: new Map([[5, { kind: 'datePreset', preset: 'today' }]]) }, signal())
        expect(h.shown(GRID_RANGE)).toEqual([stubbedDay])
      } finally { await h.close() }
    } finally { stubbed.mockRestore() }
  }, 30_000)
})

describe('the full key view reads the snapshot', () => {
  test("'full' builds no text view: the projected sheet, with the collapsed columns of the snapshot's geometry", async () => {
    // sortWithheld carries a CF rule (B22 `cellIs > 90`), so a text view's CF scope would reserve format budget.
    const d = await open(await sortWithheld(), 'full')
    expect(d.saved.sheets[0]!.conditionalFormatting?.rules.length).toBeGreaterThan(0)
    const spy = spyOn(textView, 'sheetTextView')
    const budget = trackingBudget()
    let reserved = 0
    const counting: FormatBudget = { reserve: (b) => { reserved += b; budget.reserve(b) }, release: (b) => budget.release(b) }
    try {
      const { view, snap, projected } = await keyIn(d, { budget: counting })
      expect(spy).toHaveBeenCalledTimes(0)
      expect(reserved).toBe(0)
      expect(view.sheet).toBe(projected)
      expect(view.blankAll).toBe(false)
      expect(view.withheld).toBe(0)
      expect(view.collapsedCols).toEqual([{ start: SORT_GRID.zeroWidthCol, end: SORT_GRID.zeroWidthCol }])
      if (snap.geometry.kind !== 'ready') throw new Error('sortWithheld geometry not ready')
      expect(view.collapsedCols).toEqual(snap.geometry.metrics.cols.overrides.filter((s) => s.sizeCss === 0).map((s) => ({ start: s.start, end: s.end })))
    } finally { spy.mockRestore() }
  }, 30_000)
})
