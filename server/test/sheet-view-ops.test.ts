/**
 * Stage 4c Task 5 — the `sheet-view` pool job (plan P3, P5, P6, P13; parent §12 C1). Sort, filter, value lists and
 * range detection run through the local engine over the key view. The oracle in each test is `viewerControls` run
 * here over a copy of the fixture parsed apart from the server inputs, with its withheld cells physically deleted:
 * under `visible-only` that is B20, B21, B22, C30:C34, column E's cells, row 12's cells and hidden row 9's cells (the
 * projection strips a hidden row's cells), with D23 replaced by `****`.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, ColumnRule, FilterRange, FilterSpec, SheetView, SortSpec } from '../../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { visibilityFromHidden } from '../../src/xlsx/view/visibility'
import type { XlsxCell, XlsxDocument } from '../../src/xlsx/types'
import type { SheetOrder, WireRange, WireSortSpec } from '../../src/stream/sheet-wire'
import { Pool } from '../src/pool'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetViewJob, SheetViewReply } from '../src/sheet/view-ops'
import { orderBytes, type CanonicalView, type ResolvedFilter, type ResolvedRule } from '../src/sheet/view-spec'
import {
  DISPLAY_MERGE, displayMerge, FILTER_AUTO_HIDDEN, filterAutoHidden, ONE_MILLION_ROWS, oneMillionSort, PAINTED_FORMATS, paintedFormats, SORT_GRID,
  SORT_WITHHELD, sortGrid, sortMerges, sortUnderShape, sortWithheld,
} from './sheet-fixtures'

type Policy = 'visible-only' | 'full'
const CLOCK: SheetClock = { day: '2026-11-23', timeZone: 'UTC' }
const IDENTITY: CanonicalView = { sort: null, sortFilter: null, filter: null }
const signal = () => new AbortController().signal

const hosts: DocumentHost[] = []
const pools: Pool[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await Promise.all(pools.map((p) => p.dispose()))
})

/** The server side: an in-process `DocumentHost` holding the fixture in tiles mode. */
async function serve(bytes: Uint8Array, policy: Policy) {
  const host = new DocumentHost()
  hosts.push(host)
  const { docKey } = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  return (job: SheetViewJob) => host.sheetView(docKey, CLOCK, policy, job)
}
type Serve = Awaited<ReturnType<typeof serve>>

const viewOf = (r: SheetViewReply): { view: CanonicalView; order: SheetOrder } => {
  if (r.op !== 'view' || !('order' in r)) throw new Error(`no view: ${JSON.stringify(r)}`)
  return r
}
/** Source rows of `range`'s body in display order, under `order`. */
function displayed(order: SheetOrder, r0: number, r1: number): number[] {
  const out: number[] = []
  for (let d = r0; d <= r1; d++) {
    const band = order.bands.find((b) => b.r0 <= d && d <= b.r1)
    out.push(band ? band.sourceAtOffset[d - band.r0]! : d)
  }
  return out
}
/** Source rows of `[r0, r1]` the view shows: display slot through the bands, then the mask; a slot outside the mask
 * shows as in the identity view (`authoredHidden`, the sheet's hidden rows, stays hidden). */
function shownSources(order: SheetOrder, r0: number, r1: number, authoredHidden: ReadonlySet<number> = new Set()): Set<number> {
  const out = new Set<number>()
  const m = order.mask
  for (let d = r0; d <= r1; d++) {
    const src = displayed(order, d, d)[0]!
    if (!m || d < m.r0 || d > m.r1) { if (!authoredHidden.has(src)) out.add(src); continue }
    const i = d - m.r0
    if ((m.words[i >>> 5]! >>> (i & 31)) & 1) out.add(src)
  }
  return out
}

// ---- the oracle ----------------------------------------------------------------------------------------------

const rcOf = (ref: string): [number, number] => [Number(ref.slice(1)) - 1, ref.charCodeAt(0) - 65]
/** The fixture parsed here, with `drop(row, col)` cells physically deleted and `replace` cells rewritten. */
async function oracleDoc(bytes: Uint8Array, drop: (row: number, col: number) => boolean = () => false, replace: Record<string, string> = {}): Promise<XlsxDocument> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const sheet = doc.sheets[0]!
  const swap = new Map(Object.entries(replace).map(([ref, v]) => [rcOf(ref).join(':'), v]))
  const rows = sheet.rows.map((row) => ({ ...row, cells: row.cells.filter((c) => !drop(c.row, c.col)).map((c): XlsxCell => {
    const v = swap.get(`${c.row}:${c.col}`)
    return v === undefined ? c : { ref: c.ref, row: c.row, col: c.col, value: v, styleIndex: 0, style: { numFmtId: 0 } }
  }) }))
  return { ...doc, sheets: [{ ...sheet, rows }, ...doc.sheets.slice(1)] }
}
/** `sortWithheld` as the oracle sees it under `visible-only`. */
function withheldOracle(bytes: Uint8Array): Promise<XlsxDocument> {
  const gone = new Set([SORT_WITHHELD.covered, SORT_WITHHELD.hiddenFormat, SORT_WITHHELD.cfHidden, 'C30', 'C31', 'C32', 'C33', 'C34'].map((r) => rcOf(r).join(':')))
  return oracleDoc(bytes, (r, c) => gone.has(`${r}:${c}`) || c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow,
    { [SORT_WITHHELD.masked]: SORT_WITHHELD.maskedDisplay })
}
/** `sortGrid` (and its variants) as the oracle sees them under `visible-only`. */
const gridOracle = (bytes: Uint8Array) => oracleDoc(bytes, (r, c) => c === SORT_GRID.zeroWidthCol || r === SORT_GRID.zeroHeightRow || r === SORT_GRID.hiddenRow)

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
    view: () => view,
    async sort(spec: SortSpec) { return (await client.sort(owner(), spec, signal())).kind },
    async filter(spec: FilterSpec | null) { return (await client.filter(owner(), spec, signal())).kind },
    /** A values rule selecting the local entries whose label passes `pick`. */
    async values(range: FilterRange, col: number, pick: (label: string) => boolean, blanks: boolean): Promise<ColumnRule | null> {
      const port = await client.values(owner(), range, col, null, signal())
      const page = await port.read({ text: '', offset: 0 }, signal())
      const quiet = { text: '', offset: 10_000 }
      await port.change({ kind: 'all', selected: false, includeBlanks: false }, quiet, signal())
      for (const e of page.entries) if (pick(e.label)) await port.change({ kind: 'toggle', ordinal: e.ordinal, selected: true }, quiet, signal())
      await port.change({ kind: 'blanks', selected: blanks }, quiet, signal())
      const rule = await port.candidate(signal())
      port.transfer()
      return rule
    },
    async labels(range: FilterRange, col: number) {
      const port = await client.values(owner(), range, col, null, signal())
      try { return (await port.read({ text: '', offset: 0 }, signal())).entries.map((e) => e.label) } finally { port.cancel() }
    },
    async detect(kind: 'sort' | 'filter', active: [number, number]) {
      view = { ...view, selection: { ...emptyGridSelection(), active: { row: active[0], col: active[1] } } }
      return client.resolve(owner(), kind, signal())
    },
    /** Source rows of `[r0, r1]` the oracle shows. */
    shown(r0: number, r1: number) {
      const out = new Set<number>()
      for (let d = r0; d <= r1; d++) if (view.visibility.visible(d)) out.add(view.order.displayToSource(d))
      return out
    },
    order(r0: number, r1: number) { return Array.from({ length: r1 - r0 + 1 }, (_, i) => view.order.displayToSource(r0 + i)) },
    dispose: async () => { controls.dispose(); await source.release() },
  }
}

// ---- spec helpers ----------------------------------------------------------------------------------------------

const wire = (r0: number, r1: number, c0: number, c1: number, header: 0 | 1 = 1): WireRange => ({ r0, r1, c0, c1, header, totals: 0 })
const local = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const sortSpec = (range: WireRange, keys: WireSortSpec['keys'], caseSensitive = false): WireSortSpec => ({ range, keys, caseSensitive })
const GRID = wire(0, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
const contains = (first: string): ResolvedRule => ({ kind: 'text', op: 'contains', first, wildcards: false })
/** The server's merged value list of `(range, col)` (every page). */
async function serverList(sv: Serve, range: WireRange, col: number, view: CanonicalView = IDENTITY) {
  const items: Array<{ id: number; text: string; selected: boolean }> = []
  let offset: number | undefined = 0, blanks = { present: false, selected: false }
  while (offset !== undefined) {
    const r = await sv({ op: 'values', sheetIndex: 0, view, range, col, offset })
    if (r.op !== 'values') throw new Error('values')
    items.push(...r.page.items)
    blanks = r.page.blanks
    offset = r.next
  }
  return { items, blanks }
}
const valuesRule = (range: WireRange, col: number, positions: number[], blanks: boolean, mode: 'include' | 'exclude' = 'include'): ResolvedRule => ({ kind: 'values', list: { range, col }, positions, mode, blanks })
const filterOf = (range: WireRange, columns: Array<[number, ResolvedRule]>): ResolvedFilter => ({ range, columns })

describe('the sheet-view job (S4 §6; plan P3, P5, P6, P13)', () => {
  test('sort parity', async () => {
    const bytes = await sortWithheld()
    // sortWithheld's unfilled shape over B20 is a cell-anchored drawing: a sort whose band holds row 20 is refused
    // drawing-impact by both, so the keyed band is A21:F40 (B21, B22, D23 and the table C30:C34 inside it).
    const band = wire(20, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1, 0)
    const keySets: Array<[WireSortSpec['keys'], boolean]> = [
      [[{ col: 1, order: 'asc' }], false], [[{ col: 1, order: 'desc' }], false], [[{ col: 3, order: 'asc' }], true],
      [[{ col: 2, order: 'asc' }, { col: 1, order: 'desc' }], false],
    ]
    for (const policy of ['visible-only', 'full'] as const) {
      const sv = await serve(bytes, policy)
      const o = await oracle(policy === 'full' ? await oracleDoc(bytes) : await withheldOracle(bytes))
      try {
        const whole = await sv({ op: 'derive', sheetIndex: 0, view: { sort: sortSpec(GRID, [{ col: 1, order: 'asc' }]), sortFilter: null, filter: null } }).catch((e: Error) => e.message)
        expect(String(whole)).toContain('drawing-impact')
        expect(await o.sort({ range: local(GRID), keys: [{ col: 1, order: 'asc' }], caseSensitive: false })).toBe('refused')
        for (const [keys, caseSensitive] of keySets) {
          const s = sortSpec(band, keys, caseSensitive)
          const r = viewOf(await sv({ op: 'derive', sheetIndex: 0, view: { sort: s, sortFilter: null, filter: null } }))
          expect(await o.sort({ range: local(band), keys, caseSensitive })).toBe('committed')
          expect({ policy, keys, order: Array.from(r.order.bands[0]!.sourceAtOffset) }).toEqual({ policy, keys, order: o.order(band.r0, band.r1) })
        }
      } finally { await o.dispose() }
    }
  }, 30_000)

  test('filter parity and C1', async () => {
    const bytes = await filterAutoHidden()
    const F = wire(0, 19, 0, 2)
    const hidden = FILTER_AUTO_HIDDEN.hidden.map((r) => r - 1)
    const sv = await serve(bytes, 'visible-only')
    const o = await oracle(await oracleDoc(bytes, (r) => hidden.includes(r)))
    try {
      const list = await serverList(sv, F, 0)
      const cases: Array<[string, ResolvedFilter | null, () => Promise<FilterSpec | null>]> = [
        ['contains ""', filterOf(F, [[0, contains('')]]), async () => ({ range: local(F), columns: new Map([[0, { kind: 'text', op: 'contains', first: '', wildcards: false }]]) })],
        ['contains keep', filterOf(F, [[0, contains('keep')]]), async () => ({ range: local(F), columns: new Map([[0, { kind: 'text', op: 'contains', first: 'keep', wildcards: false }]]) })],
        ['values all', filterOf(F, [[0, valuesRule(F, 0, list.items.map((i) => i.id), true)]]), async () => null],
        // Blanks only: locally the collapsed rows' blank keys match and would show; C1 keeps them hidden.
        ['blanks only', filterOf(F, [[0, valuesRule(F, 0, [], true)]]), async () => ({ range: local(F), columns: new Map([[0, (await o.values(local(F), 0, () => false, true))!]]) })],
        ['cleared', null, async () => null],
      ]
      for (const [name, f, spec] of cases) {
        const r = viewOf(await sv({ op: 'derive', sheetIndex: 0, view: { sort: null, sortFilter: null, filter: f } }))
        expect(await o.filter(await spec())).toMatch(/committed|unchanged/)
        const shown = shownSources(r.order, 0, 19, new Set(hidden)), local = o.shown(0, 19)
        for (const row of hidden) expect({ name, row, shown: shown.has(row) }).toEqual({ name, row, shown: false })
        if (name === 'blanks only') for (const row of hidden) expect(local.has(row)).toBe(true)
        for (let row = 0; row < 20; row++) if (!hidden.includes(row)) expect({ name, row, shown: shown.has(row) }).toEqual({ name, row, shown: local.has(row) })
      }
    } finally { await o.dispose() }

    // Under 'full' the filter decides alone, as locally: `contains ""` shows rows 4 and 7; clearing it leaves them
    // hidden again (their saved `hidden="1"` is the authored state; controller ruling, Task 5).
    const full = await serve(bytes, 'full')
    const of = await oracle(await oracleDoc(bytes))
    try {
      const shownRows = await full({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(F, [[0, contains('')]]) })
      expect(await of.filter({ range: local(F), columns: new Map([[0, { kind: 'text', op: 'contains', first: '', wildcards: false }]]) })).toBe('committed')
      expect(shownSources(viewOf(shownRows).order, 0, 19)).toEqual(of.shown(0, 19))
      for (const row of hidden) expect(of.shown(0, 19).has(row)).toBe(true)
      const cleared = viewOf(await full({ op: 'filter', sheetIndex: 0, base: viewOf(shownRows).view, filter: null }))
      expect(await of.filter(null)).toBe('committed')
      expect(cleared.view).toEqual(IDENTITY)
      expect(of.shown(0, 19)).toEqual(new Set(Array.from({ length: 20 }, (_, r) => r).filter((r) => !hidden.includes(r))))
    } finally { await of.dispose() }
  }, 30_000)

  test('withheld cells sort and list as blank', async () => {
    const bytes = await sortWithheld()
    const sv = await serve(bytes, 'visible-only')
    const o = await oracle(await withheldOracle(bytes))
    try {
      const band = wire(20, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1, 0)
      const r = viewOf(await sv({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort: sortSpec(band, [{ col: 1, order: 'asc' }]) }))
      expect(await o.sort({ range: local(band), keys: [{ col: 1, order: 'asc' }], caseSensitive: false })).toBe('committed')
      const order = Array.from(r.order.bands[0]!.sourceAtOffset)
      expect(order).toEqual(o.order(band.r0, band.r1))
      const tail = order.slice(-2)
      expect(new Set(tail)).toEqual(new Set([rcOf(SORT_WITHHELD.hiddenFormat)[0], rcOf(SORT_WITHHELD.cfHidden)[0]]))

      // B20 sits under the drawing, so no band holding it sorts; its value still never lists.
      const b = await serverList(sv, GRID, 1)
      for (const v of ['1', '5', '95']) expect(b.items.map((i) => i.text)).not.toContain(v)
      expect(b.blanks.present).toBe(true)
      expect(b.items.map((i) => i.text)).toEqual(await o.labels(local(GRID), 1))
      const d = (await serverList(sv, GRID, 3)).items.map((i) => i.text)
      expect(d).toContain(SORT_WITHHELD.maskedDisplay)
      expect(d).not.toContain(SORT_WITHHELD.maskedText)
      // A filter range below the drawing (hiding rows above it would move it: drawing-impact).
      const masked = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(band, [[3, contains('MASK')]]) }))
      expect(masked.order.head.counts).toEqual({ visible: 0, total: band.r1 - band.r0 + 1 })
    } finally { await o.dispose() }
  }, 30_000)

  test('replay equals incremental', async () => {
    const run = async (bytes: Uint8Array, F: WireRange, S: WireSortSpec) => {
      const sv = await serve(bytes, 'visible-only')
      const d = (await serverList(sv, F, 3)).items
      const F1 = filterOf(F, [[3, valuesRule(F, 3, d.filter((_, i) => i % 3 === 0).map((i) => i.id), false)]])
      const F2 = filterOf(F, [[3, contains('e')]])
      const F3 = filterOf(F, [[1, valuesRule(F, 1, [], true)]])
      const steps: Array<{ name: string; job: (base: CanonicalView) => SheetViewJob }> = [
        { name: 'F1', job: (base) => ({ op: 'filter', sheetIndex: 0, base, filter: F1 }) },
        { name: 'S', job: (base) => ({ op: 'sort', sheetIndex: 0, base, sort: S }) },
        { name: 'F2', job: (base) => ({ op: 'filter', sheetIndex: 0, base, filter: F2 }) },
        { name: 'F3', job: (base) => ({ op: 'filter', sheetIndex: 0, base, filter: F3 }) },
      ]
      let base = IDENTITY
      const views: Record<string, { view: CanonicalView; order: SheetOrder }> = {}
      for (const step of steps) {
        const inc = await sv(step.job(base))
        expect({ step: step.name, refused: 'refused' in inc ? inc.refused : null }).toEqual({ step: step.name, refused: null })
        const v = viewOf(inc)
        // A fresh host replays the canonical view from the identity view (no cached derived view).
        const replay = viewOf(await (await serve(bytes, 'visible-only'))({ op: 'derive', sheetIndex: 0, view: v.view }))
        expect({ step: step.name, ...replay }).toEqual({ step: step.name, ...v })
        views[step.name] = v
        base = v.view
      }
      return { sv, F1, F2, S, views }
    }
    const grid = await run(await sortGrid(), GRID, sortSpec(GRID, [{ col: 1, order: 'asc' }]))
    // The sort pinned the rows F1 hid: without that history the order differs.
    const unpinned = viewOf(await grid.sv({ op: 'derive', sheetIndex: 0, view: { sort: grid.S, sortFilter: null, filter: grid.F2 } }))
    expect(Array.from(unpinned.order.bands[0]!.sourceAtOffset)).not.toEqual(Array.from(grid.views.F2!.order.bands[0]!.sourceAtOffset))
    // sortUnderShape: the picture covers A15:A16; every range lies below it.
    const below = wire(19, SORT_GRID.rows - 1, 0, SORT_GRID.cols - 1)
    await run(await sortUnderShape(), below, sortSpec(below, [{ col: 1, order: 'asc' }]))
  }, 30_000)

  test('the spec moves as P3 says', async () => {
    const sv = await serve(await sortGrid(), 'visible-only')
    const S = sortSpec(GRID, [{ col: 1, order: 'desc' }])
    const F1 = filterOf(GRID, [[3, contains('a')]]), F2 = filterOf(GRID, [[3, contains('e')]])
    const a = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: F1 })).view
    expect(a).toEqual({ sort: null, sortFilter: null, filter: F1 })
    const b = viewOf(await sv({ op: 'sort', sheetIndex: 0, base: a, sort: S })).view
    expect(b).toEqual({ sort: S, sortFilter: F1, filter: F1 })
    const c = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: b, filter: F2 })).view
    expect(c).toEqual({ sort: S, sortFilter: F1, filter: F2 })
    const d = viewOf(await sv({ op: 'sort', sheetIndex: 0, base: c, sort: null }))
    expect(d.view).toEqual({ sort: null, sortFilter: null, filter: F2 })
    expect(d.order.bands).toEqual([])
    expect(d.order.head.filterSummary).toEqual([{ col: 3, kind: 'text' }])
  }, 30_000)

  test('display-text merge', async () => {
    const bytes = await displayMerge()
    const sv = await serve(bytes, 'full')
    const o = await oracle(await oracleDoc(bytes))
    try {
      const list = await serverList(sv, GRID, 1)
      const merged = list.items.filter((i) => i.text === DISPLAY_MERGE.label)
      expect(merged.length).toBe(1)
      expect((await o.labels(local(GRID), 1)).filter((l) => l === DISPLAY_MERGE.label).length).toBe(2)
      const r = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(GRID, [[1, valuesRule(GRID, 1, [merged[0]!.id], false)]]) }))
      const rule = await o.values(local(GRID), 1, (l) => l === DISPLAY_MERGE.label, false)
      expect(await o.filter({ range: local(GRID), columns: new Map([[1, rule!]]) })).toBe('committed')
      const shown = shownSources(r.order, 1, SORT_GRID.rows - 1)
      expect(shown).toEqual(o.shown(1, SORT_GRID.rows - 1))
      expect(shown).toEqual(new Set(DISPLAY_MERGE.refs.map((ref) => rcOf(ref)[0])))
      // The committed rule lists its one entry as selected, the rest not.
      const again = await serverList(sv, GRID, 1, r.view)
      expect(again.items.filter((i) => i.selected).map((i) => i.text)).toEqual([DISPLAY_MERGE.label])
      expect(again.blanks.selected).toBe(false)
    } finally { await o.dispose() }
  }, 30_000)

  test('two values columns', async () => {
    const bytes = await sortGrid()
    const sv = await serve(bytes, 'visible-only')
    const o = await oracle(await gridOracle(bytes))
    try {
      const b = (await serverList(sv, GRID, 1)).items, d = (await serverList(sv, GRID, 3)).items
      const bPick = new Set(b.filter((i) => Number(i.text) > 30).map((i) => i.text))
      const dPick = new Set(d.filter((i) => /a/.test(i.text)).map((i) => i.text))
      const r = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(GRID, [
        [1, valuesRule(GRID, 1, b.filter((i) => bPick.has(i.text)).map((i) => i.id), false)],
        [3, valuesRule(GRID, 3, d.filter((i) => !dPick.has(i.text)).map((i) => i.id), false, 'exclude')],
      ]) }))
      const rb = await o.values(local(GRID), 1, (l) => bPick.has(l), false), rd = await o.values(local(GRID), 3, (l) => dPick.has(l), false)
      expect(await o.filter({ range: local(GRID), columns: new Map([[1, rb!], [3, rd!]]) })).toBe('committed')
      const shown = shownSources(r.order, 1, SORT_GRID.rows - 1)
      expect(shown.size).toBeGreaterThan(0)
      expect(shown).toEqual(o.shown(1, SORT_GRID.rows - 1))
    } finally { await o.dispose() }
  }, 30_000)

  test('local guards refuse', async () => {
    const shape = await serve(await sortUnderShape(), 'visible-only')
    expect(await shape({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort: sortSpec(GRID, [{ col: 1, order: 'asc' }]) })).toEqual({ op: 'view', refused: { reason: 'drawing-impact' } })
    const merges = await serve(await sortMerges(), 'visible-only')
    expect(await merges({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort: sortSpec(GRID, [{ col: 1, order: 'asc' }]) })).toEqual({ op: 'view', refused: { reason: 'merge', range: { r0: 4, r1: 5, c0: 1, c1: 1 } } })
  }, 30_000)

  test('range detection over servable occupancy', async () => {
    const bytes = await sortWithheld()
    const sv = await serve(bytes, 'visible-only')
    const o = await oracle(await withheldOracle(bytes))
    try {
      const at = rcOf('C31')
      const r = await sv({ op: 'detect', sheetIndex: 0, kind: 'filter', selection: { active: at, anchor: null, ranges: [] } })
      const local = await o.detect('filter', at)
      if (local.kind !== 'complete' || r.op !== 'detect' || !('detection' in r)) throw new Error(`detect: ${JSON.stringify(r)} / ${local.kind}`)
      expect(r.detection).toEqual({ range: { r0: local.range.r0, r1: local.range.r1, c0: local.range.c0, c1: local.range.c1, header: local.range.headerRows, totals: local.range.totalsRows }, source: local.source.kind, header: local.header, overrideIgnored: local.overrideIgnored })
    } finally { await o.dispose() }
  }, 30_000)

  test('painted keys list as the tile paints them (Task 1 style note)', async () => {
    const sv = await serve(await paintedFormats(), 'visible-only')
    const range = wire(0, 3, 1, 2)
    expect((await serverList(sv, range, 1)).items.map((i) => i.text)).toEqual([PAINTED_FORMATS.tablePaints])
    expect((await serverList(sv, range, 2)).items.map((i) => i.text)).toEqual([PAINTED_FORMATS.cfPaints])
  }, 30_000)

  test('the pool runs the job on the docRef worker', async () => {
    const pool = new Pool({ workers: 1 }); pools.push(pool)
    await pool.open('grid', await sortGrid(), 'visible-only', { sheetMode: 'tiles' })
    const route = { docRef: 'grid', contentPolicy: 'visible-only' as const, clock: CLOCK }
    const view: CanonicalView = { sort: sortSpec(GRID, [{ col: 1, order: 'asc' }]), sortFilter: null, filter: filterOf(GRID, [[3, contains('e')]]) }
    const r = viewOf(await pool.sheetView('s1', { ...route, op: 'derive', sheetIndex: 0, view }))
    const here = viewOf(await (await serve(await sortGrid(), 'visible-only'))({ op: 'derive', sheetIndex: 0, view }))
    expect(r.order.bands[0]!.sourceAtOffset).toBeInstanceOf(Uint32Array)
    expect(r).toEqual(here)
    expect(await pool.sheetView('s1', { ...route, op: 'sort', sheetIndex: 0, base: IDENTITY, sort: sortSpec(GRID, [{ col: 5, order: 'asc' }]) })).toMatchObject({ op: 'view' })
  }, 30_000)

  test('a million rows', async () => {
    const sv = await serve(await oneMillionSort(), 'full')
    // The port (key view, engine) is built by the first job; the measured job is the sort.
    viewOf(await sv({ op: 'derive', sheetIndex: 0, view: IDENTITY }))
    const range = wire(0, ONE_MILLION_ROWS - 1, 0, 0, 0)
    const sort = sortSpec(range, [{ col: 0, order: 'asc' }])
    Bun.gc(true)
    const baseline = process.memoryUsage().heapUsed
    let peak = baseline
    const sampler = setInterval(() => { peak = Math.max(peak, process.memoryUsage().heapUsed) }, 2)
    let r: SheetViewReply
    try { r = await sv({ op: 'sort', sheetIndex: 0, base: IDENTITY, sort }) } finally { clearInterval(sampler) }
    peak = Math.max(peak, process.memoryUsage().heapUsed)
    const v = viewOf(r)
    expect(v.order.bands[0]!.sourceAtOffset[0]).toBe(ONE_MILLION_ROWS - 1)
    expect(peak - baseline).toBeLessThan(orderBytes(v.view) + 32 * 1024 * 1024)
  }, 30_000)
})
