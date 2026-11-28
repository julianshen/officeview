/**
 * Stage 4c Task 13 — corpus row-order parity (S4 §9 4c; plan P3, P5; parent §12 C1). For every corpus XLSX, under
 * `'full'` and under `visible-only`, and for each sheet with a used range (at most 20 sheets, its first 3 columns that
 * hold content), a `sheet-view@1` session over the whole server path runs, from the identity view each time:
 * - a sort ascending and a sort descending by the column;
 * - a `values` filter selecting every list entry but the first (and the blanks), a `text contains` filter on the first
 *   entry's first character, and a blanks-only filter.
 *
 * The range is the used range from A1 with one header row. Each `sheet-order` (or refusal) must equal the local engine
 * (`viewerControls`) run here over the projected package parsed apart from the server, with the key view's withheld
 * cells deleted (and its masked or painted cells as the key view reads them), then C1: under `visible-only` a filtered
 * body row shows iff the local filter shows it and it is not an authored collapsed row. A refusal must be the local
 * refusal's reason. The test counts the files, comparisons and sheets with a withheld key cell it visited.
 */
import { afterAll, describe, expect, test } from 'bun:test'
import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { OfficePackage } from '../../src/core/zip'
import { parseXlsx } from '../../src/xlsx/parse'
import { viewerControls } from '../../src/components/xlsx/viewerControlsClient'
import type { CacheIdentity, ColumnRule, FilterRange, SheetView } from '../../src/xlsx/view/contracts'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { createIdentityOrder } from '../../src/xlsx/view/order'
import { emptyGridSelection } from '../../src/xlsx/view/selection'
import { createObjectWorkbookSource } from '../../src/xlsx/view/source'
import { visibilityFromHidden } from '../../src/xlsx/view/visibility'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
import type { XlsxCell, XlsxDocument, XlsxSheet } from '../../src/xlsx/types'
import { decodeSheetLayout, decodeSheetOrder, decodeValuesPage, type SheetOrder, type ValuesPage, type WireRange } from '../../src/stream/sheet-wire'
import { projectPackage } from '../src/policy'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import { sheetKeyView } from '../src/sheet/key-view'
import { holdsContent, servableSheet } from '../src/sheet/snapshot'
import { Client, disposePool, start } from './harness'

type Policy = 'visible-only' | 'full'
const VIEW = ['sheet-view@1']
const T0 = Date.UTC(2026, 10, 27, 12)
const MAX_SHEETS = 20, KEY_COLS = 3
const hosts: DocumentHost[] = []
afterAll(async () => {
  for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k)
  await disposePool()
})
const local = (r: WireRange): FilterRange => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1, headerRows: r.header, totalsRows: r.totals })
const signal = () => new AbortController().signal
type Text = Record<string, any>

// ---- the oracle ----------------------------------------------------------------------------------------------------

/** The key view of every sheet of `bytes` under `policy` at `clock`, from an in-process host of its own. */
async function keySheets(bytes: Uint8Array, policy: Policy, clock: SheetClock, count: number): Promise<Array<{ sheet: XlsxSheet; withheld: number }>> {
  const host = new DocumentHost(); hosts.push(host)
  const o = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  const saved = (host as unknown as { docs: Map<string, { savedParse: XlsxDocument }> }).docs.get(o.docKey)!.savedParse
  const workbook: XlsxDocument = { ...saved, sheets: saved.sheets.map((s) => servableSheet(s, policy).sheet) }
  const out: Array<{ sheet: XlsxSheet; withheld: number }> = []
  for (let i = 0; i < count; i++) {
    const snap = await host.sheetSnapshot(o.docKey, i, clock)
    const view = await sheetKeyView(snap, saved.sheets[i]!, { workbook, clock, budget: { reserve: () => {}, release: () => {} } })
    out.push({ sheet: view.sheet, withheld: view.withheld })
    view.release()
  }
  return out
}

/** Sheet `i` of `doc` (parsed apart) with the key view's withheld cells deleted and its rewritten cells taken; the
 * deleted and rewritten cells of `cols` are counted. */
function deletedSheet(parsed: XlsxSheet, key: XlsxSheet, cols: ReadonlySet<number>): { sheet: XlsxSheet; changed: number } {
  const keyed = new Map<string, XlsxCell>()
  for (const row of key.rows) for (const c of row.cells) keyed.set(`${c.row}:${c.col}`, c)
  let changed = 0
  const rows = parsed.rows.map((row) => ({ ...row, cells: row.cells.flatMap((c): XlsxCell[] => {
    if (!holdsContent(c)) return [c]
    const k = keyed.get(`${c.row}:${c.col}`)
    if (!k || !holdsContent(k)) { if (cols.has(c.col)) changed++; return [] }
    if (k.value !== c.value && cols.has(c.col)) changed++
    return [{ ...c, value: k.value, styleIndex: k.styleIndex, ...(k.style ? { style: k.style } : {}) }]
  }) }))
  return { sheet: { ...parsed, rows }, changed }
}

async function oracle(doc: XlsxDocument, sheetIndex: number) {
  const source = createObjectWorkbookSource(doc, 'oracle')
  const sheetId = source.sheets[sheetIndex]!.sheetId
  const baseHidden = source.geometry(sheetId).baseHidden
  let view: SheetView = { sheetId, revision: 0, order: createIdentityOrder(), baseHidden, visibility: visibilityFromHidden(baseHidden), sort: null, filter: null, selection: emptyGridSelection() }
  let viewGen = 0
  const owner = (): CacheIdentity => ({ docId: 'oracle', sheetId, valueGen: source.store(sheetId).stamp.valueGen, fontGen: 0, viewGen })
  const snapshot = () => ({ owner: owner(), filter: view.filter, sort: view.sort, records: { visible: 0, total: 0, completeness: 'complete' as const } })
  const publish = (next: Pick<SheetView, 'order' | 'visibility' | 'filter' | 'sort'>) => { view = { ...view, ...next }; viewGen++; return snapshot() }
  const controls = await viewerControls({ doc, source, sheetId, snapshot, view: () => view, publish, ledger: new MainMemoryLedger() }, signal())
  const client = controls.client
  const outcome = (r: Awaited<ReturnType<typeof client.sort>>) => r.kind === 'refused' ? `refused:${r.refusal.reason}` : r.kind
  return {
    async reset() { await client.filter(owner(), null, signal()); await client.sort(owner(), null as never, signal()) },
    async sort(range: WireRange, col: number, order: 'asc' | 'desc') { return outcome(await client.sort(owner(), { range: local(range), keys: [{ col, order }], caseSensitive: false }, signal())) },
    async filter(range: WireRange, col: number, rule: ColumnRule | null) { return outcome(await client.filter(owner(), rule ? { range: local(range), columns: new Map([[col, rule]]) } : null, signal())) },
    async values(range: WireRange, col: number, pick: (label: string) => boolean, blanks: boolean): Promise<ColumnRule | null> {
      const port = await client.values(owner(), local(range), col, null, signal())
      const page = await port.read({ text: '', offset: 0 }, signal())
      const quiet = { text: '', offset: 10_000 }
      await port.change({ kind: 'all', selected: false, includeBlanks: false }, quiet, signal())
      for (const e of page.entries) if (pick(e.label)) await port.change({ kind: 'toggle', ordinal: e.ordinal, selected: true }, quiet, signal())
      await port.change({ kind: 'blanks', selected: blanks }, quiet, signal())
      const rule = await port.candidate(signal())
      port.transfer()
      return rule
    },
    shown(r0: number, r1: number) {
      const out = new Set<number>()
      for (let d = r0; d <= r1; d++) if (view.visibility.visible(d)) out.add(view.order.displayToSource(d))
      return out
    },
    order(r0: number, r1: number) { return Array.from({ length: r1 - r0 + 1 }, (_, i) => view.order.displayToSource(r0 + i)) },
    dispose: async () => { controls.dispose(); await source.release() },
  }
}

/** Source rows of `[r0, r1]` an order shows (a slot outside the mask shows unless `authoredHidden` holds its source). */
function shownSources(order: SheetOrder, r0: number, r1: number, authoredHidden: ReadonlySet<number>): Set<number> {
  const out = new Set<number>()
  const m = order.mask
  for (let d = r0; d <= r1; d++) {
    const band = order.bands.find((b) => b.r0 <= d && d <= b.r1)
    const src = band ? band.sourceAtOffset[d - band.r0]! : d
    if (!m || d < m.r0 || d > m.r1) { if (!authoredHidden.has(src)) out.add(src); continue }
    const i = d - m.r0
    if ((m.words[i >>> 5]! >>> (i & 31)) & 1) out.add(src)
  }
  return out
}
const displayed = (order: SheetOrder, r0: number, r1: number) => Array.from({ length: r1 - r0 + 1 }, (_, i) => {
  const d = r0 + i, band = order.bands.find((b) => b.r0 <= d && d <= b.r1)
  return band ? band.sourceAtOffset[d - band.r0]! : d
})

// ---- the server session --------------------------------------------------------------------------------------------

async function session(bytes: Uint8Array, policy: Policy) {
  let skew = 0, ops = 0
  const h = await start({ doc: bytes, policy: { contentPolicy: policy }, now: () => T0 + skew, tokenTtlSec: 86_400 })
  const c = new Client(h.url)
  const m = await c.hello({ features: VIEW })
  await c.ready()
  while (h.stream.table.get(m.session.sid)?.status !== 'stream') await Bun.sleep(5)
  let rq = 100
  async function ask(type: string, body: Text, partial?: string): Promise<{ parts: Text[]; end?: Text }> {
    skew += 100; if (++ops % 25 === 0) skew += 61_000
    const reqId = ++rq
    c.send({ token: m.session.token, reqId, type, ...body })
    const parts: Text[] = []
    for (;;) {
      const x = await c.next((y) => 'text' in y && y.text.reqId === reqId, 30_000) as { text: Text }
      if (x.text.type !== partial) return { parts, end: x.text }
      parts.push(x.text)
      if (x.text.last === true) return { parts }
    }
  }
  return {
    sheets: (m.sheets ?? []) as Array<{ sheetId: string }>,
    clock: h.stream.sheetClockOf(m.session.sid)!,
    async layout(sheetId: string) { const r = await ask('sheet-layout', { sheetId }, 'sheet-layout'); if (r.end) throw new Error(`layout refused ${JSON.stringify(r.end)}`); return decodeSheetLayout(r.parts) },
    async list(sheetId: string, viewGen: number, range: WireRange, col: number): Promise<{ items: ValuesPage['items']; first: ValuesPage }> {
      const pages: ValuesPage[] = []
      let a = (await ask('filter-values', { sheetId, viewGen, range, col })).end!
      for (;;) {
        if (a.type !== 'filter-values-result') throw new Error(`values refused ${JSON.stringify(a)}`)
        const p = decodeValuesPage(a)
        pages.push(p)
        if (!p.cursor) break
        a = (await ask('filter-values', { sheetId, viewGen, cursor: p.cursor })).end!
      }
      return { items: pages.flatMap((p) => p.items), first: pages[0]! }
    },
    async view(type: 'sheet-sort' | 'sheet-filter', body: Text): Promise<{ order?: SheetOrder; end?: Text }> {
      const r = await ask(type, body, 'sheet-order')
      return r.end ? { end: r.end } : { order: decodeSheetOrder(r.parts) }
    },
    async close() { c.close(); await h.stop() },
  }
}

// ---- the parity loop -----------------------------------------------------------------------------------------------

describe('corpus row-order parity (S4 §9 4c)', () => {
  test('every sheet-order equals the local engine over the deleted fixture, then C1', async () => {
    const dir = resolve(import.meta.dir, '../../corpus')
    const names = readdirSync(dir).filter((n) => n.toLowerCase().endsWith('.xlsx')).sort()
    let files = 0, sheets = 0, compared = 0, committed = 0, refusedBoth = 0, withheldSheets = 0, keyWithheld = 0, c1Sheets = 0
    const refusedOpen: string[] = []
    for (const name of names) {
      const bytes = new Uint8Array(readFileSync(resolve(dir, name)))
      for (const policy of ['full', 'visible-only'] as const) {
        let s: Awaited<ReturnType<typeof session>>
        try { s = await session(bytes, policy) } catch {
          // A document the server refuses to open is refused by the host itself (`unsupported-format`, as in 4a's corpus
          // test), never by a 4c path. It is named in the log.
          const code = await new DocumentHost().open(name, bytes, policy, { sheetMode: 'tiles' }).then(() => 'opened', (e: { code?: string }) => String(e.code))
          expect({ name, policy, code }).toEqual({ name, policy, code: 'unsupported-format' })
          refusedOpen.push(`${name} (${policy})`)
          continue
        }
        try {
          const projected = await projectPackage(await OfficePackage.load(bytes), 'xlsx', policy)
          const parsed = await parseXlsx(projected.pkg, { valuePolicy: 'cached' })
          const count = Math.min(s.sheets.length, MAX_SHEETS)
          expect({ name, policy, sheets: parsed.sheets.length }).toEqual({ name, policy, sheets: s.sheets.length })
          const keys = policy === 'visible-only' ? await keySheets(bytes, policy, s.clock, count) : undefined
          files++
          for (let i = 0; i < count; i++) {
            const sheetId = String(i)
            const layout = await s.layout(sheetId)
            if (!layout.used) continue
            const { r1, c1 } = layout.used
            if (r1 < 1) continue
            const range: WireRange = { r0: 0, r1, c0: 0, c1, header: 1, totals: 0 }
            const body = (r: number) => r >= 1 && r <= r1
            const filled = new Set<number>()
            for (const row of parsed.sheets[i]!.rows) for (const c of row.cells) if (body(c.row) && c.col <= c1 && holdsContent(c)) filled.add(c.col)
            const cols = [...filled].sort((a, b) => a - b).slice(0, KEY_COLS)
            if (!cols.length) continue
            const del = keys ? deletedSheet(parsed.sheets[i]!, keys[i]!.sheet, new Set(cols)) : { sheet: parsed.sheets[i]!, changed: 0 }
            if (del.changed > 0) withheldSheets++
            keyWithheld += keys?.[i]!.withheld ?? 0
            const doc: XlsxDocument = { ...parsed, sheets: parsed.sheets.map((x, j) => j === i ? del.sheet : x) }
            const geometry = buildSheetGeometry(parsed.sheets[i]!)
            const collapsed = new Set<number>()
            if (geometry.kind === 'ready') for (const o of geometry.metrics.rows.overrides) if (o.sizeCss === 0) for (let r = o.start; r <= Math.min(o.end, r1); r++) collapsed.add(r)
            if (policy === 'visible-only' && collapsed.size > 0) c1Sheets++
            const o = await oracle(doc, i)
            sheets++
            let gen = layout.viewGen
            try {
              for (const col of cols) {
                const tag = { name, policy, sheetId, col }
                type Op = { what: string; send: (list: string) => Promise<{ order?: SheetOrder; end?: Text }>; local: () => Promise<string>; kind: 'sort' | 'filter'; listed?: true }
                const list = { ...await s.list(sheetId, gen, range, col), gen }
                expect(list.items.map((x) => x.id)).toEqual(list.items.map((_, k) => k))
                const ops: Op[] = (['asc', 'desc'] as const).map((order) => ({ what: `sort ${order}`, kind: 'sort', send: () => s.view('sheet-sort', { sheetId, viewGen: gen, range, keys: [{ col, order }] }), local: () => o.sort(range, col, order) }))
                const first = list.items[0]
                if (first) {
                  const rest = list.items.slice(1)
                  ops.push({
                    what: 'values but the first', kind: 'filter', listed: true,
                    send: (id) => s.view('sheet-filter', { sheetId, viewGen: gen, range, columns: [[col, { kind: 'values', list: id, mode: 'include', ids: rest.length ? [[rest[0]!.id, rest.at(-1)!.id]] : [], blanks: list.first.blanks.present }]] }),
                    local: async () => o.filter(range, col, await o.values(range, col, (l) => l !== first.text, list.first.blanks.present)),
                  })
                  const ch = [...first.text][0] ?? ''
                  ops.push({
                    what: `contains ${JSON.stringify(ch)}`, kind: 'filter',
                    send: () => s.view('sheet-filter', { sheetId, viewGen: gen, range, columns: [[col, { kind: 'text', op: 'contains', first: ch, wildcards: false }]] }),
                    local: () => o.filter(range, col, { kind: 'text', op: 'contains', first: ch, wildcards: false }),
                  })
                }
                ops.push({
                  what: 'blanks only', kind: 'filter', listed: true,
                  send: (id) => s.view('sheet-filter', { sheetId, viewGen: gen, range, columns: [[col, { kind: 'values', list: id, mode: 'include', ids: [], blanks: true }]] }),
                  local: async () => o.filter(range, col, await o.values(range, col, () => false, true)),
                })
                for (const op of ops) {
                  // A list is bound to the viewGen it was issued at, so a values op lists again at the current one (the
                  // view is the identity view again, so the entries are the first list's).
                  let id = list.first.list
                  if (op.listed && gen !== list.gen) {
                    const again = await s.list(sheetId, gen, range, col)
                    expect(again.items).toEqual(list.items)
                    id = again.first.list
                  }
                  await o.reset()
                  const got = await op.send(id)
                  const want = await op.local()
                  compared++
                  if (got.end) {
                    expect({ ...tag, op: op.what, server: got.end.code === 'view-refused' ? `refused:${got.end.reason}` : got.end.code }).toEqual({ ...tag, op: op.what, server: want })
                    refusedBoth++
                    continue
                  }
                  expect({ ...tag, op: op.what, local: want }).toMatchObject({ ...tag, op: op.what, local: expect.stringMatching(/^(committed|unchanged)$/) })
                  committed++
                  const order = got.order!
                  if (op.kind === 'sort') {
                    expect({ ...tag, op: op.what, order: displayed(order, 1, r1) }).toEqual({ ...tag, op: op.what, order: o.order(1, r1) })
                  } else {
                    const local = o.shown(1, r1)
                    if (policy === 'visible-only') for (const r of collapsed) local.delete(r)
                    expect({ ...tag, op: op.what, shown: shownSources(order, 1, r1, collapsed) }).toEqual({ ...tag, op: op.what, shown: local })
                    expect({ ...tag, op: op.what, counts: order.head.counts }).toEqual({ ...tag, op: op.what, counts: { visible: local.size, total: r1 } })
                  }
                  gen = order.head.viewGen
                  const clear = await s.view(op.kind === 'sort' ? 'sheet-sort' : 'sheet-filter', { sheetId, viewGen: gen, range, ...(op.kind === 'sort' ? { keys: [] } : { columns: [] }) })
                  expect({ ...tag, op: op.what, end: clear.end }).toEqual({ ...tag, op: op.what, end: undefined })
                  gen = clear.order!.head.viewGen
                }
              }
            } finally { await o.dispose() }
          }
        } finally { await s.close() }
      }
    }
    console.log(`sheet sort corpus: ${names.length} xlsx; ${files} (file, policy) sessions; ${sheets} sheet(s); ${compared} comparisons (${committed} committed, ${refusedBoth} refused by both); ${withheldSheets} visible-only sheet(s) with a withheld key cell (${keyWithheld} withheld cells in all their key views); ${c1Sheets} visible-only sheet(s) with an authored collapsed body row (C1); refused to open: ${refusedOpen.join(', ') || 'none'}`)
    expect(files).toBeGreaterThan(0)
    expect(sheets).toBeGreaterThan(0)
    expect(committed).toBeGreaterThan(0)
    expect(refusedOpen.length).toBeLessThan(files)
  }, 120_000)
})
