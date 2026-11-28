/**
 * Stage 4c final review R1 — a value list groups the local entries by display text over EVERY raw value, and the
 * public distinct-entry cap (`VALUE_DISPLAY_LIMIT`) applies to those groups. Selecting a listed entry selects every
 * row that paints its text, whatever raw value it holds, so the include and exclude forms of one selection give the
 * same view. The fixture is the reviewer's: one header and 10,002 numbers under built-in format 1 (`0`): body index
 * 1 holds 2, index 2 holds 3, every other index i holds `1 + i/100000` (10,000 distinct values painting `1`).
 */
import { afterAll, describe, expect, test } from 'bun:test'
import type { SheetOrder, WireRange } from '../../src/stream/sheet-wire'
import { VALUE_DISPLAY_LIMIT } from '../../src/xlsx/view/valueIndex'
import { buildXlsx } from '../../src/testdata/ooxml-builders'
import { DocumentHost } from '../src/render'
import type { SheetClock } from '../src/sheet/clock'
import type { SheetViewJob, SheetViewReply } from '../src/sheet/view-ops'
import type { CanonicalView, ResolvedFilter, ResolvedRule } from '../src/sheet/view-spec'
import { patchStylesXml, patchXlsx } from './sheet-fixtures'

const CLOCK: SheetClock = { day: '2026-11-28', timeZone: 'UTC' }
const IDENTITY: CanonicalView = { sort: null, sortFilter: null, filter: null }
const hosts: DocumentHost[] = []
afterAll(() => { for (const h of hosts) for (const k of [...(h as unknown as { docs: Map<string, unknown> }).docs.keys()]) h.drop(k) })

/** One header (`V`) and `values.length` numbers in column A, every number under built-in format 1 (`0`). */
async function formatZero(values: readonly number[]): Promise<Uint8Array> {
  const rows = [{ r: 1, cells: [{ ref: 'A1', t: 's' as const, v: 0 }] }, ...values.map((v, i) => ({ r: i + 2, cells: [{ ref: `A${i + 2}`, v, style: 99 }] }))]
  return patchXlsx(await buildXlsx([{ name: 'S', rows }], ['V']), async (zip) => {
    let index = -1
    await patchStylesXml(zip, (xml) => {
      const block = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)
      if (!block) throw new Error('no cellXfs')
      index = (block[1]!.match(/<xf\b/g) ?? []).length
      return xml.replace('</cellXfs>', '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>')
    })
    const path = 'xl/worksheets/sheet1.xml'
    zip.file(path, (await zip.file(path)!.async('string')).replaceAll(' s="99"', ` s="${index}"`))
  })
}
const REVIEWER = Array.from({ length: 10_002 }, (_, i) => i === 1 ? 2 : i === 2 ? 3 : 1 + i / 100_000)
const paints = (v: number) => String(Math.round(v))

async function serve(bytes: Uint8Array, policy: 'visible-only' | 'full' = 'visible-only') {
  const host = new DocumentHost()
  hosts.push(host)
  const { docKey } = await host.open('doc', bytes, policy, { sheetMode: 'tiles' })
  return (job: SheetViewJob) => host.sheetView(docKey, CLOCK, policy, job)
}
type Serve = Awaited<ReturnType<typeof serve>>
const wire = (rows: number): WireRange => ({ r0: 0, r1: rows, c0: 0, c1: 0, header: 1, totals: 0 })
const rule = (range: WireRange, positions: number[], mode: 'include' | 'exclude'): ResolvedRule => ({ kind: 'values', list: { range, col: 0 }, positions, mode, blanks: false })
const filterOf = (range: WireRange, r: ResolvedRule): ResolvedFilter => ({ range, columns: [[0, r]] })
const viewOf = (r: SheetViewReply): { view: CanonicalView; order: SheetOrder } => {
  if (r.op !== 'view' || !('order' in r)) throw new Error(`no view: ${JSON.stringify(r)}`)
  return r
}
/** Every page of the server's list of column A. */
async function list(sv: Serve, range: WireRange, search?: { text: string; wildcards: boolean }) {
  const items: Array<{ id: number; text: string }> = []
  let offset: number | undefined = 0, total = 0, truncated = false
  while (offset !== undefined) {
    const r = await sv({ op: 'values', sheetIndex: 0, view: IDENTITY, range, col: 0, offset, ...(search ? { search } : {}) })
    if (r.op !== 'values') throw new Error('values')
    items.push(...r.page.items.map((i) => ({ id: i.id, text: i.text })))
    total = r.page.total; truncated = r.page.truncated
    offset = r.next
  }
  return { items, total, truncated }
}
/** The body indexes (0-based) the order's mask shows. */
function shown(order: SheetOrder, rows: number): number[] {
  const m = order.mask!, out: number[] = []
  for (let b = 0; b < rows; b++) {
    const i = b + 1 - m.r0
    if ((m.words[i >>> 5]! >>> (i & 31)) & 1) out.push(b)
  }
  return out
}

describe('value lists group by display text over every raw value (final review R1)', () => {
  test("the reviewer's 10,002 rows: three entries, complete; include and exclude agree with exact membership", async () => {
    const sv = await serve(await formatZero(REVIEWER))
    const range = wire(REVIEWER.length)
    const l = await list(sv, range)
    expect(l.items).toEqual([{ id: 0, text: '1' }, { id: 1, text: '2' }, { id: 2, text: '3' }])
    expect([l.total, l.truncated]).toEqual([3, false])
    for (const [label, include, exclude] of [['1', [0], [1, 2]], ['2', [1], [0, 2]], ['3', [2], [0, 1]], ['1 or 3', [0, 2], [1]]] as const) {
      const a = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(range, rule(range, [...include], 'include')) }))
      const b = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(range, rule(range, [...exclude], 'exclude')) }))
      const texts = include.map((id) => l.items[id]!.text)
      const want = REVIEWER.flatMap((v, i) => texts.includes(paints(v)) ? [i] : [])
      expect({ label, counts: a.order.head.counts, shown: shown(a.order, REVIEWER.length) }).toEqual({ label, counts: { visible: want.length, total: REVIEWER.length }, shown: want })
      expect({ label, counts: b.order.head.counts, mask: Array.from(b.order.mask!.words) }).toEqual({ label, counts: a.order.head.counts, mask: Array.from(a.order.mask!.words) })
    }
  }, 30_000)

  test('a search lists every group with a matching value, past the raw cap', async () => {
    const sv = await serve(await formatZero(REVIEWER))
    const l = await list(sv, wire(REVIEWER.length), { text: '3', wildcards: false })
    expect(l.items).toEqual([{ id: 2, text: '3' }])
    const one = await list(sv, wire(REVIEWER.length), { text: '1', wildcards: false })
    expect(one.items).toEqual([{ id: 0, text: '1' }])
  }, 30_000)

  test('the cap applies to display groups: one more distinct text than the limit is truncated', async () => {
    const values = Array.from({ length: VALUE_DISPLAY_LIMIT + 1 }, (_, i) => i)
    const sv = await serve(await formatZero(values))
    const l = await list(sv, wire(values.length))
    expect([l.items.length, l.total, l.truncated]).toEqual([VALUE_DISPLAY_LIMIT, VALUE_DISPLAY_LIMIT, true])
  }, 30_000)

  test('a selection whose rows cannot be bounded in either form is refused resource-limit, not committed wrong', async () => {
    // Two display groups (`1` and `2`), each of more distinct raw values than the bound: either form flips too many.
    const half = VALUE_DISPLAY_LIMIT + 1
    const values = Array.from({ length: 2 * half }, (_, i) => (i < half ? 1 : 2) + (i % half) / 1_000_000)
    const sv = await serve(await formatZero(values))
    const range = wire(values.length)
    expect((await list(sv, range)).items).toEqual([{ id: 0, text: '1' }, { id: 1, text: '2' }])
    expect(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(range, rule(range, [0], 'include')) })).toEqual({ op: 'view', refused: { reason: 'resource-limit' } })
    // Selecting both flips nothing: every row shows.
    const all = viewOf(await sv({ op: 'filter', sheetIndex: 0, base: IDENTITY, filter: filterOf(range, rule(range, [0, 1], 'include')) }))
    expect(all.order.head.counts).toEqual({ visible: values.length, total: values.length })
  }, 30_000)
})
