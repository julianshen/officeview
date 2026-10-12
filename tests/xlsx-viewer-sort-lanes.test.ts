import { expect, it } from 'vitest'
import { LaneBuilder } from '../src/xlsx/view/source'
import { sortRowOrder } from '../src/xlsx/view/sort'
import { BOUNDED_SORT_LANES } from '../src/xlsx/view/sortLanes'
import { sortRowOrder as baseline } from './helpers/sort-baseline'
import { createFixture } from './helpers/xlsx-viewer-fixtures'
import type { ColumnEntry, SourceStore, SortSpec } from '../src/xlsx/view/contracts'

it('requests bounded lanes and matches the frozen mixed UTF16/error/blank permutation', async () => {
  const values = [null, 0, -0, NaN, Infinity, true, false, '#N/A', '2', '02', 'item-10', 'item-2', '\ud800', '\udc00', 'é', 'e\u0301']
  const raw = createFixture('F2', { rowCount: 5003 }).createStore({ read: { cellAt: (row, col) => ({ value: values[(row * 7 + col) % values.length], valueIsError: (row + col) % 13 === 0, hasCachedValue: true, styleIndex: 0 }) } })
  let laneRequests = 0
  const store = { ...raw, [BOUNDED_SORT_LANES]: true as const, queryColumn: async (req: Parameters<SourceStore['queryColumn']>[0], opts: Parameters<SourceStore['queryColumn']>[1]) => {
    const page = await raw.queryColumn(req, opts)
    if (req.encoding !== 'lanes') return page
    laneRequests++
    if (req.rows.kind !== 'intervals') throw new Error('intervals required')
    const { r0, r1 } = req.rows.intervals[0], builder = new LaneBuilder()
    for (const item of page.items) if (item.kind === 'cell' || item.kind === 'unavailable') builder.add(item.row, item.cell.value, item.cell.valueIsError, item.cell.styleIndex, item.kind === 'unavailable')
    return { ...page, items: [builder.build(r0, r1)] }
  } }
  for (const direction of ['asc', 'desc'] as const) {
    const spec: SortSpec = { range: { r0: 0, r1: 5002, c0: 0, c1: 1, headerRows: 1, totalsRows: 1 }, keys: [{ col: 0, order: direction }, { col: 1, order: 'desc' }], caseSensitive: false }
    const opts = { now: () => 0, yield: async () => {} }, current = await sortRowOrder(store, spec, opts), old = await baseline(raw, spec, opts)
    if (current.kind !== 'complete' || old.kind !== 'complete') throw new Error(JSON.stringify({ current, old }))
    expect(current.order.bands[0].sourceAtOffset).toEqual(old.order.bands[0].sourceAtOffset)
    expect(current.order.bands[0].displayOffsetOfSource).toEqual(old.order.bands[0].displayOffsetOfSource)
  }
  expect(laneRequests).toBeGreaterThan(0)
})

it('rejects malformed and overlapping lanes instead of publishing an incomplete permutation', async () => {
  const raw = createFixture('F2', { rowCount: 3 }).createStore()
  const b = new LaneBuilder(); b.add(0, 1, false, 0, false); b.add(1, 'a', false, 0, false)
  const lane = b.build(0, 2)
  const bad = [ { ...lane, tags: new Uint8Array(1) }, { ...lane, rows: Uint32Array.of(1, 0) },
    { ...lane, textStart: Uint32Array.of(0, 0, 999) }, { ...lane, tags: Uint8Array.of(6, 2) },
    { ...lane, tags: Uint8Array.of(1, 2), numbers: Float64Array.of(NaN, 0) } ]
  const spec: SortSpec = { range: { r0: 0, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }
  for (const items of [...bad.map(item => [item]), [lane, lane]] as ColumnEntry[][]) {
    const store = { ...raw, [BOUNDED_SORT_LANES]: true as const, queryColumn: async () => ({ items, stamp: raw.stamp, completeness: 'complete' as const }) }
    expect(await sortRowOrder(store, spec, { now: () => 0, yield: async () => {} })).toMatchObject({ kind: 'error', error: { code: 'page-error' } })
  }
})

it('falls back before invoking a typed producer when its bounded staging cannot be admitted', async () => {
  const raw = createFixture('F2', { rowCount: 3 }).createStore()
  let typedCalls = 0
  const store = { ...raw, [BOUNDED_SORT_LANES]: true as const, queryColumn: async (req: Parameters<SourceStore['queryColumn']>[0], opts: Parameters<SourceStore['queryColumn']>[1]) => {
    if (req.encoding !== 'lanes') return raw.queryColumn(req, opts)
    typedCalls++
    const b = new LaneBuilder(); for (let i = 0; i < 3; i++) b.add(i, i, false, 0, false)
    return { items: [b.build(0, 2)], stamp: raw.stamp, completeness: 'complete' as const }
  } }
  const spec: SortSpec = { range: { r0: 0, r1: 2, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }
  await sortRowOrder(store, spec, { maxWorkerBytes: 2000, now: () => 0, yield: async () => {} })
  expect(typedCalls).toBe(0)
})

it('preserves unavailable counts/firstAddress and aborts a typed read before publication', async () => {
  const raw = createFixture('F2', { rowCount: 5003 }).createStore()
  const store = { ...raw, [BOUNDED_SORT_LANES]: true as const, queryColumn: async (req: Parameters<SourceStore['queryColumn']>[0]) => {
    if (req.rows.kind !== 'intervals') throw new Error('intervals required')
    const { r0, r1 } = req.rows.intervals[0], b = new LaneBuilder()
    for (let row = r0; row <= r1; row++) b.add(row, row, false, 0, row === 1)
    return { items: [b.build(r0, r1)], stamp: raw.stamp, completeness: 'complete' as const }
  } }
  const spec: SortSpec = { range: { r0: 0, r1: 5002, c0: 0, c1: 1, headerRows: 0, totalsRows: 0 }, keys: [{ col: 1, order: 'asc' }, { col: 0, order: 'asc' }], caseSensitive: false }
  expect(await sortRowOrder(store, spec, { now: () => 0, yield: async () => {} })).toMatchObject({ kind: 'refused', refusal: { count: 2, firstAddress: { row: 1, col: 0 } } })
  const controller = new AbortController()
  expect(await sortRowOrder(store, spec, { now: () => 0, signal: controller.signal, yield: async phase => { if (phase === 'read') controller.abort() } })).toEqual({ kind: 'cancelled' })
})
