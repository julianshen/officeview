import { describe, expect, it, vi } from 'vitest'
import { prepareSortKeys, sortRowOrder, reduceSortCommand, hiddenSourceRowsFromDisplay, type SortOptions, type SortResult } from '../src/xlsx/view/sort'
import type { SourceStore, SortSpec, ColumnEntry, SortKey } from '../src/xlsx/view/contracts'
import { cellId, MAX_ROWS } from '../src/xlsx/view/coordinates'
import { createRowVisibility } from '../src/xlsx/view/visibility'
import { createIdentityOrder, createRowOrder } from '../src/xlsx/view/order'
import { createAxisMetrics } from '../src/xlsx/view/axes'
import { createSheetMetrics } from '../src/xlsx/view/projection'
import { createFixture, createPrng, f8Cases, type FixtureCell } from './helpers/xlsx-viewer-fixtures'

const fixtureCell = (value: FixtureCell['value'], error = false): FixtureCell => Object.freeze({ value, valueIsError: error, hasCachedValue: true, styleIndex: 0 })
function dataStore(rows: readonly (readonly FixtureCell[])[]): SourceStore {
  return createFixture('F2', { rowCount: Math.max(rows.length, 3) }).createStore({ read: {
    isPopulated: (row, col) => rows[row]?.[col] !== undefined,
    cellAt: (row, col) => rows[row]?.[col] ?? null,
  } })
}
const specFor = (length: number, keys: readonly SortKey[] = [{ col: 0, order: 'asc' }], caseSensitive = false): SortSpec => ({ range: { r0: 0, r1: length - 1, c0: 0, c1: 7, headerRows: 0, totalsRows: 0 }, keys, caseSensitive })
const quiet: SortOptions = { yield: async () => {} }
function rowsOf(result: SortResult): number[] {
  expect(result.kind).toBe('complete')
  if (result.kind !== 'complete') throw new Error(result.kind)
  return [...result.order.bands[0].sourceAtOffset]
}

describe('pure whole-row sort', () => {
  it.each(['asc', 'desc'] as const)('mixed truth table %s keeps errors then blanks terminal', async order => {
    const rows = [fixtureCell(null), fixtureCell('#Z!', true), fixtureCell(true), fixtureCell('item10'), fixtureCell(2), fixtureCell(false), fixtureCell('item2'), fixtureCell(-1), fixtureCell('#A!', true), fixtureCell('')].map(cell => Object.freeze([cell]))
    const store = dataStore(Object.freeze(rows))
    expect(rowsOf(await sortRowOrder(store, specFor(rows.length, [{ col: 0, order }]), quiet))).toEqual(order === 'asc' ? [7, 4, 9, 6, 3, 5, 2, 8, 1, 0] : [2, 5, 3, 6, 9, 4, 7, 8, 1, 0])
    expect(rows[0][0].value).toBeNull()
  })
  it('preserves distinct raw UTF16 keys while collation-equal strings share a rank', async () => {
    const values = ['A', 'a', '\ud800', '\udc00', 'é', 'e\u0301', '', '02', '2']
    const result = await prepareSortKeys(dataStore(values.map(v => [fixtureCell(v)])), specFor(values.length), quiet)
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    expect(result.prepared.lanes[0].strings).toEqual(values)
    const lane = result.prepared.lanes[0]
    expect(lane.stringRanks[0]).toBe(lane.stringRanks[1])
    expect(lane.ids![0]).not.toBe(lane.ids![1])
  })
  it.each(['en', 'sv', 'de'])('matches direct host collation in %s for randomized UTF16', async locale => {
    const pool = ['a', 'A', 'ä', 'z', 'é', 'e\u0301', 'Z2', 'z10', '😀', '\ud800', '\udc00', '', '2', '02', 'b\u0000', 'ö']
    const random = createPrng(717)
    const values = Array.from({ length: 513 }, () => pool[Math.floor(random() * pool.length)])
    for (const caseSensitive of [false, true]) for (const order of ['asc', 'desc'] as const) {
      const collator = new Intl.Collator(locale, { numeric: true, sensitivity: caseSensitive ? 'variant' : 'accent', caseFirst: 'upper' })
      const expected = values.map((_value, row) => row).sort((a, b) => (order === 'asc' ? 1 : -1) * collator.compare(values[a], values[b]) || a - b)
      expect(rowsOf(await sortRowOrder(dataStore(values.map(v => [fixtureCell(v)])), specFor(values.length, [{ col: 0, order }], caseSensitive), { ...quiet, locale }))).toEqual(expected)
    }
  })
  it('pins hidden/protected/source slots, source ties survive replacement and clearing', async () => {
    const values = [99, 4, 1, 1, 0, 90]
    const store = dataStore(values.map(v => [fixtureCell(v)]))
    const spec = { ...specFor(6), range: { ...specFor(6).range, headerRows: 1 as const, totalsRows: 1 as const } }
    const options = { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS, row => row !== 2)) }
    expect(rowsOf(await sortRowOrder(store, spec, options))).toEqual([4, 2, 3, 1])
    expect(rowsOf(await sortRowOrder(store, { ...spec, keys: [{ col: 0, order: 'desc' }] }, options))).toEqual([1, 2, 3, 4])
    const clear = await sortRowOrder(store, null, options)
    expect(clear.kind).toBe('complete')
    if (clear.kind === 'complete') expect(clear.order).toBe(createIdentityOrder())
  })
  it.each([1, 2, 3, 4, 5, 6, 7, 8])('%i keys break ties left to right then immutable source ids', async count => {
    const values = [[...Array(7).fill(0), 2], [...Array(7).fill(0), 1], [...Array(7).fill(0), 1]]
    const keys = Array.from({ length: count }, (_v, col) => ({ col, order: 'asc' as const }))
    expect(rowsOf(await sortRowOrder(dataStore(values.map(row => row.map(v => fixtureCell(v)))), specFor(3, keys), quiet))).toEqual(count === 8 ? [1, 2, 0] : [0, 1, 2])
  })
  it('refuses pending, unavailable and failed source keys without a draft', async () => {
    const fixture = createFixture('F8')
    const unavailable = await sortRowOrder(fixture.createStore({ valuePolicy: 'cached' }), specFor(3, [{ col: 1, order: 'asc' }]), quiet)
    expect(unavailable).toMatchObject({ kind: 'refused', refusal: { reason: 'formula-value-unavailable', count: 1, firstAddress: { row: 1, col: 1 } } })
    const pending = await sortRowOrder(fixture.createStore({ pending: [{ r0: 0, r1: 0, c0: 0, c1: 0 }] }), specFor(3), quiet)
    expect(pending).toEqual({ kind: 'incomplete' })
    const failed = await sortRowOrder(fixture.createStore({ errors: [{ range: { r0: 0, r1: 0, c0: 0, c1: 0 }, code: 'page-error', retryable: true }] }), specFor(3), quiet)
    expect(failed).toMatchObject({ kind: 'error', error: { code: 'page-error' } })
  })
  it('rejects stale page/current stamp and aborts during reads', async () => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const original = store.queryColumn.bind(store)
    const stale: SourceStore = { ...store, queryColumn: async (request, options) => ({ ...await original(request, options), stamp: { ...store.stamp, valueGen: 9 } }) }
    expect(await sortRowOrder(stale, specFor(2), quiet)).toMatchObject({ kind: 'error', error: { code: 'stale-source' } })
    const controller = new AbortController()
    const cancelled: SourceStore = { ...store, queryColumn: async (request, options) => { const page = await original(request, options); controller.abort(); return page } }
    expect(await sortRowOrder(cancelled, specFor(2), { ...quiet, signal: controller.signal })).toEqual({ kind: 'cancelled' })
  })
  it('cycles keys, refuses ninth key, toggles case and clears', () => {
    const range = specFor(4).range
    let state = reduceSortCommand(null, { kind: 'cycle', col: 0, range }).sort
    expect(state?.keys).toEqual([{ col: 0, order: 'asc' }])
    state = reduceSortCommand(state, { kind: 'cycle', col: 0, range }).sort
    expect(state?.keys).toEqual([{ col: 0, order: 'desc' }])
    expect(reduceSortCommand(state, { kind: 'cycle', col: 0, range }).sort).toBeNull()
    state = null
    for (let col = 0; col < 8; col++) state = reduceSortCommand(state, { kind: 'cycle', col, shift: true, range: { ...range, c1: 8 } }).sort
    expect(reduceSortCommand(state, { kind: 'cycle', col: 8, shift: true, range: { ...range, c1: 8 } })).toMatchObject({ sort: state, notice: { reason: 'too-many-sort-keys', maxKeys: 8 } })
    expect(reduceSortCommand(state, { kind: 'toggle-case' }).sort?.caseSensitive).toBe(true)
    expect(reduceSortCommand(state, { kind: 'clear' }).sort).toBeNull()
  })
})

describe('sort ownership, bounded work and resource admission', () => {
  it('only reads requested key/body chunks and detects missing coverage/cursor loops', async () => {
    const fixture = createFixture('F7', { rowCount: 9000 }), store = fixture.createStore()
    const original = store.queryColumn.bind(store)
    const requests: Array<{ col: number; r0: number; r1: number }> = []
    const bounded: SourceStore = { ...store, readRows: async () => { throw new Error('must use columns') }, queryColumn: async (request, options) => {
      expect(request.rows.kind).toBe('intervals')
      if (request.rows.kind !== 'intervals') throw new Error('unexpected rows')
      const [interval] = request.rows.intervals
      expect(interval.r1 - interval.r0 + 1).toBeLessThanOrEqual(4096)
      requests.push({ col: request.col, ...interval })
      return original(request, options)
    } }
    const spec = { ...specFor(9000, [{ col: 1, order: 'asc' }, { col: 7, order: 'desc' }]), range: { ...specFor(9000).range, headerRows: 1 as const, totalsRows: 1 as const } }
    expect((await sortRowOrder(bounded, spec, quiet)).kind).toBe('complete')
    expect(new Set(requests.map(r => r.col))).toEqual(new Set([1, 7]))
    expect(Math.min(...requests.map(r => r.r0))).toBe(1)
    expect(Math.max(...requests.map(r => r.r1))).toBe(8998)
    const missing: SourceStore = { ...store, queryColumn: async () => ({ items: [], completeness: 'complete', stamp: store.stamp }) }
    expect(await sortRowOrder(missing, specFor(2), quiet)).toEqual({ kind: 'incomplete' })
    const looping: SourceStore = { ...missing, queryColumn: async () => ({ items: [], completeness: 'complete', stamp: store.stamp, nextCursor: 'same' }) }
    expect(await sortRowOrder(looping, specFor(2), quiet)).toMatchObject({ kind: 'error', error: { code: 'invalid-cursor' } })
    const entries: ColumnEntry[] = [{ kind: 'blank', r0: 0, r1: 3 }]
    expect(await sortRowOrder({ ...missing, queryColumn: async () => ({ items: entries, completeness: 'complete', stamp: store.stamp }) }, specFor(2), quiet)).toMatchObject({ kind: 'error', error: { code: 'page-error' } })
  })
  it('snapshots caller keys/range/stamp before awaiting and catches live source mutations', async () => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const input = specFor(2), stamp = { ...store.stamp }
    const original = store.queryColumn.bind(store)
    const delayed: SourceStore = { ...store, queryColumn: async (request, options) => {
      const page = await original(request, options)
      input.range.r1 = 0
      ;(input.keys[0] as SortKey).order = 'desc'
      stamp.valueGen = 77
      return page
    } }
    const result = await sortRowOrder(delayed, input, { ...quiet, stamp })
    expect(rowsOf(result)).toEqual([1, 0])
    if (result.kind === 'complete') expect(result.metadata.sort?.range.r1).toBe(1)
    const changed: SourceStore = { ...store, queryColumn: async (request, options) => { const page = await original(request, options); store.stamp.valueGen++; return page } }
    expect(await sortRowOrder(changed, specFor(2), quiet)).toMatchObject({ kind: 'error', error: { code: 'stale-source' } })
  })
  it('reports total unavailable key cells with row-major firstAddress', async () => {
    const unavailable = Object.freeze({ ...fixtureCell(null), formula: 'X()', hasCachedValue: false })
    const store = dataStore([[fixtureCell(1), unavailable], [unavailable, unavailable]])
    const cached = { ...store, valuePolicy: 'cached' as const, queryColumn: async (request: Parameters<SourceStore['queryColumn']>[0], options: Parameters<SourceStore['queryColumn']>[1]) => {
      const page = await store.queryColumn(request, options)
      return { ...page, items: page.items.map(entry => entry.kind === 'cell' && entry.cell.formula ? { kind: 'unavailable' as const, row: entry.row, cell: entry.cell, reason: 'formula-value-unavailable' as const } : entry) }
    } }
    expect(await sortRowOrder(cached, specFor(2, [{ col: 0, order: 'asc' }, { col: 1, order: 'asc' }]), quiet)).toMatchObject({ kind: 'refused', refusal: { count: 3, firstAddress: { row: 0, col: 1 } } })
    // Engine authority is a value, not a cache-presence inference.
    expect((await sortRowOrder(store, specFor(2), quiet)).kind).toBe('complete')
  })
  it('F8 cached values preserve number/date and error-looking text classification', async () => {
    const cases = f8Cases().filter(c => ['plain-number', 'cached-formula', 'date-style', 'cached-error', 'error-looking-text'].includes(c.note))
    const cells = cases.map(c => [Object.freeze(c.cell)])
    const result = await sortRowOrder(dataStore(Object.freeze(cells)), specFor(cells.length), quiet)
    const expected = cases.map((_c, row) => row).sort((a, b) => {
      const ca = cases[a].cell, cb = cases[b].cell
      const ta = ca.valueIsError ? 4 : typeof ca.value === 'number' ? 1 : 2
      const tb = cb.valueIsError ? 4 : typeof cb.value === 'number' ? 1 : 2
      return ta - tb || (ta === 1 ? Number(ca.value) - Number(cb.value) : 0) || a - b
    })
    expect(rowsOf(result)).toEqual(expected)
    expect(cases.find(c => c.note === 'error-looking-text')?.cell.valueIsError).toBe(false)
  })
  it('F4 exact dictionary ranking stays equivalent for unique UTF16 values', async () => {
    const fixture = createFixture('F4', { rowCount: 8193 }), values = Array.from({ length: 8193 }, (_v, row) => String(fixture.cellAt(row, 0)!.value))
    const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'accent', caseFirst: 'upper' })
    expect(rowsOf(await sortRowOrder(fixture.createStore(), specFor(values.length), quiet))).toEqual(values.map((_v, row) => row).sort((a, b) => collator.compare(values[a], values[b]) || a - b))
  })
  it('F7 eight low-cardinality columns use later keys and match direct oracle', async () => {
    const fixture = createFixture('F7', { rowCount: 8193 }), keys = Array.from({ length: 8 }, (_v, col) => ({ col, order: col % 2 ? 'desc' as const : 'asc' as const }))
    const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'accent', caseFirst: 'upper' })
    const expected = Array.from({ length: 8193 }, (_v, row) => row).sort((a, b) => {
      for (const key of keys) {
        const av = fixture.cellAt(a, key.col)!.value!, bv = fixture.cellAt(b, key.col)!.value!
        const diff = typeof av === 'number' ? av - Number(bv) : collator.compare(String(av), String(bv))
        if (diff) return (key.order === 'asc' ? 1 : -1) * diff
      }
      return a - b
    })
    expect(rowsOf(await sortRowOrder(fixture.createStore(), specFor(8193, keys), quiet))).toEqual(expected)
  })
  it.each(['read', 'dictionary', 'candidates', 'merge', 'output', 'inverse', 'publish'] as const)('cancels atomically during %s', async phase => {
    const fixture = createFixture(phase === 'dictionary' ? 'F4' : 'F7', { rowCount: 8193 }), controller = new AbortController(), previous = createIdentityOrder()
    const result = await sortRowOrder(fixture.createStore(), specFor(8193), { signal: controller.signal, now: () => 0, yield: async at => { if (at === phase) controller.abort() } })
    expect(result).toEqual({ kind: 'cancelled' })
    expect(previous.displayToSource(7)).toBe(7)
    expect('order' in result).toBe(false)
  })
  it('cooperates on clock deadline and default macrotask scheduler allows queued abort', async () => {
    const store = dataStore(Array.from({ length: 100 }, (_v, row) => [fixtureCell(100 - row)]))
    let time = 0, yields = 0
    expect((await sortRowOrder(store, specFor(100), { now: () => (time += 11), yield: async () => { yields++ } })).kind).toBe('complete')
    expect(yields).toBeGreaterThan(100)
    const controller = new AbortController()
    const job = sortRowOrder(createFixture('F7', { rowCount: 10000 }).createStore(), specFor(10000), { signal: controller.signal })
    setTimeout(() => controller.abort(), 0)
    expect(await job).toEqual({ kind: 'cancelled' })
  })
  it('admits construction peak against aggregate budget and rejects invalid coordinates', async () => {
    const store = dataStore([[fixtureCell('alpha')], [fixtureCell('beta')]])
    const result = await sortRowOrder(store, specFor(2), quiet)
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    expect(result.ledger.retainedBytes).toBe(16 + 256)
    expect(result.ledger.peakBytes).toBeGreaterThan(result.ledger.retainedBytes)
    expect(await sortRowOrder(store, specFor(2), { ...quiet, residentWorkerBytes: 1000, maxWorkerBytes: 1000 + result.ledger.peakBytes - 1 })).toMatchObject({ kind: 'resource-limit' })
    const invalid = { ...specFor(2), keys: [{ col: 16384, order: 'asc' as const }] }
    expect(await sortRowOrder(store, invalid, quiet)).toMatchObject({ kind: 'error', error: { code: 'invalid-coordinate' } })
  })
  it('rejects invalid key counts, duplicate/outside columns, directions and protected counts', async () => {
    const store = dataStore([[fixtureCell(2)]])
    const invalid = [specFor(1, []), specFor(1, Array.from({ length: 9 }, (_v, col) => ({ col, order: 'asc' as const }))), specFor(1, [{ col: 0, order: 'asc' }, { col: 0, order: 'desc' }]), { ...specFor(1), keys: [{ col: 0, order: 'sideways' }] }, { ...specFor(1), range: { ...specFor(1).range, headerRows: 2 } }, { ...specFor(1), range: { ...specFor(1).range, c1: 0 }, keys: [{ col: 1, order: 'asc' }] }]
    for (const spec of invalid) expect((await sortRowOrder(store, spec as SortSpec, quiet)).kind).toBe('error')
  })
})

describe('full-axis typed sort workspace', () => {
  it('F2 million rows numeric sort uses typed workspace; maximum source row hits top viewport', async () => {
    const fixture = createFixture('F2')
    const store = fixture.createStore({ read: { isPopulated: (_row, col) => col < 8, cellAt: (row, col) => fixtureCell(row + col) } })
    const started = performance.now()
    const result = await sortRowOrder(store, specFor(MAX_ROWS, [{ col: 0, order: 'desc' }]), { ...quiet, now: () => 0 })
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    const band = result.order.bands[0]
    expect(band.sourceAtOffset).toBeInstanceOf(Uint32Array)
    expect(band.displayOffsetOfSource).toBeInstanceOf(Uint32Array)
    expect(result.order.displayToSource(0)).toBe(MAX_ROWS - 1)
    expect(result.order.sourceToDisplay(MAX_ROWS - 1)).toBe(0)
    expect(result.ledger.orderBytes + result.ledger.inverseBytes).toBe(8 * 1024 * 1024)
    expect(result.ledger.indexBytes).toBe(4 * MAX_ROWS)
    expect(result.ledger.scratchBytes).toBe(4 * MAX_ROWS)
    expect(result.ledger.numbersBytes).toBe(8 * MAX_ROWS)
    expect(result.ledger.idsBytes).toBe(0)
    expect(result.ledger.retainedBytes).toBe(8 * MAX_ROWS + 256)
    for (let slot = 0; slot < MAX_ROWS; slot++) {
      if (band.sourceAtOffset[slot] !== MAX_ROWS - 1 - slot || band.displayOffsetOfSource[MAX_ROWS - 1 - slot] !== slot) throw new Error(`Invalid permutation at ${slot}`)
    }
    console.info(`S7.2 F2 million numeric local kernel observation: ${(performance.now() - started).toFixed(1)}ms; immediate injected scheduler; peak=${result.ledger.peakBytes}B`)
    const metrics = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 20 }), cols: createAxisMetrics({ logicalCount: 16384, defaultSizeCss: 64 }), order: result.order })
    expect(metrics.logicalPointToSource(1, 1)).toEqual({ row: MAX_ROWS - 1, col: 0 })
    expect(cellId(MAX_ROWS - 1, 0)).toBe((MAX_ROWS - 1) * 16384)
  }, 30000)
  it('F2 eight full-band Float64 lanes stay <=64MiB and preserve total aggregate ledger', async () => {
    const fixture = createFixture('F2'), store = fixture.createStore({ read: { isPopulated: (_row, col) => col < 8, cellAt: (row, col) => fixtureCell(row % (col + 2)) } })
    const result = await prepareSortKeys(store, specFor(MAX_ROWS, Array.from({ length: 8 }, (_v, col) => ({ col, order: 'asc' }))), { ...quiet, now: () => 0, residentWorkerBytes: 10 * 1024 * 1024 })
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    const { lanes, ledger } = result.prepared
    expect(lanes).toHaveLength(8)
    expect(lanes.every(lane => lane.tags instanceof Uint8Array && lane.numbers instanceof Float64Array && lane.ids === undefined)).toBe(true)
    expect(ledger.numbersBytes).toBeLessThanOrEqual(64 * 1024 * 1024)
    expect(ledger.numbersBytes).toBe(64 * 1024 * 1024)
    expect(ledger.tagsBytes).toBe(8 * MAX_ROWS)
    expect(ledger.retainedBytes).toBe(ledger.numbersBytes + ledger.tagsBytes + ledger.fixedOverheadBytes)
    expect(ledger.aggregatePeakBytes).toBe(ledger.peakBytes + 10 * 1024 * 1024)
  }, 30000)
  it('small-helper and cooperative-builder paths agree with same pinned direct oracle', async () => {
    for (const length of [4096, 4097]) {
      const values = Array.from({ length }, (_v, row) => row % 5)
      const hidden = (row: number) => row % 7 === 0
      const expected = values.map((_v, row) => row).filter(row => !hidden(row)).sort((a, b) => values[a] - values[b] || a - b)
      let next = 0
      const slots = values.map((_v, row) => hidden(row) ? row : expected[next++])
      const result = await sortRowOrder(dataStore(values.map(v => [fixtureCell(v)])), specFor(length), { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS, row => !hidden(row))), now: () => 0 })
      expect(rowsOf(result)).toEqual(slots)
      if (result.kind === 'complete') for (let slot = 0; slot < length; slot++) expect(result.order.sourceToDisplay(slots[slot])).toBe(slot)
    }
  })
})

describe('construction/accounting cancellation regressions', () => {
  it('charges page UTF16 before a read yield and exposes dictionary Map/copy peaks', async () => {
    const long = 'x'.repeat(10000)
    const result = await sortRowOrder(dataStore([[fixtureCell(long)]]), specFor(1), quiet)
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    expect(result.ledger.dictionaryMapPeakBytes).toBeGreaterThanOrEqual(128)
    expect(result.ledger.pagePeakBytes).toBeGreaterThanOrEqual(long.length * 2)
    expect(result.ledger.constructionCopiesBytes).toBeGreaterThanOrEqual(12)
    expect(result.ledger.fixedOverheadBytes).toBeGreaterThan(0)
  })
  it('dictionary index, merge, and rank assignment all expose cancellation boundaries', async () => {
    for (const boundary of [1, 3, 32]) {
      const controller = new AbortController()
      let dictionaryYields = 0
      const result = await sortRowOrder(createFixture('F4', { rowCount: 8193 }).createStore(), specFor(8193), { now: () => 0, signal: controller.signal, yield: async phase => {
        if (phase === 'dictionary' && ++dictionaryYields === boundary) controller.abort()
      } })
      expect(dictionaryYields).toBe(boundary)
      expect(result).toEqual({ kind: 'cancelled' })
    }
  })
  it('every work phase yields by <=4096 checked work units even when clock stays fixed', async () => {
    let checks = 0
    const perPhase = new Map<string, number>()
    const result = await sortRowOrder(createFixture('F4', { rowCount: 16385 }).createStore(), specFor(16385), { now: () => { checks++; return 0 }, yield: async phase => {
      // Clock sampling happens once per work unit, plus scheduler reset/initialization.
      expect(checks).toBeLessThanOrEqual(4097)
      perPhase.set(phase, (perPhase.get(phase) ?? 0) + 1)
      checks = 0
    } })
    expect(result.kind).toBe('complete')
    for (const phase of ['read', 'dictionary', 'candidates', 'merge', 'output', 'inverse', 'publish']) expect(perPhase.has(phase)).toBe(true)
  })
  it('handles row-empty protected bands, initial abort, and thrown source failure without publication', async () => {
    const store = dataStore([[fixtureCell(1)]])
    const spec = { ...specFor(1), range: { ...specFor(1).range, headerRows: 1 as const } }
    const complete = await sortRowOrder(store, spec, quiet)
    expect(complete.kind).toBe('complete')
    if (complete.kind === 'complete') expect(complete.order).toBe(createIdentityOrder())
    const controller = new AbortController(); controller.abort()
    expect(await sortRowOrder(store, null, { signal: controller.signal })).toEqual({ kind: 'cancelled' })
    expect(await sortRowOrder({ ...store, queryColumn: async () => { throw new Error('disk unavailable') } }, specFor(1), quiet)).toMatchObject({ kind: 'error', error: { code: 'page-error', message: 'disk unavailable' } })
  })
})

describe('immutable option ownership', () => {
  it('keeps the original abort signal when the caller replaces options.signal', async () => {
    const controller = new AbortController()
    const options: SortOptions = { now: () => 0, signal: controller.signal, yield: async () => { options.signal = undefined; controller.abort() } }
    expect((await sortRowOrder(createFixture('F7', { rowCount: 8193 }).createStore(), specFor(8193), options)).kind).toBe('cancelled')
  })
  it('keeps the visibility generation captured before preparing keys', async () => {
    const store = dataStore([[fixtureCell(3)], [fixtureCell(2)], [fixtureCell(1)]])
    const original = store.queryColumn.bind(store)
    const options: SortOptions = { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS, row => row !== 0)) }
    const delayed: SourceStore = { ...store, queryColumn: async (request, read) => { const page = await original(request, read); options.hidden = hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS)); return page } }
    expect(rowsOf(await sortRowOrder(delayed, specFor(3), options))).toEqual([0, 2, 1])
  })
})


describe('review behavior regressions', () => {
  it('default MessageChannel scheduling avoids timers and disposes on every outcome', async () => {
    const timer = globalThis.setTimeout
    for (const outcome of ['complete', 'cancelled', 'error'] as const) {
      const controller = new AbortController()
      let channels = 0, closed = 0, messages = 0, time = 0
      class TestChannel {
        constructor() { channels++ }
        port1 = { onmessage: null as (() => void) | null, close: () => { closed++ } }
        port2 = {
          close: () => { closed++ },
          postMessage: () => {
            timer(() => {
              messages++
              if (outcome === 'cancelled') controller.abort()
              this.port1.onmessage?.()
            }, 0)
          },
        }
      }
      const timers = vi.fn(timer)
      vi.stubGlobal('setImmediate', undefined)
      vi.stubGlobal('MessageChannel', TestChannel)
      vi.stubGlobal('setTimeout', timers)
      try {
        const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
        let reads = 0
        const original = store.queryColumn.bind(store)
        const source: SourceStore = { ...store, queryColumn: async (request, options) => {
          if (outcome === 'error' && reads++) throw new Error('source failed after yield')
          return original(request, options)
        } }
        const result = await sortRowOrder(source, specFor(2, [
          { col: 0, order: 'asc' }, { col: 1, order: 'asc' },
        ]), { now: () => (time += 11), signal: controller.signal })
        expect(result.kind).toBe(outcome)
        if (result.kind === 'complete') expect(rowsOf(result)).toEqual([1, 0])
        expect(timers).not.toHaveBeenCalled()
        expect(channels).toBe(1)
        expect(messages).toBeGreaterThan(0)
        expect(closed).toBe(2)
      } finally { vi.unstubAllGlobals() }
    }
  })
  it('invalid locale is unsupported and nonretryable rather than an allocation failure', async () => {
    const store = dataStore([[fixtureCell(1)]])
    for (const run of [prepareSortKeys, sortRowOrder]) {
      expect(await run(store, specFor(1), { ...quiet, locale: 'en_US' })).toMatchObject({
        kind: 'error', error: { code: 'unsupported', retryable: false },
      })
    }
  })
  it('source RangeErrors are page failures rather than allocation failures', async () => {
    const store = dataStore([[fixtureCell(1)]])
    const source: SourceStore = { ...store, queryColumn: async () => { throw new RangeError('source decoder bounds') } }
    expect(await sortRowOrder(source, specFor(1), quiet)).toMatchObject({
      kind: 'error', error: { code: 'page-error', message: 'source decoder bounds' },
    })
  })
  it.each(['asc', 'desc'] as const)('nonfinite numbers share the real #NUM! rank in %s', async order => {
    const cells = [fixtureCell(NaN), fixtureCell(3), fixtureCell('#NUM!', true),
      fixtureCell(Infinity), fixtureCell('#DIV/0!', true), fixtureCell(-Infinity),
      fixtureCell(''), fixtureCell(null), fixtureCell(false), fixtureCell(-2)]
    const store = dataStore(cells.map(cell => [cell]))
    const spec = specFor(cells.length, [{ col: 0, order }])
    expect(rowsOf(await sortRowOrder(store, spec, quiet))).toEqual(order === 'asc'
      ? [9, 1, 6, 8, 4, 0, 2, 3, 5, 7] : [8, 6, 1, 9, 4, 0, 2, 3, 5, 7])
    const prepared = await prepareSortKeys(store, spec, quiet)
    expect(prepared.kind).toBe('complete')
    if (prepared.kind === 'complete') {
      const lane = prepared.prepared.lanes[0]
      expect(lane.errors).toEqual(['#NUM!', '#DIV/0!'])
      for (const row of [0, 2, 3, 5]) expect(lane.ids![row]).toBe(lane.ids![0])
    }
  })
  it('reducer returns typed notices for invalid commands and preserves prior state', () => {
    const sort = specFor(4)
    for (const col of [-1, 16384, 8]) {
      expect(() => reduceSortCommand(sort, { kind: 'cycle', col, range: sort.range })).not.toThrow()
      const result = reduceSortCommand(sort, { kind: 'cycle', col, range: sort.range })
      expect(result.sort).toBe(sort)
      expect(result.notice).toMatchObject({ reason: 'invalid-sort-command', error: { code: 'invalid-coordinate', retryable: false } })
    }
    const range = { ...sort.range, r1: -1 }
    const result = reduceSortCommand(sort, { kind: 'cycle', col: 0, range })
    expect(result.sort).toBe(sort)
    expect(result.notice).toMatchObject({ reason: 'invalid-sort-command', error: { code: 'invalid-source-rect' } })
    expect(reduceSortCommand(null, { kind: 'toggle-case' })).toEqual({ sort: null, notice: { reason: 'no-active-sort' } })
    const malformed = { ...sort, keys: null } as unknown as SortSpec
    expect(reduceSortCommand(malformed, { kind: 'toggle-case' }).sort).toBe(malformed)
    expect(reduceSortCommand(malformed, { kind: 'toggle-case' }).notice).toMatchObject({ reason: 'invalid-sort-command' })
  })
  it('toggle-case without a sort reports a typed notice', () => {
    expect(reduceSortCommand(null, { kind: 'toggle-case' }))
      .toEqual({ sort: null, notice: { reason: 'no-active-sort' } })
  })
  it.each([false, true])('changed range begins a fresh single ascending key even with shift=%s', shift => {
    const sort = specFor(4, [{ col: 0, order: 'asc' }, { col: 1, order: 'desc' }], true)
    const range = { ...sort.range, r1: 5 }
    expect(reduceSortCommand(sort, { kind: 'cycle', col: 0, shift, range }).sort).toEqual({
      range, keys: [{ col: 0, order: 'asc' }], caseSensitive: true,
    })
  })
  it('nonshift multi-key click collapses to that cycled key; ninth key leaves state unchanged', () => {
    const sort = specFor(4, [{ col: 0, order: 'asc' }, { col: 1, order: 'asc' }])
    expect(reduceSortCommand(sort, { kind: 'cycle', col: 1, range: sort.range }).sort?.keys)
      .toEqual([{ col: 1, order: 'desc' }])
    const full = { ...specFor(4, Array.from({ length: 8 }, (_v, col) => ({ col, order: 'asc' }))),
      range: { ...sort.range, c1: 8 } } as SortSpec
    const refused = reduceSortCommand(full, { kind: 'cycle', col: 8, shift: true, range: full.range })
    expect(refused.sort).toBe(full)
    expect(refused.notice).toEqual({ reason: 'too-many-sort-keys', maxKeys: 8 })
  })
})


describe('review policy and full-axis evidence', () => {
  it('hidden unknown and unavailable key rows conservatively block publication', async () => {
    const fixture = createFixture('F8')
    const options = { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS, row => row !== 1)) }
    expect(await sortRowOrder(fixture.createStore({ valuePolicy: 'cached' }),
      specFor(3, [{ col: 1, order: 'asc' }]), options)).toMatchObject({
      kind: 'refused', refusal: { reason: 'formula-value-unavailable', firstAddress: { row: 1, col: 1 } },
    })
    const pending = fixture.createStore({ pending: [{ r0: 1, r1: 1, c0: 0, c1: 0 }] })
    expect(await sortRowOrder(pending, specFor(3), options)).toEqual({ kind: 'incomplete' })
  })
  it('empty string remains text ahead of errors and absent/null blanks in both directions', async () => {
    const store = dataStore([[fixtureCell(null)], [fixtureCell('')], [], [fixtureCell('#NUM!', true)]])
    for (const order of ['asc', 'desc'] as const) {
      expect(rowsOf(await sortRowOrder(store, specFor(4, [{ col: 0, order }]), quiet)))
        .toEqual([1, 3, 0, 2])
    }
  })
  it('sort A then B derives ties from immutable source rather than stacking A', async () => {
    const rows = [[2, 1], [0, 1], [1, 1], [0, 0]]
    const store = dataStore(rows.map(row => row.map(value => fixtureCell(value))))
    const a = rowsOf(await sortRowOrder(store, specFor(4), quiet))
    const b = rowsOf(await sortRowOrder(store, specFor(4, [{ col: 1, order: 'asc' }]), quiet))
    const stacked = [...a].sort((left, right) => rows[left][1] - rows[right][1])
    expect(a).toEqual([1, 3, 2, 0])
    expect(b).toEqual([3, 0, 1, 2])
    expect(b).not.toEqual(stacked)
    expect(rows).toEqual([[2, 1], [0, 1], [1, 1], [0, 0]])
  })
  it.each(['asc', 'desc'] as const)('mixed multi-key %s pins hidden slots and keeps terminal categories', async order => {
    const rows = [
      [fixtureCell(1), fixtureCell('b')], [fixtureCell(1), fixtureCell('a')],
      [fixtureCell(1), fixtureCell('a')], [fixtureCell(false), fixtureCell(1)],
      [fixtureCell('2'), fixtureCell(true)], [fixtureCell('#NUM!', true), fixtureCell('b')],
      [fixtureCell('#NUM!', true), fixtureCell('a')], [fixtureCell(null), fixtureCell(0)],
      [fixtureCell(1), fixtureCell(null)], [fixtureCell(1), fixtureCell('#Z!', true)],
    ]
    const result = await sortRowOrder(dataStore(rows), specFor(rows.length, [
      { col: 0, order }, { col: 1, order: 'desc' },
    ]), { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), createRowVisibility(MAX_ROWS, row => row !== 2)) })
    expect(rowsOf(result)).toEqual(order === 'asc'
      ? [0, 1, 2, 9, 8, 4, 3, 5, 6, 7] : [3, 4, 2, 0, 1, 9, 8, 5, 6, 7])
  })
  it('genuine typed allocation RangeErrors remain nonretryable resource failures', async () => {
    vi.stubGlobal('Uint8Array', class {
      constructor() { throw new RangeError('allocation exhausted') }
    })
    try {
      expect(await prepareSortKeys(dataStore([[fixtureCell(1)]]), specFor(1), quiet))
        .toMatchObject({ kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'allocation exhausted' } })
    } finally { vi.unstubAllGlobals() }
  })
  it('unsupported object values still refuse without a draft', async () => {
    const store = dataStore([[fixtureCell(1)]])
    const original = store.queryColumn.bind(store)
    const invalid: SourceStore = { ...store, queryColumn: async (request, options) => {
      const page = await original(request, options)
      return { ...page, items: page.items.map(item => item.kind === 'cell'
        ? { ...item, cell: { ...item.cell, value: {} as unknown as typeof item.cell.value } } : item) }
    } }
    expect(await sortRowOrder(invalid, specFor(1), quiet))
      .toMatchObject({ kind: 'error', error: { code: 'page-error' } })
  })
  it('seeded random maximum-row numeric permutation matches native typed-index oracle', async () => {
    const random = createPrng(20261010)
    const values = Float64Array.from({ length: MAX_ROWS }, (_v, row) => row)
    for (let row = MAX_ROWS - 1; row > 0; row--) {
      const at = Math.floor(random() * (row + 1)), value = values[row]
      values[row] = values[at]
      values[at] = value
    }
    const oracle = Uint32Array.from({ length: MAX_ROWS }, (_v, row) => row)
      .sort((a, b) => values[a] - values[b] || a - b)
    const store = createFixture('F2').createStore({ read: {
      isPopulated: (_row, col) => col === 0,
      cellAt: row => fixtureCell(values[row]),
    } })
    const result = await sortRowOrder(store, specFor(MAX_ROWS), { ...quiet, now: () => 0 })
    expect(result.kind).toBe('complete')
    if (result.kind !== 'complete') return
    const band = result.order.bands[0], ledger = result.ledger
    expect(band.sourceAtOffset).toBeInstanceOf(Uint32Array)
    expect(band.displayOffsetOfSource).toBeInstanceOf(Uint32Array)
    expect(ledger.tagsBytes).toBe(MAX_ROWS)
    expect(ledger.numbersBytes).toBe(8 * MAX_ROWS)
    expect(ledger.idsBytes + ledger.ranksBytes + ledger.dictionaryBytes).toBe(0)
    expect(ledger.indexBytes).toBe(4 * MAX_ROWS)
    expect(ledger.scratchBytes).toBe(4 * MAX_ROWS)
    expect(ledger.orderBytes + ledger.inverseBytes).toBe(8 * MAX_ROWS)
    expect(ledger.retainedBytes).toBe(8 * MAX_ROWS + 256)
    expect(ledger.aggregatePeakBytes).toBe(ledger.peakBytes + ledger.residentWorkerBytes)
    expect(ledger.peakBytes).toBeGreaterThan(ledger.retainedBytes)
    expect(ledger.peakBytes).toBeLessThanOrEqual(512 * 1024 * 1024)
    for (let slot = 0; slot < MAX_ROWS; slot++) {
      if (band.sourceAtOffset[slot] !== oracle[slot]
        || band.displayOffsetOfSource[oracle[slot]] !== slot) {
        throw new Error(`Random typed oracle mismatch at ${slot}`)
      }
    }
  }, 30000)
})


describe('literal boundary regressions', () => {
  it('bounded helper visibility RangeErrors are source failures, not allocations', async () => {
    let calls = 0
    const base = createRowVisibility(MAX_ROWS)
    const visibility = {
      visible: () => { if (++calls === 3) throw new RangeError('visibility generation changed'); return true },
      rank: base.rank.bind(base), select: base.select.bind(base), next: base.next.bind(base),
    }
    expect(await sortRowOrder(dataStore([[fixtureCell(2)], [fixtureCell(1)]]), specFor(2),
      { ...quiet, hidden: hiddenSourceRowsFromDisplay(createIdentityOrder(), visibility) })).toMatchObject({
      kind: 'error', error: { code: 'page-error', message: 'visibility generation changed' },
    })
    expect(calls).toBe(3)
  })
  it('reducer safely refuses getters throwing an object whose conversion throws', () => {
    const thrown = { toString: () => { throw new Error('conversion must not run') } }
    const sort = specFor(4)
    const command = { get kind() { throw thrown } } as unknown as Parameters<typeof reduceSortCommand>[1]
    expect(() => reduceSortCommand(sort, command)).not.toThrow()
    expect(reduceSortCommand(sort, command)).toMatchObject({
      sort, notice: { reason: 'invalid-sort-command', error: { code: 'invalid-source-rect', retryable: false } },
    })
    const draft = { get range() { throw thrown } } as unknown as SortSpec
    expect(reduceSortCommand(draft, { kind: 'toggle-case' }).sort).toBe(draft)
  })
})


describe('round three source ownership regressions', () => {
  it('translates current display visibility before replacing a filtered sort', async () => {
    const values = [3, 1, 2], store = dataStore(values.map(value => [fixtureCell(value)]))
    const first = await sortRowOrder(store, specFor(3, [{ col: 0, order: 'desc' }]), quiet)
    expect(first.kind).toBe('complete')
    if (first.kind !== 'complete') return
    expect(rowsOf(first)).toEqual([0, 2, 1])
    const hidden = hiddenSourceRowsFromDisplay(first.order, createRowVisibility(MAX_ROWS, slot => slot !== 1))
    expect([0, 1, 2].map(row => hidden.hidden(row))).toEqual([false, false, true])
    const replacement = await sortRowOrder(store, specFor(3), { ...quiet, hidden })
    expect(rowsOf(replacement).map(row => values[row])).toEqual([1, 3, 2])
  })
  it('translates identity and a nonzero permutation band including outside rows', () => {
    const display = createRowVisibility(MAX_ROWS, slot => slot !== 6 && slot !== 2)
    const identity = hiddenSourceRowsFromDisplay(createIdentityOrder(), display)
    expect([2, 5, 6, 7].map(row => identity.hidden(row))).toEqual([true, false, true, false])
    const order = createRowOrder([{ r0: 5, r1: 7, sourceAtOffset: new Uint32Array([7, 5, 6]) }])
    const hidden = hiddenSourceRowsFromDisplay(order, display)
    expect([2, 5, 6, 7].map(row => hidden.hidden(row))).toEqual([true, true, false, false])
  })
  it('retains no key, fixed or metadata reservation for a no-work empty identity', async () => {
    const spec = { ...specFor(1), range: { ...specFor(1).range, headerRows: 1 as const } }
    const result = await sortRowOrder(dataStore([[fixtureCell(1)]]), spec, quiet)
    expect(result.kind).toBe('complete')
    if (result.kind === 'complete') expect(result.ledger.retainedBytes).toBe(0)
  })
  it('refuses a present legacy visibility key instead of silently sorting unpinned', async () => {
    const store = dataStore([[fixtureCell(3)], [fixtureCell(1)], [fixtureCell(2)]])
    const legacy = { ...quiet, visibility: createRowVisibility(3, slot => slot !== 1) } as unknown as SortOptions
    for (const run of [prepareSortKeys, sortRowOrder]) {
      expect(await run(store, specFor(3), legacy)).toMatchObject({
        kind: 'error', error: { code: 'unsupported', retryable: false },
      })
    }
    expect('order' in await sortRowOrder(store, specFor(3), legacy)).toBe(false)
  })
  it('ignores an explicitly undefined legacy visibility key', async () => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const options = { ...quiet, visibility: undefined } as unknown as SortOptions
    expect(rowsOf(await sortRowOrder(store, specFor(2), options))).toEqual([1, 0])
  })
  it('unknown source throws are nonretryable page failures', async () => {
    const store = dataStore([[fixtureCell(1)]])
    for (const error of [new Error('disk failed'), new RangeError('decoder bounds')]) {
      expect(await sortRowOrder({ ...store, queryColumn: async () => { throw error } }, specFor(1), quiet))
        .toMatchObject({ kind: 'error', error: { code: 'page-error', retryable: false } })
    }
  })
  it('bounded helper copy allocation failures are resource limits', async () => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const copy = vi.spyOn(Uint32Array.prototype, 'slice').mockImplementation(() => { throw new RangeError('order copy exhausted') })
    try {
      expect(await sortRowOrder(store, specFor(2), quiet)).toMatchObject({
        kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'order copy exhausted' },
      })
    } finally { copy.mockRestore() }
  })
  it.each([3, 4, 5])('bounded helper typed allocation %i is a resource limit', async failAt => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const TypedRows = Uint32Array
    let allocations = 0
    vi.stubGlobal('Uint32Array', new Proxy(TypedRows, {
      construct(target, args) {
        if (args[0] === 2 && ++allocations === failAt) throw new RangeError('helper buffer exhausted')
        return Reflect.construct(target, args)
      },
    }))
    try {
      expect(await sortRowOrder(store, specFor(2), quiet)).toMatchObject({
        kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'helper buffer exhausted' },
      })
      expect(allocations).toBe(failAt)
    } finally { vi.unstubAllGlobals() }
  })
  it('second order helper copy allocation is also a resource limit', async () => {
    const store = dataStore([[fixtureCell(2)], [fixtureCell(1)]])
    const original = Uint32Array.prototype.slice
    let copies = 0
    const copy = vi.spyOn(Uint32Array.prototype, 'slice').mockImplementation(function (this: Uint32Array, start?: number, end?: number) {
      if (++copies === 2) throw new RangeError('published order copy exhausted')
      return original.call(this, start, end)
    })
    try {
      expect(await sortRowOrder(store, specFor(2), quiet)).toMatchObject({
        kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'published order copy exhausted' },
      })
      expect(copies).toBe(2)
    } finally { copy.mockRestore() }
  })
  it('dictionary Map growth allocation failures are resource limits', async () => {
    const store = dataStore([[fixtureCell('item')]])
    const original = Map.prototype.set
    const growth = vi.spyOn(Map.prototype, 'set').mockImplementation(function (this: Map<unknown, unknown>, key: unknown, value: unknown) {
      if (key === 'item') throw new RangeError('dictionary exhausted')
      return original.call(this, key, value)
    })
    try {
      expect(await prepareSortKeys(store, specFor(1), quiet)).toMatchObject({
        kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'dictionary exhausted' },
      })
    } finally { growth.mockRestore() }
  })
  it('spec getter RangeErrors stay validation page failures', async () => {
    let reads = 0
    const spec = { ...specFor(1), keys: [{
      get col() { if (++reads === 5) throw new RangeError('key getter failed'); return 0 }, order: 'asc' as const,
    }] }
    expect(await sortRowOrder(dataStore([[fixtureCell(1)]]), spec, quiet)).toMatchObject({
      kind: 'error', error: { code: 'page-error', retryable: false, message: 'key getter failed' },
    })
  })
  it('owned specification key growth allocation failures are resource limits', async () => {
    const store = dataStore([[fixtureCell(1)]])
    const spec = specFor(1), original = Array.prototype.push
    Array.prototype.push = function (this: unknown[], ...items: unknown[]) {
      const item = items[0]
      if (items.length === 1 && item !== null && typeof item === 'object'
        && Object.isFrozen(item) && Object.keys(item).length === 2
        && 'col' in item && 'order' in item && item.col === 0 && item.order === 'asc') {
        throw new RangeError('owned sort keys exhausted')
      }
      return original.apply(this, items)
    }
    try {
      expect(await sortRowOrder(store, spec, quiet)).toMatchObject({
        kind: 'error', error: { code: 'resource-limit', retryable: false, message: 'owned sort keys exhausted' },
      })
    } finally { Array.prototype.push = original }
  })
  it.each([false, true])('hostile prototype probes still resolve to a typed nonretryable failure (always hostile: %s)', async always => {
    const store = dataStore([[fixtureCell(1)]])
    let probes = 0
    const thrown = new Proxy({}, { getPrototypeOf() {
      if (always || ++probes === 1) throw new Error('prototype forbidden')
      return Object.prototype
    } })
    await expect(sortRowOrder({ ...store, queryColumn: async () => { throw thrown } }, specFor(1), quiet))
      .resolves.toMatchObject({
        kind: 'error', error: { code: 'page-error', retryable: false, message: 'Sort operation failed' },
      })
  })
  it('hostile unknown throws still return a typed nonretryable failure', async () => {
    const store = dataStore([[fixtureCell(1)]])
    const error = { toString: () => { throw new Error('conversion forbidden') } }
    expect(await sortRowOrder({ ...store, queryColumn: async () => { throw error } }, specFor(1), quiet))
      .toMatchObject({ kind: 'error', error: { code: 'page-error', retryable: false, message: 'Sort operation failed' } })
  })
  it('Node immediate is preferred when a native scheduler also exists', async () => {
    let immediate = 0, native = 0
    vi.stubGlobal('setImmediate', (resume: () => void) => { immediate++; resume() })
    vi.stubGlobal('scheduler', { yield: async () => { native++ } })
    try {
      expect((await sortRowOrder(dataStore([[fixtureCell(1)]]), specFor(1))).kind).toBe('complete')
      expect(immediate).toBeGreaterThan(0)
      expect(native).toBe(0)
    } finally { vi.unstubAllGlobals() }
  })
  it('browser native scheduler is preferred over MessageChannel', async () => {
    let yielded = 0
    vi.stubGlobal('setImmediate', undefined)
    vi.stubGlobal('scheduler', { yield: async () => { yielded++ } })
    vi.stubGlobal('MessageChannel', class { constructor() { throw new Error('ports should not be created') } })
    try {
      expect((await sortRowOrder(dataStore([[fixtureCell(1)]]), specFor(1))).kind).toBe('complete')
      expect(yielded).toBeGreaterThan(0)
    } finally { vi.unstubAllGlobals() }
  })
})

// Compile-only consumer contract: public lanes never grant write access or
// methods returning the worker's mutable backing storage.
function preparedLaneTypeContract(lane: import('../src/xlsx/view/sort').SortLane): void {
  // @ts-expect-error The buffer reference is immutable too.
  lane.tags = new Uint8Array(1)
  // @ts-expect-error Element writes corrupt the resident sort keys.
  lane.tags[0] = 1
  // @ts-expect-error Numeric values are readonly too.
  lane.numbers[0] = 2
  // @ts-expect-error IDs are readonly too.
  lane.ids![0] = 0
  // @ts-expect-error Ranks are readonly too.
  lane.stringRanks[0] = 0
  // @ts-expect-error Error ranks are readonly too.
  lane.errorRanks[0] = 0
  // @ts-expect-error Mutating typed array methods are not public.
  lane.tags.fill(0)
  // @ts-expect-error Subarray views must not escape with mutable elements.
  lane.tags.subarray(0)[0] = 0
  // @ts-expect-error A writable backing ArrayBuffer must not escape.
  new Uint8Array(lane.tags.buffer)[0] = 0
  // @ts-expect-error DISPLAY-space masks cannot be passed to the old raw option.
  const raw: SortOptions = { visibility: createRowVisibility(MAX_ROWS) }
  // @ts-expect-error A DISPLAY visibility object is not branded SOURCE hidden rows.
  const wrongSpace: SortOptions = { hidden: createRowVisibility(MAX_ROWS) }
  // @ts-expect-error An unbranded predicate must be explicitly translated.
  const unbranded: SortOptions = { hidden: { hidden: () => false } }
  const inferredLegacy = { yield: async () => {}, visibility: createRowVisibility(MAX_ROWS) }
  // @ts-expect-error Legacy options are forbidden through inferred variables too.
  void sortRowOrder(dataStore([]), specFor(1), inferredLegacy)
  void raw
  void wrongSpace
  void unbranded
}
void preparedLaneTypeContract

describe('ownership-aware string accounting (S8.4)', () => {
  const values = ['item-10', 'item-2', 'beta-long-'.repeat(20), 'item-2', 'alpha']
  const rows = values.map(v => Object.freeze([fixtureCell(v)])).concat([Object.freeze([fixtureCell('#N/A', true)])])
  const unique = [...new Set(values)]

  it('charges each unique text key once in the dictionary unless the source owns the strings', async () => {
    const store = dataStore(Object.freeze(rows))
    const own = await sortRowOrder(store, specFor(rows.length), { ...quiet, sourceRetainsStrings: true })
    const copy = await sortRowOrder(store, specFor(rows.length), quiet)
    // Ordering is identical; only the ledger differs.
    expect(rowsOf(own)).toEqual(rowsOf(copy))
    if (own.kind !== 'complete' || copy.kind !== 'complete') throw new Error('incomplete')
    const errorBytes = '#N/A'.length * 2 + 40
    expect(copy.ledger.dictionaryBytes).toBe(unique.reduce((sum, v) => sum + v.length * 2 + 40, 0) + errorBytes)
    // The source owns the text keys: only the 16 B reference slot each; error texts keep the full charge.
    expect(own.ledger.dictionaryBytes).toBe(unique.length * 16 + errorBytes)
    expect(own.ledger.peakBytes).toBeLessThan(copy.ledger.peakBytes)
  })

  it('admits a sort the double-counting ledger would refuse, at the same budget', async () => {
    const store = dataStore(Object.freeze(rows))
    const copy = await sortRowOrder(store, specFor(rows.length), quiet)
    if (copy.kind !== 'complete') throw new Error('incomplete')
    const limit = copy.ledger.peakBytes - 1
    expect((await sortRowOrder(store, specFor(rows.length), { ...quiet, maxWorkerBytes: limit })).kind).toBe('resource-limit')
    expect((await sortRowOrder(store, specFor(rows.length), { ...quiet, maxWorkerBytes: limit, sourceRetainsStrings: true })).kind).toBe('complete')
  })
})
