import { describe, expect, it } from 'vitest'
import type { FilterRange, SourceStamp, SourceStore, ColumnEntry, Page, ValuesFilter } from '../src/xlsx/view/contracts'
import { FilterEngine, rawValueKey, type ColumnClassification, type PreparedValue } from '../src/xlsx/view/filter'
import {
  SOURCE_BLANK, SOURCE_BOOLEAN, SOURCE_ERROR, SOURCE_NUMBER, SOURCE_STRING, SOURCE_UNAVAILABLE, VALUE_DISPLAY_LIMIT,
  ValueIndexCache, ValueIndexError, ValueSelectionDraft, compileValuesMembership, valueSourceFromStore,
  type ValueColumnSource, type ValueIndex, type ValueIndexIdentity, type ValueLabeler, type ValueLanePage, type ValuePageRead,
} from '../src/xlsx/view/valueIndex'
import { MemoryValueRunStore } from '../src/xlsx/view/valueRuns'

const STAMP: SourceStamp = { docId: 'doc', sheetId: 's1', valueGen: 3 }
const RANGE: FilterRange = { r0: 0, r1: 1_048_575, c0: 0, c1: 3, headerRows: 1, totalsRows: 0 }
const mixed: ColumnClassification = { kind: 'text', hasNumeric: true }
const numeric: ColumnClassification = { kind: 'number', hasNumeric: true }
const date: ColumnClassification = { kind: 'date', hasNumeric: true }
const engine = new FilterEngine({ locale: 'en-US', dateSystem: '1900', timeZone: 'UTC', nowMs: 0 })
const label: ValueLabeler = (kind, value) => kind === 'boolean' ? (value ? 'TRUE' : 'FALSE')
  : kind === 'date' ? new Date(Math.round((value - 25569) * 86400000)).toISOString().slice(0, 10) : String(value)

type Cell = number | string | boolean | { error: string } | { unavailable: true } | null
function lanes(pageId: string, cells: readonly Cell[]): ValueLanePage {
  const tags = new Uint8Array(cells.length), numbers = new Float64Array(cells.length), texts: (string | undefined)[] = []
  cells.forEach((cell, i) => {
    if (cell === null) tags[i] = SOURCE_BLANK
    else if (typeof cell === 'number') { tags[i] = SOURCE_NUMBER; numbers[i] = cell }
    else if (typeof cell === 'string') { tags[i] = SOURCE_STRING; texts[i] = cell }
    else if (typeof cell === 'boolean') { tags[i] = SOURCE_BOOLEAN; numbers[i] = cell ? 1 : 0 }
    else if ('error' in cell) { tags[i] = SOURCE_ERROR; texts[i] = cell.error }
    else tags[i] = SOURCE_UNAVAILABLE
  })
  return { kind: 'values', pageId, count: cells.length, tags, numbers, texts }
}
class FakeSource implements ValueColumnSource {
  readonly reads: number[] = []
  stamp: SourceStamp = STAMP
  constructor(readonly pages: (() => ValuePageRead)[], readonly onRead?: (page: number) => void) {}
  get pageCount(): number { return this.pages.length }
  async readPage(page: number, _signal?: AbortSignal): Promise<ValuePageRead> {
    this.reads.push(page)
    this.onRead?.(page)
    return this.pages[page]()
  }
}
function fromCells(cells: readonly Cell[], pageRows = 4, onRead?: (page: number) => void): FakeSource {
  const pages: (() => ValuePageRead)[] = []
  for (let start = 0; start < cells.length; start += pageRows) {
    const slice = cells.slice(start, start + pageRows), id = `p${pages.length}`
    pages.push(() => lanes(id, slice))
  }
  return new FakeSource(pages, onRead)
}
const identity = (classification: ColumnClassification = mixed, extra: Partial<ValueIndexIdentity> = {}): ValueIndexIdentity =>
  ({ ...STAMP, range: RANGE, col: 2, classification, classificationGen: 1, ...extra })
async function built(cells: readonly Cell[], classification: ColumnClassification = mixed, pageRows = 4) {
  const cache = new ValueIndexCache(), source = fromCells(cells, pageRows)
  const index = cache.open(identity(classification), source)
  await index.build()
  return { cache, source, index }
}
const keysOf = (index: ValueIndex): string[] => Array.from({ length: index.size }, (_, i) => index.key(i))
const str = (value: string): PreparedValue => ({ kind: 'string', value })
const num = (value: number): PreparedValue => ({ kind: 'number', value })
async function membership(rule: ValuesFilter | null, store: MemoryValueRunStore, classification: ColumnClassification = mixed) {
  if (!rule) return () => true
  return compileValuesMembership(rule, classification, store, STAMP)
}
const items = (count: number): string[] => Array.from({ length: count }, (_, i) => `item-${String(i).padStart(5, '0')}`)

describe('semantic value keys', () => {
  it('uses exact type-tagged raw keys with blank separate and explicit error identity', async () => {
    const cells: Cell[] = [1.5, 'foo', true, false, { error: '#DIV/0!' }, { error: '#N/A' }, '#N/A', 'TRUE', '1.5', -0, 0, null, 1.5, 'foo', null, { error: '#N/A' }]
    const { index } = await built(cells)
    const oracle = new Set(cells.filter(cell => cell !== null).map(cell => rawValueKey(
      typeof cell === 'number' ? num(cell) : typeof cell === 'string' ? str(cell) : typeof cell === 'boolean' ? { kind: 'boolean', value: cell }
        : { kind: 'error', value: (cell as { error: string }).error }, mixed)))
    expect(new Set(keysOf(index))).toEqual(oracle)
    expect(index.size).toBe(10)
    expect(index.blankCount).toBe(2)
    expect(keysOf(index)).toEqual(expect.arrayContaining(['e:#N/A', 'e:#DIV/0!', 's:#N/A', 'b:1', 's:TRUE', 'n:1.5', 's:1.5', 'n:0']))
    expect(keysOf(index)).not.toContain('n:-0')
    const page = index.page({ offset: 0, limit: 50, label })
    expect(page.entries.find(entry => entry.key === 'n:1.5')?.count).toBe(2)
    expect(page.entries.find(entry => entry.key === 'e:#N/A')).toMatchObject({ kind: 'error', label: '#N/A', count: 2 })
    expect(page.entries.find(entry => entry.key === 's:#N/A')).toMatchObject({ kind: 'string', count: 1 })
  })
  it('keeps distinct date serials with the same display label as distinct keys', async () => {
    const cells: Cell[] = [44927.25, 44927.75, 44927.25, 44928]
    const { index } = await built(cells, date)
    expect(keysOf(index)).toEqual(['d:44927.25', 'd:44927.75', 'd:44928'])
    const entries = index.page({ offset: 0, limit: 10, label }).entries
    expect(entries.map(entry => entry.label)).toEqual(['2023-01-01', '2023-01-01', '2023-01-02'])
    expect(entries.map(entry => entry.kind)).toEqual(['date', 'date', 'date'])
    expect(() => index.page({ offset: 0, limit: 10 })).toThrow(expect.objectContaining({ code: 'missing-display-text' }))
    await expect(index.search({ text: '2023', locale: 'en-US' })).rejects.toMatchObject({ code: 'missing-display-text' })
    const day = await index.search({ text: '2023-01-01', locale: 'en-US' }, { label })
    expect(index.page({ offset: 0, limit: 10, search: day, label }).entries.map(entry => entry.key)).toEqual(['d:44927.25', 'd:44927.75'])
  })
  it('keys the cache by classification so the same serials are numeric keys under number classification', async () => {
    const cache = new ValueIndexCache(), source = fromCells([44927.25, 44927.75])
    const asDate = cache.open(identity(date), source), asNumber = cache.open(identity(numeric), source)
    expect(asDate).not.toBe(asNumber)
    await asDate.build(); await asNumber.build()
    expect(keysOf(asDate)).toEqual(['d:44927.25', 'd:44927.75'])
    expect(keysOf(asNumber)).toEqual(['n:44927.25', 'n:44927.75'])
    expect(cache.open(identity(date, { classificationGen: 2 }), source)).not.toBe(asDate)
    expect(cache.open(identity(date), source)).toBe(asDate)
  })
  it('keeps raw UTF-16 strings without collation folding', async () => {
    const { index } = await built(['A', 'a', 'é', 'e\u0301', '\ud800', '\ud800'])
    expect(keysOf(index)).toEqual(['s:A', 's:a', 's:é', 's:e\u0301', 's:\ud800'])
  })
})

describe('universe completeness', () => {
  it('never reports a universe with unknown pages complete and resumes only those pages', async () => {
    let resolved = false
    const source = new FakeSource([
      () => lanes('p0', ['a', 'b']),
      () => resolved ? lanes('p1', ['c', null]) : { kind: 'unknown', pageId: 'p1' },
      () => lanes('p2', ['d', 'a']),
    ])
    const cache = new ValueIndexCache(), index = cache.open(identity(), source)
    expect(await index.build()).toBe('loading')
    expect(index.completeness).toBe('loading')
    expect(keysOf(index)).toEqual(['s:a', 's:b', 's:d'])
    expect(index.page({ offset: 0, limit: 10 }).completeness).toBe('loading')
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    draft.toggle(0, false)
    await expect(draft.commit()).rejects.toMatchObject({ code: 'incomplete' })
    resolved = true
    expect(await cache.open(identity(), source).build()).toBe('complete')
    expect(source.reads).toEqual([0, 1, 2, 1])
    expect(keysOf(index)).toEqual(['s:a', 's:b', 's:d', 's:c'])
    expect(index.blankCount).toBe(1)
  })
  it('marks unavailable formula values incomplete and refuses commit', async () => {
    const { index } = await built(['a', { unavailable: true }, 'b'])
    expect(index.completeness).toBe('incomplete')
    expect(index.unavailableCount).toBe(1)
    const draft = await ValueSelectionDraft.open(index, null, new MemoryValueRunStore())
    draft.toggle(0, false)
    await expect(draft.commit()).rejects.toMatchObject({ code: 'incomplete' })
  })
  it('stays loading without source reads while classification is loading', async () => {
    const source = fromCells(['a'])
    const index = new ValueIndexCache().open(identity({ kind: 'loading', hasNumeric: false }), source)
    expect(await index.build()).toBe('loading')
    expect(source.reads).toEqual([])
  })
  it('publishes typed source failures and rejects a changed source owner', async () => {
    const failing = new FakeSource([() => ({ kind: 'error', failure: { code: 'page-error', retryable: true, message: 'boom' } })])
    await expect(new ValueIndexCache().open(identity(), failing).build()).rejects.toMatchObject({ code: 'page-error', detail: { retryable: true } })
    const stale = fromCells(['a', 'b', 'c', 'd', 'e'], 2, () => { stale.stamp = { ...STAMP, valueGen: 4 } })
    await expect(new ValueIndexCache().open(identity(), stale).build()).rejects.toMatchObject({ code: 'stale-source' })
    const uncovered = new FakeSource([() => ({ ...lanes('p0', ['a']), tags: new Uint8Array(1) })])
    await expect(new ValueIndexCache().open(identity(), uncovered).build()).rejects.toMatchObject({ code: 'page-error' })
  })
})

describe('draft membership semantics', () => {
  const cells: Cell[] = [...items(30), null]
  it('keeps an untouched OK unrestricted without writing exclusions', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    expect(draft.selectedCount()).toEqual({ count: 30, complete: true })
    expect(await draft.commit()).toEqual({ kind: 'unchanged', rule: null })
    expect(store.puts).toBe(0)
  })
  it('represents unchecking one entry as all-minus-exclusions', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    draft.toggle(index.find(str('item-00007')), false)
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ kind: 'values', showBlanks: true, selection: { mode: 'allExcept', excluded: { count: 1, owner: STAMP } } })
    const matches = await membership(rule, store)
    expect(matches(str('item-00007'))).toBe(false)
    expect(matches(str('item-00008'))).toBe(true)
    expect(matches({ kind: 'blank' })).toBe(true)
  })
  it('represents a small selection as explicit membership', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    draft.selectAll(false)
    draft.toggle(index.find(str('item-00002')), true)
    draft.toggle(index.find(str('item-00029')), true)
    expect(draft.blanks).toBe(false)
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ showBlanks: false, selection: { mode: 'include', keys: { count: 2 } } })
    const matches = await membership(rule, store)
    expect([str('item-00002'), str('item-00029'), str('item-00003'), { kind: 'blank' } as PreparedValue].map(matches)).toEqual([true, true, false, false])
  })
  it('normalizes a draft equivalent to all values and blanks away', async () => {
    const { index } = await built(cells)
    const draft = await ValueSelectionDraft.open(index, null, new MemoryValueRunStore())
    draft.selectAll(false); draft.selectAll(true)
    expect(await draft.commit()).toEqual({ kind: 'committed', rule: null })
    const noBlanks = await built(items(3))
    const other = await ValueSelectionDraft.open(noBlanks.index, null, new MemoryValueRunStore())
    other.setBlanks(false)
    expect(await other.commit()).toEqual({ kind: 'committed', rule: null })
  })
  it('represents blanks separately from value keys', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    draft.setBlanks(false)
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ showBlanks: false, selection: { mode: 'allExcept', excluded: { count: 0 } } })
    const matches = await membership(rule, store)
    expect(matches({ kind: 'blank' })).toBe(false)
    expect(matches(str('item-00000'))).toBe(true)
    expect(() => matches({ kind: 'unknown' })).toThrow(/incomplete/)
  })
  it('cancel closes the draft without changing the committed rule and reopen shows the applied rule', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore(), first = await ValueSelectionDraft.open(index, null, store)
    first.toggle(index.find(str('item-00004')), false)
    const committed = (await first.commit()).rule!
    const draft = await ValueSelectionDraft.open(index, committed, store)
    expect(draft.isSelected(index.find(str('item-00004')))).toBe(false)
    expect(draft.isSelected(index.find(str('item-00005')))).toBe(true)
    const puts = store.puts
    draft.selectAll(false)
    draft.cancel()
    expect(draft.closed).toBe(true)
    expect(store.puts).toBe(puts)
    expect(() => draft.toggle(0, true)).toThrow(ValueIndexError)
    await expect(draft.commit()).rejects.toMatchObject({ code: 'closed' })
    const reopened = await ValueSelectionDraft.open(index, committed, store)
    expect(reopened.isSelected(index.find(str('item-00004')))).toBe(false)
    expect(reopened.selectedCount()).toEqual({ count: 29, complete: true })
    expect(await reopened.commit()).toEqual({ kind: 'unchanged', rule: committed })
    expect(store.puts).toBe(puts)
  })
  it('rejects a committed rule owned by another source generation', async () => {
    const { index } = await built(cells)
    const store = new MemoryValueRunStore()
    const ref = await store.put({ count: 0, kinds: new Uint8Array(0), numbers: new Float64Array(0), texts: [] }, { ...STAMP, valueGen: 2 })
    const stale: ValuesFilter = { kind: 'values', showBlanks: true, selection: { mode: 'allExcept', excluded: ref } }
    await expect(ValueSelectionDraft.open(index, stale, store)).rejects.toMatchObject({ code: 'stale-source' })
  })
})

describe('complete universe beyond the displayed prefix', () => {
  const values = items(15_000)
  it('limits displayed entries only, not construction', async () => {
    const { index } = await built(values, mixed, 1024)
    expect(index.size).toBe(15_000)
    expect(index.completeness).toBe('complete')
    const first = index.page({ offset: 0, limit: 500 })
    expect(first).toMatchObject({ total: 15_000, displayed: VALUE_DISPLAY_LIMIT, truncated: true })
    expect(first.entries).toHaveLength(500)
    expect(index.page({ offset: 9_900, limit: 500 }).entries).toHaveLength(100)
    expect(index.page({ offset: VALUE_DISPLAY_LIMIT, limit: 10 }).entries).toHaveLength(0)
    expect(index.key(14_999)).toBe('s:item-14999')
  })
  it('searches all values with S8.1 text semantics, including matches after the first 10k', async () => {
    const { index } = await built([...values, 'ÍTEM-ACCENT', 'item-accent'], mixed, 1024)
    const hit = await index.search({ text: 'ITEM-12345', locale: 'en-US' })
    expect(hit.count).toBe(1)
    expect(hit.complete).toBe(true)
    expect(index.page({ offset: 0, limit: 10, search: hit }).entries.map(entry => entry.key)).toEqual(['s:item-12345'])
    const accent = await index.search({ text: 'item-accent', locale: 'en-US' })
    expect(index.page({ offset: 0, limit: 10, search: accent }).entries.map(entry => entry.key)).toEqual(['s:ÍTEM-ACCENT', 's:item-accent'].filter(key => engine.matchText('contains', 'item-accent', false, key.slice(2))))
    const subset = await index.search({ text: 'item-1', locale: 'en-US' })
    expect(subset.count).toBe(5_000)
    expect(index.page({ offset: 0, limit: 500, search: subset })).toMatchObject({ total: 5_000, displayed: 5_000, truncated: false })
  })
  it('applies searched Select All to exactly the searched subset', async () => {
    const { index } = await built(values, mixed, 1024)
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    const subset = await draft.search({ text: 'item-1', locale: 'en-US' })
    draft.selectAll(false)
    draft.selectAll(true, subset)
    expect(draft.selectedCount().count).toBe(5_000)
    expect(draft.describe()).toEqual({ base: 'none', blanks: false, patches: [{ kind: 'search', selected: true, count: 5_000 }] })
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ selection: { mode: 'include', keys: { count: 5_000 } } })
    const matches = await membership(rule, store)
    expect(matches(str('item-14999'))).toBe(true)
    expect(matches(str('item-09999'))).toBe(false)
    const deselect = await ValueSelectionDraft.open(index, null, store)
    deselect.selectAll(false, await deselect.search({ text: 'item-1', locale: 'en-US' }))
    expect(deselect.blanks).toBe(true)
    const excluded = (await deselect.commit()).rule
    expect(excluded).toMatchObject({ selection: { mode: 'allExcept', excluded: { count: 5_000 } } })
    expect((await membership(excluded, store))(str('item-00001'))).toBe(true)
  })
  it('lets the last matching patch win for individual entries and searched subsets', async () => {
    const { index } = await built(values, mixed, 1024)
    const draft = await ValueSelectionDraft.open(index, null, new MemoryValueRunStore())
    const target = index.find(str('item-12345'))
    draft.toggle(target, false)
    expect(draft.isSelected(target)).toBe(false)
    draft.selectAll(true, await draft.search({ text: 'item-123', locale: 'en-US' }))
    expect(draft.isSelected(target)).toBe(true)
    draft.toggle(target, false)
    expect(draft.isSelected(target)).toBe(false)
    expect(draft.selectedCount().count).toBe(14_999)
  })
  it('keeps undisplayed values selected on an unchanged OK of an existing rule', async () => {
    const { index } = await built(values, mixed, 1024)
    const store = new MemoryValueRunStore(), first = await ValueSelectionDraft.open(index, null, store)
    first.toggle(index.find(str('item-00003')), false)
    const committed = (await first.commit()).rule!
    const reopened = await ValueSelectionDraft.open(index, committed, store)
    await reopened.search({ text: 'item-0', locale: 'en-US' })
    const result = await reopened.commit()
    expect(result).toEqual({ kind: 'unchanged', rule: committed })
    const matches = await membership(result.rule, store)
    expect(matches(str('item-14999'))).toBe(true)
    expect(matches(str('item-00003'))).toBe(false)
  })
  it('extends a search made while loading over values merged later before commit', async () => {
    let resolved = false
    const source = new FakeSource([
      () => lanes('p0', ['alpha-1', 'beta-1']),
      () => resolved ? lanes('p1', ['alpha-2', 'beta-2']) : { kind: 'unknown', pageId: 'p1' },
    ])
    const cache = new ValueIndexCache(), index = cache.open(identity(), source)
    await index.build()
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    const partial = await draft.search({ text: 'alpha', locale: 'en-US' })
    expect(partial).toMatchObject({ count: 1, complete: false })
    draft.selectAll(false, partial)
    resolved = true
    await index.build()
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ selection: { mode: 'allExcept', excluded: { count: 2 } } })
    const matches = await membership(rule, store)
    expect([str('alpha-1'), str('alpha-2'), str('beta-2')].map(matches)).toEqual([false, false, true])
  })
})

describe('scale with typed lanes', () => {
  it('indexes and searches 100k distinct mixed values', async () => {
    const cells: Cell[] = Array.from({ length: 100_000 }, (_, i) => i % 2 ? i * 1.5 : `text-${i}`)
    cells.push(true, false, { error: '#N/A' }, { error: '#DIV/0!' }, null)
    const { index } = await built(cells, mixed, 8192)
    expect(index.size).toBe(100_004)
    expect(index.completeness).toBe('complete')
    expect(index.find(num(99_999 * 1.5))).toBeGreaterThan(VALUE_DISPLAY_LIMIT)
    const started = performance.now()
    const hit = await index.search({ text: 'text-99998', locale: 'en-US' }, { label })
    const elapsed = performance.now() - started
    expect(hit.count).toBe(1)
    expect(index.page({ offset: 0, limit: 5, search: hit, label }).entries[0].key).toBe('s:text-99998')
    expect(elapsed).toBeGreaterThan(0)
  }, 120_000)
  it('indexes 1M distinct numeric values with bounded typed allocations and a warm reopen', async () => {
    const total = 1_000_000, rows = 65_536, pages: (() => ValuePageRead)[] = []
    for (let start = 0; start < total; start += rows) {
      const count = Math.min(rows, total - start), id = `p${pages.length}`
      pages.push(() => {
        const numbers = new Float64Array(count)
        for (let i = 0; i < count; i++) numbers[i] = start + i + 0.5
        return { kind: 'values', pageId: id, count, tags: new Uint8Array(count).fill(SOURCE_NUMBER), numbers, texts: [] }
      })
    }
    const source = new FakeSource(pages), cache = new ValueIndexCache()
    const index = cache.open(identity(numeric), source)
    expect(await index.build()).toBe('complete')
    expect(index.size).toBe(total)
    expect(index.key(total - 1)).toBe('n:999999.5')
    expect(index.ledger.sourcePageReads).toBe(pages.length)
    expect(index.ledger.allocations).toBeLessThan(200)
    expect(index.ledger.chargedBytes).toBeGreaterThanOrEqual(total * 17)
    expect(cache.usedBytes).toBe(index.ledger.chargedBytes)
    const cold = { ...index.ledger }
    const warm = cache.open(identity(numeric), source)
    expect(warm).toBe(index)
    expect(await warm.build()).toBe('complete')
    expect(warm.ledger).toEqual(cold)
    expect(index.page({ offset: 0, limit: 500, label })).toMatchObject({ total, displayed: VALUE_DISPLAY_LIMIT, truncated: true })
    const store = new MemoryValueRunStore(), draft = await ValueSelectionDraft.open(index, null, store)
    draft.toggle(index.find(num(999_999.5)), false)
    const { rule } = await draft.commit()
    expect(rule).toMatchObject({ selection: { mode: 'allExcept', excluded: { count: 1 } } })
    const matches = await membership(rule, store, numeric)
    expect(matches(num(999_999.5))).toBe(false)
    expect(matches(num(500_000.5))).toBe(true)
    const include = await ValueSelectionDraft.open(index, null, store)
    include.selectAll(false)
    for (const value of [10.5, 600_000.5, 999_998.5]) include.toggle(index.find(num(value)), true)
    const selected = (await include.commit()).rule
    expect(selected).toMatchObject({ selection: { mode: 'include', keys: { count: 3 } } })
    expect((await membership(selected, store, numeric))(num(600_000.5))).toBe(true)
  }, 120_000)
})

describe('resources, cancellation and ownership', () => {
  it('fails memory admission with a typed quota error and drops the partial index', async () => {
    const cache = new ValueIndexCache({ limitBytes: 64 * 1024 })
    const index = cache.open(identity(), fromCells(items(20_000), 1024))
    const failure = await index.build().catch(error => error)
    expect(failure).toBeInstanceOf(ValueIndexError)
    expect(failure).toMatchObject({ code: 'quota', detail: { resource: 'memory', limitBytes: 64 * 1024 } })
    expect(failure.detail.requiredBytes).toBeGreaterThan(64 * 1024)
    expect(index.disposed).toBe(true)
    expect(cache.size).toBe(0)
    expect(cache.usedBytes).toBe(0)
  })
  it('evicts idle cached indexes before refusing admission', async () => {
    const cache = new ValueIndexCache({ limitBytes: 2 * 1024 * 1024 })
    const first = cache.open(identity(), fromCells(items(5_000), 1024))
    await first.build()
    const second = cache.open(identity(mixed, { col: 3 }), fromCells(items(5_000).map(value => value + 'x'), 1024))
    await second.build()
    const third = cache.open(identity(mixed, { col: 1 }), fromCells(items(5_000).map(value => value + 'y'), 1024))
    await third.build()
    expect(third.completeness).toBe('complete')
    expect(first.disposed).toBe(true)
    expect(cache.usedBytes).toBeLessThanOrEqual(2 * 1024 * 1024)
  })
  it('fails storage admission with a typed quota error and leaves the draft open', async () => {
    const { index } = await built(items(100), mixed, 32)
    const store = new MemoryValueRunStore({ limitBytes: 64 }), draft = await ValueSelectionDraft.open(index, null, store)
    draft.selectAll(false)
    for (let i = 0; i < 10; i++) draft.toggle(i, true)
    await expect(draft.commit()).rejects.toMatchObject({ code: 'quota', detail: { resource: 'storage', limitBytes: 64 } })
    expect(draft.closed).toBe(false)
    expect(store.refCount).toBe(0)
  })
  it('cancels a build between batches and resumes from valid cached pages', async () => {
    const controller = new AbortController()
    const source = fromCells(items(40), 4, page => { if (page === 2) controller.abort() })
    const cache = new ValueIndexCache(), index = cache.open(identity(), source)
    await expect(index.build({ signal: controller.signal })).rejects.toMatchObject({ code: 'cancelled' })
    expect(index.completeness).toBe('loading')
    expect(index.size).toBe(8)
    const reopened = cache.open(identity(), source)
    expect(reopened).toBe(index)
    expect(await reopened.build()).toBe('complete')
    expect(source.reads).toEqual([0, 1, 2, 2, 3, 4, 5, 6, 7, 8, 9])
    expect(index.ledger.mergedPages).toBe(10)
    expect(keysOf(index)).toEqual(items(40).map(value => `s:${value}`))
  })
  it('cancels searches and discards them when the draft is cancelled', async () => {
    const { index } = await built(items(15_000), mixed, 1024)
    const controller = new AbortController()
    const pending = index.search({ text: 'item', locale: 'en-US' }, { signal: controller.signal, yield: async () => { controller.abort() } })
    await expect(pending).rejects.toMatchObject({ code: 'cancelled' })
    const draft = await ValueSelectionDraft.open(index, null, new MemoryValueRunStore())
    const inFlight = draft.search({ text: 'item', locale: 'en-US' }, { yield: async () => { draft.cancel() } })
    await expect(inFlight).rejects.toMatchObject({ code: 'cancelled' })
  })
  it('reports long text as a typed resource limit, never a silent truncation', async () => {
    const { index } = await built(['short', 'x'.repeat(70_000)])
    await expect(index.search({ text: 'zz', locale: 'en-US' })).rejects.toMatchObject({ code: 'resource-limit' })
    const page = index.page({ offset: 0, limit: 5 })
    expect(page.entries[1]).toMatchObject({ labelTruncated: true, key: `s:${'x'.repeat(70_000)}` })
    expect(page.entries[1].label.length).toBeLessThanOrEqual(256)
  })
  it('invalidates on generation change and rejects stale draft commits', async () => {
    const cache = new ValueIndexCache(), source = fromCells(items(10))
    const index = cache.open(identity(), source)
    await index.build()
    const draft = await ValueSelectionDraft.open(index, null, new MemoryValueRunStore())
    draft.toggle(0, false)
    const next = cache.open(identity(mixed, { valueGen: 4 }), { ...source, stamp: { ...STAMP, valueGen: 4 }, pageCount: source.pageCount, readPage: (page, signal) => source.readPage(page, signal) })
    expect(next).not.toBe(index)
    cache.invalidate(owner => owner.valueGen < 4)
    expect(index.disposed).toBe(true)
    await expect(draft.commit()).rejects.toMatchObject({ code: 'stale-source' })
    await expect(ValueSelectionDraft.open(index, null, new MemoryValueRunStore())).rejects.toMatchObject({ code: 'stale-source' })
    await next.build()
    const nextSearch = await next.search({ text: 'item', locale: 'en-US' })
    const fresh = await ValueSelectionDraft.open(next, null, new MemoryValueRunStore())
    expect(() => fresh.selectAll(true, { ...nextSearch, owner: STAMP })).toThrow(ValueIndexError)
  })
})

describe('SourceStore adapter', () => {
  it('translates queryColumn pages into typed lane pages and keeps unknown pages loading', async () => {
    const cell = (row: number, value: string | number | null, error = false): ColumnEntry => ({
      kind: 'cell', row, cell: { id: row * 16384 + 2, row, col: 2, value, valueIsError: error, hasCachedValue: false, styleIndex: 7 } as never,
    })
    const pages: Record<number, ColumnEntry[]> = {
      1: [cell(1, 'a'), { kind: 'blank', r0: 2, r1: 2 }],
      3: [cell(3, 5), cell(4, '#N/A', true)],
      5: [{ kind: 'unknown', r0: 5, r1: 6, pageId: 'remote-7' }],
    }
    const store = {
      stamp: STAMP,
      async queryColumn(request: { rows: { kind: string; intervals?: { r0: number }[] } }): Promise<Page<ColumnEntry>> {
        const r0 = request.rows.intervals![0].r0
        return { items: pages[r0], completeness: pages[r0][0].kind === 'unknown' ? 'loading' : 'complete', stamp: STAMP }
      },
    } as unknown as SourceStore
    const source = valueSourceFromStore(store, 2, { r0: 1, r1: 6 }, 2)
    expect(source.pageCount).toBe(3)
    const first = await source.readPage(0, new AbortController().signal)
    expect(first).toMatchObject({ kind: 'values', count: 2, texts: ['a'] })
    expect(Array.from((first as ValueLanePage).tags)).toEqual([SOURCE_STRING, SOURCE_BLANK])
    expect(Array.from((first as ValueLanePage).styles!)).toEqual([7, 0])
    const second = await source.readPage(1, new AbortController().signal) as ValueLanePage
    expect(Array.from(second.tags)).toEqual([SOURCE_NUMBER, SOURCE_ERROR])
    expect(await source.readPage(2, new AbortController().signal)).toEqual({ kind: 'unknown', pageId: 'remote-7' })
    const index = new ValueIndexCache().open(identity(), source)
    expect(await index.build()).toBe('loading')
    expect(keysOf(index)).toEqual(['s:a', 'n:5', 'e:#N/A'])
  })
})
