import { describe, expect, it } from 'vitest'
import type { FilterRange, SourceStamp } from '../src/xlsx/view/contracts'
import type { ColumnClassification } from '../src/xlsx/view/filter'
import {
  SOURCE_STRING, ValueIndexCache, valueScope,
  type ValueColumnSource, type ValueIndex, type ValueIndexIdentity, type ValuePageRead, type ValueSearchResult,
} from '../src/xlsx/view/valueIndex'
import { EXACT_SEARCH, MemorySearchKeyStore, type SearchKeyOwner, type SearchKeyStrategy } from '../src/xlsx/view/searchKeys'

const STAMP: SourceStamp = { docId: 'doc', sheetId: 's1', valueGen: 3 }
const RANGE: FilterRange = { r0: 0, r1: 1_048_575, c0: 0, c1: 3, headerRows: 1, totalsRows: 0 }
const mixed: ColumnClassification = { kind: 'text', hasNumeric: true }
const identity = (extra: Partial<ValueIndexIdentity> = {}): ValueIndexIdentity =>
  ({ ...STAMP, range: RANGE, col: 2, classification: mixed, classificationGen: 1, ...extra })
class StringSource implements ValueColumnSource {
  readonly stamp = STAMP
  reads = 0
  constructor(readonly values: readonly string[], readonly rows = 512, readonly onRead?: (page: number) => void) {}
  get pageCount(): number { return Math.ceil(this.values.length / this.rows) }
  async readPage(page: number): Promise<ValuePageRead> {
    this.reads++
    this.onRead?.(page)
    const texts = this.values.slice(page * this.rows, (page + 1) * this.rows)
    return { kind: 'values', pageId: `p${page}`, count: texts.length, tags: new Uint8Array(texts.length).fill(SOURCE_STRING), numbers: new Float64Array(texts.length), texts }
  }
}
/** Fake strategies for the injection seam only; S8.5 owns the real key. */
const printable = /^[\x20-\x7e]*$/
const asciiCandidate = (calls = { n: 0 }): SearchKeyStrategy => ({
  id: 'fake-ascii-lower-v1', exactness: 'candidate',
  searchKeyOf: text => { calls.n++; return printable.test(text) ? text.toLowerCase() : null },
})
const digitsEquivalent: SearchKeyStrategy = { id: 'fake-digits-v1', exactness: 'equivalent', searchKeyOf: text => /^[0-9]+$/.test(text) ? text : null }
const ascii = Array.from({ length: 95 }, (_, i) => String.fromCharCode(32 + i))
const hostile = ['é', 'e', '́', '​', '­', '﻿', '⁠', 'ﬁ', 'ß', '😀', '‍', '漢', '\ud800', '\udc00', 'å', 'İ', 'Å', 'ı', 'ǅ', 'Ⅳ', '·', 'l']
function seededValues(count: number, seed: number): string[] {
  const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
  const values = new Set<string>()
  while (values.size < count) {
    const pool = values.size % 3 === 0 ? [...ascii, ...hostile] : values.size % 3 === 1 ? ascii : '0123456789'.split('')
    let text = ''
    for (let i = 1 + random(10); i > 0; i--) text += pool[random(pool.length)]
    values.add(text)
  }
  return [...values]
}
function seededQueries(values: readonly string[], count: number, seed: number): string[] {
  const random = (length: number) => { seed = (seed * 48271) % 2147483647; return seed % length }
  const queries = ['', 'a', 'A', 'ss', 'SS', 'ß', 'e', 'é', 'fi', 'ﬁ', 'aa', 'å', 'i', 'İ', '1', '12', '\ud800', '😀', ' ', '~']
  while (queries.length < count) {
    const value = values[random(values.length)], points = Array.from(value)
    const start = random(points.length), end = start + 1 + random(Math.max(1, points.length - start))
    const slice = points.slice(start, end).join('')
    queries.push(random(2) ? slice.toUpperCase() : slice)
  }
  return queries
}
const matchedOrdinals = (result: ValueSearchResult): number[] => {
  const out: number[] = []
  for (let i = 0; i < result.searchedEntries; i++) if (result.matches[i >>> 5] & (1 << (i & 31))) out.push(i)
  return out
}
async function index(values: readonly string[], options: ConstructorParameters<typeof ValueIndexCache>[0] = {}, extra: Partial<ValueIndexIdentity> = {}): Promise<ValueIndex> {
  const built = new ValueIndexCache(options).open(identity(extra), new StringSource(values))
  await built.build()
  return built
}

describe('search-key injection seam', () => {
  it('defaults to the exact matcher for every value', async () => {
    const values = seededValues(500, 11), plain = await index(values)
    const result = await plain.search({ text: 'ab', locale: 'en-US' })
    expect(result.stats).toEqual({ keyed: 0, exact: 500 })
    expect(plain.searchKeyStatus).toMatchObject({ strategyId: EXACT_SEARCH.id, state: 'none', keyed: 0 })
  })
  it('matches the exact path on seeded values incl. hostile pools (candidate and equivalent keys)', async () => {
    const values = seededValues(1_200, 17), exact = await index(values)
    const strategies = [asciiCandidate(), digitsEquivalent]
    const keyed = await Promise.all(strategies.map(strategy => index(values, { searchKeys: { strategy } })))
    keyed.forEach((built, at) => expect(built.searchKeyStatus).toMatchObject({ strategyId: strategies[at].id, state: 'computed', keyed: 1_200 }))
    let comparisons = 0
    const keyedDecisions = [0, 0]
    for (const query of seededQueries(values, 30, 101)) {
      for (const locale of ['en-US', 'de-DE', 'da-DK']) {
        const expected = matchedOrdinals(await exact.search({ text: query, locale }))
        for (let at = 0; at < keyed.length; at++) {
          const result = await keyed[at].search({ text: query, locale })
          expect(matchedOrdinals(result), `${strategies[at].id} ${JSON.stringify(query)} ${locale}`).toEqual(expected)
          expect(result.stats.keyed + result.stats.exact).toBeGreaterThanOrEqual(values.length)
          comparisons += values.length
          keyedDecisions[at] += result.stats.keyed
        }
      }
    }
    expect(comparisons).toBeGreaterThanOrEqual(200_000)
    expect(keyedDecisions.every(count => count > 0)).toBe(true)
  }, 300_000)
  it('runs the exact matcher only on null-key values and confirmed candidates', async () => {
    const values = ['apple', 'APPLE pie', 'banana', 'äpple', 'grape']
    const keyed = await index(values, { searchKeys: { strategy: asciiCandidate() } })
    const result = await keyed.search({ text: 'apple', locale: 'en-US' })
    expect(matchedOrdinals(result)).toEqual([0, 1])
    // apple and APPLE pie are key candidates; äpple has a null key; banana/grape are excluded by keys.
    expect(result.stats).toEqual({ keyed: 4, exact: 3 })
  })
})

describe('search-key store port', () => {
  const owner = (strategyId: string, extra: Partial<ValueIndexIdentity> = {}): SearchKeyOwner => {
    const id = identity(extra)
    return { docId: id.docId, sheetId: id.sheetId, valueGen: id.valueGen, col: id.col, classificationGen: id.classificationGen, scope: valueScope(id), strategyId }
  }
  it('persists keys after a complete build and reuses them on a cold index without recomputing', async () => {
    const values = seededValues(2_000, 41), store = new MemorySearchKeyStore()
    const first = { n: 0 }, firstIndex = await index(values, { searchKeys: { strategy: asciiCandidate(first), store } })
    expect(firstIndex.searchKeyStatus).toMatchObject({ state: 'persisted', keyed: 2_000 })
    expect(store.saves).toBe(1)
    expect(first.n).toBe(2_000)
    const second = { n: 0 }, warm = await index(values, { searchKeys: { strategy: asciiCandidate(second), store } })
    expect(warm.searchKeyStatus).toMatchObject({ state: 'loaded', keyed: 2_000 })
    expect(second.n).toBe(0)
    const a = await firstIndex.search({ text: 'Ab', locale: 'en-US' }), b = await warm.search({ text: 'Ab', locale: 'en-US' })
    expect(matchedOrdinals(b)).toEqual(matchedOrdinals(a))
    expect(await store.load(owner('fake-ascii-lower-v1'))).toMatchObject({ texts: expect.any(Array), keys: expect.any(Array) })
    expect(await store.load(owner('other-strategy'))).toBeNull()
    expect(await store.load(owner('fake-ascii-lower-v1', { valueGen: 4 }))).toBeNull()
  })
  it('discards persisted keys whose texts do not align exactly and recomputes', async () => {
    const values = seededValues(300, 43), store = new MemorySearchKeyStore()
    await store.save(owner('fake-ascii-lower-v1'), { texts: [...values].reverse(), keys: [...values].reverse().map(() => 'zzz') })
    const calls = { n: 0 }, built = await index(values, { searchKeys: { strategy: asciiCandidate(calls), store } })
    expect(calls.n).toBe(300)
    expect(built.searchKeyStatus.state).toBe('persisted')
    const exact = await index(values)
    for (const text of ['a', 'z', '1']) {
      expect(matchedOrdinals(await built.search({ text, locale: 'en-US' }))).toEqual(matchedOrdinals(await exact.search({ text, locale: 'en-US' })))
    }
  })
  it('reports a typed storage quota without changing results', async () => {
    const values = seededValues(400, 47), store = new MemorySearchKeyStore({ limitBytes: 128 })
    const built = await index(values, { searchKeys: { strategy: asciiCandidate(), store } })
    expect(built.completeness).toBe('complete')
    expect(built.searchKeyStatus).toMatchObject({ state: 'unpersisted', keyed: 400, error: { code: 'quota', detail: { resource: 'storage', limitBytes: 128 } } })
    expect(store.usedBytes).toBe(0)
    const exact = await index(values)
    expect(matchedOrdinals(await built.search({ text: 'b', locale: 'en-US' }))).toEqual(matchedOrdinals(await exact.search({ text: 'b', locale: 'en-US' })))
  })
  it('computes keys in a cancellable, resumable phase with progress, exact meanwhile', async () => {
    const values = seededValues(1_500, 53), controller = new AbortController(), calls = { n: 0 }
    const strategy: SearchKeyStrategy = { ...asciiCandidate(calls), searchKeyOf: text => { if (++calls.n === 700) controller.abort(); return printable.test(text) ? text.toLowerCase() : null } }
    const cache = new ValueIndexCache({ searchKeys: { strategy } }), source = new StringSource(values)
    const built = cache.open(identity(), source), phases: string[] = []
    await expect(built.build({ signal: controller.signal, sliceMs: 0, onProgress: p => phases.push(p.phase) })).rejects.toMatchObject({ code: 'cancelled' })
    expect(phases).toContain('keys')
    const partial = built.searchKeyStatus.keyed
    expect(partial).toBeGreaterThan(0)
    expect(partial).toBeLessThan(1_500)
    const exact = await index(values)
    const midway = await built.search({ text: 'a', locale: 'en-US' })
    expect(matchedOrdinals(midway)).toEqual(matchedOrdinals(await exact.search({ text: 'a', locale: 'en-US' })))
    const reads = source.reads
    expect(await cache.open(identity(), source).build()).toBe('complete')
    expect(source.reads).toBe(reads)
    expect(built.searchKeyStatus).toMatchObject({ state: 'computed', keyed: 1_500 })
  })
  it('stores immutable copies and supports delete', async () => {
    const store = new MemorySearchKeyStore(), texts = ['a', 'b'], keys: (string | null)[] = ['a', null]
    await store.save(owner('s'), { texts, keys })
    texts[0] = 'mutated'; keys[1] = 'mutated'
    expect(await store.load(owner('s'))).toEqual({ texts: ['a', 'b'], keys: ['a', null] })
    expect(store.usedBytes).toBeGreaterThan(0)
    await store.delete(owner('s'))
    expect(await store.load(owner('s'))).toBeNull()
    expect(store.usedBytes).toBe(0)
  })
})
