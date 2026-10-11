import { describe, expect, it } from 'vitest'
import { canonicalKeyStrategyId } from '../src/xlsx/view/canonicalText'
import type { FilterRange, SourceStamp, ValuePatch, ValueSearchPredicate } from '../src/xlsx/view/contracts'
import type { ColumnClassification } from '../src/xlsx/view/filter'
import {
  SOURCE_STRING, ValueIndexCache, valueScope,
  type ValueColumnSource, type ValueIndex, type ValueIndexIdentity, type ValuePageRead, type ValueSearchResult,
} from '../src/xlsx/view/valueIndex'
import {
  EXACT_SEARCH, MemorySearchKeyStore, SEARCH_HASH_CHUNK_BYTES, canonicalKeyStrategy, createSearchKeyStrategy, foldVariantOf, textContentHash,
  type SearchKeyOwner, type SearchKeyStrategy, type SearchKeyStrategySpec,
} from '../src/xlsx/view/searchKeys'

const STAMP: SourceStamp = { docId: 'doc', sheetId: 's1', valueGen: 3 }
const RANGE: FilterRange = { r0: 0, r1: 1_048_575, c0: 0, c1: 3, headerRows: 1, totalsRows: 0 }
const mixed: ColumnClassification = { kind: 'text', hasNumeric: true }
const identity = (extra: Partial<ValueIndexIdentity> = {}): ValueIndexIdentity =>
  ({ ...STAMP, range: RANGE, col: 2, classification: mixed, classificationGen: 1, ...extra })
class StringSource implements ValueColumnSource {
  stamp = STAMP
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
const asciiSpec = (calls = { n: 0 }): SearchKeyStrategySpec => ({
  canonicalizerVersion: 'fake-ascii-lower-1', unicodeVersion: '15.1', foldVariant: 'default',
  searchKeyOf: text => { calls.n++; return printable.test(text) ? text.toLowerCase() : null },
})
const asciiCandidate = (calls = { n: 0 }): SearchKeyStrategy => createSearchKeyStrategy(asciiSpec(calls))
const digits: SearchKeyStrategy = createSearchKeyStrategy({
  canonicalizerVersion: 'fake-digits-1', unicodeVersion: '15.1', foldVariant: 'default',
  searchKeyOf: text => /^[0-9]+$/.test(text) ? text : null,
})
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
async function index(values: readonly string[], options: ConstructorParameters<typeof ValueIndexCache>[0] = { searchKeys: { strategy: EXACT_SEARCH } }, extra: Partial<ValueIndexIdentity> = {}): Promise<ValueIndex> {
  const built = new ValueIndexCache(options).open(identity(extra), new StringSource(values))
  await built.build()
  return built
}
const owner = (strategy: SearchKeyStrategy, extra: Partial<ValueIndexIdentity> = {}): SearchKeyOwner => {
  const id = identity(extra)
  return { docId: id.docId, sheetId: id.sheetId, valueGen: id.valueGen, col: id.col, classificationGen: id.classificationGen, scope: valueScope(id), strategyId: strategy.id }
}

describe('search-key strategies', () => {
  it('defaults to the exact matcher with no resident keys; canonical keys are opt-in per fold variant', async () => {
    const values = seededValues(500, 11)
    const plain = await index(values, {})
    const exact = await plain.search({ text: 'ab', locale: 'en-US' })
    expect(new ValueIndexCache().searchKeyStrategy).toBe(EXACT_SEARCH)
    expect(exact.stats).toEqual({ keyed: 0, exact: 500 })
    expect(plain.searchKeyStatus).toMatchObject({ strategyId: EXACT_SEARCH.id, state: 'none', keyed: 0 })
    const canonical = await index(values, { searchKeys: { strategy: canonicalKeyStrategy(foldVariantOf('en-US')) } })
    const result = await canonical.search({ text: 'ab', locale: 'en-US' })
    expect(canonical.searchKeyStatus).toMatchObject({ strategyId: canonicalKeyStrategyId('default'), state: 'computed', keyed: 500 })
    expect(result.stats).toEqual({ keyed: 500, exact: 0 })
    expect(canonicalKeyStrategy(foldVariantOf('tr-TR')).id).toBe(canonicalKeyStrategyId('turkic'))
    expect(matchedOrdinals(result)).toEqual(matchedOrdinals(exact))
  })
  it('creates frozen, versioned strategies with a fold-variant dimension and no equivalent mode', () => {
    const strategy = asciiCandidate()
    expect(strategy.id).toBe('fake-ascii-lower-1|unicode-15.1|default')
    expect(strategy.foldVariant).toBe('default')
    expect(Object.isFrozen(strategy)).toBe(true)
    expect('exactness' in strategy).toBe(false)
    expect(foldVariantOf('tr-TR')).toBe('turkic')
    expect(foldVariantOf('az-Latn-AZ')).toBe('turkic')
    expect(foldVariantOf('en-US')).toBe('default')
    expect(() => createSearchKeyStrategy({ ...asciiSpec(), foldVariant: 'other' as never })).toThrow(expect.objectContaining({ code: 'invalid-rule' }))
  })
  it('refuses strategies that are not created by the certifying factory', () => {
    const forged = { id: 'forged', foldVariant: 'default', searchKeyOf: () => 'x', matchKey: () => false } as unknown as SearchKeyStrategy
    expect(() => new ValueIndexCache({ searchKeys: { strategy: forged } })).toThrow(expect.objectContaining({ code: 'invalid-rule' }))
  })
  it('refuses a key function whose misses contradict the exact matcher on the probe set (codex r2 1)', () => {
    expect(() => createSearchKeyStrategy({ ...asciiSpec(), searchKeyOf: text => `#${text.length}` }))
      .toThrow(expect.objectContaining({ code: 'invalid-rule' }))
    expect(() => createSearchKeyStrategy({ ...asciiSpec(), searchKeyOf: () => 'alpha', matchKey: () => false }))
      .toThrow(expect.objectContaining({ code: 'invalid-rule' }))
  })
  it('fixes the strategy at construction; the exposed strategy cannot be swapped', async () => {
    const cache = new ValueIndexCache({ searchKeys: { strategy: asciiCandidate() } })
    expect(Object.isFrozen(cache.searchKeyStrategy)).toBe(true)
    expect(() => { (cache as unknown as { searchKeyStrategy: unknown }).searchKeyStrategy = EXACT_SEARCH }).toThrow()
    const built = cache.open(identity(), new StringSource(['ALPHA', 'beta']))
    await built.build()
    expect(matchedOrdinals(await built.search({ text: 'alpha', locale: 'en-US' }))).toEqual([0])
  })
  it('matches the exact path on seeded values incl. hostile pools with candidate keys', async () => {
    const values = seededValues(1_200, 17), exact = await index(values)
    const strategies = [asciiCandidate(), digits]
    const keyed = await Promise.all(strategies.map(strategy => index(values, { searchKeys: { strategy } })))
    for (const built of keyed) await built.search({ text: 'a', locale: 'en-US' })
    keyed.forEach((built, at) => expect(built.searchKeyStatus).toMatchObject({ strategyId: strategies[at].id, state: 'computed', keyed: 1_200 }))
    let comparisons = 0
    const keyedDecisions = [0, 0]
    for (const query of seededQueries(values, 30, 101)) {
      for (const locale of ['en-US', 'de-DE', 'da-DK']) {
        const expected = matchedOrdinals(await exact.search({ text: query, locale }))
        for (let at = 0; at < keyed.length; at++) {
          const result = await keyed[at].search({ text: query, locale })
          expect(matchedOrdinals(result), `${strategies[at].id} ${JSON.stringify(query)} ${locale}`).toEqual(expected)
          comparisons += values.length
          keyedDecisions[at] += result.stats.keyed
        }
      }
    }
    expect(comparisons).toBeGreaterThanOrEqual(200_000)
    expect(keyedDecisions.every(count => count > 0)).toBe(true)
  }, 300_000)
  it('runs the exact matcher only on null-key values and confirmed candidates', async () => {
    const keyed = await index(['apple', 'APPLE pie', 'banana', 'äpple', 'grape'], { searchKeys: { strategy: asciiCandidate() } })
    const result = await keyed.search({ text: 'apple', locale: 'en-US' })
    expect(matchedOrdinals(result)).toEqual([0, 1])
    expect(result.stats).toEqual({ keyed: 4, exact: 3 })
  })
  it('uses keys only for queries of the strategy fold variant', async () => {
    const values = ['Istanbul', 'ıstanbul', 'istanbul', 'İstanbul']
    const keyed = await index(values, { searchKeys: { strategy: asciiCandidate() } }), exact = await index(values)
    const turkish = await keyed.search({ text: 'is', locale: 'tr-TR' })
    expect(turkish.stats.keyed).toBe(0)
    expect(matchedOrdinals(turkish)).toEqual(matchedOrdinals(await exact.search({ text: 'is', locale: 'tr-TR' })))
    expect((await keyed.search({ text: 'is', locale: 'en-US' })).stats.keyed).toBeGreaterThan(0)
  })
  it('searches wildcard patterns through the exact matcher only', async () => {
    const values = ['item-105', 'item-15', 'item-1x5', 'item-25', 'a*b', 'a-b']
    const keyed = await index(values, { searchKeys: { strategy: asciiCandidate() } })
    const predicate: ValueSearchPredicate = { text: 'item-1*5', locale: 'en-US', wildcards: true }
    const star = await keyed.search(predicate)
    expect(matchedOrdinals(star)).toEqual([0, 1, 2])
    expect(star.stats.keyed).toBe(0)
    const patch: ValuePatch = { domain: { kind: 'search', query: star.query, owner: star.owner }, selected: false }
    expect(JSON.parse(JSON.stringify(patch.domain))).toMatchObject({ query: { text: 'item-1*5', locale: 'en-US', wildcards: true } })
    expect(matchedOrdinals(await keyed.search({ text: 'item-1?5', locale: 'en-US', wildcards: true }))).toEqual([0, 2])
    expect(matchedOrdinals(await keyed.search({ text: 'a~*b', locale: 'en-US', wildcards: true }))).toEqual([4])
    expect(matchedOrdinals(await keyed.search({ text: 'a*b', locale: 'en-US' }))).toEqual([4])
    await expect(keyed.search({ text: 'a~', locale: 'en-US', wildcards: true })).rejects.toMatchObject({ code: 'invalid-rule' })
  })
  it('builds keys lazily on the first plain search only, never in build() (spike decision)', async () => {
    const calls = { n: 0 }, strategy = asciiCandidate(calls), probeCalls = calls.n
    const values = seededValues(800, 83)
    const cache = new ValueIndexCache({ searchKeys: { strategy } }), built = cache.open(identity(), new StringSource(values))
    expect(await built.build()).toBe('complete')
    expect(calls.n - probeCalls).toBe(0)
    expect(built.searchKeyStatus).toMatchObject({ state: 'none', keyed: 0 })
    await built.search({ text: '', locale: 'en-US' })
    await built.search({ text: 'a*b', locale: 'en-US', wildcards: true })
    await built.search({ text: 'a', locale: 'tr-TR' })
    expect(built.searchKeyStatus).toMatchObject({ state: 'none', keyed: 0 })
    expect(calls.n - probeCalls).toBe(0)
    const progress: string[] = []
    const first = await built.search({ text: 'a', locale: 'en-US' }, { onKeyProgress: p => progress.push(p.phase) })
    expect(progress).toContain('keys')
    expect(built.searchKeyStatus).toMatchObject({ state: 'computed', keyed: 800 })
    expect(first.stats.keyed).toBe(values.filter(value => printable.test(value)).length)
    expect(calls.n - probeCalls).toBe(801)
    await built.search({ text: 'b', locale: 'en-US' })
    expect(calls.n - probeCalls).toBe(802)
  })
  it('falls back to streaming keys under memory pressure with identical results', async () => {
    const values = seededValues(1_000, 61)
    const probe = new ValueIndexCache(), sized = probe.open(identity(), new StringSource(values))
    await sized.build()
    const cache = new ValueIndexCache({ limitBytes: probe.usedBytes + 8 * 1024, searchKeys: { strategy: asciiCandidate() } })
    const streaming = cache.open(identity(), new StringSource(values))
    expect(await streaming.build()).toBe('complete')
    expect(streaming.searchKeyStatus).toMatchObject({ state: 'none', keyed: 0 })
    await streaming.search({ text: 'b', locale: 'en-US' })
    expect(streaming.searchKeyStatus).toMatchObject({ state: 'streaming', keyed: 0 })
    expect(cache.usedBytes).toBeLessThanOrEqual(cache.limitBytes)
    for (const text of ['a', 'Z1', 'é', '']) {
      const result = await streaming.search({ text, locale: 'en-US' })
      expect(matchedOrdinals(result)).toEqual(matchedOrdinals(await sized.search({ text, locale: 'en-US' })))
      if (text === 'a') expect(result.stats.keyed).toBeGreaterThan(0)
    }
  }, 60_000)
})

describe('search-key store port', () => {
  it('persists keys with a content hash after a complete build and reuses them on a cold index', async () => {
    const values = seededValues(2_000, 41), store = new MemorySearchKeyStore()
    const first = { n: 0 }, strategy = asciiCandidate(first), firstIndex = await index(values, { searchKeys: { strategy, store } })
    expect(store.saves).toBe(0)
    const firstHit = await firstIndex.search({ text: 'Ab', locale: 'en-US' })
    expect(firstIndex.searchKeyStatus).toMatchObject({ state: 'persisted', keyed: 2_000 })
    expect(store.saves).toBe(1)
    const stored = await store.load(owner(strategy))
    expect(stored).toMatchObject({ count: 2_000, hash: expect.stringMatching(/^[0-9a-f]{64}$/) })
    expect('texts' in stored!).toBe(false)
    const second = { n: 0 }, secondStrategy = asciiCandidate(second), probeCalls = second.n
    const warm = await index(values, { searchKeys: { strategy: secondStrategy, store } })
    const warmHit = await warm.search({ text: 'Ab', locale: 'en-US' })
    expect(warm.searchKeyStatus).toMatchObject({ state: 'loaded', keyed: 2_000 })
    // Only the query key is computed; value keys come from the store.
    expect(second.n - probeCalls).toBe(1)
    expect(matchedOrdinals(warmHit)).toEqual(matchedOrdinals(firstHit))
    expect(await store.load(owner(strategy, { valueGen: 4 }))).toBeNull()
  })
  it('discards persisted keys whose content hash differs and recomputes', async () => {
    const values = seededValues(300, 43), store = new MemorySearchKeyStore(), strategy = asciiCandidate()
    await store.save(owner(strategy), { count: 300, hash: '0'.repeat(64), keys: values.map(() => 'zzz') })
    const calls = { n: 0 }, counted = asciiCandidate(calls), probeCalls = calls.n
    const built = await index(values, { searchKeys: { strategy: counted, store } })
    await built.search({ text: 'q', locale: 'en-US' })
    expect(calls.n - probeCalls).toBe(301)
    expect(built.searchKeyStatus.state).toBe('persisted')
    const exact = await index(values)
    for (const text of ['a', 'z', '1']) {
      expect(matchedOrdinals(await built.search({ text, locale: 'en-US' }))).toEqual(matchedOrdinals(await exact.search({ text, locale: 'en-US' })))
    }
  })
  it('hashes lone surrogates and code-unit order losslessly', async () => {
    const store = new MemorySearchKeyStore(), strategy = asciiCandidate()
    await (await index(['\ud800', 'a'], { searchKeys: { strategy, store } })).search({ text: 'a', locale: 'en-US' })
    const first = (await store.load(owner(strategy)))!.hash
    const other = new MemorySearchKeyStore()
    await (await index(['\udc00', 'a'], { searchKeys: { strategy, store: other } })).search({ text: 'a', locale: 'en-US' })
    expect((await other.load(owner(strategy)))!.hash).not.toBe(first)
  })
  it('reports a typed storage quota without changing results', async () => {
    const values = seededValues(400, 47), store = new MemorySearchKeyStore({ limitBytes: 128 })
    const built = await index(values, { searchKeys: { strategy: asciiCandidate(), store } })
    expect(built.completeness).toBe('complete')
    await built.search({ text: 'q', locale: 'en-US' })
    expect(built.searchKeyStatus).toMatchObject({ state: 'unpersisted', keyed: 400, error: { code: 'quota', detail: { resource: 'storage', limitBytes: 128 } } })
    expect(store.usedBytes).toBe(0)
    const exact = await index(values)
    expect(matchedOrdinals(await built.search({ text: 'b', locale: 'en-US' }))).toEqual(matchedOrdinals(await exact.search({ text: 'b', locale: 'en-US' })))
  })
  it('charges the in-memory key store to the owning cache pool and deletes keys on owner invalidation (codex r2 3)', async () => {
    const store = new MemorySearchKeyStore(), strategy = asciiCandidate()
    const keyed = new ValueIndexCache({ searchKeys: { strategy, store } })
    const built = keyed.open(identity(), new StringSource(seededValues(1_000, 71)))
    await built.build()
    built.release(await built.search({ text: 'q', locale: 'en-US' }))
    expect(store.usedBytes).toBeGreaterThan(0)
    expect(keyed.usedBytes).toBe(built.ledger.chargedBytes + store.usedBytes)
    keyed.invalidate()
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(await store.load(owner(strategy))).toBeNull()
    expect(store.usedBytes).toBe(0)
    expect(keyed.usedBytes).toBe(0)
  })
  it('computes keys lazily in a cancellable, resumable phase with progress, exact meanwhile', async () => {
    const values = seededValues(1_500, 53), controller = new AbortController(), calls = { n: 0 }
    const strategy = createSearchKeyStrategy({ ...asciiSpec(), searchKeyOf: text => { if (++calls.n === 700) controller.abort(); return printable.test(text) ? text.toLowerCase() : null } })
    const cache = new ValueIndexCache({ searchKeys: { strategy } }), source = new StringSource(values)
    const built = cache.open(identity(), source), phases: string[] = []
    await built.build()
    await expect(built.search({ text: 'a', locale: 'en-US' }, { signal: controller.signal, sliceMs: 0, onKeyProgress: p => phases.push(p.phase) })).rejects.toMatchObject({ code: 'cancelled' })
    expect(phases).toContain('keys')
    const partial = built.searchKeyStatus.keyed
    expect(partial).toBeGreaterThan(0)
    expect(partial).toBeLessThan(1_500)
    const exact = await index(values)
    const reads = source.reads
    expect(matchedOrdinals(await built.search({ text: 'a', locale: 'en-US' }))).toEqual(matchedOrdinals(await exact.search({ text: 'a', locale: 'en-US' })))
    expect(source.reads).toBe(reads)
    expect(built.searchKeyStatus).toMatchObject({ state: 'computed', keyed: 1_500 })
  })
  it('stores immutable copies and supports delete', async () => {
    const store = new MemorySearchKeyStore(), strategy = asciiCandidate(), keys: (string | null)[] = ['a', null]
    await store.save(owner(strategy), { count: 2, hash: 'f'.repeat(64), keys })
    keys[1] = 'mutated'
    expect(await store.load(owner(strategy))).toEqual({ count: 2, hash: 'f'.repeat(64), keys: ['a', null] })
    expect(store.usedBytes).toBeGreaterThan(0)
    await store.delete(owner(strategy))
    expect(await store.load(owner(strategy))).toBeNull()
    expect(store.usedBytes).toBe(0)
  })
})

describe('review round 3 findings', () => {
  const big = (count: number, units: number) => Array.from({ length: count }, (_, i) => String(i).padStart(8, '0') + 'x'.repeat(units - 8))
  it('charges hash scratch: no room for one chunk falls back to streaming (codex r3 2)', async () => {
    const values = big(256, 4096), probe = new ValueIndexCache(), sized = probe.open(identity(), new StringSource(values, 64))
    await sized.build()
    // Tiny keys (8 units, never decisive) so only the hash scratch can overflow.
    const tiny = createSearchKeyStrategy({ ...asciiSpec(), canonicalizerVersion: 'fake-prefix-1', searchKeyOf: text => text.slice(0, 8), matchKey: () => true })
    const store = new MemorySearchKeyStore({ account: new ValueIndexCache().reserve() })
    const cache = new ValueIndexCache({ limitBytes: probe.usedBytes + 64 * 1024, searchKeys: { strategy: tiny, store } })
    const index = cache.open(identity(), new StringSource(values, 64))
    await index.build()
    await index.search({ text: '0001', locale: 'en-US' })
    expect(index.searchKeyStatus.state).toBe('streaming')
    expect(cache.usedBytes).toBeLessThanOrEqual(cache.limitBytes)
  })
  it('hashes in bounded chunks that fit far below the full text size (codex r3 2)', async () => {
    const values = big(1024, 4096), total = values.length * (4 + 4096 * 2)
    expect(total).toBeGreaterThan(4 * SEARCH_HASH_CHUNK_BYTES)
    const probe = new ValueIndexCache(), sized = probe.open(identity(), new StringSource(values, 128))
    await sized.build()
    const keys = values.length * (4096 * 2 + 48)
    const limit = probe.usedBytes + keys + 3 * SEARCH_HASH_CHUNK_BYTES + 64 * 1024
    expect(limit).toBeLessThan(probe.usedBytes + keys + 2 * total)
    // The store's copy is charged to a separate pool so only hash scratch competes here.
    const store = new MemorySearchKeyStore({ account: new ValueIndexCache().reserve() })
    const cache = new ValueIndexCache({ limitBytes: limit, searchKeys: { strategy: asciiCandidate(), store } })
    const index = cache.open(identity(), new StringSource(values, 128))
    await index.build()
    await index.search({ text: '0001', locale: 'en-US' })
    expect(index.searchKeyStatus.state).toBe('persisted')
  }, 60_000)
  it('cancels hashing within a slice (codex r3 2)', async () => {
    const values = big(4096, 4096), store = new MemorySearchKeyStore()
    const cache = new ValueIndexCache({ searchKeys: { strategy: asciiCandidate(), store } })
    const index = cache.open(identity(), new StringSource(values, 256))
    await index.build()
    const controller = new AbortController()
    // Measured from the scheduled time: synchronous hashing delays the timer itself.
    const scheduledAt = performance.now() + 5
    setTimeout(() => controller.abort(), 5)
    await expect(index.search({ text: '0001', locale: 'en-US' }, { signal: controller.signal, sliceMs: 0 })).rejects.toMatchObject({ code: 'cancelled' })
    expect(performance.now() - scheduledAt).toBeLessThan(150)
    expect(index.searchKeyStatus.state).not.toBe('persisted')
  }, 60_000)
  it('replaces stored runs atomically under concurrent saves (codex r3 3)', async () => {
    const pool = new ValueIndexCache(), store = new MemorySearchKeyStore({ account: pool.reserve() }), strategy = asciiCandidate()
    const run = (key: string) => ({ count: 2, hash: 'a'.repeat(64), keys: [key, null] })
    await store.save(owner(strategy), run('first'))
    await Promise.all([store.save(owner(strategy), run('second')), store.save(owner(strategy), run('third-longer')), store.save(owner(strategy), run('x'))])
    expect(store.usedBytes).toBe(pool.usedBytes)
    await store.delete(owner(strategy))
    expect(store.usedBytes).toBe(0)
    expect(pool.usedBytes).toBe(0)
  })
})

describe('review round 4 findings', () => {
  const recorder = () => {
    const account = { live: 0, peak: 0, charge(bytes: number) { this.live += bytes; this.peak = Math.max(this.peak, this.live) }, release(bytes: number) { this.live -= bytes } }
    return account
  }
  it('bounds hash scratch even for one value far larger than a chunk (codex r4 2)', async () => {
    const huge = 'x'.repeat(4 * 1024 * 1024) + 'y', account = recorder()
    let chunks = 0
    const hash = await textContentHash(['a', huge, 'b'], { account, afterChunk: () => { chunks++ } })
    expect(account.peak).toBeLessThanOrEqual(2 * SEARCH_HASH_CHUNK_BYTES + 128)
    expect(account.live).toBe(0)
    expect(chunks).toBeGreaterThanOrEqual(8)
    expect(await textContentHash(['a', huge, 'b'], { account: recorder(), afterChunk: () => {} })).toBe(hash)
    const changed = 'x'.repeat(4 * 1024 * 1024) + 'z'
    expect(await textContentHash(['a', changed, 'b'], { account: recorder(), afterChunk: () => {} })).not.toBe(hash)
    // Segment boundaries stay injective: moving a code unit across values changes the hash.
    expect(await textContentHash(['ax', 'b'], { account: recorder(), afterChunk: () => {} }))
      .not.toBe(await textContentHash(['a', 'xb'], { account: recorder(), afterChunk: () => {} }))
  }, 60_000)
  it('stops hashing a huge value at the first chunk boundary after cancellation (codex r4 2)', async () => {
    const huge = 'x'.repeat(8 * 1024 * 1024)
    let calls = 0
    await expect(textContentHash([huge], { account: recorder(), afterChunk: () => { if (++calls === 2) throw new Error('stop') } })).rejects.toThrow('stop')
    expect(calls).toBe(2)
  }, 60_000)
  it('reconciles a replaced run that was shed during its own admission (codex r4 1)', async () => {
    const store = new MemorySearchKeyStore(), strategy = asciiCandidate()
    const pool = new ValueIndexCache({ limitBytes: 300, searchKeys: { strategy, store } })
    const run = (key: string) => ({ count: 1, hash: 'a'.repeat(64), keys: [key] })
    await store.save(owner(strategy), run('k'))
    const one = store.usedBytes
    expect(pool.usedBytes).toBe(one)
    await store.save(owner(strategy), run('j'))
    expect(await store.load(owner(strategy))).toMatchObject({ keys: ['j'] })
    expect(store.usedBytes).toBe(one)
    expect(pool.usedBytes).toBe(one)
    await store.delete(owner(strategy))
    expect(store.usedBytes).toBe(0)
    expect(pool.usedBytes).toBe(0)
  })
})

describe('S8.5 canonical keys', () => {
  it('match the raw canonical predicate exactly over >=100k seeded cases, both fold variants, plain and wildcard', async () => {
    const values = seededValues(2_000, 97)
    values.push('İstanbul', 'ISTANBUL', 'ıstanbul', 'i̇stanbul', 'Straße', 'STRASSE', 'e\u0301', 'é', '👩‍👩‍👧', '🇹🇼', 'a\u200db', 'ﬁ')
    const exactDefault = await index(values), exactTurkic = await index(values)
    const keyedDefault = await index(values, { searchKeys: { strategy: canonicalKeyStrategy('default') } })
    const keyedTurkic = await index(values, { searchKeys: { strategy: canonicalKeyStrategy('turkic') } })
    const random = (() => { let seed = 7; return (n: number) => { seed = (seed * 48271) % 2147483647; return seed % n } })()
    const queries = seededQueries(values, 30, 131)
    for (let i = 0; i < 10; i++) {
      const points = Array.from(values[random(values.length)])
      queries.push(points.map(point => random(4) === 0 ? '?' : random(6) === 0 ? '*' : point).join(''))
    }
    let comparisons = 0
    for (const query of queries) {
      for (const wildcards of [false, true]) {
        for (const [locale, exact, keyed] of [['en-US', exactDefault, keyedDefault], ['tr-TR', exactTurkic, keyedTurkic]] as const) {
          let expected: number[]
          try { expected = matchedOrdinals(await exact.search({ text: query, locale, wildcards })) } catch (error) {
            await expect(keyed.search({ text: query, locale, wildcards })).rejects.toMatchObject({ code: (error as { code: string }).code })
            continue
          }
          const result = await keyed.search({ text: query, locale, wildcards })
          expect(matchedOrdinals(result), `${JSON.stringify(query)} ${locale} wildcards=${wildcards}`).toEqual(expected)
          if (query) expect(result.stats.keyed).toBeGreaterThan(0)
          comparisons += values.length
        }
      }
    }
    expect(comparisons).toBeGreaterThanOrEqual(100_000)
  }, 300_000)
})

describe('review round 5 findings', () => {
  const mixedSource = (texts: readonly string[], numbers: readonly number[]): ValueColumnSource => ({
    stamp: STAMP, pageCount: 1,
    async readPage() {
      const count = texts.length + numbers.length, tags = new Uint8Array(count), lanes = new Float64Array(count), cells: string[] = []
      texts.forEach((text, i) => { tags[i] = SOURCE_STRING; cells[i] = text })
      numbers.forEach((value, i) => { tags[texts.length + i] = 1; lanes[texts.length + i] = value })
      return { kind: 'values', pageId: 'p0', count, tags, numbers: lanes, texts: cells }
    },
  })
  it.each([['exact', EXACT_SEARCH], ['canonical', canonicalKeyStrategy('default')]] as const)(
    'treats canonical-empty queries as no search term (%s strategy, codex r5 1)', async (_name, strategy) => {
      const cache = new ValueIndexCache({ searchKeys: { strategy } })
      const index = cache.open(identity(), mixedSource(['alpha', 'x'.repeat(70_000)], [1.5, 2.5]))
      await index.build()
      for (const text of ['\u200b', '\u034f', '\u00ad', '\u200b\u00ad\u034f']) {
        for (const wildcards of [false, true]) {
          const result = await index.search({ text, locale: 'en-US', wildcards })
          expect(result.count, JSON.stringify(text)).toBe(4)
          expect(result.stats.keyed).toBe(0)
        }
      }
      expect(index.searchKeyStatus.state).toBe('none')
      await expect(index.search({ text: '\u200b*', locale: 'en-US', wildcards: true })).rejects.toMatchObject({ code: 'missing-display-text' })
    })
  it('preflights raw length before generating keys, resident and streaming (codex r5 2)', async () => {
    const huge = '가'.repeat(1 << 20), seen = { max: 0 }
    const spy = createSearchKeyStrategy({ ...asciiSpec(), canonicalizerVersion: 'spy-1', searchKeyOf: text => { seen.max = Math.max(seen.max, text.length); return printable.test(text) ? text.toLowerCase() : null } })
    for (const strategy of [spy, canonicalKeyStrategy('default')]) {
      const probe = new ValueIndexCache({ searchKeys: { strategy: EXACT_SEARCH } }), sized = probe.open(identity(), new StringSource(['alpha', huge]))
      await sized.build()
      await expect(sized.search({ text: 'al', locale: 'en-US' })).rejects.toMatchObject({ code: 'resource-limit' })
      for (const room of [256 * 1024, 64]) {
        seen.max = 0
        const cache = new ValueIndexCache({ limitBytes: probe.usedBytes + room, searchKeys: { strategy } })
        const index = cache.open(identity(), new StringSource(['alpha', huge]))
        await index.build()
        await expect(index.search({ text: 'al', locale: 'en-US' })).rejects.toMatchObject({ code: 'resource-limit' })
        expect(cache.usedBytes).toBeLessThanOrEqual(cache.limitBytes)
        if (strategy === spy) expect(seen.max).toBeLessThan(huge.length)
      }
    }
  }, 60_000)
})
