import { expect, it, vi } from 'vitest'
import { sortRowOrder, type SortResult } from '../src/xlsx/view/sort'
import { sortRowOrder as baseline } from './helpers/sort-baseline'
import { createFixture, createPrng, type FixtureCell } from './helpers/xlsx-viewer-fixtures'
import type { SortSpec } from '../src/xlsx/view/contracts'

const values = [null, '', -Infinity, Infinity, NaN, -0, 0, -4.5, 42, 45292, true, false,
  '#N/A', '#Z!', 'item-2', 'item-02', 'item-10', 'a', 'A', 'ä', 'é', 'e\u0301',
  'ß', 'ss', 'İ', 'i', '中文', 'ไทย', '\ud800', '\udc00', '😀', 'a\r\nb', '\u0000'] as const
function permutation(result: SortResult | Awaited<ReturnType<typeof baseline>>): Uint32Array {
  if (result.kind !== 'complete') throw new Error(JSON.stringify(result))
  const band = result.order.bands[0]
  return band ? Uint32Array.from(band.sourceAtOffset) : new Uint32Array()
}

it('matches the frozen kernel byte-for-byte across seeded mixed 1–8-key sorts', async () => {
  for (let seed = 1; seed <= 12; seed++) {
    const random = createPrng(seed * 717)
    const n = seed === 12 ? 5003 : 129 + seed
    const cells: FixtureCell[][] = Array.from({ length: n }, () => Array.from({ length: 8 }, () => {
      const value = values[Math.floor(random() * values.length)]
      return { value, styleIndex: typeof value === 'number' ? 3 : 0, hasCachedValue: true,
        valueIsError: typeof value === 'string' && value.startsWith('#') && random() < 0.5 }
    }))
    // Exact duplicates force source-row stable ties across merge boundaries.
    for (let row = 1; row < n; row += 7) cells[row] = cells[row - 1]
    const store = createFixture('F2', { rowCount: n }).createStore({ read: { cellAt: (row, col) => cells[row]?.[col] ?? null } })
    for (let count = 1; count <= 8; count++) for (const desc of [false, true]) {
      const spec: SortSpec = { range: { r0: 0, r1: n - 1, c0: 0, c1: 7, headerRows: seed % 2 as 0 | 1, totalsRows: (seed + 1) % 2 as 0 | 1 },
        keys: Array.from({ length: count }, (_, col) => ({ col, order: (col % 2 === 0) === desc ? 'desc' : 'asc' })), caseSensitive: seed % 2 === 0 }
      const options = { locale: ['en', 'sv', 'de', 'tr', 'th', 'zh'][seed % 6], now: () => 0, yield: async () => {},
        hidden: { hidden: (row: number) => row % 17 === 0 } as NonNullable<Parameters<typeof sortRowOrder>[2]>['hidden'] }
      const old = await baseline(store, spec, options), current = await sortRowOrder(store, spec, options)
      expect(current.kind, `seed ${seed}, ${count} keys, descending ${desc}`).toBe(old.kind)
      expect(permutation(current)).toEqual(permutation(old))
      if (current.kind === 'complete' && old.kind === 'complete') {
        expect(current.order.bands[0].displayOffsetOfSource).toEqual(old.order.bands[0].displayOffsetOfSource)
      }
    }
  }
}, 30_000)

async function textParity(values: readonly string[], locale: string, caseSensitive: boolean, order: 'asc' | 'desc') {
  const store = createFixture('F4', { rowCount: values.length }).createStore({ read: { cellAt: row => ({ value: values[row], styleIndex: 0, hasCachedValue: true, valueIsError: false }) } })
  const spec: SortSpec = { range: { r0: 0, r1: values.length - 1, c0: 0, c1: 0, headerRows: 0, totalsRows: 0 }, keys: [{ col: 0, order }], caseSensitive }
  const opts = { locale, now: () => 0, yield: async () => {} }
  expect(permutation(await sortRowOrder(store, spec, opts))).toEqual(permutation(await baseline(store, spec, opts)))
}
it('matches frozen text ranks with fast numeric keys, Unicode fallbacks and stable duplicates', async () => {
  const random = createPrng(8386)
  const values = Array.from({ length: 5003 }, (_, i) => `${['item-', 'alpha-', 'é', '😀', '\ud800', 'price-'][Math.floor(random() * 6)]}${Math.floor(random() * 1000000)}${i % 3 ? '' : '.02'}`)
  for (let i = 1; i < values.length; i += 13) values[i] = values[i - 1]
  for (const locale of ['en', 'de', 'sv', 'tr']) for (const cs of [false, true]) for (const direction of ['asc', 'desc'] as const) await textParity(values, locale, cs, direction)
}, 30_000)

it('verifies the observed dictionary against the host even when finite numeric certification misses a host difference', async () => {
  const Native = Intl.Collator
  const replacement = vi.spyOn(Intl, 'Collator').mockImplementation(function (locale, options) {
    const host = new Native(locale, options)
    return { resolvedOptions: () => host.resolvedOptions(), compare: (a: string, b: string) => {
      const first = /^item-(\d+)$/.exec(a), second = /^item-(\d+)$/.exec(b)
      return first && second && +first[1] >= 1000 && +second[1] >= 1000 ? +second[1] - +first[1] : host.compare(a, b)
    } } as Intl.Collator
  })
  try { await textParity(Array.from({ length: 160 }, (_, i) => `item-${1000 + i}`), 'en', false, 'asc') }
  finally { replacement.mockRestore() }
})
it('preserves the base admission limit for an empty protected body', async () => {
  const store = createFixture('F2', { rowCount: 3 }).createStore()
  const spec: SortSpec = { range: { r0: 0, r1: 0, c0: 0, c1: 7, headerRows: 1, totalsRows: 0 }, keys: [{ col: 0, order: 'asc' }], caseSensitive: false }
  const opts = { now: () => 0, yield: async () => {} }
  const old = await baseline(store, spec, opts)
  if (old.kind !== 'complete') throw new Error(old.kind)
  expect((await sortRowOrder(store, spec, { ...opts, maxWorkerBytes: old.ledger.peakBytes })).kind).toBe('complete')
})
