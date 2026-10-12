import { expect, it } from 'vitest'
import { numericTextKeys, NUMERIC_ASCII } from '../src/xlsx/view/sortText'
import { CooperativeSlice } from '../src/xlsx/view/scheduler'
import { createPrng } from './helpers/xlsx-viewer-fixtures'

async function keys(values: string[], locale = 'en', caseSensitive = false, custom?: Intl.Collator) {
  let held = 0, calls = 0
  const host = custom ?? new Intl.Collator(locale, { numeric: true, sensitivity: caseSensitive ? 'variant' : 'accent', caseFirst: 'upper' })
  const collator = { resolvedOptions: () => host.resolvedOptions(), compare: (a: string, b: string) => { calls++; return host.compare(a, b) } } as Intl.Collator
  const result = await numericTextKeys(values, collator,
    { charge: bytes => { held += bytes }, release: bytes => { held -= bytes }, allocate: create => create() },
    new CooperativeSlice({ checkEvery: 256, now: () => 0, yield: async () => {} }), 'dictionary')
  const preparedCalls = calls
  return { ...result, host, held: () => held, calls: () => calls - preparedCalls }
}

it('uses certified decimal keys after identical ASCII prefixes and releases its charged lanes', async () => {
  const k = await keys(['item-2', 'item-10', 'item-02', 'long-9-' + 'y'.repeat(20), 'long-11-y'])
  expect(Math.sign(k.compare(0, 1))).toBe(-1)
  expect(Math.sign(k.compare(3, 4))).toBe(-1)
  expect(k.calls()).toBe(0)
  expect(k.held()).toBe(5 * 12)
  expect(k.compare(0, 2)).toBe(0)
  expect(k.calls()).toBe(1)
  k.dispose()
  expect(k.held()).toBe(0)
})

it('matches exhaustive single-unit and pairwise checks for its claimed ASCII set', async () => {
  for (const locale of ['en', 'sv', 'de', 'tr']) for (const cs of [false, true]) {
    const values = [...NUMERIC_ASCII, ...Array.from({ length: 100 }, (_, i) => String(i)),
      ...[...NUMERIC_ASCII].flatMap(c => Array.from({ length: 10 }, (_, i) => c + i))]
    const k = await keys(values, locale, cs)
    for (let a = 0; a < values.length; a++) for (let b = 0; b < values.length; b++) {
      if (Math.sign(k.compare(a, b)) !== Math.sign(k.host.compare(values[a], values[b]))) throw new Error(`${locale}/${cs}: ${JSON.stringify([values[a], values[b]])}`)
    }
    k.dispose()
  }
}, 30_000)

it('falls back for Unicode, long numeric runs, differing prefixes, numeric ties and uncertified hosts', async () => {
  const random = createPrng(829)
  const values = ['é2', 'é10', '😀2', '😀10', '\ud8002', '\udc0010', 'a0002x', 'a2y', 'A2', 'a10',
    '1'.repeat(255), '9'.repeat(254), 'a\u00002', 'a\u000010', 'x9007199254740993', 'x9007199254740992',
    ...Array.from({ length: 1000 }, () => `${['item-', 'a', 'price-', 'é', '😀'][Math.floor(random() * 5)]}${Math.floor(random() * 10000)}${['', '.2', '-3', 'y', 'é'][Math.floor(random() * 5)]}`)]
  const k = await keys(values)
  for (let i = 0; i < 10000; i++) {
    const a = Math.floor(random() * values.length), b = Math.floor(random() * values.length)
    expect(Math.sign(k.compare(a, b))).toBe(Math.sign(k.host.compare(values[a], values[b])))
  }
  k.dispose()
  const host = new Intl.Collator('en', { numeric: false })
  const fallback = await keys(['item-2', 'item-10'], 'en', false, host)
  expect(fallback.compare(0, 1)).toBe(host.compare('item-2', 'item-10'))
  expect(fallback.held()).toBe(0)
})
it('keeps long-string validation outside the fast path instead of scanning blobs', async () => {
  const values = ['long-9-' + 'y'.repeat(500), 'long-11-' + 'y'.repeat(400)]
  const k = await keys(values)
  expect(Math.sign(k.compare(0, 1))).toBe(Math.sign(k.host.compare(values[0], values[1])))
  expect(k.calls()).toBe(1)
  k.dispose()
})
