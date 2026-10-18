// @vitest-environment node
import { expect, it, vi } from 'vitest'
import { RowCacheLedger } from '../src/worker/viewerRemoteSource'

function harness(limit: number) {
  const ledger = new RowCacheLedger(limit), dropped: number[] = []
  const store = { sheetId: 's', drop: (row: number) => dropped.push(row) } as unknown as Parameters<RowCacheLedger['charge']>[0]
  return { ledger, store, dropped }
}

/** Count resident work regardless of which Map traversal API admission uses. */
function residentWork(ledger: RowCacheLedger) {
  const entries = (ledger as unknown as { entries: Map<string, { bytes: number }> }).entries
  let visited = 0
  const iterate = entries[Symbol.iterator].bind(entries), pairs = entries.entries.bind(entries)
  const keys = entries.keys.bind(entries), values = entries.values.bind(entries), each = entries.forEach.bind(entries)
  vi.spyOn(entries, Symbol.iterator).mockImplementation(function* () { for (const item of iterate()) { visited++; yield item }; return undefined })
  vi.spyOn(entries, 'entries').mockImplementation(function* () { for (const item of pairs()) { visited++; yield item }; return undefined })
  vi.spyOn(entries, 'keys').mockImplementation(function* () { for (const item of keys()) { visited++; yield item }; return undefined })
  vi.spyOn(entries, 'values').mockImplementation(function* () { for (const item of values()) { visited++; yield item }; return undefined })
  vi.spyOn(entries, 'forEach').mockImplementation((callback, receiver) => each((value, key, map) => { visited++; callback.call(receiver, value, key, map) }))
  const get = vi.spyOn(entries, 'get')
  return { visited: () => visited, lookups: () => get.mock.calls.length }
}

it('SLICE10 zero-growth one-row admission does no resident traversal in a 262144-row cache', () => {
  const { ledger, store, dropped } = harness(16 * 1024 * 1024)
  for (let row = 0; row < 262144; row++) ledger.charge(store, row, 64)
  const work = residentWork(ledger)
  try {
    const token = ledger.reserve(0, store, [0])!
    expect(token).toBeDefined()
    expect(work.visited()).toBe(0)
    expect(work.lookups()).toBe(0)
    expect(dropped).toEqual([])
    expect(ledger.bytes).toBe(16 * 1024 * 1024)
    token.close()
  } finally { vi.restoreAllMocks() }
})

it.each([1024, 65536])('SLICE10 near-full positive-growth admission only reads active pins with %s resident rows', count => {
  const { ledger, store, dropped } = harness(count * 256)
  for (let row = 0; row < count; row++) ledger.charge(store, row, 256)
  const first = ledger.reserve(64, store, [0])!
  expect(first).toBeDefined()
  const work = residentWork(ledger)
  try {
    // Resident cache is full: the fast-fit path cannot admit either promise.
    const second = ledger.reserve(128, store, [1])!
    expect(second).toBeDefined()
    expect(work.visited()).toBe(0)
    expect(work.lookups()).toBeLessThanOrEqual(2)
    expect(dropped).toEqual([])
    expect(ledger.bytes).toBe(count * 256)
    second.close()
    first.close()
  } finally { vi.restoreAllMocks() }
})

it('SLICE10 overlapping duplicate pins count each resident charge once at exact capacity', () => {
  const { ledger, store, dropped } = harness(1000)
  ledger.charge(store, 0, 300)
  ledger.charge(store, 1, 200)
  ledger.charge(store, 2, 100)
  ledger.charge(store, 3, 400)
  const first = ledger.reserve(100, store, [0, 0, 1])!
  const second = ledger.reserve(100, store, [1, 2, 2])!
  expect(first).toBeDefined()
  expect(second).toBeDefined()
  const work = residentWork(ledger)
  try {
    // Union of all pins is600B, prior promises200B, this promise200B.
    const exact = ledger.reserve(200, store, [0, 2, 0])!
    expect(exact).toBeDefined()
    expect(work.visited()).toBe(0)
    expect(work.lookups()).toBe(3)
    expect(ledger.reserve(1, store, [0, 1, 2])).toBeUndefined()
    expect(dropped).toEqual([])
    expect(ledger.bytes).toBe(1000)
    exact.close()
    second.close()
    first.close()
    const all = ledger.reserve(1000)!
    expect(all).toBeDefined()
    all.close()
  } finally { vi.restoreAllMocks() }
})

it('SLICE10 prior overlapping pins remain protected when the new promise has no pins', () => {
  const { ledger, store, dropped } = harness(1000)
  ledger.charge(store, 0, 300)
  ledger.charge(store, 1, 200)
  ledger.charge(store, 2, 500)
  const first = ledger.reserve(100, store, [0, 1])!
  const second = ledger.reserve(100, store, [0, 1])!
  expect(first).toBeDefined()
  expect(second).toBeDefined()
  const work = residentWork(ledger)
  try {
    expect(ledger.reserve(301)).toBeUndefined()
    expect(work.visited()).toBe(0)
    expect(work.lookups()).toBe(2)
    const exact = ledger.reserve(300)!
    expect(exact).toBeDefined()
    expect(dropped).toEqual([])
    exact.close()
    first.close()
    // Closing one overlapping owner must retain the other's protection.
    expect(ledger.reserve(401)).toBeUndefined()
    second.close()
    const all = ledger.reserve(1000)!
    expect(all).toBeDefined()
    all.close()
  } finally { vi.restoreAllMocks() }
})
