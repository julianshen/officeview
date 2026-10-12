import { expect, it } from 'vitest'
import { buildSortPrefix } from '../src/xlsx/view/sortPrefix'
import type { SortLane } from '../src/xlsx/view/sort'
import { CooperativeSlice } from '../src/xlsx/view/scheduler'
import { createPrng } from './helpers/xlsx-viewer-fixtures'

function lane(tags: number[], numbers: number[]): SortLane {
  return { tags: Uint8Array.from(tags), numbers: Float64Array.from(numbers), strings: [], errors: [], stringRanks: new Uint32Array(), errorRanks: new Uint32Array() }
}
function memory(limit = 1e9) {
  let held = 0
  return { canCharge: (n: number) => held + n <= limit, charge: (n: number) => { held += n }, release: (n: number) => { held -= n }, allocate: <T>(create: () => T) => create(), held: () => held }
}
const slice = () => new CooperativeSlice({ checkEvery: 256, now: () => 0, yield: async () => {} })
it('encodes tags and values in exact lexicographic order, keeping ties equal', async () => {
  const random = createPrng(8613), n = 1000
  const lanes = Array.from({ length: 4 }, (_, key) => lane(Array.from({ length: n }, () => 1 + Math.floor(random() * 5)), Array.from({ length: n }, () => Math.floor(random() * (key + 2)) - 2)))
  const mem = memory(), prefix = await buildSortPrefix(lanes, n, mem, slice(), 'read')
  expect(prefix).toBeDefined()
  if (!prefix) throw new Error('missing prefix')
  expect(prefix.keys).toBe(4)
  for (let a = 0; a < n; a++) for (let b = 0; b < n; b += 19) {
    let comparison = 0
    for (const lane of lanes) { comparison = lane.tags[a] - lane.tags[b] || lane.numbers[a] - lane.numbers[b]; if (comparison) break }
    expect(Math.sign(prefix.codes[a] - prefix.codes[b])).toBe(Math.sign(comparison))
  }
  expect(mem.held()).toBe(prefix.bytes)
  mem.release(prefix.bytes)
  expect(mem.held()).toBe(0)
})
it('stops before fractional/oversized lanes and never wraps the 32-bit prefix', async () => {
  const small = lane([1, 1, 1], [0, 1, 2]), wide = lane([1, 1, 1], [0, 4294967295, 1]), fractional = lane([1, 1, 1], [0, 0.1, 1])
  const mem = memory()
  const prefix = await buildSortPrefix([small, small, fractional, small], 3, mem, slice(), 'read')
  expect(prefix?.keys).toBe(2)
  const none = memory()
  expect(await buildSortPrefix([wide, small, small, small], 3, none, slice(), 'read')).toBeUndefined()
  expect(none.held()).toBe(0)
  const tight = memory(1)
  expect(await buildSortPrefix([small, small, small, small], 3, tight, slice(), 'read')).toBeUndefined()
  expect(tight.held()).toBe(0)
})
it('keeps checks within 256 units and work yields within 4096 while building a large prefix', async () => {
  let units = 0, checked = 0, paused = 0
  class Checked extends CooperativeSlice {
    override due(n = 1) { units += n; return super.due(n) }
    override check() { expect(units - checked).toBeLessThanOrEqual(256); checked = units }
  }
  const s = new Checked({ checkEvery: 256, now: () => 0, yield: async () => { expect(units - paused).toBeLessThanOrEqual(4096); paused = units } })
  const l = lane(Array(9003).fill(1), Array.from({ length: 9003 }, (_, i) => i % 3))
  expect(await buildSortPrefix([l, l, l, l], 9003, memory(), s, 'read')).toBeDefined()
  expect(paused).toBeGreaterThan(0)
})
it('reclaims prefix staging when cancellation interrupts code construction', async () => {
  const n = 9003, l = lane(Array(n).fill(1), Array.from({ length: n }, (_, i) => i % 3)), mem = memory()
  let pauses = 0
  await expect(buildSortPrefix([l, l, l, l], n, mem,
    new CooperativeSlice({ checkEvery: 256, now: () => 0, yield: async () => { if (++pauses === 9) throw new Error('cancelled') } }), 'read')).rejects.toThrow('cancelled')
  expect(mem.held()).toBe(0)
})
