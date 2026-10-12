import { expect, it } from 'vitest'
import { mergeIndexRuns } from '../src/xlsx/view/sortMerge'
import { CooperativeSlice } from '../src/xlsx/view/scheduler'
import { createPrng } from './helpers/xlsx-viewer-fixtures'

async function sorted(values: readonly number[]) {
  let comparisons = 0, units = 0
  const slice = new CooperativeSlice({ checkEvery: 256, now: () => 0, yield: async () => {} })
  const due = slice.due.bind(slice)
  slice.due = n => { units += n ?? 1; return due(n) }
  const order = await mergeIndexRuns(Uint32Array.from(values, (_, i) => i), new Uint32Array(values.length),
    new Uint32Array(Math.ceil(values.length / 2) + 1), values.length,
    (a, b) => { comparisons++; return values[a] - values[b] || a - b }, slice, 'merge')
  return { order: [...order], comparisons, units }
}
it('stably merges descending, equal, uneven and seeded interleaved runs', async () => {
  const random = createPrng(128)
  for (const values of [[], [1], [4, 3, 2, 1], [1, 1, 0, 0, 1], Array.from({ length: 9003 }, () => Math.floor(random() * 47))]) {
    expect((await sorted(values)).order).toEqual(values.map((_, i) => i).sort((a, b) => values[a] - values[b] || a - b))
  }
})
it('takes linear comparison work for a presorted or reverse-sorted run', async () => {
  for (const values of [Array.from({ length: 8193 }, (_, i) => i), Array.from({ length: 8193 }, (_, i) => 8193 - i)]) {
    const result = await sorted(values)
    expect(result.comparisons).toBeGreaterThan(0)
    expect(result.comparisons).toBeLessThanOrEqual(values.length)
    expect(result.units).toBeGreaterThanOrEqual(values.length - 1)
    expect(result.order).toEqual(values.map((_, i) => i).sort((a, b) => values[a] - values[b]))
  }
})
it('checks within 256 work units while copying unpaired runs and reversing runs', async () => {
  let units = 0, lastCheck = 0
  class Checked extends CooperativeSlice {
    override due(count = 1) { units += count; return super.due(count) }
    override check() { expect(units - lastCheck).toBeLessThanOrEqual(256); lastCheck = units }
  }
  const values = [...Array.from({ length: 1000 }, (_, i) => 1000 + i), ...Array.from({ length: 1000 }, (_, i) => i), ...Array.from({ length: 1000 }, (_, i) => i - 1000)]
  const slice = new Checked({ checkEvery: 256, now: () => 0, yield: async () => {} })
  await mergeIndexRuns(Uint32Array.from(values, (_, i) => i), new Uint32Array(values.length), new Uint32Array(values.length), values.length,
    (a, b) => values[a] - values[b], slice, 'merge')
})
