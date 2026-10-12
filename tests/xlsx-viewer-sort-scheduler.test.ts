import { expect, it } from 'vitest'
import { CooperativeSlice, YIELD_WORK_UNITS } from '../src/xlsx/view/scheduler'

it('amortizes clock and owner checks to 256 units without exceeding the work yield bound', async () => {
  let units = 0, clockSamples = 0
  const checked: number[] = [], paused: number[] = []
  class Checked extends CooperativeSlice { override check() { checked.push(units) } }
  const slice = new Checked({ checkEvery: 256, now: () => { clockSamples++; return 0 }, yield: async () => { paused.push(units) } })
  for (units = 1; units <= YIELD_WORK_UNITS * 2; units++) if (slice.due()) await slice.pause('merge')
  expect(checked.slice(0, 3)).toEqual([256, 512, 768])
  expect(paused).toEqual([4096, 8192])
  expect(clockSamples).toBe(1 + 32 + 2)
  for (let i = 1; i < checked.length; i++) expect(checked[i] - checked[i - 1]).toBeLessThanOrEqual(256)
})

it('checks at both sides of every pause and at the first batch after a clock deadline', async () => {
  let units = 0, clock = 0
  const checked: number[] = []
  class Checked extends CooperativeSlice { override check() { checked.push(units) } }
  const slice = new Checked({ checkEvery: 256, now: () => clock, yield: async () => { clock += 100 } })
  for (units = 1; units < 256; units++) { clock = units >= 128 ? 10 : 0; expect(slice.due()).toBe(false) }
  expect(slice.due()).toBe(true)
  await slice.pause('merge')
  expect(checked).toEqual([256, 256, 256])
  units++
  expect(slice.due()).toBe(false)
  await slice.flush('merge')
  expect(checked.slice(-2)).toEqual([257, 257])
})
