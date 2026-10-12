import { expect, it } from 'vitest'
import { compactPreparedPage, preparedRowAt, preparedMetadataAt, type RawPreparedPage } from '../src/worker/preparedColumn'
import { CooperativeSlice } from '../src/xlsx/view/scheduler'

function rawPage(values: readonly number[], sparse = false): RawPreparedPage {
  const n = values.length
  return { count: n, rows: Uint32Array.from(values, (_, i) => sparse ? i * 3 : i + 500),
    numbers: Float64Array.from(values), tags: new Uint8Array(n).fill(1), stringIds: new Uint32Array(n),
    styles: new Uint32Array(n).fill(4294967295), flags: new Uint8Array(n).fill(127) }
}
const quiet = () => new CooperativeSlice({ now: () => 0, yield: async () => {} })
it('packs dense rows, numeric widths and constant metadata without retaining redundant lanes', async () => {
  const raw = rawPage(Array.from({ length: 4096 }, (_, i) => i))
  const page = (await compactPreparedPage(raw, quiet(), () => false))!
  expect(page.rows).toBeUndefined()
  expect(page.values).toBeInstanceOf(Uint16Array)
  expect(page.laneBytes).toBeLessThanOrEqual(raw.count * 2 + 6)
  expect(page.metadataIds).toBeUndefined()
  expect(page.styles).toEqual(Uint32Array.of(4294967295))
  raw.numbers[0] = 999
  expect(page.values[0]).toBe(0)
})
it('round-trips scalar values exactly including negative zero and non-Float32 fractions', async () => {
  for (const values of [[0, 1, 255], [0, 256, 65535], [0, 65536, 4294967295], [-1, -2147483648, 7], [-0, 0, 1.5], [0.1, Math.PI, 4294967296.25], [NaN, Infinity, -Infinity]]) {
    const raw = rawPage(values, true), page = (await compactPreparedPage(raw, quiet(), () => false))!
    for (let i = 0; i < values.length; i++) {
      expect(Object.is(page.values[i], values[i])).toBe(true)
      expect(preparedRowAt(page, i)).toBe(raw.rows[i])
      const m = preparedMetadataAt(page, i)
      expect(page.styles[m]).toBe(raw.styles[i])
      expect(page.codes[m]).toBe(raw.flags[i] * 4 + raw.tags[i])
    }
  }
})
it('dictionary-encodes repeated metadata, with a dense fallback for many styles', async () => {
  for (const cardinality of [2, 100]) {
    const raw = rawPage(Array.from({ length: 1000 }, (_, i) => i))
    for (let i = 0; i < raw.count; i++) { raw.styles[i] = i % cardinality; raw.flags[i] = i % 2 ? 127 : 0; raw.tags[i] = i % 2 ? 2 : 1; raw.stringIds[i] = 9 }
    const page = (await compactPreparedPage(raw, quiet(), () => false))!
    expect(page.flatMetadata).toBe(cardinality > 16)
    for (let i = 0; i < raw.count; i++) {
      const m = preparedMetadataAt(page, i)
      expect(page.styles[m]).toBe(raw.styles[i])
      expect(page.codes[m]).toBe(raw.flags[i] * 4 + raw.tags[i])
      expect(page.values[i]).toBe(raw.tags[i] === 2 ? 9 : i)
    }
  }
})
it('interrupts page compaction at a bounded task pause', async () => {
  let halted = false
  const page = await compactPreparedPage(rawPage(Array.from({ length: 4096 }, (_, i) => i)),
    new CooperativeSlice({ now: () => 0, yield: async () => { halted = true } }), () => halted)
  expect(page).toBeUndefined()
})

it('walks a prepared page with linear row-lane reads and exact blank/membership boundaries', async () => {
  const { visitPreparedRows } = await import('../src/worker/preparedColumn')
  let reads = 0
  const items: unknown[] = []
  visitPreparedRows(1000, at => { reads++; return at }, [{ r0: 0, r1: 999 }], 1000,
    ordinal => items.push(ordinal), (r0, r1) => items.push({ r0, r1 }))
  expect(items).toEqual(Array.from({ length: 1000 }, (_, i) => i))
  expect(reads).toBeLessThanOrEqual(1064)
  const sparse: unknown[] = [], rows = [2, 3, 9]
  visitPreparedRows(rows.length, at => rows[at], [{ r0: 0, r1: 5 }, { r0: 8, r1: 10 }], 7,
    ordinal => sparse.push(ordinal), (r0, r1) => sparse.push({ r0, r1 }))
  expect(sparse).toEqual([{ r0: 0, r1: 1 }, 0, 1, { r0: 4, r1: 5 }, { r0: 8, r1: 8 }, 2, { r0: 10, r1: 10 }])
})
