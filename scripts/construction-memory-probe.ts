/**
 * Bun heap probe for the construction-memory model (S13.4, spec §12.4).
 *
 *   bun scripts/construction-memory-probe.ts [case ...]
 *
 * Each case builds one kind of construction structure, keeps it reachable,
 * forces a full GC and reports the live heap delta (heapSize + extraMemorySize
 * from bun:jsc) next to the model's charge for the same structure. Sizes include
 * array-capacity transitions (just past a power of two and the review's 98,304).
 * Prints one JSON line: [{ name, count, observed, model }]. Bun/JSC only; it is
 * not a Chrome heap measurement.
 */
import { heapStats } from 'bun:jsc'
import {
  DECODED_CHUNK_OVERHEAD_BYTES, DECODED_PAGE_OVERHEAD_BYTES, PageBuilder, K_EMPTY, K_NUMBER, K_TEXT, decodePage, decodeStrings, encodePage, encodeStrings,
  summarize, type PageSummary,
} from '../src/xlsx/storage/sourcePages'
import { MemoryPool } from '../src/xlsx/view/tileCache'
import { PAGE_SUMMARY_BYTES, PENDING_STRING_BYTES } from '../src/xlsx/storage/constructionMemory'

declare const Bun: { gc(force: boolean): void }
const live = (): number => {
  Bun.gc(true)
  Bun.gc(true)
  const s = heapStats() as { heapSize: number; extraMemorySize: number }
  return s.heapSize + s.extraMemorySize
}
/** A flat (non-rope) string, like the tokenizer's output. */
const flat = (s: string): string => { s.charCodeAt(s.length - 1); return s }

interface Result { name: string; count: number; observed: number; model: number }
/** `baseline`: a heap reading taken inside the case after setup that must not be counted. */
type Case = (n: number) => { keep: unknown; model: number; baseline?: number }

const CASES: Record<string, { sizes: number[]; run: Case }> = {
  'builder-numeric': {
    sizes: [65_537, 98_304, 131_073],
    run: (n) => {
      const b = new PageBuilder()
      for (let i = 0; i < n; i++) b.add(i >> 3, i & 7, K_NUMBER, 0, i * 1.5, '', true, undefined)
      return { keep: b, model: b.memoryBytes }
    },
  },
  'builder-text': {
    sizes: [65_537, 100_000],
    run: (n) => {
      const b = new PageBuilder()
      for (let i = 0; i < n; i++) b.add(i >> 3, i & 7, K_TEXT, 0, 0, flat(`item-${i}-漢`), true, undefined)
      return { keep: b, model: b.memoryBytes }
    },
  },
  'builder-formula': {
    sizes: [50_000],
    run: (n) => {
      const b = new PageBuilder()
      for (let i = 0; i < n; i++) b.add(i, 0, K_EMPTY, 0, 0, '', false, { formula: flat(`SUM(A${i}:B${i})*2`) })
      return { keep: b, model: b.memoryBytes }
    },
  },
  'summaries': {
    sizes: [65_537, 100_000],
    run: (n) => {
      const payload = new PageBuilder().build()
      const list: PageSummary[] = []
      for (let i = 0; i < n; i++) list.push(summarize('1', flat(`p${i}`), i * 10, i * 10 + 9, payload, 1024))
      return { keep: list, model: n * PAGE_SUMMARY_BYTES }
    },
  },
  // Decoded pages/chunks viewing buffers a memory backend already holds (the
  // buffers are created before the measurement): the decoded objects, a
  // resident-cache record and a pool entry each, as SourcePageStore keeps them.
  'decoded-pages': {
    sizes: [1000, 5000],
    run: (n) => {
      const b = new PageBuilder()
      b.add(0, 0, K_NUMBER, 0, 1, '', true, undefined)
      const held = Array.from({ length: n }, () => encodePage(b.build()))
      const before = live()
      const pool = new MemoryPool({ limitBytes: 1 << 30 })
      const resident = new Map<string, { value: unknown; bytes: number; entry: unknown }>()
      held.forEach((bytes, i) => {
        const key = flat(`page/1/p${i}`)
        resident.set(key, { value: decodePage(bytes), bytes: DECODED_PAGE_OVERHEAD_BYTES, entry: pool.retain(key, { tier: 'source', bytes: DECODED_PAGE_OVERHEAD_BYTES }) })
      })
      return { keep: [held, resident, pool], model: n * DECODED_PAGE_OVERHEAD_BYTES, baseline: before }
    },
  },
  'decoded-chunks': {
    sizes: [1000, 5000],
    run: (n) => {
      const held = Array.from({ length: n }, (_, i) => encodeStrings([flat(`v${i}`)]))
      const before = live()
      const pool = new MemoryPool({ limitBytes: 1 << 30 })
      const resident = new Map<string, { value: unknown; bytes: number; entry: unknown }>()
      held.forEach((bytes, i) => {
        const key = flat(`blob/${i}`)
        resident.set(key, { value: decodeStrings(bytes), bytes: DECODED_CHUNK_OVERHEAD_BYTES, entry: pool.retain(key, { tier: 'source', bytes: DECODED_CHUNK_OVERHEAD_BYTES }) })
      })
      return { keep: [held, resident, pool], model: n * DECODED_CHUNK_OVERHEAD_BYTES, baseline: before }
    },
  },
  'pending-strings': {
    sizes: [65_537, 100_000],
    run: (n) => {
      const chunks: string[][] = [[]]
      let model = 0
      for (let i = 0; i < n; i++) {
        const s = flat(`shared string ${i}`)
        model += 2 * s.length + PENDING_STRING_BYTES
        if (chunks[chunks.length - 1].length === 4096) chunks.push([])
        chunks[chunks.length - 1].push(s)
      }
      return { keep: chunks, model }
    },
  },
}

const wanted = process.argv.slice(2)
const results: Result[] = []
const sink: unknown[] = []
for (const [name, { sizes, run }] of Object.entries(CASES)) {
  if (wanted.length && !wanted.includes(name)) continue
  for (const count of sizes) {
    // Warm the code path so JIT/compiler allocations are not attributed.
    sink.push(run(Math.min(count, 2048)).keep)
    sink.length = 0
    const start = live()
    const { keep, model, baseline } = run(count)
    const observed = live() - (baseline ?? start)
    sink.push(keep)
    results.push({ name, count, observed, model })
    sink.length = 0
  }
}
console.log(JSON.stringify(results))
