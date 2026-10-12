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
import { PageBuilder, K_EMPTY, K_NUMBER, K_TEXT, summarize, type PageSummary } from '../src/xlsx/storage/sourcePages'
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
type Case = (n: number) => { keep: unknown; model: number }

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
    const before = live()
    const { keep, model } = run(count)
    const observed = live() - before
    sink.push(keep)
    results.push({ name, count, observed, model })
    sink.length = 0
  }
}
console.log(JSON.stringify(results))
