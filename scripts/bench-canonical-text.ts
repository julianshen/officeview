/** Run: bun scripts/bench-canonical-text.ts
 * Non-browser (Bun/Node) per-value cost of canonical text matching (plan S8.5).
 * Host microbenchmark only: it is not a real-browser measurement and sets no
 * threshold. Prints one JSON row per pool/length/operation with the matcher's
 * charged work units and the mean microseconds per evaluated value. */
import { TextFilterMatcher, WorkBudget } from '../src/xlsx/view/textFilter'

const pools: Record<string, readonly string[]> = {
  latin: Array.from('abcdefghijklmnopqrstuvwxyz ABCDEFGHIJ.,'),
  german: ['ß', 'Straße ', 'e\u0301', 'ü', 'u\u0308', 'ﬁ', ' ', 'Ä', 'x'],
  cjk: Array.from({ length: 150 }, (_, index) => String.fromCodePoint(0x4e00 + index * 7)),
  emoji: ['👍🏽', '🇺🇸', '👨\u200d👩\u200d👧', '😀', ' ', 'a'],
}
function sample(pool: readonly string[], length: number): string {
  let seed = 20261011, text = ''
  while (text.length < length) { seed = seed * 48271 % 2147483647; text += pool[seed % pool.length] }
  return text.slice(0, length).replace(/[\ud800-\udbff]$/, 'x')
}
const host = process.versions.bun ? `Bun ${process.versions.bun}` : `Node ${process.versions.node}`
const operations = [
  ['contains-miss', 'contains', () => 'zzz', false],
  ['contains-hit', 'contains', (text: string) => text.slice(text.length - 12, text.length - 4), false],
  ['equals', 'eq', (text: string) => text.toUpperCase(), false],
  ['wildcard', 'eq', () => '*e*?*z', true],
] as const
for (const [name, pool] of Object.entries(pools)) {
  for (const length of [100, 300, 1000, 2000]) {
    const text = sample(pool, length)
    for (const [label, op, pattern, wildcards] of operations) {
      const matcher = new TextFilterMatcher('en-US'), work = { used: 0 }
      const predicate = matcher.prepare(op, pattern(text), wildcards, new WorkBudget())
      predicate(text, new WorkBudget({ work }))
      for (let i = 0; i < 200; i++) predicate(text, new WorkBudget())
      // Median of five timed trials; each trial evaluates the same value repeatedly.
      const repeats = Math.max(200, Math.round(400000 / length)), trials: number[] = []
      for (let trial = 0; trial < 5; trial++) {
        const started = performance.now()
        for (let i = 0; i < repeats; i++) predicate(text, new WorkBudget())
        trials.push((performance.now() - started) * 1000 / repeats)
      }
      const us = trials.sort((a, b) => a - b)[2]
      console.log(JSON.stringify({ host: `${host} (non-browser)`, pool: name, length, op: label, work: work.used, usPerValue: Number(us.toFixed(2)) }))
    }
  }
}
