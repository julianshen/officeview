/**
 * S13.5 Bun microbenchmark (non-browser): the worker filter evaluator over a
 * 1M-row object source with a numeric (0.00), a date (numFmt 14) and a mixed
 * column. Numeric/date filters must not format display text; a text rule on the
 * mixed column builds the display lane on its first apply and reuses it after.
 *
 *   bun scripts/bench-display-lane.ts [repo-root] [label]
 *
 * Pass another checkout's root to measure it with the same harness. The heap
 * delta is transient allocation around one apply, not retained memory.
 */
import { resolve } from 'node:path'
const root = resolve(process.argv[2] ?? '.')
const label = process.argv[3] ?? 'run'
const { createObjectWorkbookSource } = await import(`${root}/src/xlsx/view/source.ts`)
const { createWorkerFilterEvaluator } = await import(`${root}/src/worker/viewerFilterEvaluator.ts`)

const ROWS = 1_000_000
const styles = [{ numFmtId: 0 }, { numFmtId: 2 }, { numFmtId: 14 }]
const rows: unknown[] = [{ index: 0, cells: [
  { ref: '', row: 0, col: 0, value: 'amount', styleIndex: 0, style: styles[0] },
  { ref: '', row: 0, col: 1, value: 'day', styleIndex: 0, style: styles[0] },
  { ref: '', row: 0, col: 2, value: 'mixed', styleIndex: 0, style: styles[0] },
] }]
for (let r = 1; r <= ROWS; r++) rows.push({ index: r, cells: [
  { ref: '', row: r, col: 0, value: (r * 7919) % 100000 / 100, styleIndex: 1, style: styles[1] },
  { ref: '', row: r, col: 1, value: 40000 + (r % 5000), styleIndex: 2, style: styles[2] },
  r === 1 ? { ref: '', row: r, col: 2, value: 'text', styleIndex: 0, style: styles[0] } : { ref: '', row: r, col: 2, value: (r * 7919) % 100000 / 100, styleIndex: 1, style: styles[1] },
] })
const doc = { sheets: [{ name: 'S', sheetId: '1', workbookIndex: 0, cols: [], merges: [], mergeRanges: [], rows }], tables: [] }
const source = createObjectWorkbookSource(doc as never, 'bench')
const store = source.store('1')
const formatting = { dateSystem: '1900' as const, styles }
const range = { r0: 0, r1: ROWS, c0: 0, c1: 2, headerRows: 1, totalsRows: 0 }
const cases = [
  ['number gt', { range, columns: new Map([[0, { kind: 'number', op: 'gt', first: 500 }]]) }],
  ['date between', { range, columns: new Map([[1, { kind: 'date', op: 'between', first: { year: 2010, month: 1, day: 1 }, second: { year: 2011, month: 12, day: 31 } }]]) }],
] as const
for (const [name, spec] of cases) {
  const samples: number[] = []
  let heapDelta = 0
  for (let i = 0; i < 4; i++) {
    // Fresh evaluator per sample: classification cache (if any) is cold.
    const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting)
    Bun.gc(true)
    const before = process.memoryUsage().heapUsed
    const t0 = performance.now()
    const result = await evaluate({ store, spec: spec as never, nowMs: Date.UTC(2026, 9, 13), signal: new AbortController().signal })
    samples.push(performance.now() - t0)
    if (result.kind !== 'complete') throw new Error(`${name}: ${JSON.stringify(result)}`)
    heapDelta = Math.max(heapDelta, process.memoryUsage().heapUsed - before)
  }
  samples.sort((a, b) => a - b)
  console.log(JSON.stringify({ label, case: name, rows: ROWS, msSamples: samples.map(v => Math.round(v)), medianMs: Math.round((samples[1] + samples[2]) / 2), maxHeapDeltaMiB: +(heapDelta / 1048576).toFixed(1) }))
}

{
  // Text contains on a mixed column: first apply (lane build) vs repeat apply (lane reuse), one evaluator.
  const spec = { range, columns: new Map([[2, { kind: 'text', op: 'contains', first: '25', wildcards: false }]]) }
  const evaluate = createWorkerFilterEvaluator(() => new Map(), formatting)
  const times: number[] = []
  for (let i = 0; i < 3; i++) {
    const t0 = performance.now()
    const result = await evaluate({ store, spec: spec as never, nowMs: Date.UTC(2026, 9, 13), signal: new AbortController().signal })
    times.push(Math.round(performance.now() - t0))
    if (result.kind !== 'complete') throw new Error(JSON.stringify(result))
  }
  console.log(JSON.stringify({ label, case: 'text contains mixed (apply 1, 2, 3)', rows: ROWS, ms: times }))
}
