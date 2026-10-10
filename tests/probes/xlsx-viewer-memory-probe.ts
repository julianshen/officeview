// Child-process probe for tests/xlsx-viewer-memory.test.ts. Run under `bun` because the
// vitest worker cannot expose gc: Bun.gc(true) is synchronous and
// process.memoryUsage().external counts typed-array backing stores (Bun reports 0 for
// arrayBuffers). Prints one JSON line of retained-byte deltas measured after a forced gc,
// with every caller-owned input already dropped, so an accidental retention of an input
// buffer or an allocation outside the modules' own counters shows up.
import { createAxisMetrics } from '../../src/xlsx/view/axes'
import { MAX_COLS, MAX_ROWS } from '../../src/xlsx/view/coordinates'
import { createRowOrder } from '../../src/xlsx/view/order'
import { createSheetMetrics } from '../../src/xlsx/view/projection'
import { createRowVisibility } from '../../src/xlsx/view/visibility'

const bun = (globalThis as unknown as { Bun: { gc(sync: boolean): void } }).Bun
function settle(): NodeJS.MemoryUsage {
  bun.gc(true)
  bun.gc(true)
  return process.memoryUsage()
}
function shuffled(length: number, seed: number): Uint32Array {
  const rows = Uint32Array.from({ length }, (_, i) => i)
  let state = seed
  for (let i = length - 1; i > 0; i--) {
    state = (state * 1664525 + 1013904223) % 4294967296
    const j = state % (i + 1)
    const value = rows[i]; rows[i] = rows[j]; rows[j] = value
  }
  return rows
}

const scenario = process.argv[2]
const before = settle()
let keep: unknown
if (scenario === 'identity') {
  keep = createSheetMetrics({ rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24 }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }) })
} else if (scenario === 'order') {
  // The input is built inline and never named: only a retained closure could keep it alive.
  keep = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1, sourceAtOffset: shuffled(MAX_ROWS, 23) }])
} else if (scenario === 'projection') {
  const half = MAX_ROWS / 2
  const order = createRowOrder([{ r0: 0, r1: MAX_ROWS - 1,
    sourceAtOffset: Uint32Array.from({ length: MAX_ROWS }, (_, slot) => slot % 2 ? slot / 2 - 0.5 : half + slot / 2) }])
  const visibility = createRowVisibility()
  const sheet = createSheetMetrics({
    rows: createAxisMetrics({ logicalCount: MAX_ROWS, defaultSizeCss: 24, overrides: [{ start: half, end: MAX_ROWS - 1, sizeCss: 48 }] }),
    cols: createAxisMetrics({ logicalCount: MAX_COLS, defaultSizeCss: 64 }), order, visibility })
  keep = [order, visibility, sheet]
} else throw new Error(`Unknown scenario ${scenario}`)
const after = settle()
console.log(JSON.stringify({ scenario, externalBytes: after.external - before.external, heapBytes: after.heapUsed - before.heapUsed, alive: keep !== undefined }))
