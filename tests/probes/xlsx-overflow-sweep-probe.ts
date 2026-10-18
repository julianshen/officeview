/// <reference lib="es2021.weakref" />
import { createDocumentOverflowCache } from '../../src/xlsx/view/documentOverflow'
import { MainMemoryLedger } from '../../src/xlsx/view/mainLedger'
import { bindOverflowCache } from '../../src/xlsx/view/overflowCache'
import { createRowOrder } from '../../src/xlsx/view/order'
import { buildSheetGeometry, rowOverflowInput } from '../../src/xlsx/view/viewport'

// Isolated child process: deliberately suppress automatic collection callbacks
// while actual weak references and GC still determine owner reachability.
const globals = globalThis as unknown as { FinalizationRegistry: typeof FinalizationRegistry; Bun: { gc(sync: boolean): void } }
globals.FinalizationRegistry = class {
  register(): void {}
  unregister(): boolean { return true }
} as unknown as typeof FinalizationRegistry
const ledger = new MainMemoryLedger(), cache = createDocumentOverflowCache(ledger, 'D')
const refs: Array<{ geometry: WeakRef<object>; metrics: WeakRef<object> }> = []
function make() {
  const count = 131072
  const order = createRowOrder([{ r0: 0, r1: count - 1, sourceAtOffset: Uint32Array.from({ length: count }, (_, row) => count - row - 1) }])
  const geometry = buildSheetGeometry({ name: 's', rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', styleIndex: 0, value: 'a' }] }], cols: [], merges: [], mergeRanges: [] }, { order })
  if (geometry.kind !== 'ready') throw Error('geometry')
  bindOverflowCache(geometry, cache)
  rowOverflowInput(geometry, 0)
  return geometry
}
function once() {
  const geometry = make()
  refs.push({ geometry: new WeakRef(geometry), metrics: new WeakRef(geometry.metrics) })
}
for (let i = 0; i < 8; i++) once()
for (let i = 0; i < 10; i++) { await new Promise(resolve => setTimeout(resolve, 0)); globals.Bun.gc(true); globals.Bun.gc(true) }
const snapshot = () => ({ entries: cache.pool.stats().entries, bytes: cache.bytes, pool: cache.pool.totalBytes, main: ledger.bytes })
const before = snapshot()
const deadGeometry = refs.filter(ref => ref.geometry.deref()).length, deadMetrics = refs.filter(ref => ref.metrics.deref()).length
const live = make() // A real miss, not a direct call to sweep().
const after = snapshot()
cache.close()
console.log(JSON.stringify({ deadGeometry, deadMetrics, before, after, closed: snapshot(), alive: live.kind === 'ready' }))
