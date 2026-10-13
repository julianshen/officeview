/// <reference lib="es2021.weakref" />
import { OverflowCache } from '../../src/xlsx/view/overflowCache'
import { createRowOrder } from '../../src/xlsx/view/order'
import { MemoryPool } from '../../src/xlsx/view/tileCache'
import type { XlsxCell } from '../../src/xlsx/types'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
const bun = (globalThis as unknown as { Bun: { gc(sync: boolean): void } }).Bun
const cache = new OverflowCache(new MemoryPool({ limitBytes: 2 * 1024 * 1024 }))
const refs: Array<{ geometry: WeakRef<object>; metrics: WeakRef<object> }> = []
function once(): void {
  const n = 131072
  const order = createRowOrder([{ r0: 0, r1: n - 1, sourceAtOffset: Uint32Array.from({ length: n }, (_, i) => n - 1 - i) }])
  const g = buildSheetGeometry({ name: 's', rows: [{ index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 'a', styleIndex: 0 }] }], cols: [], merges: [], mergeRanges: [] }, { order })
  if (g.kind !== 'ready') throw Error('geometry')
  cache.get(g, 0)
  refs.push({ geometry: new WeakRef(g), metrics: new WeakRef(g.metrics) })
}
let droppedGeometry: ReturnType<typeof buildSheetGeometry> | undefined
let row: XlsxCell[] | undefined
if (process.argv[2] === 'drop') {
  const built = buildSheetGeometry({ name: 's', rows: [], cols: [], merges: [], mergeRanges: [] })
  if (built.kind !== 'ready') throw Error('geometry')
  row = [{ row: 0, col: 10, ref: 'K1', styleIndex: 0, value: 'block' }]
  droppedGeometry = { ...built, cellsOf: () => row }
  cache.get(droppedGeometry, 0)
  row = undefined
} else for (let i = 0; i < 8; i++) once()
for (let i = 0; i < 10; i++) {
  await new Promise(resolve => setTimeout(resolve, 0))
  bun.gc(true); bun.gc(true)
}
if (droppedGeometry?.kind === 'ready') console.log(JSON.stringify({ blockers: cache.get(droppedGeometry, 0).blockers.length }))
else console.log(JSON.stringify({ geometries: refs.filter(ref => ref.geometry.deref()).length, metrics: refs.filter(ref => ref.metrics.deref()).length, bytes: cache.bytes, pool: cache.pool.stats().totalBytes }))
