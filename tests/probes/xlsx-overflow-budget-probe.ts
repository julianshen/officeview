// Isolated Bun/JSC ownership measurement. heapSize already includes external
// backing stores; adding extraMemorySize again would double-count them.
import { OverflowCache } from '../../src/xlsx/view/overflowCache'
import { MemoryPool } from '../../src/xlsx/view/tileCache'
import { buildSheetGeometry } from '../../src/xlsx/view/viewport'
const { heapStats } = await import('bun:' + 'jsc') as { heapStats(): { heapSize: number; extraMemorySize: number } }
const bun = (globalThis as unknown as { Bun: { gc(sync: boolean): void } }).Bun
const count = Number(process.argv[2] ?? 1000)
const width = Number(process.argv[3] ?? 1)
const built = buildSheetGeometry({ name: 's', rows: [], cols: [], merges: [], mergeRanges: [] })
if (built.kind !== 'ready') throw Error('geometry')
const geometry = { ...built, cellsOf: (row: number) => Array.from({ length: width }, (_, col) => ({ row, col, ref: '', styleIndex: 0, value: 1 })) }
const make = () => new OverflowCache(new MemoryPool({ limitBytes: 64 * 1024 * 1024, highWatermark: 1, lowWatermark: 1 }), width === 1 ? 32 * 1024 * 1024 : 2 * 1024 * 1024)
async function settle() {
  for (let i = 0; i < 10; i++) {
    await new Promise(resolve => setTimeout(resolve, 0))
    bun.gc(true); bun.gc(true)
  }
  return heapStats()
}
const warm = make()
for (let row = 0; row < Math.min(count, 512); row++) warm.get(geometry, row)
warm.close()
const before = await settle()
const cache = make()
for (let row = 0; row < count; row++) cache.get(geometry, row)
const after = await settle()
console.log(JSON.stringify({ count, width, growth: after.heapSize - before.heapSize, externalGrowth: after.extraMemorySize - before.extraMemorySize, charged: cache.bytes }))
cache.close()
