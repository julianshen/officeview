import type { SourceStore } from '../xlsx/view/contracts'
import { FindError } from '../xlsx/view/find'

/** Wire/workspace admission within the existing document pool; never a larger
 * worker budget. Large source pages and long hit texts refuse explicitly. */
export const WORKER_QUERY_BYTES = { D: 8 * 1024 * 1024, M: 4 * 1024 * 1024 } as const

/** Bounded, cycle-safe clone estimate with an early stop, including UTF-16 text
 * and complete backing buffers (not just a small typed-array view). */
export function queryPayloadBytes(value: unknown, limit: number): number {
  const seen = new Set<object>(), stack: unknown[] = [value]
  let bytes = 0
  while (stack.length && bytes <= limit) {
    const item = stack.pop()
    if (typeof item === 'string') { bytes += 16 + 2 * item.length; continue }
    if (typeof item !== 'object' || item === null) { bytes += 8; continue }
    if (seen.has(item)) continue
    seen.add(item)
    if (ArrayBuffer.isView(item)) { stack.push(item.buffer); bytes += 24; continue }
    if (Object.prototype.toString.call(item) === '[object ArrayBuffer]') { bytes += (item as ArrayBuffer).byteLength; continue }
    const values = Array.isArray(item) ? item : Object.values(item)
    bytes += 24 + values.length * 16
    if (bytes > limit) break
    for (const child of values) stack.push(child)
  }
  return bytes
}

export function boundedQueryStore(store: SourceStore, budget: number, emitted?: (bytes: number) => void, maxRecords = 20_000, fullDomain = false, loadedOnly = false): SourceStore {
  const checked = <T>(value: T, emission = false): T => {
    const bytes = queryPayloadBytes(value, budget)
    if (bytes > budget) throw new FindError('resource-limit', 'Source query page exceeds its bounded byte budget')
    if (emission) emitted?.(2 * bytes)
    return value
  }
  return {
    get stamp() { return store.stamp },
    get extent() { if (!fullDomain) return store.extent; const { used: _used, ...extent } = store.extent; return extent },
    get valuePolicy() { return store.valuePolicy },
    readCell: async (at, options) => checked(await store.readCell(at, options), true),
    readRows: async (range, cursor, options) => {
      const page = await store.readRows(range, cursor, options)
      if (page.items.reduce((count, row) => count + row.cells.length, 0) > maxRecords) throw new FindError('resource-limit', 'Viewport page exceeds its record budget')
      return checked(page)
    },
    queryColumn: async (request, options) => checked(await store.queryColumn(request, options)),
    queryStructural: async (rect, options) => {
      const result = await store.queryStructural(rect, options)
      if (result.merges.length + result.tables.length + result.spills.length + result.drawings.length > 4096) throw new FindError('resource-limit', 'Viewport structural hits exceed the 4096-record batch')
      return checked(result)
    },
    queryOverview: async (request, options) => checked(await store.queryOverview(request, options)),
    ensureComplete: (rect, options) => store.ensureComplete(rect, options),
    async *scanPopulated(request, options) {
      for await (const page of store.scanPopulated(loadedOnly ? { ...request, availability: 'loaded' } : request, options)) {
        if (page.items.length > 4096) throw new FindError('resource-limit', 'Source scan exceeds the 4096-cell page limit')
        yield checked({ ...page, items: page.items.map(cell => store.valuePolicy === 'engine' && cell.formula !== undefined && cell.value === null && cell.hasCachedValue === false && !cell.unavailable ? { ...cell, unavailable: true } : cell) })
      }
    },
    ...(store.scanOccupied ? { async *scanOccupied(request, options) {
      for await (const page of store.scanOccupied!(loadedOnly ? { ...request, availability: 'loaded' } : request, options)) {
        if (page.items.length > 4096) throw new FindError('resource-limit', 'Source occupancy scan exceeds its page limit')
        yield checked(page)
      }
    } } satisfies Pick<SourceStore, 'scanOccupied'> : {}),
    ...(store.settle ? { settle: (payload: object) => store.settle!(payload) } : {}),
    release: () => store.release(),
  }
}
