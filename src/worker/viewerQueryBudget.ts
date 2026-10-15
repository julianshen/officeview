import type { SourceStore } from '../xlsx/view/contracts'
import { FindError } from '../xlsx/view/find'
import { queryPayloadBytes } from '../xlsx/view/queryPayload'

/** Wire/workspace admission within the existing document pool; never a larger
 * worker budget. Large source pages and long hit texts refuse explicitly. */
export const WORKER_QUERY_BYTES = { D: 8 * 1024 * 1024, M: 4 * 1024 * 1024 } as const

export { queryPayloadBytes } from '../xlsx/view/queryPayload'

/** Viewport/search value-page facade. Drawings use the original interruptible
 * store and their separate drawing budget; they are not query-page payloads. */
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
