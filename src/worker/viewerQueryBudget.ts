import type { AuthoritativeCell, Page, ReadOptions, SourceStore } from '../xlsx/view/contracts'
import { FindError } from '../xlsx/view/find'
import { CheckpointBudget } from '../xlsx/view/scheduler'
import { walkPayload, type WalkOutcome } from './replyMeasure'

/** Wire/workspace admission within the existing document pool; never a larger
 * worker budget. Large source pages and long hit texts refuse explicitly. */
export const WORKER_QUERY_BYTES = { D: 8 * 1024 * 1024, M: 4 * 1024 * 1024 } as const

export { queryPayloadBytes } from '../xlsx/view/queryPayload'

/** Viewport/search value-page facade. Drawings use the original interruptible
 * store and their separate drawing budget; they are not query-page payloads. */
export function boundedQueryStore(store: SourceStore, budget: number, emitted?: (bytes: number) => void, maxRecords = 20_000, fullDomain = false, loadedOnly = false): SourceStore {
  // queryPayloadBytes accounting through the worker's one payload walker (spec
  // §10.2): checkpointed (cancellation, owner) with real yields. Its workspace
  // is bounded by the byte limit at which the walk stops.
  // A page that needs no pause is accounted synchronously (no extra task turns).
  const checked = <T>(value: T, options: ReadOptions, emission = false): T | Promise<T> => {
    const pass = new CheckpointBudget(() => {
      if (options.signal?.aborted) throw new FindError('cancelled', 'Source query cancelled')
      const now = store.stamp
      if (options.stamp && (options.stamp.docId !== now.docId || options.stamp.sheetId !== now.sheetId || options.stamp.valueGen !== now.valueGen)) throw new FindError('stale', 'Source owner changed')
    }, undefined, {}, 'query')
    const done = (outcome: WalkOutcome): T => {
      pass.dispose()
      const bytes = outcome.kind === 'measured' ? outcome.bytes : Number.POSITIVE_INFINITY
      if (bytes > budget) throw new FindError('resource-limit', 'Source query page exceeds its bounded byte budget')
      if (emission) emitted?.(2 * bytes)
      return value
    }
    let run: WalkOutcome | Promise<WalkOutcome>
    try { run = walkPayload(value, () => true, pass, { model: 'query', limit: budget }) } catch (error) { pass.dispose(); throw error }
    return run instanceof Promise ? run.then(done, error => { pass.dispose(); throw error }) : done(run)
  }
  return {
    get stamp() { return store.stamp },
    get extent() { if (!fullDomain) return store.extent; const { used: _used, ...extent } = store.extent; return extent },
    get valuePolicy() { return store.valuePolicy },
    readCell: async (at, options) => checked(await store.readCell(at, options), options, true),
    readRows: async (range, cursor, options) => {
      const page = await store.readRows(range, cursor, options)
      if (page.items.reduce((count, row) => count + row.cells.length, 0) > maxRecords) throw new FindError('resource-limit', 'Viewport page exceeds its record budget')
      return checked(page, options)
    },
    queryColumn: async (request, options) => checked(await store.queryColumn(request, options), options),
    queryStructural: async (rect, options) => {
      const result = await store.queryStructural(rect, options)
      if (result.merges.length + result.tables.length + result.spills.length + result.drawings.length > 4096) throw new FindError('resource-limit', 'Viewport structural hits exceed the 4096-record batch')
      return checked(result, options)
    },
    queryOverview: async (request, options) => checked(await store.queryOverview(request, options), options),
    ensureComplete: (rect, options) => store.ensureComplete(rect, options),
    async *scanPopulated(request, options) {
      for await (const page of store.scanPopulated(loadedOnly ? { ...request, availability: 'loaded' } : request, options)) {
        if (page.items.length > 4096) throw new FindError('resource-limit', 'Source scan exceeds the 4096-cell page limit')
        const accounted = checked<Page<AuthoritativeCell>>({ ...page, items: page.items.map(cell => store.valuePolicy === 'engine' && cell.formula !== undefined && cell.value === null && cell.hasCachedValue === false && !cell.unavailable ? { ...cell, unavailable: true } : cell) }, options)
        yield accounted instanceof Promise ? await accounted : accounted
      }
    },
    ...(store.scanOccupied ? { async *scanOccupied(request, options) {
      for await (const page of store.scanOccupied!(loadedOnly ? { ...request, availability: 'loaded' } : request, options)) {
        if (page.items.length > 4096) throw new FindError('resource-limit', 'Source occupancy scan exceeds its page limit')
        const accounted = checked(page, options)
        yield accounted instanceof Promise ? await accounted : accounted
      }
    } } satisfies Pick<SourceStore, 'scanOccupied'> : {}),
    ...(store.settle ? { settle: (payload: object) => store.settle!(payload) } : {}),
    release: () => store.release(),
  }
}
