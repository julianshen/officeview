import type { AuthoritativeCell, Page, ReadOptions, SourceStore } from '../xlsx/view/contracts'
import { FindError } from '../xlsx/view/find'
import { CheckpointBudget } from '../xlsx/view/scheduler'
import { StopWalk, walkAdmission, walkPayload, type WalkOutcome } from './replyMeasure'
import { rememberQueryBytes } from '../xlsx/view/queryBytesMemo'
import type { PoolCharge } from '../xlsx/view/tileCache'
import type { MainMemoryLedger, MainPool } from '../xlsx/view/mainLedger'

/** Wire/workspace admission within the existing document pool; never a larger
 * worker budget. Large source pages and long hit texts refuse explicitly. */
export const WORKER_QUERY_BYTES = { D: 8 * 1024 * 1024, M: 4 * 1024 * 1024 } as const

export { queryPayloadBytes } from '../xlsx/view/queryPayload'

/** A pool whose transient shares admit a walk's workspace (the document pool,
 * or the main ledger through ledgerPool). */
export interface QueryPool { charge(): Pick<PoolCharge, 'set' | 'release'> }
let ledgerSerial = 0
/** Main-thread walks admit their workspace to the main ledger: a transient
 * owner in `pool`, refused once the ledger would exceed `limitBytes`. */
export function ledgerPool(ledger: MainMemoryLedger, pool: MainPool, limitBytes: number): QueryPool {
  return { charge: () => {
    const owner = `walk:${++ledgerSerial}`
    let held = 0
    return {
      set(bytes: number) {
        if (ledger.bytes - held + bytes > limitBytes) return false
        held = bytes
        ledger.set(pool, owner, bytes)
        return true
      },
      release() { ledger.release(pool, owner); held = 0 },
    }
  } }
}
/** queryPayloadBytes of `value` through the worker's one payload walker (spec
 * §10.2): checkpointed by `check` (it throws on cancellation or a stale owner)
 * with real yields (the host's, or the session's injected `hostYield`), its workspace admitted to `pool` in chunks and released
 * after the walk. Synchronous when no pause is due. `undefined`: the pool
 * refused the workspace; a StopWalk from `check` is rethrown. With `shared`,
 * the walk charges that budget (its own check applies). */
export function accountQueryBytes(value: unknown, limit: number, check: () => void, pool?: QueryPool, hostYield?: () => Promise<void>, shared?: CheckpointBudget): number | undefined | Promise<number | undefined> {
  const workspace = pool?.charge()
  // A shared budget carries one operation's work across several passes.
  const pass = shared ?? new CheckpointBudget(check, undefined, hostYield ? { yield: hostYield } : {}, 'query')
  const end = () => { if (!shared) pass.dispose(); workspace?.release() }
  const done = (outcome: WalkOutcome): number | undefined => {
    end()
    // A stop (the check threw StopWalk) is not a workspace refusal.
    if (outcome.kind === 'stopped') throw new StopWalk()
    return outcome.kind === 'measured' ? outcome.bytes : undefined
  }
  let run: WalkOutcome | Promise<WalkOutcome>
  try { run = walkPayload(value, walkAdmission(workspace), pass, { model: 'query', limit }) } catch (error) { end(); throw error }
  return run instanceof Promise ? run.then(done, error => { end(); throw error }) : done(run)
}
/** A typed retryable refusal: the walk's workspace does not fit the pool now. */
export const queryWorkspaceRefusal = () => new FindError('resource-limit', 'Source query accounting exceeds the worker memory budget; retry', true)

/** Viewport/search value-page facade. Drawings use the original interruptible
 * store and their separate drawing budget; they are not query-page payloads. */
export function boundedQueryStore(store: SourceStore, budget: number, emitted?: (bytes: number) => void, maxRecords = 20_000, fullDomain = false, loadedOnly = false, pool?: QueryPool, hostYield?: () => Promise<void>): SourceStore {
  // Accounting through the one walker, checkpointed (cancellation, owner) with
  // real yields and its workspace admitted to the document pool; a page that
  // needs no pause is accounted synchronously (no extra task turns). A refused
  // workspace publishes no page: a typed retryable resource-limit.
  const checked = <T>(value: T, options: ReadOptions, emission = false): T | Promise<T> => {
    const done = (bytes: number | undefined): T => {
      if (bytes === undefined) throw queryWorkspaceRefusal()
      if (bytes > budget) throw new FindError('resource-limit', 'Source query page exceeds its bounded byte budget')
      if (emission) emitted?.(2 * bytes)
      rememberQueryBytes(value, budget, bytes)
      return value
    }
    const accounted = accountQueryBytes(value, budget, () => {
      if (options.signal?.aborted) throw new FindError('cancelled', 'Source query cancelled')
      const now = store.stamp
      if (options.stamp && (options.stamp.docId !== now.docId || options.stamp.sheetId !== now.sheetId || options.stamp.valueGen !== now.valueGen)) throw new FindError('stale', 'Source owner changed')
    }, pool, hostYield)
    return accounted instanceof Promise ? accounted.then(done) : done(accounted)
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
