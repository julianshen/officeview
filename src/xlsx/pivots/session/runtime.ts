/**
 * What the lazy `pivots` chunk keeps per loaded workbook for interaction (plan Task 25, spec §7.4, caps ruling): the
 * parsed cache definitions, and a records source read on demand. Never structured-cloned and never put on the
 * `XlsxSheet` (which is): it lives in a WeakMap keyed by the document's sheets array and by each sheet.
 *
 * - `records()` decodes a cache's records part when asked, under `RECORD_LIMITS[profile]` (default D), and charges
 *   nothing: a sheet that never opens a session reads no record bytes.
 * - `lease()` is how a session holds a store (§7.4: pivots on one cache share one store). The first lease decodes; every
 *   later lease on the same cache shares that decode and store. The store is charged once to the main-thread ledger's
 *   `document` pool (`chargeWhileHeld`), each lease holds it, and the charge ends at the last release, when the
 *   runtime also drops the store (the next lease decodes again). The ledger has no pivot pool: adding one is in
 *   `src/xlsx/view/mainLedger.ts` (Task 26's lane), so the store counts as document memory.
 */
import { PackageError } from '../../../core/package-error'
import type { PivotCacheDefinition, PivotDiagnostic, PivotTableDefinition } from '../model'
import { decodeRecords, RECORD_LIMITS, type RecordLimits, type RecordResult, type RecordStore } from '../engine/records'
import type { MainMemoryLedger } from '../../view/mainLedger'

/** One holder's claim on a shared store. Releasing twice is a no-op. */
export interface RecordLease { readonly store: RecordStore; release(): void }
export type RecordLeaseResult = { ok: true; lease: RecordLease } | { ok: false; diagnostic: PivotDiagnostic }
export type RuntimeLedger = Pick<MainMemoryLedger, 'chargeWhileHeld' | 'hold' | 'chargeWhileReachable' | 'set'>
export interface PivotRuntime {
  /** The workbook's date system (decided after the pivots load; read when needed). */
  readonly date1904: boolean
  /** Bytes the runtime itself retains for its lifetime (the records parts' compressed copies, Ruling S3-E). */
  readonly retainedBytes: number
  /** Charges `retainedBytes` to `ledger`'s document pool while the runtime is reachable; once per ledger. */
  chargeRetained(ledger: RuntimeLedger): void
  cache(def: PivotTableDefinition): PivotCacheDefinition | undefined
  records(cache: PivotCacheDefinition, signal: AbortSignal, limits?: RecordLimits): Promise<RecordResult>
  lease(cache: PivotCacheDefinition, ledger: RuntimeLedger, signal: AbortSignal, limits?: RecordLimits): Promise<RecordLeaseResult>
}
/** A records part's bytes as a stream, or undefined when the part is not in the package. */
export type RecordsSource = (path: string, signal: AbortSignal) => AsyncIterable<Uint8Array> | PivotDiagnostic | undefined

const cancelled = () => new PackageError('cancelled', 'operation cancelled')
const absent = (cache: PivotCacheDefinition): RecordResult => ({ ok: false, diagnostic: { kind: 'unsupported', feature: 'pivot-records-absent', message: `Pivot cache ${cache.cacheId} has no records part`, partPath: cache.partPath } })

function waitOrAbort<T>(pending: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) return Promise.reject(cancelled())
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(cancelled())
    signal.addEventListener('abort', abort, { once: true })
    pending.then(value => { signal.removeEventListener('abort', abort); resolve(value) },
      error => { signal.removeEventListener('abort', abort); reject(error) })
  })
}

interface SharedStore { pending: Promise<RecordResult>; controller: AbortController; waiting: number; holders: number }

export function createPivotRuntime(caches: ReadonlyMap<number, PivotCacheDefinition>, source: RecordsSource, date1904: () => boolean, retainedBytes = 0): PivotRuntime {
  const charged = new WeakSet<object>()
  const byPath = new Map<string, PivotCacheDefinition>()
  for (const cache of caches.values()) if (!byPath.has(cache.partPath)) byPath.set(cache.partPath, cache)
  const shared = new Map<PivotCacheDefinition, SharedStore>()
  const drop = (cache: PivotCacheDefinition, entry: SharedStore) => {
    if (entry.waiting || entry.holders || shared.get(cache) !== entry) return
    shared.delete(cache)
    entry.controller.abort()
  }
  const runtime: PivotRuntime = {
    get date1904() { return date1904() },
    retainedBytes,
    chargeRetained(ledger) {
      if (charged.has(ledger)) return
      charged.add(ledger)
      if (retainedBytes > 0) ledger.chargeWhileReachable('document', [runtime], retainedBytes)
    },
    cache(def) {
      return (def.cachePartPath !== undefined ? byPath.get(def.cachePartPath) : undefined) ?? caches.get(def.cacheId)
    },
    async records(cache, signal, limits = RECORD_LIMITS.D) {
      if (signal.aborted) throw cancelled()
      const path = cache.saveData ? cache.recordsPartPath : undefined
      const chunks = path === undefined ? undefined : source(path, signal)
      if (!chunks) return absent(cache)
      if ('feature' in chunks) return { ok: false, diagnostic: chunks }
      const result = await decodeRecords(chunks, cache, limits, date1904(), signal)
      if (signal.aborted) throw cancelled()
      return result
    },
    async lease(cache, ledger, signal, limits) {
      if (signal.aborted) throw cancelled()
      runtime.chargeRetained(ledger)
      let entry = shared.get(cache)
      if (!entry) {
        const controller = new AbortController()
        // The first lease's limits decide the shared decode.
        entry = { pending: runtime.records(cache, controller.signal, limits), controller, waiting: 0, holders: 0 }
        // A decode nobody waits for any more (every waiter aborted) is abandoned; its rejection is not a failure.
        entry.pending.catch(() => {})
        shared.set(cache, entry)
      }
      const current = entry
      current.waiting++
      let result: RecordResult
      try { result = await waitOrAbort(current.pending, signal) }
      catch (error) { current.waiting--; drop(cache, current); throw error }
      current.waiting--
      if (!result.ok) {
        // A failed decode is not kept: a later lease tries again.
        if (shared.get(cache) === current && !current.holders) shared.delete(cache)
        return result
      }
      const store = result.store
      let hold = ledger.hold(store)
      if (!hold) { ledger.chargeWhileHeld('document', [store], store.bytes); hold = ledger.hold(store)! }
      current.holders++
      let held = true
      return { ok: true, lease: { store, release: () => {
        if (!held) return
        held = false
        hold!.release()
        current.holders--
        drop(cache, current)
      } } }
    },
  }
  return runtime
}

const runtimes = new WeakMap<object, PivotRuntime>()
/** Registers `runtime` under each owner (the document's sheets array, each of its sheets). */
export function bindPivotRuntime(owners: readonly object[], runtime: PivotRuntime): void {
  for (const owner of owners) runtimes.set(owner, runtime)
}
/** The runtime of a loaded resident workbook: by its document (through its `sheets` array) or by one of its sheets. */
export function pivotRuntimeFor(doc: object): PivotRuntime | undefined {
  const direct = runtimes.get(doc)
  if (direct) return direct
  const sheets = (doc as { sheets?: unknown }).sheets
  return typeof sheets === 'object' && sheets !== null ? runtimes.get(sheets) : undefined
}
