import { PackageError } from '../../core/package-error'
import type { CfMemory, CfReservation } from './limits'
import type { CfLease, CfPrepared, CfStamp } from './model'

export function sameCfStamp(a: CfStamp, b: CfStamp): boolean {
  return a.owner === b.owner && a.sheetId === b.sheetId && a.workbookValueGen === b.workbookValueGen && a.cfGen === b.cfGen
    && a.viewGen === b.viewGen && a.structureGen === b.structureGen && a.clockGen === b.clockGen
}
const cancelled = () => new PackageError('cancelled', 'CF preparation is no longer current')
const quota = () => new PackageError('resource-limit', 'CF preparation admission refused')
interface Entry { key: string; stamp: CfStamp; lease: CfLease<CfPrepared>; metadata: CfReservation; refs: number }
interface Waiter { signal: AbortSignal; abort(): void; hold: CfReservation; resolve(v: CfLease<CfPrepared>): void; reject(e: unknown): void }
interface Job { stamp: CfStamp; controller: AbortController; waiters: Set<Waiter>; metadata: CfReservation }

/** One sheet/session owner. Cache references, consumer references, and live producers
 * are independent owners. Aborting a producer never releases its allocations. */
export function createCfCache(memory: CfMemory, maxEntries = 16) {
  const ready = new Map<string, Entry>(), jobs = new Map<string, Job>()
  let current: CfStamp | undefined, closed = false
  const isCurrent = (stamp: CfStamp) => !closed && current !== undefined && sameCfStamp(current, stamp)
  function drop(entry: Entry) {
    if (--entry.refs !== 0) return
    try { entry.lease.value.dispose() } finally { entry.lease.release(); entry.metadata.release() }
  }
  function detach(job: Job, w: Waiter) { job.waiters.delete(w); w.signal.removeEventListener('abort', w.abort) }
  function stop(job: Job) {
    job.controller.abort()
    for (const w of job.waiters) { w.signal.removeEventListener('abort', w.abort); w.hold.release(); w.reject(cancelled()) }
    job.waiters.clear()
  }
  function activate(stamp: CfStamp) {
    if (closed) throw cancelled()
    if (current && sameCfStamp(current, stamp)) return
    current = Object.freeze({ ...stamp })
    for (const entry of ready.values()) drop(entry)
    ready.clear()
    for (const job of jobs.values()) stop(job)
    jobs.clear() // jobs retain their own metadata until their producers settle
  }
  function borrow(entry: Entry, hold: CfReservation): CfLease<CfPrepared> {
    entry.refs++
    let released = false
    const release = () => { if (!released) { released = true; hold.release(); drop(entry) } }
    const value: CfPrepared = Object.freeze({ stamp: entry.stamp,
      at: (row: number, col: number) => released || !isCurrent(entry.stamp)
        ? { kind: 'unresolved' as const, reason: 'cf-stale-generation' } : entry.lease.value.at(row, col), dispose: release })
    return { value, bytes: entry.lease.bytes, release,
      transfer: () => ({ kind: 'unresolved', reason: 'cf-preparation-transfer-requires-snapshot' }) }
  }
  function prepare(key: string, stamp: CfStamp, signal: AbortSignal,
    producer: (signal: AbortSignal) => Promise<CfLease<CfPrepared>>): Promise<CfLease<CfPrepared>> {
    if (closed || signal.aborted) return Promise.reject(cancelled())
    activate(stamp)
    if (key.length > 65536) return Promise.reject(quota())
    const waiterHold = memory.reserve(256)
    if (waiterHold.kind !== 'reserved') return Promise.reject(quota())
    const cached = ready.get(key)
    if (cached) { ready.delete(key); ready.set(key, cached); return Promise.resolve(borrow(cached, waiterHold)) }
    let job = jobs.get(key), start = false
    if (job && job.controller.signal.aborted) { jobs.delete(key); job = undefined }
    if (!job) {
      const metadata = memory.reserve(1024 + key.length * 2 + stamp.owner.length * 2 + stamp.sheetId.length * 2)
      if (metadata.kind !== 'reserved') { waiterHold.release(); return Promise.reject(quota()) }
      job = { stamp: Object.freeze({ ...stamp }), controller: new AbortController(), waiters: new Set(), metadata }
      jobs.set(key, job); start = true
    }
    const active = job
    const promise = new Promise<CfLease<CfPrepared>>((resolve, reject) => {
      const waiter: Waiter = { signal, hold: waiterHold, resolve, reject, abort() {
        detach(active, waiter); waiterHold.release(); reject(cancelled())
        if (active.waiters.size === 0) active.controller.abort()
      } }
      active.waiters.add(waiter); signal.addEventListener('abort', waiter.abort, { once: true })
    })
    if (start) {
      const failure = (error: unknown) => {
        if (jobs.get(key) === active) jobs.delete(key)
        active.metadata.release()
        for (const w of active.waiters) { w.signal.removeEventListener('abort', w.abort); w.hold.release(); w.reject(error) }
        active.waiters.clear()
      }
      let produced: Promise<CfLease<CfPrepared>>
      try { produced = producer(active.controller.signal) } catch (error) { failure(error); return promise }
      void produced.then(lease => {
        if (jobs.get(key) === active) jobs.delete(key)
        if (active.controller.signal.aborted || !isCurrent(active.stamp) || !sameCfStamp(lease.value.stamp, active.stamp) || !active.waiters.size) {
          try { lease.value.dispose() } finally { lease.release(); failure(cancelled()) }
          return
        }
        const entry: Entry = { key, stamp: active.stamp, lease, metadata: active.metadata, refs: 1 }
        ready.set(key, entry)
        for (const w of active.waiters) { w.signal.removeEventListener('abort', w.abort); w.resolve(borrow(entry, w.hold)) }
        active.waiters.clear()
        while (ready.size > maxEntries) { const oldest = ready.entries().next().value!; ready.delete(oldest[0]); drop(oldest[1]) }
      }, failure)
    }
    return promise
  }
  function at(stamp: CfStamp, row: number, col: number) {
    if (isCurrent(stamp)) for (const entry of ready.values()) {
      const result = entry.lease.value.at(row, col)
      if (result.kind !== 'unresolved' || result.reason !== 'cf-outside-preparation') return result
    }
    return { kind: 'unresolved' as const, reason: 'cf-pending-data' }
  }
  function dispose() {
    if (closed) return
    closed = true
    for (const entry of ready.values()) drop(entry)
    ready.clear()
    for (const job of jobs.values()) stop(job)
    jobs.clear()
  }
  return { prepare, activate, at, dispose }
}
