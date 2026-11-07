import type {
  CacheIdentity,
  CfReadDiagnostic,
  CfReadResult,
  CfReadStamp,
  WorkerReply,
} from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import type { CommandOf } from './viewerProtocol'

/** The receiver-owned CF budget is an upper bound, not a target: preparations
 * are admitted against their stamped size and evicted oldest-first. A
 * completed preparation stays cached until evicted, invalidated, replaced or
 * closed; an in-flight load belongs to its entry, so an aborted waiter merely
 * detaches while the shared build runs on for later getters. */
const MAX_ENTRIES = 64
const MAX_NOTIFY_RECTS = 16

/** Host surface the cache needs: routing identity and load state, a receiver
 * grant before any request leaves, and deferred invalidation once a
 * preparation completes, fails or is evicted. Notifications are always
 * scheduled outside render (queueMicrotask), never synchronously. */
export interface CfReadHost {
  identity(sheetId: string): CacheIdentity
  loadState(sheetId: string): { completeness: 'complete' | 'loading' | 'incomplete' }
  hasConditionalFormatting(sheetId: string): boolean
  /** Reserve receiver bytes synchronously, before a request can leave. */
  grant(requestId: string, minimumBytes?: number): number | null
  releaseGrant(requestId: string): void
  request(command: CommandOf<'read'>, signal?: AbortSignal): Promise<WorkerReply<'read'> | symbol>
  nextRequestId(): string
  /** End the ledger charge of a reply the cache no longer references. */
  release(reply: WorkerReply<'read'>): void
  changed(sheetId: string, rects: readonly SourceRect[], reason?: 'updated' | 'evicted' | 'status'): void
}

export interface PagedCfState {
  readonly state: 'waiting' | 'loading' | 'ready' | 'failed'
  readonly gen: number
  readonly result?: CfReadResult
  readonly diagnostics: readonly CfReadDiagnostic[]
  readonly pendingSheetIds?: readonly string[]
}

interface Entry {
  key: string
  sheetId: string
  rects: SourceRect[]
  stamp: CfReadStamp
  state: PagedCfState['state']
  gen: number
  owner?: CacheIdentity
  seq: number
  refs: number
  detached: boolean
  dropped: boolean
  bytes: number
  reply?: WorkerReply<'read'>
  releaseReply?: () => void
  result?: CfReadResult
  diagnostics: CfReadDiagnostic[]
  pendingSheetIds?: readonly string[]
  controller?: AbortController
}

const empty = (): PagedCfState => ({ state: 'waiting', gen: 0, diagnostics: [] })
const sameOwner = (a: CacheIdentity, b: CacheIdentity) => a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen

function rectsKey(rects: readonly SourceRect[]): string {
  return rects.map(r => `${r.r0},${r.r1},${r.c0},${r.c1}`).join(';')
}

function overlaps(a: readonly SourceRect[], b: readonly SourceRect[]): boolean {
  return a.some(x => b.some(y => x.r0 <= y.r1 && y.r0 <= x.r1 && x.c0 <= y.c1 && y.c0 <= x.c1))
}

/** Cache key: the sheet, the requested scope, and every generation that can
 * change the result. View and clock generations are intentionally excluded:
 * paint-only commits and clock ticks never change composed states (Task 10
 * owns time-based invalidation). */
function entryKey(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): string {
  return `${sheetId}|${rectsKey(rects)}|${stamp.owner}|${stamp.workbookValueGen}|${stamp.cfGen}|${stamp.structureGen}`
}

function unionRects(rects: readonly SourceRect[]): SourceRect[] {
  if (rects.length <= MAX_NOTIFY_RECTS) return [...rects]
  let r0 = Number.MAX_SAFE_INTEGER, r1 = 0, c0 = Number.MAX_SAFE_INTEGER, c1 = 0
  for (const r of rects) { r0 = Math.min(r0, r.r0); r1 = Math.max(r1, r.r1); c0 = Math.min(c0, r.c0); c1 = Math.max(c1, r.c1) }
  return [{ r0, r1, c0, c1 }]
}

/** Bounded receiver cache for worker CF preparations: one shared in-flight
 * build per key, refcounted entries pinned across definition replacement,
 * LRU eviction under the receiver budget, and deferred affected-tile
 * notification. */
export class RemoteCfCache {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  private notifyScheduled = false
  private readonly pendingNotifies = new Map<string, { rects: SourceRect[]; reason: 'updated' | 'evicted' | 'status' }>()
  constructor(private readonly host: CfReadHost, readonly budgetBytes: number = Number.MAX_SAFE_INTEGER) {}

  get(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): PagedCfState {
    if (!this.host.hasConditionalFormatting(sheetId)) return empty()
    const entry = this.entryFor(sheetId, rects, stamp)
    if (entry.state === 'waiting' && !entry.controller) void this.load(entry)
    return entry
  }

  acquire(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp, signal?: AbortSignal): { state: PagedCfState; release(): void } {
    if (!this.host.hasConditionalFormatting(sheetId)) {
      const state = empty()
      return { state, release() {} }
    }
    if (signal?.aborted) return { state: this.get(sheetId, rects, stamp), release() {} }
    const entry = this.entryFor(sheetId, rects, stamp)
    entry.refs++
    if (entry.state === 'waiting' && !entry.controller) void this.load(entry)
    let released = false
    let onAbort: (() => void) | undefined
    const release = () => {
      if (released) return
      released = true
      if (onAbort && signal) signal.removeEventListener('abort', onAbort)
      entry.refs--
      if (entry.refs <= 0 && entry.detached && !entry.dropped) this.drop(entry)
    };
    if (signal) {
      onAbort = () => release()
      signal.addEventListener('abort', onAbort, { once: true })
    }
    return { state: entry, release }
  }

  /** Forget cached preparations (scoped to rects when given); held versions
   * survive until their last release, so pinned consumers never see torn data. */
  invalidate(sheetId: string, rects?: readonly SourceRect[]): void {
    for (const [key, entry] of [...this.entries]) {
      if (entry.sheetId !== sheetId) continue
      if (rects && rects.length && !overlaps(entry.rects, rects)) continue
      this.entries.delete(key)
      entry.detached = true
      entry.gen++
      this.deferChanged(sheetId, entry.rects, 'updated')
      if (entry.refs <= 0 && !entry.dropped) this.drop(entry)
    }
  }

  close(): void {
    if (this.closed) return
    this.closed = true
    for (const entry of this.entries.values()) {
      entry.detached = true
      if (!entry.dropped) this.drop(entry)
    }
    this.entries.clear()
    this.pendingNotifies.clear()
  }

  get retainedEntries(): number { return [...this.entries.values()].filter(e => e.state === 'ready').length }
  get retainedBytes(): number { return [...this.entries.values()].reduce((n, e) => n + e.bytes, 0) }

  private entryFor(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): Entry {
    const key = entryKey(sheetId, rects, stamp)
    let entry = this.entries.get(key)
    if (!entry) {
      // A new generation replaces same-scope entries: unreferenced ones drop
      // now (releasing once); referenced ones detach and release on last use.
      // This is what keeps an old pinned definition charged while the new
      // preparation builds beside it.
      for (const [otherKey, other] of [...this.entries]) {
        if (other.sheetId !== sheetId || !sameScope(other.rects, rects)) continue
        this.entries.delete(otherKey)
        other.detached = true
        other.gen++
        this.deferChanged(sheetId, other.rects, 'updated')
        if (other.refs <= 0 && !other.dropped) this.drop(other)
      }
      this.entries.set(key, entry = {
        key, sheetId, rects: rects.map(r => ({ ...r })), stamp: { ...stamp },
        state: 'waiting', gen: 0, seq: 0, refs: 0, detached: false, dropped: false, bytes: 0, diagnostics: [],
      })
      while (this.entries.size > MAX_ENTRIES) {
        const oldest = [...this.entries.values()].find(e => e.refs <= 0 && e !== entry)
        if (!oldest) break
        this.entries.delete(oldest.key)
        oldest.detached = true
        this.deferChanged(oldest.sheetId, oldest.rects, 'evicted')
        if (!oldest.dropped) this.drop(oldest)
      }
    } else {
      this.entries.delete(key)
      this.entries.set(key, entry)
    }
    return entry
  }

  private drop(entry: Entry): void {
    if (entry.dropped) return
    entry.dropped = true
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    entry.releaseReply?.()
    entry.releaseReply = undefined
    entry.reply = undefined
    entry.result = undefined
  }

  private deferChanged(sheetId: string, rects: readonly SourceRect[], reason: 'updated' | 'evicted' | 'status'): void {
    const pending = this.pendingNotifies.get(sheetId)
    if (pending) {
      pending.rects.push(...rects)
      if (reason === 'updated') pending.reason = 'updated'
      else if (reason === 'evicted' && pending.reason === 'status') pending.reason = 'evicted'
    } else {
      this.pendingNotifies.set(sheetId, { rects: [...rects], reason })
    }
    if (this.notifyScheduled) return
    this.notifyScheduled = true
    queueMicrotask(() => {
      this.notifyScheduled = false
      if (this.closed) return
      for (const [id, note] of this.pendingNotifies) {
        this.pendingNotifies.delete(id)
        try {
          this.host.changed(id, unionRects(note.rects), note.reason)
        } catch { /* observer failures must not change cache ownership */ }
      }
    })
  }

  private unpack(reply: WorkerReply<'read'>): CfReadResult | undefined {
    if (reply.kind !== 'complete' || reply.result.kind !== 'read' || reply.result.result.op !== 'conditionalFormatting') return undefined
    return reply.result.result.result
  }

  private async load(entry: Entry): Promise<void> {
    const { sheetId } = entry
    if (this.closed) return
    if (this.host.loadState(sheetId).completeness !== 'complete') return
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) {
      entry.seq++
      entry.controller?.abort()
      entry.controller = undefined
      entry.releaseReply?.()
      entry.releaseReply = undefined
      entry.reply = undefined
      entry.result = undefined
      entry.diagnostics = []
      entry.pendingSheetIds = undefined
      entry.owner = undefined
      entry.state = 'waiting'
      entry.bytes = 0
      entry.gen++
    }
    const requestId = this.host.nextRequestId()
    const grantBytes = this.host.grant(requestId)
    if (grantBytes === null) return
    const seq = ++entry.seq
    const controller = entry.controller = new AbortController()
    entry.owner = owner
    entry.state = 'loading'
    let grantOpen = true
    const endGrant = () => {
      if (!grantOpen) return
      grantOpen = false
      this.host.releaseGrant(requestId)
    }
    try {
      const out = await this.host.request({
        kind: 'read',
        requestId,
        owner,
        query: { op: 'conditionalFormatting', rects: entry.rects.map(r => ({ ...r })), stamp: { ...entry.stamp }, grantBytes },
      }, controller.signal)
      if (this.closed || entry.seq !== seq) {
        if (typeof out === 'object') this.host.release(out)
        return
      }
      entry.controller = undefined
      if (typeof out !== 'object') {
        // Cancelled or closed while in flight: back to waiting, never failed
        // or stale — a later get retries with the current stamp.
        if (entry.state === 'loading') entry.state = 'waiting'
        entry.owner = undefined
        return
      }
      const releaseOnce = (() => { let done = false; return () => { if (done) return; done = true; this.host.release(out) } })()
      if (out.kind === 'cancelled' || out.kind === 'incomplete') {
        releaseOnce()
        if (entry.state === 'loading') entry.state = 'waiting'
        entry.owner = undefined
        return
      }
      if (out.kind === 'error') {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-read-failed', message: out.diagnostics[0]?.message ?? `CF read failed (${out.code})` }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      if (out.kind === 'quota') {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'resource-limit', feature: 'cf-grant-required', message: `CF reply exceeds the receiver grant.${out.requiredBytes !== undefined ? ` Retry with at least ${out.requiredBytes} bytes.` : ''}` }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      const result = this.unpack(out)
      if (!result || result.sheetId !== sheetId) {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-read-failed', message: 'Conditional formatting read did not complete.' }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      if (!Number.isSafeInteger(out.replyBytes) || (out.replyBytes ?? -1) < 0) {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-read-failed', message: 'Conditional formatting reply had no valid byte estimate.' }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      if (result.kind === 'pending') {
        releaseOnce()
        entry.state = 'waiting'
        entry.pendingSheetIds = result.pendingSheetIds
        entry.owner = undefined
        return
      }
      if (result.kind !== 'ready') {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = result.diagnostics.length
          ? [...result.diagnostics]
          : [{ kind: 'unsupported', feature: 'cf-preparation-withheld', message: 'Conditional formatting withheld the preparation.' }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      const bytes = out.replyBytes ?? 0
      if (!this.makeRoom(entry, bytes)) {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'resource-limit', feature: 'cf-cache-budget', message: 'Conditional formatting data exceeds the main-thread cache budget.' }]
        entry.gen++
        this.deferChanged(sheetId, entry.rects, 'status')
        return
      }
      entry.reply = out
      entry.releaseReply = releaseOnce
      entry.result = result
      entry.bytes = bytes
      entry.diagnostics = [...result.diagnostics]
      entry.pendingSheetIds = undefined
      entry.state = 'ready'
      entry.gen++
      this.deferChanged(sheetId, entry.rects, 'updated')
    } catch {
      if (entry.seq === seq && !this.closed) {
        entry.controller = undefined
        if (entry.state === 'loading') entry.state = 'waiting'
        entry.owner = undefined
      }
    } finally {
      endGrant()
    }
  }

  private makeRoom(current: Entry, needed: number): boolean {
    for (const [id, other] of this.entries) {
      if (this.retainedBytes + needed <= this.budgetBytes) break
      if (other === current || other.state !== 'ready' || other.refs > 0) continue
      this.entries.delete(id)
      other.detached = true
      other.gen++
      this.deferChanged(other.sheetId, other.rects, 'evicted')
      if (!other.dropped) this.drop(other)
    }
    return this.retainedBytes + needed <= this.budgetBytes
  }
}

function sameScope(a: readonly SourceRect[], b: readonly SourceRect[]): boolean {
  return rectsKey(a) === rectsKey(b)
}
