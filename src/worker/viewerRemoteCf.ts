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
/** Admitted bytes of one cache entry's own metadata (key, scope copy, stamp
 * copy, result holder): counted in retainedBytes from publication, so the
 * 64-entry bound is a byte bound too and pinned entries stay charged. */
export const CF_CACHE_ENTRY_OVERHEAD_BYTES = 512

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

/** Cache key: the sheet, the requested scope, and the full stamp. Every
 * generation rides the key — including view and clock — so no committed
 * change can ever serve a stale preparation. */
function entryKey(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): string {
  return `${sheetId}|${rectsKey(rects)}|${stamp.owner}|${stamp.workbookValueGen}|${stamp.cfGen}|${stamp.viewGen}|${stamp.structureGen}|${stamp.clockGen}`
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
  /** Detached but consumer-held entries: excluded from lookup, included in
   * retained accounting and the entry cap until their last release. */
  private readonly pinned = new Set<Entry>()
  private closed = false
  private notifyScheduled = false
  private readonly pendingNotifies = new Map<string, { rects: SourceRect[]; reason: 'updated' | 'evicted' | 'status' }>()
  constructor(private readonly host: CfReadHost, readonly budgetBytes: number = Number.MAX_SAFE_INTEGER) {}

  get(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): PagedCfState {
    if (!this.host.hasConditionalFormatting(sheetId)) return empty()
    if (this.closed) return empty()
    const entry = this.entryFor(sheetId, rects, stamp)
    if (!entry) return { state: 'failed', gen: 0, diagnostics: [{ kind: 'resource-limit', feature: 'cf-cache-budget', message: 'Conditional formatting cache is pinned at capacity.' }] }
    // Liveness gate: a changed owner invalidates the cached preparation even
    // before its key would miss, so no stale paint survives a commit.
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) {
      this.reset(entry)
      entry.gen++
      this.deferChanged(sheetId, entry.rects, 'updated')
    }
    if (entry.state === 'waiting' && !entry.controller) void this.load(entry)
    return entry
  }

  acquire(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp, signal?: AbortSignal): { state: PagedCfState; release(): void } {
    if (!this.host.hasConditionalFormatting(sheetId) || this.closed) {
      const state = empty()
      return { state, release() {} }
    }
    if (signal?.aborted) return { state: this.get(sheetId, rects, stamp), release() {} }
    const entry = this.entryFor(sheetId, rects, stamp)
    if (!entry) {
      const failed: PagedCfState = { state: 'failed', gen: 0, diagnostics: [{ kind: 'resource-limit', feature: 'cf-cache-budget', message: 'Conditional formatting cache is pinned at capacity.' }] }
      return { state: failed, release() {} }
    }
    // Liveness gate (mirrors get): a changed owner invalidates before use.
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) {
      this.reset(entry)
      entry.gen++
      this.deferChanged(sheetId, entry.rects, 'updated')
    }
    entry.refs++
    if (entry.state === 'waiting' && !entry.controller) void this.load(entry)
    let released = false
    let onAbort: (() => void) | undefined
    const release = () => {
      if (released) return
      released = true
      if (onAbort && signal) signal.removeEventListener('abort', onAbort)
      entry.refs--
      if (entry.refs <= 0 && entry.detached && !entry.dropped) {
        this.pinned.delete(entry)
        this.drop(entry)
      }
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
      else if (entry.refs > 0) this.pinned.add(entry)
    }
  }

  close(): void {
    if (this.closed) return
    this.closed = true
    for (const entry of this.entries.values()) {
      entry.detached = true
      if (entry.refs <= 0) {
        if (!entry.dropped) this.drop(entry)
      } else {
        // Held entries keep their charge until the last consumer releases,
        // even across close; no new reads start after this point. Abort the
        // in-flight load so no late reply can publish into the closed cache.
        entry.controller?.abort()
        entry.controller = undefined
        this.pinned.add(entry)
      }
    }
    this.entries.clear()
    this.pendingNotifies.clear()
  }

  get retainedEntries(): number { return [...this.entries.values()].filter(e => e.state === 'ready').length + [...this.pinned].filter(e => e.state === 'ready').length }
  get retainedBytes(): number { return [...this.entries.values()].reduce((n, e) => n + e.bytes, 0) + [...this.pinned].reduce((n, e) => n + e.bytes, 0) }

  private entryFor(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): Entry | undefined {
    const key = entryKey(sheetId, rects, stamp)
    let entry = this.entries.get(key)
    if (!entry) {
      // A new generation replaces same-scope entries: unreferenced ones drop
      // now (releasing once); referenced ones detach into the pinned set and
      // release on last use. This is what keeps an old pinned definition
      // charged while the new preparation builds beside it.
      for (const [otherKey, other] of [...this.entries]) {
        if (other.sheetId !== sheetId || !sameScope(other.rects, rects)) continue
        this.entries.delete(otherKey)
        other.detached = true
        other.gen++
        this.deferChanged(sheetId, other.rects, 'updated')
        if (other.refs <= 0) {
          if (!other.dropped) this.drop(other)
        } else {
          this.pinned.add(other)
        }
      }
      // Enforce the entry cap before creating: evict unreferenced entries
      // first; if everything is pinned, refuse the new entry rather than
      // growing without bound.
      while (this.entries.size + this.pinned.size >= MAX_ENTRIES) {
        const oldest = [...this.entries.values()].find(e => e.refs <= 0);
        if (!oldest) return undefined
        this.entries.delete(oldest.key)
        oldest.detached = true
        this.deferChanged(oldest.sheetId, oldest.rects, 'evicted')
        if (!oldest.dropped) this.drop(oldest)
      }
      this.entries.set(key, entry = {
        key, sheetId, rects: rects.map(r => ({ ...r })), stamp: { ...stamp },
        state: 'waiting', gen: 0, seq: 0, refs: 0, detached: false, dropped: false, bytes: 0, diagnostics: [],
      })
    } else {
      this.entries.delete(key)
      this.entries.set(key, entry)
    }
    return entry
  }

  private drop(entry: Entry): void {
    if (entry.dropped) return
    entry.dropped = true
    this.pinned.delete(entry)
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    entry.releaseReply?.()
    entry.releaseReply = undefined
    entry.reply = undefined
    entry.result = undefined
  }

  private reset(entry: Entry): void {
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    if (entry.releaseReply) {
      const release = entry.releaseReply
      entry.releaseReply = undefined
      entry.reply = undefined
      entry.result = undefined
      release()
    } else {
      entry.reply = undefined
      entry.result = undefined
    }
    entry.diagnostics = []
    entry.pendingSheetIds = undefined
    entry.owner = undefined
    entry.state = 'waiting'
    entry.bytes = 0
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
    // Grant ownership stays with the transport until a terminal response or
    // a proven not-sent failure: the transport settles the grant when the
    // reply arrives (even after our waiter detached), so the cache only
    // releases when the request provably never left (synchronous throw).
    let pending: Promise<WorkerReply<'read'> | symbol>
    try {
      pending = this.host.request({
        kind: 'read',
        requestId,
        owner,
        query: { op: 'conditionalFormatting', rects: entry.rects.map(r => ({ ...r })), stamp: { ...entry.stamp }, grantBytes },
      }, controller.signal)
    } catch {
      this.host.releaseGrant(requestId)
      if (entry.seq === seq && !this.closed) {
        entry.controller = undefined
        if (entry.state === 'loading') entry.state = 'waiting'
        entry.owner = undefined
      }
      return
    }
    try {
      const out = await pending
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
      // Full-stamp validation: the reply must match the requested identity on
      // every field, or it is dropped without publishing — never paint stale.
      // A mismatch means the caller is behind; it must re-request with live
      // generations (failing closed here breaks retry loops on stale stamps).
      const want = entry.stamp
      const got = result.stamp
      const stampMismatch = !got || got.owner !== want.owner || got.sheetId !== want.sheetId
        || got.workbookValueGen !== want.workbookValueGen || got.cfGen !== want.cfGen
        || got.viewGen !== want.viewGen || got.structureGen !== want.structureGen || got.clockGen !== want.clockGen
      if (stampMismatch || !sameOwner(out.owner as CacheIdentity, owner)) {
        releaseOnce()
        entry.state = 'failed'
        entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-stale-generation', message: 'Conditional formatting reply did not match the requested generations.' }]
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
      const bytes = (out.replyBytes ?? 0) + CF_CACHE_ENTRY_OVERHEAD_BYTES
      if (!Number.isSafeInteger(bytes) || !this.makeRoom(entry, bytes)) {
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
