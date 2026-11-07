import type {
  CacheIdentity,
  CfReadDiagnostic,
  CfReadResult,
  CfReadStamp,
  CfSharedDefinition,
  WorkerReply,
} from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import { MEASURE_ENTRY_BYTES, MEASURE_FIRST_CHUNK_ENTRIES, measurePayloadBytes } from '../xlsx/view/payloadBytes'
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
  /** Admit one cache entry's metadata (key string, scope/rect copies, stamp
   * copy, entry shell) on the main ledger before the cache allocates it;
   * null refuses and the entry is not created. The returned release ends
   * that charge with the entry. */
  admitEntry(key: string, bytes: number): (() => void) | null
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
  /** Admitted metadata bytes (key, scope, stamp, shell); the ledger release
   * sits beside them. */
  meta: number
  releaseEntry?: () => void
  /** Admitted per-entry states charge (framing, states and diagnostics, but
   * not shared definitions): released with the entry. Absent on the legacy
   * path, where the full reply charge is kept instead. */
  releaseStates?: () => void
  /** Adopted version whose shared definitions this entry reads; the registry
   * holds the array while any entry references it. */
  sharedVersion?: number
  /** A version adoption already consumed one bounded retry on this entry. */
  adopted: boolean
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

/** Admitted bytes of one cache entry's metadata, computed from input sizes
 * before the cache allocates anything: the key string, the scope/rect
 * copies, the stamp copy and the entry shell. Waiting and failed entries
 * carry exactly this; ready entries add their reply bytes. */
export function entryMetaBytes(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): number {
  // Shell plus the key string, one copy per scope rect (four numbers plus
  // object slots) and the stamp copy (identities plus generation numbers).
  return CF_CACHE_ENTRY_OVERHEAD_BYTES + (sheetId.length + stamp.owner.length) * 2 + rects.length * 128 + 256
}

function unionRects(rects: readonly SourceRect[]): SourceRect[] {
  if (rects.length <= MAX_NOTIFY_RECTS) return [...rects]
  let r0 = Number.MAX_SAFE_INTEGER, r1 = 0, c0 = Number.MAX_SAFE_INTEGER, c1 = 0
  for (const r of rects) { r0 = Math.min(r0, r.r0); r1 = Math.max(r1, r.r1); c0 = Math.min(c0, r.c0); c1 = Math.max(c1, r.c1) }
  return [{ r0, r1, c0, c1 }]
}

/** Scope equality without building key strings: same length, same
 * coordinates in order. */
function sameScopeRects(a: readonly SourceRect[], b: readonly SourceRect[]): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) {
    const x = a[i]!, y = b[i]!
    if (x.r0 !== y.r0 || x.r1 !== y.r1 || x.c0 !== y.c0 || x.c1 !== y.c1) return false
  }
  return true
}

/** Full stamp equality without building key strings. */
function sameStamp(a: CfReadStamp, b: CfReadStamp): boolean {
  return a.owner === b.owner && a.sheetId === b.sheetId
    && a.workbookValueGen === b.workbookValueGen && a.cfGen === b.cfGen
    && a.viewGen === b.viewGen && a.structureGen === b.structureGen && a.clockGen === b.clockGen
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
  /** Version-keyed shared definitions: one frozen, ledger-charged array per
   * definitions version, refcounted across viewport consumers with
   * independent state leases. The registry owns its bytes through a ledger
   * admission per version (bounded by live versions); per-entry reply
   * charges stay conservative on top, so no accounting path under-counts.
   * Dropping the last referencing entry releases the version. */
  private readonly sharedDefs = new Map<number, { definitions: readonly CfSharedDefinition[]; refs: number; bytes: number; release: () => void }>()
  private closed = false
  private notifyScheduled = false
  private readonly pendingNotifies = new Map<string, { rects: SourceRect[]; reason: 'updated' | 'evicted' | 'status' }>()
  /** Monotonic entry admission owner: rekeys keep the original owner. */
  private entrySeq = 0
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
  get retainedBytes(): number {
    let total = [...this.entries.values()].reduce((n, e) => n + e.bytes, 0) + [...this.pinned].reduce((n, e) => n + e.bytes, 0)
    for (const shared of this.sharedDefs.values()) total += shared.bytes
    return total
  }
  /** Test seam: live version-keyed shared definition entries. */
  get sharedDefinitionVersions(): number { return this.sharedDefs.size }

  private entryFor(sheetId: string, rects: readonly SourceRect[], stamp: CfReadStamp): Entry | undefined {
    // Warm hits never touch the ledger: compare the bounded scope and stamp
    // fields directly (no key strings, no copies, no admission) before
    // allocating anything for a miss.
    for (const [key, existing] of this.entries) {
      if (existing.sheetId === sheetId && sameScopeRects(existing.rects, rects) && sameStamp(existing.stamp, stamp)) {
        this.entries.delete(key)
        this.entries.set(key, existing)
        return existing
      }
    }
    // Admit variable metadata on the main ledger before allocating the key
    // string, the scope/rect copies or the entry shell. Waiting and failed
    // entries keep exactly this charge until the entry drops. The admission
    // owner is entry-scoped (not key-scoped), so rekeys keep it.
    const meta = entryMetaBytes(sheetId, rects, stamp)
    const releaseEntry = this.host.admitEntry(`cf-entry:${this.entrySeq++}`, meta)
    if (!releaseEntry) return undefined
    const key = entryKey(sheetId, rects, stamp)
    const existing = this.entries.get(key)
    if (existing) {
      // A lookup hit allocates nothing new: end the spare admission at once.
      try { releaseEntry() } catch { /* release-only path */ }
      this.entries.delete(key)
      this.entries.set(key, existing)
      return existing
    }
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
      if (!oldest) {
        try { releaseEntry() } catch { /* release-only path */ }
        return undefined
      }
      this.entries.delete(oldest.key)
      oldest.detached = true
      this.deferChanged(oldest.sheetId, oldest.rects, 'evicted')
      if (!oldest.dropped) this.drop(oldest)
    }
    const entry: Entry = {
      key, sheetId, rects: rects.map(r => ({ ...r })), stamp: { ...stamp },
      state: 'waiting', gen: 0, seq: 0, refs: 0, detached: false, dropped: false,
      meta, releaseEntry, adopted: false, bytes: meta, diagnostics: [],
    }
    this.entries.set(key, entry)
    return entry
  }

  /** Moves a live entry to a newly adopted stamp (cold-start or first
   * discovery): the old key is forgotten and the new key installed. At most
   * one entry exists per key, so a clash overwrites — unreachable in current
   * flows, where adoption only runs on live loading entries. */
  private rekey(entry: Entry, stamp: CfReadStamp): void {
    this.entries.delete(entry.key)
    entry.stamp = { ...stamp }
    entry.key = entryKey(entry.sheetId, entry.rects, entry.stamp)
    this.entries.set(entry.key, entry)
  }

  /** Deep-freezes a cloned definition array: the structured-clone boundary
   * thaws what the worker froze, so the registry restores immutability at
   * every level before sharing. */
  private static freezeShared(definitions: readonly CfSharedDefinition[]): readonly CfSharedDefinition[] {
    for (const definition of definitions) {
      if (definition.properties && typeof definition.properties === 'object') Object.freeze(definition.properties)
      if (definition.writes && typeof (definition.writes as unknown[]).length === 'number') Object.freeze(definition.writes)
      Object.freeze(definition)
    }
    return Object.freeze([...definitions])
  }

  /** Admitted walk of one payload subgraph: each traversal chunk is admitted
   * on the main ledger before it grows and released with the walk, so the
   * walk never runs on unadmitted workspace. Returns undefined when the walk
   * cannot be admitted or exceeds its cap. */
  private walkSeq = 0
  private measuredWalk(value: unknown, cap: number, label: string): number | undefined {
    let admitted = 0
    const releases: Array<() => void> = []
    try {
      return measurePayloadBytes(value, extra => {
        const next = admitted + extra
        if (next > cap) return false
        const hold = this.host.admitEntry(`${label}:${this.walkSeq++}:${admitted}`, extra)
        if (!hold) return false
        releases.push(hold)
        admitted = next
        return true
      }) ?? undefined
    } finally {
      for (const release of releases.splice(0)) {
        try { release() } catch { /* release-only path */ }
      }
    }
  }

  /** Component sizes of a ready payload for single-charge publication: the
   * states (plus framing and diagnostics) charged per entry, and the
   * definitions charged once per version in the registry. Walks are capped
   * at the reply they arrived in (floored at the walker's first chunk so
   * small shared payloads still share). */
  private measuredSplit(result: CfReadResult, replyBytes: number): { statesBytes: number; defsBytes: number } | undefined {
    const cap = Math.max(replyBytes, MEASURE_FIRST_CHUNK_ENTRIES * MEASURE_ENTRY_BYTES)
    const statesBytes = this.measuredWalk({ ...result, definitions: [] }, cap, 'cf-walk-states')
    if (statesBytes === undefined) return undefined
    const defsBytes = this.measuredWalk(result.definitions, cap, 'cf-defs-walk')
    if (defsBytes === undefined) return undefined
    return { statesBytes, defsBytes }
  }

  /** Shares the reply's definitions by version: returns the shared array
   * (adopting it on first use with a ledger admission), or undefined when
   * sharing is unavailable. Callers unshare on any later failure. */
  private shareDefinitions(version: number, definitions: readonly CfSharedDefinition[], defsBytes: number): readonly CfSharedDefinition[] | undefined {
    const existing = this.sharedDefs.get(version)
    if (existing) {
      existing.refs++
      return existing.definitions
    }
    const release = this.host.admitEntry(`cf-defs:${version}`, defsBytes)
    if (!release) return undefined
    const shared = RemoteCfCache.freezeShared(definitions)
    this.sharedDefs.set(version, { definitions: shared, refs: 1, bytes: defsBytes, release })
    return shared
  }

  /** Ends one entry's shared-definitions reference; the last release drops
   * the version and settles its ledger charge. */
  private unshare(entry: Entry): void {
    if (entry.sharedVersion === undefined) return
    const version = entry.sharedVersion
    entry.sharedVersion = undefined
    const shared = this.sharedDefs.get(version)
    if (!shared) return
    shared.refs--
    if (shared.refs <= 0) {
      this.sharedDefs.delete(version)
      try { shared.release() } catch { /* release-only path */ }
    }
  }

  private drop(entry: Entry): void {
    if (entry.dropped) return
    entry.dropped = true
    this.pinned.delete(entry)
    this.unshare(entry)
    if (entry.releaseStates) {
      try { entry.releaseStates() } catch { /* release-only path */ }
      entry.releaseStates = undefined
    }
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    if (entry.releaseEntry) {
      try { entry.releaseEntry() } catch { /* release-only path */ }
      entry.releaseEntry = undefined
    }
    entry.releaseReply?.()
    entry.releaseReply = undefined
    entry.reply = undefined
    entry.result = undefined
  }

  private reset(entry: Entry): void {
    entry.seq++
    entry.controller?.abort()
    entry.controller = undefined
    this.unshare(entry)
    if (entry.releaseStates) {
      try { entry.releaseStates() } catch { /* release-only path */ }
      entry.releaseStates = undefined
    }
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
    entry.adopted = false
    // The entry (and its ledger-admitted metadata) is reused, not dropped.
    entry.bytes = entry.meta
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

  /** Retry pinned preparations when structure or receiver headroom changes. */
  wakeWaiting(): void {
    if (this.closed) return
    for (const entry of this.entries.values()) {
      if (entry.refs > 0 && entry.state === 'waiting' && !entry.controller) void this.load(entry)
    }
  }

  private async load(entry: Entry): Promise<void> {
    const { sheetId } = entry
    if (this.closed) return
    const completeness = this.host.loadState(sheetId).completeness
    if (completeness === 'loading') return
    if (completeness === 'incomplete') {
      entry.state = 'failed'
      entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-read-failed', message: 'Conditional formatting is unavailable because the sheet did not finish loading.' }]
      entry.gen++
      this.deferChanged(sheetId, entry.rects, 'status')
      return
    }
    const owner = this.host.identity(sheetId)
    if (entry.owner && !sameOwner(entry.owner, owner)) {
      this.reset(entry)
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
      if (this.closed || entry.detached || entry.seq !== seq) {
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
      // Bounded version adoption: a cold request (cfGen 0, "unknown, adopt
      // live") publishes the live reply under the adopted stamp; a first
      // stale-version reply rekeys and reloads once so a provider that raced
      // ahead still completes. Anything else mismatched fails closed, naming
      // the discovered version so the caller can re-request with it. Every
      // other stamp field and the transport owner must already match —
      // adoption never crosses owners, sheets or value generations.
      const want = entry.stamp
      const got = result.stamp
      const sameIdentity = !!got && Number.isSafeInteger(got.cfGen) && got.cfGen !== 0
        && got.owner === want.owner && got.sheetId === want.sheetId
        && got.workbookValueGen === want.workbookValueGen && got.viewGen === want.viewGen
        && got.structureGen === want.structureGen && got.clockGen === want.clockGen
        && sameOwner(out.owner as CacheIdentity, owner)
      if (sameIdentity && got.cfGen !== want.cfGen) {
        if (want.cfGen === 0 || !entry.adopted) {
          entry.adopted = true
          this.rekey(entry, { ...got })
          entry.gen++
          if (want.cfGen !== 0) {
            // Stale known version: the live reply belongs to the adopted
            // stamp, not this request — reload once under the new key.
            releaseOnce()
            entry.controller = undefined
            entry.state = 'waiting'
            entry.owner = undefined
            entry.pendingSheetIds = undefined
            this.deferChanged(sheetId, entry.rects, 'status')
            void this.load(entry)
            return
          }
          // Cold start: fall through and publish the live reply below.
        } else {
          releaseOnce()
          entry.state = 'failed'
          entry.diagnostics = [{ kind: 'unsupported', feature: 'cf-stale-generation', message: `Conditional formatting reply carried cfGen ${got.cfGen}; re-request with the live generations.` }]
          entry.gen++
          this.deferChanged(sheetId, entry.rects, 'status')
          return
        }
      }
      // Full-stamp validation: the reply must match the requested identity on
      // every field, or it is dropped without publishing — never paint stale.
      // A mismatch means the caller is behind; it must re-request with live
      // generations (failing closed here breaks retry loops on stale stamps).
      const live = entry.stamp
      const stampMismatch = !got || got.cfGen === 0 || got.owner !== live.owner || got.sheetId !== live.sheetId
        || got.workbookValueGen !== live.workbookValueGen || got.cfGen !== live.cfGen
        || got.viewGen !== live.viewGen || got.structureGen !== live.structureGen || got.clockGen !== live.clockGen
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
      // Single-charge publication: independent states charged per entry,
      // shared definitions charged once per version in the registry — never
      // a full-reply charge on top of a shared charge. Component sizes are
      // measured with admitted walk scratch (released right after); the
      // registry and states charges are admitted before publication. If any
      // part cannot be measured or admitted, the entry keeps its full reply
      // charge instead (legacy fallback that is still fail-closed, never
      // double-counted).
      const replyBytes = out.replyBytes ?? 0
      const split = this.measuredSplit(result, replyBytes)
      let shared: readonly CfSharedDefinition[] | undefined
      if (split !== undefined) {
        shared = this.shareDefinitions(result.definitionsVersion, result.definitions, split.defsBytes)
        if (shared !== undefined) entry.sharedVersion = result.definitionsVersion
      }
      let releaseStates: (() => void) | undefined
      if (shared !== undefined && split !== undefined) {
        const admitted = this.host.admitEntry(`cf-states:${requestId}`, split.statesBytes)
        if (admitted === null) {
          this.unshare(entry)
          shared = undefined
        } else {
          releaseStates = admitted
        }
      }
      if (shared === undefined || releaseStates === undefined || split === undefined) {
        // Legacy fallback: keep the transport's full reply charge.
        entry.sharedVersion = undefined
        const bytes = replyBytes + entry.meta
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
    } else {
      // Split ownership: settle the reply now that its components carry
      // the charge, then publish with per-entry states plus the shared
      // version. makeRoom sees the registry charge (admitted above) and
      // judges the entry on its own new total.
      const statesBytes = split.statesBytes
      const needed = entry.meta + statesBytes
      if (!Number.isSafeInteger(needed) || !this.makeRoom(entry, needed)) {
          this.unshare(entry)
          try { releaseStates() } catch { /* release-only path */ }
          releaseOnce()
          entry.state = 'failed'
          entry.diagnostics = [{ kind: 'resource-limit', feature: 'cf-cache-budget', message: 'Conditional formatting data exceeds the main-thread cache budget.' }]
          entry.gen++
          this.deferChanged(sheetId, entry.rects, 'status')
          return
        }
        this.host.release(out)
        entry.reply = undefined
        entry.releaseReply = undefined
        entry.releaseStates = releaseStates
        entry.result = { ...result, definitions: shared }
        entry.bytes = needed
      }
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
    // The current entry's own admitted metadata is already counted in
    // retainedBytes: room is judged on everyone else plus the new total.
    for (const [id, other] of this.entries) {
      if (this.retainedBytes - current.bytes + needed <= this.budgetBytes) break
      if (other === current || other.state !== 'ready' || other.refs > 0) continue
      this.entries.delete(id)
      other.detached = true
      other.gen++
      this.deferChanged(other.sheetId, other.rects, 'evicted')
      if (!other.dropped) this.drop(other)
    }
    return this.retainedBytes - current.bytes + needed <= this.budgetBytes
  }
}

function sameScope(a: readonly SourceRect[], b: readonly SourceRect[]): boolean {
  return rectsKey(a) === rectsKey(b)
}
