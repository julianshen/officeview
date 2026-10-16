import { validateOverflowProof, viewportPlanAdmission, VIEWPORT_MEASUREMENT_BYTES, type OverflowViewportResult, type OverflowEntry } from '../../xlsx/view/viewportOverflow'
/** Worker-owned bounded tile planning; host-owned canonical text layout/paint.
 * Plans share the existing wire byte cap and are charged to the main ledger.
 * Tiles the row cache already proves never ask the worker (tilePlanSource).
 * A tile plan spends the per-tile page cap from the SAME per-frame/per-tile
 * allowance the row cache's reads spend (§11.1); ring work may only use half
 * of the frame, so a visible tile always finds room. */
import type { CacheIdentity, LogicalViewport } from '../../xlsx/view/contracts'
import { identityFields, sameIdentity } from '../../xlsx/view/cacheIdentity'
import { pageTileKey, type RemoteWorkbook, type RemoteSheetStore } from '../../worker/viewerRemoteSource'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from '../../worker/viewerQueryBudget'
import { emptyWork, type CellPlacement, type ViewportPlan } from '../../xlsx/view/viewport'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { formatA1 } from '../../xlsx/view/coordinates'
import { TILE_DEVICE_PX, TILE_PREFETCH_RING, MAIN_THREAD_BUDGETS } from '../../xlsx/view/tileCache'

const pending = (): ViewportPlan => ({ kind: 'pending', diagnostics: [] })
/** At most this many tile entries (plans, failures, requests) are retained. */
const MAX_ENTRIES = 32
/** A retryable tile failure expires after a backoff that doubles per
 * consecutive failure of the same tile, bounded by the maximum. */
export const PLAN_RETRY_MS = 250
export const PLAN_RETRY_MAX_MS = 4000
/** Whether a settled tile outcome may be re-requested: anything but an
 * explicitly terminal failure (a cancelled read, a timeout, a stale owner or a
 * page error can clear; an error with retryable: false cannot). */
function retryableOutcome(out: unknown): boolean {
  if (typeof out !== 'object' || out === null) return true
  if ((out as { kind?: string }).kind === 'error') return (out as { retryable?: boolean }).retryable !== false
  return true
}
const scopeKey = (owner: CacheIdentity, merges: number, window: LogicalViewport, layout: number): string => JSON.stringify([...identityFields.map(field => owner[field]), merges, window.zoom, window.deviceScale, layout])
interface Entry { plan?: ViewportPlan; bytes: number; controller: AbortController; timer?: ReturnType<typeof setTimeout> }
let serial = 0
export class WorkerViewportPlans {
  private readonly entries = new Map<string, Entry>()
  private scope = ''
  private bytes = 0
  private ownershipRevision = 0
  private closed = false
  /** The host's session-size generation (spec §5.6): a plan requested under
   * another layout is out of scope, so its reply is refused. */
  private layoutGen = 0
  private readonly ledgerOwner = `viewport-plans:${++serial}`
  constructor(private readonly remote: RemoteWorkbook) {}
  /** Retryable failures per tile key, for the backoff (bounded like entries). */
  private readonly failures = new Map<string, number>()
  beginFrame(owner: CacheIdentity, window: LogicalViewport, layoutGen = 0): void {
    // One presented frame: plans and row reads share its page allowance.
    this.remote.beginPageFrame()
    this.layoutGen = layoutGen
    this.select(owner, window)
    // Requests and failures off the pane and its prefetch ring are obsolete;
    // settled plans stay until the byte/entry caps or a scope change.
    const reach = TILE_PREFETCH_RING * TILE_DEVICE_PX / (window.zoom * window.deviceScale)
    const box = { x: window.box.x - reach, y: window.box.y - reach, width: window.box.width + 2 * reach, height: window.box.height + 2 * reach }
    for (const [key, entry] of [...this.entries]) {
      if (entry.plan?.kind === 'detail' && !entry.plan.pending?.length && !entry.timer || entry.plan?.kind === 'coarse') continue
      const [, x, y, width, height] = JSON.parse(key) as [string, number, number, number, number]
      if (x >= box.x + box.width || x + width <= box.x || y >= box.y + box.height || y + height <= box.y) {
        this.drop(key, entry)
        this.failures.delete(key)
      }
    }
  }
  private select(owner: CacheIdentity, window: LogicalViewport): string {
    const scope = scopeKey(owner, this.remote.merges(owner.sheetId).gen, window, this.layoutGen)
    if (scope !== this.scope) { this.clear(); this.scope = scope }
    return scope
  }
  /** Forgets a tile the row cache now proves: its plan (and bytes) or its
   * request are no longer needed. */
  release(owner: CacheIdentity, window: LogicalViewport): void {
    if (this.closed) return
    this.select(owner, window)
    const key = pageTileKey(owner.sheetId, window.box), entry = this.entries.get(key)
    if (entry) this.drop(key, entry)
    this.failures.delete(key)
  }
  get(owner: CacheIdentity, window: LogicalViewport, extent: { width: number; height: number }, visible: boolean): ViewportPlan {
    if (this.closed || this.remote.isClosed) return pending()
    const scope = this.select(owner, window)
    const expectedOwner = { ...owner }, mergeGeneration = this.remote.merges(owner.sheetId).gen, zoom = window.zoom, deviceScale = window.deviceScale
    const key = pageTileKey(owner.sheetId, window.box), existing = this.entries.get(key)
    if (existing) { this.entries.delete(key); this.entries.set(key, existing); return existing.plan ?? pending() }
    // The worker may spend up to the per-tile page cap on this plan, so it
    // needs the tile's whole unspent allowance in this frame (a tile whose rows
    // were already requested this frame gets no plan), within the ring share
    // for a prefetch tile.
    const pages = this.remote.pageLimits.tile, frame = this.remote.currentPageFrame
    const budgetKey = pageTileKey(owner.sheetId, window.box)
    if (!this.remote.admitPages(frame, budgetKey, pages, !visible, false)) return pending()
    if (this.entries.size >= MAX_ENTRIES) {
      const oldest = [...this.entries].find(([, e]) => e.plan)
      if (oldest) this.drop(oldest[0], oldest[1])
      else return pending()
    }
    this.remote.admitPages(frame, budgetKey, pages, !visible)
    const entry: Entry = { bytes: 0, controller: new AbortController() }
    this.entries.set(key, entry)
    this.ownershipRevision++
    const store = this.remote.store(owner.sheetId)
    const sourceRequest = store.captureViewportRequest(entry.controller.signal)
    // Host fonts scope the cached plan; source geometry and its proof carry
    // the worker's font owner. Keep a private wire snapshot so neither a caller
    // mutation nor a newer worker view can be published under this host frame.
    const sourceOwner = { ...expectedOwner, fontGen: this.remote.identity(owner.sheetId).fontGen }
    void this.remote.viewport({ ...sourceOwner }, window, entry.controller.signal, 'tile').then(out => {
      if (this.closed || this.remote.isClosed || scope !== this.scope || this.entries.get(key) !== entry || entry.controller.signal.aborted) return
      if (typeof out !== 'object' || out.kind !== 'complete') {
        this.fail(key, entry, typeof out === 'object' ? out.diagnostics : [], retryableOutcome(out))
        return
      }
      const result = out.result as OverflowViewportResult
      if (!sourceRequest.live()) { this.fail(key, entry, [{ code: 'stale-source', message: 'Viewport source cache was cleared' }], true); return }
      const cap = WORKER_QUERY_BYTES[this.remote.profile]
      const reservation = `${this.ledgerOwner}:stage:${++serial}`
      const measurement = `${reservation}:measure`
      let measurementLive = false, reservationLive = false
      let adoptionToken: ReturnType<typeof store.stageViewport>
      try {
        // The small bounded measurement stack is itself pre-admitted. Larger
        // indexes, grouping, copies and sizing walks wait for the joint bound.
        if (this.remote.mainLedger.bytes + VIEWPORT_MEASUREMENT_BYTES > MAIN_THREAD_BUDGETS[this.remote.profile]) {
          this.fail(key, entry, [{ code: 'resource-limit', message: 'Viewport measurement exceeds the main memory budget' }], true)
          return
        }
        measurementLive = true
        this.remote.mainLedger.set('layout', measurement, VIEWPORT_MEASUREMENT_BYTES)
        const cells = result.content.kind === 'detail' ? result.content.cells : []
        const layout = this.remote.layout(owner.sheetId)
        const admission = viewportPlanAdmission(out, result, this.remote.formatting().styles, store.viewportAdoptionBytes(cells), store.viewportRollbackBytes(), layout.colOverrides.length + (layout.hiddenCols?.length ?? 0) + 1, cap, this.bytes)
        if (admission.payload > cap || !Number.isSafeInteger(admission.joint) || admission.joint < 0) throw new Error('Viewport reply exceeds its bounded byte budget')
        // The measurement owner was charged above and remains live here;
        // joint replaces that charge rather than adding a second stack.
        if (this.remote.mainLedger.bytes - VIEWPORT_MEASUREMENT_BYTES + admission.joint > MAIN_THREAD_BUDGETS[this.remote.profile]) {
          this.fail(key, entry, [{ code: 'resource-limit', message: 'Viewport staging exceeds the main memory budget' }], true)
          return
        }
        reservationLive = true
        this.remote.mainLedger.set('layout', reservation, admission.reserve)
        this.remote.mainLedger.release('layout', measurement)
        measurementLive = false
        if (queryPayloadBytes(out, cap) > cap) throw new Error('Viewport reply exceeds its bounded byte budget')
        if (!sameIdentity(out.owner, sourceOwner)) throw new Error('Viewport reply owner mismatch')
        // Shared with remote.geometry(), but created only after joint
        // admission and without materializing full-sheet row geometry.
        const columns = this.remote.columns(owner.sheetId)
        const proof = validateOverflowProof(result, sourceOwner, window, this.remote.profile, 'tile', columns)
        const plan = this.decode(result, extent, proof)
        proof.clear()
        const bytes = queryPayloadBytes(plan, cap)
        if (bytes > admission.planUpper) throw new Error('Viewport construction exceeded its admitted estimate')
        let retainedBytes = this.bytes + bytes
        const evictions: Array<[string, Entry]> = []
        for (const [oldKey, oldEntry] of this.entries) {
          if (retainedBytes <= cap) break
          if (oldKey === key || !oldEntry.plan) continue
          evictions.push([oldKey, oldEntry])
          retainedBytes -= oldEntry.bytes
        }
        if (retainedBytes > cap) throw new Error('Viewport plans exceed their bounded memory budget')
        const guard = () => sourceRequest.live() && !this.closed && !this.remote.isClosed && scope === this.scope && this.entries.get(key) === entry && !entry.controller.signal.aborted && sameIdentity(this.remote.identity(owner.sheetId), sourceOwner) && sameIdentity(owner, expectedOwner) && this.remote.merges(owner.sheetId).gen === mergeGeneration && window.zoom === zoom && window.deviceScale === deviceScale
        adoptionToken = store.stageViewport(cells, entry.controller.signal, { live: guard })
        if (!guard()) {
          // Same retryable refusal, but name the actual cause: a cleared
          // source cache is not an owner change.
          const superseded = this.closed || this.remote.isClosed || scope !== this.scope || this.entries.get(key) !== entry || entry.controller.signal.aborted
          const message = superseded ? 'Viewport request was superseded. Retrying.'
            : !sourceRequest.live() ? 'Viewport source cache was cleared'
              : !sameIdentity(this.remote.identity(owner.sheetId), sourceOwner) || !sameIdentity(owner, expectedOwner) ? 'Viewport owner changed. Retrying.'
                : 'Viewport layout changed. Retrying.'
          this.fail(key, entry, [{ code: 'stale-source', message }], true)
          return
        }
        const previousPlan = entry.plan, previousEntryBytes = entry.bytes
        const rollback = () => {
          if (this.remote.isClosed) { this.dispose(); return }
          if (this.closed || scope !== this.scope) {
            this.remote.mainLedger.set('layout', this.ledgerOwner, this.bytes)
            return
          }
          if (this.entries.get(key) === entry && !entry.controller.signal.aborted) {
            entry.bytes = previousEntryBytes
            entry.plan = previousPlan
          }
          for (const [oldKey, oldEntry] of evictions) {
            if (!oldEntry.controller.signal.aborted && !this.entries.has(oldKey)) this.entries.set(oldKey, oldEntry)
          }
          this.bytes = 0
          for (const retained of this.entries.values()) this.bytes += retained.bytes
          this.ownershipRevision++
          this.remote.mainLedger.set('layout', this.ledgerOwner, this.bytes)
        }
        const currentRetainedBytes = () => {
          let total = this.bytes + bytes
          for (const [oldKey, oldEntry] of evictions) if (this.entries.get(oldKey) === oldEntry) total -= oldEntry.bytes
          return total
        }
        const publish = () => {
          if (!guard()) throw new Error('Viewport owner changed before publication')
          retainedBytes = currentRetainedBytes()
          const ownership = this.ownershipRevision
          this.remote.mainLedger.set('layout', this.ledgerOwner, retainedBytes)
          if (!guard()) throw new Error('Viewport owner changed during publication')
          const reconciled = currentRetainedBytes()
          if (reconciled !== retainedBytes || this.ownershipRevision !== ownership) {
            const reconciliationOwnership = this.ownershipRevision
            this.remote.mainLedger.set('layout', this.ledgerOwner, reconciled)
            if (!guard() || currentRetainedBytes() !== reconciled || this.ownershipRevision !== reconciliationOwnership) throw new Error('Viewport plan accounting changed during publication')
            retainedBytes = reconciled
          }
          entry.bytes = bytes
          this.bytes = retainedBytes
          entry.plan = plan
          for (const [oldKey, oldEntry] of evictions) if (this.entries.get(oldKey) === oldEntry) this.entries.delete(oldKey)
          this.ownershipRevision++
        }
        try {
          if (adoptionToken) {
            if (!adoptionToken.commit(publish)) {
              rollback()
              const error = adoptionToken.error
              this.fail(key, entry, [{ code: error?.code ?? 'stale-source', message: error?.message ?? 'Viewport could not be updated. Retrying.' }], error?.retryable ?? true)
              return
            }
          } else publish()
        } catch (cause) { rollback(); throw cause }
        // Abort evicted owners only after success; rollback must keep their
        // original plans, timers and controllers alive.
        for (const [, oldEntry] of evictions) { oldEntry.controller.abort(); clearTimeout(oldEntry.timer) }
        this.remote.mainLedger.settle(out)
      } catch (cause) {
        this.fail(key, entry, [{ code: 'resource-limit', message: cause instanceof Error ? cause.message : 'Invalid worker viewport' }], false)
        return
      } finally {
        adoptionToken?.close()
        if (reservationLive) this.remote.mainLedger.release('layout', reservation)
        if (measurementLive) this.remote.mainLedger.release('layout', measurement)
      }
      if (out.result.completeness === 'complete') { this.failures.delete(key); return }
      // A partial result that failed (a retryable page error) keeps its error
      // shading and backs off like any other failure (r3 #4); a terminal one stays.
      if (out.result.error) { if (out.result.error.retryable) this.backoff(key, entry); return }
      // Loading snapshots refresh at a bounded cadence; an import can finish
      // after its last metadata notification, so own this last retry timer.
      if (this.remote.loadState(owner.sheetId).completeness !== 'incomplete') entry.timer = setTimeout(() => {
        if (this.entries.get(key) === entry) { this.drop(key, entry); this.remote.notify() }
      }, 100)
    }).catch(() => {
      if (!this.closed && this.entries.get(key) === entry) this.fail(key, entry, [{ code: 'page-error', message: 'Viewport query failed' }], true)
    }).finally(() => { sourceRequest.close(); this.remote.releasePages(pages); if (!this.closed) this.remote.notify() })
    return pending()
  }
  /** A failed tile shows its diagnostic; a retryable one expires after a
   * bounded backoff and the repaint that follows asks again (MINOR-2), so a
   * transient timeout or page error does not stick while the tile stays on
   * screen. A terminal failure stays until the tile leaves or the scope changes. */
  private fail(key: string, entry: Entry, diagnostics: ViewportPlan['diagnostics'], retryable: boolean): void {
    if (this.closed || this.remote.isClosed || this.entries.get(key) !== entry || entry.controller.signal.aborted) return
    entry.plan = { kind: 'diagnostic', diagnostics }
    if (retryable) this.backoff(key, entry)
  }
  /** Expires a failed entry after a backoff doubling per consecutive failure
   * of the same tile, bounded by PLAN_RETRY_MAX_MS. */
  private backoff(key: string, entry: Entry): void {
    const attempt = (this.failures.get(key) ?? 0) + 1
    this.failures.delete(key)
    this.failures.set(key, attempt)
    if (this.failures.size > MAX_ENTRIES) this.failures.delete(this.failures.keys().next().value as string)
    entry.timer = setTimeout(() => {
      if (this.entries.get(key) === entry) { this.drop(key, entry); this.remote.notify() }
    }, Math.min(PLAN_RETRY_MAX_MS, PLAN_RETRY_MS * 2 ** (attempt - 1)))
  }
  private decode(result: OverflowViewportResult, extent: { width: number; height: number }, proof: Map<number, OverflowEntry>): ViewportPlan {
    const cells = result.content.kind === 'detail' ? result.content.cells : []
    const unresolved = [...(result.overflowProof?.pending ?? []).map(item => ({ bounds: { ...item.bounds }, error: item.error })), ...cells.filter(cell => cell.kind === 'unknown' || cell.kind === 'error').map(cell => ({ bounds: { ...cell.bounds }, error: cell.kind === 'error' }))]
    const mergeAnchors = new Map((result.merges ?? []).map(m => [`${m.range.minRow}:${m.range.minCol}`, m.range]))
    const placed: CellPlacement[] = cells.flatMap(item => {
      if (item.kind !== 'cell' && item.kind !== 'unavailable') return []
      const style = this.remote.styleOf(item.cell.styleIndex)
      const merge = mergeAnchors.get(`${item.source.row}:${item.source.col}`)
      const overflow = proof.get(item.source.row * 16384 + item.source.col)!
      return [{ overflow: { left: overflow.left, right: overflow.right }, source: { ...item.source }, ...item.bounds, cell: { ...item.cell, ...(item.cell.sharedFormula ? { sharedFormula: { ...item.cell.sharedFormula } } : {}), ref: formatA1(item.source), value: item.cell.value ?? null, ...(style ? { style } : {}) }, ...(merge ? { merge: { ...merge } } : {}) }]
    })
    const coarse = (result.coarse ?? []).map(group => ({ ...group, bounds: { ...group.bounds }, slots: { ...group.slots }, cols: { ...group.cols } })), work = { ...(result.work ?? emptyWork()) }
    return { kind: 'detail', ...(unresolved.length ? { pending: unresolved } : {}), cells: placed, merges: (result.merges ?? []).map(m => ({ range: { ...m.range }, ...m.bounds })), drawings: [], rowLines: [...(result.rowLines ?? [])], colLines: [...(result.colLines ?? [])], extent, diagnostics: [], coarse, work }
  }
  private drop(key: string, entry: Entry): void {
    this.ownershipRevision++
    entry.controller.abort()
    clearTimeout(entry.timer)
    this.entries.delete(key)
    this.bytes -= entry.bytes
    this.remote.mainLedger.set('layout', this.ledgerOwner, this.bytes)
  }
  private clear(): void { for (const [key, entry] of [...this.entries]) this.drop(key, entry); this.failures.clear() }
  dispose(): void { this.closed = true; this.clear() }
}
/** The frame's tile plan source (S13.3 review round 2). A tile the row cache
 * already proves paints from the main-thread planner at once: no worker
 * command, no wait (warm scroll and revisits, §12.2). A tile the row cache
 * does not know asks the worker for its plan; until a detail plan lands it
 * paints from the main-thread planner too, whose `pendingIn` requests the rows
 * that make its cells (and the accessible grid) known, so a slow or failed
 * plan never leaves blanks reading "Loading…" (MAJOR-1). Only a tile spanning
 * more rows than the row cache may cover depends on the worker plan alone and
 * shows its pending or diagnostic state. */
export function tilePlanSource(options: {
  store: RemoteSheetStore
  metrics: ProjectedSheetMetrics
  scale: number
  identity: CacheIdentity
  extent: { width: number; height: number }
  plans: WorkerViewportPlans
  /** The sheet has view-only session sizes (spec §5.6): worker plans are laid
   * out in the workbook geometry, so the main-thread planner is authoritative
   * and the worker is never asked. A tile only a worker plan could cover
   * (oversized) is summarized coarse until zoomed in. */
  resized?: boolean
}): (window: LogicalViewport, visible: boolean) => ViewportPlan | undefined {
  const { store, metrics, scale, identity, extent, plans } = options
  return (window, visible) => {
    const coverage = store.coverage(metrics, window.box, scale)
    if (options.resized) {
      return coverage === 'oversized'
        ? { kind: 'coarse', diagnostics: [], groups: [{ groupId: 'resized-oversized', bounds: { ...window.box }, slots: { start: 0, end: 0 }, cols: { c0: 0, c1: 0 } }], extent, work: emptyWork() }
        : undefined
    }
    if (coverage === 'proven') { plans.release(identity, window); return undefined }
    const plan = plans.get(identity, window, extent, visible)
    return coverage === 'oversized' || plan.kind === 'detail' ? plan : undefined
  }
}
