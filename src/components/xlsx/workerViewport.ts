/** Worker-owned bounded tile planning; host-owned canonical text layout/paint.
 * Plans share the existing wire byte cap and are charged to the main ledger.
 * Tiles the row cache already proves never ask the worker (tilePlanSource).
 * A tile plan spends the per-tile page cap from the SAME per-frame/per-tile
 * allowance the row cache's reads spend (§11.1); ring work may only use half
 * of the frame, so a visible tile always finds room. */
import type { CacheIdentity, LogicalViewport, ViewportResult } from '../../xlsx/view/contracts'
import { pageTileKey, type RemoteWorkbook, type RemoteSheetStore } from '../../worker/viewerRemoteSource'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from '../../worker/viewerQueryBudget'
import { emptyWork, type CellPlacement, type ViewportPlan } from '../../xlsx/view/viewport'
import type { ProjectedSheetMetrics } from '../../xlsx/view/projection'
import { formatA1 } from '../../xlsx/view/coordinates'
import { TILE_DEVICE_PX, TILE_PREFETCH_RING } from '../../xlsx/view/tileCache'

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
interface Entry { plan?: ViewportPlan; bytes: number; controller: AbortController; timer?: ReturnType<typeof setTimeout> }
let serial = 0
export class WorkerViewportPlans {
  private readonly entries = new Map<string, Entry>()
  private scope = ''
  private bytes = 0
  private closed = false
  private readonly ledgerOwner = `viewport-plans:${++serial}`
  constructor(private readonly remote: RemoteWorkbook) {}
  /** Retryable failures per tile key, for the backoff (bounded like entries). */
  private readonly failures = new Map<string, number>()
  beginFrame(owner: CacheIdentity, window: LogicalViewport): void {
    // One presented frame: plans and row reads share its page allowance.
    this.remote.beginPageFrame()
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
    const scope = JSON.stringify([owner, this.remote.merges(owner.sheetId).gen, window.zoom, window.deviceScale])
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
    // The request carries the frame owner its plan is stored under: a plan for
    // a newer worker view would otherwise be cached as this frame's (MINOR-3).
    void this.remote.viewport(owner, window, entry.controller.signal, 'tile').then(out => {
      if (this.closed || this.remote.isClosed || scope !== this.scope || this.entries.get(key) !== entry || entry.controller.signal.aborted) return
      if (typeof out !== 'object' || out.kind !== 'complete') {
        this.fail(key, entry, typeof out === 'object' ? out.diagnostics : [], retryableOutcome(out))
        return
      }
      if (out.result.content.kind === 'detail') this.remote.store(owner.sheetId).adoptViewport(out.result.content.cells)
      const plan = this.decode(out.result, extent)
      const bytes = queryPayloadBytes(plan, WORKER_QUERY_BYTES[this.remote.profile])
      while (this.bytes + bytes > WORKER_QUERY_BYTES[this.remote.profile] || this.entries.size > MAX_ENTRIES) {
        const oldest = [...this.entries].find(([k, e]) => k !== key && e.plan)
        if (!oldest) break
        this.drop(oldest[0], oldest[1])
      }
      if (this.bytes + bytes > WORKER_QUERY_BYTES[this.remote.profile]) {
        entry.plan = { kind: 'diagnostic', diagnostics: [{ code: 'resource-limit', message: 'Viewport plans exceed their bounded memory budget' }] }
        this.remote.mainLedger.settle(out)
        return
      }
      entry.bytes = bytes
      this.bytes += bytes
      this.remote.mainLedger.set('layout', this.ledgerOwner, this.bytes)
      entry.plan = plan
      this.remote.mainLedger.settle(out)
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
    }).finally(() => { this.remote.releasePages(pages); if (!this.closed) this.remote.notify() })
    return pending()
  }
  /** A failed tile shows its diagnostic; a retryable one expires after a
   * bounded backoff and the repaint that follows asks again (MINOR-2), so a
   * transient timeout or page error does not stick while the tile stays on
   * screen. A terminal failure stays until the tile leaves or the scope changes. */
  private fail(key: string, entry: Entry, diagnostics: ViewportPlan['diagnostics'], retryable: boolean): void {
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
  private decode(result: ViewportResult, extent: { width: number; height: number }): ViewportPlan {
    const cells = result.content.kind === 'detail' ? result.content.cells : []
    const unresolved = cells.filter(cell => cell.kind === 'unknown' || cell.kind === 'error').map(cell => ({ bounds: cell.bounds, error: cell.kind === 'error' }))
    const mergeAnchors = new Map((result.merges ?? []).map(m => [`${m.range.minRow}:${m.range.minCol}`, m.range]))
    const placed: CellPlacement[] = cells.flatMap(item => {
      if (item.kind !== 'cell' && item.kind !== 'unavailable') return []
      const style = this.remote.styleOf(item.cell.styleIndex)
      const merge = mergeAnchors.get(`${item.source.row}:${item.source.col}`)
      return [{ source: item.source, ...item.bounds, cell: { ...item.cell, ref: formatA1(item.source), value: item.cell.value ?? null, ...(style ? { style } : {}) }, ...(merge ? { merge } : {}) }]
    })
    const coarse = result.coarse ?? [], work = { ...(result.work ?? emptyWork()) }
    return { kind: 'detail', ...(unresolved.length ? { pending: unresolved } : {}), cells: placed, merges: (result.merges ?? []).map(m => ({ range: m.range, ...m.bounds })), drawings: [], rowLines: result.rowLines ?? [], colLines: result.colLines ?? [], extent, diagnostics: [], coarse, work }
  }
  private drop(key: string, entry: Entry): void {
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
}): (window: LogicalViewport, visible: boolean) => ViewportPlan | undefined {
  const { store, metrics, scale, identity, extent, plans } = options
  return (window, visible) => {
    const coverage = store.coverage(metrics, window.box, scale)
    if (coverage === 'proven') { plans.release(identity, window); return undefined }
    const plan = plans.get(identity, window, extent, visible)
    return coverage === 'oversized' || plan.kind === 'detail' ? plan : undefined
  }
}
