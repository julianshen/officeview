/** Worker-owned bounded tile planning; host-owned canonical text layout/paint.
 * Plans share the existing wire byte cap and are charged to the main ledger.
 * Only visible tiles request work; ring tiles may reuse an already loaded plan. */
import type { CacheIdentity, LogicalViewport, ViewportResult } from '../../xlsx/view/contracts'
import type { RemoteWorkbook } from '../../worker/viewerRemoteSource'
import { queryPayloadBytes, WORKER_QUERY_BYTES } from '../../worker/viewerQueryBudget'
import { emptyWork, type CellPlacement, type ViewportPlan } from '../../xlsx/view/viewport'
import { formatA1 } from '../../xlsx/view/coordinates'

const pending = (): ViewportPlan => ({ kind: 'pending', diagnostics: [] })
interface Entry { plan?: ViewportPlan; bytes: number; controller: AbortController; timer?: ReturnType<typeof setTimeout> }
let serial = 0
export class WorkerViewportPlans {
  private readonly entries = new Map<string, Entry>()
  private scope = ''
  private bytes = 0
  private inFlight = 0
  private closed = false
  private readonly ledgerOwner = `viewport-plans:${++serial}`
  constructor(private readonly remote: RemoteWorkbook) {}
  beginFrame(owner: CacheIdentity, window: LogicalViewport): void {
    this.select(owner, window)
    const box = window.box
    for (const [key, entry] of [...this.entries]) {
      if (entry.plan?.kind === 'detail' && !entry.plan.pending?.length && !entry.timer || entry.plan?.kind === 'coarse') continue
      const tile = JSON.parse(key) as LogicalViewport['box']
      if (tile.x >= box.x + box.width || tile.x + tile.width <= box.x || tile.y >= box.y + box.height || tile.y + tile.height <= box.y) this.drop(key, entry)
    }
  }
  private select(owner: CacheIdentity, window: LogicalViewport): string {
    const scope = JSON.stringify([owner, this.remote.merges(owner.sheetId).gen, window.zoom, window.deviceScale])
    if (scope !== this.scope) { this.clear(); this.scope = scope }
    return scope
  }
  get(owner: CacheIdentity, window: LogicalViewport, extent: { width: number; height: number }, visible: boolean): ViewportPlan {
    if (this.closed || this.remote.isClosed) return pending()
    const scope = this.select(owner, window)
    const key = JSON.stringify(window.box), existing = this.entries.get(key)
    if (existing) { this.entries.delete(key); this.entries.set(key, existing); return existing.plan ?? pending() }
    // Four tile requests consume at most the profile's per-frame page budget.
    if (!visible || this.inFlight >= 4) return pending()
    if (this.entries.size >= 32) {
      const oldest = [...this.entries].find(([, e]) => e.plan)
      if (oldest) this.drop(oldest[0], oldest[1])
      else return pending()
    }
    const entry: Entry = { bytes: 0, controller: new AbortController() }
    this.entries.set(key, entry)
    this.inFlight++
    void this.remote.viewport(this.remote.identity(owner.sheetId), window, entry.controller.signal, 'tile').then(out => {
      if (this.closed || this.remote.isClosed || scope !== this.scope || this.entries.get(key) !== entry || entry.controller.signal.aborted) return
      if (typeof out !== 'object' || out.kind !== 'complete') {
        entry.plan = { kind: 'diagnostic', diagnostics: typeof out === 'object' ? out.diagnostics : [] }
        return
      }
      if (out.result.content.kind === 'detail') this.remote.store(owner.sheetId).adoptViewport(out.result.content.cells)
      const plan = this.decode(out.result, extent)
      const bytes = queryPayloadBytes(plan, WORKER_QUERY_BYTES[this.remote.profile])
      while (this.bytes + bytes > WORKER_QUERY_BYTES[this.remote.profile] || this.entries.size > 32) {
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
      // Loading snapshots refresh at a bounded cadence; an import can finish
      // after its last metadata notification, so own this last retry timer.
      if (out.result.completeness !== 'complete' && (!out.result.error || out.result.error.retryable) && this.remote.loadState(owner.sheetId).completeness !== 'incomplete') entry.timer = setTimeout(() => {
        if (this.entries.get(key) === entry) { this.drop(key, entry); this.remote.notify() }
      }, 100)
    }).catch(() => { if (this.entries.get(key) === entry) entry.plan = { kind: 'diagnostic', diagnostics: [{ code: 'page-error', message: 'Viewport query failed' }] } }).finally(() => { this.inFlight--; if (!this.closed) this.remote.notify() })
    return pending()
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
  private clear(): void { for (const [key, entry] of [...this.entries]) this.drop(key, entry) }
  dispose(): void { this.closed = true; this.clear() }
}
