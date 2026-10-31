/** Main-thread two-phase client for worker-owned interactive pivot sessions. */
import type { PivotDelta } from '../xlsx/pivots/session/delta'
import type { PivotTableDefinition } from '../xlsx/pivots/model'
import type { PivotMenuModel } from '../xlsx/pivots/session/menu'
import type { ChartModel } from '../charts/model'
import type { ValueOverlayCell, WorkerReply } from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import type { ViewerClient } from './viewerClient'
import { WORKER_BUDGETS } from '../xlsx/view/tileCache'

export interface RemotePivotOpen { readonly kind: 'open'; readonly revision: number; readonly def: PivotTableDefinition; /** The menus' model (Task 28). */ readonly menu: PivotMenuModel; /** The session's layout cells (Task 29 autofit). */ readonly cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[] }
export type RemotePivotApply =
  | { readonly kind: 'committed'; readonly revision: number; readonly notice?: string; readonly def: PivotTableDefinition; readonly menu: PivotMenuModel; readonly rects: readonly SourceRect[]; readonly protectedRects: readonly SourceRect[]; readonly overlayExtent: SourceRect | null; readonly merges: readonly SourceRect[]; readonly usedRows: number; readonly usedCols: number; readonly charts: readonly RemotePivotChart[]; readonly cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[]; readonly valueGen: number }
  | { readonly kind: 'stale' | 'cancelled' }
  | { readonly kind: 'refused'; readonly message: string }

export interface RemotePivotApplyPayload {
  readonly revision: number
  readonly def: PivotTableDefinition
  readonly menu: PivotMenuModel
  readonly rects: readonly SourceRect[]
  readonly protectedRects: readonly SourceRect[]
  readonly overlayExtent: SourceRect | null
  readonly merges: readonly SourceRect[]
  readonly usedRows: number
  readonly usedCols: number
  readonly charts: readonly RemotePivotChart[]
  readonly cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[]
}
export interface RemotePivotChart { readonly sheetId: string; readonly partPath: string; readonly treePath: string; readonly chart: ChartModel }
export interface RemotePivotHost {
  readonly docId: string
  readonly profile: 'D' | 'M'
  readonly client: ViewerClient
  isClosed(): boolean
  /** Invoked only after the worker acknowledges the matching commit. */
  committed?(sheetId: number, pivot: number, result: RemotePivotApplyPayload, valueGen: number): void
  released?(sheetId: number, pivot: number, valueGen?: number): void
}
interface Entry { generation: number; revision: number; requestRevision: number; opened: boolean; openRequest?: string; applyRequest?: string; open?: Promise<RemotePivotOpen | { kind: 'refused'; message: string }> }
const keyOf = (sheetId: number, pivot: number) => JSON.stringify([sheetId, pivot])
const messageOf = (reply: WorkerReply): string => reply.kind === 'error'
  ? reply.diagnostics[0]?.message ?? `Pivot request failed (${reply.code})`
  : reply.kind === 'quota' ? 'The pivot operation exceeds the worker memory budget.'
    : reply.kind === 'cancelled' ? 'Pivot operation cancelled.' : 'Unexpected pivot worker reply.'

/** Owns request ordering on the receiver. A staged result is committed only if
 * it is still the latest request when its reply arrives; sending is not commit. */
export class RemotePivotSessions {
  private readonly entries = new Map<string, Entry>()
  private closed = false
  constructor(private readonly host: RemotePivotHost) {}

  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<RemotePivotOpen | { kind: 'refused'; message: string }> {
    const key = keyOf(sheetId, pivot)
    let entry = this.entries.get(key)
    if (entry?.open) return entry.open
    if (!entry) this.entries.set(key, entry = { generation: 0, revision: 0, requestRevision: 0, opened: false })
    if (this.closed || this.host.isClosed()) return Promise.resolve({ kind: 'refused', message: 'The workbook is closed.' })
    const requestId = this.host.client.nextRequestId()
    const generation = entry.generation
    entry.openRequest = requestId
    const request = this.host.client.send({ kind: 'pivotOpen', requestId, docId: this.host.docId, sheetId: String(sheetId), pivot, grantBytes: WORKER_BUDGETS[this.host.profile] / 4 })
    entry.open = this.awaitRequest(request, requestId, signal).then(reply => {
      if (generation !== entry!.generation || this.entries.get(key) !== entry) return { kind: 'refused' as const, message: 'The pivot session was released while opening.' }
      if (reply.kind === 'cancelled') return { kind: 'refused' as const, message: messageOf(reply) }
      if (reply.kind !== 'complete' || reply.result.kind !== 'pivotOpen') return { kind: 'refused' as const, message: messageOf(reply) }
      entry!.revision = reply.result.revision
      entry!.opened = true
      return { kind: 'open' as const, revision: reply.result.revision, def: reply.result.def, menu: reply.result.menu, cells: reply.result.cells }
    }).finally(() => { if (this.entries.get(key) === entry) { entry!.open = undefined; entry!.openRequest = undefined } })
    return entry.open
  }

  async apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<RemotePivotApply> {
    const key = keyOf(sheetId, pivot), entry = this.entries.get(key)
    if (this.closed || this.host.isClosed() || !entry?.opened) return { kind: 'refused', message: 'Open the pivot before changing it.' }
    const generation = ++entry.generation, baseRevision = entry.revision, requestRevision = ++entry.requestRevision
    const requestId = this.host.client.nextRequestId()
    entry.applyRequest = requestId
    let staged: RemotePivotApplyPayload | undefined
    try {
      const reply = await this.host.client.send({ kind: 'pivotApply', requestId, docId: this.host.docId, sheetId: String(sheetId), pivot, baseRevision, revision: requestRevision, delta, grantBytes: WORKER_BUDGETS[this.host.profile] / 4 })
      if (generation !== entry.generation) return { kind: 'stale' }
      if (reply.kind !== 'complete' || reply.result.kind !== 'pivotApply') return reply.kind === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', message: messageOf(reply) }
      staged = reply.result
      const revision = staged.revision
      const committed = await this.host.client.send({ kind: 'pivotCommit', requestId: this.host.client.nextRequestId(), docId: this.host.docId, sheetId: String(sheetId), pivot, revision })
      if (committed.kind !== 'complete' || committed.result.kind !== 'pivotCommit') return committed.kind === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', message: messageOf(committed) }
      if ('refused' in committed.result) return { kind: 'stale' }
      // An acknowledgement is the worker's committed fact. Publish it even when a newer
      // request was sent; only suppress an older acknowledgement already overtaken by a
      // later committed revision, or one whose owner was released.
      const registered = this.entries.get(key) === entry
      const publish = registered && revision > entry.revision
      if (publish) {
        entry.revision = revision
        this.host.committed?.(sheetId, pivot, staged, committed.result.valueGen)
      }
      return generation !== entry.generation || !registered ? { kind: 'stale' } : { ...staged, kind: 'committed', valueGen: committed.result.valueGen }
    } finally { if (entry.applyRequest === requestId) entry.applyRequest = undefined }
  }

  release(sheetId: number, pivot: number): void {
    const key = keyOf(sheetId, pivot), entry = this.entries.get(key)
    if (!entry) return
    this.entries.delete(key); entry.generation++
    if (entry.openRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.openRequest }).catch(() => {})
    if (entry.applyRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.applyRequest }).catch(() => {})
    this.host.released?.(sheetId, pivot)
    if (!this.closed && !this.host.isClosed()) void this.host.client.send({ kind: 'pivotRelease', requestId: this.host.client.nextRequestId(), docId: this.host.docId, sheetId: String(sheetId), pivot }).then(reply => {
      if (reply.kind === 'complete' && reply.result.kind === 'pivotRelease') this.host.released?.(sheetId, pivot, reply.result.valueGen)
    }).catch(() => {})
  }

  releaseExcept(sheetId: number): void {
    for (const [key] of [...this.entries]) {
      const [candidate, pivot] = JSON.parse(key) as [number, number]
      if (candidate !== sheetId) this.release(candidate, pivot)
    }
  }

  close(): void {
    if (this.closed) return
    for (const [key] of this.entries) {
      const [sheet, pivot] = JSON.parse(key) as [number, number]
      this.release(sheet, pivot)
    }
    this.closed = true
    this.entries.clear()
  }

  private async awaitRequest(request: Promise<WorkerReply<'pivotOpen'>>, requestId: string, signal: AbortSignal): Promise<WorkerReply<'pivotOpen'>> {
    if (signal.aborted) {
      void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: requestId }).catch(() => {})
      return { kind: 'cancelled', commandKind: 'pivotOpen', requestId, owner: { docId: this.host.docId }, diagnostics: [] }
    }
    let abort!: () => void
    const aborted = new Promise<WorkerReply<'pivotOpen'>>(resolve => {
      abort = () => {
        void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: requestId }).catch(() => {})
        resolve({ kind: 'cancelled', commandKind: 'pivotOpen', requestId, owner: { docId: this.host.docId }, diagnostics: [] })
      }
      signal.addEventListener('abort', abort, { once: true })
    })
    try { return await Promise.race([request, aborted]) }
    finally { signal.removeEventListener('abort', abort) }
  }
}
