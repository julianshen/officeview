/** Main-thread two-phase client for worker-owned interactive pivot sessions. */
import type { PivotDelta } from '../xlsx/pivots/session/delta'
import type { PivotTableDefinition } from '../xlsx/pivots/model'
import type { ChartModel } from '../charts/model'
import type { ValueOverlayCell, WorkerReply, ValueOverlay } from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import type { ViewerClient } from './viewerClient'
import { WORKER_BUDGETS } from '../xlsx/view/tileCache'

export interface RemotePivotOpen { readonly kind: 'open'; readonly revision: number; readonly def: PivotTableDefinition }
export type RemotePivotApply =
  | { readonly kind: 'committed'; readonly revision: number; readonly def: PivotTableDefinition; readonly rects: readonly SourceRect[]; readonly protectedRects: readonly SourceRect[]; readonly merges: readonly SourceRect[]; readonly usedRows: number; readonly usedCols: number; readonly charts: readonly ChartModel[]; readonly cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[]; readonly valueGen: number }
  | { readonly kind: 'stale' | 'cancelled' }
  | { readonly kind: 'refused'; readonly message: string }

export interface RemotePivotApplyPayload {
  readonly revision: number
  readonly def: PivotTableDefinition
  readonly rects: readonly SourceRect[]
  readonly protectedRects: readonly SourceRect[]
  readonly merges: readonly SourceRect[]
  readonly usedRows: number
  readonly usedCols: number
  readonly charts: readonly ChartModel[]
  readonly cells: readonly { row: number; col: number; value: ValueOverlayCell | null }[]
}
export interface RemotePivotHost {
  readonly docId: string
  readonly profile: 'D' | 'M'
  readonly client: ViewerClient
  isClosed(): boolean
  /** Invoked only after the worker acknowledges the matching commit. */
  committed?(sheetId: number, pivot: number, result: RemotePivotApplyPayload, valueGen: number): void
  /** Non-paged object documents are cloned to the main thread. Mirror the accepted
   * worker commit into that clone's store; paged workers omit this hook. */
  mirror?(sheetId: number, overlay: ValueOverlay | undefined): void
  released?(sheetId: number, pivot: number, valueGen?: number): void
}
interface Entry { generation: number; revision: number; opened: boolean; openRequest?: string; applyRequest?: string; open?: Promise<RemotePivotOpen | { kind: 'refused'; message: string }> }
const keyOf = (sheetId: number, pivot: number) => JSON.stringify([sheetId, pivot])
const messageOf = (reply: WorkerReply): string => reply.kind === 'error'
  ? reply.diagnostics[0]?.message ?? `Pivot request failed (${reply.code})`
  : reply.kind === 'quota' ? 'The pivot operation exceeds the worker memory budget.'
    : reply.kind === 'cancelled' ? 'Pivot operation cancelled.' : 'Unexpected pivot worker reply.'

/** Owns request ordering on the receiver. A staged result is committed only if
 * it is still the latest request when its reply arrives; sending is not commit. */
export class RemotePivotSessions {
  private readonly entries = new Map<string, Entry>()
  private readonly mirrored = new Map<number, Map<number, ValueOverlay>>()
  private closed = false
  constructor(private readonly host: RemotePivotHost) {}

  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<RemotePivotOpen | { kind: 'refused'; message: string }> {
    const key = keyOf(sheetId, pivot)
    let entry = this.entries.get(key)
    if (entry?.open) return entry.open
    if (!entry) this.entries.set(key, entry = { generation: 0, revision: 0, opened: false })
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
      return { kind: 'open' as const, revision: reply.result.revision, def: reply.result.def }
    }).finally(() => { if (this.entries.get(key) === entry) { entry!.open = undefined; entry!.openRequest = undefined } })
    return entry.open
  }

  async apply(sheetId: number, pivot: number, delta: PivotDelta | 'reset'): Promise<RemotePivotApply> {
    const key = keyOf(sheetId, pivot), entry = this.entries.get(key)
    if (this.closed || this.host.isClosed() || !entry?.opened) return { kind: 'refused', message: 'Open the pivot before changing it.' }
    const generation = ++entry.generation, revision = entry.revision + 1
    const requestId = this.host.client.nextRequestId()
    entry.applyRequest = requestId
    let staged: RemotePivotApplyPayload | undefined
    try {
      const reply = await this.host.client.send({ kind: 'pivotApply', requestId, docId: this.host.docId, sheetId: String(sheetId), pivot, revision, delta, grantBytes: WORKER_BUDGETS[this.host.profile] / 4 })
      if (generation !== entry.generation) return { kind: 'stale' }
      if (reply.kind !== 'complete' || reply.result.kind !== 'pivotApply') return reply.kind === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', message: messageOf(reply) }
      staged = reply.result
      const committed = await this.host.client.send({ kind: 'pivotCommit', requestId: this.host.client.nextRequestId(), docId: this.host.docId, sheetId: String(sheetId), pivot, revision })
      if (generation !== entry.generation) return { kind: 'stale' }
      if (committed.kind !== 'complete' || committed.result.kind !== 'pivotCommit') return committed.kind === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', message: messageOf(committed) }
      if ('refused' in committed.result) return { kind: 'stale' }
      entry.revision = revision
      let sheetOverlays = this.mirrored.get(sheetId)
      if (!sheetOverlays) this.mirrored.set(sheetId, sheetOverlays = new Map())
      sheetOverlays.set(pivot, valueOverlayFromPivotApply(staged))
      this.host.mirror?.(sheetId, this.combinedMirror(sheetId))
      this.host.committed?.(sheetId, pivot, staged, committed.result.valueGen)
      return { ...staged, kind: 'committed', valueGen: committed.result.valueGen }
    } finally { if (entry.applyRequest === requestId) entry.applyRequest = undefined }
  }

  release(sheetId: number, pivot: number): void {
    const key = keyOf(sheetId, pivot), entry = this.entries.get(key)
    if (!entry) return
    this.entries.delete(key); entry.generation++
    if (entry.openRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.openRequest }).catch(() => {})
    if (entry.applyRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.applyRequest }).catch(() => {})
    const sheetOverlays = this.mirrored.get(sheetId)
    sheetOverlays?.delete(pivot)
    if (!sheetOverlays?.size) this.mirrored.delete(sheetId)
    this.host.mirror?.(sheetId, this.combinedMirror(sheetId))
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

  private combinedMirror(sheetId: number): ValueOverlay | undefined {
    const overlays = [...(this.mirrored.get(sheetId)?.values() ?? [])]
    if (!overlays.length) return undefined
    return Object.freeze({
      rects: overlays.flatMap(overlay => overlay.rects),
      protectedRects: overlays.flatMap(overlay => overlay.protectedRects ?? overlay.rects),
      merges: overlays.flatMap(overlay => overlay.merges),
      usedRows: Math.max(...overlays.map(overlay => overlay.usedRows)),
      usedCols: Math.max(...overlays.map(overlay => overlay.usedCols)),
      cell(row: number, col: number) {
        for (let i = overlays.length - 1; i >= 0; i--) {
          const overlay = overlays[i]!
          if (overlay.rects.some(rect => row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1)) return overlay.cell(row, col)
        }
        return undefined
      },
    })
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

/** Rebuild the callable overlay only on the main thread after its worker commit. */
export function valueOverlayFromPivotApply(result: RemotePivotApplyPayload): ValueOverlay {
  const values = new Map(result.cells.map(cell => [`${cell.row},${cell.col}`, cell.value]))
  return Object.freeze({
    rects: result.rects, protectedRects: result.protectedRects, merges: result.merges,
    usedRows: result.usedRows, usedCols: result.usedCols,
    cell: (row: number, col: number) => values.get(`${row},${col}`) ?? null,
  })
}
