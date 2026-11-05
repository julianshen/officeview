/** Main-thread two-phase client for worker-owned interactive pivot sessions. */
import type { PivotDelta } from '../xlsx/pivots/session/delta'
import type { PivotTableDefinition } from '../xlsx/pivots/model'
import type { PivotMenuModel } from '../xlsx/pivots/session/menu'
import type { ChartModel } from '../charts/model'
import { LINKED_CONTROL_REPLY_BYTES, LINKED_CONTROL_WORKER_GRANT_BYTES, MAX_LINKED_CONTROL_TARGETS, PIVOT_FIT_REPLY_BYTES, type LinkedControlPivotTarget, type LinkedControlSelection, type ValueOverlayCell, type WorkerCommand, type WorkerReply } from '../xlsx/view/contracts'
import type { SourceRect } from '../xlsx/view/coordinates'
import { ViewerNotSentError, type ViewerClient } from './viewerClient'
import { WORKER_BUDGETS } from '../xlsx/view/tileCache'

export interface RemotePivotOpen { readonly kind: 'open'; readonly revision: number; readonly def: PivotTableDefinition; /** The menus' model (Task 28). */ readonly menu: PivotMenuModel; /** The session's value cells a fit reads (Task 29 autofit; absent past the payload cap). */ readonly cells?: readonly { row: number; col: number; value: ValueOverlayCell | null }[] }
export type RemotePivotApply =
  | { readonly kind: 'committed'; readonly revision: number; readonly notice?: string; readonly atSaved?: true; readonly def: PivotTableDefinition; readonly menu: PivotMenuModel; readonly rects: readonly SourceRect[]; readonly protectedRects: readonly SourceRect[]; readonly overlayExtent: SourceRect | null; readonly merges: readonly SourceRect[]; readonly usedRows: number; readonly usedCols: number; readonly charts: readonly RemotePivotChart[]; readonly cells?: readonly { row: number; col: number; value: ValueOverlayCell | null }[]; readonly valueGen: number }
  | { readonly kind: 'stale' | 'cancelled' }
  | { readonly kind: 'refused'; readonly message: string }
export type RemoteLinkedControlApply =
  | { readonly kind: 'committed'; readonly revision: number; readonly selection: LinkedControlSelection; readonly targets: readonly { id: string; sheetId: number; pivot: number; result: RemotePivotApplyPayload; valueGen: number }[] }
  | { readonly kind: 'stale' | 'cancelled' }
  | { readonly kind: 'refused'; readonly message: string }

export interface RemotePivotApplyPayload {
  readonly revision: number
  /** The committed candidate is the pivot's saved view (Ruling S3-AK). */
  readonly atSaved?: true
  readonly def: PivotTableDefinition
  readonly menu: PivotMenuModel
  readonly rects: readonly SourceRect[]
  readonly protectedRects: readonly SourceRect[]
  readonly overlayExtent: SourceRect | null
  readonly merges: readonly SourceRect[]
  readonly usedRows: number
  readonly usedCols: number
  readonly charts: readonly RemotePivotChart[]
  readonly cells?: readonly { row: number; col: number; value: ValueOverlayCell | null }[]
  readonly hasFormulaCellsOutsidePivots?: boolean
}
export interface RemotePivotChart { readonly sheetId: string; readonly partPath: string; readonly treePath: string; readonly chart: ChartModel }
export interface RemotePivotHost {
  readonly docId: string
  readonly profile: 'D' | 'M'
  readonly client: ViewerClient
  isClosed(): boolean
  /** Invoked only after the worker acknowledges the matching commit. */
  /** Main-ledger receiving grants (Task 29 r2), as `RemoteWorkbook.request` grants its replies: `reserve` admits the
   * fit cells' grant before a command is posted (false: none, the reply carries no cells); `settle` charges the
   * arrived reply's stamped size from arrival, adopting the grant. Absent: no grant is sent (the worker's default). */
  readonly replyGrant?: { reserve(owner: string, bytes: number): boolean; settle(owner: string, reply: WorkerReply): void; release(owner: string): void }
  committed?(sheetId: number, pivot: number, result: RemotePivotApplyPayload, valueGen: number): void
  /** Publishes a linked control's complete fan-out in one main-thread state change. */
  linkedCommitted?(targets: readonly { sheetId: number; pivot: number; result: RemotePivotApplyPayload; valueGen: number }[]): void
  formulaFlagChanged?(sheetId: number, hasFormulaCellsOutsidePivots: boolean): void
  /** The pivot's committed display is withdrawn on this thread, once, when its release is sent. */
  released?(sheetId: number, pivot: number): void
  /** The worker acknowledged the release: its sheet's values moved to `valueGen` (final review minor 3: the value
   * generation only; the display was withdrawn when the release was sent). */
  releaseAcknowledged?(sheetId: number, valueGen: number): void
}
interface Entry { generation: number; revision: number; requestRevision: number; opened: boolean; def?: PivotTableDefinition; openRequest?: string; applyRequest?: string; open?: Promise<RemotePivotOpen | { kind: 'refused'; message: string }>
  /** Its last published commit shows a change (not the saved view): a sheet switch keeps it (Ruling S3-AK). */
  changed?: boolean }
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
  private linkedControlRevision = 0
  private linkedControlPending?: { revision: number; requestId: string; targetIds: readonly string[] }
  constructor(private readonly host: RemotePivotHost) {}

  /** Sends a pivot open or apply under a main-ledger grant for its fit cells (see `RemotePivotHost.replyGrant`). */
  private sendGranted<K extends 'pivotOpen' | 'pivotApply'>(command: Extract<import('../xlsx/view/contracts').WorkerCommand, { kind: 'pivotOpen' | 'pivotApply' }> & { kind: K }): Promise<WorkerReply<K>> {
    const grant = this.host.replyGrant
    const send = (sent: typeof command) => this.host.client.send(sent as never) as Promise<WorkerReply<K>>
    if (!grant) return send(command)
    const owner = `pivot-reply:${command.requestId}`
    const fitGrantBytes = grant.reserve(owner, PIVOT_FIT_REPLY_BYTES) ? PIVOT_FIT_REPLY_BYTES : 0
    return send({ ...command, fitGrantBytes }).then(
      reply => { grant.settle(owner, reply); return reply },
      // Refunded only when the command provably never left (as `RemoteWorkbook.request`): a disposed client may still
      // have the reply in flight, so its reservation stays until the reply arrives or the ledger is torn down.
      error => { if (error instanceof ViewerNotSentError) grant.release(owner); throw error })
  }

  /** Linked-control replies carry every candidate result, so they require a receiver grant (not only a worker grant). */
  private sendLinkedGranted(command: Extract<WorkerCommand, { kind: 'controlApply' | 'controlCommit' }>): Promise<WorkerReply<'controlApply' | 'controlCommit'>> | undefined {
    const grant = this.host.replyGrant
    if (!grant) return undefined
    const owner = `pivot-control-reply:${command.requestId}`
    if (!grant.reserve(owner, LINKED_CONTROL_REPLY_BYTES)) return undefined
    const sent = this.host.client.send(command as never) as Promise<WorkerReply<'controlApply' | 'controlCommit'>>
    return sent.then(reply => { grant.settle(owner, reply); return reply }, error => {
      // A disposed receiver can still have an in-flight reply; only refund a command known not to have left.
      if (error instanceof ViewerNotSentError) grant.release(owner)
      throw error
    })
  }

  open(sheetId: number, pivot: number, signal: AbortSignal): Promise<RemotePivotOpen | { kind: 'refused'; message: string }> {
    const key = keyOf(sheetId, pivot)
    let entry = this.entries.get(key)
    if (entry?.open) return entry.open
    if (!entry) this.entries.set(key, entry = { generation: 0, revision: 0, requestRevision: 0, opened: false })
    if (this.closed || this.host.isClosed()) return Promise.resolve({ kind: 'refused', message: 'The workbook is closed.' })
    const requestId = this.host.client.nextRequestId()
    const generation = entry.generation
    entry.openRequest = requestId
    const request = this.sendGranted<'pivotOpen'>({ kind: 'pivotOpen', requestId, docId: this.host.docId, sheetId: String(sheetId), pivot, grantBytes: WORKER_BUDGETS[this.host.profile] / 4 })
    entry.open = this.awaitRequest(request, requestId, signal).then(reply => {
      if (generation !== entry!.generation || this.entries.get(key) !== entry) return { kind: 'refused' as const, message: 'The pivot session was released while opening.' }
      if (reply.kind === 'cancelled') return { kind: 'refused' as const, message: messageOf(reply) }
      if (reply.kind !== 'complete' || reply.result.kind !== 'pivotOpen') return { kind: 'refused' as const, message: messageOf(reply) }
      entry!.revision = reply.result.revision
      entry!.opened = true
      entry!.def = reply.result.def
      return { kind: 'open' as const, revision: reply.result.revision, def: reply.result.def, menu: reply.result.menu, ...(reply.result.cells ? { cells: reply.result.cells } : {}) }
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
      const reply = await this.sendGranted<'pivotApply'>({ kind: 'pivotApply', requestId, docId: this.host.docId, sheetId: String(sheetId), pivot, baseRevision, revision: requestRevision, delta, grantBytes: WORKER_BUDGETS[this.host.profile] / 4 })
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
        entry.changed = !staged.atSaved
        staged = { ...staged, ...(committed.result.hasFormulaCellsOutsidePivots === undefined ? {} : { hasFormulaCellsOutsidePivots: committed.result.hasFormulaCellsOutsidePivots }) }
        this.host.committed?.(sheetId, pivot, staged, committed.result.valueGen)
        if (committed.result.hasFormulaCellsOutsidePivots === undefined) this.followFormulaFlag(sheetId)
      }
      return generation !== entry.generation || !registered ? { kind: 'stale' } : { ...staged, kind: 'committed', valueGen: committed.result.valueGen }
    } finally { if (entry.applyRequest === requestId) entry.applyRequest = undefined }
  }

  /** A workbook-scoped linked control is staged by one worker command and committed by exactly one command. */
  async applyLinked(selection: LinkedControlSelection, targets: readonly { sheetId: number; pivot: number; delta: PivotDelta | 'reset' }[], signal?: AbortSignal): Promise<RemoteLinkedControlApply> {
    if (this.closed || this.host.isClosed() || signal?.aborted) return { kind: 'cancelled' }
    if (!targets.length || targets.length > MAX_LINKED_CONTROL_TARGETS) return { kind: 'refused', message: 'Linked control target count exceeds the supported limit.' }
    const localKeys = targets.map(target => keyOf(target.sheetId, target.pivot))
    const wireKeys = targets.map(target => JSON.stringify([String(target.sheetId), target.pivot]))
    if (new Set(localKeys).size !== localKeys.length) return { kind: 'refused', message: 'Linked control contains duplicate pivot targets.' }
    const wireTargets: LinkedControlPivotTarget[] = []
    for (let index = 0; index < targets.length; index++) {
      const target = targets[index]!, entry = this.entries.get(localKeys[index]!)
      if (!entry?.opened || !entry.def || !entry.def.name) return { kind: 'refused', message: 'Open every linked pivot before applying a control.' }
      wireTargets.push({ sheetId: String(target.sheetId), pivot: target.pivot, name: entry.def.name, cacheId: entry.def.cacheId,
        baseRevision: entry.revision, requestRevision: ++entry.requestRevision, delta: target.delta })
    }
    if (this.linkedControlPending) {
      void this.host.client.send({ kind: 'controlCancel', requestId: this.host.client.nextRequestId(), docId: this.host.docId, revision: this.linkedControlPending.revision }).catch(() => {})
    }
    const revision = ++this.linkedControlRevision, requestId = this.host.client.nextRequestId()
    const pending = this.linkedControlPending = { revision, requestId, targetIds: wireKeys }
    const cancel = () => {
      void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: requestId }).catch(() => {})
      void this.host.client.send({ kind: 'controlCancel', requestId: this.host.client.nextRequestId(), docId: this.host.docId, revision }).catch(() => {})
    }
    let staged = false, committed = false
    signal?.addEventListener('abort', cancel, { once: true })
    try {
      const applyReply = this.sendLinkedGranted({ kind: 'controlApply', requestId, docId: this.host.docId, revision, selection, targets: wireTargets, grantBytes: LINKED_CONTROL_WORKER_GRANT_BYTES })
      if (!applyReply) return { kind: 'refused', message: 'Not enough main-thread memory to receive the linked-control result.' }
      const reply = await applyReply
      if (this.linkedControlPending !== pending) return { kind: 'stale' }
      if (signal?.aborted || reply.kind === 'cancelled') return { kind: 'cancelled' }
      if (reply.kind === 'error' && reply.code === 'stale-source') return { kind: 'stale' }
      if (reply.kind !== 'complete' || reply.result.kind !== 'controlApply') return { kind: 'refused', message: messageOf(reply) }
      staged = true
      if (reply.result.targets.length !== targets.length || reply.result.targets.some((target, index) => target.id !== wireKeys[index])) {
        return { kind: 'refused', message: 'Worker linked-control reply did not match the staged targets.' }
      }
      if (signal?.aborted) { cancel(); return { kind: 'cancelled' } }
      // Once sent, commit is a fact; wait for its acknowledgement even if the caller aborts meanwhile.
      const commitResponse = this.sendLinkedGranted({ kind: 'controlCommit', requestId: this.host.client.nextRequestId(), docId: this.host.docId, revision })
      if (!commitResponse) return { kind: 'refused', message: 'Not enough main-thread memory to receive the linked-control commit.' }
      const commitReply = await commitResponse
      if (commitReply.kind === 'error' && commitReply.code === 'stale-source') return { kind: 'stale' }
      if (commitReply.kind !== 'complete' || commitReply.result.kind !== 'controlCommit' || commitReply.result.revision !== revision) {
        return commitReply.kind === 'cancelled' ? { kind: 'cancelled' } : { kind: 'refused', message: messageOf(commitReply) }
      }
      committed = true
      const valueGens = new Map(commitReply.result.valueGens.map(item => [item.sheetId, item.valueGen]))
      const published = reply.result.targets.map((item, index) => {
        const target = targets[index]!, sheetId = String(target.sheetId), valueGen = valueGens.get(sheetId)
        if (valueGen === undefined || item.result.kind !== 'pivotApply') throw new Error('Linked-control commit omitted a target result or sheet generation')
        const result: RemotePivotApplyPayload = item.result
        const entry = this.entries.get(localKeys[index]!)
        if (entry) { entry.revision = item.revision; entry.changed = !result.atSaved }
        return { id: item.id, sheetId: target.sheetId, pivot: target.pivot, result, valueGen }
      })
      if (this.host.linkedCommitted) this.host.linkedCommitted(published.map(({ sheetId, pivot, result, valueGen }) => ({ sheetId, pivot, result, valueGen })))
      else for (const target of published) this.host.committed?.(target.sheetId, target.pivot, target.result, target.valueGen)
      // As on the ordinary path: a linked commit can leave formulas outside the
      // pivots stale, so follow the flag for every committed sheet.
      for (const sheetId of new Set(published.map(target => target.sheetId))) this.followFormulaFlag(sheetId)
      return { kind: 'committed', revision, selection: commitReply.result.selection, targets: published }
    } finally {
      signal?.removeEventListener('abort', cancel)
      if (staged && !committed) void this.host.client.send({ kind: 'controlCancel', requestId: this.host.client.nextRequestId(), docId: this.host.docId, revision }).catch(() => {})
      if (this.linkedControlPending === pending) this.linkedControlPending = undefined
    }
  }

  /** The workbook formula flag (S3-S) once the worker has reported it: it is document-wide and does not depend on the
   * commit or the pivot entry that asked (S3-AG), so it is kept for the document's life. */
  private formulaFlag?: boolean
  /** The one outstanding status request; every committed sheet waiting on it, each told when it settles. */
  private formulaStatus?: { sheets: Set<number> }
  /** A committed sheet's notice status (M-5): applied however the sheet's pivot entries change meanwhile (a release
   * right after the commit, a reopen); only the workbook's close drops it. */
  private followFormulaFlag(sheetId: number): void {
    if (this.formulaFlag !== undefined) { this.host.formulaFlagChanged?.(sheetId, this.formulaFlag); return }
    if (this.formulaStatus) { this.formulaStatus.sheets.add(sheetId); return }
    const pending = this.formulaStatus = { sheets: new Set([sheetId]) }
    void this.host.client.send({ kind: 'pivotFormulaStatus', requestId: this.host.client.nextRequestId(), docId: this.host.docId, sheetId: String(sheetId) }).then(reply => {
      if (this.formulaStatus === pending) this.formulaStatus = undefined
      if (this.closed || this.host.isClosed() || reply.kind !== 'complete' || reply.result.kind !== 'pivotFormulaStatus' || reply.result.hasFormulaCellsOutsidePivots === undefined) return
      const flag = this.formulaFlag = reply.result.hasFormulaCellsOutsidePivots
      for (const sheet of pending.sheets) this.host.formulaFlagChanged?.(sheet, flag)
    }, () => { if (this.formulaStatus === pending) this.formulaStatus = undefined })
  }

  release(sheetId: number, pivot: number): void {
    const key = keyOf(sheetId, pivot), entry = this.entries.get(key)
    if (!entry) return
    if (this.linkedControlPending?.targetIds.includes(JSON.stringify([String(sheetId), pivot]))) this.cancelLinkedControlPending()
    this.entries.delete(key); entry.generation++
    if (entry.openRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.openRequest }).catch(() => {})
    if (entry.applyRequest) void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: entry.applyRequest }).catch(() => {})
    this.host.released?.(sheetId, pivot)
    if (!this.closed && !this.host.isClosed()) void this.host.client.send({ kind: 'pivotRelease', requestId: this.host.client.nextRequestId(), docId: this.host.docId, sheetId: String(sheetId), pivot }).then(reply => {
      if (reply.kind === 'complete' && reply.result.kind === 'pivotRelease' && reply.result.valueGen !== undefined) this.host.releaseAcknowledged?.(sheetId, reply.result.valueGen)
    }).catch(() => {})
  }

  /** A sheet switch (Ruling S3-AK): the other sheets' pivots at their saved view are released; a pivot showing a
   * committed change keeps its session (the worker keeps its overlay, charged) until it is reset or the workbook
   * closes, as on the resident path, so the menus, the autofit's "before" and the fitted widths stay true. */
  releaseExcept(sheetId: number): void {
    if (this.linkedControlPending?.targetIds.some(id => (JSON.parse(id) as [string, number])[0] !== String(sheetId))) this.cancelLinkedControlPending()
    for (const [key, entry] of [...this.entries]) {
      const [candidate, pivot] = JSON.parse(key) as [number, number]
      if (candidate !== sheetId && !entry.changed) this.release(candidate, pivot)
    }
  }

  close(): void {
    if (this.closed) return
    this.cancelLinkedControlPending()
    for (const [key] of this.entries) {
      const [sheet, pivot] = JSON.parse(key) as [number, number]
      this.release(sheet, pivot)
    }
    this.closed = true
    this.entries.clear()
  }

  private cancelLinkedControlPending(): void {
    const pending = this.linkedControlPending
    if (!pending) return
    this.linkedControlPending = undefined
    void this.host.client.send({ kind: 'cancel', requestId: this.host.client.nextRequestId(), targetRequestId: pending.requestId }).catch(() => {})
    void this.host.client.send({ kind: 'controlCancel', requestId: this.host.client.nextRequestId(), docId: this.host.docId, revision: pending.revision }).catch(() => {})
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
