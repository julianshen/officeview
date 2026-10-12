/**
 * Resident viewer kernel: the worker-side command dispatcher (spec §10.2,
 * plans S4.1/S4.2). It owns one DocumentSession per open docId, maps session
 * outcomes to typed replies and accounts every ArrayBuffer copy and transfer.
 *
 * S4.2 publication rules: every request is bound at receipt to the session
 * instance (document epoch) current for its docId; a reply or progress is
 * published only while that same epoch is open and, for view-owned commands,
 * the captured owner is still current. Messages arrive FIFO, so binding at
 * receipt is exact. A new sort/filter whose owner is current supersedes older
 * in-flight drafts on the same sheet; a new value query replaces an older one
 * on the same column; close cancels everything of its document. Cancelled
 * replies name their reason in a diagnostic.
 *
 * Queries abort with a terminal, diagnosed cancellation after 120 seconds.
 * Paged reads, overview, viewport and workbook search use bounded resident
 * commands; continuations and projections stay in the worker.
 *
 * No worker globals: worker-entry.ts is a thin adapter, and tests drive the
 * kernel in-process (directly, or through connectInProcess, which reproduces
 * structured-clone/transfer semantics including buffer detachment).
 */
import type {
  CacheIdentity, FilterKeySet, OperationRefusal, ProfileId, SourceHandle, ViewCommitDelta, WorkerCommand, WorkerCommandKind,
  WorkerCommandOwner, WorkerDiagnostic, WorkerOwner, WorkerReply,
} from '../xlsx/view/contracts'
import type { XlsxDocument } from '../xlsx/types'
import {
  displayBitsetsFromVisibility, isUint16Array, refusalReply, reply, transferablesOf, TransferAccounting,
  type CommandOf, type ReplyBody, type TransferLedger,
} from './viewerProtocol'
import {
  DocumentSession, VIEW_OPERATION_STAGES, type CommittedView, type OperationOutcome, type PrepareOutcome, type SessionOptions,
  type WorkbookSource,
  workbookIndexKey,
} from './viewerSession'
import { createObjectWorkbookSource } from './viewerObjectSource'
import { createWorkerFilterEvaluator } from './viewerFilterEvaluator'
import { MAX_ROWS } from '../xlsx/view/coordinates'
import { awaitAbortable, bindQueryDeadline } from './viewerAbort'
import { queryPayloadBytes } from './viewerQueryBudget'
import type { PoolCharge } from '../xlsx/view/tileCache'

class KeySetQuota extends Error {
  constructor(readonly requiredBytes: number, readonly limitBytes: number) { super('Filter keys exceed the worker memory budget') }
}

export const QUERY_DEADLINE_MS = 120_000

/** A typed open refusal (admission or storage); the open replies with its code. */
export class OpenRefusal extends Error {
  constructor(readonly code: string, message: string) { super(message); this.name = 'OpenRefusal' }
}
const OPEN_FAILURE_CODES: ReadonlySet<string> = new Set(['resource-limit', 'unsupported', 'parse-error', 'storage-error', 'quota'])
export type ReplySink = (reply: WorkerReply, transfer: ArrayBuffer[]) => void
export interface ViewerKernelEnv {
  /** Parses package bytes once per open; the only parse entry point. */
  parse(bytes: Uint8Array): Promise<XlsxDocument>
  /** Optional source override (S1 paged adapters, instrumented tests). */
  openSource?(handle: SourceHandle, docId: string, signal: AbortSignal): Promise<WorkbookSource>
  /** Instrumentation seam between dispatch and emission (tests pin the final
   * ownership recheck by closing a document exactly here). */
  afterDispatch?(command: WorkerCommand): void | Promise<void>
  /** 'in-thread' is the no-worker fallback for small admitted documents: the
   * same chunked kernels, but paged (large-mode) sources are refused. */
  host?: 'worker' | 'in-thread'
  /** Blob opens run admission first (S1.2); false keeps the parse-only path. */
  admitBlobs?: boolean
  /** Paged-import storage and limit overrides (tests inject fakes). */
  pageStorage?: import('../xlsx/storage/sourcePages').PageStorageEnv
  importLimits?: import('../xlsx/storage/importSession').ImportLimitOverrides
}
/** Why an in-flight request was aborted; echoed as the cancelled diagnostic. */
export type CancelReason = 'cancel-requested' | 'superseded' | 'range-replaced' | 'document-closed' | 'query-timeout'
const CANCEL_MESSAGES: Readonly<Record<CancelReason, string>> = Object.freeze({
  'cancel-requested': 'Cancelled by request',
  superseded: 'Superseded by a newer view draft on the same sheet',
  'range-replaced': 'Replaced by a newer query on the same column',
  'document-closed': 'The document was closed',
  'query-timeout': 'Query timed out after 120 seconds. Retry or choose a smaller range.',
})
/** Recently finished request owners kept so a late cancel can still answer
 * with its target's owner (bounded; oldest dropped first). */
const FINISHED_OWNERS_LIMIT = 1024
export interface KernelStats { opens: number; parses: number; sourceReads: number }
export interface KernelLifecycleStats { sessions: number; inflight: number; opening: number; retainedBytes: number; transientBytes: number }
interface InFlight {
  controller: AbortController
  owner: WorkerOwner
  kind: WorkerCommandKind
  /** Session (document epoch) current for the command's docId at receipt. */
  session: DocumentSession | undefined
  /** Supersession scope: a newer current command with the same key aborts this. */
  key: string | undefined
  reason?: CancelReason
}
type Report = (completed: number, total: number) => void

const errorReply = <K extends WorkerCommandKind>(
  command: CommandOf<K>, owner: WorkerCommandOwner<K>, code: string, retryable: boolean, diagnostics: readonly WorkerDiagnostic[] = [],
): WorkerReply<K> => reply(command, owner, { kind: 'error', code, retryable } as unknown as ReplyBody<K>, diagnostics)

/** The document a non-open command addresses (undefined when malformed). */
function documentOf(command: Exclude<WorkerCommand, { kind: 'cancel' | 'open' }>): string | undefined {
  if (command.kind === 'close') return command.docId
  const owner = (command as { owner?: { docId?: unknown } }).owner
  return typeof owner?.docId === 'string' ? owner.docId : undefined
}
/** Supersession scope of a command whose owner is current at receipt: view
 * drafts per sheet, value queries per column (range replacement). A stale
 * command never supersedes a current one. */
function supersessionKey(command: Exclude<WorkerCommand, { kind: 'cancel' }>, session: DocumentSession | undefined): string | undefined {
  if (session && command.kind === 'search' && workbookIndexKey(command.owner) === workbookIndexKey(session.workbookOwner())) return JSON.stringify(['search', session.epoch, command.owner.docId])
  if (!session || (command.kind !== 'sort' && command.kind !== 'filter' && command.kind !== 'valueQuery')) return undefined
  if (!session.isCurrent(command.owner)) return undefined
  const scope = [session.epoch, command.owner.docId, command.owner.sheetId]
  return JSON.stringify(command.kind === 'valueQuery' ? ['query', ...scope, command.query?.col] : ['view', ...scope])
}
/** Owner captured from the command itself (§10.2); cancel uses its target's. */
function capturedOwner(command: WorkerCommand, inflight?: InFlight): WorkerOwner {
  switch (command.kind) {
    case 'open': case 'close': return { docId: command.docId }
    case 'cancel': return inflight?.owner ?? { docId: '' }
    default: return command.owner
  }
}

export class ViewerKernel {
  private readonly sessions = new Map<string, DocumentSession>()
  private readonly inflight = new Map<string, InFlight>()
  private readonly opening = new Set<string>()
  private readonly finished = new Map<string, WorkerOwner>()
  private readonly accounting = new TransferAccounting()
  /** Sources opened through the paged route: their views live here, so filter
   * rules are evaluated here too (S5.4); object sessions keep the S9 transport. */
  private readonly pagedSources = new WeakSet<WorkbookSource>()
  /** Value-list keys shipped with in-flight filter commands, per document and
   * reference-counted: overlapping commands (a superseded draft and its
   * replacement) may ship the same committed ref, and one finishing must not
   * remove the keys the other still evaluates. */
  private readonly keySets = new Map<string, Map<string, { set: FilterKeySet; holders: number; charge: PoolCharge }>>()
  private readonly foreground = new Set<AbortSignal>()
  private readonly foregroundWaiters = new Set<() => void>()
  private async waitForForeground(signal: AbortSignal): Promise<void> {
    while (this.foreground.size && !this.foreground.has(signal) && !signal.aborted) {
      await awaitAbortable(new Promise<void>(resolve => {
        const wake = () => { signal.removeEventListener('abort', wake); this.foregroundWaiters.delete(wake); resolve() }
        this.foregroundWaiters.add(wake)
        signal.addEventListener('abort', wake, { once: true })
      }), signal)
    }
  }
  private epochs = 0
  readonly stats: KernelStats = { opens: 0, parses: 0, sourceReads: 0 }

  constructor(private readonly env: ViewerKernelEnv, private readonly options: SessionOptions = {}) {}

  session(docId: string): DocumentSession | undefined { return this.sessions.get(docId) }
  ledger(): TransferLedger { return this.accounting.snapshot() }
  /** S4.3 lifecycle ledger: open sessions, pending requests and the pooled
   * bytes of every open document (a closed document contributes nothing). */
  lifecycleStats(): KernelLifecycleStats {
    let retainedBytes = 0, transientBytes = 0
    for (const session of this.sessions.values()) {
      const memory = session.memoryStats()
      retainedBytes += memory.retainedBytes
      transientBytes += memory.transientBytes
    }
    return { sessions: this.sessions.size, inflight: this.inflight.size, opening: this.opening.size, retainedBytes, transientBytes }
  }

  /** Handles one command, emitting its replies through `post`. Never throws:
   * unexpected failures become an error reply for the command's owner. */
  async handle(command: WorkerCommand, post: ReplySink): Promise<void> {
    const emit = (out: WorkerReply) => {
      const transfer = transferablesOf(out)
      this.accounting.transfer(transfer)
      post(out, transfer)
    }
    let entry: InFlight | undefined
    let deadline: ReturnType<typeof setTimeout> | undefined
    let report: ReturnType<ViewerKernel['reporter']> | undefined
    let dispatched: WorkerReply | undefined
    try {
      if (command.kind === 'cancel') { emit(this.cancel(command)); return }
      if (this.inflight.has(command.requestId)) {
        emit(errorReply(command, capturedOwner(command) as never, 'invalid-cursor', false, [{ code: 'duplicate-request', message: 'requestId is already in flight' }]))
        return
      }
      const docId = command.kind === 'open' ? undefined : documentOf(command)
      const session = docId === undefined ? undefined : this.sessions.get(docId)
      const key = supersessionKey(command, session)
      if (key !== undefined) {
        for (const other of this.inflight.values()) {
          if (other.key === key && (command.kind !== 'search' || command.cursor === undefined)) this.abort(other, command.kind === 'valueQuery' ? 'range-replaced' : 'superseded')
        }
      }
      entry = { controller: new AbortController(), owner: capturedOwner(command), kind: command.kind, session, key }
      this.inflight.set(command.requestId, entry)
      const expired = new AbortController()
      bindQueryDeadline(entry.controller.signal, expired.signal)
      if (command.kind === 'viewport' || command.kind === 'read' && command.query.op === 'cell') this.foreground.add(entry.controller.signal)
      if (command.kind !== 'open' && command.kind !== 'close') {
        const captured = entry
        deadline = setTimeout(() => { this.abort(captured, 'query-timeout'); expired.abort() }, QUERY_DEADLINE_MS)
      }
      report = this.reporter(command, entry, emit)
      const work = this.dispatch(command, entry.controller.signal, report)
      const out = command.kind === 'open' || command.kind === 'close' ? await work : await awaitAbortable(work, expired.signal)
      dispatched = out
      const after = this.env.afterDispatch?.(command)
      if (after) {
        try { await awaitAbortable(Promise.resolve(after), expired.signal) }
        catch (cause) {
          // A landed view commit is irreversible. Deliver its actual current
          // outcome even if its instrumentation/publication wait expires.
          if (!(expired.signal.aborted && out.kind === 'complete' && out.result.kind === 'view')) throw cause
        }
      }
      const final = this.finalCheck(command, out, entry)
      report.close()
      const source = docId ? this.sessions.get(docId)?.source : undefined
      const status = source?.valueStatus?.()
      const diagnosed = final.kind === 'complete' && status?.policy === 'cached' ? { ...final, diagnostics: [
        ...final.diagnostics.filter(d => d.code !== 'cached-value-policy'),
        ...(source?.openDiagnostics ?? []).filter(d => d.code !== 'cached-value-policy' && !final.diagnostics.some(existing => existing.code === d.code && existing.message === d.message)),
        { code: 'cached-value-policy', message: status.message! },
      ] } : final
      emit(this.withReason(diagnosed, entry))
    } catch (cause) {
      if (entry?.controller.signal.aborted) {
        emit(this.withReason(reply(command as never, entry.owner as never, { kind: 'cancelled' }), entry))
        return
      }
      if (cause instanceof KeySetQuota) {
        emit(reply(command as never, capturedOwner(command) as never, { kind: 'quota', resource: 'memory', requiredBytes: cause.requiredBytes, limitBytes: cause.limitBytes }))
        return
      }
      emit(errorReply(command, capturedOwner(command) as never, 'page-error', false, [{
        code: 'worker-exception', message: cause instanceof Error ? cause.message : String(cause),
      }]))
    } finally {
      clearTimeout(deadline)
      report?.close()
      if (entry?.controller.signal.aborted && dispatched?.kind === 'complete' && dispatched.result.kind === 'search' && dispatched.result.nextCursor) {
        entry.session?.releaseSearch(dispatched.result.nextCursor)
      }
      if (entry) {
        this.foreground.delete(entry.controller.signal)
        if (!this.foreground.size) for (const wake of [...this.foregroundWaiters]) wake()
        this.inflight.delete(command.requestId)
        this.finished.set(command.requestId, entry.owner)
        if (this.finished.size > FINISHED_OWNERS_LIMIT) this.finished.delete(this.finished.keys().next().value as string)
      }
    }
  }
  private abort(entry: InFlight, reason: CancelReason): void {
    if (entry.controller.signal.aborted) return
    entry.reason = reason
    entry.controller.abort(reason)
  }
  /** Progress publication guard (§10.4): only completed-batch counts, strictly
   * increasing, never after cancellation, after the terminal reply, outside the
   * bound epoch or once the captured view owner is superseded. */
  private reporter(command: Exclude<WorkerCommand, { kind: 'cancel' }>, entry: InFlight, emit: (out: WorkerReply) => void): Report & { close(): void } {
    let last = -1, open = true
    const report = ((completed: number, total: number) => {
      if (!open || entry.controller.signal.aborted || completed <= last || completed > total) return
      if (command.kind !== 'sort' && command.kind !== 'filter' && command.kind !== 'prepareColumns') return
      const session = entry.session
      if (!session || this.sessions.get(command.owner.docId) !== session || !session.isCurrent(command.owner)) return
      last = completed
      emit(reply(command, command.owner, { kind: 'progress', completed, total } as ReplyBody<typeof command.kind>))
    }) as Report & { close(): void }
    report.close = () => { open = false }
    return report
  }
  /** Cancelled replies carry the abort reason so the UI can say why. */
  private withReason(out: WorkerReply, entry: InFlight): WorkerReply {
    if (out.kind !== 'cancelled' || !entry.reason || out.diagnostics.length) return out
    return { ...out, diagnostics: [{ code: entry.reason, message: CANCEL_MESSAGES[entry.reason] }] }
  }
  /** Last ownership check, immediately before emission (§10.1): a completion is
   * never published for a cancelled request, a closed or replaced document epoch
   * or an owner a newer commit superseded. Sort/filter completions already
   * committed their view; they are withheld when their epoch is gone or a newer
   * commit replaced their delta (the client then learns the newer state). */
  private finalCheck(command: Exclude<WorkerCommand, { kind: 'cancel' }>, out: WorkerReply, entry: InFlight): WorkerReply {
    if (out.kind !== 'complete') return out
    const signal = entry.controller.signal
    const live = (docId: string) => entry.session !== undefined && !entry.session.isClosed && this.sessions.get(docId) === entry.session
    switch (command.kind) {
      case 'open':
        return this.sessions.has(command.docId) ? out : reply(command, { docId: command.docId }, { kind: 'cancelled' })
      case 'close':
        return out
      case 'sort':
      case 'filter': {
        // The command owner is superseded by its own commit; the committed
        // delta's owner must still be the sheet's current one, so a delta that a
        // newer commit replaced before emission is never published.
        const delta = out.result.kind === 'view' ? out.result.delta : undefined
        return live(command.owner.docId) && delta !== undefined && entry.session!.isCurrent(delta.owner)
          ? out : this.stale(command, command.owner)
      }
      case 'activateSheet': {
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        return live(command.owner.docId) ? out : errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document closed' }])
      }
      case 'prepareColumns':
      case 'viewport':
      case 'valueQuery': {
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        return live(command.owner.docId) && entry.session!.isCurrent(command.owner) ? out : this.stale(command, command.owner)
      }
      case 'read': {
        // A read finishing after cancel, close, replacement or a newer commit is withheld (§10.1).
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        if (command.query?.op === 'releaseScan' || command.query?.op === 'releaseSearch') return out
        return live(command.owner.docId) && entry.session!.isCurrent(command.owner)
          ? out : errorReply(command, command.owner, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }])
      }
      case 'search': {
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        return live(command.owner.docId) && workbookIndexKey(command.owner) === workbookIndexKey(entry.session!.workbookOwner())
          ? out : errorReply(command, command.owner, 'stale-source', true, [{ code: 'stale-owner', message: 'Workbook search owner changed' }])
      }
      default:
        return out
    }
  }

  private async dispatch(command: Exclude<WorkerCommand, { kind: 'cancel' }>, signal: AbortSignal, report: Report): Promise<WorkerReply> {
    switch (command.kind) {
      case 'open': return this.open(command, signal)
      case 'close': return this.close(command)
      case 'activateSheet': {
        // Activation selects resident state; it publishes nothing that depends on
        // view generations, so the document owner is the relevant check. An
        // evicted view (S4.3) is rebuilt under its unchanged identity first, and
        // the active sheet switches only once that succeeded: a refused rebuild
        // leaves the previous (resident, protected) sheet active.
        const session = this.sessions.get(command.owner.docId)
        if (!session || session.isClosed) return errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document is not open' }])
        if (!session.hasSheet(command.sheetId)) return errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-sheet', message: `No sheet ${command.sheetId}` }])
        const resident = await session.ensureView(command.sheetId, signal)
        switch (resident.kind) {
          case 'resident':
            if (session.isClosed) return errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document closed' }])
            session.activate(command.sheetId)
            return reply(command, command.owner, { kind: 'complete', result: { kind: 'activateSheet', sheetId: command.sheetId, layout: session.layout(command.sheetId) } }, resident.diagnostics ?? [])
          case 'cancelled': return reply(command, command.owner, { kind: 'cancelled' })
          case 'stale': return errorReply(command, command.owner, 'stale-source', true, [{ code: 'view-rebuild-stale', message: 'The sheet view changed while it was being rebuilt' }])
          case 'incomplete': return reply(command, command.owner, { kind: 'incomplete', missing: resident.missing })
          case 'quota': return reply(command, command.owner, { kind: 'quota', resource: 'memory', limitBytes: resident.limitBytes, requiredBytes: resident.requiredBytes })
          case 'error': return errorReply(command, command.owner, resident.error.code === 'operation-refused' ? 'page-error' : resident.error.code, resident.error.retryable, [{ code: 'view-rebuild-failed', message: resident.error.message }])
        }
      }
      case 'prepareColumns': {
        const session = this.current(command.owner)
        if (!session) return this.stale(command, command.owner)
        if (!isUint16Array(command.cols)) {
          return errorReply(command, command.owner, 'invalid-coordinate', false, [{ code: 'invalid-columns', message: 'cols must be a Uint16Array of column indexes' }])
        }
        // The zero-batch acknowledgment comes from the session once this batch is
        // admitted past the serialized preparation queue, never while queued.
        const outcome = await session.prepareColumns(command.owner, [...command.cols], signal, { progress: report })
        if (outcome.kind === 'resource-limit') return errorReply(command, command.owner, 'resource-limit', false, [{ code: 'prepare-batch-limit', message: outcome.message }])
        if (outcome.kind !== 'complete') return this.failed(command, command.owner, outcome)
        // Last check before emission: never complete for a cancelled request or
        // an owner that a newer commit (or close) superseded meanwhile.
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        if (!session.isCurrent(command.owner)) return this.stale(command, command.owner)
        // The command buffer arrived as the worker's own; reply with a fresh copy
        // so a resident list is never detached by the outgoing transfer.
        const cols = command.cols.slice()
        this.accounting.copy(cols.byteLength)
        return reply(command, command.owner, { kind: 'complete', result: { kind: 'prepareColumns', cols } })
      }
      case 'sort':
      case 'filter': {
        const session = this.current(command.owner)
        if (!session) return this.stale(command, command.owner)
        report(0, VIEW_OPERATION_STAGES)
        const hooks = { progress: report }
        let shipped: string[] = []
        const incoming = session.memory.charge()
        let outcome: OperationOutcome
        try {
          if (command.kind === 'filter' && command.keySets?.length) {
            const bytes = queryPayloadBytes(command.keySets, session.memory.limitBytes)
            if (!incoming.set(bytes)) throw new KeySetQuota(bytes, session.memory.limitBytes)
            shipped = this.installKeys(command.owner.docId, command.keySets)
          }
          outcome = command.kind === 'sort'
            ? await session.sort(command.owner, command.spec, signal, hooks)
            : await session.filter(command.owner, command.spec, command.nowMs, signal, hooks)
        } finally { this.removeKeys(command.owner.docId, shipped); incoming.release() }
        if (outcome.kind === 'committed') {
          return reply(command, command.owner, { kind: 'complete', result: { kind: 'view', delta: this.delta(outcome.view) } } as ReplyBody<typeof command.kind>)
        }
        if (outcome.kind === 'refused') return refusalReply(command, command.owner, outcome.refusal, refusalDiagnostics(outcome.refusal))
        return this.failed(command, command.owner, outcome)
      }
      case 'valueQuery': {
        const session = this.current(command.owner)
        if (!session) return this.stale(command, command.owner)
        const page = await session.valueQuery(command.owner, command.query, signal)
        if ('kind' in page) {
          return page.kind === 'cancelled' ? reply(command, command.owner, { kind: 'cancelled' }) : this.stale(command, command.owner)
        }
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        if (!session.isCurrent(command.owner)) return this.stale(command, command.owner)
        return reply(command, command.owner, { kind: 'complete', result: { kind: 'valueQuery', page } })
      }
      case 'read': {
        if (command.query?.op === 'releaseScan' || command.query?.op === 'releaseSearch') {
          // Cleanup is document-owned: it must succeed after a newer commit too.
          const session = this.sessions.get(command.owner.docId)
          const released = command.query.op === 'releaseScan' ? session?.releaseScan(command.query.cursor) ?? false : session?.releaseSearch(command.query.cursor) ?? false
          return reply(command, command.owner, { kind: 'complete', result: { kind: 'read', result: { op: command.query.op, released } } })
        }
        const session = this.current(command.owner)
        if (!session) return errorReply(command, command.owner, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }])
        const out = await session.read(command.owner, command.query, signal)
        if ('kind' in out) {
          return out.kind === 'cancelled' ? reply(command, command.owner, { kind: 'cancelled' })
            : errorReply(command, command.owner, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }])
        }
        return reply(command, command.owner, { kind: 'complete', result: { kind: 'read', result: out } })
      }
      case 'viewport': {
        const session = this.current(command.owner)
        if (!session) return this.stale(command, command.owner)
        const out = await session.viewport(command.owner, command.window, signal, command.scope)
        return 'result' in out ? reply(command, command.owner, { kind: 'complete', result: out.result }, out.diagnostics) : this.failed(command, command.owner, out)
      }
      case 'search': {
        const session = this.sessions.get(command.owner.docId)
        if (!session) return errorReply(command, command.owner, 'stale-source', false)
        const out = await session.search(command.owner, command.query, command.cursor, signal)
        if ('kind' in out) return this.failed(command, command.owner, out)
        return reply(command, command.owner, { kind: 'complete', result: { kind: 'search', items: out.hits, unavailableCount: out.unavailableCount, page: out, ...(out.cursor ? { nextCursor: out.cursor } : {}) } })
      }
      default: {
        // Unknown kinds (malformed or newer clients) get a typed error, never undefined.
        const unknown = command as unknown as { kind: unknown; requestId: string; owner?: WorkerOwner; docId?: unknown }
        const owner = unknown.owner ?? { docId: typeof unknown.docId === 'string' ? unknown.docId : '' }
        return errorReply(unknown as never, owner as never, 'unsupported', false, [{ code: 'unknown-command', message: `Unknown command kind ${String(unknown.kind)}` }])
      }
    }
  }

  private async open(command: CommandOf<'open'>, signal: AbortSignal): Promise<WorkerReply> {
    const owner = { docId: command.docId }
    if (typeof command.docId !== 'string' || !command.docId) return errorReply(command, owner, 'invalid-cursor', false, [{ code: 'invalid-document-id', message: 'docId must be a nonempty string' }])
    if (this.sessions.has(command.docId) || this.opening.has(command.docId)) {
      return errorReply(command, owner, 'unsupported', false, [{ code: 'duplicate-document', message: 'docId is already open or opening' }])
    }
    if (this.env.host === 'in-thread' && command.source?.kind === 'paged') {
      // No-worker mode admits small documents only; large paged sources need the
      // resident worker and its page storage (spec §§4.2, 10.2).
      return errorReply(command, owner, 'unsupported', false, [{ code: 'worker-required', message: 'Large paged sources require the resident viewer worker' }])
    }
    // The opening owner is tracked so close (or cancel) can abandon it: an open
    // that finishes after close releases its source and never installs a session.
    this.opening.add(command.docId)
    let source: WorkbookSource
    try {
      source = await this.openSource(command.source, command.docId, signal, command.profile)
    } catch (cause) {
      if (signal.aborted) return reply(command, owner, { kind: 'cancelled' })
      // Typed refusals (admission limits, storage, quota) reply with their code;
      // nothing is opened and no blank document is published.
      const code = (cause as { code?: unknown } | null)?.code
      if (cause instanceof OpenRefusal || (typeof code === 'string' && OPEN_FAILURE_CODES.has(code))) {
        const typed = String(code)
        return errorReply(command, owner, typed, typed === 'storage-error' || typed === 'quota', [{ code: 'open-refused', message: (cause as Error).message }])
      }
      throw cause
    } finally {
      this.opening.delete(command.docId)
    }
    if (signal.aborted) { await source.release(); return reply(command, owner, { kind: 'cancelled' }) }
    let session: DocumentSession
    const evaluator = !this.options.filterEvaluator && this.pagedSources.has(source)
      ? { filterEvaluator: createWorkerFilterEvaluator(stamp => this.keysFor(stamp.docId), source.formatting?.()) } : {}
    try {
      session = new DocumentSession(command.docId, source, { ...this.options, ...evaluator, waitForForeground: signal => this.waitForForeground(signal), profile: command.profile, epoch: ++this.epochs })
    } catch (cause) {
      // Ownership never transferred: release what open acquired.
      await source.release()
      throw cause
    }
    this.sessions.set(command.docId, session)
    this.stats.opens++
    const formatting = source.formatting?.()
    return reply(command, owner, { kind: 'complete', result: { kind: 'open', valuePolicy: source.valuePolicy, sheets: source.sheets, ...(formatting ? { formatting } : {}) } }, source.openDiagnostics ?? [])
  }
  /** Open is the only place source bytes are read or parsed. A blob is copied
   * into worker memory exactly once (accounted), parsed, then dropped. */
  private async openSource(handle: SourceHandle, docId: string, signal: AbortSignal, profile: ProfileId = 'D'): Promise<WorkbookSource> {
    if (this.env.openSource) return this.env.openSource(handle, docId, signal)
    switch (handle.kind) {
      case 'object': return createObjectWorkbookSource(handle.document, docId)
      case 'blob': {
        let openDiagnostics: WorkerDiagnostic[] = []
        if (this.env.admitBlobs !== false && this.env.host !== 'in-thread') {
          // Admission reads only the central directory and (for object
          // candidates) one bounded verification pass; a paged decision never
          // copies the blob into worker memory (S1.2).
          const { openLargeBlob } = await import('./viewerLargeSource')
          const routed = await openLargeBlob(handle.blob, docId, { profile, signal, limits: this.env.importLimits, storage: this.env.pageStorage })
          if (routed.kind === 'paged') { this.pagedSources.add(routed.source); return routed.source }
          if (routed.kind === 'refuse') throw new OpenRefusal(routed.code, routed.message)
          openDiagnostics = routed.diagnostics
        }
        this.stats.sourceReads++
        const bytes = await handle.blob.arrayBuffer()
        this.accounting.copy(bytes.byteLength)
        this.stats.parses++
        const source = createObjectWorkbookSource(await this.env.parse(new Uint8Array(bytes)), docId)
        return openDiagnostics.length ? { ...source, openDiagnostics: [...(source.openDiagnostics ?? []), ...openDiagnostics] } : source
      }
      default: throw new RangeError('Paged source handles are delivered by S1.2')
    }
  }
  private async close(command: CommandOf<'close'>): Promise<WorkerReply> {
    const owner = { docId: command.docId }
    // Abort first, including a pending open of this docId, before deciding.
    let abandonedOpen = false
    for (const [requestId, entry] of this.inflight) {
      if (requestId === command.requestId || entry.owner.docId !== command.docId) continue
      if (entry.kind === 'open') abandonedOpen = true
      this.abort(entry, 'document-closed')
    }
    const session = this.sessions.get(command.docId)
    if (!session) {
      return abandonedOpen
        ? reply(command, owner, { kind: 'complete', result: { kind: 'close' } })
        : errorReply(command, owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document is not open' }])
    }
    this.sessions.delete(command.docId)
    await session.release()
    return reply(command, owner, { kind: 'complete', result: { kind: 'close' } })
  }
  /** Acknowledges cancellation of an in-flight request immediately; the target
   * stops at its next cooperative check and finishes with its own cancelled
   * reply (reason `cancel-requested`), or with its commit if it already landed. */
  private cancel(command: CommandOf<'cancel'>): WorkerReply {
    const target = this.inflight.get(command.targetRequestId)
    if (!target) {
      // Defined outcome for a target that already finished (its reply may still
      // be queued): an error carrying the target's owner. A target this worker
      // never saw has no owner; the reply then uses the document-less owner.
      const done = this.finished.get(command.targetRequestId)
      return done
        ? errorReply(command, done, 'stale-source', false, [{ code: 'target-finished', message: 'The target request already finished' }])
        : errorReply(command, { docId: '' }, 'stale-source', false, [{ code: 'unknown-request', message: 'No such request' }])
    }
    this.abort(target, 'cancel-requested')
    return reply(command, target.owner, { kind: 'cancelled' })
  }
  /** Installs well-formed shipped keys for one filter command; returns their ids. */
  private installKeys(docId: string, keySets: readonly FilterKeySet[] | undefined): string[] {
    if (!Array.isArray(keySets) || !keySets.length) return []
    let map = this.keySets.get(docId)
    if (!map) this.keySets.set(docId, map = new Map())
    const installed: string[] = []
    try {
      for (const set of keySets) {
        const run = set?.run
        if (typeof set?.refId !== 'string' || typeof set.scope !== 'string' || !run || !Number.isSafeInteger(run.count) || run.count < 0
          || !isTyped(run.kinds, 'Uint8Array') || !isTyped(run.numbers, 'Float64Array') || !Array.isArray(run.texts)
          || run.kinds.length < run.count || run.numbers.length < run.count || run.texts.length < run.count) continue
        if (installed.includes(set.refId)) continue
        const held = map.get(set.refId)
        // A committed ref is immutable: a second holder shares the installed keys.
        if (held) held.holders++
        else {
          const session = this.sessions.get(docId)!
          const charge = session.memory.charge(), bytes = queryPayloadBytes(set, session.memory.limitBytes)
          if (!charge.set(bytes)) { charge.release(); throw new KeySetQuota(bytes, session.memory.limitBytes) }
          map.set(set.refId, { set, holders: 1, charge })
        }
        installed.push(set.refId)
      }
      return installed
    } catch (cause) { this.removeKeys(docId, installed); throw cause }
  }
  private removeKeys(docId: string, refIds: readonly string[]): void {
    const map = this.keySets.get(docId)
    if (!map) return
    for (const refId of refIds) {
      const held = map.get(refId)
      if (held && --held.holders <= 0) { held.charge.release(); map.delete(refId) }
    }
    if (!map.size) this.keySets.delete(docId)
  }
  private keysFor(docId: string): ReadonlyMap<string, FilterKeySet> {
    const map = this.keySets.get(docId)
    return new Map([...map?.entries() ?? []].map(([refId, held]) => [refId, held.set]))
  }
  private current(owner: CacheIdentity): DocumentSession | undefined {
    const session = this.sessions.get(owner.docId)
    return session && session.isCurrent(owner) ? session : undefined
  }
  private stale<K extends 'prepareColumns' | 'sort' | 'filter' | 'valueQuery' | 'viewport'>(command: CommandOf<K>, owner: CacheIdentity): WorkerReply<K> {
    return errorReply(command, owner as WorkerCommandOwner<K>, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }])
  }
  private failed<K extends 'prepareColumns' | 'sort' | 'filter' | 'viewport' | 'search'>(
    command: CommandOf<K>, owner: WorkerCommandOwner<K>,
    outcome: Exclude<OperationOutcome, { kind: 'committed' | 'refused' }> | Exclude<PrepareOutcome, { kind: 'complete' | 'resource-limit' }>,
  ): WorkerReply<K> {
    const own = owner as WorkerCommandOwner<K>
    switch (outcome.kind) {
      case 'stale': return errorReply(command, own, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not current' }])
      case 'cancelled': return reply(command, own, { kind: 'cancelled' } as unknown as ReplyBody<K>)
      case 'incomplete':
        return reply(command, own, { kind: 'incomplete', missing: outcome.missing } as unknown as ReplyBody<K>,
          'diagnostics' in outcome ? outcome.diagnostics ?? [] : [])
      case 'quota': return reply(command, own, {
        kind: 'quota', resource: 'memory', limitBytes: outcome.limitBytes, requiredBytes: outcome.requiredBytes,
      } as unknown as ReplyBody<K>, 'diagnostics' in outcome ? outcome.diagnostics ?? [] : [])
      case 'error': return errorReply(command, own, outcome.error.code === 'operation-refused' ? 'page-error' : outcome.error.code, outcome.error.retryable, [{ code: 'source-error', message: outcome.error.message }])
    }
  }
  /** Deep copy of plain wire data: every typed array becomes an accounted copy,
   * so a later transfer can never detach resident state (e.g. selection rows). */
  private wireCopy<T>(value: T): T {
    if (ArrayBuffer.isView(value)) {
      const copy = (value as unknown as Uint8Array).slice()
      this.accounting.copy(copy.byteLength)
      return copy as unknown as T
    }
    if (Array.isArray(value)) return value.map(entry => this.wireCopy(entry)) as unknown as T
    if (typeof value === 'object' && value !== null) {
      const out: Record<string, unknown> = {}
      for (const [key, entry] of Object.entries(value)) out[key] = this.wireCopy(entry)
      return out as T
    }
    return value
  }
  /** Compact transferable delta. Resident order/mask buffers stay with the
   * session; the delta carries accounted copies that are then transferred.
   * Wire conventions: `metrics` are SOURCE-space dimension overlays (immutable
   * source dimensions; the client projects them through `bands`, as
   * createSheetMetrics does); `visibility` lists display slots outside every
   * bitset as visible; `counts.total` is the logical row axis (1,048,576), not a
   * filter range, and `counts.visible` its visible display slots. S9.1 record
   * counts ("N of M records") need a body/filter-range total, which the frozen
   * ViewCommitDelta has no field for: that is a reviewed additive contract
   * change for S9.1, deliberately not made here. */
  private delta(view: CommittedView): ViewCommitDelta {
    const bands = view.order.bands.map(band => {
      const sourceAtOffset = band.sourceAtOffset.slice()
      const displayOffsetOfSource = band.displayOffsetOfSource.slice()
      this.accounting.copy(sourceAtOffset.byteLength)
      this.accounting.copy(displayOffsetOfSource.byteLength)
      return { r0: band.r0, r1: band.r1, sourceAtOffset, displayOffsetOfSource }
    })
    return {
      bands,
      visibility: displayBitsetsFromVisibility(view.visibility, this.accounting),
      metrics: { rows: view.geometry.rowOverrides.map(i => ({ ...i })), cols: view.geometry.colOverrides.map(i => ({ ...i })) },
      counts: { visible: view.visibility.rank(view.visibility.logicalCount), total: MAX_ROWS },
      selection: this.wireCopy(view.selection),
      owner: view.identity,
      workbookOwner: view.workbookOwner,
    }
  }
}
/** Realm-independent typed-array tag check (commands are structured clones). */
const isTyped = (value: unknown, tag: 'Uint8Array' | 'Float64Array'): boolean => Object.prototype.toString.call(value) === `[object ${tag}]`
function refusalDiagnostics(refusal: OperationRefusal): WorkerDiagnostic[] {
  const message = 'message' in refusal ? refusal.message : `Operation refused: ${refusal.reason}`
  return [{ code: `refused-${refusal.reason}`, message }]
}

/** In-process port with real worker message semantics: commands and replies are
 * structured-cloned and their transfer lists detach the sender's buffers. */
export interface InProcessPort {
  postMessage(message: unknown, transfer: Transferable[]): void
  addEventListener(type: 'message', listener: (event: { data: unknown }) => void): void
  removeEventListener(type: 'message', listener: (event: { data: unknown }) => void): void
  /** Resolves once every command posted so far has finished. */
  idle(): Promise<void>
}
export function connectInProcess(kernel: ViewerKernel): InProcessPort {
  const listeners = new Set<(event: { data: unknown }) => void>()
  const pending = new Set<Promise<void>>()
  return {
    postMessage(message, transfer) {
      const command = structuredClone(message, { transfer: transfer as ArrayBuffer[] }) as WorkerCommand
      const task = kernel.handle(command, (out, replyTransfer) => {
        const data = structuredClone(out, { transfer: replyTransfer })
        for (const listener of listeners) listener({ data })
      })
      pending.add(task)
      void task.finally(() => pending.delete(task))
    },
    addEventListener(_type, listener) { listeners.add(listener) },
    removeEventListener(_type, listener) { listeners.delete(listener) },
    async idle() { while (pending.size) await Promise.all([...pending]) },
  }
}
/** No-worker fallback for small admitted documents: the same resident kernel
 * and chunked cooperative kernels on the calling thread, behind the same
 * message semantics as a worker. Paged sources are refused (worker-required). */
export function createInThreadPort(env: Omit<ViewerKernelEnv, 'host'>, options: SessionOptions = {}): InProcessPort {
  return connectInProcess(new ViewerKernel({ ...env, host: 'in-thread' }, options))
}
