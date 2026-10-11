/**
 * Resident viewer kernel: the worker-side command dispatcher (spec §10.2,
 * plan S4.1). It owns one DocumentSession per open docId, maps session
 * outcomes to typed replies and accounts every ArrayBuffer copy and transfer.
 *
 * No worker globals: worker-entry.ts is a thin adapter, and tests drive the
 * kernel in-process (directly, or through connectInProcess, which reproduces
 * structured-clone/transfer semantics including buffer detachment).
 */
import type {
  CacheIdentity, OperationRefusal, SourceHandle, ViewCommitDelta, WorkerCommand, WorkerCommandKind,
  WorkerCommandOwner, WorkerDiagnostic, WorkerOwner, WorkerReply,
} from '../xlsx/view/contracts'
import type { XlsxDocument } from '../xlsx/types'
import {
  displayBitsetsFromVisibility, isUint16Array, refusalReply, reply, transferablesOf, TransferAccounting,
  type CommandOf, type ReplyBody, type TransferLedger,
} from './viewerProtocol'
import {
  DocumentSession, type CommittedView, type OperationOutcome, type PrepareOutcome, type SessionOptions, type WorkbookSource,
} from './viewerSession'
import { createObjectWorkbookSource } from './viewerObjectSource'
import { MAX_ROWS } from '../xlsx/view/coordinates'

export type ReplySink = (reply: WorkerReply, transfer: ArrayBuffer[]) => void
export interface ViewerKernelEnv {
  /** Parses package bytes once per open; the only parse entry point. */
  parse(bytes: Uint8Array): Promise<XlsxDocument>
  /** Optional source override (S1 paged adapters, instrumented tests). */
  openSource?(handle: SourceHandle, docId: string, signal: AbortSignal): Promise<WorkbookSource>
  /** Instrumentation seam between dispatch and emission (tests pin the final
   * ownership recheck by closing a document exactly here). */
  afterDispatch?(command: WorkerCommand): void | Promise<void>
}
/** Recently finished request owners kept so a late cancel can still answer
 * with its target's owner (bounded; oldest dropped first). */
const FINISHED_OWNERS_LIMIT = 1024
export interface KernelStats { opens: number; parses: number; sourceReads: number }
interface InFlight { controller: AbortController; owner: WorkerOwner; kind: WorkerCommandKind }

const errorReply = <K extends WorkerCommandKind>(
  command: CommandOf<K>, owner: WorkerCommandOwner<K>, code: string, retryable: boolean, diagnostics: readonly WorkerDiagnostic[] = [],
): WorkerReply<K> => reply(command, owner, { kind: 'error', code, retryable } as unknown as ReplyBody<K>, diagnostics)

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
  readonly stats: KernelStats = { opens: 0, parses: 0, sourceReads: 0 }

  constructor(private readonly env: ViewerKernelEnv, private readonly options: SessionOptions = {}) {}

  session(docId: string): DocumentSession | undefined { return this.sessions.get(docId) }
  ledger(): TransferLedger { return this.accounting.snapshot() }

  /** Handles one command, emitting its replies through `post`. Never throws:
   * unexpected failures become an error reply for the command's owner. */
  async handle(command: WorkerCommand, post: ReplySink): Promise<void> {
    const emit = (out: WorkerReply) => {
      const transfer = transferablesOf(out)
      this.accounting.transfer(transfer)
      post(out, transfer)
    }
    let tracked = false
    try {
      if (command.kind === 'cancel') { emit(this.cancel(command)); return }
      if (this.inflight.has(command.requestId)) {
        emit(errorReply(command, capturedOwner(command) as never, 'invalid-cursor', false, [{ code: 'duplicate-request', message: 'requestId is already in flight' }]))
        return
      }
      const controller = new AbortController()
      this.inflight.set(command.requestId, { controller, owner: capturedOwner(command), kind: command.kind })
      tracked = true
      const out = await this.dispatch(command, controller.signal)
      await this.env.afterDispatch?.(command)
      emit(this.finalCheck(command, out, controller.signal))
    } catch (cause) {
      emit(errorReply(command, capturedOwner(command) as never, 'page-error', false, [{
        code: 'worker-exception', message: cause instanceof Error ? cause.message : String(cause),
      }]))
    } finally {
      if (tracked) {
        const entry = this.inflight.get(command.requestId)
        this.inflight.delete(command.requestId)
        if (entry) {
          this.finished.set(command.requestId, entry.owner)
          if (this.finished.size > FINISHED_OWNERS_LIMIT) this.finished.delete(this.finished.keys().next().value as string)
        }
      }
    }
  }
  /** Last ownership check, immediately before emission (§10.1): a completion is
   * never published for a cancelled request, a closed document or an owner a
   * newer commit superseded. Sort/filter completions already committed their
   * view; they are only withheld when the document itself is gone. */
  private finalCheck(command: Exclude<WorkerCommand, { kind: 'cancel' }>, out: WorkerReply, signal: AbortSignal): WorkerReply {
    if (out.kind !== 'complete') return out
    switch (command.kind) {
      case 'open':
        return this.sessions.has(command.docId) ? out : reply(command, { docId: command.docId }, { kind: 'cancelled' })
      case 'close':
        return out
      case 'sort':
      case 'filter': {
        const session = this.sessions.get(command.owner.docId)
        return session && !session.isClosed ? out : this.stale(command, command.owner)
      }
      case 'activateSheet': {
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        const session = this.sessions.get(command.owner.docId)
        return session && !session.isClosed ? out : errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document closed' }])
      }
      case 'prepareColumns':
      case 'valueQuery': {
        if (signal.aborted) return reply(command, command.owner, { kind: 'cancelled' })
        const session = this.sessions.get(command.owner.docId)
        return session && session.isCurrent(command.owner) ? out : this.stale(command, command.owner)
      }
      default:
        return out
    }
  }

  private async dispatch(command: Exclude<WorkerCommand, { kind: 'cancel' }>, signal: AbortSignal): Promise<WorkerReply> {
    switch (command.kind) {
      case 'open': return this.open(command, signal)
      case 'close': return this.close(command)
      case 'activateSheet': {
        // Activation selects resident state only; it publishes nothing that depends
        // on view generations, so the document owner is the relevant check.
        const session = this.sessions.get(command.owner.docId)
        if (!session || session.isClosed) return errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-document', message: 'Document is not open' }])
        if (!session.hasSheet(command.sheetId)) return errorReply(command, command.owner, 'stale-source', false, [{ code: 'unknown-sheet', message: `No sheet ${command.sheetId}` }])
        session.activate(command.sheetId)
        return reply(command, command.owner, { kind: 'complete', result: { kind: 'activateSheet', sheetId: command.sheetId } })
      }
      case 'prepareColumns': {
        const session = this.current(command.owner)
        if (!session) return this.stale(command, command.owner)
        if (!isUint16Array(command.cols)) {
          return errorReply(command, command.owner, 'invalid-coordinate', false, [{ code: 'invalid-columns', message: 'cols must be a Uint16Array of column indexes' }])
        }
        const outcome = await session.prepareColumns(command.owner, [...command.cols], signal)
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
        const outcome = command.kind === 'sort'
          ? await session.sort(command.owner, command.spec, signal)
          : await session.filter(command.owner, command.spec, command.nowMs, signal)
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
      case 'viewport':
      case 'search':
        // Owned by S5 (canonical viewport records and global find).
        return errorReply(command, command.owner as never, 'unsupported', false, [{ code: `${command.kind}-not-implemented`, message: `${command.kind} is delivered by S5` }])
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
    // The opening owner is tracked so close (or cancel) can abandon it: an open
    // that finishes after close releases its source and never installs a session.
    this.opening.add(command.docId)
    let source: WorkbookSource
    try {
      source = await this.openSource(command.source, command.docId, signal)
    } catch (cause) {
      if (signal.aborted) return reply(command, owner, { kind: 'cancelled' })
      throw cause
    } finally {
      this.opening.delete(command.docId)
    }
    if (signal.aborted) { await source.release(); return reply(command, owner, { kind: 'cancelled' }) }
    let session: DocumentSession
    try {
      session = new DocumentSession(command.docId, source, { ...this.options, profile: command.profile })
    } catch (cause) {
      // Ownership never transferred: release what open acquired.
      await source.release()
      throw cause
    }
    this.sessions.set(command.docId, session)
    this.stats.opens++
    return reply(command, owner, { kind: 'complete', result: { kind: 'open', valuePolicy: source.valuePolicy, sheets: source.sheets } })
  }
  /** Open is the only place source bytes are read or parsed. A blob is copied
   * into worker memory exactly once (accounted), parsed, then dropped. */
  private async openSource(handle: SourceHandle, docId: string, signal: AbortSignal): Promise<WorkbookSource> {
    if (this.env.openSource) return this.env.openSource(handle, docId, signal)
    switch (handle.kind) {
      case 'object': return createObjectWorkbookSource(handle.document, docId)
      case 'blob': {
        this.stats.sourceReads++
        const bytes = await handle.blob.arrayBuffer()
        this.accounting.copy(bytes.byteLength)
        this.stats.parses++
        return createObjectWorkbookSource(await this.env.parse(new Uint8Array(bytes)), docId)
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
      entry.controller.abort()
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
  /** Acknowledges cancellation of an in-flight request; the target finishes with
   * its own cancelled reply. Scheduling-level cancellation is S4.2. */
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
    target.controller.abort()
    return reply(command, target.owner, { kind: 'cancelled' })
  }
  private current(owner: CacheIdentity): DocumentSession | undefined {
    const session = this.sessions.get(owner.docId)
    return session && session.isCurrent(owner) ? session : undefined
  }
  private stale<K extends 'prepareColumns' | 'sort' | 'filter' | 'valueQuery'>(command: CommandOf<K>, owner: CacheIdentity): WorkerReply<K> {
    return errorReply(command, owner as WorkerCommandOwner<K>, 'stale-source', true, [{ code: 'stale-owner', message: 'Command owner is not the current committed view' }])
  }
  private failed<K extends 'prepareColumns' | 'sort' | 'filter'>(
    command: CommandOf<K>, owner: CacheIdentity,
    outcome: Exclude<OperationOutcome, { kind: 'committed' | 'refused' }> | Exclude<PrepareOutcome, { kind: 'complete' | 'resource-limit' }>,
  ): WorkerReply<K> {
    const own = owner as WorkerCommandOwner<K>
    switch (outcome.kind) {
      case 'stale': return this.stale(command, owner)
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
