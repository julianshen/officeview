/**
 * Main-thread client for the resident viewer worker (spec §§10.2–10.4, plans
 * S4.1/S4.2).
 *
 * Correlates replies to commands by requestId, surfaces progress, ignores late
 * or foreign messages (including legacy render responses on the same worker)
 * and accounts every ArrayBuffer it transfers or leaves to structured clone.
 * ViewerPublications is the main-thread half of atomic publication: the one
 * place committed order/counts/status/search geometry change, guarded by the
 * document epoch and monotonic owners. selectViewerHost defines no-worker
 * support. Main-safe: imports only the protocol module.
 */
import type {
  CacheIdentity, ViewCommitDelta, WorkbookOwner, WorkerCommand, WorkerCommandKind, WorkerOwner, WorkerReply,
} from '../xlsx/view/contracts'
import {
  initialIdentity, initialWorkbookOwner, isArrayBuffer, isTerminalReply, isViewerReply, TransferAccounting,
  type CommandOf, type TransferLedger,
} from './viewerProtocol'

type MessageListener = (event: { data: unknown }) => void
/** Worker, MessagePort or the in-process port. */
export interface ViewerPort {
  postMessage(message: unknown, transfer: Transferable[]): void
  addEventListener?(type: 'message', listener: MessageListener): void
  removeEventListener?(type: 'message', listener: MessageListener): void
  onmessage?: ((event: MessageEvent) => void) | null
}
interface Pending {
  kind: WorkerCommandKind
  owner: WorkerOwner | undefined
  resolve: (reply: WorkerReply) => void
  reject: (error: Error) => void
  /** Last progress delivered; progress must strictly increase. */
  completed: number
}
/** Completed authoritative work reported by the worker for one request. */
export interface ViewerProgress { completed: number; total?: number }
export type ProgressListener = (progress: ViewerProgress) => void

export class ViewerClient {
  private readonly pending = new Map<string, Pending>()
  private readonly listeners = new Map<string, Set<ProgressListener>>()
  private readonly accounting = new TransferAccounting()
  private counter = 0
  private ignored = 0
  private rejected = 0
  private readonly detach: () => void

  constructor(private readonly port: ViewerPort, private readonly prefix = 'viewer') {
    const listener: MessageListener = event => this.receive(event.data)
    if (typeof port.addEventListener === 'function') {
      port.addEventListener('message', listener)
      this.detach = () => port.removeEventListener?.('message', listener)
    } else {
      // Chain like the legacy renderPageBitmap client so both share one worker.
      const previous = port.onmessage
      port.onmessage = (event: MessageEvent) => {
        if (typeof previous === 'function') previous.call(port, event)
        listener(event)
      }
      this.detach = () => { port.onmessage = previous ?? null }
    }
  }
  nextRequestId(): string { return `${this.prefix}-${++this.counter}` }
  ledger(): TransferLedger { return this.accounting.snapshot() }
  /** Replies with no pending request (late, duplicate or foreign). */
  get ignoredReplies(): number { return this.ignored }
  get pendingCount(): number { return this.pending.size }
  /** Replies whose commandKind or owner did not match their pending request. */
  get rejectedReplies(): number { return this.rejected }

  /** Resolves with the request's terminal reply. Typed arrays that own their whole
   * buffer are transferred (the worker takes ownership; the caller's view detaches).
   * Everything else is structured-cloned, which copies each unique backing buffer
   * in FULL, so those are accounted by full byteLength once per buffer. An open
   * command's object source document is cloned as one estimated copy. */
  send<K extends WorkerCommandKind>(command: CommandOf<K>): Promise<WorkerReply<K>> {
    if (this.pending.has(command.requestId)) return Promise.reject(new Error(`Duplicate requestId ${command.requestId}`))
    const transfer: ArrayBuffer[] = []
    const message = command as WorkerCommand
    if (message.kind === 'open' && message.source.kind === 'object') {
      this.accounting.copy(estimateCloneBytes(message.source.document), true)
    } else {
      const cloned = new Set<ArrayBuffer>()
      const seen = new Set<object>()
      const visit = (value: unknown): void => {
        if (typeof value !== 'object' || value === null || seen.has(value)) return
        seen.add(value)
        if (ArrayBuffer.isView(value)) {
          const buffer = value.buffer
          if (!isArrayBuffer(buffer)) return
          if (value.byteOffset === 0 && value.byteLength === buffer.byteLength) {
            if (!transfer.includes(buffer)) transfer.push(buffer)
          } else cloned.add(buffer)
          return
        }
        if (isArrayBuffer(value)) { cloned.add(value); return }
        if (value instanceof Map) { for (const entry of value.values()) visit(entry); return }
        // A Blob is passed by handle; its bytes are copied once inside the worker.
        if (typeof Blob !== 'undefined' && value instanceof Blob) return
        for (const entry of Array.isArray(value) ? value : Object.values(value)) visit(entry)
      }
      visit(command)
      // A buffer both transferred and partially viewed moves whole; no clone.
      for (const buffer of cloned) if (!transfer.includes(buffer)) this.accounting.copy(buffer.byteLength)
    }
    return new Promise<WorkerReply<K>>((resolve, reject) => {
      this.pending.set(command.requestId, {
        kind: command.kind, owner: this.expectedOwner(command),
        resolve: resolve as (reply: WorkerReply) => void, reject, completed: -1,
      })
      try {
        this.accounting.transfer(transfer)
        this.port.postMessage(command, transfer)
      } catch (error) {
        this.pending.delete(command.requestId)
        reject(error instanceof Error ? error : new Error(String(error)))
      }
    })
  }
  /** The owner the worker must echo (§10.2); undefined when it cannot be known
   * (cancel of a request this client is not tracking). */
  private expectedOwner(command: CommandOf<WorkerCommandKind>): WorkerOwner | undefined {
    switch (command.kind) {
      case 'open': case 'close': return { docId: command.docId }
      case 'cancel': return this.pending.get(command.targetRequestId)?.owner
      default: return command.owner
    }
  }
  private receive(data: unknown): void {
    if (!isViewerReply(data)) return
    const entry = this.pending.get(data.requestId)
    if (!entry) { this.ignored++; return }
    // Correlate on kind and owner too: a reply for another command or owner
    // that reuses this requestId must never settle the request.
    if (data.commandKind !== entry.kind || (entry.owner !== undefined && !sameOwner(data.owner, entry.owner)
      && !isDocumentlessCancelError(data))) {
      this.rejected++
      return
    }
    // Progress never settles a request; it is delivered while the request is
    // pending, strictly increasing (a reordered or repeated count is dropped).
    if (!isTerminalReply(data)) {
      if (data.kind !== 'progress' || data.completed <= entry.completed) return
      entry.completed = data.completed
      const progress: ViewerProgress = data.total === undefined ? { completed: data.completed } : { completed: data.completed, total: data.total }
      for (const listener of [...(this.listeners.get(data.requestId) ?? [])]) listener(progress)
      return
    }
    this.pending.delete(data.requestId)
    this.listeners.delete(data.requestId)
    entry.resolve(data)
  }
  /** Subscribes to completed-batch progress of one request (subscribe before
   * send to see the immediate acknowledgment). Listeners are dropped when the
   * request settles; the returned function unsubscribes earlier. */
  onProgress(requestId: string, listener: ProgressListener): () => void {
    let set = this.listeners.get(requestId)
    if (!set) this.listeners.set(requestId, set = new Set())
    set.add(listener)
    return () => {
      const current = this.listeners.get(requestId)
      current?.delete(listener)
      if (current && !current.size) this.listeners.delete(requestId)
    }
  }
  /** Detaches from the port and rejects every pending request. */
  dispose(): void {
    this.detach()
    for (const entry of this.pending.values()) entry.reject(new Error('Viewer client disposed'))
    this.pending.clear()
    this.listeners.clear()
  }
}
export function createViewerClient(port: ViewerPort, prefix?: string): ViewerClient {
  return new ViewerClient(port, prefix)
}

/** The defined cancel outcome for a target the worker never saw: an error whose
 * owner is the document-less `{ docId: '' }` (a cancel names no document). */
function isDocumentlessCancelError(reply: WorkerReply): boolean {
  const owner = reply.owner as { docId?: unknown }
  return reply.commandKind === 'cancel' && reply.kind === 'error' && owner.docId === '' && Object.keys(owner).length === 1
}
/** Structural equality over plain owner data (no Maps or class instances). */
function sameOwner(a: unknown, b: unknown): boolean {
  if (a === b) return true
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false
  if (Array.isArray(a) !== Array.isArray(b)) return false
  const keysA = Object.keys(a), keysB = Object.keys(b)
  if (keysA.length !== keysB.length) return false
  return keysA.every(key => Object.prototype.hasOwnProperty.call(b, key)
    && sameOwner((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key]))
}
/** Approximate structured-clone size of a plain object graph: UTF-16 strings,
 * 8-byte numbers, small object/array headers and each unique buffer in full. */
function estimateCloneBytes(root: unknown): number {
  let bytes = 0
  const seen = new Set<unknown>()
  const stack: unknown[] = [root]
  while (stack.length) {
    const value = stack.pop()
    if (typeof value === 'string') { bytes += value.length * 2; continue }
    if (typeof value === 'number') { bytes += 8; continue }
    if (typeof value === 'boolean') { bytes += 4; continue }
    if (typeof value !== 'object' || value === null || seen.has(value)) continue
    seen.add(value)
    if (ArrayBuffer.isView(value)) {
      if (!seen.has(value.buffer)) { seen.add(value.buffer); bytes += value.buffer.byteLength }
      continue
    }
    if (isArrayBuffer(value)) { bytes += value.byteLength; continue }
    bytes += 16
    if (value instanceof Map) { for (const [key, entry] of value) stack.push(key, entry); continue }
    for (const [key, entry] of Object.entries(value)) { bytes += key.length * 2; stack.push(entry) }
  }
  return bytes
}

/** One request's claim on a sheet publication, bound to the document epoch. */
export interface PublicationTicket {
  readonly docId: string
  readonly sheetId: string
  readonly epoch: number
  readonly requestId: string
}
export type SheetStatus =
  | { readonly kind: 'ready' }
  | { readonly kind: 'pending'; readonly requestId: string; readonly progress: ViewerProgress | null }
/** One status notice for a request that ended without a commit (§§6.3, 10.4). */
export interface SheetNotice {
  readonly requestId: string
  readonly kind: 'cancelled' | 'refused' | 'stale' | 'incomplete' | 'quota' | 'error'
  readonly code?: string
}
/** Geometry-dependent search/layout records, owned by an exact view and font. */
export interface SheetGeometry {
  readonly owner: CacheIdentity
  readonly fontGen: number
  readonly payload: unknown
}
/** The committed main-thread view of one sheet; replaced as a whole. */
export interface SheetPublication {
  readonly epoch: number
  readonly owner: CacheIdentity
  readonly workbookOwner: WorkbookOwner
  readonly bands: ViewCommitDelta['bands']
  readonly visibility: ViewCommitDelta['visibility']
  /** Null until the first committed view delta reports them. */
  readonly counts: ViewCommitDelta['counts'] | null
  readonly status: SheetStatus
  readonly notice: SheetNotice | null
  readonly geometry: SheetGeometry | null
}
interface DocumentPublications { epoch: number; fontGen: number; sheets: Map<string, SheetPublication> }
const sameIdentityValue = (a: CacheIdentity, b: CacheIdentity): boolean =>
  a.docId === b.docId && a.sheetId === b.sheetId && a.valueGen === b.valueGen && a.fontGen === b.fontGen && a.viewGen === b.viewGen

/**
 * Main-thread atomic publication (§§10.1, 10.4). Input is acknowledged at once
 * (begin → pending) while the last committed view stays visible. A completion
 * replaces order/mask/counts as ONE record swap, and only when its ticket
 * belongs to the current document epoch and its owner is strictly newer than
 * the committed one; delayed, reordered, superseded or previous-epoch replies
 * never replace order, counts, status or geometry. A request that ends without
 * a commit settles status (with one notice) only if it is still the sheet's
 * current pending request.
 */
export class ViewerPublications {
  private readonly documents = new Map<string, DocumentPublications>()
  /** Epoch counters survive close, so a reopened docId never reuses one. */
  private readonly epochs = new Map<string, number>()

  /** Starts a new document epoch with generation-0 owners (protocol convention). */
  open(docId: string, sheets: readonly { sheetId: string }[]): number {
    const epoch = (this.epochs.get(docId) ?? 0) + 1
    this.epochs.set(docId, epoch)
    const workbookOwner = initialWorkbookOwner(docId, sheets)
    const records = new Map<string, SheetPublication>()
    for (const { sheetId } of sheets) {
      records.set(sheetId, Object.freeze({
        epoch, owner: initialIdentity(docId, sheetId), workbookOwner, bands: [], visibility: [], counts: null,
        status: READY, notice: null, geometry: null,
      }))
    }
    this.documents.set(docId, { epoch, fontGen: 0, sheets: records })
    return epoch
  }
  close(docId: string): void { this.documents.delete(docId) }
  epoch(docId: string): number { return this.documents.get(docId)?.epoch ?? 0 }
  fontGen(docId: string): number { return this.documents.get(docId)?.fontGen ?? 0 }
  get(docId: string, sheetId: string): SheetPublication | undefined {
    return this.documents.get(docId)?.sheets.get(sheetId)
  }
  /** Immediate input acknowledgment: the sheet shows pending for this request
   * (superseding any earlier pending request) and keeps its committed view. */
  begin(docId: string, sheetId: string, requestId: string): PublicationTicket {
    const document = this.documents.get(docId)
    const current = document?.sheets.get(sheetId)
    if (!document || !current) throw new RangeError(`No open sheet ${docId}/${sheetId}`)
    this.replace(document, sheetId, { ...current, status: Object.freeze({ kind: 'pending', requestId, progress: null }) })
    return Object.freeze({ docId, sheetId, epoch: document.epoch, requestId })
  }
  /** Records progress only for the sheet's current pending request. */
  progress(ticket: PublicationTicket, progress: ViewerProgress): boolean {
    const found = this.pendingFor(ticket)
    if (!found) return false
    const previous = found.current.status.kind === 'pending' ? found.current.status.progress : null
    if (previous && progress.completed <= previous.completed) return false
    this.replace(found.document, ticket.sheetId, {
      ...found.current, status: Object.freeze({ kind: 'pending', requestId: ticket.requestId, progress: Object.freeze({ ...progress }) }),
    })
    return true
  }
  /** Applies a terminal sort/filter reply. 'applied': a newer committed view
   * replaced the record; 'status': a non-commit ended the current pending
   * request; 'ignored': stale, superseded or from another epoch. */
  settle(ticket: PublicationTicket, out: WorkerReply): 'applied' | 'status' | 'ignored' {
    const document = this.documents.get(ticket.docId)
    const current = document?.sheets.get(ticket.sheetId)
    if (!document || !current || document.epoch !== ticket.epoch || out.requestId !== ticket.requestId) return 'ignored'
    const isCurrentPending = current.status.kind === 'pending' && current.status.requestId === ticket.requestId
    if (out.kind === 'complete') {
      if (out.result.kind !== 'view') return 'ignored'
      const delta = out.result.delta
      const newer = delta.owner.docId === ticket.docId && delta.owner.sheetId === ticket.sheetId
        && delta.owner.viewGen > current.owner.viewGen && delta.workbookOwner.globalViewGen > current.workbookOwner.globalViewGen
      if (!newer) return 'ignored'
      this.replace(document, ticket.sheetId, {
        epoch: document.epoch, owner: delta.owner, workbookOwner: delta.workbookOwner,
        bands: delta.bands, visibility: delta.visibility, counts: delta.counts,
        // A newer commit settles its own request; another pending request stays pending.
        status: isCurrentPending ? READY : current.status,
        notice: isCurrentPending ? null : current.notice,
        // An applied view change invalidates geometry-dependent search layout (§10.1).
        geometry: null,
      })
      return 'applied'
    }
    if (!isCurrentPending || out.kind === 'progress') return 'ignored'
    this.replace(document, ticket.sheetId, { ...current, status: READY, notice: Object.freeze(noticeOf(out)) })
    return 'status'
  }
  /** Accepts geometry only for the exact committed owner and current font. */
  publishGeometry(docId: string, sheetId: string, owner: CacheIdentity, fontGen: number, payload: unknown): boolean {
    const document = this.documents.get(docId)
    const current = document?.sheets.get(sheetId)
    if (!document || !current || fontGen !== document.fontGen || !sameIdentityValue(owner, current.owner)) return false
    this.replace(document, sheetId, { ...current, geometry: Object.freeze({ owner, fontGen, payload }) })
    return true
  }
  /** Font readiness invalidates geometry only; order, counts and status stay. */
  fontChanged(docId: string): void {
    const document = this.documents.get(docId)
    if (!document) return
    document.fontGen++
    for (const [sheetId, current] of document.sheets) {
      if (current.geometry) this.replace(document, sheetId, { ...current, geometry: null })
    }
  }
  private pendingFor(ticket: PublicationTicket): { document: DocumentPublications; current: SheetPublication } | null {
    const document = this.documents.get(ticket.docId)
    const current = document?.sheets.get(ticket.sheetId)
    if (!document || !current || document.epoch !== ticket.epoch) return null
    if (current.status.kind !== 'pending' || current.status.requestId !== ticket.requestId) return null
    return { document, current }
  }
  private replace(document: DocumentPublications, sheetId: string, next: SheetPublication): void {
    document.sheets.set(sheetId, Object.freeze({ ...next }))
  }
}
const READY: SheetStatus = Object.freeze({ kind: 'ready' })
function noticeOf(out: WorkerReply): SheetNotice {
  switch (out.kind) {
    case 'cancelled': return { requestId: out.requestId, kind: 'cancelled' }
    case 'incomplete': return { requestId: out.requestId, kind: 'incomplete' }
    case 'quota': return { requestId: out.requestId, kind: 'quota' }
    case 'error':
      if (out.code === 'operation-refused') return { requestId: out.requestId, kind: 'refused', code: out.code }
      if (out.code === 'stale-source') return { requestId: out.requestId, kind: 'stale', code: out.code }
      return { requestId: out.requestId, kind: 'error', code: out.code }
    default: return { requestId: out.requestId, kind: 'error' }
  }
}

/** Host selection for the viewer (plan S4.2 no-worker support). Large mode
 * needs the resident worker and persistent page storage and refuses visibly
 * with a typed reason when either is missing. Small admitted documents use the
 * worker when available, else the in-thread port running the same chunked
 * TypeScript kernels (`createInThreadPort`). Ports are created lazily. */
export type ViewerHostChoice =
  | { kind: 'worker'; port: ViewerPort }
  | { kind: 'in-thread'; port: ViewerPort }
  | { kind: 'refused'; code: 'unsupported'; reason: 'worker-required' | 'storage-required'; message: string }
export interface ViewerHostInput {
  mode: 'small' | 'large'
  /** Worker factory, or null when the host cannot start a module worker. */
  worker: (() => ViewerPort) | null
  /** OPFS/IndexedDB page storage is available to the worker. */
  persistentStorage: boolean
  inThread: () => ViewerPort
}
export function selectViewerHost(input: ViewerHostInput): ViewerHostChoice {
  if (input.mode === 'large') {
    if (!input.worker) {
      return { kind: 'refused', code: 'unsupported', reason: 'worker-required', message: 'Large workbooks need the resident viewer worker, which this browser cannot start' }
    }
    if (!input.persistentStorage) {
      return { kind: 'refused', code: 'unsupported', reason: 'storage-required', message: 'Large workbooks need persistent page storage (OPFS or IndexedDB), which is unavailable' }
    }
    return { kind: 'worker', port: input.worker() }
  }
  return input.worker ? { kind: 'worker', port: input.worker() } : { kind: 'in-thread', port: input.inThread() }
}
