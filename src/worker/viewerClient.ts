/**
 * Main-thread client for the resident viewer worker (spec §10.2, plan S4.1).
 *
 * Correlates replies to commands by requestId, surfaces progress, ignores late
 * or foreign messages (including legacy render responses on the same worker)
 * and accounts every ArrayBuffer it transfers or leaves to structured clone.
 * Main-safe: imports only the protocol module.
 */
import type { WorkerCommand, WorkerCommandKind, WorkerOwner, WorkerReply } from '../xlsx/view/contracts'
import {
  isArrayBuffer, isTerminalReply, isViewerReply, TransferAccounting, type CommandOf, type TransferLedger,
} from './viewerProtocol'

type MessageListener = (event: { data: unknown }) => void
/** Worker, MessagePort or the in-process port. */
export interface ViewerPort {
  postMessage(message: unknown, transfer: Transferable[]): void
  addEventListener?(type: 'message', listener: MessageListener): void
  removeEventListener?(type: 'message', listener: MessageListener): void
  onmessage?: ((event: MessageEvent) => void) | null
}
export interface SendOptions<K extends WorkerCommandKind> {
  onProgress?: (reply: Extract<WorkerReply<K>, { kind: 'progress' }>) => void
}
interface Pending {
  kind: WorkerCommandKind
  owner: WorkerOwner | undefined
  resolve: (reply: WorkerReply) => void
  reject: (error: Error) => void
  onProgress?: (reply: WorkerReply) => void
}

export class ViewerClient {
  private readonly pending = new Map<string, Pending>()
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
  send<K extends WorkerCommandKind>(command: CommandOf<K>, options: SendOptions<K> = {}): Promise<WorkerReply<K>> {
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
        resolve: resolve as (reply: WorkerReply) => void, reject,
        onProgress: options.onProgress as ((reply: WorkerReply) => void) | undefined,
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
    if (data.commandKind !== entry.kind || (entry.owner !== undefined && !sameOwner(data.owner, entry.owner))) {
      this.rejected++
      return
    }
    if (!isTerminalReply(data)) { entry.onProgress?.(data); return }
    this.pending.delete(data.requestId)
    entry.resolve(data)
  }
  /** Detaches from the port and rejects every pending request. */
  dispose(): void {
    this.detach()
    for (const entry of this.pending.values()) entry.reject(new Error('Viewer client disposed'))
    this.pending.clear()
  }
}
export function createViewerClient(port: ViewerPort, prefix?: string): ViewerClient {
  return new ViewerClient(port, prefix)
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
