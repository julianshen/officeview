/**
 * Main-thread client for the resident viewer worker (spec §10.2, plan S4.1).
 *
 * Correlates replies to commands by requestId, surfaces progress, ignores late
 * or foreign messages (including legacy render responses on the same worker)
 * and accounts every ArrayBuffer it transfers or leaves to structured clone.
 * Main-safe: imports only the protocol module.
 */
import type { WorkerCommandKind, WorkerReply } from '../xlsx/view/contracts'
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
  resolve: (reply: WorkerReply) => void
  reject: (error: Error) => void
  onProgress?: (reply: WorkerReply) => void
}

export class ViewerClient {
  private readonly pending = new Map<string, Pending>()
  private readonly accounting = new TransferAccounting()
  private counter = 0
  private ignored = 0
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

  /** Resolves with the request's terminal reply. Typed arrays that own their whole
   * buffer are transferred (the worker takes ownership; the caller's view detaches);
   * partial views, and everything inside an open command's source document, are
   * left to structured clone and accounted as copies. */
  send<K extends WorkerCommandKind>(command: CommandOf<K>, options: SendOptions<K> = {}): Promise<WorkerReply<K>> {
    if (this.pending.has(command.requestId)) return Promise.reject(new Error(`Duplicate requestId ${command.requestId}`))
    const transfer: ArrayBuffer[] = []
    const seen = new Set<object>()
    const ownership = command.kind !== 'open'
    const visit = (value: unknown): void => {
      if (typeof value !== 'object' || value === null || seen.has(value)) return
      seen.add(value)
      if (ArrayBuffer.isView(value)) {
        const buffer = value.buffer
        if (ownership && isArrayBuffer(buffer) && value.byteOffset === 0 && value.byteLength === buffer.byteLength) {
          if (!transfer.includes(buffer)) transfer.push(buffer)
        } else this.accounting.copy(value.byteLength)
        return
      }
      if (value instanceof Map) { for (const entry of value.values()) visit(entry); return }
      // A Blob is passed by handle; its bytes are copied once inside the worker.
      if (typeof Blob !== 'undefined' && value instanceof Blob) return
      for (const entry of Array.isArray(value) ? value : Object.values(value)) visit(entry)
    }
    visit(command)
    return new Promise<WorkerReply<K>>((resolve, reject) => {
      this.pending.set(command.requestId, {
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
  private receive(data: unknown): void {
    if (!isViewerReply(data)) return
    const entry = this.pending.get(data.requestId)
    if (!entry) { this.ignored++; return }
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
