/** Render worker (§9.2): one DocumentHost per Bun Worker; one job at a time. */
import { markRaster } from '../mark/mark-raster'
import { DocumentHost, RenderRefusal, type MarkSpec, type NotesJob, type Projection } from '../render'
import type { TextJob } from '../text/served'

export type WorkerRequest =
  | { type: 'open'; id: number; docRef: string; bytes: Uint8Array; contentPolicy: 'visible-only' | 'full'; servedCacheBytes?: number }
  | { type: 'render'; id: number; docKey: string; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number; crash?: boolean }
  | { type: 'drop'; id: number; docKey: string }
  /** Stage 3c notes job: plan or draw unit i's notes image. */
  | ({ type: 'notes'; id: number; docKey: string } & NotesJob)
  /** Stage 3a text job against the worker's served index (§4.1). */
  | { type: 'text'; id: number; docKey: string; job: TextJob; /** Admission bytes reserved for a prepare. */ grant?: number }
  /** Layer-1 encode job (§9.4): canonical PNG → session-marked PNG. */
  | { type: 'encode'; id: number; png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; hangMs?: number; crash?: boolean }
export type WorkerReply = { id: number; ok: true; result: unknown; stats: { parses: number; pages: number; docs: number } } | { id: number; ok: false; error: string; refusal?: string; stats: { parses: number; pages: number; docs: number } }

declare const self: Worker
const host = new DocumentHost()
const stats = () => ({ parses: host.parses, pages: host.pages, docs: host.docCount })

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const m = e.data
  try {
    let result: unknown
    if (m.type === 'open') { if (m.servedCacheBytes !== undefined) host.setServedCacheBytes(m.servedCacheBytes); result = await host.open(m.docRef, m.bytes, m.contentPolicy) }
    else if (m.type === 'drop') { host.drop(m.docKey); result = null }
    // A prepare without a grant is refused (grant 0) — the pool always sends one.
    else if (m.type === 'text') result = host.text(m.docKey, m.job, m.grant ?? 0)
    else if (m.type === 'notes') result = await host.notes(m.docKey, m)
    else {
      // Test hook: the worker dies mid-job (Bun defers a worker exit, so never reply).
      if (m.crash) { process.exit(1); await new Promise(() => {}) }
      if (m.hangMs) { const until = performance.now() + m.hangMs; while (performance.now() < until) { /* test hook: a wedged paint */ } }
      result = m.type === 'encode' ? markRaster(m) : host.render(m.docKey, m)
    }
    const transfer: ArrayBuffer[] = []
    const r = result as { png?: Uint8Array; ops?: Uint8Array } | null
    if (r?.png) transfer.push(r.png.buffer as ArrayBuffer)
    if (r?.ops) transfer.push(r.ops.buffer as ArrayBuffer)
    postMessage({ id: m.id, ok: true, result, stats: stats() } satisfies WorkerReply, transfer)
  } catch (err) {
    postMessage({ id: m.id, ok: false, error: (err as Error).message, ...(err instanceof RenderRefusal ? { refusal: err.code } : {}), stats: stats() } satisfies WorkerReply)
  }
}
