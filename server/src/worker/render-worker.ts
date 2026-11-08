/** Render worker (§9.2): one DocumentHost per Bun Worker; one job at a time. */
import { DocumentHost, RenderRefusal, type MarkSpec, type Projection } from '../render'

export type WorkerRequest =
  | { type: 'open'; id: number; docRef: string; bytes: Uint8Array; contentPolicy: 'visible-only' | 'full' }
  | { type: 'render'; id: number; docKey: string; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number }
  | { type: 'drop'; id: number; docKey: string }
export type WorkerReply = { id: number; ok: true; result: unknown; stats: { parses: number; pages: number; docs: number } } | { id: number; ok: false; error: string; refusal?: string; stats: { parses: number; pages: number; docs: number } }

declare const self: Worker
const host = new DocumentHost()
const stats = () => ({ parses: host.parses, pages: host.pages, docs: host.docCount })

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const m = e.data
  try {
    let result: unknown
    if (m.type === 'open') result = await host.open(m.docRef, m.bytes, m.contentPolicy)
    else if (m.type === 'drop') { host.drop(m.docKey); result = null }
    else {
      if (m.hangMs) { const until = performance.now() + m.hangMs; while (performance.now() < until) { /* test hook: a wedged paint */ } }
      result = host.render(m.docKey, m)
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
