/**
 * Dev-only worker for the S10.2 performance page (src/debug/viewer-perf.html).
 *
 * It runs the shipped viewer kernel exactly as dist/worker.js does
 * (createWorkerMessageHandler + a ViewerKernel built the same way), and adds a
 * read-only side channel the protocol deliberately lacks: `__perf` messages
 * answer the worker-owned memory ledger (kernel lifecycle stats and each open
 * session's pool), run a dedicated importer throughput pass over the same
 * openImportSession the paged route uses, and trigger a GC when the browser
 * exposes one. Viewer commands are never inspected beyond their docId, and
 * `__perf` replies carry no viewer reply fields, so the viewer client ignores
 * them. Excluded from the library build (tsconfig.build.json excludes src/debug).
 */
import { OfficePackage } from '../core/zip'
import { parseXlsx } from '../xlsx/parse'
import { openImportSession } from '../xlsx/storage/pagedImport'
import type { ImportProgress } from '../xlsx/storage/importSession'
import type { ProfileId } from '../xlsx/view/contracts'
import { createWorkerMessageHandler } from '../worker/worker-entry'
import { ViewerKernel } from '../worker/viewerKernel'
// The kernel imports its large-source opener on demand; named here so the
// harness server's locked-down module graph serves it.
import '../worker/viewerLargeSource'

export type PerfRequest =
  | { __perf: 'stats'; id: number }
  | { __perf: 'gc'; id: number }
  | { __perf: 'import'; id: number; blob: Blob; profile: ProfileId }

export interface PerfImportEvent { t: number; phase: ImportProgress['phase']; expandedBytes: number; populatedCells: number; compressedRead: number }
export interface PerfSessionMemory { docId: string; memory: ReturnType<NonNullable<ReturnType<ViewerKernel['session']>>['memoryStats']> }
export type PerfReply =
  | { __perfReply: number; kind: 'stats'; lifecycle: ReturnType<ViewerKernel['lifecycleStats']>; sessions: PerfSessionMemory[]; transfer: ReturnType<ViewerKernel['ledger']> }
  | { __perfReply: number; kind: 'gc'; collected: boolean }
  | { __perfReply: number; kind: 'import'; startedAt: number; events: PerfImportEvent[]; outcome: string; expandedTotal: number | 'unknown' }
  | { __perfReply: number; kind: 'error'; message: string }

const scope = globalThis as unknown as {
  onmessage: ((ev: MessageEvent) => void) | null
  postMessage: (message: unknown, transfer?: Transferable[]) => void
  gc?: () => void
}
const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) })
const viewer = createWorkerMessageHandler((message, transfer) => scope.postMessage(message, transfer), kernel)
/** Documents this worker has been asked to open (ledger lookup keys). */
const docIds = new Set<string>()
let importCount = 0

async function runImport(id: number, blob: Blob, profile: ProfileId): Promise<PerfReply> {
  const events: PerfImportEvent[] = []
  let expandedTotal: number | 'unknown' = 'unknown'
  const startedAt = performance.now()
  const session = await openImportSession({ kind: 'blob', blob }, {
    docId: `perf-import-${++importCount}`, profile,
    onProgress: p => {
      expandedTotal = p.expandedTotal
      events.push({ t: performance.now(), phase: p.phase, expandedBytes: p.expandedBytes, populatedCells: p.populatedCells, compressedRead: p.compressedRead })
    },
  })
  try {
    const outcome = await session.done
    return { __perfReply: id, kind: 'import', startedAt, events, outcome: outcome.kind === 'incomplete' ? `incomplete: ${outcome.failure.code}` : outcome.kind, expandedTotal }
  } finally {
    await session.release()
  }
}

async function answer(request: PerfRequest): Promise<PerfReply> {
  switch (request.__perf) {
    case 'stats': {
      const sessions: PerfSessionMemory[] = []
      for (const docId of docIds) {
        const session = kernel.session(docId)
        if (session) sessions.push({ docId, memory: session.memoryStats() })
      }
      return { __perfReply: request.id, kind: 'stats', lifecycle: kernel.lifecycleStats(), sessions, transfer: kernel.ledger() }
    }
    case 'gc':
      scope.gc?.()
      return { __perfReply: request.id, kind: 'gc', collected: typeof scope.gc === 'function' }
    case 'import':
      return runImport(request.id, request.blob, request.profile)
  }
}

scope.onmessage = (ev: MessageEvent) => {
  const data = ev.data as Partial<PerfRequest> & { kind?: unknown; docId?: unknown }
  if (data && typeof data === 'object' && typeof data.__perf === 'string' && typeof data.id === 'number') {
    const id = data.id
    answer(data as PerfRequest).then(
      reply => scope.postMessage(reply),
      error => scope.postMessage({ __perfReply: id, kind: 'error', message: error instanceof Error ? error.message : String(error) } satisfies PerfReply),
    )
    return
  }
  if (data && data.kind === 'open' && typeof data.docId === 'string') docIds.add(data.docId)
  viewer(ev)
}
