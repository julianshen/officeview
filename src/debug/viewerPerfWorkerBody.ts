/**
 * Runtime of the S10.2 perf-page worker (see viewerPerfWorker.ts, which loads
 * this module on demand so no lazy chunk ever imports the worker entry).
 */
import { OfficePackage } from '../core/zip'
import { parseXlsx } from '../xlsx/parse'
import { openImportSession } from '../xlsx/storage/pagedImport'
import type { ProfileId } from '../xlsx/view/contracts'
import { createWorkerMessageHandler } from '../worker/workerMessages'
import { ViewerKernel } from '../worker/viewerKernel'
// The kernel imports its large-source opener on demand; named here so the
// harness server's locked-down module graph serves it.
import '../worker/viewerLargeSource'

import type { PerfImportEvent, PerfReply, PerfRequest, PerfSessionMemory } from './viewerPerfWorker'

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

export function perfMessageHandler(ev: MessageEvent): void {
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
