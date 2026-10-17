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
import type { ImportProgress } from '../xlsx/storage/importSession'
import type { ProfileId } from '../xlsx/view/contracts'
import type { ViewerKernel } from '../worker/viewerKernel'

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


// Bootstrap only (as src/worker/worker-entry.ts): the runtime loads on demand, so
// the perf build never has a lazy chunk importing this entry, which WebKit would
// re-evaluate (replacing onmessage and losing open sessions). Early messages queue.
const scope = globalThis as unknown as { onmessage: ((ev: MessageEvent) => void) | null }
let handler: ((ev: MessageEvent) => void) | null = null
const queued: MessageEvent[] = []
scope.onmessage = ev => { if (handler) handler(ev); else queued.push(ev) }
import('./viewerPerfWorkerBody').then(
  ({ perfMessageHandler }) => { handler = perfMessageHandler; for (const ev of queued.splice(0)) perfMessageHandler(ev) },
  (error: unknown) => { setTimeout(() => { throw error }) },
)
