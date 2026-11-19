/** Render worker (§9.2): one DocumentHost per Bun Worker; one job at a time. */
import { markLayerRaster, markRaster, type LayerCrop } from '../mark/mark-raster'
import { DocumentHost, RenderRefusal, type MarkSpec, type NotesJob, type Projection } from '../render'
import type { FfmpegInfo } from '../media/ffmpeg'
import { observeSpawns } from '../media/sandbox'
import type { ShowPlanJob } from '../show/build'
import type { MorphJob } from '../show/morph'
import type { TextJob } from '../text/served'

export type WorkerRequest =
  /** Piece 5 (D21): the worker's init, posted first. `media`: ffmpeg and the pool's media temp dir; absent = media off. */
  | {
    type: 'init'; id: number
    /** `prefix`: this worker's file-name prefix in `dir` (Task 13); the pool deletes what bears it when the worker goes. */
    media?: { info: FfmpegInfo; dir: string; prefix?: string }
    /** Task 13: slots the pool shares with this worker, holding the pid of each sandboxed run in progress. */
    pids?: SharedArrayBuffer
    /** TEST-ONLY (the pool sends it only with `testHooks`): each probe's hold. */ mediaHoldMs?: number
  }
  /** Piece 5 (D21, D22): a unit's audio uses and their probed entries (a short read job). */
  | { type: 'media-refs'; id: number; docKey: string; unit: number }
  /** Piece 5 (D21): the projected part written 0600 into `dir` (the media dir) as `name` (the pool's choice, so it can
   * delete the file if no reply comes); the caller adopts it. */
  | { type: 'media-part'; id: number; docKey: string; partPath: string; dir: string; name?: string }
  | { type: 'open'; id: number; docRef: string; bytes: Uint8Array; contentPolicy: 'visible-only' | 'full'; servedCacheBytes?: number }
  | { type: 'render'; id: number; docKey: string; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number; crash?: boolean }
  | { type: 'drop'; id: number; docKey: string }
  /** Stage 3c notes job: plan or draw unit i's notes image. */
  | ({ type: 'notes'; id: number; docKey: string } & NotesJob)
  /** Stage 3a text job against the worker's served index (§4.1). */
  | { type: 'text'; id: number; docKey: string; job: TextJob; /** Admission bytes reserved for a prepare. */ grant?: number }
  /** Layer-1 encode job (§9.4): canonical PNG → session-marked PNG. */
  | { type: 'encode'; id: number; png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; hangMs?: number; crash?: boolean }
  /** Piece 4: one slide's show-plan bundle (plan, entry frame, visibly marked layers). */
  | { type: 'show-plan'; id: number; docKey: string; job: ShowPlanJob; /** The admitted bytes the build's ledger holds (PR #78 round 13). */ grant?: number; hangMs?: number; crash?: boolean }
  /** Piece 4: the keyframe track of one composited step. */
  | { type: 'show-track'; id: number; docKey: string; job: ShowPlanJob & { step: number }; grant?: number; hangMs?: number; crash?: boolean }
  /** Piece 4 Task 7: one Morph pair (`from`, `to`, `fromStep`): its record and assets, or its composited frames. */
  | { type: 'show-morph'; id: number; docKey: string; job: MorphJob; grant?: number; hangMs?: number; crash?: boolean }
  /** Piece 4 D7: a show-plan layer's per-session Layer-1 (the page basis cropped to its rest). */
  | { type: 'mark-layer'; id: number; png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; crop: LayerCrop; hangMs?: number; crash?: boolean }
export type WorkerReply = { id: number; ok: true; result: unknown; stats: WorkerStats } | { id: number; ok: false; error: string; refusal?: string; stats: WorkerStats }
export interface WorkerStats { parses: number; pages: number; docs: number; /** Piece 5: media probes run. */ probes: number }

/**
 * Task 13 (carried from Task 3): every sandboxed run of this worker is listed, by its group leader's pid, in memory the
 * pool shares, from its spawn to its exit. A terminated worker's wall timers and watchdog die with it; the pool then
 * kills each listed group. Written synchronously in the spawning step, so no message can be lost on the way. With
 * every slot taken a run goes unlisted (the pool's runs are serial per worker; it is bounded by its CPU limit).
 */
function trackChildren(slots: Int32Array): void {
  observeSpawns(({ pid, exited }) => {
    for (let i = 0; i < slots.length; i++) {
      if (Atomics.compareExchange(slots, i, 0, pid) !== 0) continue
      void exited.then(() => { Atomics.compareExchange(slots, i, pid, 0) })
      return
    }
  })
}

declare const self: Worker
const host = new DocumentHost()
const stats = (): WorkerStats => ({ parses: host.parses, pages: host.pages, docs: host.docCount, probes: host.mediaProbes })

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const m = e.data
  try {
    let result: unknown
    if (m.type === 'init') { if (m.pids) trackChildren(new Int32Array(m.pids)); host.configureMedia(m.media); if (m.mediaHoldMs) host.mediaSeams = { hangMs: m.mediaHoldMs }; result = null }
    else if (m.type === 'media-refs') result = await host.mediaRefs(m.docKey, m.unit)
    else if (m.type === 'media-part') result = await host.mediaPart(m.docKey, m.partPath, m.dir, m.name)
    else if (m.type === 'open') { if (m.servedCacheBytes !== undefined) host.setServedCacheBytes(m.servedCacheBytes); result = await host.open(m.docRef, m.bytes, m.contentPolicy) }
    else if (m.type === 'drop') { host.drop(m.docKey); result = null }
    // A prepare without a grant is refused (grant 0) — the pool always sends one.
    else if (m.type === 'text') result = host.text(m.docKey, m.job, m.grant ?? 0)
    else if (m.type === 'notes') result = await host.notes(m.docKey, m)
    else {
      // Test hook: the worker dies mid-job (Bun defers a worker exit, so never reply).
      if (m.crash) { process.exit(1); await new Promise(() => {}) }
      if (m.hangMs) { const until = performance.now() + m.hangMs; while (performance.now() < until) { /* test hook: a wedged paint */ } }
      // A bundle's PNGs are nested (and the worker's cache keeps them): they are copied, never transferred.
      result = m.type === 'encode' ? markRaster(m)
        : m.type === 'mark-layer' ? markLayerRaster(m)
        : m.type === 'show-plan' ? await host.showPlan(m.docKey, m.job, m.grant)
        : m.type === 'show-track' ? await host.showTrack(m.docKey, m.job, m.grant)
        : m.type === 'show-morph' ? await host.showMorph(m.docKey, m.job, m.grant)
        : host.render(m.docKey, m)
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
