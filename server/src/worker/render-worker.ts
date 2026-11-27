/** Render worker (§9.2): one DocumentHost per Bun Worker; one job at a time. */
import { markLayerRaster, markRaster, type LayerCrop } from '../mark/mark-raster'
import { DocumentHost, PaintablesUnbuilt, RenderRefusal, type MarkSpec, type NotesJob, type Projection, type SheetMode } from '../render'
import type { FfmpegInfo } from '../media/ffmpeg'
import type { ProbeKillRequest, ProbeRunReply, ProbeRunRequest } from '../media/probe-host'
import type { ProbeRunner, ProbeRunResult } from '../media/probe'
import type { ShowPlanJob } from '../show/build'
import type { MorphJob } from '../show/morph'
import type { TextJob } from '../text/served'
import type { SheetClock } from '../sheet/clock'
import type { TileJob } from '../sheet/tile'
import type { SheetTextJob } from '../sheet/text-ops'
import type { SheetViewJob } from '../sheet/view-ops'

export type WorkerRequest =
  /** Piece 5 (D21): the worker's init, posted first. `media`: ffmpeg and the pool's media temp dir; absent = media off. */
  | {
    type: 'init'; id: number
    /** `prefix`: this worker's file-name prefix in `dir` (Task 13): every file it stages bears it, and the pool deletes
     * them when the worker goes. Its probes run on the pool's thread (fix round 2): asked for by `probe-run`. */
    media?: { info: FfmpegInfo; dir: string; prefix?: string }
    /** TEST-ONLY (the pool sends it only with `testHooks`): each probe's hold. */ mediaHoldMs?: number
    /** TEST-ONLY (the pool sends it only with `testHooks`): `mediaPart` never returns once it wrote its file. */
    testWedge?: 'after-part'
  }
  /** Piece 5 (D21, D22): a unit's audio uses and their probed entries (a short read job). */
  | { type: 'media-refs'; id: number; docKey: string; unit: number }
  /** Piece 5 (D21): the projected part written 0600 into `dir` (the media dir) as `name` (the pool's choice, so it can
   * delete the file if no reply comes); the caller adopts it. */
  | { type: 'media-part'; id: number; docKey: string; partPath: string; dir: string; name?: string }
  | { type: 'open'; id: number; docRef: string; bytes: Uint8Array; contentPolicy: 'visible-only' | 'full'; servedCacheBytes?: number; /** Stage 4a (R5): absent = `'units'`. */ sheetMode?: SheetMode }
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
  /** Stage 4a (F23): one sheet's `sheet-layout` (from its servable snapshot under the session clock). */
  | { type: 'layout'; id: number; docKey: string; sheetIndex: number; clock: SheetClock }
  /** Stage 4a (S4 §6.5): one 512² sheet tile (canonical, pre-forensic) from its servable snapshot. */
  | { type: 'tile'; id: number; docKey: string; job: TileJob; hangMs?: number; crash?: boolean }
  /** Stage 4b (P2, P9): one text op (find, copy, a11y, fit) over a sheet's text view under the session clock. */
  | { type: 'sheet-text'; id: number; docKey: string; clock: SheetClock; job: SheetTextJob }
  /** Stage 4c (P3, P6): one sort, filter, derive, value-list or range-detect job over a sheet's key view. */
  | { type: 'sheet-view'; id: number; docKey: string; clock: SheetClock; contentPolicy: 'visible-only' | 'full'; job: SheetViewJob }
  /** Stage 4c Task 14a: prepare sheet `sheetIndex`'s sort data (snapshot, key view, `sheet-view` port) under the
   * session clock; a warm port answers at once. `hangMs` (test hook): a cold build waits this long first. */
  | { type: 'sheet-view-prep'; id: number; docKey: string; clock: SheetClock; contentPolicy: 'visible-only' | 'full'; sheetIndex: number; hangMs?: number }
  /** Stage 4c final review R2: drop what the document retains for derived view `viewHash` of `(sheetIndex, clock)`
   * (its last server holder is gone). */
  | { type: 'sheet-view-release'; id: number; docKey: string; clock: SheetClock; sheetIndex: number; viewHash: string }
  /** TEST-ONLY (the pool sends it only with `testHooks`): what the document retains for derived views. */
  | { type: 'sheet-view-retained'; id: number; docKey: string }
  /** Fix round 2: the pool's answer to this worker's `probe-run` (not a job: no reply). */
  | ProbeRunReply
export type WorkerReply = { id: number; ok: true; result: unknown; stats: WorkerStats } | { id: number; ok: false; error: string; refusal?: string; stats: WorkerStats }
export interface WorkerStats { parses: number; pages: number; docs: number; /** Piece 5: media probes run. */ probes: number }

declare const self: Worker
const host = new DocumentHost()

/**
 * Fix round 2 (controller ruling): this worker spawns nothing. A probe is asked of the pool, whose thread spawns,
 * supervises and reaps it and outlives this worker. Aborting the signal asks the pool to kill the run.
 */
const probes = new Map<number, { ok: (r: ProbeRunResult) => void; fail: (e: Error) => void }>()
let probeSeq = 0
const poolProbe: ProbeRunner = (req, signal) => new Promise<ProbeRunResult>((ok, fail) => {
  const id = ++probeSeq
  probes.set(id, { ok, fail })
  signal.addEventListener('abort', () => { postMessage({ kind: 'probe-kill', id } satisfies ProbeKillRequest) }, { once: true })
  postMessage({ kind: 'probe-run', id, ...req } satisfies ProbeRunRequest)
})
/** A result for a run no longer awaited is dropped. */
function probeAnswered(r: ProbeRunReply): void {
  const p = probes.get(r.id)
  if (!p) return
  probes.delete(r.id)
  if (r.ok) p.ok(r.result); else p.fail(new Error(`probe: ${r.error}`))
}
/** A units path (R5): run it; if the entry's paintables are not built yet (a tiles-mode open), build them once and
 * run again. The job's own checks run first, so a refused job never builds. */
async function unitsPath<T>(docKey: string, run: () => T): Promise<T> {
  try { return run() } catch (e) {
    if (!(e instanceof PaintablesUnbuilt)) throw e
    await host.ensurePaintables(docKey)
    return run()
  }
}
const stats = (): WorkerStats => ({ parses: host.parses, pages: host.pages, docs: host.docCount, probes: host.mediaProbes })

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const m = e.data
  if (m.type === 'probe-result') { probeAnswered(m); return }
  try {
    let result: unknown
    if (m.type === 'init') {
      host.configureMedia(m.media, m.media ? poolProbe : undefined)
      if (m.mediaHoldMs) host.mediaSeams = { hangMs: m.mediaHoldMs }
      // Bun (1.3.10, measured) terminates a worker only once its JS yields: the wedge awaits, never spins.
      if (m.testWedge === 'after-part') host.afterPartForTests = () => new Promise<void>(() => {})
      result = null
    }
    else if (m.type === 'media-refs') result = await host.mediaRefs(m.docKey, m.unit)
    else if (m.type === 'media-part') result = await host.mediaPart(m.docKey, m.partPath, m.dir, m.name)
    else if (m.type === 'open') { if (m.servedCacheBytes !== undefined) host.setServedCacheBytes(m.servedCacheBytes); result = await host.open(m.docRef, m.bytes, m.contentPolicy, m.sheetMode ? { sheetMode: m.sheetMode } : {}) }
    else if (m.type === 'drop') { host.drop(m.docKey); result = null }
    // A prepare without a grant is refused (grant 0) — the pool always sends one.
    else if (m.type === 'text') result = await unitsPath(m.docKey, () => host.text(m.docKey, m.job, m.grant ?? 0))
    else if (m.type === 'notes') result = await host.notes(m.docKey, m)
    // Before the final `else`: an unlisted type would fall through to `render`.
    else if (m.type === 'layout') result = await host.sheetLayout(m.docKey, m.sheetIndex, m.clock)
    else if (m.type === 'sheet-text') result = await host.sheetText(m.docKey, m.clock, m.job)
    // A derived view's order is cached here: its buffers are copied, never transferred.
    else if (m.type === 'sheet-view') result = await host.sheetView(m.docKey, m.clock, m.contentPolicy, m.job)
    else if (m.type === 'sheet-view-prep') result = await host.sheetViewPrep(m.docKey, m.clock, m.contentPolicy, m.sheetIndex, m.hangMs)
    else if (m.type === 'sheet-view-release') { await host.releaseView(m.docKey, m.sheetIndex, m.clock, m.viewHash); result = null }
    else if (m.type === 'sheet-view-retained') result = await host.retainedViewsForTests(m.docKey)
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
        : m.type === 'tile' ? await host.renderTile(m.docKey, m.job)
        : await unitsPath(m.docKey, () => host.render(m.docKey, m))
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
