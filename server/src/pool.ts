/**
 * Render pool (§9.2): Bun Workers with document affinity, admission through
 * the shared AdmissionModel (§14 caps), a 2 s render deadline with
 * drain-then-refund retirement, and recycling after 256 pages.
 *
 * [call] Stage 1 parses in the render worker that owns a docRef (affinity);
 * §9.2's separate parse worker is Stage 2. Stage 1 registers no embedded
 * fonts (the throwing registrar in render.ts), so the 64 MiB font-budget
 * recycle trigger cannot fire yet; the page trigger is live.
 */
import { randomBytes } from 'node:crypto'
import { unlinkSync } from 'node:fs'
import { availableParallelism } from 'node:os'
import { join, resolve } from 'node:path'
import { AdmissionModel, DEFAULT_CAPS, type Caps, type Job } from './admission'
import type { LayerCrop, MarkLayerResult, MarkRasterResult } from './mark/mark-raster'
import { SHOW_BUILD_BYTES, showRequestBytes } from './show/grant'
import { SHOW_PLAN_DEADLINE_MS, type BuiltShowPlan, type ShowPlanJob, type ShowRefusal } from './show/build'
import type { BuiltMorph, MorphJob } from './show/morph'
import type { MarkSpec, MediaUnitRefs, NotesJob, NotesPlan, NotesRaster, OpenResult, Projection, RenderedUnit, SheetMode } from './render'
import type { FfmpegInfo } from './media/ffmpeg'
import { PROBE_DECK_MS, PROBE_PART_MS } from './media/probe'
import { createWorkerProbes, type ProbeKillRequest, type ProbeRunRequest, type WorkerProbes } from './media/probe-host'
import { sweepPrefixed } from './media/temp'
import type { WorkerReply, WorkerRequest } from './worker/render-worker'
import type { SheetClock } from './sheet/clock'
import type { RenderedTile, TileJob } from './sheet/tile'
import type { SheetTextJob, SheetTextReply } from './sheet/text-ops'
import type { SheetViewJob, SheetViewReply } from './sheet/view-ops'
import { canonicalView, isIdentityView, orderBytes, type CanonicalView } from './sheet/view-spec'
import type { StreamSheetLayout } from '../../src/stream/sheet-wire'
import { prepareWorkspace, TEXT_PREPARE_ALLOWANCE, type TextJob, type TextReply } from './text/served'
import { startWorker, type StartOptions, type StartedWorker } from './worker-start'

export class AdmissionRefused extends Error { constructor(readonly reason: 'session-queue' | 'process-queue' | 'too-large') { super(`admission refused: ${reason}`); this.name = 'AdmissionRefused' } }
export class RenderTimeout extends Error { constructor() { super('render deadline exceeded'); this.name = 'RenderTimeout' } }
export class RenderFailed extends Error { constructor(m: string, readonly refusal?: string) { super(m); this.name = 'RenderFailed' } }
/** Piece 5 (Task 8): a show job's media warm-up passed its own deadline. Never remembered as the slide's refusal. */
class MediaWarmTimeout extends Error { constructor() { super('media warm-up deadline exceeded'); this.name = 'MediaWarmTimeout' } }
/** A job whose socket closed before it reached a worker: nothing was run (PR #78 rounds 4–5: show, unit, notes, text). */
export class ShowCancelled extends RenderFailed { constructor() { super('cancelled: the requesting socket closed'); this.name = 'ShowCancelled' } }
/** Stage 4c Task 14a: a sheet's sort data (its key view and `sheet-view` port) could not be prepared inside the prep
 * deadline (`sheetPrepDeadlineMs`), or its build failed. The socket answers `view-refused {reason: 'too-large'}`. */
export class SheetPrepRefused extends Error { constructor(readonly cause: 'timeout' | 'failed') { super(`sheet prep ${cause === 'timeout' ? 'deadline exceeded' : 'failed'}`); this.name = 'SheetPrepRefused' } }
/** A text op whose units do not all fit the worker's served-index cache at once (Task 5 maps it to copy-too-large). */
export class TextTooLarge extends Error { constructor(readonly missing: number[]) { super(`text range exceeds the served-index cache (${missing.length} units missing)`); this.name = 'TextTooLarge' } }

export interface RenderJob { docRef: string; contentPolicy: 'visible-only' | 'full'; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number; crash?: boolean }
/** A notes plan or draw (Stage 3c) on the docRef's worker. */
export type PoolNotesJob = { docRef: string; contentPolicy: 'visible-only' | 'full' } & NotesJob
/** A per-session Layer-1 encode (§9.4); no document affinity. */
export interface EncodeJob { png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; hangMs?: number; crash?: boolean }
/** Piece 4: a show-plan bundle or track job on the docRef's worker. `media` (piece 5, D21: `can.media`) joins the
 * refusal and remembered keys; absent is false. */
export type PoolShowJob = { docRef: string; contentPolicy: 'visible-only' | 'full'; media?: boolean; hangMs?: number; crash?: boolean } & ShowPlanJob
/** Piece 5 (D21): a media read job's deadline. `mediaRefs` may probe a deck's whole budget (10 s) plus one part's
 * staging; `mediaPart` streams up to 512 MiB. Neither is an encode. */
export const MEDIA_JOB_DEADLINE_MS = PROBE_DECK_MS + 2 * PROBE_PART_MS + 1000
/** Piece 4 Task 7: a Morph pair job on the docRef's worker. */
export type PoolMorphJob = { docRef: string; contentPolicy: 'visible-only' | 'full'; hangMs?: number; crash?: boolean } & MorphJob
/** Piece 4 D7: a show-plan layer's per-session Layer-1 encode. */
export interface MarkLayerJob { png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; crop: LayerCrop; hangMs?: number; crash?: boolean }
export { SHOW_PLAN_DEADLINE_MS }
export type PoolTextJob = { docRef: string; contentPolicy: 'visible-only' | 'full' } & TextJob
/** Admission charge of one text op itself (index assembly, answer text);
 * every prepare it needs is charged separately (`prepareWorkspace`). */
export const TEXT_JOB_BYTES = TEXT_PREPARE_ALLOWANCE
/** Stage 4a (F23): a `sheet-layout` job on the docRef's worker: sheet `sheetIndex` under the session clock. */
export interface PoolLayoutJob { docRef: string; contentPolicy: 'visible-only' | 'full'; sheetIndex: number; clock: SheetClock }
/** Stage 4a (S4 §6.5): a tile job on the docRef's worker. `hangMs`/`crash`: the `RenderJob` test hooks (F25). */
export type PoolTileJob = { docRef: string; contentPolicy: 'visible-only' | 'full'; hangMs?: number; crash?: boolean } & TileJob
export { TILE_BYTES_ESTIMATE } from './sheet/tile'
/** Admission charge of one layout job (the snapshot it may build, and the layout it copies out). */
export const LAYOUT_JOB_BYTES = 16 * 1024 * 1024
/** Stage 4b: a `sheet-text` job on the docRef's worker: one text op under the session clock. */
export type PoolSheetTextJob = { docRef: string; contentPolicy: 'visible-only' | 'full'; clock: SheetClock } & SheetTextJob
/** Admission charge of one sheet-text job: what a layout job charges (the snapshot and text view it may build) plus a
 * fit's measurement scratch (`FIT_MEASURE_BYTES`, 16 MiB, in `sheet/text-ops.ts`, not imported here: the pool thread
 * never loads the text engines). */
export const SHEET_TEXT_JOB_BYTES = LAYOUT_JOB_BYTES + 16 * 1024 * 1024
/** Stage 4c: a `sheet-view` job on the docRef's worker: one sort, filter, derive, value-list or range-detect op. */
export type PoolSheetViewJob = { docRef: string; contentPolicy: 'visible-only' | 'full'; clock: SheetClock } & SheetViewJob
/** The fixed part of a sheet-view job's admission charge (the key view, the port and the engine's working set). */
export const SHEET_VIEW_JOB_BYTES = 16 * 1024 * 1024
/** The view a sheet-view job leaves (P3: a sort sets `sortFilter := base.filter`, a filter keeps the sort). */
export function sheetViewNext(job: SheetViewJob): CanonicalView {
  switch (job.op) {
    case 'derive': case 'values': case 'extent': return job.view
    case 'sort': return canonicalView(job.sort ? { sort: job.sort, sortFilter: job.base.filter, filter: job.base.filter } : { sort: null, sortFilter: null, filter: job.base.filter })
    case 'filter': return canonicalView({ sort: job.base.sort, sortFilter: job.base.sortFilter, filter: job.filter })
    case 'detect': return { sort: null, sortFilter: null, filter: null }
  }
}
/**
 * Stage 4c Task 14a (owner ruling 2026-11-28): the prep job (snapshot, key view, `sheet-view` port) has its own
 * deadline, `renderDeadlineMs + PREP_MS_PER_MIB` per MiB of the document's admitted bytes, at most
 * `PREP_DEADLINE_CAP_MS`. Measured on `oneMillionSort` (7.6 MiB admitted, 1,048,576 rows), its snapshot already
 * built: a cold prep took 0.55–0.95 s under `'full'` and 1.6–2.05 s under `'visible-only'`, loaded runs included; a
 * cold snapshot adds 0.35–0.6 s. Its deadline is 2,000 + 7.6 × 1,000 ≈ 9.6 s, 3.7× the slowest prep with a cold
 * snapshot. The cap is the parse's own default deadline (`openDeadlineMs`).
 */
export const PREP_MS_PER_MIB = 1000
export const PREP_DEADLINE_CAP_MS = 30_000
/** The prep deadline of a document of `bytes` admitted bytes (`cap`: `PREP_DEADLINE_CAP_MS` unless a test injects one). */
export function sheetPrepDeadlineMs(bytes: number, renderDeadlineMs: number, cap = PREP_DEADLINE_CAP_MS): number {
  return Math.min(cap, renderDeadlineMs + Math.ceil(PREP_MS_PER_MIB * bytes / (1024 * 1024)))
}
/** Whether a sheet-view job reads the sheet's port (every op but an identity-view extent). */
const readsPort = (job: SheetViewJob) => job.op !== 'extent' || !isIdentityView(canonicalView(job.view))
/** A sheet-view job's admission charge: `orderBytes(next view) + 16 MiB`. */
export const sheetViewJobBytes = (job: SheetViewJob) => orderBytes(sheetViewNext(job)) + SHEET_VIEW_JOB_BYTES
/** Stage 4c (Task 6 fix round 1): a tile or sheet-text job under a sort or filter view may re-derive that view's order in
 * the worker (a derived-view cache miss, not knowable here), so it is charged `orderBytes(view)` on top of its estimate;
 * 0 for no view or the identity view. */
const viewOrderBytes = (view: CanonicalView | undefined) => view ? orderBytes(view) : 0
/** `servedCacheBytes`: the worker's served-index bound per document (R6 default 32 MiB). */
export interface PoolOptions { /** Piece 4 (tests shorten it): the show-plan job deadline, default `SHOW_PLAN_DEADLINE_MS`. */ showDeadlineMs?: number; workers?: number; caps?: Caps; recyclePages?: number; now?: () => number; testHooks?: boolean; openDeadlineMs?: number; servedCacheBytes?: number; /** Piece 5 (D21): ffmpeg and the media temp dir, forwarded in each worker's init. */ media?: { info: FfmpegInfo; dir: string }; /** TEST-ONLY (honoured with `testHooks`): every media probe is held this long (`MediaTableSeams.hangMs`). */ mediaHoldMs?: number; /** TEST-ONLY (honoured with `testHooks`): the media warm-up and media job deadline, default `MEDIA_JOB_DEADLINE_MS`. */ mediaDeadlineMs?: number; /** TEST-ONLY (honoured with `testHooks`): `after-part`, `mediaPart` never answers once it wrote its file. */ mediaWedge?: 'after-part'; /** TEST-ONLY (honoured with `testHooks`): builds each worker in place of `new Worker` (`worker-start.ts` seam). */ workerFactoryForTests?: (url: string, options: WorkerOptions) => Worker; /** TEST-ONLY (honoured with `testHooks`): the start bound, default `WORKER_LOAD_BOUND_MS`. */ workerLoadBoundMsForTests?: number; /** TEST-ONLY (honoured with `testHooks`): the prep deadline's cap, default `PREP_DEADLINE_CAP_MS`. */ prepDeadlineCapMsForTests?: number; /** TEST-ONLY (honoured with `testHooks`): each cold port build waits this long first. */ prepHangMsForTests?: number }

const WORKER_URL = resolve(import.meta.dir, 'worker/render-worker.ts')
const PRELOAD = resolve(import.meta.dir, '../../scripts/stream-s0/napi-preload.ts')
export const defaultWorkers = () => Math.max(1, Math.min(8, Math.floor(availableParallelism() / 2)))

interface Slot {
  index: number
  /** Started through the process's start gate (`worker-start.ts`, a Bun 1.3.10 workaround): one load at a time, and a
   * terminate during a load waits for it to end. */
  worker: StartedWorker
  busy: boolean
  queue: Array<() => void>
  pending: Map<number, { ok: (r: WorkerReply) => void }>
  /** Set once the current worker's `close` event fired (terminate or death). */
  closed: boolean
  /** Resolves when the current worker's `close` event fires. */
  whenClosed: Promise<void>
  docs: Map<string, string>       // docRef|policy → docKey on this worker
  pages: number
  parses: number
  /** Piece 5: media probes run by the current worker. */
  probes: number
  generation: number
  /** The preparation grant of the show job holding this slot (0: none). */
  prep: number
  /** Fix round 2: retires the current worker's probes (refuses its later requests, kills its running ones). */
  retire?: () => void
}

/** Round 13d: show requests a session may have waiting for its show slot (one more is running). */
export const showQueueBound = (caps: Caps) => caps.perSessionOutstanding + caps.perSessionQueue - 1

export class Pool {
  readonly admission: AdmissionModel
  private slots: Slot[] = []
  private affinity = new Map<string, number>()             // docRef → slot index
  private sources = new Map<string, Uint8Array>()          // docRef → bytes (re-open after recycle)
  private nextId = 1
  private retired = 0
  private recycled = 0
  private encodes = 0
  private readonly recyclePages: number
  private readonly openDeadlineMs: number
  private readonly now: () => number
  private readonly testHooks: boolean
  private readonly servedCacheBytes?: number
  private readonly showDeadlineMs: number
  private readonly media?: { info: FfmpegInfo; dir: string }
  private readonly mediaHoldMs?: number
  private readonly mediaWedge?: 'after-part'
  private readonly mediaDeadlineMs: number
  private readonly startOptions: StartOptions
  private readonly prepCapMs: number
  private readonly prepHangMs?: number
  /**
   * Piece 4 (controller ruling): show-plan refusals by slide (`docRef|policy|unit|scale|motion|strictMarks`: the
   * L1 tuple; the fingerprint is per process). A slide that timed out retired its worker, and with it the worker's
   * cache: a repeat request answers here, with no admission and no worker.
   */
  private showRefusals = new Map<string, ShowRefusal>()
  /**
   * Task 7 follow-up (controller ruling): track and Morph-pair refusals, the same way. A track is keyed by the
   * slide's L1 tuple and its step, a pair by `docRef|policy|morph|from|to|fromStep|scale|motion|strictMarks`.
   * `timeout`: the job passed the show deadline (its worker was retired); a repeat throws `RenderTimeout` at once.
   * `morph-fallback`: `planMorph` refused the pair (deterministic). Cleared with `showRefusals` (re-register, forget).
   */
  private jobRefusals = new Map<string, 'timeout' | 'morph-fallback'>()
  /** Runs `run` unless `key` is known refused; a timeout (and a Morph refusal) is remembered while the document is registered. */
  private async remembered<T>(docRef: string, key: string, run: () => Promise<T>): Promise<T> {
    const known = this.jobRefusals.get(key)
    if (known === 'timeout') throw new RenderTimeout()
    if (known === 'morph-fallback') return { refused: 'morph-fallback' } as T
    try {
      const r = await run()
      if (r && typeof r === 'object' && (r as { refused?: unknown }).refused === 'morph-fallback' && this.sources.has(docRef)) this.jobRefusals.set(key, 'morph-fallback')
      return r
    } catch (e) {
      if (e instanceof RenderTimeout && this.sources.has(docRef)) this.jobRefusals.set(key, 'timeout')
      throw e
    }
  }
  /** `${docRef}|${policy}` → unit sizes from the last parse (prepare charges). */
  private unitSizes = new Map<string, OpenResult['units']>()
  private disposed = false
  /** `dispose()`'s result: every slot's worker closed. */
  private disposedWait?: Promise<void>
  /** Queued admission jobs waiting to start: resolve(true) = started, resolve(false) = cancelled. */
  private waiters = new Map<number, (started: boolean) => void>()
  /** Worker-time preparation reservations (owner ruling, round 13b). */
  private prep = { reserved: 0, peak: 0, peakPerWorker: 0, peakPerSession: 0 }
  private prepBySession = new Map<string, number>()
  /** Round 13c: each session's show slot (one running show job) and its waiters, in order. */
  private showSessions = new Map<string, { busy: boolean; waiters: Array<() => void> }>()

  constructor(opts: PoolOptions = {}) {
    this.admission = new AdmissionModel(opts.caps ?? DEFAULT_CAPS)
    this.recyclePages = opts.recyclePages ?? 256
    this.openDeadlineMs = opts.openDeadlineMs ?? 30_000
    this.now = opts.now ?? (() => performance.now())
    this.testHooks = opts.testHooks ?? false
    this.servedCacheBytes = opts.servedCacheBytes
    this.showDeadlineMs = opts.showDeadlineMs ?? SHOW_PLAN_DEADLINE_MS
    this.media = opts.media
    this.mediaHoldMs = this.testHooks ? opts.mediaHoldMs : undefined
    this.mediaWedge = this.testHooks ? opts.mediaWedge : undefined
    this.mediaDeadlineMs = (this.testHooks ? opts.mediaDeadlineMs : undefined) ?? MEDIA_JOB_DEADLINE_MS
    this.prepCapMs = (this.testHooks ? opts.prepDeadlineCapMsForTests : undefined) ?? PREP_DEADLINE_CAP_MS
    this.prepHangMs = this.testHooks ? opts.prepHangMsForTests : undefined
    this.startOptions = this.testHooks ? { ...(opts.workerFactoryForTests ? { create: opts.workerFactoryForTests } : {}), ...(opts.workerLoadBoundMsForTests ? { boundMs: opts.workerLoadBoundMsForTests } : {}) } : {}
    for (let i = 0; i < (opts.workers ?? defaultWorkers()); i++) {
      const slot = { index: i, busy: false, queue: [], pending: new Map(), docs: new Map(), pages: 0, parses: 0, probes: 0, generation: 0, closed: false, prep: 0 } as unknown as Slot
      this.attach(slot)
      this.slots.push(slot)
    }
  }

  private static failure(error: string): WorkerReply { return { id: -1, ok: false, error, stats: { parses: 0, pages: 0, docs: 0, probes: 0 } } }

  /** Settles every in-flight post on `slot` with a failure reply. */
  private settle(slot: Slot, error: string) {
    const pending = [...slot.pending.values()]
    slot.pending.clear()
    for (const p of pending) p.ok(Pool.failure(error))
  }

  /** Starts a worker for `slot`; replies resolve the slot's pending map; an
   * unexpected death settles its posts (nothing waits forever). */
  private attach(slot: Slot): void {
    // The events are wired below; `worker` is assigned before any of them can fire (they fire on later turns).
    const on = { message: (_e: MessageEvent) => {}, close: () => {} }
    const worker: StartedWorker = startWorker(WORKER_URL, { preload: [PRELOAD] } as WorkerOptions, { message: (e) => on.message(e), close: () => on.close() }, this.startOptions)
    slot.worker = worker
    slot.closed = false
    let markClosed!: () => void
    slot.whenClosed = new Promise<void>((r) => { markClosed = r })
    // Task 13 (carried from Task 3; fix rounds 1–2): this worker's file prefix (every file it stages bears it; swept
    // when it goes), and its probes, which run on this thread (`probe-host.ts`): killed when it goes, reaped here.
    const prefix = this.media ? `w${randomBytes(6).toString('hex')}-` : undefined
    // Released from the registry when retired with no run left (fix round 3), never by a stats() read.
    const probes: WorkerProbes | undefined = this.media && prefix ? createWorkerProbes(this.media.info, this.media.dir, prefix, (r) => worker.post(r), () => { if (probes) this.probeSets.delete(probes) }) : undefined
    if (probes) this.probeSets.add(probes)
    slot.retire = () => probes?.retire()
    on.message = (e: MessageEvent<WorkerReply | ProbeRunRequest | ProbeKillRequest>) => {
      const d = e.data
      if ('kind' in d) {
        // A request after its worker was retired is refused there (spawns nothing); media off has no probes at all.
        if (d.kind === 'probe-run') probes?.request(d); else if (d.kind === 'probe-kill') probes?.kill(d.id)
        return
      }
      if (slot.worker !== worker) return
      const p = slot.pending.get(d.id)
      if (!p) return
      slot.pending.delete(d.id)
      slot.pages = d.stats.pages; slot.parses = d.stats.parses; slot.probes = d.stats.probes
      p.ok(d)
    }
    // The init goes first (messages are ordered), so every later job sees the media config. Its reply is unawaited here
    // (the start gate takes it as the load signal); a worker whose init fails has media off (every entry
    // `media-unavailable`).
    worker.post({ type: 'init', id: this.nextId++, ...(this.media && prefix ? { media: { ...this.media, prefix } } : {}), ...(this.mediaHoldMs ? { mediaHoldMs: this.mediaHoldMs } : {}), ...(this.mediaWedge ? { testWedge: this.mediaWedge } : {}) } satisfies WorkerRequest)
    on.close = () => {
      // First, before anything awaiting the close (`replace`, a dying job's refund) resumes: its probes are killed
      // (their exits are reaped on this thread) and its staged files deleted.
      probes?.retire()
      if (this.media && prefix) sweepPrefixed(this.media.dir, prefix)
      markClosed()
      if (slot.worker !== worker) return
      slot.closed = true
      slot.docs.clear()
      this.settle(slot, 'worker died')
    }
  }

  /** Every worker's probe handles, until a retired worker's last run has settled (its host removes itself). */
  private probeSets = new Set<WorkerProbes>()
  /** Read-only. */
  private probesRunning(): number {
    let n = 0
    for (const p of this.probeSets) n += p.running
    return n
  }
  /** TEST-ONLY, side-effect free: how many probe hosts the registry holds. */
  probeHostsForTests(): number { return this.probeSets.size }

  /** Runs `fn` when the slot is idle; one in-flight message per worker. */
  private async exclusive<T>(slot: Slot, fn: () => Promise<T>): Promise<T> {
    if (slot.busy) await new Promise<void>((r) => slot.queue.push(r))
    slot.busy = true
    try {
      if (this.disposed) throw new RenderFailed('pool disposed')
      if (slot.closed) await this.replace(slot)
      // A worker still waiting its turn at the start gate, or loading, has not begun the job: its deadline starts after.
      // A loaded worker costs no extra turn (the job is posted in the same step as before).
      if (!slot.worker.loaded) await this.startedFor(slot)
      return await fn()
    } finally {
      slot.busy = false
      slot.queue.shift()?.()
    }
  }

  /**
   * Round 1 (R1, R3): waits for the slot's worker to load, inside the slot's exclusive turn, and re-checks the pool and
   * the slot after every await. Startup-failure policy (see `worker-start.ts`):
   * - closed while loading (or a construction that threw): one fresh worker; if that one does not load either, the
   *   job fails with `RenderFailed('worker failed to start')` and nothing is posted (the slot is replaced again by the
   *   next job, since it is closed);
   * - stalled past the start bound: the job fails with `RenderFailed('worker start timed out')`, and so does every job
   *   until the worker answers (then the slot works) or closes (then it is replaced). The stalled worker is never
   *   terminated while it loads; a dispose's terminate waits for its answer or close;
   * - disposed meanwhile: `RenderFailed('pool disposed')`.
   */
  private async startedFor(slot: Slot): Promise<void> {
    for (let attempt = 0; ; attempt++) {
      const r = await slot.worker.ready
      if (this.disposed) throw new RenderFailed('pool disposed')
      if (slot.worker.loaded && !slot.closed) return
      if (r === 'stalled' && !slot.closed) throw new RenderFailed('worker start timed out')
      // Closed while loading.
      if (attempt >= 1) throw new RenderFailed('worker failed to start')
      await this.replace(slot)
      if (this.disposed) throw new RenderFailed('pool disposed')
    }
  }

  /** Posts and waits at most `ms`; on timeout the worker is retired (the
   * caller learns via 'timeout'). A closed worker is sent nothing: the reply is the death failure at once, with no
   * pending entry (Round 1, R3). */
  private async post(slot: Slot, msg: WorkerRequest, ms?: number): Promise<WorkerReply | 'timeout'> {
    if (slot.closed) return Pool.failure('worker died')
    const reply = new Promise<WorkerReply>((ok) => { slot.pending.set(msg.id, { ok }); slot.worker.post(msg) })
    if (ms === undefined) return reply
    let timer: ReturnType<typeof setTimeout> | undefined
    const timeout = new Promise<'timeout'>((r) => { timer = setTimeout(() => r('timeout'), ms) })
    try { return await Promise.race([reply, timeout]) } finally { clearTimeout(timer) }
  }

  private pick(docRef: string): Slot {
    const known = this.affinity.get(docRef)
    if (known !== undefined) return this.slots[known]!
    const slot = [...this.slots].sort((a, b) => (a.docs.size - b.docs.size) || (Number(a.busy) - Number(b.busy)) || a.index - b.index)[0]!
    this.affinity.set(docRef, slot.index)
    return slot
  }

  /** Opens (parses) under its own deadline; a wedged parse retires the worker. */
  /** `sheetMode`: a re-open (after a worker recycle) is `'tiles'`: an XLSX entry's paintables come back lazily on its
   * first units use (R5). Only `open` passes the caller's mode. */
  private async ensureOpen(slot: Slot, docRef: string, policy: 'visible-only' | 'full', force = false, sheetMode: SheetMode = 'tiles'): Promise<{ docKey: string; open?: OpenResult }> {
    const k = `${docRef}|${policy}`, docKey = slot.docs.get(k)
    if (docKey && !force) return { docKey }
    const bytes = this.sources.get(docRef)
    if (!bytes) throw new RenderFailed('document not registered')
    const reply = await this.post(slot, { type: 'open', id: this.nextId++, docRef, bytes, contentPolicy: policy, sheetMode, ...(this.servedCacheBytes !== undefined ? { servedCacheBytes: this.servedCacheBytes } : {}) }, this.openDeadlineMs)
    if (reply === 'timeout') { await this.replace(slot); this.retired++; throw new RenderTimeout() }
    if (!reply.ok) throw new RenderFailed(reply.error, reply.refusal)
    const open = reply.result as OpenResult
    slot.docs.set(k, open.docKey)
    // Ruling B1: a tiles-mode open advertises an over-limit unit as the refusal unit; the prepare grants (units-mode
    // text ops) keep the true sizes, so a tiles (re)open never replaces sizes a units open recorded.
    if (sheetMode === 'units' || !this.unitSizes.has(k)) this.unitSizes.set(k, open.units)
    return { docKey: open.docKey, open }
  }

  /** Registers the bytes and parses on the docRef's worker. */
  async open(docRef: string, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full', o: { sheetMode?: SheetMode } = {}): Promise<OpenResult> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    this.sources.set(docRef, bytes)
    this.dropShowRefusals(docRef)
    const slot = this.pick(docRef)
    return this.exclusive(slot, async () => (await this.ensureOpen(slot, docRef, contentPolicy, true, o.sheetMode ?? 'units')).open!)
  }

  /** Admits, waits for the worker, renders under the deadline. Every exit
   * path refunds exactly once: here for normal ends, in `renderOn` (after the
   * worker is gone) for a retired job, in `admission.cancel` for a queued job
   * that never started. */
  async render(sessionId: string, job: RenderJob, bytesEstimate: number, signal?: AbortSignal): Promise<RenderedUnit> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, () => {
        // Admitted, then waited for the worker: a socket that closed meanwhile has no use for the render.
        if (signal?.aborted) throw new ShowCancelled()
        return this.renderOn(slot, admitted, job, signal)
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** A notes job (Stage 3c): the same affinity, admission, deadline and recycle paths as `render`. */
  async notes(sessionId: string, job: PoolNotesJob, bytesEstimate: number, signal?: AbortSignal): Promise<NotesPlan | NotesRaster> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
        if (signal?.aborted) throw new ShowCancelled()
        const r = await this.runJob(slot, admitted, { type: 'notes', id: this.nextId++, docKey, unitIndex: job.unitIndex, scale: job.scale, mark: job.mark, draw: job.draw })
        if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        return r as NotesPlan | NotesRaster
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** Stage 4a (F23): one sheet's layout, with the same affinity, admission, deadline and recycle paths as `notes`. */
  async sheetLayout(sessionId: string, job: PoolLayoutJob, bytesEstimate = LAYOUT_JOB_BYTES, signal?: AbortSignal): Promise<StreamSheetLayout> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
        if (signal?.aborted) throw new ShowCancelled()
        const r = await this.runJob(slot, admitted, { type: 'layout', id: this.nextId++, docKey, sheetIndex: job.sheetIndex, clock: { day: job.clock.day, timeZone: job.clock.timeZone } })
        if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        return r as StreamSheetLayout
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** Stage 4b: one `sheet-text` op (find, copy, a11y, fit) over a sheet's text view, with the same affinity, admission,
   * deadline and recycle paths as `sheetLayout`. Every job field but the routing is forwarded. Stage 4c: a job under a
   * view is charged `bytesEstimate + orderBytes(view)` (`viewOrderBytes`). */
  async sheetText(sessionId: string, job: PoolSheetTextJob, bytesEstimate = SHEET_TEXT_JOB_BYTES, signal?: AbortSignal): Promise<SheetTextReply> {
    const { docRef, contentPolicy, clock, ...text } = job
    const admitted = await this.admit(sessionId, bytesEstimate + viewOrderBytes(job.view), signal)
    const slot = this.pick(docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const { docKey } = await this.ensureOpen(slot, docRef, contentPolicy)
        if (signal?.aborted) throw new ShowCancelled()
        const r = await this.runJob(slot, admitted, { type: 'sheet-text', id: this.nextId++, docKey, clock: { day: clock.day, timeZone: clock.timeZone }, job: text as SheetTextJob })
        if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        return r as SheetTextReply
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** Stage 4c: one `sheet-view` job, with the same affinity, admission, deadline and recycle paths as `sheetText`.
   * `bytesEstimate` defaults to `sheetViewJobBytes`. Task 14a: a job that reads the sheet's port runs as two worker
   * jobs under the one admission and slot hold: the prep (`sheet-view-prep`: snapshot, key view, port; a warm port
   * answers at once) under `sheetPrepDeadlineMs` of the document's admitted bytes, then the op under the render
   * deadline. A prep past its deadline, or a failed prep, throws `SheetPrepRefused`; an abort during prep runs no op. */
  async sheetView(sessionId: string, job: PoolSheetViewJob, bytesEstimate?: number, signal?: AbortSignal): Promise<SheetViewReply> {
    const { docRef, contentPolicy, clock, ...view } = job
    const admitted = await this.admit(sessionId, bytesEstimate ?? sheetViewJobBytes(view as SheetViewJob), signal)
    const slot = this.pick(docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const { docKey } = await this.ensureOpen(slot, docRef, contentPolicy)
        if (signal?.aborted) throw new ShowCancelled()
        if (readsPort(view as SheetViewJob)) {
          await this.sheetPrep(slot, admitted, docRef, { type: 'sheet-view-prep', id: this.nextId++, docKey, clock: { day: clock.day, timeZone: clock.timeZone }, contentPolicy, sheetIndex: view.sheetIndex, ...(this.prepHangMs ? { hangMs: this.prepHangMs } : {}) })
          if (signal?.aborted) throw new ShowCancelled()
        }
        const r = await this.runJob(slot, admitted, { type: 'sheet-view', id: this.nextId++, docKey, clock: { day: clock.day, timeZone: clock.timeZone }, contentPolicy, job: view as SheetViewJob })
        if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        return r as SheetViewReply
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** Task 14a: the prep job under its own deadline. A timeout (the worker is retired, as for any job) or a failure the
   * worker reports is `SheetPrepRefused`; a worker refusal code (`unknown-doc`, `bad-job`) stays `RenderFailed`. */
  private async sheetPrep(slot: Slot, admitted: Job, docRef: string, msg: WorkerRequest): Promise<void> {
    try {
      await this.runJob(slot, admitted, msg, sheetPrepDeadlineMs(this.sources.get(docRef)?.length ?? 0, this.admission.caps.renderDeadlineMs, this.prepCapMs))
    } catch (e) {
      if (e instanceof RenderTimeout) throw new SheetPrepRefused('timeout')
      if (e instanceof RenderFailed && e.refusal === undefined && !this.disposed) throw new SheetPrepRefused('failed')
      throw e
    }
  }

  /** Stage 4a (S4 §6.5): one canonical sheet tile, with the same affinity, admission, deadline and recycle paths as
   * `render` (`bytesEstimate`: `TILE_BYTES_ESTIMATE`; Stage 4c: plus `orderBytes(view)` under a view). The result carries the
   * job's palette (F24). */
  async tile(sessionId: string, job: PoolTileJob, bytesEstimate: number, signal?: AbortSignal): Promise<RenderedTile> {
    const admitted = await this.admit(sessionId, bytesEstimate + viewOrderBytes(job.view), signal)
    const { docRef, contentPolicy, hangMs, crash, ...rest } = job
    const tile: TileJob = { sheetIndex: rest.sheetIndex, viewGen: rest.viewGen, zoomStep: rest.zoomStep, dpr: rest.dpr, palette: rest.palette, col: rest.col, row: rest.row, mark: rest.mark, clock: rest.clock && { day: rest.clock.day, timeZone: rest.clock.timeZone }, ...(rest.sizes !== undefined ? { sizes: rest.sizes } : {}), ...(rest.view !== undefined ? { view: rest.view } : {}) }
    const slot = this.pick(docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const { docKey } = await this.ensureOpen(slot, docRef, contentPolicy)
        // Again after the (re)open: a socket that closed while the document was parsed gets no tile.
        if (signal?.aborted) throw new ShowCancelled()
        const r = await this.runJob(slot, admitted, { type: 'tile', id: this.nextId++, docKey, job: tile, ...this.hooks(hangMs, crash) })
        if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        return r as RenderedTile
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** Same admission, deadline, crash and retire paths as `render`, on the
   * least-loaded worker. `bytesEstimate` = w·h·4·3 (decode + marked copy +
   * encode buffers). The PNG is copied to the worker, never transferred: the
   * caller's bytes are the shared canonical cache entry. */
  async encode(sessionId: string, job: EncodeJob, bytesEstimate: number, signal?: AbortSignal): Promise<MarkRasterResult> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = [...this.slots].sort((a, b) => (Number(a.busy) - Number(b.busy)) || (a.queue.length - b.queue.length) || a.index - b.index)[0]!
    try {
      return await this.exclusive(slot, () => {
        if (signal?.aborted) throw new ShowCancelled()
        const msg: WorkerRequest = { type: 'encode', id: this.nextId++, png: job.png, w: job.w, h: job.h, kBasis: job.kBasis, code: job.code, ...(this.testHooks && job.hangMs ? { hangMs: job.hangMs } : {}), ...(this.testHooks && job.crash ? { crash: true } : {}) }
        this.encodes++
        return this.runJob(slot, admitted, msg) as Promise<MarkRasterResult>
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /**
   * Piece 4: one slide's show-plan bundle, with the same document affinity, admission and recycle paths as
   * `render`, under `SHOW_PLAN_DEADLINE_MS` (10 s, against 2 s for units). A bundle past the deadline retires
   * its worker and plays its still: `{ refused: 'show-plan-timeout' }`.
   */
  async showPlan(sessionId: string, job: PoolShowJob, bytesEstimate: number, signal?: AbortSignal): Promise<BuiltShowPlan | ShowRefusal> {
    const { docRef, contentPolicy, hangMs, crash, ...show } = job
    const key = [docRef, contentPolicy, show.unitIndex, show.scale, show.motion, show.strictMarks, show.media === true].join('|')
    const known = this.showRefusals.get(key)
    if (known) return known
    let r: BuiltShowPlan | ShowRefusal
    try {
      r = await this.showJob(sessionId, docRef, contentPolicy, bytesEstimate, (docKey) => ({ type: 'show-plan', id: this.nextId++, docKey, job: show, grant: bytesEstimate, ...this.hooks(hangMs, crash) }), signal, [show.unitIndex]) as BuiltShowPlan | ShowRefusal
    } catch (e) {
      // A warm-up timeout answers this request only (Task 8): never remembered.
      if (e instanceof MediaWarmTimeout) return { refused: 'show-plan-timeout', diagnostics: [{ code: 'show-plan-timeout' }] }
      if (!(e instanceof RenderTimeout)) throw e
      r = { refused: 'show-plan-timeout', diagnostics: [{ code: 'show-plan-timeout' }] }
    }
    if ('refused' in r && this.sources.has(docRef)) this.showRefusals.set(key, r)
    return r
  }

  /** Piece 4: the keyframe track of one composited step (affinity, admission, `SHOW_PLAN_DEADLINE_MS`; a timeout throws `RenderTimeout`). */
  async showTrack(sessionId: string, job: PoolShowJob & { step: number }, bytesEstimate: number, signal?: AbortSignal): Promise<Array<{ assetId: number; w: number; h: number; png: Uint8Array }>> {
    const { docRef, contentPolicy, hangMs, crash, ...show } = job
    const plan = [docRef, contentPolicy, show.unitIndex, show.scale, show.motion, show.strictMarks, show.media === true].join('|')
    // A slide whose bundle is refused has no track.
    if (this.showRefusals.has(plan)) return []
    return await this.unwarmed(() => this.remembered(docRef, `${plan}|track|${show.step}`, async () => await this.showJob(sessionId, docRef, contentPolicy, bytesEstimate, (docKey) => ({ type: 'show-track', id: this.nextId++, docKey, job: show, grant: bytesEstimate, ...this.hooks(hangMs, crash) }), signal, [show.unitIndex]) as Array<{ assetId: number; w: number; h: number; png: Uint8Array }>))
  }

  /** Piece 4 Task 7: one Morph pair (affinity, admission, `SHOW_PLAN_DEADLINE_MS`; a timeout throws `RenderTimeout`). */
  async showMorph(sessionId: string, job: PoolMorphJob, bytesEstimate: number, signal?: AbortSignal): Promise<BuiltMorph> {
    const { docRef, contentPolicy, hangMs, crash, ...morph } = job
    const key = [docRef, contentPolicy, 'morph', morph.from, morph.to, morph.fromStep, morph.scale, morph.motion, morph.strictMarks === true].join('|')
    return await this.unwarmed(() => this.remembered(docRef, key, async () => await this.showJob(sessionId, docRef, contentPolicy, bytesEstimate, (docKey) => ({ type: 'show-morph', id: this.nextId++, docKey, job: morph, grant: bytesEstimate, ...this.hooks(hangMs, crash) }), signal, [morph.from, morph.to]) as BuiltMorph))
  }

  /**
   * Piece 4 D7: a layer's per-session Layer-1 encode. It reads no document, so it runs on the least-loaded
   * worker as `encode` does, with the same admission, under `SHOW_PLAN_DEADLINE_MS`.
   */
  async markLayer(sessionId: string, job: MarkLayerJob, bytesEstimate: number, signal?: AbortSignal): Promise<MarkLayerResult> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = [...this.slots].sort((a, b) => (Number(a.busy) - Number(b.busy)) || (a.queue.length - b.queue.length) || a.index - b.index)[0]!
    try {
      return await this.exclusive(slot, () => {
        if (signal?.aborted) throw new ShowCancelled()
        const msg: WorkerRequest = { type: 'mark-layer', id: this.nextId++, png: job.png, w: job.w, h: job.h, kBasis: job.kBasis, code: job.code, crop: job.crop, ...this.hooks(job.hangMs, job.crash) }
        return this.runJob(slot, admitted, msg, this.showDeadlineMs) as Promise<MarkLayerResult>
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /**
   * Piece 5 (D21, D22): unit `unit`'s audio uses and their probed entries, from the document's table on its affinity
   * worker (one table per document). A short read job admitted like `text`; not an encode.
   */
  async mediaRefs(sessionId: string, docRef: string, contentPolicy: 'visible-only' | 'full', unit: number, signal?: AbortSignal): Promise<MediaUnitRefs> {
    return await this.mediaJob(sessionId, docRef, contentPolicy, (docKey) => ({ type: 'media-refs', id: this.nextId++, docKey, unit }), signal) as MediaUnitRefs
  }

  /**
   * Piece 5 (D21): the worker writes the projected part to a 0600 file in `dir` (the media temp dir), so part bytes
   * never cross the worker boundary; the caller adopts it into its `MediaTemp` ledger. A short read job like `text`.
   */
  async mediaPart(sessionId: string, docRef: string, contentPolicy: 'visible-only' | 'full', partPath: string, dir: string, signal?: AbortSignal): Promise<{ path: string; bytes: number; sourceHash: string }> {
    // Task 13: the pool names the file, so a job with no reply (a timeout or death retires the worker first) leaves none.
    const name = `${randomBytes(16).toString('hex')}.bin`
    try {
      return await this.mediaJob(sessionId, docRef, contentPolicy, (docKey) => ({ type: 'media-part', id: this.nextId++, docKey, partPath, dir, name }), signal) as { path: string; bytes: number; sourceHash: string }
    } catch (e) {
      try { unlinkSync(join(dir, name)) } catch { /* never written, or removed by the worker */ }
      throw e
    }
  }

  private async mediaJob(sessionId: string, docRef: string, contentPolicy: 'visible-only' | 'full', msg: (docKey: string) => WorkerRequest, signal?: AbortSignal): Promise<unknown> {
    const admitted = await this.admit(sessionId, TEXT_JOB_BYTES, signal)
    const slot = this.pick(docRef)
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const generation = slot.generation
        try {
          const { docKey } = await this.ensureOpen(slot, docRef, contentPolicy)
          if (signal?.aborted) throw new ShowCancelled()
          return await this.runJob(slot, admitted, msg(docKey), this.mediaDeadlineMs)
        } finally {
          if (slot.generation === generation && !slot.closed && slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        }
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** A warm-up timeout passes `remembered` uncached (it remembers only `RenderTimeout`); the caller sees a timeout. */
  private async unwarmed<T>(run: () => Promise<T>): Promise<T> {
    try { return await run() } catch (e) { if (e instanceof MediaWarmTimeout) throw new RenderTimeout(); throw e }
  }

  private hooks(hangMs?: number, crash?: boolean) { return { ...(this.testHooks && hangMs ? { hangMs } : {}), ...(this.testHooks && crash ? { crash: true } : {}) } }

  /** A show job on the docRef's worker: admission, affinity, the show deadline, the recycle check. */
  /**
   * Admission for one job of `sessionId`: refused, started, or queued until it starts. `signal` (the requesting
   * socket's): aborted while queued, the job leaves the queue unstarted (nothing charged) and `ShowCancelled` is
   * thrown; aborted before the call, nothing is requested (PR #78 rounds 4–5: a closed page's backlog never runs).
   */
  private async admit(sessionId: string, bytesEstimate: number, signal?: AbortSignal): Promise<Job> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    if (signal?.aborted) throw new ShowCancelled()
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    const admitted = res.job
    if (res.status === 'admitted') return admitted
    let onAbort: (() => void) | undefined
    const started = await new Promise<boolean>((r) => {
      this.waiters.set(admitted.id, r)
      onAbort = () => { if (this.waiters.get(admitted.id) === r && this.admission.cancel(admitted.id)) { this.waiters.delete(admitted.id); r(false) } }
      signal?.addEventListener('abort', onAbort, { once: true })
    })
    if (onAbort) signal?.removeEventListener('abort', onAbort)
    if (!started) throw this.disposed ? new RenderFailed('pool disposed') : new ShowCancelled()
    return admitted
  }

  /**
   * Owner ruling (PR #78 round 13b, reserve per worker): `grant` is the job's preparation grant. The session is
   * admitted only for `showRequestBytes(grant)` (its queue entry and encoded reply); the grant itself is reserved
   * when the job holds its worker's slot (`exclusive`: one job per worker) and released when it ends, so preparation
   * never holds more than workers × `SHOW_BUILD_BYTES` (`stats().prep`).
   */
  private async showJob(sessionId: string, docRef: string, contentPolicy: 'visible-only' | 'full', grant: number, msg: (docKey: string) => WorkerRequest, signal?: AbortSignal, warm?: readonly number[]): Promise<unknown> {
    // Round 13c: one running show job per session (its session slot first, then admission, then its worker's), so a
    // viewer holds one grant at most whatever the worker count. Round 13d: a job waiting for the session slot holds
    // no general admission (no running place, no bytes), so the session's unit renders are not held behind it.
    const leave = await this.sessionShowSlot(sessionId, signal)
    try {
      const admitted = await this.admit(sessionId, showRequestBytes(grant), signal)
      try {
        const slot = this.pick(docRef)
        return await this.exclusive(slot, async () => {
          // Admitted, then waited for the worker: a socket that closed meanwhile has no use for the build.
          if (signal?.aborted) throw new ShowCancelled()
          const held = this.reservePrep(slot, grant, sessionId)
          try { return await this.showOnSlot(slot, admitted, docRef, contentPolicy, msg, signal, warm) } finally { this.releasePrep(slot, held, sessionId) }
        })
      } finally {
        if (admitted.state === 'running') this.finish(admitted)
      }
    } finally { leave() }
  }

  /**
   * Round 13c: the session's show slot (one running show job per session). Resolves with its release, which hands
   * the slot to the session's next waiter in order. `signal` aborted while waiting: the waiter leaves at once and
   * `ShowCancelled` is thrown (it never took the slot).
   */
  private async sessionShowSlot(sessionId: string, signal?: AbortSignal): Promise<() => void> {
    if (signal?.aborted) throw new ShowCancelled()
    let q = this.showSessions.get(sessionId)
    if (!q) this.showSessions.set(sessionId, q = { busy: false, waiters: [] })
    const s = q
    let released = false
    const release = () => {
      if (released) return
      released = true
      const next = s.waiters.shift()
      if (next) next()
      else { s.busy = false; this.showSessions.delete(sessionId) }
    }
    if (!s.busy) { s.busy = true; return release }
    // Round 13d: bounded so a session keeps as many show requests in flight as general admission allowed it before
    // (its running places plus its queue: 8 + 32), and refused the same way (`session-queue`).
    if (s.waiters.length >= showQueueBound(this.admission.caps)) { this.admission.refusals['session-queue']++; throw new AdmissionRefused('session-queue') }
    await new Promise<void>((ok, fail) => {
      const onAbort = () => {
        const i = s.waiters.indexOf(wake)
        if (i >= 0) s.waiters.splice(i, 1)
        fail(new ShowCancelled())
      }
      const wake = () => { signal?.removeEventListener('abort', onAbort); ok() }
      s.waiters.push(wake)
      signal?.addEventListener('abort', onAbort, { once: true })
    })
    return release
  }

  /** The worker's preparation reservation for the job on its slot (at most `SHOW_BUILD_BYTES`; one per worker). */
  private reservePrep(slot: Slot, grant: number, sessionId: string): number {
    const bytes = Math.max(0, Math.min(SHOW_BUILD_BYTES, Number.isFinite(grant) ? grant : SHOW_BUILD_BYTES))
    if (slot.prep !== 0) throw new Error('pool: a worker holds two preparation grants')
    slot.prep = bytes
    this.prep.reserved += bytes
    if (this.prep.reserved > this.slots.length * SHOW_BUILD_BYTES) throw new Error('pool: preparation past workers × SHOW_BUILD_BYTES')
    this.prep.peak = Math.max(this.prep.peak, this.prep.reserved)
    this.prep.peakPerWorker = Math.max(this.prep.peakPerWorker, bytes)
    const mine = (this.prepBySession.get(sessionId) ?? 0) + bytes
    if (mine > SHOW_BUILD_BYTES) throw new Error('pool: a session past one preparation grant')
    this.prepBySession.set(sessionId, mine)
    this.prep.peakPerSession = Math.max(this.prep.peakPerSession, mine)
    return bytes
  }
  private releasePrep(slot: Slot, bytes: number, sessionId: string): void {
    slot.prep = 0
    this.prep.reserved -= bytes
    const mine = (this.prepBySession.get(sessionId) ?? 0) - bytes
    if (mine > 0) this.prepBySession.set(sessionId, mine); else this.prepBySession.delete(sessionId)
  }

  private async showOnSlot(slot: Slot, admitted: Job, docRef: string, contentPolicy: 'visible-only' | 'full', msg: (docKey: string) => WorkerRequest, signal?: AbortSignal, warm?: readonly number[]): Promise<unknown> {
    const { docKey } = await this.ensureOpen(slot, docRef, contentPolicy)
    // Again after the (re)open: a socket that closed while the document was parsed gets no build.
    if (signal?.aborted) throw new ShowCancelled()
    // Piece 5 (Task 8; the deadline risk carried from Task 3): the build's `loadShow` probes the slide's media parts
    // (up to the 10 s deck budget), which would share the 10 s show deadline. So, when the pool has media, the probes
    // run first in this same turn (same session slot, admission, worker and open document), as a media read job under
    // the media deadline; the build then finds the table warm and the show deadline covers it alone. A warm-up
    // timeout (the worker is retired) is `MediaWarmTimeout`, never remembered. A refusal (unit range, a non-show
    // document) is left to the build to report. Without media nothing runs (piece 4 unchanged).
    if (this.media && warm) {
      for (const unit of warm) {
        try { await this.runJob(slot, admitted, { type: 'media-refs', id: this.nextId++, docKey, unit }, this.mediaDeadlineMs) } catch (e) {
          if (e instanceof RenderTimeout) throw new MediaWarmTimeout()
          if (!(e instanceof RenderFailed && e.refusal !== undefined)) throw e
        }
        if (signal?.aborted) throw new ShowCancelled()
      }
    }
    const r = await this.runJob(slot, admitted, msg(docKey), this.showDeadlineMs)
    if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
    return r
  }


  /** A text op on the docRef's worker (affinity, admission, deadline as
   * `render`). The op is charged `bytesEstimate`; every prepare (explicit or
   * for a `{ missing }` answer) is charged `prepareWorkspace(w, h)` from the
   * unit's parsed size through the same admission, and the worker is sent
   * that grant to enforce. A missing answer prepares each missing unit (one
   * job each, render deadline) and retries once; still missing ⇒
   * TextTooLarge. Everything runs in one exclusive turn of the slot, and the
   * recycle threshold is checked on every exit, success or not. */
  async text(sessionId: string, job: PoolTextJob, bytesEstimate = TEXT_JOB_BYTES, signal?: AbortSignal): Promise<TextReply> {
    const admitted = await this.admit(sessionId, bytesEstimate, signal)
    const slot = this.pick(job.docRef)
    const { docRef: _d, contentPolicy: _c, ...rest } = job
    const op = rest as TextJob
    const k = `${job.docRef}|${job.contentPolicy}`
    try {
      return await this.exclusive(slot, async () => {
        if (signal?.aborted) throw new ShowCancelled()
        const generation = slot.generation
        try {
          const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
          // Checked before every post (after the reopen, and between a prepare and its retry): a closed socket's
          // text op stops at the next await.
          const post = (j: TextJob, owner: Job, grant?: number) => {
            if (signal?.aborted) return Promise.reject(new ShowCancelled())
            return this.runJob(slot, owner, { type: 'text', id: this.nextId++, docKey, job: j, ...(grant !== undefined ? { grant } : {}) }) as Promise<TextReply>
          }
          const prepare = (unit: number) => this.prepareOn(sessionId, k, unit, (owner, grant) => post({ op: 'prepare', unit }, owner, grant))
          if (op.op === 'prepare') return await prepare(op.unit)
          let r = await post(op, admitted)
          if ('missing' in r) {
            for (const unit of r.missing) await prepare(unit)
            r = await post(op, admitted)
            if ('missing' in r) throw new TextTooLarge(r.missing)
          }
          return r
        } finally {
          // Only the worker these jobs ran on; a timeout/death already retired it.
          if (slot.generation === generation && !slot.closed && slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
        }
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** One prepare under its own admission charge, sized from the unit. Inside
   * the slot's exclusive turn a queued charge could wait on work that waits
   * for this slot, so a charge that cannot start at once is refused. */
  private async prepareOn(sessionId: string, k: string, unit: number, run: (owner: Job, grant: number) => Promise<TextReply>): Promise<TextReply> {
    const size = this.unitSizes.get(k)?.[unit]
    if (!size) throw new RenderFailed('unit-range', 'unit-range')
    const grant = prepareWorkspace(size.w, size.h)
    const res = this.admission.request(sessionId, grant, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    if (res.status === 'queued') { this.admission.cancel(res.job.id); throw new AdmissionRefused('session-queue') }
    try {
      return await run(res.job, grant)
    } finally {
      if (res.job.state === 'running') this.finish(res.job)
    }
  }

  private finish(job: Job): void {
    const started = this.admission.finish(job, this.now())
    if (this.disposed) return
    for (const j of started) { this.waiters.get(j.id)?.(true); this.waiters.delete(j.id) }
  }

  private async renderOn(slot: Slot, admitted: Job, job: RenderJob, signal?: AbortSignal): Promise<RenderedUnit> {
    const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
    // Again after the (re)open: a socket that closed while the document was parsed gets no render.
    if (signal?.aborted) throw new ShowCancelled()
    const msg: WorkerRequest = { type: 'render', id: this.nextId++, docKey, unitIndex: job.unitIndex, scale: job.scale, projection: job.projection, mark: job.mark, ...(job.missingFaces ? { missingFaces: job.missingFaces } : {}), ...(this.testHooks && job.hangMs ? { hangMs: job.hangMs } : {}), ...(this.testHooks && job.crash ? { crash: true } : {}) }
    const r = await this.runJob(slot, admitted, msg)
    if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
    return r as RenderedUnit
  }

  /** Posts one admitted job under the render deadline; owns the retire paths. */
  private async runJob(slot: Slot, admitted: Job, msg: WorkerRequest, deadlineMs = this.admission.caps.renderDeadlineMs): Promise<unknown> {
    // The deadline covers the paint only (the worker is idle when we post).
    admitted.startedAt = this.now()
    const reply = await this.post(slot, msg, deadlineMs + 1)
    if (reply === 'timeout') {
      // Retirement owner: only THIS job retires; its charge stays reserved
      // until the worker is gone, then is refunded once.
      admitted.state = 'retiring'; admitted.timedOut = true
      await this.replace(slot)
      this.retired++
      this.finish(admitted)
      throw new RenderTimeout()
    }
    if (!reply.ok && reply.id === -1) {
      // The worker is going away (dispose, death): the paint may still be
      // running, so the charge moves to a retirement owner and is refunded
      // only once the worker has closed (review r2-4).
      admitted.state = 'retiring'
      await slot.whenClosed
      this.finish(admitted)
      throw new RenderFailed(reply.error)
    }
    if (!reply.ok) throw new RenderFailed(reply.error, reply.refusal)
    return reply.result
  }

  /** Terminates and respawns a slot; its documents re-open on demand.
   * Resolves only after the old worker is gone (drain-then-refund, §10). */
  private async replace(slot: Slot): Promise<void> {
    const old = slot.worker
    // Its probes first: a request it posts from now on is refused, and its running probes are killed.
    slot.retire?.()
    this.settle(slot, 'worker retired')
    if (!slot.closed) {
      const closed = slot.whenClosed
      old.terminate()
      await closed
    }
    slot.docs.clear(); slot.pages = 0; slot.parses = 0; slot.probes = 0; slot.generation++
    if (!this.disposed) this.attach(slot)
  }

  /** Releases a parse (document-ledger eviction); the bytes stay, so a later
   * request re-parses on the same worker. */
  dropParsed(docRef: string, policy: 'visible-only' | 'full'): void {
    const i = this.affinity.get(docRef)
    if (i === undefined) return
    const slot = this.slots[i]!, k = `${docRef}|${policy}`, docKey = slot.docs.get(k)
    if (!docKey) return
    slot.docs.delete(k)
    void this.exclusive(slot, () => this.post(slot, { type: 'drop', id: this.nextId++, docKey })).catch(() => {})
  }

  /** Forgets a document everywhere (no session references it any more). */
  forget(docRef: string): void {
    this.dropParsed(docRef, 'visible-only'); this.dropParsed(docRef, 'full')
    // Size metadata goes with the document; a worker recycle keeps it (the
    // next ensureOpen re-parses and regenerates it), so does dropParsed.
    this.unitSizes.delete(`${docRef}|visible-only`); this.unitSizes.delete(`${docRef}|full`)
    this.sources.delete(docRef)
    this.affinity.delete(docRef)
    this.dropShowRefusals(docRef)
  }

  private dropShowRefusals(docRef: string) {
    for (const k of this.showRefusals.keys()) if (k.startsWith(`${docRef}|`)) this.showRefusals.delete(k)
    for (const k of this.jobRefusals.keys()) if (k.startsWith(`${docRef}|`)) this.jobRefusals.delete(k)
  }

  stats() {
    return {
      workers: this.slots.map((s) => ({ index: s.index, generation: s.generation, pages: s.pages, parses: s.parses, probes: s.probes, docs: s.docs.size, busy: s.busy, closed: s.closed })),
      reserved: this.admission.reserved, running: this.admission.runningCount, queued: this.admission.queued,
      prep: { ...this.prep, bound: this.slots.length * SHOW_BUILD_BYTES },
      retired: this.retired, recycled: this.recycled, encodes: this.encodes, affinity: Object.fromEntries(this.affinity), unitSizes: this.unitSizes.size,
      /** Fix round 2: probe runs spawned for workers and not yet exited (reaped), retired workers' included. */
      probesRunning: this.probesRunning(),
    }
  }

  /** Shuts down: queued jobs are cancelled (refunded, rejected), in-flight
   * posts settle as failures (their callers refund), slot waiters wake and
   * reject; no worker is respawned. The returned promise resolves once every slot's worker has fired `close` (a
   * terminate deferred by the start gate during a load included: it resolves when that load ends and the worker
   * closes); a later call returns the same promise. Callers that ignore it are unchanged. Nothing attaches a worker
   * once `disposed` is set (`replace` checks it after its close), so each slot's current `whenClosed` is its last. */
  dispose(): Promise<void> {
    if (this.disposedWait) return this.disposedWait
    this.disposed = true
    this.unitSizes.clear()
    for (const [id, wake] of this.waiters) { this.admission.cancel(id); wake(false) }
    this.waiters.clear()
    for (const s of this.slots) {
      this.settle(s, 'pool disposed')
      s.retire?.()
      s.worker.terminate()
      for (const wake of s.queue.splice(0)) wake()
    }
    return this.disposedWait = Promise.all(this.slots.map((s) => s.whenClosed)).then(() => {})
  }
}
