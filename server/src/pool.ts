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
import type { MarkSpec, MediaUnitRefs, NotesJob, NotesPlan, NotesRaster, OpenResult, Projection, RenderedUnit } from './render'
import type { FfmpegInfo } from './media/ffmpeg'
import { PROBE_DECK_MS, PROBE_PART_MS } from './media/probe'
import { createWorkerProbes, type ProbeKillRequest, type ProbeRunRequest, type WorkerProbes } from './media/probe-host'
import { sweepPrefixed } from './media/temp'
import type { WorkerReply, WorkerRequest } from './worker/render-worker'
import { prepareWorkspace, TEXT_PREPARE_ALLOWANCE, type TextJob, type TextReply } from './text/served'
import { startWorker, type StartedWorker } from './worker-start'

export class AdmissionRefused extends Error { constructor(readonly reason: 'session-queue' | 'process-queue' | 'too-large') { super(`admission refused: ${reason}`); this.name = 'AdmissionRefused' } }
export class RenderTimeout extends Error { constructor() { super('render deadline exceeded'); this.name = 'RenderTimeout' } }
export class RenderFailed extends Error { constructor(m: string, readonly refusal?: string) { super(m); this.name = 'RenderFailed' } }
/** Piece 5 (Task 8): a show job's media warm-up passed its own deadline. Never remembered as the slide's refusal. */
class MediaWarmTimeout extends Error { constructor() { super('media warm-up deadline exceeded'); this.name = 'MediaWarmTimeout' } }
/** A job whose socket closed before it reached a worker: nothing was run (PR #78 rounds 4–5: show, unit, notes, text). */
export class ShowCancelled extends RenderFailed { constructor() { super('cancelled: the requesting socket closed'); this.name = 'ShowCancelled' } }
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
/** `servedCacheBytes`: the worker's served-index bound per document (R6 default 32 MiB). */
export interface PoolOptions { /** Piece 4 (tests shorten it): the show-plan job deadline, default `SHOW_PLAN_DEADLINE_MS`. */ showDeadlineMs?: number; workers?: number; caps?: Caps; recyclePages?: number; now?: () => number; testHooks?: boolean; openDeadlineMs?: number; servedCacheBytes?: number; /** Piece 5 (D21): ffmpeg and the media temp dir, forwarded in each worker's init. */ media?: { info: FfmpegInfo; dir: string }; /** TEST-ONLY (honoured with `testHooks`): every media probe is held this long (`MediaTableSeams.hangMs`). */ mediaHoldMs?: number; /** TEST-ONLY (honoured with `testHooks`): the media warm-up and media job deadline, default `MEDIA_JOB_DEADLINE_MS`. */ mediaDeadlineMs?: number; /** TEST-ONLY (honoured with `testHooks`): `after-part`, `mediaPart` never answers once it wrote its file. */ mediaWedge?: 'after-part' }

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
    const worker: StartedWorker = startWorker(WORKER_URL, { preload: [PRELOAD] } as WorkerOptions, { message: (e) => on.message(e), close: () => on.close() })
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
      if (!slot.worker.loaded) {
        await slot.worker.ready
        if (this.disposed) throw new RenderFailed('pool disposed')
        // It died while loading: once more, on a fresh worker.
        if (slot.closed) { await this.replace(slot); await slot.worker.ready }
      }
      return await fn()
    } finally {
      slot.busy = false
      slot.queue.shift()?.()
    }
  }

  /** Posts and waits at most `ms`; on timeout the worker is retired (the
   * caller learns via 'timeout'). */
  private async post(slot: Slot, msg: WorkerRequest, ms?: number): Promise<WorkerReply | 'timeout'> {
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
  private async ensureOpen(slot: Slot, docRef: string, policy: 'visible-only' | 'full', force = false): Promise<{ docKey: string; open?: OpenResult }> {
    const k = `${docRef}|${policy}`, docKey = slot.docs.get(k)
    if (docKey && !force) return { docKey }
    const bytes = this.sources.get(docRef)
    if (!bytes) throw new RenderFailed('document not registered')
    const reply = await this.post(slot, { type: 'open', id: this.nextId++, docRef, bytes, contentPolicy: policy, ...(this.servedCacheBytes !== undefined ? { servedCacheBytes: this.servedCacheBytes } : {}) }, this.openDeadlineMs)
    if (reply === 'timeout') { await this.replace(slot); this.retired++; throw new RenderTimeout() }
    if (!reply.ok) throw new RenderFailed(reply.error, reply.refusal)
    const open = reply.result as OpenResult
    slot.docs.set(k, open.docKey)
    this.unitSizes.set(k, open.units)
    return { docKey: open.docKey, open }
  }

  /** Registers the bytes and parses on the docRef's worker. */
  async open(docRef: string, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full'): Promise<OpenResult> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    this.sources.set(docRef, bytes)
    this.dropShowRefusals(docRef)
    const slot = this.pick(docRef)
    return this.exclusive(slot, async () => (await this.ensureOpen(slot, docRef, contentPolicy, true)).open!)
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
   * reject; no worker is respawned. */
  dispose(): void {
    if (this.disposed) return
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
  }
}
