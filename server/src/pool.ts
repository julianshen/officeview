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
import { availableParallelism } from 'node:os'
import { resolve } from 'node:path'
import { AdmissionModel, DEFAULT_CAPS, type Caps, type Job } from './admission'
import type { MarkRasterResult } from './mark/mark-raster'
import type { MarkSpec, NotesJob, NotesPlan, NotesRaster, OpenResult, Projection, RenderedUnit } from './render'
import type { WorkerReply, WorkerRequest } from './worker/render-worker'
import { prepareWorkspace, TEXT_PREPARE_ALLOWANCE, type TextJob, type TextReply } from './text/served'

export class AdmissionRefused extends Error { constructor(readonly reason: 'session-queue' | 'process-queue' | 'too-large') { super(`admission refused: ${reason}`); this.name = 'AdmissionRefused' } }
export class RenderTimeout extends Error { constructor() { super('render deadline exceeded'); this.name = 'RenderTimeout' } }
export class RenderFailed extends Error { constructor(m: string, readonly refusal?: string) { super(m); this.name = 'RenderFailed' } }
/** A text op whose units do not all fit the worker's served-index cache at once (Task 5 maps it to copy-too-large). */
export class TextTooLarge extends Error { constructor(readonly missing: number[]) { super(`text range exceeds the served-index cache (${missing.length} units missing)`); this.name = 'TextTooLarge' } }

export interface RenderJob { docRef: string; contentPolicy: 'visible-only' | 'full'; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number; crash?: boolean }
/** A notes plan or draw (Stage 3c) on the docRef's worker. */
export type PoolNotesJob = { docRef: string; contentPolicy: 'visible-only' | 'full' } & NotesJob
/** A per-session Layer-1 encode (§9.4); no document affinity. */
export interface EncodeJob { png: Uint8Array; w: number; h: number; kBasis: Uint8Array; code: Int8Array; hangMs?: number; crash?: boolean }
export type PoolTextJob = { docRef: string; contentPolicy: 'visible-only' | 'full' } & TextJob
/** Admission charge of one text op itself (index assembly, answer text);
 * every prepare it needs is charged separately (`prepareWorkspace`). */
export const TEXT_JOB_BYTES = TEXT_PREPARE_ALLOWANCE
/** `servedCacheBytes`: the worker's served-index bound per document (R6 default 32 MiB). */
export interface PoolOptions { workers?: number; caps?: Caps; recyclePages?: number; now?: () => number; testHooks?: boolean; openDeadlineMs?: number; servedCacheBytes?: number }

const WORKER_URL = resolve(import.meta.dir, 'worker/render-worker.ts')
const PRELOAD = resolve(import.meta.dir, '../../scripts/stream-s0/napi-preload.ts')
export const defaultWorkers = () => Math.max(1, Math.min(8, Math.floor(availableParallelism() / 2)))

interface Slot {
  index: number
  worker: Worker
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
  generation: number
}

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
  /** `${docRef}|${policy}` → unit sizes from the last parse (prepare charges). */
  private unitSizes = new Map<string, OpenResult['units']>()
  private disposed = false
  /** Queued admission jobs waiting to start: resolve(true) = started, resolve(false) = cancelled. */
  private waiters = new Map<number, (started: boolean) => void>()

  constructor(opts: PoolOptions = {}) {
    this.admission = new AdmissionModel(opts.caps ?? DEFAULT_CAPS)
    this.recyclePages = opts.recyclePages ?? 256
    this.openDeadlineMs = opts.openDeadlineMs ?? 30_000
    this.now = opts.now ?? (() => performance.now())
    this.testHooks = opts.testHooks ?? false
    this.servedCacheBytes = opts.servedCacheBytes
    for (let i = 0; i < (opts.workers ?? defaultWorkers()); i++) {
      const slot = { index: i, busy: false, queue: [], pending: new Map(), docs: new Map(), pages: 0, parses: 0, generation: 0, closed: false } as unknown as Slot
      this.attach(slot)
      this.slots.push(slot)
    }
  }

  private static failure(error: string): WorkerReply { return { id: -1, ok: false, error, stats: { parses: 0, pages: 0, docs: 0 } } }

  /** Settles every in-flight post on `slot` with a failure reply. */
  private settle(slot: Slot, error: string) {
    const pending = [...slot.pending.values()]
    slot.pending.clear()
    for (const p of pending) p.ok(Pool.failure(error))
  }

  /** Starts a worker for `slot`; replies resolve the slot's pending map; an
   * unexpected death settles its posts (nothing waits forever). */
  private attach(slot: Slot): void {
    const worker = new Worker(WORKER_URL, { preload: [PRELOAD] } as WorkerOptions)
    slot.worker = worker
    slot.closed = false
    let markClosed!: () => void
    slot.whenClosed = new Promise<void>((r) => { markClosed = r })
    worker.onmessage = (e: MessageEvent<WorkerReply>) => {
      if (slot.worker !== worker) return
      const p = slot.pending.get(e.data.id)
      if (!p) return
      slot.pending.delete(e.data.id)
      slot.pages = e.data.stats.pages; slot.parses = e.data.stats.parses
      p.ok(e.data)
    }
    worker.addEventListener('close', () => {
      markClosed()
      if (slot.worker !== worker) return
      slot.closed = true
      slot.docs.clear()
      this.settle(slot, 'worker died')
    })
  }

  /** Runs `fn` when the slot is idle; one in-flight message per worker. */
  private async exclusive<T>(slot: Slot, fn: () => Promise<T>): Promise<T> {
    if (slot.busy) await new Promise<void>((r) => slot.queue.push(r))
    slot.busy = true
    try {
      if (this.disposed) throw new RenderFailed('pool disposed')
      if (slot.closed) await this.replace(slot)
      return await fn()
    } finally {
      slot.busy = false
      slot.queue.shift()?.()
    }
  }

  /** Posts and waits at most `ms`; on timeout the worker is retired (the
   * caller learns via 'timeout'). */
  private async post(slot: Slot, msg: WorkerRequest, ms?: number): Promise<WorkerReply | 'timeout'> {
    const reply = new Promise<WorkerReply>((ok) => { slot.pending.set(msg.id, { ok }); slot.worker.postMessage(msg) })
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
    const slot = this.pick(docRef)
    return this.exclusive(slot, async () => (await this.ensureOpen(slot, docRef, contentPolicy, true)).open!)
  }

  /** Admits, waits for the worker, renders under the deadline. Every exit
   * path refunds exactly once: here for normal ends, in `renderOn` (after the
   * worker is gone) for a retired job, in `admission.cancel` for a queued job
   * that never started. */
  async render(sessionId: string, job: RenderJob, bytesEstimate: number): Promise<RenderedUnit> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    const admitted = res.job
    if (res.status === 'queued' && !(await new Promise<boolean>((r) => this.waiters.set(admitted.id, r)))) throw new RenderFailed('pool disposed')
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, () => this.renderOn(slot, admitted, job))
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** A notes job (Stage 3c): the same affinity, admission, deadline and recycle paths as `render`. */
  async notes(sessionId: string, job: PoolNotesJob, bytesEstimate: number): Promise<NotesPlan | NotesRaster> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    const admitted = res.job
    if (res.status === 'queued' && !(await new Promise<boolean>((r) => this.waiters.set(admitted.id, r)))) throw new RenderFailed('pool disposed')
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, async () => {
        const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
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
  async encode(sessionId: string, job: EncodeJob, bytesEstimate: number): Promise<MarkRasterResult> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    const admitted = res.job
    if (res.status === 'queued' && !(await new Promise<boolean>((r) => this.waiters.set(admitted.id, r)))) throw new RenderFailed('pool disposed')
    const slot = [...this.slots].sort((a, b) => (Number(a.busy) - Number(b.busy)) || (a.queue.length - b.queue.length) || a.index - b.index)[0]!
    try {
      return await this.exclusive(slot, () => {
        const msg: WorkerRequest = { type: 'encode', id: this.nextId++, png: job.png, w: job.w, h: job.h, kBasis: job.kBasis, code: job.code, ...(this.testHooks && job.hangMs ? { hangMs: job.hangMs } : {}), ...(this.testHooks && job.crash ? { crash: true } : {}) }
        this.encodes++
        return this.runJob(slot, admitted, msg) as Promise<MarkRasterResult>
      })
    } finally {
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  /** A text op on the docRef's worker (affinity, admission, deadline as
   * `render`). The op is charged `bytesEstimate`; every prepare (explicit or
   * for a `{ missing }` answer) is charged `prepareWorkspace(w, h)` from the
   * unit's parsed size through the same admission, and the worker is sent
   * that grant to enforce. A missing answer prepares each missing unit (one
   * job each, render deadline) and retries once; still missing ⇒
   * TextTooLarge. Everything runs in one exclusive turn of the slot, and the
   * recycle threshold is checked on every exit, success or not. */
  async text(sessionId: string, job: PoolTextJob, bytesEstimate = TEXT_JOB_BYTES): Promise<TextReply> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    const admitted = res.job
    if (res.status === 'queued' && !(await new Promise<boolean>((r) => this.waiters.set(admitted.id, r)))) throw new RenderFailed('pool disposed')
    const slot = this.pick(job.docRef)
    const { docRef: _d, contentPolicy: _c, ...rest } = job
    const op = rest as TextJob
    const k = `${job.docRef}|${job.contentPolicy}`
    try {
      return await this.exclusive(slot, async () => {
        const generation = slot.generation
        try {
          const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
          const post = (j: TextJob, owner: Job, grant?: number) => this.runJob(slot, owner, { type: 'text', id: this.nextId++, docKey, job: j, ...(grant !== undefined ? { grant } : {}) }) as Promise<TextReply>
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

  private async renderOn(slot: Slot, admitted: Job, job: RenderJob): Promise<RenderedUnit> {
    const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
    const msg: WorkerRequest = { type: 'render', id: this.nextId++, docKey, unitIndex: job.unitIndex, scale: job.scale, projection: job.projection, mark: job.mark, ...(job.missingFaces ? { missingFaces: job.missingFaces } : {}), ...(this.testHooks && job.hangMs ? { hangMs: job.hangMs } : {}), ...(this.testHooks && job.crash ? { crash: true } : {}) }
    const r = await this.runJob(slot, admitted, msg)
    if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
    return r as RenderedUnit
  }

  /** Posts one admitted job under the render deadline; owns the retire paths. */
  private async runJob(slot: Slot, admitted: Job, msg: WorkerRequest): Promise<unknown> {
    // The deadline covers the paint only (the worker is idle when we post).
    admitted.startedAt = this.now()
    const reply = await this.post(slot, msg, this.admission.caps.renderDeadlineMs + 1)
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
    this.settle(slot, 'worker retired')
    if (!slot.closed) {
      const closed = slot.whenClosed
      old.terminate()
      await closed
    }
    slot.docs.clear(); slot.pages = 0; slot.parses = 0; slot.generation++
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
  }

  stats() {
    return {
      workers: this.slots.map((s) => ({ index: s.index, generation: s.generation, pages: s.pages, parses: s.parses, docs: s.docs.size, busy: s.busy, closed: s.closed })),
      reserved: this.admission.reserved, running: this.admission.runningCount, queued: this.admission.queued,
      retired: this.retired, recycled: this.recycled, encodes: this.encodes, affinity: Object.fromEntries(this.affinity), unitSizes: this.unitSizes.size,
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
      s.worker.terminate()
      for (const wake of s.queue.splice(0)) wake()
    }
  }
}
