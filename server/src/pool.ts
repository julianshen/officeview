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
import type { MarkSpec, OpenResult, Projection, RenderedUnit } from './render'
import type { WorkerReply, WorkerRequest } from './worker/render-worker'

export class AdmissionRefused extends Error { constructor(readonly reason: 'session-queue' | 'process-queue' | 'too-large') { super(`admission refused: ${reason}`); this.name = 'AdmissionRefused' } }
export class RenderTimeout extends Error { constructor() { super('render deadline exceeded'); this.name = 'RenderTimeout' } }
export class RenderFailed extends Error { constructor(m: string, readonly refusal?: string) { super(m); this.name = 'RenderFailed' } }

export interface RenderJob { docRef: string; contentPolicy: 'visible-only' | 'full'; unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: string[]; hangMs?: number }
export interface PoolOptions { workers?: number; caps?: Caps; recyclePages?: number; now?: () => number; testHooks?: boolean }

const WORKER_URL = resolve(import.meta.dir, 'worker/render-worker.ts')
const PRELOAD = resolve(import.meta.dir, '../../scripts/stream-s0/napi-preload.ts')
export const defaultWorkers = () => Math.max(1, Math.min(8, Math.floor(availableParallelism() / 2)))

interface Slot {
  index: number
  worker: Worker
  busy: boolean
  queue: Array<() => void>
  pending: Map<number, { ok: (r: WorkerReply) => void }>
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
  private readonly recyclePages: number
  private readonly now: () => number
  private readonly testHooks: boolean
  private disposed = false

  constructor(opts: PoolOptions = {}) {
    this.admission = new AdmissionModel(opts.caps ?? DEFAULT_CAPS)
    this.recyclePages = opts.recyclePages ?? 256
    this.now = opts.now ?? (() => performance.now())
    this.testHooks = opts.testHooks ?? false
    for (let i = 0; i < (opts.workers ?? defaultWorkers()); i++) this.slots.push(this.spawn(i, 0))
  }

  private spawn(index: number, generation: number): Slot {
    const slot = { index, busy: false, queue: [], pending: new Map(), docs: new Map(), pages: 0, parses: 0, generation } as unknown as Slot
    this.attach(slot)
    return slot
  }

  /** Starts a worker for `slot`; replies resolve the slot's pending map. */
  private attach(slot: Slot): void {
    slot.worker = new Worker(WORKER_URL, { preload: [PRELOAD] } as WorkerOptions)
    slot.worker.onmessage = (e: MessageEvent<WorkerReply>) => {
      const p = slot.pending.get(e.data.id)
      if (!p) return
      slot.pending.delete(e.data.id)
      slot.pages = e.data.stats.pages; slot.parses = e.data.stats.parses
      p.ok(e.data)
    }
  }

  /** Runs `fn` when the slot is idle; one in-flight message per worker. */
  private async exclusive<T>(slot: Slot, fn: () => Promise<T>): Promise<T> {
    if (slot.busy) await new Promise<void>((r) => slot.queue.push(r))
    slot.busy = true
    try { return await fn() } finally {
      slot.busy = false
      slot.queue.shift()?.()
    }
  }

  private post(slot: Slot, msg: WorkerRequest, transfer: Transferable[] = []): Promise<WorkerReply> {
    return new Promise((ok) => { slot.pending.set(msg.id, { ok }); slot.worker.postMessage(msg, transfer) })
  }

  private pick(docRef: string): Slot {
    const known = this.affinity.get(docRef)
    if (known !== undefined) return this.slots[known]!
    const slot = [...this.slots].sort((a, b) => (a.docs.size - b.docs.size) || (Number(a.busy) - Number(b.busy)) || a.index - b.index)[0]!
    this.affinity.set(docRef, slot.index)
    return slot
  }

  private async ensureOpen(slot: Slot, docRef: string, policy: 'visible-only' | 'full'): Promise<{ docKey: string; open?: OpenResult }> {
    const k = `${docRef}|${policy}`, docKey = slot.docs.get(k)
    if (docKey) return { docKey }
    const bytes = this.sources.get(docRef)
    if (!bytes) throw new RenderFailed('document not registered')
    const reply = await this.post(slot, { type: 'open', id: this.nextId++, docRef, bytes, contentPolicy: policy })
    if (!reply.ok) throw new RenderFailed(reply.error, reply.refusal)
    const open = reply.result as OpenResult
    slot.docs.set(k, open.docKey)
    return { docKey: open.docKey, open }
  }

  /** Registers the bytes and parses on the docRef's worker. */
  async open(docRef: string, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full'): Promise<OpenResult> {
    this.sources.set(docRef, bytes)
    const slot = this.pick(docRef)
    return this.exclusive(slot, async () => {
      const k = `${docRef}|${contentPolicy}`
      slot.docs.delete(k)   // force a reply carrying the OpenResult
      const r = await this.ensureOpen(slot, docRef, contentPolicy)
      return r.open!
    })
  }

  /** Admits, waits for the worker, renders under the deadline. */
  async render(sessionId: string, job: RenderJob, bytesEstimate: number): Promise<RenderedUnit> {
    if (this.disposed) throw new RenderFailed('pool disposed')
    const res = this.admission.request(sessionId, bytesEstimate, this.now())
    if (res.status === 'refused') throw new AdmissionRefused(res.reason)
    if (res.status === 'queued') await this.started(res.job)
    const admitted = res.job
    const slot = this.pick(job.docRef)
    try {
      return await this.exclusive(slot, () => this.renderOn(slot, admitted, job))
    } finally {
      // A retired job was refunded by its retirement owner; everything else here.
      if (admitted.state === 'running') this.finish(admitted)
    }
  }

  private waiters = new Map<number, () => void>()
  private started(job: Job): Promise<void> {
    return new Promise((r) => this.waiters.set(job.id, r))
  }
  private finish(job: Job): void {
    for (const j of this.admission.finish(job, this.now())) { this.waiters.get(j.id)?.(); this.waiters.delete(j.id) }
  }

  private async renderOn(slot: Slot, admitted: Job, job: RenderJob): Promise<RenderedUnit> {
    const { docKey } = await this.ensureOpen(slot, job.docRef, job.contentPolicy)
    const deadline = this.admission.caps.renderDeadlineMs
    const msg: WorkerRequest = { type: 'render', id: this.nextId++, docKey, unitIndex: job.unitIndex, scale: job.scale, projection: job.projection, mark: job.mark, ...(job.missingFaces ? { missingFaces: job.missingFaces } : {}), ...(this.testHooks && job.hangMs ? { hangMs: job.hangMs } : {}) }
    admitted.startedAt = this.now()
    let timer: ReturnType<typeof setTimeout> | undefined
    const timeout = new Promise<'timeout'>((r) => { timer = setTimeout(() => r('timeout'), deadline + 1) })
    const reply = await Promise.race([this.post(slot, msg), timeout])
    clearTimeout(timer)
    if (reply === 'timeout') {
      // Retirement owner: the charge stays reserved until the worker is gone.
      this.admission.expire(this.now())
      if (admitted.state === 'running') { admitted.state = 'retiring'; admitted.timedOut = true }
      await this.replace(slot)
      this.retired++
      this.finish(admitted)
      throw new RenderTimeout()
    }
    if (!reply.ok) throw new RenderFailed(reply.error, reply.refusal)
    if (slot.pages >= this.recyclePages) { await this.replace(slot); this.recycled++ }
    return reply.result as RenderedUnit
  }

  /** Terminates and respawns a slot; its documents re-open on demand.
   * Resolves only after the old worker is gone (drain-then-refund, §10). */
  private async replace(slot: Slot): Promise<void> {
    const old = slot.worker
    for (const p of slot.pending.values()) p.ok({ id: -1, ok: false, error: 'worker retired', stats: { parses: 0, pages: 0, docs: 0 } })
    slot.pending.clear()
    const closed = new Promise<void>((r) => old.addEventListener('close', () => r(), { once: true }))
    old.terminate()
    await closed
    slot.docs.clear(); slot.pages = 0; slot.parses = 0; slot.generation++
    this.attach(slot)
  }

  /** Releases a parse (document-ledger eviction); the bytes stay, so a later
   * request re-parses on the same worker. */
  dropParsed(docRef: string, policy: 'visible-only' | 'full'): void {
    const i = this.affinity.get(docRef)
    if (i === undefined) return
    const slot = this.slots[i]!, k = `${docRef}|${policy}`, docKey = slot.docs.get(k)
    if (!docKey) return
    slot.docs.delete(k)
    void this.exclusive(slot, () => this.post(slot, { type: 'drop', id: this.nextId++, docKey }))
  }

  /** Forgets a document everywhere (no session references it any more). */
  forget(docRef: string): void {
    this.dropParsed(docRef, 'visible-only'); this.dropParsed(docRef, 'full')
    this.sources.delete(docRef)
    this.affinity.delete(docRef)
  }

  stats() {
    return {
      workers: this.slots.map((s) => ({ index: s.index, generation: s.generation, pages: s.pages, parses: s.parses, docs: s.docs.size, busy: s.busy })),
      reserved: this.admission.reserved, running: this.admission.runningCount, queued: this.admission.queued,
      retired: this.retired, recycled: this.recycled, affinity: Object.fromEntries(this.affinity),
    }
  }

  dispose(): void {
    this.disposed = true
    for (const s of this.slots) s.worker.terminate()
  }
}
