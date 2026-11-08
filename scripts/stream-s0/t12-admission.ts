/**
 * T12 (Stage-0 scope call 2): a SYNTHETIC admission probe. There is no server
 * at Stage 0, so this is an in-process model of the pool's §14 / §9.2
 * reservation and queue rules, driven by 1,000 concurrent requests over 40
 * sessions plus 5 rogue sessions bursting 100 each. Admitted jobs run a REAL
 * @napi-rs/canvas paint of a corpus unit in one of 4 worker threads.
 *
 *   bun run s0:napi scripts/stream-s0/t12-admission.ts
 *
 * It proves the caps are enforceable and measures peak RSS and latency; it does
 * not test a server.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { resolveBackend, writeResult } from './results'

export interface Caps {
  perSessionOutstanding: number
  perSessionQueue: number
  sessionTransientBytes: number
  renderReservationBytes: number
  processQueuedUnits: number
  renderDeadlineMs: number
}

/** §14 / §9.2 values. */
export const DEFAULT_CAPS: Caps = {
  perSessionOutstanding: 8,
  perSessionQueue: 32,
  sessionTransientBytes: 96 * 1024 * 1024,
  renderReservationBytes: 512 * 1024 * 1024,
  processQueuedUnits: 4096,
  renderDeadlineMs: 2000,
}

export type JobState = 'running' | 'queued' | 'retiring' | 'done'
export interface Job { id: number; sessionId: string; bytes: number; state: JobState; startedAt?: number; timedOut: boolean }
export type RequestResult = { status: 'admitted' | 'queued'; job: Job } | { status: 'refused'; reason: 'session-queue' | 'process-queue' | 'too-large' }

interface Session { running: number; queued: number; reserved: number }

/**
 * Charges are reserved when a job STARTS and refunded only when its worker
 * reports it finished — a job past its deadline becomes `retiring` and keeps
 * its charge until then (drain-then-refund, §9.2/§10).
 */
export class AdmissionModel {
  readonly caps: Caps
  private readonly sessions = new Map<string, Session>()
  private readonly queue: Job[] = []
  private readonly running = new Set<Job>()
  private nextId = 1
  reserved = 0
  peakReserved = 0
  peakSessionReserved = 0
  peakSessionRunning = 0
  peakSessionQueued = 0
  peakProcessQueued = 0
  refusals = { 'session-queue': 0, 'process-queue': 0, 'too-large': 0 }
  timeouts = 0

  constructor(caps: Caps = DEFAULT_CAPS) { this.caps = caps }

  private session(id: string): Session {
    let s = this.sessions.get(id)
    if (!s) { s = { running: 0, queued: 0, reserved: 0 }; this.sessions.set(id, s) }
    return s
  }

  private fits(s: Session, bytes: number): boolean {
    return s.running < this.caps.perSessionOutstanding
      && s.reserved + bytes <= this.caps.sessionTransientBytes
      && this.reserved + bytes <= this.caps.renderReservationBytes
  }

  private start(job: Job, s: Session, now: number): void {
    job.state = 'running'
    job.startedAt = now
    s.running++
    this.running.add(job)
    s.reserved += job.bytes
    this.reserved += job.bytes
    this.peakReserved = Math.max(this.peakReserved, this.reserved)
    this.peakSessionReserved = Math.max(this.peakSessionReserved, s.reserved)
    this.peakSessionRunning = Math.max(this.peakSessionRunning, s.running)
  }

  request(sessionId: string, bytes: number, now = 0): RequestResult {
    if (bytes > this.caps.sessionTransientBytes) { this.refusals['too-large']++; return { status: 'refused', reason: 'too-large' } }
    const s = this.session(sessionId)
    const job: Job = { id: this.nextId++, sessionId, bytes, state: 'queued', timedOut: false }
    // FIFO fairness: never jump ahead of this session's own queued work.
    if (s.queued === 0 && this.fits(s, bytes)) { this.start(job, s, now); return { status: 'admitted', job } }
    if (s.queued >= this.caps.perSessionQueue) { this.refusals['session-queue']++; return { status: 'refused', reason: 'session-queue' } }
    if (this.queue.length >= this.caps.processQueuedUnits) { this.refusals['process-queue']++; return { status: 'refused', reason: 'process-queue' } }
    s.queued++
    this.queue.push(job)
    this.peakSessionQueued = Math.max(this.peakSessionQueued, s.queued)
    this.peakProcessQueued = Math.max(this.peakProcessQueued, this.queue.length)
    return { status: 'queued', job }
  }

  /** Marks running jobs past the deadline `retiring` (charge kept). Returns them. */
  expire(now: number): Job[] {
    const out: Job[] = []
    for (const job of this.running) {
      if (job.state === 'running' && job.startedAt !== undefined && now - job.startedAt > this.caps.renderDeadlineMs) {
        job.state = 'retiring'
        job.timedOut = true
        this.timeouts++
        out.push(job)
      }
    }
    return out
  }

  /** Called once the worker has actually finished (drained). Refunds, then pumps the queue; returns newly started jobs. */
  finish(job: Job, now = 0): Job[] {
    if (job.state !== 'running' && job.state !== 'retiring') throw new Error(`finish on ${job.state} job ${job.id}`)
    const s = this.session(job.sessionId)
    job.state = 'done'
    this.running.delete(job)
    s.running--
    s.reserved -= job.bytes
    this.reserved -= job.bytes
    return this.pump(now)
  }

  /** Starts queued jobs in FIFO order wherever their session and the process have room. */
  pump(now = 0): Job[] {
    const started: Job[] = []
    const blocked = new Set<string>()
    for (let i = 0; i < this.queue.length;) {
      const job = this.queue[i]!
      const s = this.session(job.sessionId)
      if (!blocked.has(job.sessionId) && this.fits(s, job.bytes)) {
        this.queue.splice(i, 1)
        s.queued--
        this.start(job, s, now)
        started.push(job)
      } else {
        blocked.add(job.sessionId) // keep per-session FIFO
        i++
      }
    }
    return started
  }

  get queued(): number { return this.queue.length }
  get runningCount(): number { return this.running.size }
}

// ---------------------------------------------------------------------------
// Driver (real paints in worker threads)

interface WorkerReply { id: number; ms: number; backend: string; error?: string }

async function main(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t12] refused: run with `bun run s0:napi scripts/stream-s0/t12-admission.ts`')
    return 2
  }
  const WORKERS = 4
  const preload = join(fileURLToPath(new URL('.', import.meta.url)), 'napi-preload.ts')
  const workerUrl = new URL('./t12-worker.ts', import.meta.url).href
  const workers = Array.from({ length: WORKERS }, () => new Worker(workerUrl, { preload: [preload] } as WorkerOptions))
  const ready = await Promise.all(workers.map((w) => new Promise<{ backend: string; units: Array<{ key: string; bytes: number }> }>((resolve, reject) => {
    w.onmessage = (e) => resolve(e.data)
    w.onerror = (e) => reject(e)
  })))
  const backends = ready.map((r) => r.backend)
  if (!backends.every((b) => b === '@napi-rs/canvas')) throw new Error(`worker backends ${backends.join(',')}`)
  const units = ready[0]!.units
  const model = new AdmissionModel()
  const t0 = performance.now()
  const now = () => performance.now() - t0
  const latency = new Map<number, number>() // job id → request time
  const done: number[] = []
  const failures: string[] = []
  let rssPeak = process.memoryUsage().rss
  const sampler = setInterval(() => { rssPeak = Math.max(rssPeak, process.memoryUsage().rss) }, 25)

  const idle: Worker[] = [...workers]
  const waiting: Job[] = [] // started (charged) jobs not yet on a worker
  const unitOf = new Map<number, string>()
  const inflight = new Map<number, Job>()
  let resolveAll!: () => void
  const allDone = new Promise<void>((r) => { resolveAll = r })
  let outstanding = 0

  const dispatch = () => {
    while (idle.length > 0 && waiting.length > 0) {
      const job = waiting.shift()!
      const w = idle.pop()!
      inflight.set(job.id, job)
      ;(w as Worker & { jobId?: number }).jobId = job.id
      w.postMessage({ id: job.id, key: unitOf.get(job.id) })
    }
  }
  for (const w of workers) {
    w.onmessage = (e: MessageEvent<WorkerReply>) => {
      const reply = e.data
      const job = inflight.get(reply.id)!
      inflight.delete(reply.id)
      if (reply.error) failures.push(reply.error)
      done.push(now() - latency.get(job.id)!)
      for (const started of model.finish(job, now())) waiting.push(started)
      idle.push(w)
      outstanding--
      dispatch()
      if (outstanding === 0) resolveAll()
    }
  }
  const deadlineTimer = setInterval(() => model.expire(now()), 100)

  const sessions: Array<{ id: string; count: number }> = [
    ...Array.from({ length: 40 }, (_, i) => ({ id: `s${i}`, count: 25 })),
    ...Array.from({ length: 5 }, (_, i) => ({ id: `rogue${i}`, count: 100 })),
  ]
  const results = { admitted: 0, queued: 0, refused: 0 }
  // All requests arrive "concurrently": interleave sessions round-robin at t≈0.
  const maxCount = Math.max(...sessions.map((s) => s.count))
  let n = 0
  for (let k = 0; k < maxCount; k++) {
    for (const s of sessions) {
      if (k >= s.count) continue
      const unit = units[n++ % units.length]!
      const r = model.request(s.id, unit.bytes, now())
      results[r.status]++
      if (r.status === 'refused') continue
      latency.set(r.job.id, now())
      unitOf.set(r.job.id, unit.key)
      outstanding++
      if (r.status === 'admitted') waiting.push(r.job)
    }
  }
  // Queued jobs carry their unit key too (assigned above); start dispatching.
  dispatch()
  if (outstanding > 0) await allDone
  clearInterval(sampler)
  clearInterval(deadlineTimer)
  for (const w of workers) w.terminate()

  done.sort((a, b) => a - b)
  const pct = (p: number) => done[Math.min(done.length - 1, Math.floor(done.length * p))] ?? 0
  const caps = DEFAULT_CAPS
  const within = model.peakReserved <= caps.renderReservationBytes
    && model.peakSessionRunning <= caps.perSessionOutstanding
    && model.peakSessionQueued <= caps.perSessionQueue
    && model.peakSessionReserved <= caps.sessionTransientBytes
    && model.peakProcessQueued <= caps.processQueuedUnits
  const green = within && failures.length === 0
  writeResult('t12', {
    status: green ? 'GREEN' : 'RED',
    synthetic: true,
    caps,
    requests: { total: n, ...results, refusalsByReason: model.refusals },
    peak: {
      reservedBytes: model.peakReserved,
      sessionReservedBytes: model.peakSessionReserved,
      sessionRunning: model.peakSessionRunning,
      sessionQueued: model.peakSessionQueued,
      processQueued: model.peakProcessQueued,
      rssBytes: rssPeak,
    },
    timeouts: model.timeouts,
    paintFailures: failures.slice(0, 10),
    latencyMs: { p50: pct(0.5), p95: pct(0.95), max: done[done.length - 1] ?? 0, completed: done.length },
    workers: { count: WORKERS, backends },
    unitsUsed: units.length,
    note: 'Synthetic in-process model of §14/§9.2 admission (no server at Stage 0); paints are real @napi-rs/canvas renders in worker threads. Latency is request→finish wall clock and depends on machine load.',
  })
  console.log(`[t12] ${green ? 'GREEN' : 'RED'} requests ${n} admitted ${results.admitted} queued ${results.queued} refused ${results.refused} · peak reserved ${(model.peakReserved / 2 ** 20).toFixed(1)} MiB · session running≤${model.peakSessionRunning} queued≤${model.peakSessionQueued} · rss ${(rssPeak / 2 ** 20).toFixed(0)} MiB · p95 ${pct(0.95).toFixed(0)} ms · timeouts ${model.timeouts}`)
  return green ? 0 : 3
}

if (import.meta.main) process.exit(await main())
