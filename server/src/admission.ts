/**
 * Pool admission (spec §9.2/§14): per-session outstanding/queue caps, process
 * queue, transient and render reservations, and drain-then-refund retirement.
 * First built as the Stage-0 T12 model; the server pool and T12 share it.
 */
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
    // A non-integer/negative estimate, or one no idle pool could ever fit, is refused (never queued forever).
    if (!Number.isSafeInteger(bytes) || bytes < 0 || bytes > this.caps.sessionTransientBytes || bytes > this.caps.renderReservationBytes) { this.refusals['too-large']++; return { status: 'refused', reason: 'too-large' } }
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

  /** Removes a queued (never started) job; its session slot is released. */
  cancel(jobId: number): boolean {
    const i = this.queue.findIndex((j) => j.id === jobId)
    if (i < 0) return false
    const [job] = this.queue.splice(i, 1)
    this.session(job!.sessionId).queued--
    job!.state = 'done'
    return true
  }

  get queued(): number { return this.queue.length }
  get runningCount(): number { return this.running.size }
}
