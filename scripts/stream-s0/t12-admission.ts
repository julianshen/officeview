export interface Caps {
  perSessionOutstanding: number
  perSessionQueue: number
  sessionTransientBytes: number
  renderReservationBytes: number
  processSockets: number
  processQueuedUnits: number
  renderDeadlineMs: number
}

export const DEFAULT_CAPS: Caps = {
  perSessionOutstanding: 8,
  perSessionQueue: 32,
  sessionTransientBytes: 96 * 1024 * 1024,
  renderReservationBytes: 512 * 1024 * 1024,
  processSockets: 256,
  processQueuedUnits: 4096,
  renderDeadlineMs: 2000,
}

interface SessionState { outstanding: number; queued: number; reservedBytes: number }

export class AdmissionModel {
  caps: Caps
  sessions = new Map<string, SessionState>()
  globalQueued = 0
  globalReserved = 0
  outstandingMax = 0
  queueMax = 0
  refusals = 0
  completed = 0

  constructor(caps: Caps = DEFAULT_CAPS) { this.caps = caps }

  private sess(id: string): SessionState {
    let s = this.sessions.get(id)
    if (!s) { s = { outstanding: 0, queued: 0, reservedBytes: 0 }; this.sessions.set(id, s) }
    return s
  }

  request(sessionId: string, bytes: number): 'admitted' | 'queued' | 'refused' {
    const s = this.sess(sessionId)
    if (s.outstanding < this.caps.perSessionOutstanding) {
      s.outstanding++
      s.reservedBytes += bytes
      this.globalReserved += bytes
      if (s.outstanding > this.outstandingMax) this.outstandingMax = s.outstanding
      if (this.globalReserved > this.caps.renderReservationBytes) {
        // would exceed; refund and refuse
        s.outstanding--
        s.reservedBytes -= bytes
        this.globalReserved -= bytes
        this.refusals++
        return 'refused'
      }
      return 'admitted'
    }
    if (s.queued < this.caps.perSessionQueue && this.globalQueued < this.caps.processQueuedUnits) {
      s.queued++
      this.globalQueued++
      if (s.queued > this.queueMax) this.queueMax = s.queued
      if (this.globalQueued > this.queueMax) this.queueMax = this.globalQueued
      return 'queued'
    }
    this.refusals++
    return 'refused'
  }

  complete(id: string, bytes: number, sessionId: string): void {
    const s = this.sess(sessionId)
    // drain: if queued, promote one
    if (s.outstanding > 0) {
      s.outstanding--
      s.reservedBytes -= bytes
      this.globalReserved -= bytes
      if (s.reservedBytes < 0) s.reservedBytes = 0
      if (this.globalReserved < 0) this.globalReserved = 0
    }
    if (s.queued > 0) {
      s.queued--
      this.globalQueued--
      s.outstanding++
      s.reservedBytes += bytes
      this.globalReserved += bytes
      if (s.outstanding > this.outstandingMax) this.outstandingMax = s.outstanding
    }
    this.completed++
  }

  peak() {
    return { reserved: this.globalReserved, queued: this.globalQueued, outstandingMax: this.outstandingMax, queueMax: this.queueMax, refusals: this.refusals }
  }
}
