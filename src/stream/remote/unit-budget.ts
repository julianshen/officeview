/**
 * R28: the unit send budget of one stream session (lazy remote path only).
 *
 * The server's `'unit'` rate class allows 60 requests in any 60 s per sid and
 * closes the session 4008 on the third refusal inside 60 s (server/src/rate.ts).
 * Its window lives as long as the sid: across show restarts and resume. So the
 * client keeps ONE sliding log per StreamSession, stamped on every
 * server-charged unit send from any caller, and holds sends to it.
 */

/** The server's `'unit'` limit per 60 s window. */
export const UNIT_SERVER_LIMIT = 60
/** The client window: the server's 60 s plus a margin, as the server stamps a
 * request on receipt, later than the client sent it. */
export const UNIT_SPAN_MS = 62_000
/** First sends of a request (any caller) are held to this many per window. */
export const UNIT_FIRST_CAP = 48
/** Re-sends of a request already sent (validation, resume) may go up to this. */
export const UNIT_RESEND_CAP = 56
/** After a refusal the log explains, no charged send for at least this long. */
export const UNIT_REFUSAL_BACKOFF_MS = 2_000

/** What a caller sees of a budget: room now, and how long until there is some. */
export interface UnitBudgetView {
  /** Sends the budget admits now beyond `pending` already spoken for (0 while blocked). */
  room(pending?: number): number
  /** ms until `room(pending)` > 0; Infinity when only `pending` settling can make room. */
  waitMs(pending?: number): number
}

export class UnitSendLog implements UnitBudgetView {
  /** Charged send times inside the window, oldest first. */
  private readonly stamps: number[] = []
  private blockedUntil = -Infinity

  constructor(private readonly now: () => number, readonly cap: number, readonly spanMs: number) {}

  /** A stamp at t counts through t + spanMs inclusive (conservative at the boundary). */
  private prune(t: number) { while (this.stamps.length && this.stamps[0]! + this.spanMs < t) this.stamps.shift() }

  /** One server-charged send, now. */
  record() { this.stamps.push(this.now()) }

  room(pending = 0, cap = this.cap): number {
    const t = this.now()
    this.prune(t)
    if (t < this.blockedUntil) return 0
    return Math.max(0, cap - this.stamps.length - pending)
  }

  waitMs(pending = 0, cap = this.cap): number {
    const t = this.now()
    this.prune(t)
    // Room needs the stamp at this index (and all older) to have expired.
    const i = this.stamps.length + pending - cap
    const byLog = i < 0 ? 0 : i < this.stamps.length ? this.stamps[i]! + this.spanMs + 1 - t : Infinity
    return Math.max(0, byLog, this.blockedUntil - t)
  }

  /**
   * A `'rate-limited'` unit refusal. If the log holds the server's whole limit
   * it explains the refusal: no send for UNIT_REFUSAL_BACKOFF_MS, and then only
   * once the log has room. Otherwise the server counted sends this log never saw
   * (its history is unknown here), so the window is treated as full for a whole
   * span, after which every send it counts is one this log stamped.
   */
  refused() {
    const t = this.now()
    this.prune(t)
    const known = this.stamps.length >= UNIT_SERVER_LIMIT
    this.blockedUntil = Math.max(this.blockedUntil, t + (known ? UNIT_REFUSAL_BACKOFF_MS : this.spanMs))
  }
}
