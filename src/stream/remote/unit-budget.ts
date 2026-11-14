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
/** How long an answered send keeps its slot after its first reply arrived:
 * the server's 60 s window plus a 2 s margin (R29). */
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

/** Sends tracked at once; past this the budget is treated as full (the server gate bounds it anyway). */
export const UNIT_TRACKED_MAX = 256

/**
 * R29: a send is aged from evidence of server receipt, not from the client
 * send. A send recorded with its reqId occupies a slot until the first server
 * reply for that reqId arrives (any reply, also to a request aborted locally),
 * then for a further `spanMs` from that reply: the server received it no later
 * than its reply arrived. An unanswered send never expires on time alone.
 * A send recorded without an id is taken as answered at once (the standalone
 * fetcher's log of its own charges).
 */
export class UnitSendLog implements UnitBudgetView {
  /** Send key → when its first reply arrived (undefined: not yet), in send order. */
  private readonly sends = new Map<number | object, number | undefined>()
  private blockedUntil = -Infinity

  constructor(private readonly now: () => number, readonly cap: number, readonly spanMs: number) {}

  /** An answered send at t counts through t + spanMs inclusive (conservative at the boundary). */
  private prune(t: number) {
    for (const [k, at] of this.sends) if (at !== undefined && at + this.spanMs < t) this.sends.delete(k)
  }

  /** One server-charged send, now; `reqId` ⇒ it waits for its reply (R29). */
  record(reqId?: number) {
    const t = this.now()
    if (reqId === undefined) this.sends.set({}, t)
    else { this.sends.delete(reqId); this.sends.set(reqId, undefined) }
  }

  /** The first reply for `reqId` arrived now; later replies change nothing. */
  answered(reqId: number) { if (this.sends.has(reqId) && this.sends.get(reqId) === undefined) this.sends.set(reqId, this.now()) }

  /**
   * R30: the send under `from` was refused before the server's rate gate (a
   * stale token) and goes out again as `to`: its slot, still unanswered, moves
   * to `to` and ages only from `to`'s reply. A later reply to `from` frees nothing.
   */
  transfer(from: number, to: number) {
    if (!this.sends.has(from)) return
    const at = this.sends.get(from)
    this.sends.delete(from)
    this.sends.set(to, at)
  }

  /** True while `reqId`'s send still occupies a slot. */
  has(reqId: number) { this.prune(this.now()); return this.sends.has(reqId) }

  /** Resume: every unanswered send is taken as received now. */
  settleAll() { const t = this.now(); for (const [k, at] of this.sends) if (at === undefined) this.sends.set(k, t) }

  room(pending = 0, cap = this.cap): number {
    const t = this.now()
    this.prune(t)
    if (t < this.blockedUntil || this.sends.size >= UNIT_TRACKED_MAX) return 0
    return Math.max(0, cap - this.sends.size - pending)
  }

  waitMs(pending = 0, cap = this.cap): number {
    const t = this.now()
    this.prune(t)
    const block = Math.max(0, this.blockedUntil - t)
    // Room needs this many sends to expire; only answered ones ever do, soonest first.
    const need = Math.max(this.sends.size + pending - cap + 1, this.sends.size - UNIT_TRACKED_MAX + 1)
    if (need <= 0) return block
    const expiries: number[] = []
    for (const at of this.sends.values()) if (at !== undefined) expiries.push(at + this.spanMs + 1 - t)
    if (expiries.length < need) return Infinity
    expiries.sort((a, b) => a - b)
    return Math.max(block, expiries[need - 1]!)
  }

  /**
   * A `'rate-limited'` unit refusal. If the log holds the server's whole limit
   * it explains the refusal: no send for UNIT_REFUSAL_BACKOFF_MS, and then only
   * once the log has room. Otherwise the server counted sends this log cannot
   * account for, so the window is treated as full for a whole span.
   */
  refused() {
    const t = this.now()
    this.prune(t)
    const known = this.sends.size >= UNIT_SERVER_LIMIT
    this.blockedUntil = Math.max(this.blockedUntil, t + (known ? UNIT_REFUSAL_BACKOFF_MS : this.spanMs))
  }
}
