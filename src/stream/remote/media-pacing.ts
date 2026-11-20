/**
 * Piece 5 (plan D12; coordinator ruling 3, Task 10): the session's `media-request` send log (lazy remote path only).
 *
 * The server's `media` class allows 6 requests in any 60 s per sid (refuse-only). The client holds its sends to one
 * log per StreamSession. As R29, a send holds its slot until its first reply arrives and for `spanMs` after it: the
 * server counted it no later than its reply. Look-ahead sends (for a slide other than the viewer's) may hold at most
 * `aheadCap` of the `cap` slots, so the current slide's requests always have the rest. A `rate-limited` refusal
 * blocks the log only until its oldest answered slot ages out.
 */

interface Slot { at: number | undefined; ahead: boolean }

export class MediaSendLog {
  /** reqId → when its first reply arrived (undefined: not yet), in send order. */
  private readonly sends = new Map<number, Slot>()
  private blockedUntil = -Infinity

  constructor(private readonly now: () => number, readonly cap: number, readonly aheadCap: number, readonly spanMs: number) {}

  /** An answered send at t counts through t + spanMs inclusive. */
  private prune(t: number) { for (const [k, s] of this.sends) if (s.at !== undefined && s.at + this.spanMs < t) this.sends.delete(k) }

  record(reqId: number, ahead: boolean) { this.sends.delete(reqId); this.sends.set(reqId, { at: undefined, ahead }) }
  /** True when this was the send's first reply. */
  answered(reqId: number): boolean { const s = this.sends.get(reqId); if (!s || s.at !== undefined) return false; s.at = this.now(); return true }
  /** The stale-token re-send `to` takes over `from`'s slot, unanswered, aging only from `to`'s reply (the server
   * never counted `from`). */
  transfer(from: number, to: number) { const s = this.sends.get(from); if (!s) return; this.sends.delete(from); this.sends.set(to, { at: undefined, ahead: s.ahead }) }
  has(reqId: number) { this.prune(this.now()); return this.sends.has(reqId) }
  /** Every unanswered send is taken as received now (resume, media-end). */
  settleAll() { const t = this.now(); for (const s of this.sends.values()) if (s.at === undefined) s.at = t }

  private aheadHeld() { let n = 0; for (const s of this.sends.values()) if (s.ahead) n++; return n }

  /** True when a send of this class may go out now. */
  room(ahead: boolean): boolean {
    const t = this.now()
    this.prune(t)
    if (t < this.blockedUntil || this.sends.size >= this.cap) return false
    return !ahead || this.aheadHeld() < this.aheadCap
  }

  /** ms until `room(ahead)`; Infinity when only a reply can make room. */
  waitMs(ahead: boolean): number {
    const t = this.now()
    this.prune(t)
    const block = Math.max(0, this.blockedUntil - t)
    const expiry = (s: Slot) => s.at! + this.spanMs + 1 - t
    const kth = (want: (s: Slot) => boolean, need: number) => {
      if (need <= 0) return 0
      const e = [...this.sends.values()].filter((s) => want(s) && s.at !== undefined).map(expiry).sort((a, b) => a - b)
      return e.length < need ? Infinity : e[need - 1]!
    }
    const total = kth(() => true, this.sends.size - this.cap + 1)
    const aheadWait = ahead ? kth((s) => s.ahead, this.aheadHeld() - this.aheadCap + 1) : 0
    return Math.max(block, total, aheadWait)
  }

  /** A `rate-limited` media refusal: no send until the oldest answered slot ages out (a full span if none is). */
  refused() {
    const t = this.now()
    this.prune(t)
    let oldest = Infinity
    for (const s of this.sends.values()) if (s.at !== undefined) oldest = Math.min(oldest, s.at + this.spanMs + 1)
    this.blockedUntil = Math.max(this.blockedUntil, Number.isFinite(oldest) ? oldest : t + this.spanMs)
  }
}
