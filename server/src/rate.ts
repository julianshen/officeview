/** §14 rate gates. Sliding windows over an injected clock (ms). */
import { NOTES_RATE_PER_MIN } from './notes'
import { PRINT_RATE_PER_MIN } from './print'

export class Window {
  private hits: number[] = []
  constructor(readonly limit: number, readonly windowMs: number) {}
  /** Records a hit at `now` and reports whether it stays within the limit. */
  hit(now: number): boolean {
    const from = now - this.windowMs
    while (this.hits.length && this.hits[0]! <= from) this.hits.shift()
    if (this.hits.length >= this.limit) return false
    this.hits.push(now)
    return true
  }
}

/** Many hits at once against a sliding window: what fits, and when more would (Stage 4a `'tile'`). */
class CountWindow {
  private hits: number[] = []
  constructor(readonly limit: number, readonly windowMs: number) {}
  /** Same expiry as `Window`: a hit at `t` frees at `t + windowMs`. */
  private prune(now: number) { const from = now - this.windowMs; while (this.hits.length && this.hits[0]! <= from) this.hits.shift() }
  room(now: number) { this.prune(now); return Math.max(0, this.limit - this.hits.length) }
  take(k: number, now: number) { for (let i = 0; i < k; i++) this.hits.push(now) }
  /** ms from `now` until `need` more hits fit (0: they fit now); past the limit, until the whole window frees. */
  retryFor(need: number, now: number): number {
    this.prune(now)
    if (need <= this.limit - this.hits.length) return 0
    if (this.hits.length === 0) return this.windowMs
    const i = need > this.limit ? this.hits.length - 1 : this.hits.length - this.limit + need - 1
    return this.hits[i]! + this.windowMs - now
  }
}

/** Stage 4a (S4 §6.8, R4): `'tile'` is counted per tile, 600 per 10 s with a 120-per-second burst. Refuse-only. */
export const TILE_RATE = Object.freeze({ perWindow: 600, windowMs: 10_000, burst: 120, burstMs: 1_000 })
/** TEST and measure only (`StreamServerOptions.rates`): one window replaces a class's windows, its burst included. */
export type RateOverrides = Partial<Record<'message' | 'tile', { perWindow: number; windowMs: number }>>

export type RateClass = 'message' | 'unit' | 'text' | 'print' | 'notes' | 'media'
/** Piece 5 (spec §4.3): `media-request`s per minute per session; refuse-only, never toward 4008, apart from `unit`. */
export const MEDIA_RATE_PER_MIN = 6

/** Per-sid gates: 120 msgs/10 s with a 20/s burst, 60 unit requests/min,
 * 60 text ops/min (R5: search, copy, a11y). A breach is an op refusal; the
 * third breach inside 60 s is a CLOSE (4008). A `'text'` overflow is only a
 * refusal: it never counts toward the breach window. `'print'` (Stage 3b):
 * 10 jobs/min, refuse-only like `'text'`. `'notes'` (Stage 3c): 30 notes
 * images/min, refuse-only. `'media'` (piece 5): 6 `media-request`s/min, refuse-only. */
export class SessionRates {
  private messages: Window
  private burst: Window | undefined
  private tiles: CountWindow[]
  private units = new Window(60, 60_000)
  private texts = new Window(60, 60_000)
  private prints = new Window(PRINT_RATE_PER_MIN, 60_000)
  private notes = new Window(NOTES_RATE_PER_MIN, 60_000)
  private media = new Window(MEDIA_RATE_PER_MIN, 60_000)
  private breaches = new Window(2, 60_000)
  /** `o` (TEST and measure only): a `message` or `tile` override is that class's only window. */
  constructor(o: RateOverrides = {}) {
    this.messages = o.message ? new Window(o.message.perWindow, o.message.windowMs) : new Window(120, 10_000)
    this.burst = o.message ? undefined : new Window(20, 1_000)
    this.tiles = o.tile ? [new CountWindow(o.tile.perWindow, o.tile.windowMs)] : [new CountWindow(TILE_RATE.perWindow, TILE_RATE.windowMs), new CountWindow(TILE_RATE.burst, TILE_RATE.burstMs)]
  }
  /**
   * Stage 4a (R4): charges up to `n` tiles at `now`, as many as every `'tile'` window has room for; the rest are
   * refused, and `retryAfterMs` (0 when none is) is when the windows have room for all of them. Refuse-only: it never
   * touches the breach window (so never 4008), and is separate from `'unit'`.
   */
  takeTiles(n: number, now: number): { admitted: number; retryAfterMs: number } {
    const admitted = Math.max(0, Math.min(n, ...this.tiles.map((w) => w.room(now))))
    for (const w of this.tiles) w.take(admitted, now)
    const refused = n - admitted
    return { admitted, retryAfterMs: refused > 0 ? Math.max(...this.tiles.map((w) => w.retryFor(refused, now))) : 0 }
  }
  /** 'ok' | 'refuse' | 'close' */
  check(kind: RateClass, now: number): 'ok' | 'refuse' | 'close' {
    if (kind === 'text') return this.texts.hit(now) ? 'ok' : 'refuse'
    if (kind === 'print') return this.prints.hit(now) ? 'ok' : 'refuse'
    if (kind === 'notes') return this.notes.hit(now) ? 'ok' : 'refuse'
    if (kind === 'media') return this.media.hit(now) ? 'ok' : 'refuse'
    const ok = kind === 'unit' ? this.units.hit(now) : ((this.burst?.hit(now) ?? true) && this.messages.hit(now))
    if (ok) return 'ok'
    return this.breaches.hit(now) ? 'refuse' : 'close'
  }
}

/** Per-IP: 30 new sessions/min, 8 concurrent connects. */
export class IpRates {
  private sessions = new Map<string, Window>()
  private live = new Map<string, number>()
  constructor(readonly perMinute = 30, readonly concurrent = 8) {}
  admit(ip: string, now: number): 'ok' | 'rate' | 'concurrent' {
    if ((this.live.get(ip) ?? 0) >= this.concurrent) return 'concurrent'
    let w = this.sessions.get(ip)
    if (!w) { w = new Window(this.perMinute, 60_000); this.sessions.set(ip, w) }
    if (!w.hit(now)) return 'rate'
    this.live.set(ip, (this.live.get(ip) ?? 0) + 1)
    return 'ok'
  }
  release(ip: string) { const n = (this.live.get(ip) ?? 1) - 1; if (n <= 0) this.live.delete(ip); else this.live.set(ip, n) }
}
