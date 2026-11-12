/** §14 rate gates. Sliding windows over an injected clock (ms). */

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

export type RateClass = 'message' | 'unit' | 'text'

/** Per-sid gates: 120 msgs/10 s with a 20/s burst, 60 unit requests/min,
 * 60 text ops/min (R5: search, copy, a11y). A breach is an op refusal; the
 * third breach inside 60 s is a CLOSE (4008). A `'text'` overflow is only a
 * refusal: it never counts toward the breach window. */
export class SessionRates {
  private messages = new Window(120, 10_000)
  private burst = new Window(20, 1_000)
  private units = new Window(60, 60_000)
  private texts = new Window(60, 60_000)
  private breaches = new Window(2, 60_000)
  /** 'ok' | 'refuse' | 'close' */
  check(kind: RateClass, now: number): 'ok' | 'refuse' | 'close' {
    if (kind === 'text') return this.texts.hit(now) ? 'ok' : 'refuse'
    const ok = kind === 'unit' ? this.units.hit(now) : (this.burst.hit(now) && this.messages.hit(now))
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
