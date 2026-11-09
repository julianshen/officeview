/** §14 client admission: one aggregate transient reservation per session
 * (≤ 96 MiB) covering every in-flight receive/inflate/decode, plus the
 * outstanding-request cap (8). Requests that do not fit wait, in order. */
export const SESSION_TRANSIENT_BYTES = 96 * 1024 * 1024
export const MAX_OUTSTANDING = 8
const MiB = 1024 * 1024
const RASTER_DECODED_CAP = 64 * MiB, ENCODED_CAP = 8 * MiB, COMMANDS_DECODED = 8 * MiB, ASSETS = 9 * MiB
/** Decoded ASSET_DEFINE bytes one unit request may receive in total. */
export const ASSET_ALLOWANCE = ASSETS

/** The bytes one unit request may need (§14): the raster's RGBA at the bucket
 * (capped at the 64 MiB decoded cap) plus its encoded bytes (≤ the 8 MiB
 * encoded cap, never more than the raw RGBA + PNG/gzip overhead). A session
 * whose projection can yield commands also reserves the commands unit plus
 * its assets (8 + 9 MiB decoded), since either may arrive. */
export function unitCost(w: number, h: number, bucket: number, mayBeCommands: boolean): number {
  const rgba = Math.min(RASTER_DECODED_CAP, w * h * bucket * bucket * 4)
  const decoded = mayBeCommands ? Math.max(rgba, COMMANDS_DECODED + ASSETS) : rgba
  return Math.min(ENCODED_CAP, Math.ceil(decoded * 1.01) + 65536) + decoded
}

export class Admission {
  private used = 0
  private outstanding = 0
  private queue: Array<{ cost: number; go: () => void; fail: (e: unknown) => void; dropped: boolean }> = []
  constructor(private readonly limit = SESSION_TRANSIENT_BYTES, private readonly max = MAX_OUTSTANDING) {}
  get inFlightBytes() { return this.used }
  get inFlight() { return this.outstanding }
  /** Resolves with a release function once the request fits. */
  acquire(cost: number, signal?: AbortSignal): Promise<() => void> {
    return new Promise((resolve, reject) => {
      const entry = { cost: Math.min(cost, this.limit), dropped: false, go: () => {}, fail: reject }
      entry.go = () => {
        this.used += entry.cost; this.outstanding++
        let done = false
        resolve(() => { if (done) return; done = true; this.used -= entry.cost; this.outstanding--; this.pump() })
      }
      signal?.addEventListener('abort', () => { entry.dropped = true; reject(signal.reason) }, { once: true })
      this.queue.push(entry)
      this.pump()
    })
  }
  private pump() {
    while (this.queue.length) {
      const next = this.queue[0]!
      if (next.dropped) { this.queue.shift(); continue }
      if (this.outstanding >= this.max || this.used + next.cost > this.limit) return
      this.queue.shift(); next.go()
    }
  }
  /** Reject every waiter (session closed). */
  drain(reason: unknown) { for (const e of this.queue.splice(0)) if (!e.dropped) { e.dropped = true; e.fail(reason) } }
}
