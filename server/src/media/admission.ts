/**
 * Piece 5 (spec §6, plan D12): the server-wide media admission class. ⌈cores / 4⌉ distinct encodes run at once; a
 * queue of 16 places holds the rest. A joiner of a running encode never comes here (it takes no slot and no place).
 *
 * - `request` (a viewer's `media-request`): a free slot runs at once; otherwise it takes a queue place; with the queue
 *   full it is refused `media-busy`. It is the only mode that can be refused.
 * - `ahead` (show-start encode-ahead): never refused. A free slot runs at once (when nobody is queued); otherwise it
 *   takes a queue place if its session holds none and one is free; otherwise it waits silently, parked, holding no
 *   place. A session therefore holds at most one place for encode-ahead.
 *
 * Grants go to queued places first, in order, then to parked waiters. An aborted waiter leaves without a grant (its
 * promise rejects with the signal's reason). `release()` is idempotent.
 */
export const MEDIA_QUEUE = 16

export interface MediaAdmission {
  slot(sid: string /* the encode's starter */, mode: 'ahead' | 'request', signal: AbortSignal): Promise<{ release(): void }> | { refused: 'media-busy' }
  /**
   * A request arrived for the session holding the pending `ticket` (fix round 2): its place is switched to the request
   * without giving anything up. `held`: granted, or holding a queue place (now in request mode), or a parked ahead
   * waiter just moved into a free place. `busy`: parked with the queue full; it stays parked as encode-ahead.
   */
  upgrade(ticket: Promise<{ release(): void }>): 'held' | 'busy'
  readonly running: number
  readonly queued: number
}

interface Waiter { sid: string; mode: 'ahead' | 'request'; grant(): void; signal: AbortSignal; onAbort(): void }

const abortError = (signal: AbortSignal) => signal.reason ?? new DOMException('This operation was aborted', 'AbortError')

export function createMediaAdmission(o: { slots: number; queue: number }): MediaAdmission {
  if (!Number.isInteger(o.slots) || o.slots < 1) throw new RangeError('media admission: slots must be a positive integer')
  if (!Number.isInteger(o.queue) || o.queue < 0) throw new RangeError('media admission: queue must be a non-negative integer')
  let running = 0
  const queue: Waiter[] = []
  const parked: Waiter[] = []
  const tickets = new WeakMap<Promise<{ release(): void }>, Waiter>()

  const lease = () => {
    running++
    let released = false
    return { release() { if (released) return; released = true; running--; pump() } }
  }
  const pump = () => {
    while (running < o.slots) {
      const w = queue.shift() ?? parked.shift()
      if (!w) break
      w.signal.removeEventListener('abort', w.onAbort)
      w.grant()
    }
    // A freed place goes to the oldest parked waiter whose session holds none.
    while (queue.length < o.queue) {
      const i = parked.findIndex((p) => !queue.some((q) => q.sid === p.sid))
      if (i < 0) break
      queue.push(parked.splice(i, 1)[0]!)
    }
  }
  const wait = (sid: string, mode: 'ahead' | 'request', signal: AbortSignal, into: Waiter[]) => {
    let made!: Waiter
    const p = new Promise<{ release(): void }>((ok, fail) => {
      const w: Waiter = {
        sid, mode, signal,
        grant: () => ok(lease()),
        onAbort: () => {
          for (const list of [queue, parked]) { const i = list.indexOf(w); if (i >= 0) list.splice(i, 1) }
          fail(abortError(signal))
          pump()
        },
      }
      signal.addEventListener('abort', w.onAbort, { once: true })
      into.push(w)
      made = w
    })
    tickets.set(p, made)
    return p
  }

  return {
    slot(sid, mode, signal) {
      if (signal.aborted) return Promise.reject(abortError(signal))
      if (running < o.slots && queue.length === 0 && (mode === 'request' || parked.length === 0)) return Promise.resolve(lease())
      if (mode === 'request') {
        if (queue.length >= o.queue) return { refused: 'media-busy' }
        return wait(sid, mode, signal, queue)
      }
      const holds = queue.some((q) => q.sid === sid)
      return wait(sid, mode, signal, !holds && queue.length < o.queue ? queue : parked)
    },
    upgrade(ticket) {
      const w = tickets.get(ticket)
      if (!w || queue.includes(w)) { if (w) w.mode = 'request'; return 'held' }
      const i = parked.indexOf(w)
      if (i < 0) return 'held' // granted (or gone)
      if (queue.length >= o.queue) return 'busy'
      parked.splice(i, 1)
      w.mode = 'request'
      queue.push(w)
      pump()
      return 'held'
    },
    get running() { return running },
    get queued() { return queue.length },
  }
}
