/** 3c Task 4: show slide cache and unit request budget (§12a.7). Pure: fake clock, fake io. */
import { describe, expect, it, vi, type Mock } from 'vitest'
import { createShowFetcher, type ImageBitmapLike } from '../../src/stream/remote/show-fetch'
import { StreamSessionError } from '../../src/stream/remote/errors'
// The real server rate gates (server/src/rate.ts), loaded by a non-literal path so
// the client typecheck does not walk the Bun-only server graph. The module imports
// the notes/print renderers only for two per-minute constants; those pull Bun-only
// modules, so stand in for them here ('unit' gating does not use them).
const RATE_MODULE = '../../server/src/rate'
const loadSessionRates = async () => (await import(/* @vite-ignore */ RATE_MODULE) as {
  SessionRates: new () => { check(kind: 'unit', now: number): 'ok' | 'refuse' | 'close' }
}).SessionRates
vi.mock('../../server/src/notes', () => ({ NOTES_RATE_PER_MIN: 30 }))
vi.mock('../../server/src/print', () => ({ PRINT_RATE_PER_MIN: 10 }))

const flush = () => new Promise<void>((r) => setTimeout(r, 0))

function fakeClock() {
  let t = 0
  let seq = 0
  const timers: { at: number; seq: number; cb: () => void }[] = []
  return {
    now: () => t,
    wake(ms: number, cb: () => void) {
      const timer = { at: t + ms, seq: seq++, cb }
      timers.push(timer)
      return () => { const i = timers.indexOf(timer); if (i >= 0) timers.splice(i, 1) }
    },
    async advance(ms: number) {
      const end = t + ms
      for (;;) {
        timers.sort((a, b) => a.at - b.at || a.seq - b.seq)
        const next = timers[0]
        if (!next || next.at > end) break
        timers.shift()
        t = Math.max(t, next.at)
        next.cb()
        await flush()
      }
      t = end
      await flush()
    },
  }
}

interface Call {
  unit: number; scale: number; dpr: number; signal: AbortSignal; at: number
  resolve(b?: ImageBitmapLike): ImageBitmapLike; reject(e: unknown): void
}
type Bmp = ImageBitmapLike & { close: Mock<() => void> }
const bitmap = (width = 100): Bmp => ({ width, height: width / 2, close: vi.fn<() => void>() })

function fakeIo(clock: ReturnType<typeof fakeClock>, autoMs?: number) {
  const calls: Call[] = []
  return {
    calls,
    paint(unit: number, scale: number, dpr: number, signal: AbortSignal, charge?: () => void) {
      charge?.() // one server-charged send, at once
      return new Promise<ImageBitmapLike>((res, rej) => {
        const call: Call = {
          unit, scale, dpr, signal, at: clock.now(),
          resolve: (b = bitmap(100 * scale * dpr)) => { res(b); return b },
          reject: rej,
        }
        calls.push(call)
        if (autoMs !== undefined) clock.wake(autoMs, () => { call.resolve() })
      })
    },
  }
}

function setup(opts?: { budgetPerMin?: number; maxCached?: number }, autoMs?: number) {
  const clock = fakeClock()
  const io = fakeIo(clock, autoMs)
  const f = createShowFetcher(io, clock, opts)
  const ready: number[] = []
  const readyAt: [number, number][] = []
  const errors: [number, string][] = []
  f.onReady((u) => { ready.push(u); readyAt.push([u, clock.now()]) })
  f.onError((u, c) => errors.push([u, c]))
  f.setBucket(1, 1)
  return { clock, io, f, ready, readyAt, errors }
}

/** Largest number of issue times in any closed 60 s window [s, s + 60 000]. */
function maxInWindow(times: number[]) {
  let max = 0
  for (const s of times) max = Math.max(max, times.filter((t) => t >= s && t <= s + 60_000).length)
  return max
}

describe('createShowFetcher', () => {
  it('issues the target before any prefetch, and prefetch once the target settles', async () => {
    const { clock, io, f, ready } = setup()
    f.want([5, 6, 7, 4])
    expect(io.calls.map((c) => c.unit)).toEqual([5])
    io.calls[0]!.resolve()
    await flush()
    expect(ready).toEqual([5])
    await clock.advance(1_000) // R14: prefetch after the wanted set is unchanged 1 s
    expect(io.calls.map((c) => c.unit)).toEqual([5, 6, 7, 4])
    expect(f.cached(5)).toBeDefined()
  })

  it('a new target jumps ahead of prefetch still queued', async () => {
    const { clock, io, f } = setup()
    f.want([0, 1])
    await clock.advance(400) // R14: past the target gap
    f.want([9, 10])
    expect(io.calls.map((c) => c.unit)).toEqual([0, 9])
    expect(io.calls[0]!.signal.aborted).toBe(true)
  })

  it('200 want() calls of fast stepping over 30 s on a 40-unit deck stay ≤ 48 requests in any 60 s window', async () => {
    const { clock, io, f } = setup(undefined, 80)
    for (let i = 0; i < 200; i++) {
      const t = i % 40
      f.want([t, (t + 1) % 40, (t + 2) % 40, (t + 39) % 40])
      await clock.advance(150)
    }
    await clock.advance(100_000) // idle long past the window: wakes keep pacing
    const times = io.calls.map((c) => c.at)
    expect(times.length).toBeGreaterThan(48) // the budget refilled and fetching went on
    expect(maxInWindow(times)).toBeLessThanOrEqual(48)
    // The last target is eventually fetched without another want().
    expect(f.cached(199 % 40)).toBeDefined()
  })

  describe('R14 pacing', () => {
    it('(a) held → for 5 s at 20 steps/s then stop on k: k issued ≤ 650 ms after the last want, ready on resolve, ≤ 48 per window', async () => {
      const { clock, io, f, readyAt } = setup(undefined, 80)
      const k = 99
      for (let i = 0; i <= k; i++) {
        if (i) await clock.advance(50)
        f.want([i, i + 1, i + 2, Math.max(0, i - 1)])
      }
      const lastWant = clock.now()
      await clock.advance(5_000)
      const callK = io.calls.find((c) => c.unit === k)
      expect(callK).toBeDefined()
      expect(callK!.at - lastWant).toBeLessThanOrEqual(650)
      expect(readyAt.find(([u]) => u === k)?.[1]).toBe(callK!.at + 80)
      expect(maxInWindow(io.calls.map((c) => c.at))).toBeLessThanOrEqual(48)
    })

    it('(b) paced navigation (a step every 2 s): each target issued at once, prefetch 1 s after the step', async () => {
      const { clock, io, f } = setup(undefined, 80)
      for (let step = 0; step < 5; step++) {
        const t = step * 4 + 1 // jumps of 4: no target is already prefetched
        const at = clock.now()
        const before = io.calls.length
        f.want([t, t + 1, t + 2, t - 1])
        expect(io.calls.slice(before).map((c) => [c.unit, c.at])).toEqual([[t, at]])
        await clock.advance(999)
        expect(io.calls.length).toBe(before + 1) // target ready at +80, prefetch still held
        await clock.advance(1)
        expect(io.calls.slice(before + 1).map((c) => [c.unit, c.at])).toEqual([[t + 1, at + 1_000], [t + 2, at + 1_000], [t - 1, at + 1_000]])
        await clock.advance(1_000)
      }
    })

    it('(c) a single step after idle is issued with zero delay, also after a burst', async () => {
      const { clock, io, f } = setup(undefined, 80)
      await clock.advance(10_000)
      f.want([3, 4])
      expect(io.calls.map((c) => [c.unit, c.at])).toEqual([[3, 10_000]])
      for (let i = 10; i < 30; i++) { f.want([i, i + 1]); await clock.advance(50) }
      await clock.advance(5_000)
      const at = clock.now()
      f.want([50, 51])
      expect(io.calls.at(-1)).toMatchObject({ unit: 50, at })
    })
  })

  it('with the budget at the reserve, prefetch is withheld but the target is still issued', async () => {
    const { clock, io, f } = setup({ budgetPerMin: 10 })
    f.want([0]); io.calls[0]!.resolve(); await clock.advance(1_000)
    f.want([1]); io.calls[1]!.resolve(); await clock.advance(1_000)
    // 8 left = the reserve
    f.want([2, 3, 4]); expect(io.calls.map((c) => c.unit)).toEqual([0, 1, 2])
    io.calls[2]!.resolve(); await clock.advance(1_000)
    expect(io.calls).toHaveLength(3) // prefetch 3, 4 withheld
    f.want([5, 6]); expect(io.calls.map((c) => c.unit)).toEqual([0, 1, 2, 5])
  })

  it('aborts requests for units dropped from want and closes a result that arrives anyway', async () => {
    const { clock, io, f, ready } = setup()
    f.want([0, 1, 2]); io.calls[0]!.resolve(); await clock.advance(1_000)
    const [c1, c2] = [io.calls[1]!, io.calls[2]!]
    f.want([3])
    expect(c1.signal.aborted).toBe(true)
    expect(c2.signal.aborted).toBe(true)
    const late = c1.resolve(bitmap()) as Bmp
    c2.reject(new DOMException('aborted', 'AbortError'))
    await flush()
    expect(late.close).toHaveBeenCalledOnce()
    expect(ready).toEqual([0])
    expect(f.cached(1)).toBeUndefined()
  })

  it('A→B→A: the aborted first A resolving late is closed; one bitmap kept, one onReady', async () => {
    const { clock, io, f, ready } = setup()
    f.want([0]); await clock.advance(400); f.want([1]); await clock.advance(400); f.want([0])
    expect(io.calls.map((c) => c.unit)).toEqual([0, 1, 0])
    const kept = io.calls[2]!.resolve(bitmap()) as Bmp
    const stale = io.calls[0]!.resolve(bitmap()) as Bmp
    await flush()
    expect(stale.close).toHaveBeenCalledOnce()
    expect(kept.close).not.toHaveBeenCalled()
    expect(f.cached(0)).toBe(kept)
    expect(ready).toEqual([0])
  })

  it('a bucket change re-requests the current unit; cached() returns the old bitmap meanwhile, closed once replaced', async () => {
    const { clock, io, f, ready } = setup()
    f.want([0]); const old = io.calls[0]!.resolve(bitmap(100)) as Bmp; await clock.advance(400)
    f.setBucket(2, 1)
    expect(io.calls.map((c) => [c.unit, c.scale, c.dpr])).toEqual([[0, 1, 1], [0, 2, 1]])
    expect(f.cached(0)).toBe(old)
    const neu = io.calls[1]!.resolve(bitmap(200))
    await flush()
    expect(f.cached(0)).toBe(neu)
    expect(old.close).toHaveBeenCalledOnce()
    expect(ready).toEqual([0, 0])
  })

  it('a bucket change aborts in-flight requests at the old bucket and drops their late results', async () => {
    const { clock, io, f, ready } = setup()
    f.want([0])
    f.setBucket(1, 2)
    expect(io.calls[0]!.signal.aborted).toBe(true)
    const stale = io.calls[0]!.resolve(bitmap()) as Bmp
    await flush()
    expect(stale.close).toHaveBeenCalledOnce()
    expect(ready).toEqual([])
    await clock.advance(250) // R14: the re-request of a just-issued target waits for it to settle
    expect(io.calls.map((c) => c.dpr)).toEqual([1, 2])
  })

  it('a unit error fires onError once, with no retry storm: retried only on a later want ≥ 2 s on', async () => {
    const { clock, io, f, errors } = setup()
    f.want([0, 1])
    io.calls[0]!.reject(new StreamSessionError('unit-failed'))
    await flush()
    expect(errors).toEqual([[0, 'unit-failed']])
    // The failed target counts as settled: prefetch proceeds (after R14's 1 s).
    await clock.advance(1_000)
    expect(io.calls.map((c) => c.unit)).toEqual([0, 1])
    f.want([0, 1]); await clock.advance(999); f.want([0, 1])
    expect(io.calls.filter((c) => c.unit === 0)).toHaveLength(1)
    await clock.advance(10_000) // no automatic retry on any wake
    expect(io.calls.filter((c) => c.unit === 0)).toHaveLength(1)
    f.want([0, 1])
    expect(io.calls.filter((c) => c.unit === 0)).toHaveLength(2)
    expect(errors).toHaveLength(1)
  })

  it('maps a non-session rejection to unit-failed', async () => {
    const { io, f, errors } = setup()
    f.want([3]); io.calls[0]!.reject(new Error('boom')); await flush()
    expect(errors).toEqual([[3, 'unit-failed']])
  })

  it('a rate-limited refusal pauses all issuing ≥ 2 s, fires onError once, then re-issues the unit', async () => {
    const { clock, io, f, errors } = setup()
    f.want([0, 1])
    io.calls[0]!.reject(new StreamSessionError('rate-limited'))
    await flush()
    expect(errors).toEqual([[0, 'rate-limited']])
    expect(io.calls).toHaveLength(1)
    f.want([2, 3]) // another unit inside the pause: not issued either
    await clock.advance(1_999)
    expect(io.calls).toHaveLength(1)
    f.want([0, 1])
    await clock.advance(1) // pause over; target 0 has been the target ≥ 250 ms
    expect(io.calls.map((c) => c.unit)).toEqual([0, 0])
    expect(errors).toHaveLength(1)
  })

  it('a target blocked by the budget is issued by a wake when the window frees, with no further want', async () => {
    const { clock, io, f } = setup({ budgetPerMin: 2 })
    f.want([0]); io.calls[0]!.resolve(); await flush()
    await clock.advance(400)
    f.want([1]); io.calls[1]!.resolve(); await flush()
    f.want([2])
    expect(io.calls).toHaveLength(2)
    await clock.advance(59_600)
    expect(io.calls).toHaveLength(2)
    await clock.advance(1)
    expect(io.calls.map((c) => c.unit)).toEqual([0, 1, 2])
  })

  it('evicts least-recently-used bitmaps beyond maxCached, closing them, never a wanted unit', async () => {
    const { clock, io, f } = setup({ maxCached: 2 })
    const bmps: Bmp[] = []
    f.want([0]); bmps.push(io.calls[0]!.resolve(bitmap()) as Bmp); await clock.advance(400)
    f.want([1, 0]); bmps.push(io.calls[1]!.resolve(bitmap()) as Bmp); await clock.advance(400)
    f.want([2, 0]); bmps.push(io.calls[2]!.resolve(bitmap()) as Bmp); await flush()
    // 0 is older than 1 but still wanted: 1 goes.
    expect(bmps[1]!.close).toHaveBeenCalledOnce()
    expect(bmps[0]!.close).not.toHaveBeenCalled()
    expect(f.cached(0)).toBe(bmps[0])
    expect(f.cached(1)).toBeUndefined()
  })

  it('dispose closes cached bitmaps, aborts in flight, drops late results and never fires onReady after', async () => {
    const { clock, io, f, ready } = setup()
    f.want([0, 1]); const a = io.calls[0]!.resolve(bitmap()) as Bmp; await clock.advance(1_000)
    const inflight = io.calls[1]!
    f.dispose()
    expect(a.close).toHaveBeenCalledOnce()
    expect(inflight.signal.aborted).toBe(true)
    const late = inflight.resolve(bitmap()) as Bmp
    await flush()
    expect(late.close).toHaveBeenCalledOnce()
    expect(ready).toEqual([0])
    f.want([5])
    expect(io.calls).toHaveLength(2)
    expect(f.cached(0)).toBeUndefined()
  })

  it('issues nothing until a bucket is set, and rejects a bad bucket', () => {
    const clock = fakeClock()
    const io = fakeIo(clock)
    const f = createShowFetcher(io, clock)
    f.want([0])
    expect(io.calls).toHaveLength(0)
    expect(() => f.setBucket(0, 1)).toThrow(RangeError)
    expect(() => f.setBucket(1, Number.NaN)).toThrow(RangeError)
    f.setBucket(2, 1)
    expect(io.calls.map((c) => c.scale)).toEqual([2])
  })

  describe('fix round 1: no callback after dispose', () => {
    it('a ready listener that disposes stops the dispatch: no later ready listener runs', async () => {
      const { io, f } = setup()
      const events: string[] = []
      let disposed = false
      f.onReady(() => { events.push('first'); f.dispose(); disposed = true })
      f.onReady(() => { events.push(disposed ? 'AFTER DISPOSE' : 'second') })
      f.want([7])
      io.calls.at(-1)!.resolve()
      await flush()
      expect(events).toEqual(['first'])
    })

    it('an error listener that disposes stops the dispatch: no later error listener runs', async () => {
      const { io, f } = setup()
      const events: string[] = []
      let disposed = false
      f.onError(() => { events.push('first'); f.dispose(); disposed = true })
      f.onError(() => { events.push(disposed ? 'AFTER DISPOSE' : 'second') })
      f.want([7])
      io.calls.at(-1)!.reject(new StreamSessionError('unit-failed'))
      await flush()
      expect(events).toEqual(['first'])
    })

    it('a listener unsubscribed during dispatch is skipped', async () => {
      const { io, f } = setup()
      const events: string[] = []
      let off = () => {}
      f.onReady(() => { events.push('first'); off() })
      off = f.onReady(() => { events.push('second') })
      f.want([7])
      io.calls.at(-1)!.resolve()
      await flush()
      expect(events).toEqual(['first'])
    })

    it('onReady/onError after dispose register nothing and return a callable no-op', () => {
      const { f } = setup()
      f.dispose()
      expect(() => f.onReady(() => {})()).not.toThrow()
      expect(() => f.onError(() => {})()).not.toThrow()
    })
  })

  describe('fix round 1: R16 every server-charged send consumes budget', () => {
    it('each charge() is stamped: two charges per paint halve the requests the budget admits', async () => {
      const clock = fakeClock()
      const calls: number[] = []
      const f = createShowFetcher({
        paint(unit, _s, _d, _sig, charge) { calls.push(unit); charge?.(); charge?.(); return new Promise(() => {}) },
      }, clock, { budgetPerMin: 4 })
      f.setBucket(1, 1)
      for (let i = 0; i < 4; i++) { f.want([i]); await clock.advance(400) }
      expect(calls).toEqual([0, 1])
    })

    it('a request issued but not yet charged holds its budget until it charges', async () => {
      const clock = fakeClock()
      const calls: number[] = []
      const f = createShowFetcher({
        paint(unit) { calls.push(unit); return new Promise(() => {}) }, // waiting on admission: no send yet
      }, clock, { budgetPerMin: 2 })
      f.setBucket(1, 1)
      // 0 stays wanted (as prefetch) and unsent, so it keeps its budget share.
      f.want([0]); await clock.advance(400)
      f.want([1, 0]); await clock.advance(400)
      f.want([2, 1, 0]); await clock.advance(400)
      expect(calls).toEqual([0, 1])
    })

    it('after a rate-limited refusal, no retry into a full window even once 2 s have passed', async () => {
      const clock = fakeClock()
      const calls: { unit: number; at: number; reject(e: unknown): void; resolve(b: ImageBitmapLike): void }[] = []
      const f = createShowFetcher({
        paint(unit, _s, _d, _sig, charge) {
          charge?.()
          return new Promise((resolve, reject) => { calls.push({ unit, at: clock.now(), resolve, reject }) })
        },
      }, clock, { budgetPerMin: 2 })
      f.setBucket(1, 1)
      f.want([0]); calls[0]!.resolve(bitmap()); await clock.advance(400)
      f.want([1]); calls[1]!.reject(new StreamSessionError('rate-limited')); await flush()
      await clock.advance(10_000)
      expect(calls).toHaveLength(2)
      await clock.advance(60_000)
      expect(calls.map((c) => c.unit)).toEqual([0, 1, 1])
    })

    it('reviewer model: every paint charges twice (validation retry), real SessionRates — no close over 60 s of 400 ms navigation', async () => {
      const clock = fakeClock()
      const rates = new (await loadSessionRates())()
      const verdicts: { at: number; result: string }[] = []
      const charges: number[] = []
      const f = createShowFetcher({
        paint(_unit, _s, _d, signal, charge) {
          return new Promise<ImageBitmapLike>((resolve, reject) => {
            const send = () => {
              if (signal.aborted) return false
              charge?.(); charges.push(clock.now())
              const result = rates.check('unit', clock.now())
              if (result !== 'ok') { verdicts.push({ at: clock.now(), result }); reject(new StreamSessionError('rate-limited')); return false }
              return true
            }
            if (!send()) return
            // The first answer fails validation; the session re-sends once, which is charged too.
            clock.wake(40, () => { if (send()) clock.wake(40, () => resolve(bitmap())) })
          })
        },
      }, clock)
      f.setBucket(1, 1)
      for (let i = 0; i < 150; i++) { f.want([i % 40, (i + 1) % 40, (i + 2) % 40, (i + 39) % 40]); await clock.advance(400) }
      await clock.advance(5_000)
      f.want([5, 6, 7, 4]); await clock.advance(70_000) // idle: prefetch with retries
      expect(verdicts.filter((v) => v.result === 'close')).toEqual([])
      expect(verdicts).toEqual([])
      expect(maxInWindow(charges)).toBeLessThanOrEqual(56)
    })
  })
})
