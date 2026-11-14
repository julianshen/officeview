/**
 * Piece 4 Task 12: rate behaviour under animation. Bundles, keyframe tracks and Morph pairs are each one charged
 * `'unit'` send on the session's one budget. Pure: fake clock, fake io, the real server rate gates (4008) and the real
 * `UnitSendLog`; the fetcher's `room(kind)` is what the show asks before a track or a Morph pair.
 */
import { describe, expect, it, vi } from 'vitest'
import { createShowFetcher, MORPH_RESERVE, RESERVE, TRACK_RESERVE } from '../../src/stream/remote/show-fetch'
import { StreamSessionError } from '../../src/stream/remote/errors'
import { UNIT_FIRST_CAP, UNIT_SPAN_MS, UnitSendLog } from '../../src/stream/remote/unit-budget'
import type { ShowBundle } from '../../src/stream/remote/show-bundle'

const RATE_MODULE = '../../server/src/rate'
type Rates = { check(kind: 'unit', now: number): 'ok' | 'refuse' | 'close' }
const loadSessionRates = async () => (await import(/* @vite-ignore */ RATE_MODULE) as { SessionRates: new () => Rates }).SessionRates
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

type Kind = 'target' | 'prefetch' | 'track' | 'morph'
const UNITS = 40
/** A third of a second of server latency per request. */
const LATENCY_MS = 150

/** The session as the show sees it: one `UnitSendLog` (62 s span) stamped on every charged send, over the real gates. */
async function harness() {
  const clock = fakeClock()
  const rates = new (await loadSessionRates())()
  const log = new UnitSendLog(clock.now, UNIT_FIRST_CAP, UNIT_SPAN_MS)
  const sends: { at: number; kind: Kind; roomBefore: number }[] = []
  const verdicts: string[] = []
  const issued = new Set<number>()
  let wanted: number[] = []
  const stamp = (kind: Kind) => {
    sends.push({ at: clock.now(), kind, roomBefore: log.room() })
    log.record()
    const v = rates.check('unit', clock.now())
    if (v !== 'ok') verdicts.push(v)
  }
  const bundles: { unit: number; kind: Kind; at: number }[] = []
  const priorityBreaks: string[] = []
  const io = {
    paint: () => Promise.reject(new Error('every unit is animated')),
    budget: log,
    bundle(unit: number, _scale: number, _dpr: number, _m: 'full' | 'reduced', _signal: AbortSignal, charge: () => void) {
      const kind: Kind = unit === wanted[0] ? 'target' : 'prefetch'
      // The target bundle goes out before any prefetch.
      if (kind === 'prefetch' && wanted[0] !== undefined && !issued.has(wanted[0])) priorityBreaks.push(`prefetch of ${unit} before target ${wanted[0]}`)
      issued.add(unit)
      stamp(kind)
      charge()
      bundles.push({ unit, kind, at: clock.now() })
      return new Promise<ShowBundle>((res) => {
        clock.wake(LATENCY_MS, () => res({ unit, plan: {} as ShowBundle['plan'], assets: { bitmap: () => undefined, bytes: () => 0 }, decodedBytes: 1000, release() {} }))
      })
    },
  }
  const f = createShowFetcher(io, clock, { animated: () => true, maxCached: 12 })
  f.setBucket(1, 1)
  return {
    clock, f, log, sends, verdicts, bundles, priorityBreaks, stamp,
    want(units: number[]) { wanted = units; f.want(units) },
  }
}

/** The server's own counting: most sends inside any 60 s window. */
const maxInWindow = (times: number[]) => {
  let max = 0
  for (const s of times) max = Math.max(max, times.filter((t) => t > s - 60_000 && t <= s).length)
  return max
}

describe('Task 12: animated show fetching stays inside the unit budget', () => {
  it('the reserves are strictly ordered: target (0) < track < Morph pair < prefetch', () => {
    expect(0).toBeLessThan(TRACK_RESERVE)
    expect(TRACK_RESERVE).toBeLessThan(MORPH_RESERVE)
    expect(MORPH_RESERVE).toBeLessThan(RESERVE)
  })

  it('holding next over 40 animated slides stays under the unit budget', async () => {
    const h = await harness()
    let slide = -1
    let step = 0
    let waiting: number | undefined
    let deferredTracks = 0
    let deferredMorphs = 0
    let elapsed = 0
    const wantAround = (s: number) => h.want([s, s + 1, s + 2, s - 1].filter((u) => u >= 0 && u < UNITS))
    const enter = (s: number) => {
      slide = s; step = 0
      // D4: the tracks of the first two composited steps, each one charged send, only while the budget has room.
      for (let k = 0; k < 2; k++) { if (h.f.room('track')) h.stamp('track'); else deferredTracks++ }
    }
    wantAround(0); waiting = 0
    // → held down: one press every 100 ms until the last slide, over a 40-slide deck (a Morph into every 5th slide).
    for (let ms = 0; ms < 300_000; ms += 100) {
      await h.clock.advance(100)
      if (waiting !== undefined) { if (h.f.bundle(waiting)) { enter(waiting); waiting = undefined } continue }
      if (step < 2) { step++; continue }
      const n = slide + 1
      if (n >= UNITS) break
      elapsed = ms
      if (n % 5 === 0) { if (h.f.room('morph')) h.stamp('morph'); else deferredMorphs++ }
      wantAround(n); waiting = n
    }
    await h.clock.advance(70_000)

    expect(h.verdicts).toEqual([]) // no refusal, so never a 4008
    expect(maxInWindow(h.sends.map((s) => s.at))).toBeLessThanOrEqual(UNIT_FIRST_CAP)
    expect(h.priorityBreaks).toEqual([])
    // Each class asked for room above its own reserve: the target for any, a track above TRACK_RESERVE, …
    const floor: Record<Kind, number> = { target: 1, track: TRACK_RESERVE + 1, morph: MORPH_RESERVE + 1, prefetch: RESERVE + 1 }
    for (const s of h.sends) expect(s.roomBefore, `${s.kind} at ${s.at}`).toBeGreaterThanOrEqual(floor[s.kind])
    // The hold was real: it was slowed by the budget, deferred optional work, and still got through the deck.
    expect(slide).toBe(UNITS - 1)
    expect(elapsed).toBeGreaterThan(60_000) // 40 slides at 3 presses of 100 ms would be 12 s: the budget slowed it
    expect(deferredTracks + deferredMorphs).toBeGreaterThan(0)
    expect(h.sends.some((s) => s.kind === 'track')).toBe(true)
    expect(h.sends.filter((s) => s.kind === 'target').length).toBeGreaterThan(5)
  })

  it('budget at reserve: tracks step-deferred, Morph faded, target still issued', async () => {
    const h = await harness()
    // Leave exactly TRACK_RESERVE sends of room.
    for (let i = 0; i < UNIT_FIRST_CAP - TRACK_RESERVE; i++) h.log.record()
    expect(h.log.room()).toBe(TRACK_RESERVE)
    expect(h.f.room('track')).toBe(false)
    expect(h.f.room('morph')).toBe(false)
    h.want([7, 8, 9])
    await h.clock.advance(2_000)
    // The target went out on the reserve; prefetch did not.
    expect(h.bundles.map((b) => [b.unit, b.kind])).toEqual([[7, 'target']])
    expect(h.log.room()).toBe(TRACK_RESERVE - 1)
    expect(h.f.room('track')).toBe(false)
    expect(h.verdicts).toEqual([])
  })

  it('the reserves close in order as the budget runs down: prefetch, then Morph pair, then track, then the target', async () => {
    const h = await harness()
    h.want([0])
    await h.clock.advance(200)
    const fill = (room: number) => { while (h.log.room() > room) h.log.record() }
    fill(RESERVE)
    expect(h.f.room('morph')).toBe(true)
    expect(h.f.room('track')).toBe(true)
    // At RESERVE prefetch is closed (more than RESERVE is needed): the wanted neighbours are never asked.
    h.want([0, 1, 2])
    await h.clock.advance(3_000)
    expect(h.bundles.map((b) => b.unit)).toEqual([0])
    fill(MORPH_RESERVE)
    expect(h.f.room('morph')).toBe(false)
    expect(h.f.room('track')).toBe(true)
    fill(TRACK_RESERVE)
    expect(h.f.room('track')).toBe(false)
    // The target still goes out on what is left.
    h.want([5])
    await h.clock.advance(500)
    expect(h.bundles.map((b) => b.unit)).toEqual([0, 5])
  })

  it('no track or Morph pair takes the slot of a target that has yet to go out', async () => {
    const h = await harness()
    h.want([0])
    await h.clock.advance(200)
    expect(h.f.room('track')).toBe(true)
    // A new target 200 ms after the last one is paced (R14): it waits to go out.
    h.want([1])
    expect(h.bundles).toHaveLength(1)
    expect(h.f.room('track')).toBe(false)
    expect(h.f.room('morph')).toBe(false)
    await h.clock.advance(300)
    expect(h.bundles.map((b) => b.unit)).toEqual([0, 1])
    expect(h.f.room('track')).toBe(true)
  })

  it('a rate-limited refusal pauses issuing 2 s', async () => {
    const clock = fakeClock()
    const log = new UnitSendLog(clock.now, UNIT_FIRST_CAP, UNIT_SPAN_MS)
    const asked: [number, number][] = []
    let refuse = true
    const f = createShowFetcher({
      paint: () => Promise.reject(new Error('no')),
      budget: log,
      bundle(unit, _s, _d, _m, _sig, charge) {
        asked.push([unit, clock.now()]); charge(); log.record()
        if (refuse) { refuse = false; return Promise.reject(new StreamSessionError('rate-limited')) }
        return Promise.resolve({ unit, plan: {} as ShowBundle['plan'], assets: { bitmap: () => undefined, bytes: () => 0 }, decodedBytes: 1, release() {} })
      },
    }, clock, { animated: () => true })
    f.setBucket(1, 1)
    f.want([0, 1])
    await flush()
    expect(asked).toEqual([[0, 0]])
    await clock.advance(1_900)
    expect(asked).toEqual([[0, 0]])
    expect(f.room('track')).toBe(false)
    expect(f.room('morph')).toBe(false)
    await clock.advance(200)
    expect(asked).toEqual([[0, 0], [0, 2_000], [1, 2_000]]) // the refused target again, then its prefetch
    // A track or Morph pair refused rate-limited (the show reports it) pauses the same way.
    f.paused()
    expect(f.room('track')).toBe(false)
    await clock.advance(2_100)
    expect(f.room('track')).toBe(true)
  })
})
