/**
 * S1-2 foundation contracts: the injected clock, the seeded random number
 * generator, the deduplicating diagnostic sink, target keys, and the
 * workspace/lease contract.
 *
 * The workspace implementation is S1-16's (`src/slideshow/workspace.ts`).
 * Here the lease contract is pinned against a reference ledger written in
 * this file, typed by the S1-2 interfaces, so S1-16 inherits the exact
 * semantics: release returns the bytes exactly once, transfer never charges
 * twice, and a SubOwner never reaches the parent ledger.
 */
import { describe, expect, it } from 'vitest'
import { ManualClock, RafClock } from '../src/slideshow/clock'
import { createRng } from '../src/slideshow/rng'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { targetKey } from '../src/slideshow/target'
import type { AllocationOwner, Lease, ShowClock, SlideShowWorkspace, SubOwner, TargetRef, TerminalReserve } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Reference ledger (test-only). Every byte a lease holds is counted in
// `charged`; `events` records every parent-ledger movement.
// ---------------------------------------------------------------------------
interface RefLedger { charged: number; events: string[] }

function makeLease(ledger: RefLedger, pool: { used: number; limit: number }, bytes: number, tag: string, record: boolean): Lease {
  let held = bytes
  let live = true
  const lease: Lease = {
    get tag() { return tag },
    get bytes() { return live ? held : 0 },
    grow(more) {
      if (!live || !(more >= 0) || pool.used + more > pool.limit) return false
      pool.used += more
      held += more
      if (record) { ledger.charged += more; ledger.events.push(`grow ${tag} ${more}`) }
      return true
    },
    shrink(less) {
      if (!live || !(less >= 0)) return
      const n = Math.min(less, held)
      pool.used -= n
      held -= n
      if (record) { ledger.charged -= n; ledger.events.push(`shrink ${tag} ${n}`) }
    },
    transfer(next) {
      if (!live) throw new Error('released lease cannot transfer')
      live = false
      return makeLease(ledger, pool, held, next, record)
    },
    release() {
      if (!live) return
      live = false
      pool.used -= held
      if (record) { ledger.charged -= held; ledger.events.push(`release ${tag} ${held}`) }
      held = 0
    },
  }
  return lease
}

function makeSubOwner(ledger: RefLedger, limit: number): SubOwner {
  const pool = { used: 0, limit }
  return {
    lease(bytes, tag) {
      if (!(bytes >= 0) || pool.used + bytes > pool.limit) return undefined
      pool.used += bytes
      return makeLease(ledger, pool, bytes, tag, false)
    },
  }
}

function makeReferenceWorkspace(limit: number, ledger: RefLedger): SlideShowWorkspace {
  const pool = { used: 0, limit }
  return {
    limit,
    lease(bytes, tag) {
      if (!(bytes >= 0) || pool.used + bytes > pool.limit) return undefined
      pool.used += bytes
      ledger.charged += bytes
      ledger.events.push(`lease ${tag} ${bytes}`)
      return makeLease(ledger, pool, bytes, tag, true)
    },
    reserveTerminal(outputBytes) {
      const scratchBytes = 64
      const base = this.lease(outputBytes + scratchBytes, 'terminal')
      if (!base) return undefined
      let out = outputBytes
      const reserve: TerminalReserve = {
        get outputBytes() { return out },
        scratch: makeSubOwner(ledger, scratchBytes),
        resize(next) {
          if (next > out && !base.grow(next - out)) return false
          if (next < out) base.shrink(out - next)
          out = next
          return true
        },
        release() { base.release() },
      }
      return reserve
    },
    used: () => pool.used,
  }
}

describe('slideshow foundation contracts', () => {
  it('ManualClock advances only when told and stands still while paused', () => {
    const clock = new ManualClock()
    const asShow: ShowClock = clock
    expect(asShow.now()).toBe(0)
    expect(asShow.paused).toBe(false)
    expect(clock.now()).toBe(0) // reading never advances
    clock.advance(16)
    expect(clock.now()).toBe(16)
    clock.pause()
    expect(clock.paused).toBe(true)
    clock.advance(1000)
    expect(clock.now()).toBe(16)
    clock.resume()
    expect(clock.paused).toBe(false)
    clock.advance(4)
    expect(clock.now()).toBe(20)
    expect(() => clock.advance(-1)).toThrow(RangeError)
    expect(() => clock.advance(Number.NaN)).toThrow(RangeError)
    expect(clock.now()).toBe(20)

    // RafClock reads the rAF timestamps of the window it is given (realm-safe),
    // never the global one; it starts at 0 and stands still while paused.
    const frames = new Map<number, (t: number) => void>()
    let nextId = 0
    const win = {
      requestAnimationFrame: (cb: (t: number) => void) => { frames.set(++nextId, cb); return nextId },
      cancelAnimationFrame: (id: number) => { frames.delete(id) },
    } as unknown as Window
    const raf = new RafClock(win)
    const tick = (t: number) => { const [id, cb] = frames.entries().next().value!; frames.delete(id); cb(t) }
    tick(1000)
    expect(raf.now()).toBe(0)
    tick(1016)
    expect(raf.now()).toBe(16)
    raf.pause()
    tick(2000)
    expect(raf.now()).toBe(16)
    raf.resume()
    tick(2010)
    expect(raf.now()).toBe(26)
    raf.dispose()
    expect(frames.size).toBe(0)
  })

  it('a lease releases exactly once; transfer does not double-charge; a SubOwner never reaches the parent', () => {
    const ledger: RefLedger = { charged: 0, events: [] }
    const ws = makeReferenceWorkspace(1000, ledger)
    const owner: AllocationOwner = ws

    const a = owner.lease(300, 'layer:a')!
    expect(a.bytes).toBe(300)
    expect(ws.used()).toBe(300)
    expect(ledger.charged).toBe(300)

    // Over the limit: refused, nothing charged.
    expect(ws.lease(701, 'too-big')).toBeUndefined()
    expect(ws.used()).toBe(300)

    // grow and shrink move the same lease.
    expect(a.grow(100)).toBe(true)
    expect(a.grow(10_000)).toBe(false)
    a.shrink(50)
    expect(a.bytes).toBe(350)
    expect(ws.used()).toBe(350)

    // transfer: one charge, new tag, the old lease is dead.
    const b = a.transfer('texture:a')
    expect(b.tag).toBe('texture:a')
    expect(b.bytes).toBe(350)
    expect(a.bytes).toBe(0)
    expect(ws.used()).toBe(350)
    expect(ledger.charged).toBe(350)
    a.release() // releasing the transferred-from lease returns nothing
    expect(ws.used()).toBe(350)

    // release is idempotent and returns the bytes exactly once.
    b.release()
    b.release()
    expect(ws.used()).toBe(0)
    expect(ledger.charged).toBe(0)
    expect(ledger.events.filter((e) => e.startsWith('release')).length).toBe(1)

    // Terminal reserve: the scratch SubOwner leases inside bytes already charged.
    const terminal = ws.reserveTerminal(200)!
    const chargedWithReserve = ledger.charged
    const eventsBefore = ledger.events.length
    const scratch: AllocationOwner = terminal.scratch
    const s1 = scratch.lease(40, 'scratch:decode')!
    expect(scratch.lease(30, 'scratch:over')).toBeUndefined() // bounded by the sub-owner
    s1.release()
    s1.release()
    expect(ledger.charged).toBe(chargedWithReserve)
    expect(ledger.events.length).toBe(eventsBefore) // never reached the parent
    expect(terminal.resize(300)).toBe(true)
    expect(terminal.outputBytes).toBe(300)
    terminal.release()
    terminal.release()
    expect(ws.used()).toBe(0)
    expect(ledger.charged).toBe(0)
  })

  it('createRng is deterministic per seed and differs across seeds', () => {
    const take = (seed: number) => { const r = createRng(seed); return Array.from({ length: 8 }, () => r()) }
    const a = take(42)
    expect(take(42)).toEqual(a)
    expect(take(43)).not.toEqual(a)
    for (const v of [...a, ...take(0), ...take(0xffffffff)]) {
      expect(v).toBeGreaterThanOrEqual(0)
      expect(v).toBeLessThan(1)
    }
    // mulberry32 reference value for seed 1.
    expect(createRng(1)()).toBeCloseTo(0.6270739405881613, 15)
  })

  it('DiagnosticSink deduplicates per code, slide and effect', () => {
    const sink = createDiagnosticSink()
    sink.add({ code: 'stand-in', slideIndex: 1, effect: 'vortex', message: 'first' })
    sink.add({ code: 'stand-in', slideIndex: 1, effect: 'vortex', message: 'again' })
    sink.add({ code: 'stand-in', slideIndex: 2, effect: 'vortex', message: 'other slide' })
    sink.add({ code: 'stand-in', slideIndex: 1, effect: 'shred', message: 'other effect' })
    sink.add({ code: 'approximate', slideIndex: 1, effect: 'vortex', message: 'other code' })
    sink.add({ code: 'no-package', message: 'deck-level' })
    sink.add({ code: 'no-package', message: 'deck-level again' })
    // An undefined slide or effect is not confused with a literal value.
    sink.add({ code: 'no-package', effect: 'undefined', message: 'effect named undefined' })
    sink.add({ code: 'no-package', slideIndex: 0, message: 'slide 0' })
    const list = sink.list()
    expect(list.map((d) => d.message)).toEqual(['first', 'other slide', 'other effect', 'other code', 'deck-level', 'effect named undefined', 'slide 0'])
  })

  it('targetKey is distinct for shape, paragraph range, char range and bg', () => {
    const refs: TargetRef[] = [
      { spid: '4' },
      { spid: '4', text: { kind: 'p', st: 0, end: 0 } },
      { spid: '4', text: { kind: 'p', st: 0, end: 1 } },
      { spid: '4', text: { kind: 'char', st: 0, end: 0 } },
      { spid: '4', sub: 'bg' },
      { spid: '4', sub: 'chart' },
      { spid: '40' },
      { spid: '4/p:0-0' }, // hostile spid cannot forge a range key
      { bg: true },
      { sld: true },
      { snd: 'rId1' },
    ]
    const keys = refs.map(targetKey)
    expect(new Set(keys).size).toBe(refs.length)
    for (const k of keys) expect(typeof k).toBe('string')
    // Stable: the same ref always yields the same key.
    expect(targetKey({ spid: '4', text: { kind: 'p', st: 2, end: 3 } })).toBe(targetKey({ spid: '4', text: { kind: 'p', st: 2, end: 3 } }))
  })
})
