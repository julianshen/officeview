/**
 * S1-16: the slideshow workspace, the fallback ladder and the permanent reserve
 * (slideshow spec 7). Every budget here is injected (profile plus limit
 * overrides); nothing reads device memory.
 */
import { describe, expect, it } from 'vitest'
import {
  LAYER_LIMIT, RERASTER_LIMIT, SLIDESHOW_WORKSPACE_LIMIT, TERMINAL_SCRATCH,
  admitShow, createFrameLimits, createSlideShowWorkspace, surfaceBytes, surfaceLease, transitionLeases, uploadLeases,
} from '../src/slideshow/workspace'
import { createResidency, type LadderLevel } from '../src/slideshow/residency'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import type { AllocationOwner, DiagnosticSink, MetadataHandle, ShowSlide, SlideMetadataStore, SlideShowDiagnostic } from '../src/slideshow/types'

const MiB = 1024 * 1024

// ---------------------------------------------------------------------------
// A fake metadata store implementing the S1-2 interface. `log` records every
// acquire, handle release and evict, in order.
// ---------------------------------------------------------------------------
interface FakeStore extends SlideMetadataStore {
  log: string[]
  live: Map<number, number>
  /** Pending acquisitions when `deferred`; `settle(i)` resolves slide i's. */
  settle(index: number): void
}
function fakeStore(options: { deferred?: boolean; log?: string[]; failOnce?: number } = {}): FakeStore {
  const log = options.log ?? []
  const live = new Map<number, number>()
  const cache = new Set<number>()
  const pending = new Map<number, Array<() => void>>()
  const slide = (index: number): ShowSlide => ({ index, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] })
  const handle = (index: number): MetadataHandle => {
    live.set(index, (live.get(index) ?? 0) + 1)
    let open = true
    return {
      slide: slide(index),
      attach: () => open,
      release() {
        if (!open) return
        open = false
        live.set(index, live.get(index)! - 1)
        log.push(`release ${index}`)
      },
    }
  }
  return {
    log,
    live,
    settle(index) { const list = pending.get(index) ?? []; pending.delete(index); for (const go of list) go() },
    acquire(index) {
      log.push(`acquire ${index}`)
      if (options.failOnce === index) { options.failOnce = undefined; return Promise.reject(new Error('read failed')) }
      cache.add(index)
      if (!options.deferred) return Promise.resolve(handle(index))
      return new Promise(resolve => { const list = pending.get(index) ?? []; list.push(() => resolve(handle(index))); pending.set(index, list) })
    },
    evict(index) { log.push(`evict ${index}`); cache.delete(index) },
    resident: () => [...cache].sort((a, b) => a - b),
    close() {},
  }
}

/** A sink that keeps every add (no dedup), to prove callers emit once. */
function countingSink(): DiagnosticSink & { all: SlideShowDiagnostic[] } {
  const all: SlideShowDiagnostic[] = []
  return { all, add: d => { all.push(d) }, list: () => all }
}

const codes = (sink: DiagnosticSink) => sink.list().map(d => d.code)
const deferred = () => { let resolve!: () => void; const promise = new Promise<void>(r => { resolve = r }); return { promise, resolve } }
const flush = async () => { for (let i = 0; i < 5; i++) await Promise.resolve() }

describe('slideshow workspace (spec 7)', () => {
  it('refuses one byte below each peak: layer+texture during upload, mirror, transition surfaces', () => {
    expect(SLIDESHOW_WORKSPACE_LIMIT).toEqual({ desktop: 512 * MiB, mobile: 256 * MiB })
    expect(TERMINAL_SCRATCH).toEqual({ desktop: 64 * MiB, mobile: 32 * MiB })
    const w = 800, h = 600
    const S = surfaceBytes(w, h)
    expect(S).toBe(800 * 600 * 4)
    const reserve = 4000 + 1000 // output + scratch
    const peaks: Array<[string, number, (owner: AllocationOwner) => { release(): void } | undefined]> = [
      ['upload', 2 * S, owner => { const u = uploadLeases(owner, w, h, 'layer:1'); return u && { release: () => { u.cpu.release(); u.texture.release() } } }],
      ['mirror', S, owner => surfaceLease(owner, w, h, 'mirror')],
      ['transition', 2 * S, owner => { const t = transitionLeases(owner, w, h, 'transition'); return t && { release: () => { t.from.release(); t.to.release() } } }],
    ]
    for (const [name, peak, take] of peaks) {
      const short = createSlideShowWorkspace('desktop', { limit: reserve + peak - 1, scratch: 1000 })
      expect(short.reserveTerminal(4000), name).toBeDefined()
      expect(take(short), name).toBeUndefined()
      expect(short.used(), `${name}: a refusal charges nothing`).toBe(reserve)

      const exact = createSlideShowWorkspace('desktop', { limit: reserve + peak, scratch: 1000 })
      exact.reserveTerminal(4000)
      const got = take(exact)
      expect(got, name).toBeDefined()
      expect(exact.used(), name).toBe(reserve + peak)
      got!.release()
      expect(exact.used(), name).toBe(reserve)
    }
    // Upload: both copies are charged during upload; the CPU copy goes after it.
    const ws = createSlideShowWorkspace('mobile', { limit: 2 * S, scratch: 0 })
    const u = uploadLeases(ws, w, h, 'layer:2')!
    expect(ws.used()).toBe(2 * S)
    u.cpu.release()
    u.cpu.release() // idempotent
    expect(ws.used()).toBe(S)
    expect(u.texture.bytes).toBe(S)
  })

  it('the terminal reserve is output plus scratch and is never lent to layers, textures, metadata or media', () => {
    const ws = createSlideShowWorkspace('desktop', { limit: 10_000, scratch: 1000 })
    expect(ws.limit).toBe(10_000)
    const terminal = ws.reserveTerminal(4000)!
    expect(terminal.outputBytes).toBe(4000)
    expect(ws.used()).toBe(5000)
    // Nothing reaches into the reserve, whatever the tag.
    for (const tag of ['layer:1', 'texture:1', 'metadata:3', 'media:rId2']) expect(ws.lease(5001, tag), tag).toBeUndefined()
    const all = ws.lease(5000, 'layer:all')!
    expect(ws.used()).toBe(10_000)
    expect(ws.lease(1, 'media:x')).toBeUndefined()
    expect(all.grow(1)).toBe(false)
    // The scratch allowance is a sub-owner of bytes already reserved: it never touches the ledger.
    const decode = terminal.scratch.lease(1000, 'decode')!
    expect(ws.used()).toBe(10_000)
    expect(terminal.scratch.lease(1, 'decode:2')).toBeUndefined()
    decode.release()
    expect(terminal.scratch.lease(1000, 'decode:3')).toBeDefined()
    // A second reserve cannot be squeezed out of leased bytes either.
    expect(ws.reserveTerminal(0)).toBeUndefined()
    // Lease contract (S1-2): release once, transfer without a second charge.
    const moved = all.transfer('layer:moved')
    expect(all.bytes).toBe(0)
    expect(moved.bytes).toBe(5000)
    expect(ws.used()).toBe(10_000)
    expect(() => all.transfer('again')).toThrow()
    all.release()
    expect(ws.used()).toBe(10_000)
    moved.shrink(2000)
    expect(ws.used()).toBe(8000)
    moved.release()
    moved.release()
    expect(ws.used()).toBe(5000)
    expect(ws.lease(-1, 'bad')).toBeUndefined()
    expect(ws.lease(Number.NaN, 'bad')).toBeUndefined()
    // Releasing the reserve refunds output and scratch together and closes the scratch owner.
    terminal.release()
    terminal.release()
    expect(ws.used()).toBe(0)
    expect(terminal.scratch.lease(1, 'late')).toBeUndefined()
    expect(terminal.resize(10)).toBe(false)
    expect(() => createSlideShowWorkspace('tablet' as 'desktop')).toThrow(RangeError)
  })

  it('reserve at 3840x2160 is 33177600 plus scratch; failing that steps the DPR down; failing DPR 1 refuses', () => {
    const ws = createSlideShowWorkspace('desktop')
    expect(ws.limit).toBe(512 * MiB)
    const ok = admitShow(ws, { w: 1920, h: 1080 }, 3) // DPR capped at 2
    expect(ok.ok).toBe(true)
    if (!ok.ok) return
    expect(ok.backing).toEqual({ w: 3840, h: 2160, dpr: 2 })
    expect(ok.terminal.outputBytes).toBe(33_177_600)
    expect(ws.used()).toBe(33_177_600 + 64 * MiB)

    // The 3840 x 2160 cap applies even below DPR 2.
    const wide = admitShow(createSlideShowWorkspace('desktop'), { w: 2560, h: 1440 }, 2)
    expect(wide.ok && wide.backing).toEqual({ w: 3840, h: 2160, dpr: 1.5 })

    // One byte short of the full reserve: the DPR steps down, never below 1.
    const short = createSlideShowWorkspace('desktop', { limit: 33_177_600 + 64 * MiB - 1 })
    const stepped = admitShow(short, { w: 1920, h: 1080 }, 2)
    expect(stepped.ok).toBe(true)
    if (!stepped.ok) return
    expect(stepped.backing).toEqual({ w: 3360, h: 1890, dpr: 1.75 })
    expect(stepped.terminal.outputBytes).toBe(3360 * 1890 * 4)
    expect(short.used()).toBe(3360 * 1890 * 4 + 64 * MiB)

    // Exactly DPR 1 at the CSS size fits.
    const floor = 1920 * 1080 * 4 + 32 * MiB
    const one = admitShow(createSlideShowWorkspace('mobile', { limit: floor }), { w: 1920, h: 1080 }, 2)
    expect(one.ok && one.backing).toEqual({ w: 1920, h: 1080, dpr: 1 })

    // One byte below DPR 1: refused, nothing charged, a workspace diagnostic.
    const sink = createDiagnosticSink()
    const none = createSlideShowWorkspace('mobile', { limit: floor - 1 })
    expect(admitShow(none, { w: 1920, h: 1080 }, 2, sink)).toEqual({ ok: false })
    expect(none.used()).toBe(0)
    expect(codes(sink)).toEqual(['workspace'])

    // Invalid CSS sizes never reach the ledger.
    for (const css of [{ w: 0, h: 10 }, { w: Number.NaN, h: 10 }, { w: 10, h: -1 }, { w: Infinity, h: 10 }]) {
      expect(admitShow(none, css, 1)).toEqual({ ok: false })
    }
    expect(none.used()).toBe(0)
  })

  it('ladder: reduced, then flattened, then terminal, each diagnosed', async () => {
    const cost: Record<LadderLevel, number> = { full: 1000, reduced: 600, flattened: 300, terminal: 0 }
    const run = async (free: number) => {
      const ws = createSlideShowWorkspace('desktop', { limit: 100 + free, scratch: 50 })
      const terminal = ws.reserveTerminal(50)!
      const sink = createDiagnosticSink()
      const residency = createResidency(ws, fakeStore(), terminal, { sink, cost: (_slide, level) => cost[level] })
      residency.show(2)
      const level = await residency.ensure(2, 'full')
      return { level, ws, sink, residency }
    }
    const full = await run(1000)
    expect(full.level).toBe('full')
    expect(full.sink.list()).toEqual([])
    expect(full.ws.used()).toBe(1100)

    const reduced = await run(700)
    expect(reduced.level).toBe('reduced')
    expect(codes(reduced.sink)).toEqual(['ladder-reduced'])
    expect(reduced.sink.list()[0]!.slideIndex).toBe(2)
    expect(reduced.ws.used()).toBe(700)

    const flattened = await run(400)
    expect(flattened.level).toBe('flattened')
    expect(codes(flattened.sink)).toEqual(['ladder-reduced', 'ladder-flattened'])

    const terminal = await run(100)
    expect(terminal.level).toBe('terminal')
    expect(codes(terminal.sink)).toEqual(['ladder-reduced', 'ladder-flattened', 'ladder-terminal'])
    // The terminal step allocates nothing: it paints into the reserve.
    expect(terminal.ws.used()).toBe(100)
    expect(terminal.residency.owner(2)).toBeUndefined()

    // Re-ensuring a fitted slide at a lower level shrinks its lease in place.
    expect(await full.residency.ensure(2, 'flattened')).toBe('flattened')
    expect(full.ws.used()).toBe(400)
  })

  it('residency keeps current, previous and next; evicts neighbours first', async () => {
    const store = fakeStore()
    const ws = createSlideShowWorkspace('desktop', { limit: 3000, scratch: 0 })
    const terminal = ws.reserveTerminal(0)!
    const sink = createDiagnosticSink()
    const sizes = new Map<number, number>([[4, 500], [5, 500], [6, 500], [7, 500]])
    const residency = createResidency(ws, store, terminal, { sink, cost: slide => sizes.get(slide)! })
    residency.show(5)
    for (const i of [4, 5, 6]) expect(await residency.ensure(i, 'full')).toBe('full')
    expect(residency.resident()).toEqual([4, 5, 6])
    expect(ws.used()).toBe(1500)
    // A slide outside the window is never made resident.
    await expect(residency.ensure(9, 'full')).rejects.toThrow('cancelled')
    expect(store.live.get(9) ?? 0).toBe(0)

    store.log.length = 0
    residency.show(6)
    expect(store.log).toEqual(['release 4', 'evict 4'])
    expect(residency.resident()).toEqual([5, 6])
    expect(ws.used()).toBe(1000)
    expect(await residency.ensure(7, 'full')).toBe('full')
    expect(residency.resident()).toEqual([5, 6, 7])

    // The current slide needs more than is free: neighbours go before it steps down.
    store.log.length = 0
    sizes.set(6, 2500)
    expect(await residency.ensure(6, 'full')).toBe('full')
    expect(sink.list()).toEqual([])
    expect(store.log).toEqual(['release 5', 'evict 5', 'release 7', 'evict 7'])
    expect(residency.resident()).toEqual([6])
    expect(ws.used()).toBe(2500)
    expect(store.live.get(6)).toBe(1)

    // Evicting the current slide frees its layers but keeps its metadata.
    residency.evict(6)
    expect(ws.used()).toBe(0)
    expect(store.live.get(6)).toBe(1)
    residency.dispose()
    expect(store.live.get(6)).toBe(0)
  })

  it('rapid navigation: an allocation completing after its slide was left is released and discarded', async () => {
    // (a) The metadata read completes after the slide was left.
    const store = fakeStore({ deferred: true })
    const ws = createSlideShowWorkspace('desktop', { limit: 10_000, scratch: 0 })
    const terminal = ws.reserveTerminal(0)!
    const residency = createResidency(ws, store, terminal, { cost: () => 1000 })
    residency.show(3)
    const late = residency.ensure(4, 'full')
    residency.show(8)
    store.settle(4)
    await expect(late).rejects.toThrow('cancelled')
    expect(store.live.get(4)).toBe(0)
    expect(ws.used()).toBe(0)
    expect(residency.resident()).toEqual([])

    // (b) The layer build completes after the slide was left: its reservation
    // stays charged while the build still allocates, then is released.
    const store2 = fakeStore()
    const ws2 = createSlideShowWorkspace('desktop', { limit: 10_000, scratch: 0 })
    const gate = deferred()
    const builds: Array<{ slide: number; owner: AllocationOwner }> = []
    const residency2 = createResidency(ws2, store2, ws2.reserveTerminal(0)!, {
      cost: () => 1000,
      build: async (slide, _level, owner) => { builds.push({ slide, owner }); owner.lease(400, `layer:${slide}`); await gate.promise },
    })
    residency2.show(3)
    const building = residency2.ensure(4, 'full')
    await flush()
    expect(builds.map(b => b.slide)).toEqual([4])
    expect(ws2.used()).toBe(1000)
    residency2.show(9)
    expect(ws2.used()).toBe(1000) // still allocating: still charged
    gate.resolve()
    await expect(building).rejects.toThrow('cancelled')
    expect(ws2.used()).toBe(0)
    expect(store2.live.get(4)).toBe(0)
    // The stale build's owner refuses anything further.
    expect(builds[0]!.owner.lease(1, 'late')).toBeUndefined()

    // (c) A failed metadata read does not poison the slide: the next ensure reads afresh.
    const store3 = fakeStore({ failOnce: 5 })
    const ws3 = createSlideShowWorkspace('desktop', { limit: 10_000, scratch: 0 })
    const residency3 = createResidency(ws3, store3, ws3.reserveTerminal(0)!, { cost: () => 1000 })
    residency3.show(5)
    await expect(residency3.ensure(5, 'full')).rejects.toThrow('read failed')
    expect(residency3.resident()).toEqual([])
    expect(ws3.used()).toBe(0)
    expect(await residency3.ensure(5, 'full')).toBe('full')
    expect(residency3.resident()).toEqual([5])
  })

  it('resize that cannot get a larger reserve keeps the old backing store with a diagnostic', () => {
    const setup = (limit: number) => {
      const ws = createSlideShowWorkspace('desktop', { limit, scratch: 1000 })
      const admitted = admitShow(ws, { w: 100, h: 100 }, 1)
      if (!admitted.ok) throw new Error('not admitted')
      const sink = createDiagnosticSink()
      const residency = createResidency(ws, fakeStore(), admitted.terminal, { sink, backing: admitted.backing })
      return { ws, sink, residency, terminal: admitted.terminal }
    }
    // 100x100 -> 110x110: 40000 -> 48400 B. The delta (8400) fits, but the new
    // reserve is taken before the old is released, so old + new must fit.
    const short = setup(41_000 + 48_400 - 1)
    expect(short.residency.resize({ w: 110, h: 110 }, 1)).toBe('kept-old')
    expect(codes(short.sink)).toEqual(['resize-kept-old'])
    expect(short.terminal.outputBytes).toBe(40_000)
    expect(short.residency.backing).toEqual({ w: 100, h: 100, dpr: 1 })
    expect(short.ws.used()).toBe(41_000)

    const room = setup(41_000 + 48_400)
    expect(room.residency.resize({ w: 110, h: 110 }, 1)).toBe('resized')
    expect(room.terminal.outputBytes).toBe(48_400)
    expect(room.residency.backing).toEqual({ w: 110, h: 110, dpr: 1 })
    expect(room.ws.used()).toBe(49_400)
    expect(room.sink.list()).toEqual([])
    // A smaller output shrinks the reserve.
    expect(room.residency.resize({ w: 50, h: 50 }, 1)).toBe('resized')
    expect(room.terminal.outputBytes).toBe(10_000)
    expect(room.ws.used()).toBe(11_000)
  })

  it('pressure stops cross-slide media, then drops neighbour metadata, never current', async () => {
    const log: string[] = []
    const store = fakeStore({ log })
    const ws = createSlideShowWorkspace('desktop', { limit: 10_000, scratch: 0 })
    const residency = createResidency(ws, store, ws.reserveTerminal(0)!, { cost: () => 100, stopCrossSlideMedia: () => { log.push('stop media') } })
    residency.show(5)
    for (const i of [4, 5, 6]) await residency.ensure(i, 'full')
    log.length = 0
    residency.onPressure()
    expect(log).toEqual(['stop media', 'release 4', 'evict 4', 'release 6', 'evict 6'])
    expect(store.live.get(5)).toBe(1)
    expect(store.resident()).toEqual([5])
    expect(residency.resident()).toEqual([5])
    expect(ws.used()).toBe(100)

    // Inside ensure the relief is staged: media first, and if that suffices the neighbours stay.
    const log2: string[] = []
    const store2 = fakeStore({ log: log2 })
    const ws2 = createSlideShowWorkspace('desktop', { limit: 1000, scratch: 0 })
    const media = ws2.lease(400, 'media:loop')!
    const residency2 = createResidency(ws2, store2, ws2.reserveTerminal(0)!, {
      cost: slide => slide === 5 ? 600 : 100,
      stopCrossSlideMedia: () => { log2.push('stop media'); media.release() },
    })
    residency2.show(5)
    await residency2.ensure(4, 'full')
    log2.length = 0
    expect(await residency2.ensure(5, 'full')).toBe('full')
    expect(log2).toEqual(['acquire 5', 'stop media'])
    expect(residency2.resident()).toEqual([4, 5])

    // A neighbour prefetch that does not fit steps down silently, without relief; only a shown slide is diagnosed.
    const sink3 = createDiagnosticSink()
    const ws3 = createSlideShowWorkspace('desktop', { limit: 150, scratch: 0 })
    const residency3 = createResidency(ws3, fakeStore(), ws3.reserveTerminal(0)!, { sink: sink3, cost: () => 100 })
    residency3.show(5)
    await residency3.ensure(5, 'full')
    expect(await residency3.ensure(6, 'full')).toBe('terminal')
    expect(sink3.list()).toEqual([])
    expect(residency3.resident()).toEqual([5, 6])
  })

  it('129 layers or 9 re-rasters in a frame coarsen with a diagnostic', () => {
    expect(LAYER_LIMIT).toBe(128)
    expect(RERASTER_LIMIT).toBe(8)
    const sink = countingSink()
    const limits = createFrameLimits(sink)
    expect(limits.check(1, { layers: 128, reRasters: 8 })).toEqual({ mergeLayers: false, coarsenColor: false })
    expect(sink.all).toEqual([])
    expect(limits.check(1, { layers: 129, reRasters: 8 })).toEqual({ mergeLayers: true, coarsenColor: false })
    expect(limits.check(1, { layers: 128, reRasters: 9 })).toEqual({ mergeLayers: false, coarsenColor: true })
    // Every later frame still coarsens, but the diagnostic is emitted once per slide.
    expect(limits.check(1, { layers: 200, reRasters: 20 })).toEqual({ mergeLayers: true, coarsenColor: true })
    expect(sink.all.map(d => [d.code, d.slideIndex])).toEqual([['layer-limit', 1], ['reraster-limit', 1]])
    limits.check(2, { layers: 129, reRasters: 0 })
    expect(sink.all.map(d => [d.code, d.slideIndex])).toEqual([['layer-limit', 1], ['reraster-limit', 1], ['layer-limit', 2]])
  })
})
