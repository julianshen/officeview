/**
 * The package-free show driver (streaming Stage 3 piece 4 Task 3): the
 * navigation logic the local session and the streamed show share, driven
 * here through a fake slide source and a fake stage that log their calls.
 * These tests are the spec for the stage call order on a slide change:
 * `flattenCurrent`, then `leave`, then `enter`.
 */
import { describe, expect, it } from 'vitest'
import { ManualClock } from '../src/slideshow/clock'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createShowDriver, type DriverHooks, type DriverStage, type HitResult, type SlideSource } from '../src/slideshow/player/driver'
import type { AudioBehaviour, BehaviourNode, CTn, Condition, ContainerNode, ShowSlide, TimeNode, TimingTree, TransitionSpec, VideoBehaviour } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Timing trees
// ---------------------------------------------------------------------------

type CtnOver = Partial<Omit<CTn, 'id'>>
const ctn = (id: number, over: CtnOver = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over })
const WAIT: Condition = { delayMs: 'indefinite' }
const at = (delayMs: number): Condition => ({ delayMs })
const par = (id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode => ({ kind: 'par', ctn: ctn(id, { fill: 'hold', ...over }), children })
const show = (id: number, spid: string, over: CtnOver = {}): BehaviourNode =>
  ({ kind: 'set', ctn: ctn(id, { durMs: 1, fill: 'hold', ...over }), target: { spid }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' })
/** A click step: waits for "next", then shows shape `spid`. */
const clickStep = (id: number, spid: string) => par(id, [par(id + 1, [show(id + 2, spid)], { start: [at(0)] })], { start: [WAIT] })
const mainSeq = (children: TimeNode[]): ContainerNode => ({
  kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', nodeType: 'mainSeq' }), children, seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})
/** An interactive sequence on trigger shape `spid`. */
const interactive = (id: number, spid: string, target: string): ContainerNode => ({
  kind: 'seq',
  ctn: ctn(id, { nodeType: 'interactiveSeq', restart: 'whenNotActive', start: [{ delayMs: 0, evt: 'onClick', target: { spid } }] }),
  children: [par(id + 1, [par(id + 2, [show(id + 3, target)], { nodeType: 'clickEffect', start: [at(0)] })], { start: [at(0)] })],
  seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})
const tree = (...children: TimeNode[]): TimingTree => ({ root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot', fill: undefined }), build: [] })
const media = (kind: 'audio' | 'video', id: number, spid: string): BehaviourNode =>
  ({ kind, ctn: ctn(id, { fill: 'hold', start: [at(0)] }), target: { spid }, attrNames: [], additive: 'base', accumulate: 'none' }) as AudioBehaviour | VideoBehaviour
/** Shape 99 becomes visible once node `tn` ends. */
const successor = (id: number, tn: number): BehaviourNode =>
  ({ kind: 'set', ctn: ctn(id, { durMs: 500, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn }] }), target: { spid: '99' }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' })

const transition = (over: Partial<TransitionSpec> = {}): TransitionSpec => ({ element: 'fade', ns: 'p', attrs: {}, durationMs: 500, advClick: true, ...over })
const slideOf = (index: number, over: Partial<ShowSlide> = {}): ShowSlide =>
  ({ index, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [], ...over })

// ---------------------------------------------------------------------------
// Fakes
// ---------------------------------------------------------------------------

interface Rig {
  log: string[]
  acquired: number[]
  released: number[]
  hooks: string[]
  sink: ReturnType<typeof createDiagnosticSink>
  clock: ManualClock
  driver: ReturnType<typeof createShowDriver>
  /** What the stage's hit test answers. */
  hit: { result: HitResult | undefined }
  /** Holds slide `i`'s acquire until `open(i)`. */
  gate(i: number): void
  open(i: number): void
}

interface RigOptions {
  slides: ShowSlide[]
  hidden?: number[]
  settings?: Partial<SlideSource['settings'] & { kioskRestartMs: number }>
  kiosk?: boolean
  hooks?: Partial<DriverHooks>
}

function rig(o: RigOptions): Rig {
  const log: string[] = []
  const acquired: number[] = []
  const released: number[] = []
  const hookLog: string[] = []
  const sink = createDiagnosticSink()
  const clock = new ManualClock()
  const gates = new Map<number, () => void>()
  const held = new Set<number>()
  const hit: Rig['hit'] = { result: undefined }
  let origin = Number.POSITIVE_INFINITY
  const settings = { showAnimation: true, useTimings: true, loop: false, type: 'present' as const, kioskRestartMs: 0, ...o.settings }
  const source: SlideSource = {
    count: o.slides.length,
    settings,
    hidden: (i) => o.hidden?.includes(i) ?? false,
    async acquire(i) {
      acquired.push(i)
      if (held.has(i)) await new Promise<void>((resolve) => gates.set(i, resolve))
      return { slide: o.slides[i]!, charge: () => true, release: () => { released.push(i) } }
    },
  }
  const stage: DriverStage = {
    prepare(i) { log.push(`prepare:${i}`) },
    async enter(model, t, neighbours) {
      log.push(`enter:${model.index}:${model.cursor.step}:[${neighbours.join(',')}]`)
      await Promise.resolve()
      origin = t + (model.slide.transition?.durationMs ?? 0)
    },
    frame(cursor) { log.push(`frame:${cursor ? cursor.step : '-'}`); return new Map() },
    flattenCurrent() { log.push('flattenCurrent') },
    leave() { log.push('leave'); origin = Number.POSITIVE_INFINITY },
    hitTest() { log.push('hitTest'); return hit.result },
    slideTime: (t) => t - origin,
  }
  const hooks: DriverHooks = {
    onSlideChange: (i, b) => { hookLog.push(`slide:${i}:${b}`) },
    onVisit: (v) => { hookLog.push(v ? `visit:${v.index}` : 'visit:-') },
    onCursor: (v) => { hookLog.push(`cursor:${v.index}:${v.cursor?.step}`) },
    runLink: (l) => { hookLog.push(`link:${l.spid}`) },
    onEnd: () => { hookLog.push('end') },
    ...o.hooks,
  }
  const driver = createShowDriver({ source, stage, clock, sink, reducedMotion: false, kiosk: o.kiosk ?? false, hooks })
  return {
    log, acquired, released, hooks: hookLog, sink, clock, driver, hit,
    gate: (i) => { held.add(i) },
    open: (i) => { held.delete(i); gates.get(i)?.() },
  }
}

const flush = async () => { for (let i = 0; i < 10; i++) await Promise.resolve() }
const steps = (n: number, from = 10) => tree(mainSeq(Array.from({ length: n }, (_, k) => clickStep(from + k * 10, String(100 + k)))))

// ---------------------------------------------------------------------------
// Navigation
// ---------------------------------------------------------------------------

describe('show driver', () => {
  it('next past the last step leaves forward: flattenCurrent then leave then enter(i+1)', async () => {
    const r = rig({ slides: [slideOf(0, { timing: steps(1) }), slideOf(1)] })
    await r.driver.start(0)
    expect(r.driver.slide).toBe(0)
    r.log.length = 0
    r.driver.next()
    // The one click step plays: no slide change.
    expect(r.driver.visit?.cursor?.step).toBe(1)
    expect(r.log).toEqual([])
    // Once its effect has run, the next Next leaves the slide.
    r.clock.advance(10)
    r.driver.next()
    expect(r.driver.slide).toBe(1)
    await flush()
    expect(r.log.slice(0, 3)).toEqual(['flattenCurrent', 'leave', 'enter:1:0:[0]'])
    // The visible neighbours are prepared once the slide is entered.
    expect(r.log).toContain('prepare:0')
    expect(r.released).toEqual([0])
  })

  it('previous at step 0 enters the previous visible slide at its last checkpoint', async () => {
    const r = rig({ slides: [slideOf(0, { timing: steps(2) }), slideOf(1)] })
    await r.driver.start(1)
    r.log.length = 0
    r.hooks.length = 0
    r.driver.previous()
    await flush()
    expect(r.log.slice(0, 3)).toEqual(['flattenCurrent', 'leave', 'enter:0:2:[1]'])
    expect(r.driver.slide).toBe(0)
    expect(r.driver.visit?.cursor?.step).toBe(2)
    expect(r.hooks).toContain('slide:0:2')
    // Previous within the slide moves to the checkpoint before.
    r.driver.previous()
    expect(r.driver.visit?.cursor?.step).toBe(1)
    expect(r.hooks).toContain('slide:0:1')
  })

  it('goTo resets the cursor and supersedes queued inputs', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1, { timing: steps(2) }), slideOf(2, { timing: steps(2) })] })
    await r.driver.start(0)
    r.gate(1)
    r.driver.next()
    await flush()
    // Slide 1 is loading: step inputs queue on it.
    r.driver.next()
    r.driver.next()
    expect(r.driver.visit?.cursor).toBeUndefined()
    r.driver.goTo(2)
    await flush()
    r.open(1)
    await flush()
    expect(r.driver.slide).toBe(2)
    expect(r.driver.visit?.cursor?.step).toBe(0)
    // The stale load is released unused; the queued steps never ran on slide 2.
    expect(r.released).toContain(1)
    expect(r.log.filter((l) => l.startsWith('enter:'))).toEqual(['enter:0:0:[1]', 'enter:2:0:[1]'])
    // An index outside the deck is diagnosed and ignored.
    r.driver.goTo(7)
    expect(r.driver.slide).toBe(2)
    expect(r.sink.list().map((d) => d.code)).toContain('invalid-slide')
  })

  it('hidden slides are skipped and never acquired', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1), slideOf(2)], hidden: [1] })
    await r.driver.start(0)
    expect(r.log).toContain('enter:0:0:[2]')
    r.driver.next()
    await flush()
    expect(r.driver.slide).toBe(2)
    r.driver.previous()
    await flush()
    expect(r.driver.slide).toBe(0)
    expect(r.acquired).not.toContain(1)
    expect(r.log).not.toContain('prepare:1')
  })

  it('after-timer auto-advance fires at advTmMs after the transition', async () => {
    const r = rig({ slides: [slideOf(0, { transition: transition({ durationMs: 500, advTmMs: 2000 }) }), slideOf(1)] })
    await r.driver.start(0)
    r.clock.advance(2499)
    r.driver.tick()
    expect(r.driver.slide).toBe(0)
    r.clock.advance(1)
    r.driver.tick()
    expect(r.driver.slide).toBe(1)
    await flush()
    expect(r.log).toContain('enter:1:0:[0]')
  })

  it('a blanked screen holds the after-timer auto-advance', async () => {
    let blank = true
    const r = rig({ slides: [slideOf(0, { transition: transition({ durationMs: 0, advTmMs: 100 }) }), slideOf(1)], hooks: { held: () => blank } })
    await r.driver.start(0)
    r.clock.advance(200)
    r.driver.tick()
    expect(r.driver.slide).toBe(0)
    blank = false
    r.driver.tick()
    expect(r.driver.slide).toBe(1)
  })

  // -------------------------------------------------------------------------
  // Clicks
  // -------------------------------------------------------------------------

  it('a trigger hit runs timeline.trigger and does not advance', async () => {
    const r = rig({ slides: [slideOf(0, { timing: tree(interactive(50, '5', '6')) }), slideOf(1)] })
    await r.driver.start(0)
    r.hit.result = { kind: 'trigger', spid: '5' }
    r.log.length = 0
    r.driver.click(10, 10)
    expect(r.log).toEqual(['hitTest'])
    expect(r.driver.slide).toBe(0)
    expect(r.driver.visit?.cursor?.events.some((e) => e.kind === 'trigger')).toBe(true)
    expect(r.hooks).toContain('cursor:0:0')
  })

  it('a link hit runs the link and does not advance', async () => {
    const link = { spid: '7', action: { kind: 'slide' as const, index: 1 } }
    const r = rig({ slides: [slideOf(0, { links: [link] }), slideOf(1)] })
    await r.driver.start(0)
    r.hit.result = { kind: 'link', link }
    r.driver.click(10, 10)
    expect(r.hooks).toContain('link:7')
    expect(r.driver.slide).toBe(0)
  })

  it('a withheld hit consumes the click', async () => {
    const link = { spid: '7', action: { kind: 'slide' as const, index: 1 } }
    const r = rig({ slides: [slideOf(0, { links: [link] }), slideOf(1)] })
    await r.driver.start(0)
    r.hit.result = { kind: 'withheld' }
    r.log.length = 0
    r.driver.click(10, 10)
    expect(r.driver.slide).toBe(0)
    expect(r.log).not.toContain('leave')
    expect(r.sink.list().map((d) => d.code)).toContain('link-withheld')
  })

  it('kiosk consumes trigger and link hits without running them', async () => {
    const link = { spid: '7', action: { kind: 'slide' as const, index: 1 } }
    const r = rig({ slides: [slideOf(0, { timing: tree(interactive(50, '5', '6')), links: [link] }), slideOf(1)], kiosk: true, settings: { type: 'kiosk', kioskRestartMs: 60000 } })
    await r.driver.start(0)
    r.hit.result = { kind: 'trigger', spid: '5' }
    r.driver.click(10, 10)
    r.hit.result = { kind: 'link', link }
    r.driver.click(10, 10)
    expect(r.driver.slide).toBe(0)
    expect(r.driver.visit?.cursor?.events.some((e) => e.kind === 'trigger')).toBe(false)
    expect(r.hooks).not.toContain('link:7')
  })

  it('advClick=0 blocks click advance, not next()', async () => {
    const r = rig({ slides: [slideOf(0, { transition: transition({ advClick: false }) }), slideOf(1)] })
    await r.driver.start(0)
    r.driver.click(10, 10)
    expect(r.driver.slide).toBe(0)
    r.driver.next()
    expect(r.driver.slide).toBe(1)
  })

  it('a click with no trigger or link under it advances', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1)] })
    await r.driver.start(0)
    r.log.length = 0
    r.driver.click(10, 10)
    // Nothing interactive on the slide: the stage is not asked.
    expect(r.log[0]).toBe('flattenCurrent')
    expect(r.driver.slide).toBe(1)
  })

  // -------------------------------------------------------------------------
  // End, loop and kiosk
  // -------------------------------------------------------------------------

  it('next past the last slide reaches the end, and next there ends the show', async () => {
    const r = rig({ slides: [slideOf(0)] })
    await r.driver.start(0)
    r.driver.next()
    expect(r.driver.slide).toBe('end')
    expect(r.driver.visit).toBeUndefined()
    r.driver.next()
    expect(r.hooks).toContain('end')
    // Previous from the end enters the last visible slide at its last checkpoint.
    r.driver.previous()
    await flush()
    expect(r.driver.slide).toBe(0)
  })

  it('loop wraps to the first visible slide', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1)], settings: { loop: true } })
    await r.driver.start(1)
    r.driver.next()
    expect(r.driver.slide).toBe(0)
  })

  it('kiosk restart re-enters order[0]', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1), slideOf(2)], hidden: [0], kiosk: true, settings: { type: 'kiosk', kioskRestartMs: 1000 } })
    await r.driver.start(1)
    r.driver.next()
    await flush()
    expect(r.driver.slide).toBe(2)
    r.clock.advance(999)
    r.driver.tick()
    expect(r.driver.slide).toBe(2)
    r.clock.advance(1)
    r.log.length = 0
    expect(r.driver.tick()).toBeUndefined()
    expect(r.driver.slide).toBe(1)
    await flush()
    expect(r.log.slice(0, 3)).toEqual(['flattenCurrent', 'leave', 'enter:1:0:[2]'])
  })

  // -------------------------------------------------------------------------
  // Media (D16)
  // -------------------------------------------------------------------------

  it('media nodes end synthetically without a player', async () => {
    const r = rig({ slides: [slideOf(0, { timing: tree(media('video', 3, '4'), successor(5, 3)) })] })
    await r.driver.start(0)
    r.clock.advance(100)
    r.driver.tick()
    const ends = r.driver.visit!.cursor!.events.filter((e) => e.kind === 'mediaEnd')
    expect(ends).toEqual([{ kind: 'mediaEnd', nodeId: 3, t: 100, synthetic: true }])
    expect(r.hooks).toContain('cursor:0:0')
    // Once per activation.
    r.clock.advance(100)
    r.driver.tick()
    expect(r.driver.visit!.cursor!.events.filter((e) => e.kind === 'mediaEnd')).toHaveLength(1)
  })

  it('a media node restarted within the slide ends synthetically again, and its successor runs each time', async () => {
    // An interactive sequence on trigger 5 plays audio 53; shape 99 shows (for 500 ms) once it ends.
    const seq: ContainerNode = {
      kind: 'seq',
      ctn: ctn(50, { nodeType: 'interactiveSeq', restart: 'whenNotActive', start: [{ delayMs: 0, evt: 'onClick', target: { spid: '5' } }] }),
      children: [par(51, [par(52, [media('audio', 53, '4'), successor(54, 53)], { nodeType: 'clickEffect', start: [at(0)] })], { start: [at(0)] })],
      seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
    }
    const r = rig({ slides: [slideOf(0, { timing: tree(seq) })] })
    await r.driver.start(0)
    const successorRuns = () => r.driver.visit!.timeline!.stateAt(r.driver.visit!.cursor!, Math.max(0, r.clock.now())).active.some((e) => e.node.ctn.id === 54)
    for (const round of [1, 2]) {
      r.driver.trigger('5')
      r.clock.advance(10)
      r.driver.tick()
      r.driver.tick()
      expect(r.driver.visit!.cursor!.events.filter((e) => e.kind === 'mediaEnd' && e.synthetic)).toHaveLength(round)
      r.clock.advance(10)
      expect(successorRuns()).toBe(true)
      // Let the successor finish, so the sequence is no longer active and can restart.
      r.clock.advance(1000)
      r.driver.tick()
    }
  })

  it('a restart between ticks is a new activation: it ends synthetically again without an inactive frame between', async () => {
    // The re-review's timing: the second trigger lands before any tick samples the node inactive.
    const quick: BehaviourNode = { ...successor(64, 63), ctn: ctn(64, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 63 }] }) }
    const seq: ContainerNode = {
      kind: 'seq',
      ctn: ctn(60, { nodeType: 'interactiveSeq', restart: 'whenNotActive', start: [{ delayMs: 0, evt: 'onClick', target: { spid: '7' } }] }),
      children: [par(61, [par(62, [media('video', 63, '4'), quick], { nodeType: 'clickEffect', start: [at(0)] })], { start: [at(0)] })],
      seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
    }
    const r = rig({ slides: [slideOf(0, { timing: tree(seq) })] })
    await r.driver.start(0)
    const ends = () => r.driver.visit!.cursor!.events.filter((e) => e.kind === 'mediaEnd')
    const snap = (t: number) => r.driver.visit!.timeline!.stateAt(r.driver.visit!.cursor!, t)
    r.clock.advance(100)
    r.driver.trigger('7')
    r.driver.tick()
    expect(ends()).toEqual([{ kind: 'mediaEnd', nodeId: 63, t: 100, synthetic: true }])
    expect(snap(105).active.some((e) => e.node.ctn.id === 64)).toBe(true)
    r.clock.advance(12)
    expect(snap(112).ended.has(64)).toBe(true)
    r.driver.trigger('7')
    r.clock.advance(4)
    r.driver.tick()
    expect(ends()).toEqual([{ kind: 'mediaEnd', nodeId: 63, t: 100, synthetic: true }, { kind: 'mediaEnd', nodeId: 63, t: 116, synthetic: true }])
    expect(snap(120).active.some((e) => e.node.ctn.id === 64)).toBe(true)
    // Later ticks of the same activation add nothing.
    r.clock.advance(100)
    r.driver.tick()
    r.driver.tick()
    expect(ends()).toHaveLength(2)
  })

  it('a supplied media player owns media ends: none is synthesised', async () => {
    const synced: number[] = []
    let onEnd: ((id: number, synthetic: boolean, t: number) => void) | undefined
    const r = rig({
      slides: [slideOf(0, { timing: tree(media('video', 3, '4'), successor(5, 3)) })],
      hooks: { media: () => ({ wants: true, sync: (_s, t) => { synced.push(t) }, onEnd: (cb) => { onEnd = cb }, leave: () => {} }) },
    })
    await r.driver.start(0)
    r.clock.advance(100)
    r.driver.tick()
    expect(synced).toEqual([100])
    expect(r.driver.visit!.cursor!.events.some((e) => e.kind === 'mediaEnd')).toBe(false)
    onEnd!(3, false, 150)
    expect(r.driver.visit!.cursor!.events.filter((e) => e.kind === 'mediaEnd')).toEqual([{ kind: 'mediaEnd', nodeId: 3, t: 150, synthetic: false }])
  })

  it('dispose releases the shown slide and stops navigation', async () => {
    const r = rig({ slides: [slideOf(0), slideOf(1)] })
    await r.driver.start(0)
    const cursor = r.driver.visit?.cursor
    r.driver.dispose()
    expect(r.released).toEqual([0])
    // The last visit stays readable after dispose (a host reads buildIndex in onEnd), and is never released twice.
    expect(r.driver.visit?.cursor).toBe(cursor)
    expect(cursor).toBeDefined()
    r.driver.dispose()
    expect(r.released).toEqual([0])
    r.driver.next()
    expect(r.driver.slide).toBe(0)
    expect(r.driver.tick()).toBeUndefined()
  })
})
