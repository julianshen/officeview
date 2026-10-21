/**
 * S1-10 auto-advance, advClick, kiosk restart and pause (slideshow spec 4.4,
 * working assumption pending D4 captures). Every clock is injected.
 */
import { describe, expect, it } from 'vitest'
import { ManualClock } from '../src/slideshow/clock'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createAutoAdvance, createKioskTimer } from '../src/slideshow/engine/advance'
import { createSlideTimeline } from '../src/slideshow/engine/steps'
import type { BehaviourNode, CTn, Condition, ContainerNode, ShowSettings, ShowSlide, TimeNode, TimingTree, TransitionSpec } from '../src/slideshow/types'

type CtnOver = Partial<Omit<CTn, 'id'>>
const ctn = (id: number, over: CtnOver = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over })
const bhv = (id: number, over: CtnOver = {}): BehaviourNode =>
  ({ kind: 'set', ctn: ctn(id, { fill: 'hold', ...over }), target: { spid: String(id) }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' })
const par = (id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode => ({ kind: 'par', ctn: ctn(id, { fill: 'hold', ...over }), children })
const WAIT: Condition = { delayMs: 'indefinite' }
const step = (id: number, effects: TimeNode[]) => par(id, [par(id + 1, [par(id + 2, effects, { nodeType: 'clickEffect', start: [{ delayMs: 0 }] })], { start: [{ delayMs: 0 }] })], { start: [WAIT] })
const mainSeq = (children: TimeNode[]): ContainerNode =>
  ({ kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', nodeType: 'mainSeq' }), children, seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' } })
const tree = (...children: TimeNode[]): TimingTree => ({ root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot', fill: undefined }), build: [] })
const interactive = (id: number, spid: string, children: TimeNode[]): ContainerNode => ({
  kind: 'seq',
  ctn: ctn(id, { nodeType: 'interactiveSeq', restart: 'whenNotActive', start: [{ delayMs: 0, evt: 'onClick', target: { spid } }] }),
  children: [par(id + 1, children, { start: [{ delayMs: 0 }] })],
  seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})

const transition = (over: Partial<TransitionSpec> = {}): TransitionSpec => ({ element: 'fade', ns: 'p', attrs: {}, durationMs: 500, advClick: true, ...over })
const slide = (timing: TimingTree | undefined, t?: TransitionSpec): ShowSlide => ({ index: 0, hidden: false, transition: t, timing, notes: [], links: [], media: [], builds: [] })
const settings = (over: Partial<ShowSettings> = {}): ShowSettings => ({ loop: false, showAnimation: true, useTimings: true, showNarration: false, type: 'present', kioskRestartMs: 300000, ...over })
const NORMAL = { showAnimation: true, reducedMotion: false }

describe('auto-advance', () => {
  it('advTm missing never advances; advTm=0 advances at transition end; pending click steps defer; indefinite repeat does not hold', () => {
    const none = createSlideTimeline(undefined, NORMAL)
    // advTm missing: never.
    expect(createAutoAdvance(slide(undefined, transition()), settings(), none).due(none.initial(), 500, 1e9)).toBe(false)
    expect(createAutoAdvance(slide(undefined), settings(), none).due(none.initial(), 0, 1e9)).toBe(false)
    // advTm=0 on a static slide: due exactly when the transition ends (500).
    const zero = createAutoAdvance(slide(undefined, transition({ advTmMs: 0 })), settings(), none)
    expect(zero.due(none.initial(), 500, 499)).toBe(false)
    expect(zero.due(none.initial(), 500, 500)).toBe(true)
    // advTm=2000: the timer starts at transition end.
    const two = createAutoAdvance(slide(undefined, transition({ advTmMs: 2000 })), settings(), none)
    expect(two.due(none.initial(), 500, 2499)).toBe(false)
    expect(two.due(none.initial(), 500, 2500)).toBe(true)
    // useTimings off: never.
    expect(createAutoAdvance(slide(undefined, transition({ advTmMs: 0 })), settings({ useTimings: false }), none).due(none.initial(), 0, 1e9)).toBe(false)

    // One click step with a 1000 ms fade: the pending click defers; then the running finite effect holds.
    const fadeTree = tree(mainSeq([step(3, [bhv(6, { durMs: 1000 })])]))
    const fade = createSlideTimeline(fadeTree, NORMAL)
    const fa = createAutoAdvance(slide(fadeTree, transition({ advTmMs: 0 })), settings(), fade)
    expect(fa.due(fade.initial(), 500, 1e9)).toBe(false) // unplayed click step
    const f1 = fade.next(fade.initial(), 600) // fade runs [600, 1600)
    expect(fa.due(f1, 500, 1000)).toBe(false)
    expect(fa.due(f1, 500, 1599)).toBe(false)
    expect(fa.due(f1, 500, 1600)).toBe(true)
    // Settling the last step by Next releases the hold at once.
    const f2 = fade.next(f1, 700)
    expect(fa.due(f2, 500, 700)).toBe(true)
    // A checkpoint is settled.
    expect(fa.due(fade.checkpoint(1), 500, 500)).toBe(true)

    // Only an indefinite spin: never holds.
    const spinTree = tree(mainSeq([step(3, [bhv(7, { durMs: 500, repeatCount: 'indefinite', fill: undefined })])]))
    const spin = createSlideTimeline(spinTree, NORMAL)
    const sa = createAutoAdvance(slide(spinTree, transition({ advTmMs: 0 })), settings(), spin)
    const s1 = spin.next(spin.initial(), 600)
    expect(spin.stateAt(s1, 700).active.map(e => e.node.ctn.id)).toEqual([7])
    expect(sa.due(s1, 500, 700)).toBe(true)

    // An interactive sequence never holds; nor does media waiting for its end.
    const trigTree = tree(mainSeq([par(3, [{ kind: 'audio', ctn: ctn(40, { durMs: 'indefinite' }), target: { spid: '40' }, attrNames: [], additive: 'base', accumulate: 'none' }], { start: [{ delayMs: 0 }] })]), interactive(30, '9', [bhv(33, { durMs: 10000 })]))
    const trig = createSlideTimeline(trigTree, NORMAL)
    expect(trig.steps).toBe(0)
    const ta = createAutoAdvance(slide(trigTree, transition({ advTmMs: 0 })), settings(), trig)
    const tc = trig.trigger(trig.initial(), '9', 100)!
    expect(trig.stateAt(tc, 500).active.some(e => e.node.ctn.id === 33)).toBe(true)
    expect(ta.due(tc, 0, 500)).toBe(true)
  })

  it('advClick=0 refuses clicks but accepts keyboard and API next', () => {
    const t = tree(mainSeq([step(3, [bhv(6, { durMs: 100 })]), step(10, [bhv(13, { durMs: 100 })])]))
    const tl = createSlideTimeline(t, NORMAL)
    const refusing = createAutoAdvance(slide(t, transition({ advClick: false })), settings(), tl)
    expect(refusing.clickAdvances).toBe(false)
    // Keyboard and API next are not gated: the timeline still advances.
    expect(tl.next(tl.initial(), 0).step).toBe(1)
    expect(createAutoAdvance(slide(t, transition({ advClick: true })), settings(), tl).clickAdvances).toBe(true)
    expect(createAutoAdvance(slide(t), settings(), tl).clickAdvances).toBe(true)
  })
})

describe('kiosk', () => {
  it('kiosk restarts restart ms after the last start or restart; each loop restart resets the counter; pause stops it', () => {
    const clock = new ManualClock()
    const k = createKioskTimer(settings({ type: 'kiosk', loop: true }), clock.now())
    clock.advance(299_999)
    expect(k.restartDue(clock.now())).toBe(false)
    clock.advance(1)
    expect(k.restartDue(clock.now())).toBe(true) // 300,000 after the start
    k.restarted(clock.now())
    expect(k.restartDue(clock.now())).toBe(false)
    // A loop restart 100,000 later resets the counter too.
    clock.advance(100_000)
    k.restarted(clock.now()) // at 400,000
    clock.advance(299_999)
    expect(k.restartDue(clock.now())).toBe(false)
    // Pause stops the show clock, so the counter stops with it.
    clock.pause()
    clock.advance(1_000_000)
    expect(k.restartDue(clock.now())).toBe(false)
    clock.resume()
    clock.advance(1)
    expect(k.restartDue(clock.now())).toBe(true) // 700,000

    // A custom restart.
    const c = createKioskTimer(settings({ type: 'kiosk', kioskRestartMs: 5000 }), 1000)
    expect(c.restartDue(5999)).toBe(false)
    expect(c.restartDue(6000)).toBe(true)
    // Not a kiosk: never.
    expect(createKioskTimer(settings(), 0).restartDue(1e12)).toBe(false)
    // restart=0 would restart on every frame: refused (never due) and diagnosed.
    const sink = createDiagnosticSink()
    const z = createKioskTimer(settings({ type: 'kiosk', kioskRestartMs: 0 }), 0, sink)
    expect(z.restartDue(1e12)).toBe(false)
    expect(sink.list().map(d => d.code)).toEqual(['kiosk-restart'])
  })
})
