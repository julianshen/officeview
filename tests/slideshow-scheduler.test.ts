/**
 * S1-9 timing scheduler core (slideshow spec 6.2). Every tree is built here
 * as a `TimingTree` value; time is plain numbers (milliseconds of show
 * clock), never wall time. Expected values are hand-computed in comments.
 */
import { describe, expect, it } from 'vitest'
import { createScheduler, type SchedulerSnapshot } from '../src/slideshow/engine/scheduler'
import { easeProgress } from '../src/slideshow/engine/ease'
import { ManualClock } from '../src/slideshow/clock'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import type { BehaviourNode, CTn, Condition, ContainerNode, ShowEvent, TimeNode, TimingTree } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Builders
// ---------------------------------------------------------------------------

type CtnOver = Partial<Omit<CTn, 'id'>>

function ctn(id: number, over: CtnOver = {}): CTn {
  return { id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over }
}

function bhv(id: number, over: CtnOver = {}): BehaviourNode {
  return { kind: 'set', ctn: ctn(id, over), target: { spid: String(id) }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' }
}

function par(id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode {
  return { kind: 'par', ctn: ctn(id, over), children }
}

function seq(id: number, children: TimeNode[], over: CtnOver = {}, opts?: ContainerNode['seq']): ContainerNode {
  return { kind: 'seq', ctn: ctn(id, over), children, seq: opts }
}

function excl(id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode {
  return { kind: 'excl', ctn: ctn(id, over), children }
}

/** The usual `tmRoot`: an indefinite par that never restarts. */
function tree(...children: TimeNode[]): TimingTree {
  return { root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot' }), build: [] }
}

const delay = (delayMs: Condition['delayMs']): Condition => ({ delayMs })
const WAIT: Condition = { delayMs: 'indefinite' }

type Seen = { state: 'active' | 'frozen'; progress: number; iteration: number } | { state: 'none' }

function find(snap: SchedulerSnapshot, id: number): Seen {
  const a = snap.active.filter((e) => e.node.ctn.id === id)
  const f = snap.frozen.filter((e) => e.node.ctn.id === id)
  expect(a.length + f.length).toBeLessThanOrEqual(1)
  if (a[0]) return { state: 'active', progress: a[0].progress, iteration: a[0].iteration }
  if (f[0]) return { state: 'frozen', progress: f[0].progress, iteration: f[0].iteration }
  return { state: 'none' }
}

function expectSeen(snap: SchedulerSnapshot, id: number, state: 'active' | 'frozen', progress: number, iteration = 0): void {
  const s = find(snap, id)
  expect(s.state, `node ${id} state`).toBe(state)
  if (s.state === 'none') return
  expect(s.progress, `node ${id} progress`).toBeCloseTo(progress, 9)
  expect(s.iteration, `node ${id} iteration`).toBe(iteration)
}

// ---------------------------------------------------------------------------

describe('scheduler', () => {
  it('par children start together; seq children chain; excl pauses siblings', () => {
    // par: both children begin at 0.
    const p = createScheduler(tree(par(2, [bhv(10, { durMs: 500, fill: 'hold' }), bhv(11, { durMs: 1000, fill: 'hold' })], { fill: 'hold' })))
    let s = p.evaluate([], 250)
    expectSeen(s, 10, 'active', 0.5) // 250/500
    expectSeen(s, 11, 'active', 0.25) // 250/1000
    expect(p.endTimeOf(2, [])).toBe(1000) // implicit par duration = last child end

    // seq: A [0,500); B begins 100 after A ends, [600,900).
    const q = createScheduler(tree(seq(3, [bhv(20, { durMs: 500, fill: 'hold' }), bhv(21, { durMs: 300, fill: 'hold', start: [delay(100)] })], { fill: 'hold' })))
    s = q.evaluate([], 550)
    expectSeen(s, 20, 'frozen', 1)
    expect(find(s, 21).state).toBe('none')
    s = q.evaluate([], 750)
    expectSeen(s, 21, 'active', 0.5) // (750-600)/300
    expect(q.endTimeOf(21, [])).toBe(900)
    expect(q.endTimeOf(3, [])).toBe(900)

    // excl: A [0,1000) is paused while B [300,800) plays, then resumes; A ends at 1500.
    const x = createScheduler(tree(excl(4, [bhv(30, { durMs: 1000, fill: 'hold' }), bhv(31, { durMs: 500, fill: 'hold', start: [delay(300)] })], { fill: 'hold' })))
    s = x.evaluate([], 400)
    expectSeen(s, 30, 'active', 0.3) // paused at elapsed 300
    expectSeen(s, 31, 'active', 0.2) // 100/500
    s = x.evaluate([], 900)
    expectSeen(s, 31, 'frozen', 1)
    expectSeen(s, 30, 'active', 0.4) // 300 + (900-800)
    expect(x.endTimeOf(30, [])).toBe(1500)
    expect(x.endTimeOf(4, [])).toBe(1500)
    expectSeen(x.evaluate([], 1500), 30, 'frozen', 1)
  })

  it('delay, dur and indefinite; endSync and end conditions cut a node short', () => {
    const sch = createScheduler(tree(
      bhv(10, { start: [delay(200)], durMs: 500, fill: 'hold' }),
      bhv(11, { durMs: 'indefinite' }),
      bhv(12, { durMs: 1000, end: [delay(400)], fill: 'hold' }),
      // endSync first: the par ends when its first child ends (300); E is cut at 300.
      par(5, [bhv(13, { durMs: 300, fill: 'hold' }), bhv(14, { durMs: 900, fill: 'hold' })], { durMs: 'indefinite', fill: 'hold', endSync: { delayMs: 0, evt: 'end', rtn: 'first' } }),
      // endSync on a named child: ends when node 16 ends (200).
      par(6, [bhv(15, { durMs: 800, fill: 'hold' }), bhv(16, { durMs: 200, fill: 'hold' })], { durMs: 'indefinite', fill: 'hold', endSync: { delayMs: 0, evt: 'end', tn: 16 } }),
      // endSync all waits for a child that never begins.
      par(7, [bhv(17, { durMs: 100 }), bhv(18, { durMs: 100, start: [WAIT] })], { endSync: { delayMs: 0, evt: 'end', rtn: 'all' } }),
    ))
    expect(find(sch.evaluate([], 100), 10).state).toBe('none')
    expectSeen(sch.evaluate([], 200), 10, 'active', 0) // active on [begin, end)
    expectSeen(sch.evaluate([], 450), 10, 'active', 0.5)
    let s = sch.evaluate([], 700)
    expectSeen(s, 10, 'frozen', 1) // at t === end: ended, frozen
    expect(s.ended.has(10)).toBe(true)
    expect(sch.endTimeOf(10, [])).toBe(700)

    s = sch.evaluate([], 1_000_000)
    expectSeen(s, 11, 'active', 0) // indefinite simple duration: progress stays 0
    expect(sch.endTimeOf(11, [])).toBe('indefinite')

    expectSeen(sch.evaluate([], 300), 12, 'active', 0.3)
    expectSeen(sch.evaluate([], 400), 12, 'frozen', 0.4) // end condition at 400 cuts it
    expect(sch.endTimeOf(12, [])).toBe(400)

    s = sch.evaluate([], 600)
    expectSeen(s, 13, 'frozen', 1)
    expectSeen(s, 14, 'frozen', 300 / 900) // frozen at the progress it reached
    expect(s.ended.has(5)).toBe(true)
    expect(sch.endTimeOf(5, [])).toBe(300)
    expect(sch.endTimeOf(14, [])).toBe(300)

    s = sch.evaluate([], 250)
    expectSeen(s, 15, 'frozen', 0.25) // 200/800
    expect(sch.endTimeOf(6, [])).toBe(200)

    expect(sch.endTimeOf(7, [])).toBe('indefinite')
    expect(sch.endTimeOf(18, [])).toBeUndefined() // never begins
  })

  it('repeat 2.5 iterations ends mid-iteration; repeatDur caps; indefinite never ends', () => {
    const sch = createScheduler(tree(
      bhv(10, { durMs: 1000, repeatCount: 2.5, fill: 'hold' }),
      bhv(11, { durMs: 1000, repeatCount: 'indefinite', repeatDurMs: 1800, fill: 'hold' }),
      bhv(12, { durMs: 1000, repeatCount: 'indefinite' }),
      // A repeating container restarts its children each iteration.
      par(5, [bhv(13, { durMs: 400, fill: 'hold' })], { repeatCount: 2, fill: 'hold' }),
    ))
    expectSeen(sch.evaluate([], 1200), 10, 'active', 0.2, 1)
    expectSeen(sch.evaluate([], 2500), 10, 'frozen', 0.5, 2)
    expect(sch.endTimeOf(10, [])).toBe(2500)

    expectSeen(sch.evaluate([], 1800), 11, 'frozen', 0.8, 1)
    expect(sch.endTimeOf(11, [])).toBe(1800)

    expectSeen(sch.evaluate([], 10_000_250), 12, 'active', 0.25, 10_000)
    expect(sch.endTimeOf(12, [])).toBe('indefinite')

    expectSeen(sch.evaluate([], 300), 13, 'active', 0.75)
    expectSeen(sch.evaluate([], 600), 13, 'active', 0.5) // parent iteration 1, child local 200/400
    expectSeen(sch.evaluate([], 800), 13, 'frozen', 1)
    expect(sch.endTimeOf(5, [])).toBe(800)
    expect(sch.endTimeOf(13, [])).toBe(800)
  })

  it('autoRev doubles the active duration and mirrors progress', () => {
    const sch = createScheduler(tree(
      bhv(10, { durMs: 1000, autoRev: true, fill: 'hold' }),
      bhv(11, { durMs: 1000, autoRev: true, repeatCount: 2, fill: 'hold' }),
    ))
    expectSeen(sch.evaluate([], 500), 10, 'active', 0.5)
    expectSeen(sch.evaluate([], 1000), 10, 'active', 1)
    expectSeen(sch.evaluate([], 1500), 10, 'active', 0.5)
    expectSeen(sch.evaluate([], 1750), 10, 'active', 0.25)
    expectSeen(sch.evaluate([], 2000), 10, 'frozen', 0) // back where it started
    expect(sch.endTimeOf(10, [])).toBe(2000)
    expectSeen(sch.evaluate([], 2500), 11, 'active', 0.5, 1)
    expect(sch.endTimeOf(11, [])).toBe(4000)
    expect(easeProgress(0.75, 0, 0, true)).toBeCloseTo(0.5, 12)
  })

  it('accel 0.3 decel 0.3 matches the closed-form ease at 0.1, 0.5 and 0.9', () => {
    // Rate ramps linearly: r = 1 / (1 - 0.15 - 0.15) = 10/7.
    // 0.1 < accel: r * 0.1² / (2 * 0.3) = (10/7) / 60 = 1/42.
    // 0.5 in the middle: r * (0.5 - 0.15) = 0.5.
    // 0.9 > 1 - decel: 1 - r * 0.1² / (2 * 0.3) = 41/42.
    expect(easeProgress(0.1, 0.3, 0.3, false)).toBeCloseTo(1 / 42, 12)
    expect(easeProgress(0.5, 0.3, 0.3, false)).toBeCloseTo(0.5, 12)
    expect(easeProgress(0.9, 0.3, 0.3, false)).toBeCloseTo(41 / 42, 12)
    expect(easeProgress(0, 0.3, 0.3, false)).toBe(0)
    expect(easeProgress(1, 0.3, 0.3, false)).toBeCloseTo(1, 12)
    expect(easeProgress(0.3, 0, 0, false)).toBe(0.3)
    // Accel only, 0.5: r = 1/(1 - 0.25) = 4/3; 0.25 < 0.5 → (4/3) * 0.0625 / 1 = 1/12.
    expect(easeProgress(0.25, 0.5, 0, false)).toBeCloseTo(1 / 12, 12)

    const sch = createScheduler(tree(bhv(10, { durMs: 1000, accel: 0.3, decel: 0.3 })))
    expectSeen(sch.evaluate([], 100), 10, 'active', 1 / 42)
    expectSeen(sch.evaluate([], 500), 10, 'active', 0.5)
    expectSeen(sch.evaluate([], 900), 10, 'active', 41 / 42)
  })

  it('fill hold/freeze keeps the end value; remove drops it', () => {
    const sch = createScheduler(tree(
      bhv(10, { durMs: 500, fill: 'hold' }),
      bhv(11, { durMs: 500, fill: 'freeze' }),
      bhv(12, { durMs: 500, fill: 'remove' }),
      bhv(13, { durMs: 500, fill: 'transition' }),
      bhv(14, { durMs: 500 }), // no fill, dur given → SMIL auto → remove
      par(5, [bhv(15, {})], { durMs: 300, fill: 'hold' }), // no fill, no dur → auto → freeze when cut
      par(6, [bhv(16, { durMs: 100, fill: 'hold' })], { durMs: 300, fill: 'remove' }), // removed parent takes held children with it
    ))
    const s = sch.evaluate([], 600)
    expectSeen(s, 10, 'frozen', 1)
    expectSeen(s, 11, 'frozen', 1)
    expect(find(s, 12).state).toBe('none')
    expectSeen(s, 13, 'frozen', 1)
    expect(find(s, 14).state).toBe('none')
    expectSeen(s, 15, 'frozen', 0)
    expect(find(s, 16).state).toBe('none')
    for (const id of [10, 11, 12, 13, 14, 15, 16, 5, 6]) expect(s.ended.has(id), `ended ${id}`).toBe(true)
    expectSeen(sch.evaluate([], 200), 16, 'frozen', 1) // held while its parent is active
  })

  it('restart always, whenNotActive and never while active', () => {
    const click = (t: number): ShowEvent => ({ kind: 'trigger', spid: 'a', t })
    const onA: Condition = { delayMs: 0, evt: 'onClick', target: { spid: 'a' } }
    const make = (restart?: CTn['restart']) => createScheduler(tree(bhv(10, { durMs: 1000, fill: 'hold', start: [onA], restart })))
    const two = [click(0), click(400)]
    const three = [...two, click(1200)]

    const always = make('always')
    expectSeen(always.evaluate(two, 600), 10, 'active', 0.2) // restarted at 400
    expect(always.endTimeOf(10, two)).toBe(1400)
    expect(always.endTimeOf(10, three)).toBe(2200)
    expectSeen(make(undefined).evaluate(two, 600), 10, 'active', 0.2) // default restart is always

    const wna = make('whenNotActive')
    expectSeen(wna.evaluate(two, 600), 10, 'active', 0.6) // 400 ignored while active
    expectSeen(wna.evaluate(three, 1300), 10, 'active', 0.1) // 1200 restarts after the end
    expect(wna.endTimeOf(10, three)).toBe(2200)

    const never = make('never')
    expectSeen(never.evaluate(two, 600), 10, 'active', 0.6)
    expectSeen(never.evaluate(three, 1300), 10, 'frozen', 1)
    expect(never.endTimeOf(10, three)).toBe(1000)
  })

  it('parallel effects ending at different times', () => {
    const sch = createScheduler(tree(par(5, [
      bhv(10, { durMs: 500, fill: 'hold' }),
      bhv(11, { durMs: 1000, fill: 'hold' }),
      bhv(12, { durMs: 1500, start: [delay(250)], fill: 'remove' }),
    ], { fill: 'hold' })))
    let s = sch.evaluate([], 750)
    expectSeen(s, 10, 'frozen', 1)
    expectSeen(s, 11, 'active', 0.75)
    expectSeen(s, 12, 'active', 500 / 1500)
    s = sch.evaluate([], 1000)
    expectSeen(s, 11, 'frozen', 1)
    expect([...s.ended].sort((m, n) => m - n)).toEqual([10, 11])
    s = sch.evaluate([], 1500)
    expectSeen(s, 12, 'active', 1250 / 1500)
    s = sch.evaluate([], 1750)
    expect(find(s, 12).state).toBe('none')
    expect([...s.ended].sort((m, n) => m - n)).toEqual([5, 10, 11, 12])
    expect(sch.endTimeOf(5, [])).toBe(1750)
    // Document order inside each list (the composition order S1-11 consumes).
    expect(sch.evaluate([], 1500).frozen.map((e) => e.node.ctn.id)).toEqual([10, 11])
  })

  it('a held effect keeps its final iteration count in frozen entries', () => {
    const sch = createScheduler(tree(
      bhv(10, { durMs: 400, repeatCount: 3, fill: 'hold' }),
      bhv(11, { durMs: 400, repeatCount: 2.5, fill: 'hold' }),
      // parent L = 600 (child 300 × 2), parent AD = 1200; frozen parent leaves the child at its last iteration end.
      par(5, [bhv(12, { durMs: 300, repeatCount: 2, fill: 'hold' })], { repeatCount: 2, fill: 'hold' }),
    ))
    const s = sch.evaluate([], 2000)
    expectSeen(s, 10, 'frozen', 1, 2)
    expectSeen(s, 11, 'frozen', 0.5, 2)
    expectSeen(s, 12, 'frozen', 1, 1)
    expect(sch.endTimeOf(5, [])).toBe(1200)
  })

  it('onMouseOver, onPrev and onStopAudio start conditions fire from their events', () => {
    const sch = createScheduler(tree(
      bhv(10, { durMs: 1000, fill: 'hold', start: [{ delayMs: 0, evt: 'onMouseOver', target: { spid: 's' } }] }),
      bhv(11, { durMs: 1000, fill: 'hold', start: [{ delayMs: 200, evt: 'onPrev' }] }),
      bhv(12, { durMs: 1000, fill: 'hold', start: [{ delayMs: 0, evt: 'onStopAudio' }] }),
      bhv(13, { durMs: 1000, fill: 'hold', start: [{ delayMs: 0, evt: 'onMouseOut', target: { spid: 's' } }] }),
    ))
    const events: ShowEvent[] = [
      { kind: 'mouseOver', spid: 'other', t: 100 },
      { kind: 'mouseOver', spid: 's', t: 300 },
      { kind: 'previous', t: 500 },
      { kind: 'stopAudio', t: 900 },
      { kind: 'mouseOut', spid: 's', t: 950 },
    ]
    let s = sch.evaluate(events, 250)
    expect(s.active).toEqual([])
    s = sch.evaluate(events, 1000)
    expectSeen(s, 10, 'active', 0.7) // (1000-300)/1000
    expectSeen(s, 11, 'active', 0.3) // begins 500+200
    expectSeen(s, 12, 'active', 0.1)
    expectSeen(s, 13, 'active', 0.05)
    // Events later than t have no effect yet.
    expect(find(sch.evaluate(events, 299), 10).state).toBe('none')
  })

  it('mediaEnd ends a media node and fires what waits on its end', () => {
    const audio: BehaviourNode = { kind: 'audio', ctn: ctn(10, { fill: 'remove' }), target: { snd: 'rId3' }, attrNames: [], additive: 'base', accumulate: 'none' }
    const sch = createScheduler(tree(audio, bhv(11, { durMs: 1000, fill: 'hold', start: [{ delayMs: 100, evt: 'onEnd', tn: 10 }] })))
    expect(sch.endTimeOf(10, [])).toBe('indefinite') // intrinsic length unknown until the media says so
    expect(sch.endTimeOf(11, [])).toBeUndefined()
    const ev: ShowEvent[] = [{ kind: 'mediaEnd', nodeId: 10, t: 2000, synthetic: false }]
    expect(sch.endTimeOf(10, ev)).toBe(2000)
    const s = sch.evaluate(ev, 2600)
    expect(s.ended.has(10)).toBe(true)
    expectSeen(s, 11, 'active', 0.5) // begins 2000 + 100
    expect(find(sch.evaluate(ev, 1999), 10).state).toBe('active')
  })

  it('the main sequence advances on next; a next during an active step settles it', () => {
    // PowerPoint shape: mainSeq (concurrent, nextAc seek) of click pars waiting on delay="indefinite".
    const main = seq(2, [
      par(3, [par(4, [bhv(10, { durMs: 1000, fill: 'hold' })], { fill: 'hold', start: [delay(0)] })], { fill: 'hold', start: [WAIT] }),
      par(5, [bhv(11, { durMs: 500, fill: 'hold' }), bhv(12, { durMs: 400, repeatCount: 'indefinite', fill: 'hold' })], { fill: 'hold', start: [WAIT] }),
      par(9, [bhv(19, { durMs: 100, fill: 'hold' })], { fill: 'hold', start: [WAIT] }),
    ], { durMs: 'indefinite', nodeType: 'mainSeq' }, { concurrent: true, nextAc: 'seek', prevAc: 'none' })
    const sch = createScheduler(tree(main))
    expect(sch.evaluate([], 500).active).toEqual([])
    const n1: ShowEvent[] = [{ kind: 'next', t: 100 }]
    expectSeen(sch.evaluate(n1, 600), 10, 'active', 0.5)
    expect(sch.endTimeOf(3, n1)).toBe(1100)
    const n2: ShowEvent[] = [...n1, { kind: 'next', t: 300 }]
    let s = sch.evaluate(n2, 400)
    expectSeen(s, 10, 'frozen', 1) // seeked to its end by the second next
    expectSeen(s, 11, 'active', 0.2)
    expect(sch.endTimeOf(3, n2)).toBe(300)
    // Settling a step with an indefinite effect: finite parts end, the indefinite one keeps running.
    const n3: ShowEvent[] = [...n2, { kind: 'next', t: 500 }]
    s = sch.evaluate(n3, 700)
    expectSeen(s, 11, 'frozen', 1)
    expect(find(s, 12).state).toBe('active')
    expect(sch.endTimeOf(5, n3)).toBe('indefinite')

    // Interactive sequence: the trigger begins it, later triggers advance it.
    const btn: Condition = { delayMs: 0, evt: 'onClick', target: { spid: 'btn' } }
    const inter = seq(6, [
      par(7, [bhv(13, { durMs: 400, fill: 'hold' })], { fill: 'hold', start: [delay(0)] }),
      par(8, [bhv(14, { durMs: 400, fill: 'hold' })], { fill: 'hold', start: [WAIT] }),
    ], { restart: 'whenNotActive', fill: 'hold', nodeType: 'interactiveSeq', start: [btn], endSync: { delayMs: 0, evt: 'end', rtn: 'all' } }, { concurrent: true, nextAc: 'seek', prevAc: 'none' })
    const isch = createScheduler(tree(inter))
    const t1: ShowEvent[] = [{ kind: 'trigger', spid: 'btn', t: 1000 }]
    expectSeen(isch.evaluate(t1, 1200), 13, 'active', 0.5)
    expect(find(isch.evaluate(t1, 1200), 14).state).toBe('none')
    expect(isch.endTimeOf(6, t1)).toBe('indefinite') // waits for the second click
    const t2: ShowEvent[] = [...t1, { kind: 'trigger', spid: 'btn', t: 1100 }]
    s = isch.evaluate(t2, 1200)
    expectSeen(s, 13, 'frozen', 1)
    expectSeen(s, 14, 'active', 0.25)
    expect(isch.endTimeOf(6, t2)).toBe(1500)
  })

  it('hostile timing values and reference cycles never throw or hang', () => {
    const selfRef: Condition = { delayMs: 0, evt: 'onBegin', tn: 10 }
    const chain: TimeNode[] = []
    for (let i = 0; i < 2000; i++) chain.push(bhv(1000 + i, { durMs: 1, start: i === 0 ? [] : [{ delayMs: 0, evt: 'onEnd', tn: 1000 + i - 1 }] }))
    const reverse: TimeNode[] = []
    for (let i = 0; i < 2000; i++) reverse.push(bhv(5000 + i, { durMs: 1, start: i === 1999 ? [] : [{ delayMs: 0, evt: 'onEnd', tn: 5000 + i + 1 }] }))
    const seqChain: TimeNode[] = []
    for (let i = 0; i < 2000; i++) seqChain.push(bhv(9000 + i, { durMs: 1, fill: 'hold' }))
    const sch = createScheduler(tree(
      bhv(10, { durMs: 100, start: [selfRef] }),
      bhv(11, { durMs: 100, start: [{ delayMs: 0, evt: 'onEnd', tn: 12 }] }),
      bhv(12, { durMs: 100, start: [{ delayMs: 0, evt: 'onEnd', tn: 11 }] }),
      bhv(13, { durMs: Number.NaN, repeatCount: -3, repeatDurMs: -1, accel: 5, decel: 5, start: [delay(Number.NaN)] }),
      bhv(14, { durMs: -50, repeatCount: Number.POSITIVE_INFINITY, accel: Number.NaN }),
      bhv(15, { durMs: 100, start: [{ delayMs: 0, evt: 'onEnd', tn: 424242 }] }),
      par(16, chain),
      par(17, reverse),
      seq(18, seqChain, { fill: 'hold' }),
    ))
    const s = sch.evaluate([], 50)
    expect(find(s, 10).state).toBe('none')
    expect(find(s, 11).state).toBe('none')
    expect(find(s, 12).state).toBe('none')
    expect(find(s, 15).state).toBe('none')
    expect(() => sch.evaluate([], Number.NaN)).not.toThrow()
    expect(() => sch.evaluate([], Number.POSITIVE_INFINITY)).not.toThrow()
    for (const e of [...s.active, ...s.frozen]) {
      expect(Number.isFinite(e.progress)).toBe(true)
      expect(e.progress).toBeGreaterThanOrEqual(0)
      expect(e.progress).toBeLessThanOrEqual(1)
    }
    expect(sch.endTimeOf(1999 + 1000, [])).toBe(2000)
    expect(sch.endTimeOf(9000 + 1999, [])).toBe(2000)
    expect(sch.endTimeOf(5000, [])).toBe(2000) // reverse document order still resolves
    expect(sch.endTimeOf(424242, [])).toBeUndefined()
    expect(sch.endTimeOf(10, [])).toBeUndefined()
  })

  // -------------------------------------------------------------------------
  // Fix round 1 (codex review I1, I2)
  // -------------------------------------------------------------------------

  it('nested time bases agree: an eased parent maps samples, events and end times the same way', () => {
    // Parent dur 100, accel 1: child local time = 100 * (t/100)^2 (r = 2, ease(q) = q^2).
    const sch = createScheduler(tree(
      par(5, [bhv(10, { durMs: 50, fill: 'hold' })], { durMs: 100, accel: 1, fill: 'hold' }),
      par(6, [bhv(11, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onClick', target: { spid: 'z' } }] })], { durMs: 100, accel: 1, fill: 'hold' }),
      // A top-level node waiting on the end of a node inside the eased parent.
      bhv(12, { durMs: 100, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 10 }] }),
    ))
    const end10 = 100 * Math.sqrt(0.5) // local 50 is reached at 70.7106781...
    expect(sch.endTimeOf(10, [])).toBeCloseTo(end10, 9)
    expectSeen(sch.evaluate([], 50), 10, 'active', 0.5) // local 25 of 50
    expect(find(sch.evaluate([], end10 - 1e-6), 10).state).toBe('active')
    expectSeen(sch.evaluate([], end10 + 1e-6), 10, 'frozen', 1)

    // A trigger at absolute 50 begins the child at parent-local 25 (not 50).
    const click: ShowEvent[] = [{ kind: 'trigger', spid: 'z', t: 50 }]
    expectSeen(sch.evaluate(click, 50), 11, 'active', 0)
    expectSeen(sch.evaluate(click, 55), 11, 'active', (100 * 0.55 * 0.55 - 25) / 10) // local 30.25
    expect(sch.endTimeOf(11, click)).toBeCloseTo(100 * Math.sqrt(0.35), 9) // local 35 at 59.16...

    // The syncbase dependent starts when node 10 really ends.
    expect(sch.endTimeOf(12, [])).toBeCloseTo(end10 + 100, 9)
    expectSeen(sch.evaluate([], end10 + 10), 12, 'active', 0.1)
  })

  it('nested time bases agree under excl pauses, autoRev parents and repeating parents', () => {
    const clock = new ManualClock()
    const sch = createScheduler(tree(
      // The par (dur 100) is paused by node 13 over [25, 75): it ends at 150, and so does its child.
      excl(4, [par(6, [bhv(12, { durMs: 100, fill: 'hold' })], { durMs: 100, fill: 'hold' }), bhv(13, { durMs: 50, start: [delay(25)] })], { fill: 'hold' }),
      // autoRev parent: local = t on [0,100], 200 - t on [100,200].
      par(7, [bhv(14, { durMs: 50, fill: 'hold' })], { durMs: 100, autoRev: true, fill: 'hold' }),
      // Repeating parent: a trigger in iteration 1 maps to local 30 of that iteration.
      par(9, [bhv(17, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onClick', target: { spid: 'y' } }] })], { durMs: 100, repeatCount: 2, fill: 'hold' }),
    ))
    clock.advance(100)
    expectSeen(sch.evaluate([], clock.now()), 12, 'active', 0.5)
    expect(sch.endTimeOf(12, [])).toBe(150)
    clock.advance(50)
    expectSeen(sch.evaluate([], clock.now()), 12, 'frozen', 1)

    expectSeen(sch.evaluate([], 75), 14, 'frozen', 1)
    expectSeen(sch.evaluate([], 160), 14, 'active', 0.8) // the reverse half replays it: local 40
    expectSeen(sch.evaluate([], 200), 14, 'frozen', 0)
    expect(sch.endTimeOf(14, [])).toBe(200) // its last activation ends with the parent

    const y: ShowEvent[] = [{ kind: 'trigger', spid: 'y', t: 130 }]
    expectSeen(sch.evaluate(y, 135), 17, 'active', 0.5)
    expect(sch.endTimeOf(17, y)).toBe(140)
    // A trigger in iteration 0 does not carry into iteration 1.
    const early: ShowEvent[] = [{ kind: 'trigger', spid: 'y', t: 50 }]
    expectSeen(sch.evaluate(early, 70), 17, 'frozen', 1)
    expect(find(sch.evaluate(early, 150), 17).state).toBe('none')
  })

  it('syncbase history survives a repeating parent: earlier iterations still resolve', () => {
    // codex R1 probe. Parent: dur 100, repeat 2. Child 3 (dur 10) is clicked at 50, so it runs [50, 60) in iteration 0.
    // Top-level node 4 (dur 200) begins on child 3's end: [60, 260).
    const sch = createScheduler(tree(
      par(2, [bhv(3, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onClick', target: { spid: 'a' } }] })], { durMs: 100, repeatCount: 2, fill: 'hold' }),
      bhv(4, { durMs: 200, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 3 }] }),
      // Begin references too: node 5 begins 5 ms after each begin of time-based child 6 (iterations at 0 and 100).
      par(7, [bhv(6, { durMs: 10, fill: 'hold' })], { durMs: 100, repeatCount: 2, fill: 'hold' }),
      bhv(5, { durMs: 50, fill: 'hold', start: [{ delayMs: 5, evt: 'onBegin', tn: 6 }] }),
    ))
    const ev: ShowEvent[] = [{ kind: 'trigger', spid: 'a', t: 50 }]
    expectSeen(sch.evaluate(ev, 70), 4, 'active', 10 / 200)
    expectSeen(sch.evaluate(ev, 150), 4, 'active', 90 / 200) // iteration 1 does not erase iteration 0's end
    expect(find(sch.evaluate(ev, 150), 3).state).toBe('none') // ...and does not replay the click in iteration 1
    expect(sch.endTimeOf(3, ev)).toBe(60) // its last activation was in iteration 0
    expect(sch.endTimeOf(4, ev)).toBe(260)
    expectSeen(sch.evaluate(ev, 260), 4, 'frozen', 1)

    expectSeen(sch.evaluate([], 30), 5, 'active', 25 / 50) // began at 5
    expectSeen(sch.evaluate([], 130), 5, 'active', 25 / 50) // restarted at 105 by iteration 1's begin
    expect(sch.endTimeOf(5, [])).toBe(155)
    expect(sch.endTimeOf(6, [])).toBe(110)
  })

  it('the syncbase history window is bounded; a reference past it is refused and diagnosed, never misresolved', () => {
    // codex R2 probe: the clicked child runs [50, 60) in iteration 0 of a parent repeated 1,500 times (dur 100).
    // The history window holds the most recent 1,000 parent iterations.
    const sink = createDiagnosticSink()
    const sch = createScheduler(tree(
      par(2, [bhv(3, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onClick', target: { spid: 'a' } }] })], { durMs: 100, repeatCount: 1500, fill: 'hold' }),
      bhv(4, { durMs: 200_000, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 3 }] }),
      // A time-based child activates every iteration: past the window its reference is refused too,
      // never answered from a more recent iteration.
      par(7, [bhv(6, { durMs: 10, fill: 'hold' })], { durMs: 100, repeatCount: 1500, fill: 'hold' }),
      bhv(5, { durMs: 50, fill: 'hold', start: [{ delayMs: 5, evt: 'onBegin', tn: 6 }] }),
    ), { sink, slideIndex: 2 })
    const ev: ShowEvent[] = [{ kind: 'trigger', spid: 'a', t: 50 }]
    // Iteration 999: 1,000 iterations, all inside the window.
    expectSeen(sch.evaluate(ev, 99_950), 4, 'active', (99_950 - 60) / 200_000)
    expectSeen(sch.evaluate(ev, 99_930), 5, 'active', 25 / 50) // began at 99,905
    expect(sink.list()).toEqual([])
    // Iteration 1,001: iteration 0 has left the window.
    let s = sch.evaluate(ev, 100_150)
    expect(find(s, 4).state).toBe('none')
    expect(find(s, 5).state).toBe('none')
    s = sch.evaluate(ev, 100_250)
    expect(sch.endTimeOf(3, ev)).toBeUndefined() // the window holds no activation of node 3
    expect(sch.endTimeOf(4, ev)).toBeUndefined()
    expect(sch.endTimeOf(6, [])).toBe(149_910) // the newest activation is always inside the window
    const codes = sink.list().map((d) => `${d.code}:${d.effect}`).sort()
    expect(codes).toEqual([
      'timing-history-truncated:node 3',
      'timing-history-truncated:node 6',
      'timing-reference-outside-window:node 4',
      'timing-reference-outside-window:node 5',
    ])
    expect(sink.list().every((d) => d.slideIndex === 2)).toBe(true)
  })

  it('history construction stays inside the window for nested repeating parents', () => {
    // codex R3: outer dur 2000 x 1500, inner dur 1 x 2000 inside it, child dur 1.
    // Unbounded enumeration would build ~1,000 x 2,000 frames before trimming.
    const sink = createDiagnosticSink()
    const sch = createScheduler(tree(
      par(2, [par(3, [bhv(4, { durMs: 1, fill: 'hold' })], { durMs: 1, repeatCount: 2000, fill: 'hold' })], { durMs: 2000, repeatCount: 1500, fill: 'hold' }),
      bhv(5, { durMs: 10, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 4 }] }),
    ), { sink })
    const t0 = performance.now()
    const s = sch.evaluate([], 2_200_500)
    const elapsed = performance.now() - t0
    expect(find(s, 5).state).toBe('none')
    expect(sink.list().map((d) => d.code).sort()).toEqual(['timing-history-truncated', 'timing-reference-outside-window'])
    expect(elapsed).toBeLessThan(5_000) // generous: a bounded window is ~1,000 frames per level
  })

  it('endSync tn takes the minimum of a million in-window end times without overflowing the stack', () => {
    // codex R5: node 5 has 1,000 start offsets (0..999 ms, dur 1, restart cuts each at the next) in each
    // of 1,000 iterations of par 3 (dur 1,001): 10^6 end times, inside the history window. Spreading them
    // into Math.min overflowed the stack. The earliest end is iteration 0's first activation: 1.
    const starts: Condition[] = []
    for (let i = 0; i < 1000; i++) starts.push(delay(i))
    const sch = createScheduler(tree(
      par(2, [par(3, [bhv(5, { durMs: 1, fill: 'hold', start: starts })], { durMs: 1001, repeatCount: 1000, fill: 'hold' })], { fill: 'hold', endSync: { delayMs: 0, tn: 5 } }),
    ))
    expect(() => sch.evaluate([], 1001 * 1000 + 5)).not.toThrow()
    expect(sch.endTimeOf(2, [])).toBe(1)
  })

  it('chained history references share one per-query budget; past it they are refused and diagnosed', () => {
    // codex R4 cubic probe: B(n) holds T (onEnd W) and E(n) ⊃ [W (onEnd V), G(n) ⊃ V]; node 8 at the root
    // begins when T ends. Each level is inside the 1,000-iteration window, but the chained references
    // admit n + n^2 + n^3 history frames in one query.
    const build = (n: number, sink: ReturnType<typeof createDiagnosticSink>) => {
      const dE = n + 10
      const dB = n * dE + 10
      const sch = createScheduler(tree(
        par(2, [
          bhv(3, { durMs: 1, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 5 }] }),
          par(4, [
            bhv(5, { durMs: 1, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 7 }] }),
            par(6, [bhv(7, { durMs: 1, fill: 'hold' })], { durMs: 1, repeatCount: n, fill: 'hold' }),
          ], { durMs: dE, repeatCount: n, fill: 'hold' }),
        ], { durMs: dB, repeatCount: n, fill: 'hold' }),
        bhv(8, { durMs: 1_000_000, fill: 'hold', restart: 'never', start: [{ delayMs: 0, evt: 'onEnd', tn: 3 }] }),
      ), { sink, slideIndex: 4 })
      return { sch, t: dB * n + 10 }
    }
    // An ordinary size (n = 10, ~1,100 frames) is unaffected. V ends at 1, so W runs [1, 2) and T runs
    // [2, 3): node 8 begins at 3.
    const okSink = createDiagnosticSink()
    const ok = build(10, okSink)
    expectSeen(ok.sch.evaluate([], ok.t), 8, 'active', (ok.t - 3) / 1_000_000)
    expect(ok.sch.endTimeOf(8, [])).toBe(3 + 1_000_000)
    expect(okSink.list()).toEqual([])

    // n = 100 needs ~10^6 frames (seconds unbounded, effectively a hang at 1,000). The budget stops it,
    // and the references still unresolved when it runs out are refused, never answered from part of the history.
    const sink = createDiagnosticSink()
    const big = build(100, sink)
    const t0 = performance.now()
    const s = big.sch.evaluate([], big.t)
    expect(big.sch.endTimeOf(8, [])).toBeUndefined()
    const elapsed = performance.now() - t0
    expect(find(s, 8).state).toBe('none')
    const codes = sink.list().map((d) => d.code)
    expect(codes).toContain('timing-query-budget')
    expect(new Set(codes)).toEqual(new Set(['timing-query-budget']))
    expect(sink.list().some((d) => d.effect === 'node 8')).toBe(true)
    expect(sink.list().every((d) => d.slideIndex === 4)).toBe(true)
    // Once per node per scheduler, never per frame.
    const before = sink.list().length
    big.sch.evaluate([], big.t)
    expect(sink.list().length).toBe(before)
    expect(elapsed).toBeLessThan(3_000) // generous: the budget is 100,000 frames (~0.2 s)
  })

  it('endTimeOf does not report truncation when its answer is inside the window', () => {
    // codex M2: the newest activation of node 6 is inside the window, so 149,910 is correct and undiagnosed.
    const sink = createDiagnosticSink()
    const sch = createScheduler(tree(
      par(7, [bhv(6, { durMs: 10, fill: 'hold' })], { durMs: 100, repeatCount: 1500, fill: 'hold' }),
    ), { sink })
    expect(sch.endTimeOf(6, [])).toBe(149_910)
    expect(sink.list()).toEqual([])
  })

  it('a restart storm past 10,000 activations is refused and diagnosed, never truncated silently', () => {
    // codex M4: each start offset cuts the previous activation (restart always).
    const offsets = (n: number): Condition[] => Array.from({ length: n }, (_, i) => delay(i))
    const sink = createDiagnosticSink()
    const sch = createScheduler(tree(
      bhv(10, { durMs: 1, fill: 'hold', start: offsets(10_000) }),
      bhv(11, { durMs: 1, fill: 'hold', start: offsets(10_001) }),
      bhv(12, { durMs: 5, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 11 }] }),
    ), { sink, slideIndex: 1 })
    const s = sch.evaluate([], 20_000)
    expectSeen(s, 10, 'frozen', 1, 0) // 10,000 activations: inside the bound, last ends at 10,000
    expect(sch.endTimeOf(10, [])).toBe(10_000)
    expect(find(s, 11).state).toBe('none')
    expect(find(s, 12).state).toBe('none') // its syncbase never begins
    expect(sch.endTimeOf(11, [])).toBeUndefined()
    expect(sink.list().map((d) => `${d.code}:${d.effect}`)).toEqual(['timing-too-many-activations:node 11'])
    expect(sink.list()[0].slideIndex).toBe(1)
  })

  it('a 600-node reverse reference chain resolves exactly', () => {
    // Node 100 + i begins when node 101 + i ends; node 699 begins at 0. All 1 ms: node 100 ends at 600.
    const kids: TimeNode[] = []
    for (let i = 0; i < 600; i++) kids.push(bhv(100 + i, { durMs: 1, fill: 'hold', start: i === 599 ? [] : [{ delayMs: 0, evt: 'onEnd', tn: 101 + i }] }))
    const sch = createScheduler(tree(...kids))
    expect(sch.endTimeOf(100, [])).toBe(600)
    expect(sch.endTimeOf(400, [])).toBe(300)
    expectSeen(sch.evaluate([], 599.5), 100, 'active', 0.5)
    const s = sch.evaluate([], 600)
    expect(s.ended.has(100)).toBe(true)
    expect(s.frozen.length).toBe(600)
  })

  it('a real reference cycle is broken once, with a diagnostic, and the rest still plays', () => {
    const sink = createDiagnosticSink()
    let adds = 0
    const counting = { add: (d: Parameters<typeof sink.add>[0]) => { adds++; sink.add(d) }, list: () => sink.list() }
    const sch = createScheduler(tree(
      // A starts at 0 or when B ends; B starts when A ends: the references form a cycle.
      bhv(10, { durMs: 100, fill: 'hold', start: [delay(0), { delayMs: 0, evt: 'onEnd', tn: 11 }] }),
      bhv(11, { durMs: 100, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn: 10 }] }),
      bhv(12, { durMs: 100, fill: 'hold' }),
    ), { sink: counting, slideIndex: 3 })
    let s = sch.evaluate([], 150)
    expectSeen(s, 10, 'frozen', 1)
    expectSeen(s, 11, 'active', 0.5) // B still follows A's first end
    expectSeen(s, 12, 'frozen', 1)
    s = sch.evaluate([], 160)
    sch.endTimeOf(10, [])
    const cyc = sink.list().filter((d) => d.code === 'timing-cycle')
    expect(cyc.length).toBe(1)
    expect(cyc[0].slideIndex).toBe(3)
    expect(adds).toBe(1) // reported once per scheduler, never per frame
  })

  it('evaluate is pure: same tree, events and t give deep-equal snapshots', () => {
    const build = (): TimingTree => tree(
      seq(2, [
        par(3, [bhv(10, { durMs: 800, accel: 0.2, decel: 0.3, fill: 'hold' }), bhv(11, { durMs: 300, autoRev: true, repeatCount: 2 })], { fill: 'hold', start: [WAIT] }),
        par(4, [bhv(12, { durMs: 500, fill: 'hold' })], { fill: 'hold', start: [WAIT] }),
      ], { durMs: 'indefinite', nodeType: 'mainSeq' }, { concurrent: true, nextAc: 'seek', prevAc: 'none' }),
      excl(5, [bhv(13, { durMs: 1000, fill: 'hold' }), bhv(14, { durMs: 200, start: [{ delayMs: 0, evt: 'onClick', target: { spid: 'k' } }] })]),
    )
    const events: readonly ShowEvent[] = Object.freeze([
      { kind: 'next', t: 50 },
      { kind: 'trigger', spid: 'k', t: 250 },
      { kind: 'next', t: 400 },
    ] as ShowEvent[])
    const a = createScheduler(build())
    const b = createScheduler(build())
    for (const t of [0, 120, 260, 399, 400, 650, 2000]) {
      const first = a.evaluate(events, t)
      expect(a.evaluate(events, t)).toEqual(first)
      expect(b.evaluate(events, t)).toEqual(first)
    }
    // Event order in the input array does not matter.
    expect(a.evaluate([...events].reverse(), 650)).toEqual(a.evaluate(events, 650))
    expect(a.endTimeOf(3, events)).toEqual(b.endTimeOf(3, events))
  })

  it('S1-10: a seek at a step\'s begin instant settles it to its end', () => {
    // mainSeq [A waits, B waits (dur 0)]: two nexts at 100 begin A and B together; B seeks A from elapsed 0 to 1000.
    const a = par(3, [par(4, [bhv(8, { durMs: 1000, fill: 'hold' })], { start: [delay(0)], fill: 'hold' })], { start: [WAIT], fill: 'hold' })
    const b = par(5, [], { durMs: 0, start: [WAIT] })
    const s = createScheduler(tree(seq(2, [a, b], { durMs: 'indefinite', nodeType: 'mainSeq' }, { concurrent: false, nextAc: 'seek', prevAc: 'none' })))
    const events: ShowEvent[] = [{ kind: 'next', t: 100 }, { kind: 'next', t: 100 }]
    expectSeen(s.evaluate(events, 100), 8, 'frozen', 1)
    expectSeen(s.evaluate(events, 150), 8, 'frozen', 1)
  })

  it('S1-10: seek settles finite effects nested beside an indefinite one', () => {
    // Step A: inner par holding a 1000 ms effect (8) and an indefinite 500 ms spin (9). B seeks A at 200: settle 1000, delta 800.
    const a = par(3, [par(4, [bhv(8, { durMs: 1000, fill: 'hold' }), bhv(9, { durMs: 500, repeatCount: 'indefinite' })], { start: [delay(0)], fill: 'hold' })], { start: [WAIT], fill: 'hold' })
    const b = par(5, [bhv(10, { durMs: 1000, fill: 'hold' })], { start: [WAIT], fill: 'hold' })
    const s = createScheduler(tree(seq(2, [a, b], { durMs: 'indefinite', nodeType: 'mainSeq' }, { concurrent: false, nextAc: 'seek', prevAc: 'none' })))
    const snap = s.evaluate([{ kind: 'next', t: 0 }, { kind: 'next', t: 200 }], 300)
    expectSeen(snap, 8, 'frozen', 1)
    expectSeen(snap, 9, 'active', 0.2, 2) // local 300 + 800 = 1100: turn 2 at 100/500
    expectSeen(snap, 10, 'active', 0.1)
  })
})
