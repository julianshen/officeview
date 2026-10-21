/**
 * S1-9 timing scheduler core (slideshow spec 6.2). Every tree is built here
 * as a `TimingTree` value; time is plain numbers (milliseconds of show
 * clock), never wall time. Expected values are hand-computed in comments.
 */
import { describe, expect, it } from 'vitest'
import { createScheduler, type SchedulerSnapshot } from '../src/slideshow/engine/scheduler'
import { easeProgress } from '../src/slideshow/engine/ease'
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
    expect(() => sch.endTimeOf(5000, [])).not.toThrow()
    expect(sch.endTimeOf(424242, [])).toBeUndefined()
    expect(sch.endTimeOf(10, [])).toBeUndefined()
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
})
