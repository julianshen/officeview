/**
 * S1-10 navigation layer (slideshow spec 6.2 click steps, navigation rules
 * and slide states; 4.5 reduced motion). Trees are `TimingTree` values in
 * PowerPoint's shape (click group > inner par > effect par > behaviours);
 * time is plain show-clock milliseconds. Expected values are hand-computed
 * in comments.
 */
import { existsSync, readFileSync } from 'node:fs'
import { resolve as resolvePath } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { buildPptxDeck, mainSeqXml, timingXmlOf } from '../src/testdata/ooxml-builders'
import type { SchedulerSnapshot } from '../src/slideshow/engine/scheduler'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../src/slideshow/engine/steps'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import type { AnimMotionBehaviour, BehaviourNode, CTn, Condition, ContainerNode, TimeNode, TimingTree } from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Builders
// ---------------------------------------------------------------------------

type CtnOver = Partial<Omit<CTn, 'id'>>
const ctn = (id: number, over: CtnOver = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over })
const WAIT: Condition = { delayMs: 'indefinite' }
const at = (delayMs: number): Condition => ({ delayMs })

/** A behaviour on shape `spid` (default: its own id). */
function bhv(id: number, over: CtnOver = {}, spid = String(id)): BehaviourNode {
  return { kind: 'set', ctn: ctn(id, { fill: 'hold', ...over }), target: { spid }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' }
}
const par = (id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode => ({ kind: 'par', ctn: ctn(id, { fill: 'hold', ...over }), children })
/** An effect par (`nodeType` click/with/after) starting `delay` into its inner par. */
const effect = (id: number, nodeType: string, children: TimeNode[], delay = 0) => par(id, children, { nodeType, start: [at(delay)] })
/** A click group: waits for "next", then runs its inner pars (each with its own start delay). */
const group = (id: number, inner: ContainerNode[]) => par(id, inner, { start: [WAIT] })
const inner = (id: number, effects: ContainerNode[], delay = 0) => par(id, effects, { start: [at(delay)] })
const mainSeq = (children: TimeNode[]): ContainerNode => ({
  kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', nodeType: 'mainSeq' }), children, seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})
/** An interactive sequence on trigger shape `spid` whose one effect is `children`. */
const interactive = (id: number, spid: string, children: TimeNode[]): ContainerNode => ({
  kind: 'seq',
  ctn: ctn(id, { nodeType: 'interactiveSeq', restart: 'whenNotActive', start: [{ delayMs: 0, evt: 'onClick', target: { spid } }] }),
  children: [par(id + 1, [effect(id + 2, 'clickEffect', children)], { start: [at(0)] })],
  seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})
const tree = (...children: TimeNode[]): TimingTree => ({ root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot', fill: undefined }), build: [] })

const NORMAL = { showAnimation: true, reducedMotion: false }

type Seen = { state: 'active' | 'frozen'; progress: number; iteration: number } | { state: 'none' }
function find(snap: SchedulerSnapshot, id: number): Seen {
  const a = snap.active.find(e => e.node.ctn.id === id)
  const f = snap.frozen.find(e => e.node.ctn.id === id)
  expect(a && f, `node ${id} in one list only`).toBeFalsy()
  if (a) return { state: 'active', progress: a.progress, iteration: a.iteration }
  if (f) return { state: 'frozen', progress: f.progress, iteration: f.iteration }
  return { state: 'none' }
}
function expectSeen(snap: SchedulerSnapshot, id: number, state: 'active' | 'frozen' | 'none', progress = 0, iteration = 0): void {
  const s = find(snap, id)
  expect(s.state, `node ${id} state`).toBe(state)
  if (s.state === 'none') return
  expect(s.progress, `node ${id} progress`).toBeCloseTo(progress, 9)
  expect(s.iteration, `node ${id} iteration`).toBe(iteration)
}
const ids = (snap: SchedulerSnapshot) => [...snap.active, ...snap.frozen].map(e => e.node.ctn.id).sort((a, b) => a - b)

/** Two click steps: step 1 a fade (6, 1000 ms) with an indefinite spin (7, 500 ms per turn); step 2 a fade (8, 1000 ms). Trigger shape 9 runs effect 33. */
function spinDeck(): TimingTree {
  return tree(
    mainSeq([
      group(3, [inner(4, [effect(5, 'clickEffect', [bhv(6, { durMs: 1000 })]), effect(10, 'withEffect', [bhv(7, { durMs: 500, repeatCount: 'indefinite', fill: undefined })])])]),
      group(20, [inner(21, [effect(22, 'clickEffect', [bhv(8, { durMs: 1000 })])])]),
    ]),
    interactive(30, '9', [bhv(33, { durMs: 1000 })]),
  )
}

// ---------------------------------------------------------------------------
// Click steps
// ---------------------------------------------------------------------------

describe('click steps', () => {
  it('clickEffect starts a step, withEffect joins, afterEffect chains', () => {
    // Step 1: clickEffect 6 (1000) with withEffect 8 (500), then afterEffect 11 (500) in an inner par delayed 1000.
    // Step 2: clickEffect 23 (400).
    const tl = createSlideTimeline(tree(mainSeq([
      group(3, [
        inner(4, [effect(5, 'clickEffect', [bhv(6, { durMs: 1000 })]), effect(7, 'withEffect', [bhv(8, { durMs: 500 })])]),
        inner(9, [effect(10, 'afterEffect', [bhv(11, { durMs: 500 })])], 1000),
      ]),
      group(20, [inner(21, [effect(22, 'clickEffect', [bhv(23, { durMs: 400 })])])]),
    ])), NORMAL)
    expect(tl.steps).toBe(2)
    const init = tl.initial()
    expect(init).toMatchObject({ step: 0, mode: 'initial', events: [] })
    expect(ids(tl.stateAt(init, 500))).toEqual([])
    const c1 = tl.next(init, 100)
    expect(c1).toMatchObject({ step: 1, mode: 'live' })
    let s = tl.stateAt(c1, 350) // 250 ms into step 1
    expectSeen(s, 6, 'active', 0.25)
    expectSeen(s, 8, 'active', 0.5) // withEffect joins: same begin
    expectSeen(s, 11, 'none')
    s = tl.stateAt(c1, 1350) // afterEffect begins at 100 + 1000 = 1100
    expectSeen(s, 6, 'frozen', 1)
    expectSeen(s, 8, 'frozen', 1)
    expectSeen(s, 11, 'active', 0.5)
    expectSeen(s, 23, 'none')
    const c2 = tl.next(c1, 2000)
    expect(c2.step).toBe(2)
    expectSeen(tl.stateAt(c2, 2200), 23, 'active', 0.5)
    // Nothing is running at 2500: next has nothing left on this slide and returns the same cursor.
    expect(tl.next(c2, 2500)).toBe(c2)
  })

  it('next during an active effect settles finite effects and keeps indefinite ones running', () => {
    const tl = createSlideTimeline(spinDeck(), NORMAL)
    const c1 = tl.next(tl.initial(), 0)
    let s = tl.stateAt(c1, 200)
    expectSeen(s, 6, 'active', 0.2)
    expectSeen(s, 7, 'active', 0.4)
    // Next at 200 while step 1 runs: the step container is seeked from elapsed 200 to its finite settle point 1000 (delta 800).
    const c2 = tl.next(c1, 200)
    expect(c2.step).toBe(2)
    s = tl.stateAt(c2, 300)
    expectSeen(s, 6, 'frozen', 1)
    expectSeen(s, 7, 'active', 0.2, 2) // local 300 + 800 = 1100 → turn 2 at 100/500
    expectSeen(s, 8, 'active', 0.1)

    // The last step: next settles it and stays; the next one leaves the slide.
    const c3 = tl.next(c2, 400) // step 2 still running (8 ends at 1200)
    expect(c3).not.toBe(c2)
    expect(c3.step).toBe(2)
    s = tl.stateAt(c3, 500)
    expectSeen(s, 8, 'frozen', 1)
    expectSeen(s, 6, 'frozen', 1)
    expect(find(s, 7).state).toBe('active')
    expect(tl.next(c3, 600)).toBe(c3)
    // Settling a one-step slide while its only step runs works the same way.
    const one = createSlideTimeline(tree(mainSeq([group(3, [inner(4, [effect(5, 'clickEffect', [bhv(6, { durMs: 1000 }), bhv(7, { durMs: 500, repeatCount: 'indefinite', fill: undefined })])])])])), NORMAL)
    const o1 = one.next(one.initial(), 0)
    const o2 = one.next(o1, 200)
    expect(o2).not.toBe(o1)
    expect(o2.step).toBe(1)
    s = one.stateAt(o2, 300)
    expectSeen(s, 6, 'frozen', 1)
    expectSeen(s, 7, 'active', 0.2, 2)
    expect(one.next(o2, 300)).toBe(o2)
  })

  it('two trigger histories at the same step and time give different render states', () => {
    const tl = createSlideTimeline(spinDeck(), NORMAL)
    expect(tl.triggers()).toEqual(['9'])
    const c1 = tl.next(tl.initial(), 0)
    const a = tl.trigger(c1, '9', 100)!
    const b = tl.trigger(c1, '9', 400)!
    expect(a.step).toBe(1)
    expect(b.step).toBe(1)
    expectSeen(tl.stateAt(a, 600), 33, 'active', 0.5)
    expectSeen(tl.stateAt(b, 600), 33, 'active', 0.2)
    expect(tl.stateAt(c1, 600).active.some(e => e.node.ctn.id === 33)).toBe(false)
    expect(tl.trigger(c1, '77', 100)).toBeUndefined()
  })

  it('checkpoint n ignores trigger history and takes indefinite effects at first-iteration end', () => {
    // Step 1 also gets an autoRev spin (12): the end of its first iteration is back at progress 0.
    const t = spinDeck()
    const g1 = ((t.root as ContainerNode).children[0] as ContainerNode).children[0] as ContainerNode
    ;(g1.children[0] as ContainerNode).children.push(effect(13, 'withEffect', [bhv(12, { durMs: 500, repeatCount: 'indefinite', autoRev: true, accel: 0.5, fill: undefined })]))
    const tl = createSlideTimeline(t, NORMAL)
    const cp0 = tl.checkpoint(0)
    expect(cp0).toMatchObject({ step: 0, mode: 'checkpoint' })
    expect(ids(tl.stateAt(cp0, 0))).toEqual([])
    const cp1 = tl.checkpoint(1)
    expect(cp1).toMatchObject({ step: 1, mode: 'checkpoint' })
    for (const time of [0, 12345]) {
      const s = tl.stateAt(cp1, time)
      expectSeen(s, 6, 'frozen', 1)
      expectSeen(s, 7, 'active', 1, 0)
      expectSeen(s, 12, 'active', 0, 0)
      expectSeen(s, 8, 'none')
      expectSeen(s, 33, 'none')
    }
    const cp2 = tl.checkpoint(2)
    expectSeen(tl.stateAt(cp2, 0), 8, 'frozen', 1)
    expect(() => tl.checkpoint(3)).toThrow(RangeError)
    expect(() => tl.checkpoint(-1)).toThrow(RangeError)
    expect(() => tl.checkpoint(1.5)).toThrow(RangeError)
    // A trigger history never reaches a checkpoint.
    const live = tl.trigger(tl.next(tl.initial(), 0), '9', 100)!
    expect(ids(tl.stateAt(tl.checkpoint(live.step), 0))).not.toContain(33)
    // Leaving a checkpoint replays it at the new time: earlier steps stay settled, the next step plays.
    const c2 = tl.next(cp1, 5000)
    expect(c2).toMatchObject({ step: 2, mode: 'live' })
    const s = tl.stateAt(c2, 5100)
    expectSeen(s, 6, 'frozen', 1)
    expectSeen(s, 8, 'active', 0.1)
    expect(find(s, 7).state).toBe('active')
    // A trigger from a checkpoint keeps step 1 settled and runs the trigger.
    const tc = tl.trigger(cp1, '9', 5000)!
    expect(tc).toMatchObject({ step: 1, mode: 'live' })
    const st = tl.stateAt(tc, 5500)
    expectSeen(st, 6, 'frozen', 1)
    expectSeen(st, 33, 'active', 0.5)
  })

  it('previous across a trigger clears the trigger', () => {
    const tl = createSlideTimeline(spinDeck(), NORMAL)
    const c1 = tl.next(tl.initial(), 0)
    const tr = tl.trigger(c1, '9', 100)!
    const c2 = tl.next(tr, 300)
    expect(find(tl.stateAt(c2, 400), 33).state).toBe('active')
    // Previous snaps to the previous click checkpoint, with no trigger history.
    const back = tl.checkpoint(c2.step - 1)
    const s = tl.stateAt(back, 400)
    expectSeen(s, 33, 'none')
    expectSeen(s, 6, 'frozen', 1)
    expectSeen(s, 8, 'none')
    // Next from there replays without the trigger.
    expectSeen(tl.stateAt(tl.next(back, 1000), 1100), 33, 'none')
  })

  it('a step whose only target is unresolved still advances', async () => {
    // Step 1 targets spid 77 (absent: the reader drops the behaviour and keeps the step); step 2 targets shape 2.
    const d = await deck(timingXmlOf(mainSeqXml([xmlStep(3, 6, '77'), xmlStep(10, 13, '2')])))
    const sink = createDiagnosticSink()
    const r = readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink)
    expect(sink.list().some(x => x.code === 'unresolved-target')).toBe(true)
    const tl = createSlideTimeline(timingOf(r), NORMAL)
    expect(tl.steps).toBe(2)
    const c1 = tl.next(tl.initial(), 0)
    expect(c1.step).toBe(1)
    expect(ids(tl.stateAt(c1, 100))).toEqual([])
    const c2 = tl.next(c1, 100)
    expect(c2.step).toBe(2)
    expectSeen(tl.stateAt(c2, 350), 13, 'active', 0.5)
  })

  it('showAnimation=0 and refused timing are one static step; exit effects are not applied', () => {
    // Step 1 is an exit (hide shape 2): static never applies it.
    const exit: BehaviourNode = { ...bhv(6, { durMs: 1 }, '2'), to: 'hidden' } as BehaviourNode
    const t = tree(mainSeq([group(3, [inner(4, [par(5, [exit], { nodeType: 'clickEffect', presetClass: 'exit', start: [at(0)] })])])]), interactive(30, '9', [bhv(33, { durMs: 1000 })]))
    const timelines: SlideTimeline[] = [
      createSlideTimeline(t, { showAnimation: false, reducedMotion: false }),
      createSlideTimeline({ refused: 'limit:timeNodes' }, NORMAL),
      createSlideTimeline(undefined, NORMAL),
    ]
    for (const tl of timelines) {
      // One static step: no click steps, so the first Next leaves the slide.
      expect(tl.steps).toBe(0)
      const c = tl.initial()
      expect(c).toMatchObject({ step: 0, mode: 'static', events: [] })
      expect(tl.checkpoint(0)).toMatchObject({ step: 0, mode: 'static' })
      expect(tl.next(c, 1000)).toBe(c)
      expect(tl.triggers()).toEqual([])
      expect(tl.trigger(c, '9', 10)).toBeUndefined()
      for (const time of [0, 500, 1e6]) {
        const s = tl.stateAt(c, time)
        expect(s.active).toEqual([])
        expect(s.frozen).toEqual([])
      }
    }
  })

  it('entrance and exit on the same shape', () => {
    const entrance: BehaviourNode = bhv(6, { durMs: 1 }, '2')
    const exit: BehaviourNode = { ...bhv(8, { durMs: 1 }, '2'), to: 'hidden' } as BehaviourNode
    const tl = createSlideTimeline(tree(mainSeq([
      group(3, [inner(4, [par(5, [entrance], { nodeType: 'clickEffect', presetClass: 'entr', start: [at(0)] })])]),
      group(20, [inner(21, [par(22, [exit], { nodeType: 'clickEffect', presetClass: 'exit', start: [at(0)] })])]),
    ])), NORMAL)
    const order = (s: SchedulerSnapshot) => s.frozen.map(e => [e.node.ctn.id, e.order])
    expect(ids(tl.stateAt(tl.initial(), 0))).toEqual([])
    const s1 = tl.stateAt(tl.checkpoint(1), 0)
    expect(s1.frozen.map(e => e.node.ctn.id)).toEqual([6])
    const s2 = tl.stateAt(tl.checkpoint(2), 0)
    expect(s2.frozen.map(e => e.node.ctn.id)).toEqual([6, 8])
    const [[, o6], [, o8]] = order(s2) as Array<[number, number]>
    expect(o8).toBeGreaterThan(o6) // the exit composes over the entrance
    // Live: both frozen once step 2 has run.
    const c2 = tl.next(tl.next(tl.initial(), 0), 10)
    expect(tl.stateAt(c2, 20).frozen.map(e => e.node.ctn.id)).toEqual([6, 8])
  })

  it('reduced motion jumps each step to its settled state', () => {
    // Step 1: fade 6 (1000), afterEffect 11 delayed 1000 (500), indefinite spin 7.
    const t = tree(
      mainSeq([
        group(3, [
          inner(4, [effect(5, 'clickEffect', [bhv(6, { durMs: 1000 })]), effect(10, 'withEffect', [bhv(7, { durMs: 500, repeatCount: 'indefinite', fill: undefined })])]),
          inner(14, [effect(15, 'afterEffect', [bhv(11, { durMs: 500 })])], 1000),
        ]),
        group(20, [inner(21, [effect(22, 'clickEffect', [bhv(8, { durMs: 1000 })])])]),
      ]),
      interactive(30, '9', [bhv(33, { durMs: 1000 })]),
    )
    const tl = createSlideTimeline(t, { showAnimation: true, reducedMotion: true })
    expect(tl.steps).toBe(2)
    const c1 = tl.next(tl.initial(), 100)
    expect(c1.step).toBe(1)
    for (const time of [100, 150, 5000]) {
      const s = tl.stateAt(c1, time)
      expectSeen(s, 6, 'frozen', 1)
      expectSeen(s, 11, 'frozen', 1)
      expectSeen(s, 7, 'active', 1, 0) // no motion: held at its first iteration's end
      expectSeen(s, 8, 'none')
    }
    const c2 = tl.next(c1, 200)
    expectSeen(tl.stateAt(c2, 200), 8, 'frozen', 1)
    expect(tl.next(c2, 200)).toBe(c2)
    // A trigger's effect also lands on its end state.
    expectSeen(tl.stateAt(tl.trigger(c2, '9', 300)!, 300), 33, 'frozen', 1)
  })

  it('refuses time running backwards and non-finite times', () => {
    const tl = createSlideTimeline(spinDeck(), NORMAL)
    const c1 = tl.next(tl.initial(), 500)
    expect(() => tl.next(c1, 400)).toThrow(RangeError)
    expect(() => tl.trigger(c1, '9', 100)).toThrow(RangeError)
    expect(() => tl.next(c1, Number.NaN)).toThrow(RangeError)
    expect(() => tl.next(tl.initial(), -1)).toThrow(RangeError)
  })

  it('passes its diagnostic sink to the scheduler', () => {
    // A reference cycle: 6 begins when 7 ends, 7 begins when 6 ends.
    const sink = createDiagnosticSink()
    const t = tree(mainSeq([group(3, [inner(4, [effect(5, 'clickEffect', [
      bhv(6, { durMs: 100, start: [{ delayMs: 0, evt: 'onEnd', tn: 7 }] }),
      bhv(7, { durMs: 100, start: [{ delayMs: 0, evt: 'onEnd', tn: 6 }] }),
    ])])])]))
    const tl = createSlideTimeline(t, { ...NORMAL, sink, slideIndex: 4 })
    tl.stateAt(tl.next(tl.initial(), 0), 50)
    expect(sink.list().some(d => d.code === 'timing-cycle' && d.slideIndex === 4)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Reader -> timeline contract (S1-8 layout, S1-9 scheduler, S1-10 navigation)
// ---------------------------------------------------------------------------

/** One PowerPoint click step (ids `id`..`id+2`) whose clickEffect holds behaviour `bid` (500 ms) on `spid`. */
function xmlStep(id: number, bid: number, spid: string, body?: string): string {
  const b = body ?? `<p:set><p:cBhvr><p:cTn id="${bid}" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
  return `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${id + 2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${b}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
}
const shape = (id: number) => ({ id, prst: 'rect', off: ['0', '0'] as [string, string], ext: ['914400', '914400'] as [string, string] })
async function deck(timingXml: string) {
  const bytes = await buildPptxDeck({ slides: [{ shapes: [shape(2), shape(3)], timingXml }] })
  const zip = await JSZip.loadAsync(bytes)
  const xml = parseXmlOrdered(await zip.file('ppt/slides/slide1.xml')!.async('string'))
  const doc = await parsePptx(await OfficePackage.load(bytes))
  return { xml, slide: doc.slides[0] }
}
function timingOf(r: ReturnType<typeof readSlideTiming>): TimingTree {
  if (!r.timing || 'refused' in r.timing) throw new Error(`timing not read: ${JSON.stringify(r.timing)}`)
  return r.timing
}
function walk(n: TimeNode, f: (n: TimeNode) => void): void {
  f(n)
  if (n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl') for (const c of n.children) walk(c, f)
}

describe('reader to timeline contract', () => {
  it('reader output drives the timeline: steps, settling the last step, triggers and motion paths', async () => {
    const motion = `<p:animMotion origin="layout" path="M 0 0 L 0.5 0.25 E" pathEditMode="relative"><p:cBhvr><p:cTn id="13" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`
    const interactiveXml = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="20" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>` +
      `<p:par><p:cTn id="21" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="22" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
      `<p:set><p:cBhvr><p:cTn id="23" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
    const d = await deck(timingXmlOf(mainSeqXml([xmlStep(3, 6, '2'), xmlStep(10, 13, '2', motion)]) + interactiveXml))
    const sink = createDiagnosticSink()
    const t = timingOf(readSlideTiming(d.xml, 0, shapeResolver(d.slide), sink))
    // The default reader carries S1-6's parser: the path is parsed, not deferred.
    let path: AnimMotionBehaviour['path']
    walk(t.root, n => { if (n.kind === 'animMotion') path = n.path })
    expect(path?.segments.map(s => s.cmd)).toEqual(['M', 'L'])
    expect(sink.list().some(x => x.code === 'motion-path-deferred')).toBe(false)

    const tl = createSlideTimeline(t, NORMAL)
    expect(tl.steps).toBe(2)
    expect(tl.triggers()).toEqual(['3'])
    const init = tl.initial()
    expect(ids(tl.stateAt(init, 100))).toEqual([])
    const c1 = tl.next(init, 0)
    expectSeen(tl.stateAt(c1, 250), 6, 'active', 0.5)
    const c2 = tl.next(c1, 1000)
    expectSeen(tl.stateAt(c2, 1250), 13, 'active', 0.25)
    // Next during the last active step settles it and stays on the slide.
    const c3 = tl.next(c2, 1250)
    expect(c3).not.toBe(c2)
    expect(c3.step).toBe(2)
    expectSeen(tl.stateAt(c3, 1300), 13, 'frozen', 1)
    expect(tl.next(c3, 1300)).toBe(c3)
    // The trigger runs its interactive sequence.
    expectSeen(tl.stateAt(tl.trigger(c3, '3', 2000)!, 2500), 23, 'active', 0.5)
    // Checkpoints settle the same tree.
    expectSeen(tl.stateAt(tl.checkpoint(2), 0), 13, 'frozen', 1)
  })

  const corpus = resolvePath('corpus') + '/'
  it.skipIf(!existsSync(corpus + 'lo-tdf104788.pptx'))('tdf104788 slide 6 has six click steps', async () => {
    const bytes = new Uint8Array(readFileSync(corpus + 'lo-tdf104788.pptx'))
    const zip = await JSZip.loadAsync(bytes)
    const xml = parseXmlOrdered(await zip.file('ppt/slides/slide6.xml')!.async('string'))
    const doc = await parsePptx(await OfficePackage.load(bytes))
    const tl = createSlideTimeline(timingOf(readSlideTiming(xml, 5, shapeResolver(doc.slides[5]), createDiagnosticSink())), NORMAL)
    expect(tl.steps).toBe(6)
    let c: SlideCursor = tl.initial()
    for (let i = 1; i <= 6; i++) c = tl.next(c, i * 10_000)
    expect(c.step).toBe(6)
  })
})
