/**
 * Piece 5 (spec §5; plan D4, D22, X10): `mediaEndOf`, one case per rule — the end of an audio node (or the
 * transition sound) on the show clock, shared by the server classifier and the client driver.
 */
import { describe, expect, it } from 'vitest'
import { mediaClockOf, mediaEndOf, slideMediaEnds, TRANSITION_SOUND, type MediaClockRecord, type MediaEvent } from '../src/slideshow/engine/media-ends'
import { createSlideTimeline } from '../src/slideshow/engine/steps'
import type { MediaCommand } from '../src/slideshow/media/commands'
import type { AudioBehaviour, BehaviourNode, CTn, CmdBehaviour, ContainerNode, TimeNode } from '../src/slideshow/types'

const audio = (id: number, over: Partial<Omit<CTn, 'id'>> = {}, target: AudioBehaviour['target'] = { spid: '3' }): AudioBehaviour =>
  ({ kind: 'audio', ctn: { id, durMs: undefined, accel: 0, decel: 0, autoRev: false, fill: 'hold', start: [{ delayMs: 0 }], end: [], ...over }, target, attrNames: [], additive: 'base', accumulate: 'none' })
const rec = (durationMs: number, refused?: string): MediaClockRecord => ({ mediaId: '@a0.3', kind: 'audio', durationMs, ...(refused ? { refused } : {}) })
const cmd = (at: number, nodeId: number, c: MediaCommand): MediaEvent => ({ at, type: 'cmd', nodeId, cmd: c })

describe('mediaEndOf (spec §5)', () => {
  it('plays durationMs', () => {
    expect(mediaEndOf(audio(3), rec(4000), 100, [])).toEqual({ at: 4100 })
  })

  it('playFrom(1.5) on 4000 ms at the begin ends at 2500', () => {
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(0, 3, { verb: 'playFrom', seconds: 1.5 })])).toEqual({ at: 2500 })
    // Clamped at the duration: an offset past the clip ends the node at the command.
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'playFrom', seconds: 9 })])).toEqual({ at: 1000 })
  })

  it('playFrom(x) is an absolute seek (§10 F2): what remains is durationMs − x from the command', () => {
    // A 4 s clip, playFrom(1.5) at 1000: the player seeks to 1.5 s, so 2.5 s remain → 3500 (not 2500).
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'playFrom', seconds: 1.5 })])).toEqual({ at: 3500 })
    // playFrom(0) mid-play restarts the clip: the whole 4 s from the command.
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(2500, 3, { verb: 'playFrom', seconds: 0 })])).toEqual({ at: 6500 })
    // While paused: the seek also resumes (pause at 1000, playFrom(1) at 3000 → 3 s left → 6000).
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'pause' }), cmd(3000, 3, { verb: 'playFrom', seconds: 1 })])).toEqual({ at: 6000 })
    // Repeats: the seek is within the current iteration, as the player sets currentTime and keeps its iteration count
    // (repeatCount 2 on 4 s; playFrom(1) at 5000, in the second iteration → 3 s left → 8000).
    expect(mediaEndOf(audio(3, { repeatCount: 2 }), rec(4000), 0, [cmd(5000, 3, { verb: 'playFrom', seconds: 1 })])).toEqual({ at: 8000 })
    // In the first iteration the remaining repeats still play (playFrom(1) at 1000 → 3 s + 4 s → 8000).
    expect(mediaEndOf(audio(3, { repeatCount: 2 }), rec(4000), 0, [cmd(1000, 3, { verb: 'playFrom', seconds: 1 })])).toEqual({ at: 8000 })
    // A fractional repeat: a seek past the playback's end ends the node at the command, never before it
    // (repeatCount 2.5 on 4 s = 10 s; playFrom(3) at 9000, in the last half pass → ends at 9000).
    expect(mediaEndOf(audio(3, { repeatCount: 2.5 }), rec(4000), 0, [cmd(9000, 3, { verb: 'playFrom', seconds: 3 })])).toEqual({ at: 9000 })
  })

  it('pause 1000 ms mid-play delays the end by 1000', () => {
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'pause' }), cmd(2000, 3, { verb: 'resume' })])).toEqual({ at: 5000 })
    // A deck-command togglePause stops and restarts the clock the same way; a paused clock never ends on its own.
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'togglePause' }), cmd(2000, 3, { verb: 'togglePause' })])).toEqual({ at: 5000 })
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 3, { verb: 'pause' })])).toEqual({ open: true })
    // Another node's command changes nothing.
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1000, 4, { verb: 'pause' })])).toEqual({ at: 4000 })
  })

  it('repeatCount 3 → 3×', () => {
    expect(mediaEndOf(audio(3, { repeatCount: 3 }), rec(4000), 0, [])).toEqual({ at: 12000 })
  })

  it('indefinite / loop / untilNextSound → open', () => {
    expect(mediaEndOf(audio(3, { repeatCount: 'indefinite' }), rec(4000), 0, [])).toEqual({ open: true })
    expect(mediaEndOf(audio(TRANSITION_SOUND), rec(4000), 0, [], { transitionLoop: true })).toEqual({ open: true })
    expect(mediaEndOf(audio(90, { repeatCount: 'indefinite' }, { snd: '@a0' }), rec(4000), 0, [], { untilNextSound: true })).toEqual({ open: true })
    // Slide exit ends an open node.
    expect(mediaEndOf(audio(3, { repeatCount: 'indefinite' }), rec(4000), 0, [{ at: 7000, type: 'slide-exit' }])).toEqual({ at: 7000 })
  })

  it('dur 1500 bounds 4000 → 1500', () => {
    expect(mediaEndOf(audio(3, { durMs: 1500 }), rec(4000), 200, [])).toEqual({ at: 1700 })
    // The earlier of the two: a shorter media keeps its own end.
    expect(mediaEndOf(audio(3, { durMs: 5000 }), rec(4000), 0, [])).toEqual({ at: 4000 })
    // An open node is bounded too.
    expect(mediaEndOf(audio(3, { durMs: 1500, repeatCount: 'indefinite' }), rec(4000), 0, [])).toEqual({ at: 1500 })
  })

  it('refused → finite dur, else at once', () => {
    expect(mediaEndOf(audio(3, { durMs: 1500 }), rec(0, 'media-unavailable'), 300, [])).toEqual({ at: 1800 })
    expect(mediaEndOf(audio(3), rec(0, 'media-linked'), 300, [])).toEqual({ at: 300 })
    // D4: only a plan-refused record takes the fallback; a playable record keeps the clock end whatever happens later.
    expect(mediaEndOf(audio(3, { durMs: 1500 }), rec(4000), 300, [])).toEqual({ at: 1800 })
  })

  it('no record → at once', () => {
    expect(mediaEndOf(audio(3), undefined, 250, [])).toEqual({ at: 250 })
    expect(mediaEndOf(audio(3, { durMs: 900 }), undefined, 250, [])).toEqual({ at: 1150 })
  })

  it('stop at 1200 ends the node at 1200', () => {
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(1200, 3, { verb: 'stop' })])).toEqual({ at: 1200 })
    // While paused, too; and a stop before the node began is not its own.
    expect(mediaEndOf(audio(3), rec(4000), 0, [cmd(500, 3, { verb: 'pause' }), cmd(1200, 3, { verb: 'stop' })])).toEqual({ at: 1200 })
    expect(mediaEndOf(audio(3), rec(4000), 1000, [cmd(500, 3, { verb: 'stop' })])).toEqual({ at: 5000 })
  })

  it('onstopaudio ends every audio node and the transition sound at its time', () => {
    const stop: MediaEvent = { at: 1500, type: 'cmd', nodeId: 77, cmd: { verb: 'onstopaudio' } }
    expect(mediaEndOf(audio(3), rec(4000), 0, [stop])).toEqual({ at: 1500 })
    expect(mediaEndOf(audio(4, { repeatCount: 'indefinite' }), rec(4000), 100, [stop])).toEqual({ at: 1500 })
    expect(mediaEndOf(audio(TRANSITION_SOUND), rec(4000), 0, [stop], { transitionLoop: true })).toEqual({ at: 1500 })
  })

  it('untilNextSound ends at the next sound-start; a looping transition sound ends at the next transition sound', () => {
    const loop = audio(90, { repeatCount: 'indefinite' }, { snd: '@a0' })
    expect(mediaEndOf(loop, rec(4000), 0, [{ at: 2500, type: 'sound-start', nodeId: 91 }], { untilNextSound: true })).toEqual({ at: 2500 })
    // Its own start is not a next sound.
    expect(mediaEndOf(loop, rec(4000), 0, [{ at: 0, type: 'sound-start', nodeId: 90 }], { untilNextSound: true })).toEqual({ open: true })
    // A node that does not loop until the next sound ignores it.
    expect(mediaEndOf(audio(3), rec(4000), 0, [{ at: 2500, type: 'sound-start', nodeId: 91 }])).toEqual({ at: 4000 })
    expect(mediaEndOf(audio(TRANSITION_SOUND), rec(4000), 0, [{ at: 3000, type: 'sound-start', nodeId: TRANSITION_SOUND }], { transitionLoop: true })).toEqual({ at: 3000 })
  })
})

// ---------------------------------------------------------------------------
// The loop over a history (`slideMediaEnds().settle`), on real timelines
// ---------------------------------------------------------------------------

const ctnOf = (id: number, over: Partial<Omit<CTn, 'id'>> = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, fill: 'hold', start: [{ delayMs: 0 }], end: [], ...over })
const root = (...children: TimeNode[]): ContainerNode => ({ kind: 'par', ctn: ctnOf(1, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot', fill: undefined }), children })
/** Shape `spid` becomes visible once node `tn` ends. */
const onEndShow = (id: number, tn: number, spid = '2'): BehaviourNode =>
  ({ kind: 'set', ctn: ctnOf(id, { durMs: 1, start: [{ delayMs: 0, evt: 'onEnd', tn }] }), target: { spid }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible' })
const cmdAt = (id: number, delayMs: number, text: string, spid = '3'): CmdBehaviour =>
  ({ kind: 'cmd', cmdType: 'call', cmd: text, ctn: ctnOf(id, { durMs: 1, start: [{ delayMs }] }), target: { spid }, attrNames: [], additive: 'base', accumulate: 'none' })
const timelineOf = (r: ContainerNode) => createSlideTimeline({ root: r, build: [] }, { showAnimation: true, reducedMotion: false })
const shapeRec = (durationMs: number, spid = '3'): MediaClockRecord => ({ mediaId: `@a0.${spid}`, kind: 'audio', durationMs, spid })

describe('slideMediaEnds: the shared loop over a history', () => {
  it('an audio node ends at begin + durationMs, and what waits on its end runs then', () => {
    const r = root(audio(3, { start: [{ delayMs: 250 }] }), onEndShow(5, 3))
    const tl = timelineOf(r)
    const ends = slideMediaEnds([r], mediaClockOf([r], [shapeRec(3000)]))
    expect(ends.audio).toBe(true)
    const early = ends.settle(tl, tl.initial(), 3249)
    expect(early.ends).toEqual([])
    const done = ends.settle(tl, tl.initial(), 9000)
    expect(done.ends).toEqual([{ nodeId: 3, at: 3250 }])
    expect(done.cursor.events).toEqual([{ kind: 'mediaEnd', nodeId: 3, t: 3250, synthetic: false }])
    expect(tl.stateAt(done.cursor, 3250).active.some(e => e.node.ctn.id === 5)).toBe(true)
    expect(tl.stateAt(done.cursor, 3249).active.some(e => e.node.ctn.id === 5)).toBe(false)
  })

  it('settling in many small steps gives the history of one large step', () => {
    const r = root(audio(3), onEndShow(5, 3), { ...audio(6, { start: [{ delayMs: 0, evt: 'onEnd', tn: 3 }] }, { spid: '4' }) }, onEndShow(7, 6, '8'))
    const tl = timelineOf(r)
    const ends = slideMediaEnds([r], mediaClockOf([r], [shapeRec(1000), shapeRec(700, '4')]))
    const once = ends.settle(tl, tl.initial(), 5000).cursor
    let step = tl.initial()
    for (let t = 0; t <= 5000; t += 33) step = ends.settle(tl, step, t).cursor
    step = ends.settle(tl, step, 5000).cursor
    expect(step.events).toEqual(once.events)
    expect(once.events.map(e => [e.kind, 'nodeId' in e ? e.nodeId : -1, e.t])).toEqual([['mediaEnd', 3, 1000], ['mediaEnd', 6, 1700]])
  })

  it('every activation counts, in every iteration of a repeating parent, whenever the history is read', () => {
    const rep: ContainerNode = { kind: 'par', ctn: ctnOf(20, { durMs: 2000, repeatCount: 2 }), children: [audio(3)] }
    const r = root(rep)
    const tl = timelineOf(r)
    const ends = slideMediaEnds([r], mediaClockOf([r], [shapeRec(500)]))
    const once = ends.settle(tl, tl.initial(), 10_000)
    expect(once.ends).toEqual([{ nodeId: 3, at: 500 }, { nodeId: 3, at: 2500 }])
    let step = tl.initial()
    for (let t = 0; t <= 10_000; t += 47) step = ends.settle(tl, step, t).cursor
    expect(ends.settle(tl, step, 10_000).cursor.events).toEqual(once.cursor.events)
  })

  it('deck commands drive the clock: pause and resume, stop, and onstopaudio', () => {
    const paused = root(audio(3), cmdAt(10, 1000, 'pause'), cmdAt(11, 2000, 'resume'))
    const tlp = timelineOf(paused)
    expect(slideMediaEnds([paused], mediaClockOf([paused], [shapeRec(3000)])).settle(tlp, tlp.initial(), 9000).ends).toEqual([{ nodeId: 3, at: 4000 }])
    const stopped = root(audio(3), cmdAt(10, 1200, 'stop'))
    const tls = timelineOf(stopped)
    expect(slideMediaEnds([stopped], mediaClockOf([stopped], [shapeRec(3000)])).settle(tls, tls.initial(), 9000).ends).toEqual([{ nodeId: 3, at: 1200 }])
    const evt: CmdBehaviour = { ...cmdAt(12, 800, 'onstopaudio'), cmdType: 'evt', target: { sld: true } }
    const all = root(audio(3), audio(4, {}, { spid: '4' }), evt)
    const tla = timelineOf(all)
    expect(slideMediaEnds([all], mediaClockOf([all], [shapeRec(3000), shapeRec(5000, '4')])).settle(tla, tla.initial(), 9000).ends).toEqual([{ nodeId: 3, at: 800 }, { nodeId: 4, at: 800 }])
  })

  it('a zero-length activation (a command at its ended parent\'s last instant) fires nothing', () => {
    const box: ContainerNode = { kind: 'par', ctn: ctnOf(20, { durMs: 200 }), children: [cmdAt(10, 200, 'stop')] }
    const r = root(audio(3), box)
    const tl = timelineOf(r)
    expect(slideMediaEnds([r], mediaClockOf([r], [shapeRec(3000)])).settle(tl, tl.initial(), 9000).ends).toEqual([{ nodeId: 3, at: 3000 }])
  })

  it('a node with no record ends as it begins; a video node is never ended here (D22)', () => {
    const r = root(audio(3, { start: [{ delayMs: 400 }] }), { ...audio(9), kind: 'video' } as unknown as BehaviourNode)
    const tl = timelineOf(r)
    const ends = slideMediaEnds([r], mediaClockOf([r], []))
    expect(ends.settle(tl, tl.initial(), 9000).ends).toEqual([{ nodeId: 3, at: 400 }])
    const video = root({ ...audio(9), kind: 'video' } as unknown as BehaviourNode)
    expect(slideMediaEnds([video], mediaClockOf([video], [])).audio).toBe(false)
  })

  it('an effect sound loops until the next sound: a later effect sound ends it', () => {
    const loop = audio(90, { repeatCount: 'indefinite' }, { snd: '@a0' })
    const next = audio(91, { start: [{ delayMs: 1500 }] }, { snd: '@a1' })
    const r = root(loop, next)
    const tl = timelineOf(r)
    const records: MediaClockRecord[] = [{ mediaId: '@a0', kind: 'audio', durationMs: 1000 }, { mediaId: '@a1', kind: 'audio', durationMs: 5000 }]
    expect(slideMediaEnds([r], mediaClockOf([r], records)).settle(tl, tl.initial(), 9000).ends).toEqual([{ nodeId: 90, at: 1500 }, { nodeId: 91, at: 6500 }])
  })

  it('a checkpoint is never settled: it carries no media end, as the driver leaves it', () => {
    const r = root(audio(3), onEndShow(5, 3))
    const tl = timelineOf(r)
    const cp = tl.checkpoint(0)
    expect(slideMediaEnds([r], mediaClockOf([r], [shapeRec(10)])).settle(tl, cp, 9000).cursor).toBe(cp)
  })
})
