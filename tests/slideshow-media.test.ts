/**
 * Task S2-7: media playback and timeline integration (slideshow spec 2.1
 * goal 7; 6.1 `p:cmd`, `p:audio`, `p:video`; 5.2 one audio output; 7
 * cross-slide sound, autoplay, media end, the failed-playback escape and
 * retry; 6.5 codecs).
 *
 * Media elements are real jsdom elements whose `play`/`pause`/`load`/
 * `canPlayType` are replaced (jsdom implements none of them), created through
 * the player's `create` seam. Leases are fakes with the S2-6 contract
 * (retain / leave / release), so the timeline arithmetic is exact. Time is
 * plain slide-clock milliseconds, driven by hand.
 */
import { describe, expect, it, vi, type Mock } from 'vitest'
import { parseXmlOrdered } from '../src/core/xml'
import { createDiagnosticSink } from '../src/slideshow/diagnostics'
import { createScheduler } from '../src/slideshow/engine/scheduler'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../src/slideshow/engine/steps'
import type { MediaLease, MediaLeaseResult } from '../src/slideshow/media/lease'
import { createMediaPlayer, type MediaController, type MediaPlayer } from '../src/slideshow/media/player'
import { createTransitionSound } from '../src/slideshow/media/sounds'
import { readMedia } from '../src/slideshow/read/media'
import { readSlideTiming } from '../src/slideshow/read/timing'
import { shapeResolver } from '../src/slideshow/read/resolve'
import { readRelationships, openShowPackage } from '../src/slideshow/read/deck'
import { buildPptxDeck, timingXmlOf, mainSeqXml } from '../src/testdata/ooxml-builders'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import type {
  AudioBehaviour, BehaviourNode, CTn, CmdBehaviour, Condition, ContainerNode, DiagnosticSink, MediaRecord, ShowEvent, ShowSlide, TimeNode,
  TimingTree, VideoBehaviour,
} from '../src/slideshow/types'

// ---------------------------------------------------------------------------
// Trees
// ---------------------------------------------------------------------------

type CtnOver = Partial<Omit<CTn, 'id'>>
const ctn = (id: number, over: CtnOver = {}): CTn => ({ id, durMs: undefined, accel: 0, decel: 0, autoRev: false, start: [], end: [], ...over })
const WAIT: Condition = { delayMs: 'indefinite' }
const at = (delayMs: number): Condition => ({ delayMs })
const par = (id: number, children: TimeNode[], over: CtnOver = {}): ContainerNode => ({ kind: 'par', ctn: ctn(id, { fill: 'hold', ...over }), children })
const group = (id: number, children: TimeNode[]) => par(id, [par(id + 1, children, { start: [at(0)] })], { start: [WAIT] })
const mainSeq = (children: TimeNode[]): ContainerNode => ({
  kind: 'seq', ctn: ctn(2, { durMs: 'indefinite', nodeType: 'mainSeq' }), children, seq: { concurrent: true, nextAc: 'seek', prevAc: 'none' },
})
const tree = (...children: TimeNode[]): TimingTree => ({ root: par(1, children, { durMs: 'indefinite', restart: 'never', nodeType: 'tmRoot', fill: undefined }), build: [] })
function media(kind: 'audio' | 'video', id: number, spid: string, over: CtnOver = {}, extra: Partial<AudioBehaviour> = {}): BehaviourNode {
  return { kind, ctn: ctn(id, { fill: 'hold', start: [at(0)], ...over }), target: { spid }, attrNames: [], additive: 'base', accumulate: 'none', ...extra } as AudioBehaviour | VideoBehaviour
}
function cmd(id: number, spid: string, command: string, type: CmdBehaviour['cmdType'] = 'call'): CmdBehaviour {
  return { kind: 'cmd', cmdType: type, cmd: command, ctn: ctn(id, { durMs: 1, fill: 'hold', start: [at(0)] }), target: { spid }, attrNames: [], additive: 'base', accumulate: 'none' }
}
/** A visible-set successor on shape 99 that waits for node `tn` to end. */
const successor = (id: number, tn: number): BehaviourNode => ({
  kind: 'set', ctn: ctn(id, { durMs: 500, fill: 'hold', start: [{ delayMs: 0, evt: 'onEnd', tn }] }), target: { spid: '99' }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible',
})

const rec = (spid: string, contentType = 'video/mp4', rId = `rId${spid}`): MediaRecord => ({ rId, partPath: `ppt/media/m${spid}`, contentType, declaredBytes: 10, spid })
const slideOf = (timing: TimingTree | undefined, records: MediaRecord[], index = 0, transition?: ShowSlide['transition']): ShowSlide =>
  ({ index, hidden: false, timing, notes: [], links: [], media: records, builds: [], ...(transition ? { transition } : {}) })

// ---------------------------------------------------------------------------
// Media elements, leases and the player
// ---------------------------------------------------------------------------

type Fake = HTMLMediaElement & { play: Mock<() => Promise<void>>; pause: Mock<() => void>; load: Mock<() => void>; canPlayType: Mock<(t: string) => CanPlayTypeResult> }

interface Rig {
  player: MediaPlayer
  root: HTMLElement
  sink: DiagnosticSink
  elements: Fake[]
  leases: Array<FakeLease>
  opens: MediaRecord[]
  /** What the next `play()` does. */
  playWith: { result: () => Promise<void> }
  canPlay: { value: CanPlayTypeResult }
  openWith: { result: (r: MediaRecord) => MediaLeaseResult }
  now: { t: number }
  /** `stopTransitionSound` calls (S2-10): true = only a looping one. */
  transitionStops: boolean[]
}

interface FakeLease extends MediaLease { released: boolean; leaves: number; rec: MediaRecord }
function fakeLease(r: MediaRecord): FakeLease {
  let spans = 0
  const l: FakeLease = {
    url: `blob:${r.partPath}`, bytes: r.declaredBytes, rec: r, released: false, leaves: 0,
    get spans() { return spans },
    retain(n) { spans = n === Infinity ? Infinity : Math.max(spans, n) },
    leave() { l.leaves++; if (l.released || spans === Infinity) return; if (spans > 0) spans--; if (spans === 0) l.release() },
    release() { l.released = true },
  }
  return l
}

function rig(doc: Document = document): Rig {
  const root = doc.createElement('div')
  doc.body.append(root)
  const sink = createDiagnosticSink()
  const elements: Fake[] = []
  const leases: FakeLease[] = []
  const opens: MediaRecord[] = []
  const playWith = { result: () => Promise.resolve() }
  const canPlay = { value: 'maybe' as CanPlayTypeResult }
  const openWith = { result: (r: MediaRecord): MediaLeaseResult => { const l = fakeLease(r); leases.push(l); return l } }
  const now = { t: 0 }
  const transitionStops: boolean[] = []
  const player = createMediaPlayer({
    root,
    sink,
    stopTransitionSound: (onlyLooping) => { transitionStops.push(onlyLooping); return false },
    open: async (r) => { opens.push(r); return openWith.result(r) },
    create: (kind) => {
      const el = doc.createElement(kind) as Fake
      el.play = vi.fn(() => playWith.result())
      el.pause = vi.fn()
      el.load = vi.fn()
      el.canPlayType = vi.fn(() => canPlay.value)
      elements.push(el)
      return el
    },
  })
  return { player, root, sink, elements, leases, opens, playWith, canPlay, openWith, now, transitionStops }
}

const flush = async () => { for (let i = 0; i < 10; i++) await Promise.resolve() }
const codes = (sink: DiagnosticSink) => sink.list().map((d) => d.code)

/** Drives one slide: the timeline, its cursor and the controller, with media ends fed back as the session does. */
function drive(r: Rig, slide: ShowSlide, reducedMotion = false) {
  const tl: SlideTimeline = createSlideTimeline(slide.timing, { showAnimation: true, reducedMotion })
  const ctl: MediaController = r.player.enter(slide, () => r.now.t)
  const s = { tl, ctl, cursor: tl.initial() as SlideCursor, ends: [] as Array<[number, boolean, number]> }
  ctl.onEnd((id, synthetic, t) => {
    s.ends.push([id, synthetic, t])
    s.cursor = tl.mediaEnd(s.cursor, id, t, synthetic)
  })
  const sync = (t: number) => { r.now.t = t; ctl.sync(tl.stateAt(s.cursor, t), t) }
  const next = (t: number) => { r.now.t = t; s.cursor = tl.next(s.cursor, t); sync(t) }
  const state = (id: number, t: number) => {
    const snap = tl.stateAt(s.cursor, t)
    if (snap.active.some((e) => e.node.ctn.id === id)) return 'active'
    if (snap.frozen.some((e) => e.node.ctn.id === id)) return 'frozen'
    return snap.ended.has(id) ? 'ended' : 'none'
  }
  return Object.assign(s, { sync, next, state })
}

// ---------------------------------------------------------------------------
// The brief's tests
// ---------------------------------------------------------------------------

describe('media playback and timeline integration (S2-7)', () => {
  it('a successful clip with no dur holds its successor until the clip ends', async () => {
    const r = rig()
    const timing = tree(mainSeq([group(3, [media('video', 10, '7', { start: [at(0)] }), successor(11, 10)])]))
    const d = drive(r, slideOf(timing, [rec('7')]))
    d.next(0)
    await flush()
    expect(r.elements).toHaveLength(1)
    expect(r.elements[0].play).toHaveBeenCalledTimes(1)
    expect(r.elements[0].src).toBe('blob:ppt/media/m7')
    d.sync(1000)
    expect(d.state(11, 1000)).toBe('none')
    d.sync(2999)
    expect(d.state(10, 2999)).toBe('active')
    expect(d.ends).toEqual([])
    // The clip ends on its own at 3000 ms of slide time.
    r.now.t = 3000
    r.elements[0].dispatchEvent(new Event('ended'))
    expect(d.ends).toEqual([[10, false, 3000]])
    expect(d.state(10, 3100)).toBe('frozen')
    expect(d.state(11, 3100)).toBe('active')
    // Nothing failed; the overlay approximation is the only note.
    expect(codes(r.sink)).toEqual(['media-overlay'])
  })

  it('refused playback with no dur and refused indefinite media release successors at once, diagnosed', async () => {
    // (a) the browser refuses play() (autoplay policy); the node has no dur.
    let r = rig()
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    let d = drive(r, slideOf(tree(mainSeq([group(3, [media('audio', 10, '7'), successor(11, 10)])])), [rec('7', 'audio/wav')]))
    d.next(0)
    r.now.t = 40
    await flush()
    expect(d.ends).toEqual([[10, true, 40]])
    expect(d.state(11, 40)).toBe('active')
    expect(codes(r.sink)).toContain('media-autoplay')
    // (b) the media lease is refused; the node's dur is indefinite. The failure is found at the node's own begin.
    r = rig()
    r.openWith.result = () => ({ refused: 'media-limit' })
    d = drive(r, slideOf(tree(mainSeq([group(3, [media('video', 10, '7', { durMs: 'indefinite' }), successor(11, 10)])])), [rec('7')]))
    d.next(0)
    await flush()
    expect(d.ends).toEqual([[10, true, 0]])
    expect(d.state(10, 0)).not.toBe('active')
    expect(d.state(11, 0)).toBe('active')
    expect(codes(r.sink)).toContain('media-refused')
    expect(r.elements.every((e) => !e.isConnected)).toBe(true)
  })

  it('refused with dur 2000 releases at 2000 ms', async () => {
    const r = rig()
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    const d = drive(r, slideOf(tree(mainSeq([group(3, [media('audio', 10, '7', { durMs: 2000 }), successor(11, 10)])])), [rec('7', 'audio/wav')]))
    d.next(0)
    r.now.t = 100
    await flush()
    expect(d.ends).toEqual([])
    d.sync(1999)
    expect(d.ends).toEqual([])
    expect(d.state(11, 1999)).toBe('none')
    d.sync(2016)
    expect(d.ends).toEqual([[10, true, 2000]])
    expect(d.state(11, 2016)).toBe('active')
    d.sync(3000)
    expect(d.ends).toHaveLength(1)
  })

  it('a later retry does not fire successors twice', async () => {
    const r = rig()
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    const d = drive(r, slideOf(tree(mainSeq([group(3, [media('audio', 10, '7'), successor(11, 10)])])), [rec('7', 'audio/wav')]))
    d.next(0)
    r.now.t = 10
    await flush()
    expect(d.ends).toEqual([[10, true, 10]])
    r.playWith.result = () => Promise.resolve()
    r.now.t = 500
    d.ctl.retry(10)
    await flush()
    expect(r.elements[0].play).toHaveBeenCalledTimes(2)
    // The retried clip plays to its end; the end is not fed to the timeline again.
    r.now.t = 4000
    r.elements[0].dispatchEvent(new Event('ended'))
    d.sync(4000)
    expect(d.ends).toEqual([[10, true, 10]])
    expect(d.cursor.events.filter((e: ShowEvent) => e.kind === 'mediaEnd')).toHaveLength(1)
  })

  it('play rejection shows Click to play; the click calls play in its own activation', async () => {
    const r = rig()
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    const surfaceClicks = vi.fn()
    r.root.addEventListener('click', surfaceClicks)
    const d = drive(r, slideOf(tree(mainSeq([group(3, [media('video', 10, '7')])])), [rec('7')]))
    d.next(0)
    await flush()
    const button = r.root.querySelector<HTMLButtonElement>('button[data-officeview-show-play]')
    expect(button?.textContent).toBe('Click to play')
    expect(button?.ownerDocument).toBe(r.root.ownerDocument)
    const el = r.elements[0]
    r.playWith.result = () => Promise.resolve()
    let during = -1
    button!.addEventListener('click', () => { during = el.play.mock.calls.length })
    const before = el.play.mock.calls.length
    button!.click()
    // play() ran inside the click's own dispatch (its transient activation), before any await.
    expect(el.play.mock.calls.length).toBe(before + 1)
    expect(during).toBe(before + 1)
    // The click is the media's, never a slide advance.
    expect(surfaceClicks).not.toHaveBeenCalled()
    await flush()
    expect(r.root.querySelector('button[data-officeview-show-play]')).toBeNull()
  })

  it('cmd play, pause and stop follow the timeline', async () => {
    const r = rig()
    const timing = tree(
      mainSeq([group(3, [cmd(20, '7', 'playFrom(1.5)')]), group(5, [cmd(21, '7', 'pause')]), group(7, [cmd(22, '7', 'togglePause')]), group(9, [cmd(23, '7', 'stop')])]),
      media('audio', 10, '7', { start: [WAIT] }),
    )
    const d = drive(r, slideOf(timing, [rec('7', 'audio/wav')]))
    d.sync(0)
    expect(r.elements).toHaveLength(0)
    // The 1 ms command has already ended when the next frame samples it: it still runs, once.
    r.now.t = 100
    d.cursor = d.tl.next(d.cursor, 100)
    d.sync(116)
    await flush()
    const el = r.elements[0]
    expect(el.play).toHaveBeenCalledTimes(1)
    expect(el.currentTime).toBe(1.5)
    d.sync(200)
    expect(el.play).toHaveBeenCalledTimes(1)
    d.next(300)
    expect(el.pause).toHaveBeenCalledTimes(1)
    d.next(400)
    expect(el.play).toHaveBeenCalledTimes(2)
    d.next(500)
    expect(el.pause).toHaveBeenCalledTimes(2)
    expect(el.currentTime).toBe(0)
    expect(d.ends).toEqual([[10, false, 500]])
    expect(codes(r.sink)).toEqual([])
    // An unknown command is diagnosed, never silently dropped.
    const r2 = rig()
    const d2 = drive(r2, slideOf(tree(mainSeq([group(3, [cmd(20, '7', 'seek(3)')])]), media('audio', 10, '7', { start: [WAIT] })), [rec('7', 'audio/wav')]))
    d2.next(0)
    expect(codes(r2.sink)).toContain('media-command')
  })

  it('a looping transition sound continues to the next slide sound or endSnd', async () => {
    const r = rig()
    const sound = (rId: string, loop: boolean) => ({ element: 'fade', ns: 'p' as const, attrs: {}, durationMs: 0, advClick: true, sound: { start: { rId, loop } } })
    const slides = [
      slideOf(undefined, [rec('a', 'audio/wav', 'rIdS1')], 0, sound('rIdS1', true)),
      slideOf(undefined, [], 1),
      slideOf(undefined, [], 2, { element: 'cut', ns: 'p', attrs: {}, durationMs: 0, advClick: true, sound: { end: true } }),
      slideOf(undefined, [rec('b', 'audio/wav', 'rIdS2')], 3, sound('rIdS2', false)),
      slideOf(undefined, [rec('c', 'audio/wav', 'rIdS3')], 4, sound('rIdS3', true)),
      slideOf(undefined, [rec('d', 'audio/wav', 'rIdS4')], 5, sound('rIdS4', false)),
    ]
    const sounds = createTransitionSound({ root: r.root, sink: r.sink, slide: (i) => slides[i], open: async (m) => { r.opens.push(m); return r.openWith.result(m) }, create: (k) => {
      const el = r.root.ownerDocument.createElement(k) as Fake
      el.play = vi.fn(() => Promise.resolve()); el.pause = vi.fn(); el.load = vi.fn(); el.canPlayType = vi.fn(() => 'maybe' as CanPlayTypeResult)
      r.elements.push(el)
      return el
    } })
    sounds.start(0)
    await flush()
    const looping = r.elements[0]
    expect(looping.loop).toBe(true)
    expect(looping.play).toHaveBeenCalledTimes(1)
    sounds.stopAt(0)
    sounds.start(1)
    await flush()
    // Slide 1 has no sound: the loop plays on.
    expect(looping.pause).not.toHaveBeenCalled()
    expect(r.leases[0].released).toBe(false)
    sounds.stopAt(1)
    sounds.start(2)
    // p:endSnd stops it: the element lets go of the URL before the lease is released.
    expect(looping.pause).toHaveBeenCalled()
    expect(looping.getAttribute('src')).toBeNull()
    expect(looping.load).toHaveBeenCalled()
    expect(r.leases[0].released).toBe(true)
    // A sound that does not loop stops when its slide is left.
    sounds.stopAt(2)
    sounds.start(3)
    await flush()
    const once = r.elements[1]
    expect(once.loop).toBe(false)
    sounds.stopAt(3)
    expect(once.pause).toHaveBeenCalled()
    expect(r.leases[1].released).toBe(true)
    // A looping sound is replaced by the next slide's own sound.
    sounds.start(4)
    await flush()
    const second = r.elements[2]
    sounds.stopAt(4)
    expect(second.pause).not.toHaveBeenCalled()
    sounds.start(5)
    expect(second.pause).toHaveBeenCalled()
    expect(r.leases[2].released).toBe(true)
    await flush()
    expect(r.elements[3].play).toHaveBeenCalledTimes(1)
    sounds.dispose()
    expect(r.leases.every((l) => l.released)).toBe(true)
  })

  it('numSld media spans slides; everything else stops on slide change', async () => {
    const r = rig()
    const timing = tree(mainSeq([]), media('audio', 10, '7', {}, { numSld: 2 }), media('audio', 11, '8'))
    const a = drive(r, slideOf(timing, [rec('7', 'audio/wav'), rec('8', 'audio/wav')]))
    a.sync(0)
    await flush()
    const [spanning, local] = r.elements
    expect(spanning.play).toHaveBeenCalledTimes(1)
    expect(local.play).toHaveBeenCalledTimes(1)
    a.ctl.leave()
    expect(local.pause).toHaveBeenCalled()
    expect(local.getAttribute('src')).toBeNull()
    expect(r.leases[1].released).toBe(true)
    expect(local.isConnected).toBe(false)
    expect(spanning.pause).not.toHaveBeenCalled()
    expect(r.leases[0].released).toBe(false)
    expect(spanning.isConnected).toBe(true)
    // Its end no longer reaches the left slide's timeline.
    spanning.dispatchEvent(new Event('ended'))
    expect(a.ends).toEqual([])
    const b = drive(r, slideOf(undefined, [], 1))
    expect(spanning.pause).not.toHaveBeenCalled()
    b.ctl.leave()
    expect(spanning.pause).toHaveBeenCalled()
    expect(r.leases[0].released).toBe(true)
    expect(spanning.isConnected).toBe(false)
  })

  it('undecodable avi/wmv keeps the poster with a diagnostic', async () => {
    for (const type of ['video/x-ms-wmv', 'video/avi']) {
      const r = rig()
      r.canPlay.value = ''
      const d = drive(r, slideOf(tree(mainSeq([group(3, [media('video', 10, '7'), successor(11, 10)])])), [rec('7', type)]))
      d.next(0)
      await flush()
      // No bytes are read or charged, nothing covers the poster, and the timeline moves on.
      expect(r.opens).toEqual([])
      expect(r.root.querySelector('video, audio')).toBeNull()
      expect(r.root.querySelector('button')).toBeNull()
      expect(d.ends).toEqual([[10, true, 0]])
      expect(d.state(11, 0)).toBe('active')
      expect(r.sink.list()).toContainEqual(expect.objectContaining({ code: 'media-codec', slideIndex: 0 }))
      // Play commands are no-ops.
      d.ctl.play(10)
      await flush()
      expect(r.opens).toEqual([])
    }
  })

  it('only the audience element has audio', async () => {
    const presenter = document.implementation.createHTMLDocument('presenter')
    const r = rig()
    const d = drive(r, slideOf(tree(mainSeq([group(3, [media('video', 10, '7')])])), [rec('7')]))
    d.next(0)
    await flush()
    expect(r.root.querySelectorAll('video, audio')).toHaveLength(1)
    expect(presenter.querySelectorAll('video, audio')).toHaveLength(0)
    const el = r.elements[0]
    expect(el.ownerDocument).toBe(r.root.ownerDocument)
    expect(el.muted).toBe(false)
    // The presenter mirror copies frames from the one audience element: no second media element, no second output.
    const ctx = { drawImage: vi.fn() } as unknown as CanvasRenderingContext2D
    expect(r.player.drawFrame('7', ctx, 0, 0, 160, 90)).toBe(true)
    expect((ctx.drawImage as Mock).mock.calls[0][0]).toBe(el)
    expect(r.elements).toHaveLength(1)
    expect(r.player.drawFrame('8', ctx, 0, 0, 160, 90)).toBe(false)
  })
})

// ---------------------------------------------------------------------------
// Supporting rules
// ---------------------------------------------------------------------------

describe('media supporting rules (S2-7)', () => {
  it('the timeline records media ends as events; a static slide ignores them', () => {
    const timing = tree(mainSeq([group(3, [media('audio', 10, '7'), successor(11, 10)])]))
    const tl = createSlideTimeline(timing, { showAnimation: true, reducedMotion: false })
    const c1 = tl.next(tl.initial(), 0)
    const c2 = tl.mediaEnd(c1, 10, 250, true)
    expect(c2).not.toBe(c1)
    expect(c2.events.at(-1)).toEqual({ kind: 'mediaEnd', nodeId: 10, t: 250, synthetic: true })
    expect(c2.step).toBe(c1.step)
    expect(() => tl.mediaEnd(c2, 10, 100, false)).toThrow(RangeError)
    const still = createSlideTimeline(timing, { showAnimation: false, reducedMotion: false })
    expect(still.mediaEnd(still.initial(), 10, 0, true)).toBe(still.initial())
  })

  it('a media end at the node\'s own begin instant ends it there (the immediate escape)', () => {
    const audio = media('audio', 10, '7', { fill: 'remove' })
    const sch = createScheduler(tree(audio, successor(11, 10)))
    const ev: ShowEvent[] = [{ kind: 'mediaEnd', nodeId: 10, t: 0, synthetic: true }]
    expect(sch.endTimeOf(10, ev)).toBe(0)
    const s = sch.evaluate(ev, 100)
    expect(s.ended.has(10)).toBe(true)
    expect(s.active.map((e) => e.node.ctn.id)).toEqual([11])
  })

  it('a refused lease releases cross-slide media first, then retries once', async () => {
    const r = rig()
    const a = drive(r, slideOf(tree(mainSeq([]), media('audio', 10, '7', {}, { numSld: 3 })), [rec('7', 'audio/wav')]))
    a.sync(0)
    await flush()
    a.ctl.leave()
    expect(r.leases[0].released).toBe(false)
    let refusals = 1
    const ok = r.openWith.result
    r.openWith.result = (m) => (refusals-- > 0 ? { refused: 'resource-limit' } : ok(m))
    const b = drive(r, slideOf(tree(mainSeq([]), media('video', 20, '9')), [rec('9')], 1))
    b.sync(0)
    await flush()
    expect(r.leases[0].released).toBe(true)
    expect(r.elements[0].isConnected).toBe(false)
    expect(r.opens.map((m) => m.spid)).toEqual(['7', '9', '9'])
    expect(r.elements[1].play).toHaveBeenCalledTimes(1)
    expect(b.ends).toEqual([])
  })

  it('reduced motion still plays media whose step is entered settled', async () => {
    const r = rig()
    const d = drive(r, slideOf(tree(mainSeq([group(3, [media('audio', 10, '7', { durMs: 4000 })])])), [rec('7', 'audio/wav')]), true)
    d.next(0)
    expect(d.state(10, 0)).toBe('frozen')
    await flush()
    expect(r.elements[0].play).toHaveBeenCalledTimes(1)
    d.sync(100)
    expect(r.elements[0].pause).not.toHaveBeenCalled()
  })

  it('pausing the show pauses playing media and resumes it', async () => {
    const r = rig()
    const d = drive(r, slideOf(tree(mainSeq([]), media('audio', 10, '7')), [rec('7', 'audio/wav')]))
    d.sync(0)
    await flush()
    const el = r.elements[0]
    r.player.setPaused(true)
    expect(el.pause).toHaveBeenCalledTimes(1)
    r.player.setPaused(false)
    expect(el.play).toHaveBeenCalledTimes(2)
    // A resume the browser refuses (a hidden tab made visible, no fresh gesture) is diagnosed and offers Click to play.
    r.player.setPaused(true)
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    r.player.setPaused(false)
    await flush()
    expect(codes(r.sink)).toContain('media-autoplay')
    expect(r.root.querySelector('button[data-officeview-show-play]')?.textContent).toBe('Click to play')
  })

  it('a node that ends by its own timing stops its element; a slide left mid-read releases the late lease', async () => {
    const r = rig()
    const d = drive(r, slideOf(tree(mainSeq([]), media('audio', 10, '7', { durMs: 1000 })), [rec('7', 'audio/wav')]))
    d.sync(0)
    await flush()
    d.sync(1016)
    expect(r.elements[0].pause).toHaveBeenCalled()
    expect(d.ends).toEqual([])
    // A read still in flight when the slide is left is released on arrival, unplayed.
    const r2 = rig()
    let resolve!: (v: MediaLeaseResult) => void
    const late = fakeLease(rec('7', 'audio/wav'))
    r2.openWith.result = () => late
    const probe = Object.assign(document.createElement('audio'), { canPlayType: () => 'maybe', play: vi.fn(() => Promise.resolve()), pause: vi.fn(), load: vi.fn() })
    const player = createMediaPlayer({ root: r2.root, sink: r2.sink, open: () => new Promise((res) => { resolve = res }), create: () => probe })
    const ctl = player.enter(slideOf(tree(mainSeq([]), media('audio', 10, '7')), [rec('7', 'audio/wav')]), () => 0)
    ctl.sync(createSlideTimeline(tree(mainSeq([]), media('audio', 10, '7')), { showAnimation: true, reducedMotion: false }).stateAt({ step: 0, events: [], mode: 'initial' }, 0), 0)
    ctl.leave()
    resolve(late)
    await flush()
    expect(late.released).toBe(true)
    expect(probe.play).not.toHaveBeenCalled()
    expect(probe.isConnected).toBe(false)
  })

  it('the reader carries numSld, vol and mute; media records know their shape', async () => {
    const xml = (inner: string) => `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp></p:spTree></p:cSld>${timingXmlOf(inner)}</p:sld>`
    const sink = createDiagnosticSink()
    const raw = await buildPptxDeck({ slides: [{ shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }] }] })
    const doc = await parsePptx(await OfficePackage.load(raw))
    const r = readSlideTiming(parseXmlOrdered(xml(`<p:audio><p:cMediaNode vol="50000" mute="1" numSld="3"><p:cTn id="6" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cMediaNode></p:audio>`)), 0, shapeResolver(doc.slides[0], []), sink)
    const root = (r.timing as TimingTree).root as ContainerNode
    expect(root.children[0]).toMatchObject({ kind: 'audio', numSld: 3, volume: 0.5, muted: true })
    expect(sink.list().filter((d) => d.code === 'timing-ignored').map((d) => d.effect)).not.toContain('numSld')
    // MediaRecord.spid from a:videoFile / a:audioFile (r:link) and p14:media (r:embed).
    const deck = await buildPptxDeck({ slides: [{
      shapes: [{ id: 2, off: ['0', '0'], ext: ['1000', '1000'] }, { id: 3, off: ['0', '0'], ext: ['1000', '1000'] }],
      media: [{ shapeId: 2, kind: 'video', ext: 'mp4', data: new Uint8Array([1]), p14: true }, { shapeId: 3, kind: 'audio', ext: 'wav', data: new Uint8Array([2]) }],
      timingXml: timingXmlOf(mainSeqXml([])),
    }] })
    const handle = await openShowPackage(deck)
    const part = handle.slideParts[0]
    const rels = await readRelationships(handle, part)
    const pkg = await OfficePackage.load(deck)
    const slideXml = parseXmlOrdered((await pkg.text(part))!)
    const types = parseXmlOrdered((await pkg.text('[Content_Types].xml'))!)
    const records = readMedia(rels, types, handle.pkg, 0, createDiagnosticSink(), slideXml)
    expect(records.map((m) => [m.spid, m.contentType])).toEqual(expect.arrayContaining([['2', 'video/mp4'], ['3', 'audio/wav']]))
    expect(records.every((m) => m.spid !== undefined)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Task S2-10: effect sounds (p:audio on a p:sndTgt, in an effect's subTnLst)
// ---------------------------------------------------------------------------

/** An effect sound as PowerPoint writes it: `p:audio` on a `p:sndTgt`, begun with its effect (`evt="begin"` on the effect's node). */
function effectSound(id: number, snd: string, effectId: number, over: CtnOver = {}): AudioBehaviour {
  return { kind: 'audio', ctn: ctn(id, { start: [{ delayMs: 0, evt: 'begin', tn: effectId }], ...over }), target: { snd }, attrNames: [], additive: 'base', accumulate: 'none' } as AudioBehaviour
}
/** The fly-in's visible set on shape 4 (the effect the sound belongs to). */
const flyIn = (id: number): BehaviourNode => ({
  kind: 'set', ctn: ctn(id, { durMs: 1, fill: 'hold', start: [at(0)] }), target: { spid: '4' }, attrNames: ['style.visibility'], additive: 'base', accumulate: 'none', to: 'visible',
})
/** The sound's media record: a relationship of the slide, on no shape. */
const soundRec = (rId = 'rId2', contentType = 'audio/wav'): MediaRecord => ({ rId, partPath: `ppt/media/${rId}.wav`, contentType, declaredBytes: 10 })

describe('effect sounds (S2-10)', () => {
  it('an effect with a sndTgt plays its sound at the effect\'s begin', async () => {
    const r = rig()
    // group(3) is par 3 (the click) > par 4 (the effect); the sound begins with node 4 (tdf44223's `evt="begin" tn=5`).
    const timing = tree(mainSeq([group(3, [flyIn(10), effectSound(-1, 'rId2', 4), successor(11, -1)])]))
    const d = drive(r, slideOf(timing, [soundRec()]))
    d.sync(0)
    // Nothing plays before the effect's click.
    expect(r.elements).toHaveLength(0)
    expect(d.ctl.wants).toBe(true)
    d.next(100)
    await flush()
    expect(r.opens.map((m) => m.rId)).toEqual(['rId2'])
    expect(r.elements).toHaveLength(1)
    const el = r.elements[0]
    expect(el.tagName).toBe('AUDIO')
    expect(el.play).toHaveBeenCalledTimes(1)
    expect(el.src).toBe('blob:ppt/media/rId2.wav')
    expect(el.hidden).toBe(true)
    expect(el.ownerDocument).toBe(r.root.ownerDocument)
    expect(d.state(-1, 200)).toBe('active')
    // Its real end releases whatever waits on it.
    r.now.t = 1500
    el.dispatchEvent(new Event('ended'))
    expect(d.ends).toEqual([[-1, false, 1500]])
    expect(d.state(11, 1600)).toBe('active')
    expect(codes(r.sink)).toEqual([])
    // Everything but numSld media stops on slide change: the element lets go of its URL, then the lease goes.
    d.ctl.leave()
    expect(el.getAttribute('src')).toBeNull()
    expect(r.leases[0].released).toBe(true)
  })

  it('a refused effect sound is diagnosed, offers no Click to play, and the show continues', async () => {
    const r = rig()
    r.playWith.result = () => Promise.reject(new DOMException('no', 'NotAllowedError'))
    const timing = tree(mainSeq([group(3, [flyIn(10), effectSound(-1, 'rId2', 4), successor(11, -1)])]))
    const d = drive(r, slideOf(timing, [soundRec()]))
    d.next(0)
    r.now.t = 30
    await flush()
    expect(codes(r.sink)).toEqual(['media-autoplay'])
    expect(r.sink.list()[0].effect).toBe('behaviour -1')
    // There is no poster to click: the transition-sound convention.
    expect(r.root.querySelector('[data-officeview-show-play]')).toBeNull()
    expect(d.ends).toEqual([[-1, true, 30]])
    expect(d.state(11, 30)).toBe('active')
    // An undecodable sound reads nothing.
    const r2 = rig()
    r2.canPlay.value = ''
    const d2 = drive(r2, slideOf(timing, [soundRec('rId2', 'audio/x-ms-wma')]))
    d2.next(0)
    await flush()
    expect(r2.opens).toEqual([])
    expect(codes(r2.sink)).toEqual(['media-codec'])
    expect(d2.ends).toEqual([[-1, true, 0]])
  })

  it('a looping effect sound loops until a stop-previous-sound command', async () => {
    const r = rig()
    // "Loop until next sound" is repeatCount="indefinite"; "Stop previous sound" is p:cmd type="evt" cmd="onstopaudio".
    const stopSound: CmdBehaviour = { ...cmd(30, '0', 'onstopaudio', 'evt'), target: { sld: true } }
    const timing = tree(mainSeq([
      group(3, [flyIn(10), effectSound(-1, 'rId2', 4, { repeatCount: 'indefinite' })]),
      group(5, [flyIn(20), stopSound]),
    ]))
    const d = drive(r, slideOf(timing, [soundRec()]))
    d.next(0)
    await flush()
    const el = r.elements[0]
    expect(el.loop).toBe(true)
    expect(el.play).toHaveBeenCalledTimes(1)
    d.sync(5000)
    expect(el.pause).not.toHaveBeenCalled()
    d.next(6000)
    expect(el.pause).toHaveBeenCalled()
    expect(d.ends).toEqual([[-1, false, 6000]])
    // The effect sound stopped only a looping transition sound; "Stop previous sound" stops any.
    expect(r.transitionStops).toEqual([true, false])
    expect(codes(r.sink)).toEqual([])
  })

  it('the next effect sound stops a looping one ("loop until next sound")', async () => {
    const r = rig()
    const timing = tree(mainSeq([
      group(3, [flyIn(10), effectSound(-1, 'rId2', 4, { repeatCount: 'indefinite' })]),
      group(5, [flyIn(20), effectSound(-2, 'rId3', 6)]),
    ]))
    const d = drive(r, slideOf(timing, [soundRec('rId2'), soundRec('rId3')]))
    d.next(0)
    await flush()
    const looping = r.elements[0]
    d.next(4000)
    await flush()
    expect(looping.pause).toHaveBeenCalled()
    expect(d.ends).toEqual([[-1, false, 4000]])
    expect(r.elements).toHaveLength(2)
    expect(r.elements[1].play).toHaveBeenCalledTimes(1)
    expect(r.elements[1].loop).toBe(false)
  })

  it('an effect sound stops a looping transition sound only; Stop previous sound stops any', async () => {
    const r = rig()
    const tr = (rId: string, loop: boolean) => ({ element: 'fade', ns: 'p' as const, attrs: {}, durationMs: 0, advClick: true, sound: { start: { rId, loop } } })
    const slides = [slideOf(undefined, [rec('a', 'audio/wav', 'rIdS1')], 0, tr('rIdS1', false)), slideOf(undefined, [rec('b', 'audio/wav', 'rIdS2')], 1, tr('rIdS2', true))]
    const sounds = createTransitionSound({ root: r.root, sink: r.sink, slide: (i) => slides[i], open: async (m) => r.openWith.result(m), create: (k) => {
      const el = r.root.ownerDocument.createElement(k) as Fake
      el.play = vi.fn(() => Promise.resolve()); el.pause = vi.fn(); el.load = vi.fn(); el.canPlayType = vi.fn(() => 'maybe' as CanPlayTypeResult)
      r.elements.push(el)
      return el
    } })
    sounds.start(0)
    await flush()
    // A sound that does not loop is left alone by the next effect sound...
    expect(sounds.release(true)).toBe(false)
    expect(r.leases[0].released).toBe(false)
    // ...and stopped by "Stop previous sound".
    expect(sounds.release(false)).toBe(true)
    expect(r.leases[0].released).toBe(true)
    sounds.start(1)
    await flush()
    expect(sounds.release(true)).toBe(true)
    expect(r.elements[1].pause).toHaveBeenCalled()
    expect(r.leases[1].released).toBe(true)
  })

  it('a built-in sound or a sound with no part is diagnosed, never silent', async () => {
    const r = rig()
    const timing = tree(mainSeq([group(3, [flyIn(10), effectSound(-1, 'builtIn:wind.wav', 4), successor(11, -1)])]))
    const d = drive(r, slideOf(timing, []))
    d.next(0)
    await flush()
    expect(r.elements).toHaveLength(0)
    expect(codes(r.sink)).toEqual(['media-missing'])
    expect(r.sink.list()[0].message).toContain('wind.wav')
    expect(d.ends).toEqual([[-1, true, 0]])
    expect(d.state(11, 0)).toBe('active')
  })
})
