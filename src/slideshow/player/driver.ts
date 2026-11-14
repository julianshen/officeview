/**
 * The show driver (streaming Stage 3 piece 4, Task 3): the navigation the
 * local session and the streamed show share — entering slides, Next,
 * Previous and goTo through the visible order, clicks (triggers, links,
 * `advClick`), the after-timer auto-advance, the kiosk restart and media
 * ends — over a narrow slide source and stage.
 *
 * Package-free (D25): every import from the player, the reader or the
 * renderer is `import type`; the only value imports are the engine's pure
 * timeline modules and the diagnostic-free order helpers.
 *
 * Navigation is synchronous where the host can observe it: `slide` and the
 * stage's leave change at once. A slide's model (its click steps) arrives
 * asynchronously; step inputs typed meanwhile are queued for that slide and
 * replayed in order, and a stale load is released unused. Explicit
 * navigation (goTo, first, last) supersedes the queue.
 *
 * On every slide change the stage is told `flattenCurrent()` (the outgoing
 * frame is captured) and then `leave()` (the rest is released), in that
 * order, before the next slide's `enter`.
 */
import { createAutoAdvance, createKioskTimer, type AutoAdvance, type KioskTimer } from '../engine/advance'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../engine/steps'
import type { SchedulerSnapshot } from '../engine/scheduler'
import type { DiagnosticSink, LinkSpec, RenderState, ShowSettings, ShowSlide, TimeNode } from '../types'
import { nextVisible, previousVisible } from './order'
import type { SessionClock } from './session'
import type { Stage, StageModel } from './stage'

/** Where the driver reads slides from: the local package, or a streamed plan. */
export interface SlideSource {
  readonly count: number
  /**
   * The deck's show settings the driver plays by. `kioskRestartMs` is read
   * only for a kiosk show (absent: a kiosk show never restarts on its own).
   */
  readonly settings: Pick<ShowSettings, 'showAnimation' | 'useTimings' | 'loop' | 'type'> & { readonly kioskRestartMs?: number }
  /** Not in the visible order (hidden, or outside the show's range): never acquired by navigation. */
  hidden(i: number): boolean
  acquire(i: number, signal: AbortSignal): Promise<{ slide: ShowSlide; charge?(bytes: number, tag: string): boolean; release(): void }>
}

export type HitResult = { kind: 'trigger'; spid: string } | { kind: 'link'; link: LinkSpec } | { kind: 'withheld' }

export interface DriverStage {
  /** Prepares slide `index` in the background (a visible neighbour of the shown slide). */
  prepare(index: number): void
  enter(model: StageModel, t: number, neighbours: readonly number[]): Promise<void>
  frame(cursor: SlideCursor | undefined, t: number): RenderState | undefined
  /** The shown slide is being left: its current frame becomes the outgoing transition frame. */
  flattenCurrent(): void
  /** The shown slide is gone: everything it held but the outgoing frame is released. */
  leave(): void
  /** The topmost trigger or link at client px (clientX, clientY) on the shown slide, or `undefined` (the click passes). */
  hitTest(clientX: number, clientY: number, cursor: SlideCursor, t: number, on: { triggers: readonly string[]; links: readonly LinkSpec[] }): HitResult | undefined
  slideTime(t: number): number
  textTiming?: Stage['textTiming']
}

/** What the driver needs of a media player for one slide (`MediaController`'s following part). */
export interface DriverMedia {
  readonly wants: boolean
  sync(snap: SchedulerSnapshot, t: number, quiet?: boolean): void
  onEnd(cb: (nodeId: number, synthetic: boolean, t: number) => void): void
  leave(): void
}

export interface DriverHooks {
  /** The slide on screen and its click checkpoint changed (Next, Previous, an initial or last-checkpoint entry). */
  onSlideChange(i: number, build: number): void
  /** `v`: the visit on screen is drawable (its transition has started). `undefined`: the slide changed, nothing drawable yet. */
  onVisit(v: DriverVisit | undefined): void
  /** The cursor moved without a step report (a trigger, a media end). */
  onCursor(v: DriverVisit): void
  runLink(link: LinkSpec, v: DriverVisit): void
  /** Next on the end-of-show slide. */
  onEnd(): void
  /**
   * Additive: follows the slide's media (S2-7). Absent, audio and video nodes
   * get a synthetic end as soon as they begin (D16).
   */
  media?(slide: ShowSlide, slideTime: () => number): DriverMedia
  /** Additive: the audience screen is blanked (black or white): the after-timer auto-advance waits. */
  held?(): boolean
}

export interface DriverVisit {
  readonly index: number
  readonly slide: ShowSlide | undefined
  readonly timeline: SlideTimeline | undefined
  readonly cursor: SlideCursor | undefined
}

export interface ShowDriver {
  readonly slide: number | 'end'
  readonly visit: DriverVisit | undefined
  /** Enters the first slide (`'end'`: the show opens on end-of-show) and starts the kiosk timer. */
  start(from: number | 'end'): Promise<void>
  next(): void
  previous(): void
  goTo(i: number): void
  first(): void
  last(): void
  click(clientX: number, clientY: number): void
  trigger(spid: string): void
  /** One frame: the kiosk restart, auto-advance, media, then the stage draws; returns the state drawn. */
  tick(): RenderState | undefined
  dispose(): void
}

interface Visit {
  index: number
  at: 'initial' | 'last'
  lease?: { charge?(bytes: number, tag: string): boolean; release(): void }
  slide: ShowSlide | undefined
  timeline: SlideTimeline | undefined
  cursor: SlideCursor | undefined
  advance?: AutoAdvance
  media?: DriverMedia
  /** Without a media player: the slide's audio and video node ids, and those already ended synthetically. */
  synthetic?: { ids: ReadonlySet<number>; fired: Set<number> }
  /** Step inputs typed while the slide's model loads. */
  queue: Array<'next' | 'previous'>
  abort: AbortController
}

/** A slide with no readable metadata: static, no links, no transition. */
const staticSlide = (index: number): ShowSlide => ({ index, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] })

const MAX_DEPTH = 64

/** The ids of the audio and video nodes of a slide's timing (bounded walk). */
function mediaNodes(slide: ShowSlide): Set<number> {
  const ids = new Set<number>()
  const timing = slide.timing
  if (!timing || 'refused' in timing) return ids
  const walk = (n: TimeNode, depth: number) => {
    if (depth > MAX_DEPTH) return
    if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
    if (n.kind === 'audio' || n.kind === 'video') ids.add(n.ctn.id)
  }
  walk(timing.root, 0)
  return ids
}

export function createShowDriver(o: { source: SlideSource; stage: DriverStage; clock: SessionClock; sink: DiagnosticSink; reducedMotion: boolean; kiosk: boolean; hooks: DriverHooks }): ShowDriver {
  const { source, stage, clock, sink, hooks, kiosk } = o
  const settings = source.settings
  const count = source.count
  const order: number[] = []
  for (let i = 0; i < count; i++) if (!source.hidden(i)) order.push(i)
  const loop = settings.loop || kiosk

  let slide: number | 'end' = 'end'
  let visit: Visit | undefined
  let kioskTimer: KioskTimer | undefined
  let disposed = false

  /** The slide's own timeline time: show time since its transition ended (negative while it runs). */
  const slideTime = () => stage.slideTime(clock.now())

  /** Lets go of what visit `v` holds: everything but cross-slide media (numSld) stops (spec 7), its URLs before its lease. */
  function releaseVisit(v: Visit): void {
    v.media?.leave()
    v.abort.abort()
    const lease = v.lease
    v.lease = undefined
    lease?.release()
  }

  function leaveVisit(): void {
    if (visit) releaseVisit(visit)
    visit = undefined
  }

  /** Shows deck slide `i` (a cut, or a transition once its model is in) and loads its click steps. Resolves when the slide is drawable. */
  function enter(i: number | 'end', at: Visit['at'] = 'initial'): Promise<void> {
    leaveVisit()
    stage.flattenCurrent()
    stage.leave()
    slide = i
    hooks.onVisit(undefined)
    if (i === 'end') return Promise.resolve()
    const v: Visit = { index: i, at, slide: undefined, timeline: undefined, cursor: undefined, queue: [], abort: new AbortController() }
    visit = v
    // An initial entry is build 0 whatever the slide holds: reported at once, so a slide left before its
    // model arrives is still reported. A last-checkpoint entry is reported when its step count is known.
    if (at === 'initial') hooks.onSlideChange(i, 0)
    const live = () => visit === v && !disposed
    return source.acquire(i, v.abort.signal).then(
      (h) => {
        if (!live()) { h.release(); return undefined }
        v.lease = h
        v.slide = h.slide
        return h.slide.timing
      },
      (e) => {
        if (!live()) return undefined
        // The source degrades read and limit failures itself; anything left (aggregate refusal of even a static entry) is static.
        sink.add({ code: 'slide-refused', slideIndex: i, message: `Slide metadata refused (${e instanceof Error ? e.message : String(e)}); the slide is static` })
        return { refused: 'metadata' }
      },
    ).then((timing) => {
      if (!live()) return
      const options = { showAnimation: settings.showAnimation, reducedMotion: o.reducedMotion, sink, slideIndex: i }
      try {
        // Text by word or letter (S2-8): the stage's fanned-out working copy, charged to the slide's lease.
        const lease = v.lease
        const charge = lease?.charge
        const text = timing && !('refused' in timing) && lease && charge && v.slide && stage.textTiming
          ? stage.textTiming(i, v.slide, (bytes, tag) => charge.call(lease, bytes, tag)) : undefined
        v.timeline = createSlideTimeline(text?.tree ?? timing, text ? { ...options, orderOf: text.orderOf } : options)
      } catch (e) {
        // A bound the navigation layer refuses (spec: refuse and diagnose, never a wrong answer): the slide is static.
        sink.add({ code: 'slide-refused', slideIndex: i, message: `Slide timing refused (${e instanceof Error ? e.message : String(e)}); the slide is static` })
        v.timeline = createSlideTimeline({ refused: 'timeline' }, options)
      }
      v.slide ??= staticSlide(i)
      v.advance = createAutoAdvance(v.slide, settings as ShowSettings, v.timeline)
      if (hooks.media) {
        const media = v.media = hooks.media(v.slide, () => Math.max(0, slideTime()))
        media.onEnd((id, synthetic, t) => mediaEnded(v, id, synthetic, t))
      } else {
        const ids = mediaNodes(v.slide)
        if (ids.size) v.synthetic = { ids, fired: new Set() }
      }
      v.cursor = at === 'last' ? v.timeline.checkpoint(v.timeline.steps) : v.timeline.initial()
      const neighbours = [previousVisible(order, i), nextVisible(order, i)].filter((n): n is number => n !== undefined)
      return stage.enter({ index: i, slide: v.slide, timeline: v.timeline, cursor: v.cursor }, clock.now(), neighbours).then(() => neighbours)
    }).then((neighbours) => {
      if (!live() || !v.cursor || !neighbours) return
      // The visible neighbours are prepared once the slide is drawable.
      for (const n of neighbours) stage.prepare(n)
      hooks.onVisit(v)
      if (at === 'last') hooks.onSlideChange(i, v.cursor.step)
      // Replay the inputs typed while it loaded. One that changes the slide hands the rest to the new visit,
      // which replays them once it is ready in turn (PR #18 review 3); explicit navigation (goTo) supersedes them.
      const rest = v.queue.splice(0)
      for (let q = rest.shift(); q && !disposed; q = rest.shift()) {
        if (visit && visit !== v && !visit.cursor) { visit.queue.push(q, ...rest); break }
        q === 'next' ? next() : previous()
      }
    })
  }

  /** A media node's end, real or synthetic, joins the slide's history (spec 7); never earlier than its last event. */
  function mediaEnded(v: Visit, id: number, synthetic: boolean, t: number): void {
    if (visit !== v || disposed || !v.cursor || !v.timeline) return
    const last = v.cursor.mode === 'checkpoint' ? 0 : v.cursor.events[v.cursor.events.length - 1]?.t ?? 0
    v.cursor = v.timeline.mediaEnd(v.cursor, id, Math.max(t, last), synthetic)
    hooks.onCursor(v)
  }

  /** The next visible slide, the loop (a kiosk restart counts as one), or the end of the show. */
  function leaveForward(from: number): void {
    const after = nextVisible(order, from)
    if (after === undefined && loop) kioskTimer?.restarted(clock.now())
    void enter(after ?? (loop ? order[0] : undefined) ?? 'end')
  }

  function next(): void {
    if (disposed) return
    if (slide === 'end') { hooks.onEnd(); return }
    const v = visit!
    if (!v.cursor || !v.timeline) { v.queue.push('next'); return }
    const after = v.timeline.next(v.cursor, Math.max(0, slideTime()))
    // Identity: the timeline returns the same cursor when the slide has nothing left (S1-10).
    if (after !== v.cursor) { v.cursor = after; hooks.onSlideChange(v.index, after.step); return }
    leaveForward(v.index)
  }

  function previous(): void {
    if (disposed) return
    if (slide === 'end') {
      const to = previousVisible(order, Infinity)
      if (to !== undefined) void enter(to, 'last')
      return
    }
    const v = visit!
    if (!v.cursor || !v.timeline) { v.queue.push('previous'); return }
    if (v.cursor.step > 0) {
      v.cursor = v.timeline.checkpoint(v.cursor.step - 1)
      // A settled checkpoint: media that begins again ends again.
      v.synthetic?.fired.clear()
      hooks.onSlideChange(v.index, v.cursor.step)
      return
    }
    const to = previousVisible(order, v.index)
    if (to !== undefined) void enter(to, 'last')
  }

  function goTo(i: number): void {
    if (disposed) return
    if (!Number.isInteger(i) || i < 0 || i >= count) {
      sink.add({ code: 'invalid-slide', message: `goTo(${i}) is outside the deck; ignored.` })
      return
    }
    void enter(i)
  }

  const first = () => { if (!disposed && order.length) void enter(order[0]) }
  const last = () => { if (!disposed && order.length) void enter(order[order.length - 1]) }

  /** A trigger runs its interactive sequence (spec 4.3); never a click advance. Kiosk disables it (pending D4). */
  function trigger(spid: string): void {
    if (disposed || kiosk) return
    const v = visit
    if (!v || !v.cursor || !v.timeline || slide !== v.index) return
    const after = v.timeline.trigger(v.cursor, spid, Math.max(0, slideTime()))
    if (after) { v.cursor = after; hooks.onCursor(v) }
  }

  /** A click on the slide: the topmost trigger, else the topmost link, else an advance unless `advClick="0"`. */
  function click(clientX: number, clientY: number): void {
    if (disposed) return
    const v = visit
    if (slide !== 'end' && v?.cursor && v.timeline) {
      const triggers = v.timeline.triggers()
      const links = v.slide?.links ?? []
      const hit = triggers.length || links.length ? stage.hitTest(clientX, clientY, v.cursor, clock.now(), { triggers, links }) : undefined
      if (hit) {
        // A hit is consumed whatever it runs: kiosk disables triggers and links, and a withheld target runs nothing (D14).
        if (hit.kind === 'trigger') trigger(hit.spid)
        else if (hit.kind === 'link') { if (!kiosk && slide === v.index) hooks.runLink(hit.link, v) }
        else sink.add({ code: 'link-withheld', slideIndex: v.index, message: 'The link\'s target is not part of this show; the click does nothing' })
        return
      }
      // advClick="0": mouse clicks never advance; keyboard, API and presenter next still do (spec 4.4).
      if (v.advance && !v.advance.clickAdvances) return
    }
    next()
  }

  function tick(): RenderState | undefined {
    if (disposed) return undefined
    const t = clock.now()
    if (kioskTimer?.restartDue(t)) {
      kioskTimer.restarted(t)
      if (order.length) void enter(order[0])
      return undefined
    }
    const v = visit
    if (v && v.cursor && v.advance && slide === v.index && !hooks.held?.() && v.advance.due(v.cursor, 0, slideTime())) {
      // An automatic advance is a Next (never gated by advClick or kiosk; the show plays on its own).
      next()
    }
    // Media follows the slide's timeline (S2-7); a checkpoint entry records what has begun without playing it.
    const m = visit
    if (m && m.cursor && m.timeline && slide === m.index) {
      if (m.media?.wants) {
        const st = slideTime()
        if (st >= 0) m.media.sync(m.timeline.stateAt(m.cursor, st), st, m.cursor.mode === 'checkpoint')
      } else if (m.synthetic && m.cursor.mode !== 'checkpoint') syntheticEnds(m)
    }
    const w = visit
    const cursor = w && slide === w.index ? w.cursor : undefined
    return stage.frame(cursor, clock.now())
  }

  /** No media player (D16): an audio or video node that has begun ends at once, as local mode's refused media does. */
  function syntheticEnds(v: Visit): void {
    const st = slideTime()
    if (!(st >= 0) || !v.cursor || !v.timeline || !v.synthetic) return
    const { ids, fired } = v.synthetic
    for (const e of v.timeline.stateAt(v.cursor, st).active) {
      const id = e.node.ctn.id
      if (!ids.has(id) || fired.has(id)) continue
      fired.add(id)
      mediaEnded(v, id, true, st)
    }
  }

  return {
    get slide() { return slide },
    get visit() { return visit },
    async start(from) {
      await enter(from)
      // The kiosk restart counts from the first slide; a source without `kioskRestartMs` never restarts on its own.
      const ms = settings.kioskRestartMs
      if (!disposed && ms !== undefined) kioskTimer = createKioskTimer({ ...settings, kioskRestartMs: ms } as ShowSettings, clock.now(), sink)
    },
    next,
    previous,
    goTo,
    first,
    last,
    click,
    trigger,
    tick,
    dispose() {
      if (disposed) return
      disposed = true
      // The last visit stays readable (the host reads its cursor after the show ends); it holds nothing now.
      if (visit) releaseVisit(visit)
    },
  }
}
