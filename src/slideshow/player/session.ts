/**
 * The session controller (slideshow spec 4.1 exit and end-of-show, 4.2
 * session and starting slide, 4.3 key precedence, 4.4 hidden slides, loop,
 * kiosk input and pause, 6.2 hidden tab, 7 slide size, 8 lease fallback).
 * S1-20 wires the engine (player/stage.ts): the per-frame tick drives the
 * stage, auto-advance and the kiosk restart timer; clicks are hit-tested
 * (triggers, then links, then `advClick`); the accessibility layer and the
 * cursor follow the slide; resize and DPR changes re-render the output.
 *
 * Model assembly happens here: the show index, the leased metadata store
 * with the timing reader and a target resolver over the slide and its
 * layout and master artwork, and the per-slide navigation timeline.
 *
 * Navigation is synchronous where the host can observe it: `slideIndex`,
 * the surface and the cut change at once. A slide's metadata (its click
 * steps) arrives asynchronously; step inputs typed meanwhile are queued for
 * that slide and replayed in order, and a stale load is released unused.
 *
 * Working assumptions (pending D4): Previous from slide n+1 enters slide n
 * at its last checkpoint; Next enters a slide at its initial state.
 */
import { preparedPlacement, SlideShowError, type ShowStart, type SlideShowSession } from '../../components/present'
import { emuToPx } from '../../core/geometry'
import { attrs } from '../../core/xml'
import { throwIfAborted } from '../../core/package-error'
import { HandlePackage } from '../../pptx/appearance/handle-package'
import { masterArtwork } from '../../pptx/appearance/master-art'
import { validateSlideSize } from '../../pptx/appearance/size'
import { slideTheme } from '../../pptx/parse'
import type { PptxPackageHandle, PptxShape } from '../../pptx/types'
import { ManualClock, RafClock } from '../clock'
import { createDiagnosticSink } from '../diagnostics'
import { createAutoAdvance, createKioskTimer, type AutoAdvance, type KioskTimer } from '../engine/advance'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../engine/steps'
import { openShowPackage, readRelationships, readShowIndex } from '../read/deck'
import { shapeResolver } from '../read/resolve'
import { createSlideMetadataStore } from '../read/store'
import { readSlideTiming } from '../read/timing'
import { diagnoseInheritedTransition } from '../read/transition'
import type { DiagnosticSink, LinkSpec, MetadataHandle, RenderState, ShowClock, ShowIndex, ShowSlide, SlideShowWorkspace, TerminalReserve } from '../types'
import { admitShow, createSlideShowWorkspace, type SlideShowProfile } from '../workspace'
import { createShowA11y, showTargets, slideTitle, visibleText, type ShowA11y, type ShowTarget } from './a11y'
import { createCursorHider, type CursorHider, type CursorTimer } from './cursor'
import { hitTarget } from './hit-test'
import { createGestureState, gestureToInput, keyToInput, MODIFIER_KEYS } from './keys'
import { runLink } from './links'
import { nextVisible, previousVisible, visibleOrder } from './order'
import { createStage, type Stage, type StageDeps } from './stage'
import { createAudienceSurface, END_OF_SHOW, type AudienceSurface } from './surface'
import { adoptPresenterWindow, type PresenterWindow } from '../presenter/window'
import { mountPresenterView, type PresenterSource } from '../presenter/PresenterView'
import { createPresenterPlacement, type PresenterPlacement } from '../presenter/placement'
import { createElapsedTimer, type ElapsedTimer } from '../presenter/timer'
import type { PreviewOptions } from '../presenter/preview'
import { openMediaLease } from '../media/lease'
import { createMediaPlayer, type MediaController, type MediaPlayer } from '../media/player'
import { createTransitionSound, type TransitionSoundController } from '../media/sounds'

export type ShowInput =
  | { kind: 'next' | 'previous' | 'first' | 'last' | 'exit' | 'black' | 'white' | 'pause' }
  | { kind: 'goto'; index: number }
  | { kind: 'digit'; d: number }
  | { kind: 'enter' }
  | { kind: 'click'; x: number; y: number }

export interface SessionState {
  /** The deck index on screen, or the end-of-show slide. */
  readonly slide: number | 'end'
  /** The click checkpoint within the slide (0 while its metadata loads). */
  readonly build: number
  readonly screen: 'slide' | 'black' | 'white'
  /** Digits typed for "go to slide n". */
  readonly buffer: string
  /** S or `+`. */
  readonly paused: boolean
  /** The show clock is stopped: paused, or the document is hidden. */
  readonly clockPaused: boolean
  readonly kiosk: boolean
  /** `loop`, or kiosk (which always loops). */
  readonly loop: boolean
  /** The visible order (deck indices). */
  readonly order: readonly number[]
  readonly exited: boolean
}

/**
 * The running show. `dispatch` takes audience input (kiosk applies); the
 * `SlideShowSession` methods are the host API, never gated by kiosk: spec
 * 4.4 disables *user input* in kiosk, and names no host-API rule.
 */
export interface SessionController extends SlideShowSession {
  dispatch(input: ShowInput): void
  /** Runs one frame now (the frame driver calls it every animation frame; tests call it directly). */
  tick(): void
  /** Activates a trigger or link the way a click on its shape does (the accessibility buttons use it). */
  activate(target: { spid: string; kind: 'trigger' | 'link'; run?: { p: number; r: number } }): void
  readonly state: SessionState
  /** @internal The engine, for tests and the debug harness. */
  readonly stage: Stage | undefined
  /** @internal The shown slide's cursor. */
  readonly cursor: SlideCursor | undefined
  /** @internal What Presenter View reads and drives (S2-5). */
  readonly presenterSource: PresenterSource
}

/** A show clock the session can stop and start (pause, hidden document). */
export interface SessionClock extends ShowClock {
  pause(): void
  resume(): void
  dispose?(): void
}

/** A per-frame callback driver: returns a stop function. */
export type FrameDriver = (win: Window | null, frame: () => void) => () => void

/** How long the post-exit focus restore waits for the browser to leave full screen (S1-19 review M1). */
export const RESTORE_LISTENER_MS = 5000

/** @internal Seams for tests: the package reopen, the clock, the engine, the frame driver, timers and the workspace. */
export const sessionDeps = {
  openShowPackage: (source: Blob | ArrayBuffer | Uint8Array | string): Promise<PptxPackageHandle> => openShowPackage(source),
  clock: (win: Window | null): SessionClock => (win && typeof win.requestAnimationFrame === 'function' ? new RafClock(win) : new ManualClock()),
  stage: (deps: StageDeps): Stage => createStage(deps),
  /** Additive (S1-20): `force: 'canvas2d'` for the forced Canvas 2D path (tests, the debug harness). */
  backend: undefined as 'canvas2d' | undefined,
  frames: ((win, frame) => {
    if (!win || typeof win.requestAnimationFrame !== 'function') return () => {}
    let id = 0
    let on = true
    const loop = () => { if (!on) return; frame(); id = win.requestAnimationFrame(loop) }
    id = win.requestAnimationFrame(loop)
    return () => { on = false; win.cancelAnimationFrame(id) }
  }) as FrameDriver,
  /** Wall timers for the cursor and the restore listener; never the show clock. */
  timer: (win: Window | null): CursorTimer => ({
    set: (cb, ms) => (win ?? globalThis).setTimeout(cb, ms),
    clear: (id) => (win ?? globalThis).clearTimeout(id as number),
  }),
  /** Monotonic ms for the context-restore deadline. */
  now: (win: Window | null): number => (win?.performance ?? globalThis.performance).now(),
  /** Epoch ms for Presenter View's wall clock (S2-5). */
  wall: (): number => Date.now(),
  workspace: (profile: SlideShowProfile): SlideShowWorkspace => createSlideShowWorkspace(profile),
}

/**
 * The starting slide (spec 4.2): 'beginning' is the first visible slide,
 * or the end-of-show slide when none is visible (Review Focus 4). A number
 * (the caller resolves 'current' to one) starts on that slide even if it is
 * hidden or outside the range.
 */
export function firstSlide(model: { index: ShowIndex }, from: 'beginning' | number): number | 'end' {
  if (from !== 'beginning') return from
  return visibleOrder(model.index)[0] ?? 'end'
}

/** Bytes, a Blob or a URL: what the viewer can read again. Anything else (a Response, a stream) is never re-read. */
function reReadable(s: unknown): s is Blob | ArrayBuffer | Uint8Array | string {
  return typeof s === 'string' || s instanceof Uint8Array || s instanceof ArrayBuffer || (typeof Blob !== 'undefined' && s instanceof Blob)
}

/** Every size the show will project, checked before any allocation (spec 7). */
function sizeProblem(start: ShowStart): string | undefined {
  const { doc } = start
  const deck = validateSlideSize(doc.slideWidthEmu, doc.slideHeightEmu)
  if (!deck.ok) return deck.diagnostic.message
  for (const slide of doc.slides) {
    const check = validateSlideSize(slide?.widthEmu, slide?.heightEmu)
    if (!check.ok) return check.diagnostic.message
  }
  return undefined
}

interface Visit {
  index: number
  at: 'initial' | 'last'
  handle?: MetadataHandle
  meta?: ShowSlide
  timeline?: SlideTimeline
  cursor?: SlideCursor
  advance?: AutoAdvance
  /** S2-7: the slide's media, following its timeline. */
  media?: MediaController
  /** Step inputs typed while the slide's metadata loads. */
  queue: Array<'next' | 'previous'>
}

const NO_NOTES: ShowSlide['notes'] = []

/** A slide with no readable metadata: static, no links, no transition. */
const staticSlide = (index: number): ShowSlide => ({ index, hidden: false, timing: undefined, notes: [], links: [], media: [], builds: [] })

/**
 * The viewer's protection policy (spec 6.3). Copy comes from the start's `allowCopy`, never from CSS a consumer
 * style can override (PR #18 review 8); print from the theme root's `data-ov-print`.
 */
function protection(start: ShowStart): { copy: boolean; print: boolean } {
  return { copy: start.allowCopy !== false, print: start.host.parentElement?.getAttribute('data-ov-print') !== 'false' }
}

export async function createSession(start: ShowStart): Promise<SessionController> {
  const { host, callbacks: cb, signal, doc } = start
  const owner = host.ownerDocument
  const win = owner.defaultView
  const count = doc.slides.length

  // Diagnostics: deduplicated, then forwarded to the host once each.
  const seen = createDiagnosticSink()
  const sink: DiagnosticSink = {
    add(d) {
      const before = seen.list().length
      seen.add(d)
      if (seen.list().length > before) cb.onDiagnostic?.(d)
    },
    list: () => seen.list(),
  }

  let fullscreenSettled = false
  void start.fullscreen.then(() => { fullscreenSettled = true })

  // Spec 7: an invalid slide size refuses the show before anything is read or allocated.
  const bad = sizeProblem(start)
  if (bad !== undefined) {
    sink.add({ code: 'slide-size', message: `${bad}; the deck cannot be presented` })
    throw new SlideShowError('not-presentable', bad)
  }

  // Spec 8: a handle-less document reopens a re-readable source; otherwise its slides are static ('no-package').
  let handle = start.handle
  if (!handle && start.source !== undefined && reReadable(start.source)) {
    try {
      handle = await sessionDeps.openShowPackage(start.source)
    } catch (e) {
      throwIfAborted(signal)
      sink.add({ code: 'package-reopen', message: `The document's source could not be read again (${e instanceof Error ? e.message : String(e)}); slides are static` })
    }
  }
  if (!start.isCurrent()) throw new SlideShowError('cancelled')
  // Spec 8: the show runs the slide appearance enrichment itself (backgrounds, master and layout artwork), so a
  // show started before any viewer painted the document looks the same. Idempotent: a viewer that already ran it
  // shares the same run. A failure only costs the inherited appearance; the enrichment diagnoses it on the document.
  // Through import(), as getPaintables loads it, so the enrichment stays its own chunk.
  await import('../../pptx/appearance').then((m) => m.enrichSlideAppearance(doc)).catch(() => undefined)
  if (!start.isCurrent()) throw new SlideShowError('cancelled')

  let workspace: SlideShowWorkspace | undefined
  let terminal: TerminalReserve | undefined
  let store: ReturnType<typeof createSlideMetadataStore> | undefined
  let stage: Stage | undefined
  let surface: AudienceSurface | undefined
  let a11y: ShowA11y | undefined
  let cursorHider: CursorHider | undefined
  let stopFrames: (() => void) | undefined
  let clock: SessionClock | undefined
  /** Spec 5.1: 'window' with the popup adopted; 'swap' (single-screen) when it was wanted but blocked or closed; 'none' when not wanted. */
  let presenter: SlideShowSession['presenter'] = start.popup === undefined ? 'none' : 'swap'
  let presenterWindow: PresenterWindow | undefined
  /** Presenter View in the popup, and the single-screen swap's view over the audience surface. */
  let view: { dispose(): void } | undefined
  let swapped: { dispose(): void } | undefined
  let placement: PresenterPlacement | undefined
  let player: MediaPlayer | undefined
  let sounds: TransitionSoundController | undefined
  const listeners = new AbortController()
  const release = () => {
    listeners.abort()
    view?.dispose()
    swapped?.dispose()
    placement?.dispose()
    viewSubs.clear()
    frameSubs.clear()
    presenterWindow?.close()
    stopFrames?.()
    // Media lets go of its URLs before the leases (and the workspace) go.
    visit?.media?.leave()
    sounds?.dispose()
    player?.dispose()
    visit?.handle?.release()
    stage?.dispose()
    store?.close()
    terminal?.release()
    clock?.dispose?.()
    cursorHider?.dispose()
    a11y?.dispose()
    surface?.dispose()
  }

  let visit: Visit | undefined
  let index: ShowIndex
  try {
    index = await readShowIndex(doc, handle, sink, signal)
  } catch (e) {
    throw signal.aborted ? new SlideShowError('cancelled') : e
  }
  const order = visibleOrder(index)
  const kiosk = index.settings.type === 'kiosk'
  const loop = index.settings.loop || kiosk
  const motion = start.motion === 'auto' ? (win?.matchMedia?.('(prefers-reduced-motion: reduce)').matches ?? false) : start.motion === 'reduced'

  // Spec 7: the permanent reserve (the output surface plus scratch) is admitted before anything is shown.
  const profile: SlideShowProfile = win?.matchMedia?.('(pointer: coarse)').matches ? 'mobile' : 'desktop'
  workspace = sessionDeps.workspace(profile)
  // The host fills the viewport (full screen or fixed in-window); with no layout at all (a document without a window) the slide's own size.
  const css = { w: host.clientWidth || win?.innerWidth || Math.round(emuToPx(doc.slideWidthEmu)), h: host.clientHeight || win?.innerHeight || Math.round(emuToPx(doc.slideHeightEmu)) }
  const admitted = admitShow(workspace, css, win?.devicePixelRatio ?? 1, sink)
  if (!admitted.ok) throw new SlideShowError('load-failed', 'The slideshow workspace cannot hold the output surface')
  terminal = admitted.terminal

  // The model assembly point: the timing reader and a resolver over the slide and its layout and master artwork.
  const underlays = new Map<number, readonly PptxShape[]>()
  store = createSlideMetadataStore({
    doc, handle, index, workspace, sink, signal,
    timing: readSlideTiming,
    resolver: (slide) => shapeResolver(slide, underlays.get(doc.slides.indexOf(slide)) ?? []),
  })

  /**
   * Layout and master artwork ids per layout, parsed once per show (S1-19
   * review M2): the resolver only needs the ids, so the cache keeps an
   * id-only skeleton (no geometry, text or image bytes). A slide that hides
   * the master shapes (`showMasterSp="0"`) has no underlay.
   */
  const underlayCache = new Map<string, Promise<readonly PptxShape[]>>()
  const skeleton = (list: readonly PptxShape[], depth = 0): PptxShape[] =>
    depth > 64 ? [] : list.map((s) => ({ source: s.source && { ...s.source }, children: s.children && skeleton(s.children, depth + 1) }) as PptxShape)
  /** Layout and master artwork for the resolver; inherited transitions are diagnosed on the way (they are not inherited). */
  async function underlayOf(i: number): Promise<readonly PptxShape[]> {
    const part = index.slides[i]?.partPath
    if (!handle || !part) return []
    try {
      const rels = await readRelationships(handle, part, signal)
      const local = (list: typeof rels, kind: string) => list.find((r) => !r.external && r.type?.endsWith(`/${kind}`))?.target
      const layout = local(rels, 'slideLayout')
      const master = layout ? local(await readRelationships(handle, layout, signal), 'slideMaster') : local(rels, 'slideMaster')
      // A private, bounded view over the handle: its cache dies with this call.
      const pkg = HandlePackage.over(handle)
      for (const [path, kind] of [[layout, 'layout'], [master, 'master']] as const) {
        const root = path ? await pkg.xmlOrdered(path) : undefined
        if (root) diagnoseInheritedTransition(root, kind, i, sink)
      }
      if (['0', 'false'].includes(attrs(await pkg.xmlOrdered(part)).showMasterSp ?? '')) return []
      const key = `${layout ?? ''}|${master ?? ''}`
      let art = underlayCache.get(key)
      if (!art) {
        art = slideTheme(pkg, part).then((theme) => masterArtwork(pkg, part, theme.context)).then((a) => skeleton([...a.master, ...a.layout]))
        underlayCache.set(key, art)
        art.catch(() => underlayCache.delete(key))
      }
      return await art
    } catch (e) {
      throwIfAborted(signal)
      sink.add({ code: 'underlay', slideIndex: i, message: `Layout and master artwork could not be read (${e instanceof Error ? e.message : String(e)}); master-only targets read as missing` })
      return []
    }
  }

  async function acquire(i: number): Promise<MetadataHandle> {
    const mine = store!.resident().includes(i) ? undefined : await underlayOf(i)
    if (mine) underlays.set(i, mine)
    try {
      return await store!.acquire(i, signal)
    } finally {
      // A second visit to the same slide may have joined the same read with its own copy: leave that one.
      if (mine && underlays.get(i) === mine) underlays.delete(i)
    }
  }

  // ---------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------

  let slide: number | 'end' = 'end'
  let lastShown = 0
  let screen: SessionState['screen'] = 'slide'
  let buffer = ''
  let paused = false
  let hiddenDoc = false
  let exited = false
  /** Input waits for the first slide: before that, Esc cancels through the start's signal. */
  let started = false
  let kioskTimer: KioskTimer | undefined
  let resolveEnded!: () => void
  const ended = new Promise<void>((resolve) => { resolveEnded = resolve })

  const syncClock = () => {
    // A stopped show clock pauses media too, so a clip's end stays on the slide's timeline (S2-7).
    player?.setPaused(paused || hiddenDoc)
    sounds?.setPaused(paused || hiddenDoc)
    if (!clock) return
    if (paused || hiddenDoc) clock.pause()
    else clock.resume()
  }
  const render = () => surface?.show({ slide, count, screen })
  /**
   * Presenter View follows the same state (one owner, spec 5.1): subscribers
   * are told when anything it shows changes — never per frame otherwise.
   */
  const viewSubs = new Set<() => void>()
  const frameSubs = new Set<() => void>()
  let viewVersion = 0
  let viewKey = ''
  let timer: ElapsedTimer | undefined
  const syncPresenter = (force = false) => {
    if (!viewSubs.size) return
    const v = visit
    const t = sessionDeps.now(win)
    const busy = v?.timeline && v.cursor && stage ? v.timeline.busyUntil(v.cursor) > Math.max(0, slideTime()) : false
    const key = [slide, v?.cursor?.step ?? -1, v?.meta ? 1 : 0, busy, screen, paused, presenter, swapped ? 1 : 0, sink.list().length,
      timer ? Math.floor(timer.elapsed(t) / 1000) : 0, timer?.paused, Math.floor(sessionDeps.wall() / 60000)].join('|')
    if (!force && key === viewKey) return
    viewKey = key
    viewVersion++
    for (const cb of [...viewSubs]) cb()
  }
  const report = () => {
    if (started && !exited && visit?.cursor && slide === visit.index) cb.onSlideChange?.(visit.index, visit.cursor.step)
  }
  /** The slide's own timeline time: show time since its transition ended (negative while it runs). */
  const slideTime = () => (stage ? stage.slideTime(clock!.now()) : 0)

  /** Live region: "Slide n of m: {title}" once per change (spec 9). */
  const announce = () => {
    if (!a11y) return
    if (slide === 'end') { a11y.announce(END_OF_SHOW); return }
    const title = slideTitle(doc.slides[slide])
    a11y.announce(`Slide ${slide + 1} of ${count}${title ? `: ${title}` : ''}`)
  }

  /** The visible text list and the trigger/link buttons for what is on screen now. */
  let a11yKey = ''
  function updateA11y(force = false, drawn?: RenderState): void {
    syncPresenter()
    if (!a11y || !stage) return
    const v = visit
    if (slide === 'end' || !v || !v.cursor || screen !== 'slide') {
      if (force || a11yKey !== '') { a11yKey = ''; a11y.setText([]); a11y.setTargets([]) }
      return
    }
    // A frame passes the state it just drew for this cursor (S1-22: one state per frame); other callers compute it.
    const state = drawn ?? stage.stateAt(v.cursor, clock!.now())
    let key = `${v.index}:${v.cursor.step}:`
    for (const [k, s] of state) if (!s.visible || !(s.opacity > 0)) key += `${k},`
    if (!force && key === a11yKey) return
    a11yKey = key
    const pptx = doc.slides[v.index]
    a11y.setText(visibleText(pptx, state, v.meta?.builds ?? []))
    a11y.setTargets(kiosk ? [] : showTargets(pptx, state, v.timeline?.triggers() ?? [], v.meta?.links ?? [], v.meta?.builds ?? []))
  }

  /** Shows deck slide `i` (a cut, or a transition once its model is in) and loads its click steps. Resolves when the slide is drawable. */
  function enter(i: number | 'end', at: Visit['at'] = 'initial'): Promise<void> {
    const left = visit
    if (left) {
      // Everything but cross-slide media (numSld) and a looping transition sound stops here (spec 7).
      left.media?.leave()
      sounds?.stopAt(left.index)
    }
    visit?.handle?.release()
    visit = undefined
    stage?.leave()
    slide = i
    screen = 'slide'
    render()
    announce()
    updateA11y(true)
    if (i === 'end') return Promise.resolve()
    lastShown = i
    const v: Visit = { index: i, at, queue: [] }
    visit = v
    // An initial entry is build 0 whatever the slide holds: reported at once, so a slide left before its
    // metadata arrives is still reported. A last-checkpoint entry is reported when its step count is known.
    if (at === 'initial' && started) cb.onSlideChange?.(i, 0)
    const live = () => visit === v && !exited
    return acquire(i).then(
      (h) => {
        if (!live()) { h.release(); return undefined }
        v.handle = h
        v.meta = h.slide
        return h.slide.timing
      },
      (e) => {
        if (!live()) return undefined
        // The store degrades read and limit failures itself; anything left (aggregate refusal of even a static entry) is static.
        sink.add({ code: 'slide-refused', slideIndex: i, message: `Slide metadata refused (${e instanceof Error ? e.message : String(e)}); the slide is static` })
        return { refused: 'metadata' }
      },
    ).then((timing) => {
      if (!live()) return
      const options = { showAnimation: index.settings.showAnimation, reducedMotion: motion, sink, slideIndex: i }
      try {
        // Text by word or letter (S2-8): the stage's fanned-out working copy, charged to the slide's metadata.
        const handle = v.handle
        const text = timing && !('refused' in timing) && handle && v.meta && stage?.textTiming
          ? stage.textTiming(i, v.meta, (bytes, tag) => handle.attach(bytes, tag)) : undefined
        v.timeline = createSlideTimeline(text?.tree ?? timing, text ? { ...options, orderOf: text.orderOf } : options)
      } catch (e) {
        // A bound the navigation layer refuses (spec: refuse and diagnose, never a wrong answer): the slide is static.
        sink.add({ code: 'slide-refused', slideIndex: i, message: `Slide timing refused (${e instanceof Error ? e.message : String(e)}); the slide is static` })
        v.timeline = createSlideTimeline({ refused: 'timeline' }, options)
      }
      v.meta ??= staticSlide(i)
      v.advance = createAutoAdvance(v.meta, index.settings, v.timeline)
      if (player) {
        const media = v.media = player.enter(v.meta, () => Math.max(0, slideTime()))
        media.onEnd((id, synthetic, t) => mediaEnded(v, id, synthetic, t))
      }
      v.cursor = at === 'last' ? v.timeline.checkpoint(v.timeline.steps) : v.timeline.initial()
      const neighbours = [previousVisible(order, i), nextVisible(order, i)].filter((n): n is number => n !== undefined)
      return stage!.enter({ index: i, slide: v.meta, timeline: v.timeline, cursor: v.cursor }, clock!.now(), neighbours)
    }).then(() => {
      if (!live() || !v.cursor) return
      // The transition sound starts with the transition: the stage has just set its start, now that the incoming
      // frame can be drawn (PR #18 review P1's showNow start; PR #19 review 2).
      sounds?.start(i)
      if (at === 'last') report()
      updateA11y(true)
      // Replay the inputs typed while it loaded. One that changes the slide hands the rest to the new visit,
      // which replays them once it is ready in turn (PR #18 review 3); explicit navigation (goTo) supersedes them.
      const rest = v.queue.splice(0)
      for (let q = rest.shift(); q && !exited; q = rest.shift()) {
        if (visit && visit !== v && !visit.cursor) { visit.queue.push(q, ...rest); break }
        q === 'next' ? next() : previous()
      }
    })
  }

  /** A media node's end, real or synthetic, joins the slide's history (spec 7); never earlier than its last event. */
  function mediaEnded(v: Visit, id: number, synthetic: boolean, t: number): void {
    if (visit !== v || exited || !v.cursor || !v.timeline) return
    const last = v.cursor.mode === 'checkpoint' ? 0 : v.cursor.events[v.cursor.events.length - 1]?.t ?? 0
    v.cursor = v.timeline.mediaEnd(v.cursor, id, Math.max(t, last), synthetic)
    updateA11y()
  }

  /** The next visible slide, the loop (a kiosk restart counts as one), or the end of the show. */
  function leaveForward(from: number): void {
    const after = nextVisible(order, from)
    if (after === undefined && loop) kioskTimer?.restarted(clock!.now())
    void enter(after ?? (loop ? order[0] : undefined) ?? 'end')
  }

  function next(): void {
    if (exited) return
    if (slide === 'end') { exit(); return }
    const v = visit!
    if (!v.cursor || !v.timeline) { v.queue.push('next'); return }
    const after = v.timeline.next(v.cursor, Math.max(0, slideTime()))
    // Identity: the timeline returns the same cursor when the slide has nothing left (S1-10).
    if (after !== v.cursor) { v.cursor = after; report(); updateA11y(); return }
    leaveForward(v.index)
  }

  function previous(): void {
    if (exited) return
    if (slide === 'end') {
      const to = previousVisible(order, Infinity)
      if (to !== undefined) void enter(to, 'last')
      return
    }
    const v = visit!
    if (!v.cursor || !v.timeline) { v.queue.push('previous'); return }
    if (v.cursor.step > 0) { v.cursor = v.timeline.checkpoint(v.cursor.step - 1); report(); updateA11y(); return }
    const to = previousVisible(order, v.index)
    if (to !== undefined) void enter(to, 'last')
  }

  function goTo(i: number): void {
    if (exited) return
    if (!Number.isInteger(i) || i < 0 || i >= count) {
      sink.add({ code: 'invalid-slide', message: `goTo(${i}) is outside the deck; ignored.` })
      return
    }
    void enter(i)
  }

  const first = () => { if (order.length) void enter(order[0]) }
  const last = () => { if (order.length) void enter(order[order.length - 1]) }

  /** A trigger runs its interactive sequence; a link runs its action (spec 4.3). Neither is a click advance. */
  function activate(target: { spid: string; kind: 'trigger' | 'link'; run?: { p: number; r: number } }): void {
    if (exited || !started || kiosk) return
    const v = visit
    if (!v || !v.cursor || !v.timeline || slide !== v.index) return
    if (target.kind === 'trigger') {
      const after = v.timeline.trigger(v.cursor, target.spid, Math.max(0, slideTime()))
      if (after) { v.cursor = after; updateA11y() }
      return
    }
    // The link on the run named, else the shape's own; a target naming no run falls back to the first link.
    const links = (v.meta?.links ?? []).filter((l: LinkSpec) => l.spid === target.spid)
    const run = target.run
    const link = run ? links.find((l) => l.run?.p === run.p && l.run.r === run.r) : links.find((l) => !l.run) ?? links[0]
    if (link && win) runLink(link, { goTo, next, previous, first, last, exit }, win, { sink, slideIndex: v.index, open: start.platform.open })
  }

  /** A click on the slide: the topmost trigger, else the topmost link, else an advance unless `advClick="0"`. */
  function click(clientX: number, clientY: number): void {
    const v = visit
    if (slide !== 'end' && v?.cursor && v.timeline && stage?.plan) {
      const at = stage.toSlide(clientX, clientY)
      const triggers = v.timeline.triggers()
      const links = v.meta?.links ?? []
      const st = stage
      // A trigger beats a link on the same shape (pending D4; recorded in the catalogue). A link on text runs
      // is taken only over the laid-out run it is on (PR #18 review 3); elsewhere the shape's own link, if any,
      // else the click passes to the shapes below. Text that cannot be located keeps the first run's link.
      const accept = (spid: string, leaf: PptxShape, lx: number, ly: number, w: number, h: number): Parameters<typeof activate>[0] | undefined => {
        if (triggers.includes(spid)) return { spid, kind: 'trigger' }
        const mine = links.filter((l) => l.spid === spid)
        if (!mine.length) return undefined
        const own = mine.some((l) => !l.run)
        const runs = mine.filter((l) => l.run)
        if (runs.length && leaf.source?.id === spid) {
          const run = st.runAt(leaf, lx, ly, w, h)
          if (run && runs.some((l) => l.run!.p === run.p && l.run!.r === run.r)) return { spid, kind: 'link', run }
          if (run === undefined && !own) return { spid, kind: 'link', run: runs[0].run }
        }
        return own ? { spid, kind: 'link' } : undefined
      }
      const target = at && (triggers.length || links.length) ? hitTarget(st.plan!, st.stateAt(v.cursor, clock!.now()), at.x, at.y, accept) : undefined
      if (target) { activate(target); return }
      // advClick="0": mouse clicks never advance; keyboard, API and presenter next still do (spec 4.4).
      if (v.advance && !v.advance.clickAdvances) return
    }
    next()
  }

  function exit(): void {
    if (exited) return
    exited = true
    release()
    host.replaceChildren()
    start.cleanup()
    cb.scrollToSlide(lastShown)
    cb.restoreFocus()
    // While the host is full screen the rest of the page is inert, so focus cannot leave it yet: restore
    // again (and re-scroll, the viewport changes) once the browser has left full screen. A request still
    // in flight may enter full screen after this exit (the start then leaves it at once), so it is awaited too.
    // The listener is bounded: if the browser never reports leaving full screen, it goes after RESTORE_LISTENER_MS.
    if (!fullscreenSettled || start.token.owns(start.platform.fullscreenElement(owner))) {
      const done = new AbortController()
      const timer = sessionDeps.timer(win)
      const t = timer.set(() => done.abort(), RESTORE_LISTENER_MS)
      done.signal.addEventListener('abort', () => timer.clear(t))
      owner.addEventListener('fullscreenchange', () => {
        if (start.platform.fullscreenElement(owner)) return
        done.abort()
        cb.scrollToSlide(lastShown)
        cb.restoreFocus()
      }, { signal: done.signal })
      void start.fullscreen.then((outcome) => { if (outcome === 'refused') done.abort() })
    }
    // A throwing host callback still settles `ended` (PR #18 review 2); the error reaches the caller.
    try { cb.onEnd?.() } finally { resolveEnded() }
  }

  /** One frame: the kiosk restart, auto-advance, then the stage draws (spec 4.4; every timer is show-clock time). */
  function tick(): void {
    if (exited || !started || !clock || !stage) return
    const t = clock.now()
    if (kioskTimer?.restartDue(t)) {
      kioskTimer.restarted(t)
      if (order.length) void enter(order[0])
      return
    }
    const v = visit
    if (v && v.cursor && v.advance && slide === v.index && screen === 'slide' && v.advance.due(v.cursor, 0, slideTime())) {
      // An automatic advance is a Next (never gated by advClick or kiosk; the show plays on its own).
      next()
    }
    // Media follows the slide's timeline (S2-7); a checkpoint entry records what has begun without playing it.
    const m = visit
    if (m && m.media?.wants && m.cursor && m.timeline && slide === m.index) {
      const st = slideTime()
      if (st >= 0) m.media.sync(m.timeline.stateAt(m.cursor, st), st, m.cursor.mode === 'checkpoint')
    }
    const w = visit
    const cursor = w && slide === w.index ? w.cursor : undefined
    const drawn = stage.frame(cursor, clock.now())
    // Presenter View copies the frame now, before the browser presents (and may clear) a WebGL canvas.
    afterFrame()
    updateA11y(false, cursor ? drawn : undefined)
  }

  function afterFrame(): void {
    for (const f of frameSubs) {
      try { f() } catch (e) {
        frameSubs.delete(f)
        sink.add({ code: 'presenter-paint', message: `Presenter View stopped mirroring the audience (${e instanceof Error ? e.message : String(e)})` })
      }
    }
  }

  /** Single-screen (spec 5.1): Alt+F5 or the toggle swaps the surface between the audience and Presenter View. */
  function toggleView(): void {
    if (exited || !started || presenter === 'window' || !surface || !win) return
    presenter = 'swap'
    if (swapped) {
      swapped.dispose()
      swapped = undefined
      surface.root.hidden = false
      // The audience gets focus back: an API host may not be focusable yet.
      if (!host.hasAttribute('tabindex')) host.tabIndex = -1
      host.focus()
    } else {
      const box = owner.createElement('div')
      box.setAttribute('data-officeview-presenter-swap', '')
      // allowPrint=false: the swapped-in view and its notes stay out of print too (spec 6.3).
      if (!policy.print) box.setAttribute('data-ov-print', 'false')
      host.append(box)
      const mounted = mountPresenterView(box, { session, host: 'swap', doc: owner })
      swapped = { dispose: () => { mounted.dispose(); box.remove() } }
      surface.root.hidden = true
    }
    syncPresenter(true)
  }

  /** Audience input, in the spec 4.3 precedence: Esc, then kiosk, the number buffer, black/white resume, the table. */
  function dispatch(input: ShowInput): void { run(input, false) }

  /**
   * Input, audience or presenter (controller ruling C6): kiosk disables the
   * audience's input only; Presenter View's controls and keys still drive the show.
   */
  function run(input: ShowInput, presenterInput: boolean): void {
    if (exited || !started) return
    if (input.kind === 'exit') { exit(); return }
    // Kiosk: user input other than Esc is disabled (spec 4.4), triggers and links included (pending D4).
    if (kiosk && !presenterInput) return
    if (input.kind === 'digit') { buffer += String(input.d); return }
    let target: ShowInput = input
    if (input.kind === 'enter' && buffer) target = { kind: 'goto', index: Number(buffer) - 1 }
    if (target.kind === 'goto') {
      buffer = ''
      screen = 'slide'
      render()
      goTo(target.index)
      return
    }
    buffer = ''
    if (screen !== 'slide') {
      // Any input resumes from a black or white screen, and does nothing else.
      screen = 'slide'
      render()
      updateA11y(true)
      return
    }
    switch (target.kind) {
      case 'click': click(target.x, target.y); break
      case 'next': case 'enter': next(); break
      case 'previous': previous(); break
      case 'first': first(); break
      case 'last': last(); break
      case 'black': case 'white': screen = target.kind; render(); updateA11y(true); break
      case 'pause': paused = !paused; syncClock(); break
    }
  }

  const session: SessionController = {
    get slideIndex() { return lastShown },
    get buildIndex() { return visit?.cursor?.step ?? 0 },
    get surface() { return surfaceKind },
    get presenter() { return presenter },
    get state(): SessionState {
      return { slide, build: visit?.cursor?.step ?? 0, screen, buffer, paused, clockPaused: paused || hiddenDoc, kiosk, loop, order, exited }
    },
    get stage() { return stage },
    get cursor() { return visit?.cursor },
    next,
    previous,
    goTo,
    exit,
    dispatch,
    tick,
    activate,
    ended,
    get presenterSource() { return source },
  }

  /** The audience frame on screen: the compositor's canvas, or the terminal output when the stage is hidden. */
  const audienceFrame = (): HTMLCanvasElement | undefined => {
    const box = surface?.root.querySelector<HTMLElement>('[data-officeview-show-stage]')
    const shown = box && !box.hidden ? [...box.querySelectorAll<HTMLCanvasElement>('canvas[data-slideshow-backend]')].pop() : undefined
    return shown ?? surface?.canvas
  }
  let policy = { copy: true, print: true }
  let presenterKey: (e: KeyboardEvent) => void = () => {}
  let warnedImages = false
  let model: { index: ShowIndex; store: NonNullable<typeof store> } | undefined
  const source: PresenterSource = {
    subscribe(cb) {
      viewSubs.add(cb)
      return () => { viewSubs.delete(cb) }
    },
    get version() { return viewVersion },
    deck: doc,
    notes: () => visit?.meta?.notes ?? NO_NOTES,
    diagnostics: () => sink.list(),
    sink,
    get timer() { return timer ??= createElapsedTimer(sessionDeps.now(win)) },
    now: () => sessionDeps.now(win),
    wall: () => sessionDeps.wall(),
    slideTime: () => Math.max(0, slideTime()),
    get model() { return store ? (model ??= { index, store }) : undefined },
    get previewOptions() {
      // The preview schedules the same expanded text timing as the live slide (PR #19 review 2).
      const textTiming: PreviewOptions['textTiming'] = (i, meta, charge) => stage?.textTiming?.(i, meta, charge)
      return { showAnimation: index.settings.showAnimation, reducedMotion: motion, sink, textTiming }
    },
    get workspace() { return workspace! },
    ...(start.watermark ? { watermark: start.watermark } : {}),
    get protection() { return policy },
    audienceFrame,
    drawMedia(ctx, dest) {
      // Video plays in a DOM overlay, so the audience canvas lacks it: copy each frame from the one element (spec 5.2).
      if (!player || !stage?.mediaFrameBox) return
      for (const spid of player.videos()) {
        const b = stage.mediaFrameBox(spid)
        if (b) player.drawFrame(spid, ctx, dest.x + b.x * dest.w, dest.y + b.y * dest.h, b.w * dest.w, b.h * dest.h)
      }
    },
    onFrame(cb) {
      frameSubs.add(cb)
      return () => { frameSubs.delete(cb) }
    },
    previewOf(i, cursor, timeline) {
      const p = stage?.previewOf?.(i, cursor, timeline)
      if (p) return p
      // Best effort: a slide the stage does not hold paints its static state without decoded pictures.
      if (!warnedImages && doc.images?.length) {
        warnedImages = true
        sink.add({ code: 'presenter-thumbnail-approximate', message: 'Presenter View paints slides away from the current one without their pictures; the audience is unaffected' })
      }
      return { slide: doc.slides[i]!, state: new Map(), images: [], resolveFont: (f: string) => f, builds: [] }
    },
    key: (e) => presenterKey(e),
    input: (i) => run(i, true),
    toggleView,
    get placement() { return placement },
    touch: () => syncPresenter(true),
  }

  // ---------------------------------------------------------------------
  // Surface, clock and input
  // ---------------------------------------------------------------------

  let surfaceKind: 'fullscreen' | 'in-window' = 'in-window'
  try {
    clock = sessionDeps.clock(win)
    policy = protection(start)
    surface = createAudienceSurface(host, admitted.backing, (e) => dispatch({ kind: 'click', x: e.clientX, y: e.clientY }))
    // allowPrint=false keeps every slideshow surface out of print (spec 6.3 protection policy).
    if (!policy.print) surface.root.setAttribute('data-ov-print', 'false')
    a11y = createShowA11y(surface.root, { selectable: policy.copy, activate: (t: ShowTarget) => activate(t) })
    const watermark = start.watermark
    stage = sessionDeps.stage({
      doc, root: surface.root, output: surface.canvas, workspace, terminal, backing: admitted.backing, store, sink,
      reduced: motion, now: () => sessionDeps.now(win), showNow: () => clock!.now(),
      schedule: (fn, ms) => sessionDeps.timer(win).set(fn, ms),
      metadata: (i) => acquire(i),
      ...(watermark ? { watermark } : {}),
      ...(sessionDeps.backend ? { force: sessionDeps.backend } : {}),
    })
    // S2-7: one media player per show, on the audience surface only (spec 5.2); leases are charged to the show's workspace.
    const ws = workspace
    player = createMediaPlayer({
      root: surface.root,
      sink,
      open: (rec, s) => (handle && win ? openMediaLease(handle, rec, ws, win, s) : Promise.resolve({ refused: 'no-package' })),
      box: (spid) => stage?.mediaBox?.(spid),
      relieve: () => sounds?.release() ?? false,
      stopTransitionSound: (onlyLooping) => sounds?.release(onlyLooping) ?? false,
    })
    const mp = player
    sounds = createTransitionSound({ root: surface.root, sink, slide: (i) => (visit?.index === i ? visit.meta : undefined), open: (rec, s) => mp.open(rec, s) })
    cursorHider = createCursorHider(surface.root, sessionDeps.timer(win))
    const on = { signal: listeners.signal }
    const gestures = createGestureState(() => clock!.now())
    const keyInput = (e: KeyboardEvent, presenterInput = false) => {
      const input = keyToInput(e, buffer)
      if (input) {
        e.preventDefault()
        run(input, presenterInput)
      } else if (screen !== 'slide' && !MODIFIER_KEYS.has(e.key) && e.key !== 'Tab') {
        // "Any key except Esc resumes" (spec 4.3), including keys outside the table.
        e.preventDefault()
        run({ kind: screen }, presenterInput)
      }
    }
    /** Alt+F5 (spec 5.1): the single-screen swap, from any focus; Ctrl or Cmd with it stays the browser's. */
    const swapKey = (e: KeyboardEvent) => {
      if (e.key !== 'F5' || !e.altKey || e.ctrlKey || e.metaKey) return false
      e.preventDefault()
      toggleView()
      return true
    }
    host.addEventListener('keydown', (e) => {
      // The swapped-in Presenter View keeps the popup's rule (its buttons keep only Enter and Space).
      if ((e.target as Element | null)?.closest?.('[data-officeview-presenter-view]')) { presenterKey(e); return }
      if (swapKey(e)) return
      // A focused control (the full-screen notice's buttons, the accessibility buttons) keeps its own keys; Esc still exits through the start.
      if (e.target !== host && (e.target as Element | null)?.closest?.('button, input, textarea, select')) return
      keyInput(e)
    }, on)
    /**
     * Keys in the presenter window (spec 4.3, 5.2): Esc always exits; a text control, an editable region
     * or the scrolling speaker notes keep every other key; a focused button keeps Enter and Space.
     */
    presenterKey = (e: KeyboardEvent) => {
      if (swapKey(e)) return
      const t = e.target as Element | null
      if (e.key !== 'Escape' && (t?.closest?.('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [data-officeview-presenter-notes]') || ((e.key === 'Enter' || e.key === ' ') && t?.closest?.('button')))) return
      keyInput(e, true)
    }
    const gesture = (e: Event) => {
      // The swapped-in Presenter View scrolls its own notes and grid: wheel and swipe there are not show input.
      if ((e.target as Element | null)?.closest?.('[data-officeview-presenter-view]')) return
      const input = gestureToInput(e as WheelEvent | TouchEvent, gestures)
      if (e.type === 'wheel' && e.cancelable) e.preventDefault()
      if (input) dispatch(input)
    }
    for (const type of ['wheel', 'touchstart', 'touchend', 'touchcancel']) host.addEventListener(type, gesture, { ...on, passive: type !== 'wheel' })
    // Spec 6.2: while the document is hidden the show clock stops.
    const visibility = () => { hiddenDoc = owner.visibilityState === 'hidden'; syncClock() }
    owner.addEventListener('visibilitychange', visibility, on)
    visibility()
    // Spec 6.3: on resize, full-screen change, screen or DPR change the output follows (a new reserve first).
    const resized = () => {
      if (exited || !stage) return
      const w = host.clientWidth || win?.innerWidth || 0, h = host.clientHeight || win?.innerHeight || 0
      if (w > 0 && h > 0) stage.resize({ w, h }, win?.devicePixelRatio ?? 1)
      player?.layout()
    }
    win?.addEventListener('resize', resized, on)
    owner.addEventListener('fullscreenchange', resized, on)
    const watchDpr = () => {
      const q = win?.matchMedia?.(`(resolution: ${win.devicePixelRatio ?? 1}dppx)`)
      q?.addEventListener?.('change', () => { resized(); watchDpr() }, { ...on, once: true })
    }
    watchDpr()

    // An empty deck has no current slide: it opens on end-of-show, as a start from the beginning does (PR #18 review 4).
    const from = start.from === 'current' ? (count ? Math.min(Math.max(0, cb.currentSlide()), count - 1) : 'beginning') : start.from
    const firstShown = firstSlide({ index }, from)
    // Every start-time diagnostic reaches the host before onStart.
    await enter(firstShown)
    if (!start.isCurrent()) throw new SlideShowError('cancelled')
    // Presenter View's elapsed timer counts from the first slide (spec 5.2), on the injected monotonic clock.
    timer = createElapsedTimer(sessionDeps.now(win))
    // Spec 5.1: adopt the popup the gesture opened; blocked (null) or already closed is single-screen.
    const popup = start.popup
    if (popup === null || popup?.closed) {
      sink.add({ code: 'popup-blocked', message: 'The presenter window was blocked; presenting single-screen.' })
    } else if (popup) {
      // Best effort (owner ruling): a popup that cannot be adopted is closed and the show runs single-screen.
      try {
        const opener = win ?? popup.opener
        const pw = presenterWindow = adoptPresenterWindow(popup, opener, sink)
        const theme = host.parentElement?.getAttribute('data-ov-theme')
        if (theme) pw.root.setAttribute('data-ov-theme', theme)
        // allowPrint=false keeps the presenter window out of print too (spec 6.3).
        if (!policy.print) pw.root.setAttribute('data-ov-print', 'false')
        placement = createPresenterPlacement({
          popup, opener, host, retryFullscreen: start.retryFullscreen, fullscreen: start.fullscreen,
          prepared: preparedPlacement.get(owner), sink, changed: () => syncPresenter(true),
        })
        presenter = 'window'
        view = mountPresenterView(pw.root, { session, host: 'window', doc: pw.doc })
        // Closing the presenter window never ends the show: single-screen from here.
        pw.onClosed(() => {
          view?.dispose()
          view = undefined
          placement?.dispose()
          placement = undefined
          presenterWindow = undefined
          if (!exited) { presenter = 'swap'; syncPresenter(true) }
        })
      } catch (e) {
        view?.dispose()
        view = undefined
        placement?.dispose()
        placement = undefined
        presenterWindow = undefined
        popup.close()
        presenter = 'swap'
        sink.add({ code: 'presenter-window-failed', message: `The presenter window could not be set up (${e instanceof Error ? e.message : String(e)}); presenting single-screen.` })
      }
    }
    signal.addEventListener('abort', exit, on)
    started = true
    // Presenter View was asked for but there is one screen (blocked, or the window could not be set up): it starts swapped in.
    if (presenter === 'swap') toggleView()
    kioskTimer = createKioskTimer(index.settings, clock.now(), sink)
    stage.frame(visit?.cursor, clock.now())
    stopFrames = sessionDeps.frames(win, tick)
  } catch (e) {
    release()
    host.replaceChildren()
    throw e instanceof SlideShowError ? e : signal.aborted ? new SlideShowError('cancelled') : e
  }

  start.fullscreen.then((outcome) => {
    if (exited || !start.isCurrent()) return
    if (outcome === 'entered') { surfaceKind = 'fullscreen'; return }
    surfaceKind = 'in-window'
    sink.add({ code: 'fullscreen-refused', message: 'Full screen was refused; the show runs in the window.' })
    const notice = owner.createElement('div')
    // Removing the notice takes its focused button away: focus returns to the host, so its keys (Esc) keep working (PR #18 review 8).
    const close = () => { notice.remove(); if (!exited) host.focus() }
    notice.setAttribute('role', 'status')
    notice.setAttribute('data-officeview-show-notice', '')
    notice.textContent = 'Full screen is not available here; presenting in the window. '
    const retry = owner.createElement('button')
    retry.type = 'button'
    retry.textContent = 'Click for full screen'
    retry.onclick = () => {
      void start.retryFullscreen().then((r) => { if (r === 'entered' && !exited && start.isCurrent()) { surfaceKind = 'fullscreen'; close() } })
    }
    const dismiss = owner.createElement('button')
    dismiss.type = 'button'
    dismiss.textContent = 'Dismiss'
    dismiss.onclick = close
    notice.append(retry, dismiss)
    host.append(notice)
  })

  cb.onStart?.()
  // The first slide is reported after onStart (its metadata, and any start-time diagnostic, came first).
  if (visit) cb.onSlideChange?.(visit.index, visit.cursor?.step ?? 0)
  return session
}
