/**
 * The session controller (slideshow spec 4.1 exit and end-of-show, 4.2
 * session and starting slide, 4.3 key precedence, 4.4 hidden slides, loop,
 * kiosk input and pause, 6.2 hidden tab, 7 slide size, 8 lease fallback).
 * Checkpoint 1: static slides and cut transitions; S1-20 wires the engine.
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
import { SlideShowError, type ShowStart, type SlideShowSession } from '../../components/present'
import { emuToPx } from '../../core/geometry'
import { throwIfAborted } from '../../core/package-error'
import { HandlePackage } from '../../pptx/appearance/handle-package'
import { masterArtwork } from '../../pptx/appearance/master-art'
import { validateSlideSize } from '../../pptx/appearance/size'
import { slideTheme } from '../../pptx/parse'
import type { PptxPackageHandle, PptxShape } from '../../pptx/types'
import { ManualClock, RafClock } from '../clock'
import { createDiagnosticSink } from '../diagnostics'
import { createSlideTimeline, type SlideCursor, type SlideTimeline } from '../engine/steps'
import { openShowPackage, readRelationships, readShowIndex } from '../read/deck'
import { shapeResolver } from '../read/resolve'
import { createSlideMetadataStore } from '../read/store'
import { readSlideTiming } from '../read/timing'
import { diagnoseInheritedTransition } from '../read/transition'
import type { DiagnosticSink, MetadataHandle, ShowClock, ShowIndex, SlideShowWorkspace, TerminalReserve } from '../types'
import { admitShow, createSlideShowWorkspace, type SlideShowProfile } from '../workspace'
import { createGestureState, gestureToInput, keyToInput, MODIFIER_KEYS } from './keys'
import { nextVisible, previousVisible, visibleOrder } from './order'
import { createAudienceSurface, createStaticPainter, type AudienceSurface, type PainterContext, type SlidePainter } from './surface'

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

/** The running show. `dispatch` takes audience input (kiosk applies); the `SlideShowSession` methods are the host API (never gated by kiosk). */
export interface SessionController extends SlideShowSession {
  dispatch(input: ShowInput): void
  readonly state: SessionState
}

/** A show clock the session can stop and start (pause, hidden document). */
export interface SessionClock extends ShowClock {
  pause(): void
  resume(): void
  dispose?(): void
}

/** @internal Seams for tests: the package reopen, the clock, the painter and the workspace. */
export const sessionDeps = {
  openShowPackage: (source: Blob | ArrayBuffer | Uint8Array | string): Promise<PptxPackageHandle> => openShowPackage(source),
  clock: (win: Window | null): SessionClock => (win && typeof win.requestAnimationFrame === 'function' ? new RafClock(win) : new ManualClock()),
  painter: (ctx: PainterContext): SlidePainter => createStaticPainter(ctx),
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
  timeline?: SlideTimeline
  cursor?: SlideCursor
  /** Step inputs typed while the slide's metadata loads. */
  queue: Array<'next' | 'previous'>
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
  if (start.presenterView === true) {
    sink.add({ code: 'presenter-view-unavailable', message: 'Presenter View is not available yet; presenting audience only.' })
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

  let workspace: SlideShowWorkspace | undefined
  let terminal: TerminalReserve | undefined
  let store: ReturnType<typeof createSlideMetadataStore> | undefined
  let painter: SlidePainter | undefined
  let surface: AudienceSurface | undefined
  let clock: SessionClock | undefined
  const listeners = new AbortController()
  const release = () => {
    listeners.abort()
    visit?.handle?.release()
    store?.close()
    painter?.dispose()
    terminal?.release()
    clock?.dispose?.()
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
      const art = await masterArtwork(pkg, part, (await slideTheme(pkg, part)).context)
      return [...art.master, ...art.layout]
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
  let resolveEnded!: () => void
  const ended = new Promise<void>((resolve) => { resolveEnded = resolve })

  const syncClock = () => {
    if (!clock) return
    if (paused || hiddenDoc) clock.pause()
    else clock.resume()
  }
  const render = () => surface?.show({ slide, count, screen })
  const report = () => {
    if (started && !exited && visit?.cursor && slide === visit.index) cb.onSlideChange?.(visit.index, visit.cursor.step)
  }

  /** Shows deck slide `i` (a cut) and loads its click steps. Resolves when they are ready. */
  function enter(i: number | 'end', at: Visit['at'] = 'initial'): Promise<void> {
    visit?.handle?.release()
    visit = undefined
    slide = i
    screen = 'slide'
    render()
    if (i === 'end') return Promise.resolve()
    lastShown = i
    const v: Visit = { index: i, at, queue: [] }
    visit = v
    // An initial entry is build 0 whatever the slide holds: reported at once, so a slide left before its
    // metadata arrives is still reported. A last-checkpoint entry is reported when its step count is known.
    if (at === 'initial' && started) cb.onSlideChange?.(i, 0)
    const painted = painter!.paint(surface!.canvas, i)
    const live = () => visit === v && !exited
    const loaded = acquire(i).then(
      (h) => {
        if (!live()) { h.release(); return }
        v.handle = h
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
        v.timeline = createSlideTimeline(timing, options)
      } catch (e) {
        // A bound the navigation layer refuses (spec: refuse and diagnose, never a wrong answer): the slide is static.
        sink.add({ code: 'slide-refused', slideIndex: i, message: `Slide timing refused (${e instanceof Error ? e.message : String(e)}); the slide is static` })
        v.timeline = createSlideTimeline({ refused: 'timeline' }, options)
      }
      v.cursor = at === 'last' ? v.timeline.checkpoint(v.timeline.steps) : v.timeline.initial()
      if (at === 'last') report()
      for (const q of v.queue.splice(0)) {
        if (visit !== v || exited) break
        q === 'next' ? next() : previous()
      }
    })
    return Promise.all([painted, loaded]).then(() => {})
  }

  function next(): void {
    if (exited) return
    if (slide === 'end') { exit(); return }
    const v = visit!
    if (!v.cursor || !v.timeline) { v.queue.push('next'); return }
    const after = v.timeline.next(v.cursor, clock!.now())
    if (after !== v.cursor) { v.cursor = after; report(); return }
    // Nothing left on this slide: the next visible slide, the loop, or the end of the show.
    const to = nextVisible(order, v.index) ?? (loop ? order[0] : undefined)
    void enter(to ?? 'end')
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
    if (v.cursor.step > 0) { v.cursor = v.timeline.checkpoint(v.cursor.step - 1); report(); return }
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
    if (!fullscreenSettled || start.token.owns(start.platform.fullscreenElement(owner))) {
      const done = new AbortController()
      owner.addEventListener('fullscreenchange', () => {
        if (start.platform.fullscreenElement(owner)) return
        done.abort()
        cb.scrollToSlide(lastShown)
        cb.restoreFocus()
      }, { signal: done.signal })
      void start.fullscreen.then((outcome) => { if (outcome === 'refused') done.abort() })
    }
    cb.onEnd?.()
    resolveEnded()
  }

  /** Audience input, in the spec 4.3 precedence: Esc, then kiosk, the number buffer, black/white resume, the table. */
  function dispatch(input: ShowInput): void {
    if (exited || !started) return
    if (input.kind === 'exit') { exit(); return }
    // Kiosk: user input other than Esc is disabled (spec 4.4).
    if (kiosk) return
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
      return
    }
    switch (target.kind) {
      case 'next': case 'enter': case 'click': next(); break
      case 'previous': previous(); break
      case 'first': if (order.length) void enter(order[0]); break
      case 'last': if (order.length) void enter(order[order.length - 1]); break
      case 'black': case 'white': screen = target.kind; render(); break
      case 'pause': paused = !paused; syncClock(); break
    }
  }

  const session: SessionController = {
    get slideIndex() { return lastShown },
    get buildIndex() { return visit?.cursor?.step ?? 0 },
    get surface() { return surfaceKind },
    get presenter() { return 'none' as const },
    get state(): SessionState {
      return { slide, build: visit?.cursor?.step ?? 0, screen, buffer, paused, clockPaused: paused || hiddenDoc, kiosk, loop, order, exited }
    },
    next,
    previous,
    goTo,
    exit,
    dispatch,
    ended,
  }

  // ---------------------------------------------------------------------
  // Surface, clock and input
  // ---------------------------------------------------------------------

  let surfaceKind: 'fullscreen' | 'in-window' = 'in-window'
  try {
    clock = sessionDeps.clock(win)
    painter = sessionDeps.painter({ doc, terminal, sink })
    surface = createAudienceSurface(host, admitted.backing, (e) => dispatch({ kind: 'click', x: e.clientX, y: e.clientY }))
    const on = { signal: listeners.signal }
    const gestures = createGestureState(() => clock!.now())
    host.addEventListener('keydown', (e) => {
      // A focused control (the full-screen notice's buttons) keeps its own keys; Esc still exits through the start.
      if (e.target !== host && (e.target as Element | null)?.closest?.('button, input, textarea, select')) return
      const input = keyToInput(e, buffer)
      if (input) {
        e.preventDefault()
        dispatch(input)
      } else if (screen !== 'slide' && !MODIFIER_KEYS.has(e.key) && e.key !== 'Tab') {
        // "Any key except Esc resumes" (spec 4.3), including keys outside the table.
        e.preventDefault()
        dispatch({ kind: screen })
      }
    }, on)
    const gesture = (e: Event) => {
      const input = gestureToInput(e as WheelEvent | TouchEvent, gestures)
      if (e.type === 'wheel' && e.cancelable) e.preventDefault()
      if (input) dispatch(input)
    }
    for (const type of ['wheel', 'touchstart', 'touchend', 'touchcancel']) host.addEventListener(type, gesture, { ...on, passive: type !== 'wheel' })
    // Spec 6.2: while the document is hidden the show clock stops.
    const visibility = () => { hiddenDoc = owner.visibilityState === 'hidden'; syncClock() }
    owner.addEventListener('visibilitychange', visibility, on)
    visibility()

    const from = start.from === 'current' ? Math.min(Math.max(0, cb.currentSlide()), Math.max(0, count - 1)) : start.from
    const first = firstSlide({ index }, from)
    // Every start-time diagnostic reaches the host before onStart.
    await enter(first)
    if (!start.isCurrent()) throw new SlideShowError('cancelled')
    signal.addEventListener('abort', exit, on)
    started = true
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
    notice.setAttribute('role', 'status')
    notice.setAttribute('data-officeview-show-notice', '')
    notice.textContent = 'Full screen is not available here; presenting in the window. '
    const retry = owner.createElement('button')
    retry.type = 'button'
    retry.textContent = 'Click for full screen'
    retry.onclick = () => {
      void start.retryFullscreen().then((r) => { if (r === 'entered' && !exited && start.isCurrent()) { surfaceKind = 'fullscreen'; notice.remove() } })
    }
    const dismiss = owner.createElement('button')
    dismiss.type = 'button'
    dismiss.textContent = 'Dismiss'
    dismiss.onclick = () => notice.remove()
    notice.append(retry, dismiss)
    host.append(notice)
  })

  cb.onStart?.()
  // The first slide is reported after onStart (its metadata, and any start-time diagnostic, came first).
  if (visit) cb.onSlideChange?.(visit.index, visit.cursor?.step ?? 0)
  return session
}
