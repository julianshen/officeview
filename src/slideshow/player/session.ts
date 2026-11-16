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
 * The navigation itself is the package-free show driver (player/driver.ts,
 * piece 4 Task 3): the session supplies its slide source (the package's
 * metadata store) and stage, and keeps the package, media, sounds,
 * accessibility and presenter wiring.
 *
 * Working assumptions (pending D4): Previous from slide n+1 enters slide n
 * at its last checkpoint; Next enters a slide at its initial state.
 */
import { preparedPlacement, SlideShowError, type ShowStart, type SlideShowSession } from '../../components/present'
import { emuToPx } from '../../core/geometry'
import { throwIfAborted } from '../../core/package-error'
import { validateSlideSize } from '../../pptx/appearance/size'
import type { PptxPackageHandle } from '../../pptx/types'
import { ManualClock, RafClock } from '../clock'
import { createDiagnosticSink } from '../diagnostics'
import type { SlideCursor } from '../engine/steps'
import { openShowPackage, readShowIndex } from '../read/deck'
import { shapeResolver } from '../read/resolve'
import { createSlideMetadataStore } from '../read/store'
import { createUnderlayReader } from '../read/underlay'
import { readSlideTiming } from '../read/timing'
import type { DiagnosticSink, LinkSpec, MetadataHandle, RenderState, ShowClock, ShowIndex, ShowSlide, SlideShowWorkspace, TerminalReserve } from '../types'
import { admitShow, createSlideShowWorkspace, type SlideShowProfile } from '../workspace'
import { createShowA11y, showTargets, slideTitle, visibleText, type ShowA11y, type ShowTarget } from './a11y'
import { createCursorHider, type CursorHider, type CursorTimer } from './cursor'
import { createShowDriver, type DriverStage, type DriverVisit, type ShowDriver, type SlideSource } from './driver'
import { createGestureState, gestureToInput, keyToInput, MODIFIER_KEYS } from './keys'
import { runLink } from './links'
import { visibleOrder } from './order'
import { createStage, type Stage, type StageDeps } from './stage'
import { createAudienceSurface, END_OF_SHOW, type AudienceSurface } from './surface'
import { adoptPresenterWindow, type PresenterWindow } from '../presenter/window'
import { mountPresenterView, type PresenterSource } from '../presenter/PresenterView'
import { createPresenterPlacement, type PresenterPlacement } from '../presenter/placement'
import { createElapsedTimer, type ElapsedTimer } from '../presenter/timer'
import type { PreviewOptions } from '../presenter/preview'
import { openMediaLease } from '../media/lease'
import { createMediaPlayer, type MediaPlayer } from '../media/player'
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

const NO_NOTES: ShowSlide['notes'] = []

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
    // Media lets go of its URLs before the leases (and the workspace) go: the driver leaves the shown slide's media, then its lease.
    driver?.dispose()
    sounds?.dispose()
    player?.dispose()
    stage?.dispose()
    store?.close()
    terminal?.release()
    clock?.dispose?.()
    cursorHider?.dispose()
    a11y?.dispose()
    surface?.dispose()
  }

  /** The navigation (created once the stage is up); `slide` and `visit` below read through it. */
  let driver: ShowDriver | undefined
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
  const underlay = createUnderlayReader(handle, index, sink, signal)
  store = createSlideMetadataStore({
    doc, handle, index, workspace, sink, signal,
    timing: readSlideTiming,
    resolver: (slide) => shapeResolver(slide, underlay.underlays.get(doc.slides.indexOf(slide)) ?? []),
  })

  const acquire = (i: number): Promise<MetadataHandle> => underlay.acquire(store!, i)

  // ---------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------

  /** The deck index on screen, or the end-of-show slide. */
  const current = (): number | 'end' => driver?.slide ?? 'end'
  /** The shown slide's visit (its metadata, timeline and cursor once in). */
  const shown = (): DriverVisit | undefined => driver?.visit
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
    // A stopped show clock pauses media too, so a clip's end stays on the slide's timeline (S2-7).
    player?.setPaused(paused || hiddenDoc)
    sounds?.setPaused(paused || hiddenDoc)
    if (!clock) return
    if (paused || hiddenDoc) clock.pause()
    else clock.resume()
  }
  const render = () => surface?.show({ slide: current(), count, screen })
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
    const v = shown()
    const t = sessionDeps.now(win)
    const busy = v?.timeline && v.cursor && stage ? v.timeline.busyUntil(v.cursor) > Math.max(0, slideTime()) : false
    const key = [current(), v?.cursor?.step ?? -1, v?.slide ? 1 : 0, busy, screen, paused, presenter, swapped ? 1 : 0, sink.list().length,
      timer ? Math.floor(timer.elapsed(t) / 1000) : 0, timer?.paused, Math.floor(sessionDeps.wall() / 60000)].join('|')
    if (!force && key === viewKey) return
    viewKey = key
    viewVersion++
    for (const cb of [...viewSubs]) cb()
  }
  /** The slide's own timeline time: show time since its transition ended (negative while it runs). */
  const slideTime = () => (stage ? stage.slideTime(clock!.now()) : 0)

  /** Live region: "Slide n of m: {title}" once per change (spec 9). */
  const announce = () => {
    if (!a11y) return
    const slide = current()
    if (slide === 'end') { a11y.announce(END_OF_SHOW); return }
    const title = slideTitle(doc.slides[slide])
    a11y.announce(`Slide ${slide + 1} of ${count}${title ? `: ${title}` : ''}`)
  }

  /** The visible text list and the trigger/link buttons for what is on screen now. */
  let a11yKey = ''
  function updateA11y(force = false, drawn?: RenderState): void {
    syncPresenter()
    if (!a11y || !stage) return
    const v = shown()
    if (current() === 'end' || !v || !v.cursor || screen !== 'slide') {
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
    a11y.setText(visibleText(pptx, state, v.slide?.builds ?? []))
    a11y.setTargets(kiosk ? [] : showTargets(pptx, state, v.timeline?.triggers() ?? [], v.slide?.links ?? [], v.slide?.builds ?? []))
  }

  /**
   * What the driver tells the session (spec 4.1, 7, 9): the slide changed (surface, live region, media and
   * sounds), it became drawable (its transition sound starts with the transition), its step or cursor moved.
   */
  const hooks = {
    onSlideChange(i: number, build: number) {
      if (started && !exited && current() === i) cb.onSlideChange?.(i, build)
      updateA11y()
    },
    onVisit(v: DriverVisit | undefined) {
      if (v) {
        // The transition sound starts with the transition: the stage has just set its start, now that the incoming
        // frame can be drawn (PR #18 review P1's showNow start; PR #19 review 2).
        sounds?.start(v.index)
        updateA11y(true)
        return
      }
      // Everything but cross-slide media (numSld) and a looping transition sound stops on a slide change (spec 7).
      sounds?.stopAt(lastShown)
      const slide = current()
      if (slide !== 'end') lastShown = slide
      screen = 'slide'
      render()
      announce()
      updateA11y(true)
    },
    onCursor: () => updateA11y(),
    runLink: (link: LinkSpec, v: DriverVisit) => openLink(link, v.index),
    onEnd: () => exit(),
    held: () => screen !== 'slide',
  }

  const next = () => { if (!exited) driver?.next() }
  const previous = () => { if (!exited) driver?.previous() }
  const goTo = (i: number) => { if (!exited) driver?.goTo(i) }
  const first = () => { if (!exited) driver?.first() }
  const last = () => { if (!exited) driver?.last() }

  function openLink(link: LinkSpec, slideIndex: number): void {
    if (win) runLink(link, { goTo, next, previous, first, last, exit }, win, { sink, slideIndex, open: start.platform.open })
  }

  /** A trigger runs its interactive sequence; a link runs its action (spec 4.3). Neither is a click advance. */
  function activate(target: { spid: string; kind: 'trigger' | 'link'; run?: { p: number; r: number } }): void {
    if (exited || !started || kiosk) return
    const v = shown()
    if (!v || !v.cursor || !v.timeline || current() !== v.index) return
    if (target.kind === 'trigger') { driver!.trigger(target.spid); return }
    // The link on the run named, else the shape's own; a target naming no run falls back to the first link.
    const links = (v.slide?.links ?? []).filter((l: LinkSpec) => l.spid === target.spid)
    const run = target.run
    const link = run ? links.find((l) => l.run?.p === run.p && l.run.r === run.r) : links.find((l) => !l.run) ?? links[0]
    if (link) openLink(link, v.index)
  }

  /** A click on the slide: the topmost trigger, else the topmost link, else an advance unless `advClick="0"`. */
  function click(clientX: number, clientY: number): void {
    if (!exited) driver?.click(clientX, clientY)
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
    if (exited || !started || !clock || !stage || !driver) return
    const drawn = driver.tick()
    // Presenter View copies the frame now, before the browser presents (and may clear) a WebGL canvas.
    afterFrame()
    const w = driver.visit
    const cursor = w && driver.slide === w.index ? w.cursor : undefined
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
    get buildIndex() { return shown()?.cursor?.step ?? 0 },
    get surface() { return surfaceKind },
    get presenter() { return presenter },
    get state(): SessionState {
      return { slide: current(), build: shown()?.cursor?.step ?? 0, screen, buffer, paused, clockPaused: paused || hiddenDoc, kiosk, loop, order, exited }
    },
    get stage() { return stage },
    get cursor() { return shown()?.cursor },
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
    notes: () => shown()?.slide?.notes ?? NO_NOTES,
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
      // Video plays in a DOM overlay, so the audience canvas lacks it: copy each frame from the one element (spec 5.2),
      // at the geometry its overlay keeps (a video carried past its slide included; PR #19 review 4).
      player?.drawVideos(ctx, dest)
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
  /** Removes the fullscreen-refused notice, when one is up. */
  let closeRefusedNotice: (() => void) | undefined
  /**
   * Every later full-screen request (the notice's button, "Use both screens", an opener-side placement retry) goes
   * through here, so a success records the surface and takes the refused notice down (PR #19 review 9).
   */
  const retryFullscreen: typeof start.retryFullscreen = (o) => start.retryFullscreen(o).then((r) => {
    if (r === 'entered' && !exited && start.isCurrent()) { surfaceKind = 'fullscreen'; closeRefusedNotice?.() }
    return r
  })
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
      frameBox: (spid) => stage?.mediaFrameBox?.(spid),
      frameToBox: (f) => stage?.frameBoxToCss?.(f),
      relieve: () => sounds?.release() ?? false,
      stopTransitionSound: (onlyLooping) => sounds?.release(onlyLooping) ?? false,
    })
    const mp = player
    sounds = createTransitionSound({ root: surface.root, sink, slide: (i) => { const v = shown(); return v?.index === i ? v.slide : undefined }, open: (rec, s) => mp.open(rec, s) })
    // The driver's slide source and stage: the package's metadata store (the hidden slides and the range are off
    // the visible order), and the engine, late-bound so a test seam that wraps a stage method still sees every call.
    const slides: SlideSource = {
      count,
      settings: index.settings,
      hidden: (i) => !order.includes(i),
      acquire: (i) => acquire(i).then((h) => ({ slide: h.slide, charge: (bytes: number, tag: string) => h.attach(bytes, tag), release: () => h.release() })),
    }
    const st = stage
    const engine: DriverStage = {
      prepare: (n) => st.prepare?.(n),
      enter: (m, t, n) => st.enter(m, t, n),
      frame: (c, t) => st.frame(c, t),
      // The local stage captures the outgoing frame inside its own leave() (seams that wrap leave see it there).
      flattenCurrent: () => {},
      leave: () => st.leave(),
      hitTest: (x, y, c, t, on) => st.hitTest?.(x, y, c, t, on),
      slideTime: (t) => st.slideTime(t),
      get textTiming() { return st.textTiming && ((i: number, meta: ShowSlide, charge: (bytes: number, tag: string) => boolean) => st.textTiming!(i, meta, charge)) },
    }
    driver = createShowDriver({
      source: slides, stage: engine, clock, sink, reducedMotion: motion, kiosk,
      hooks: { ...hooks, media: (slide, now) => mp.enter(slide, now) },
    })
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
    await driver.start(firstShown)
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
          popup, opener, host, retryFullscreen, fullscreen: start.fullscreen,
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
    stage.frame(driver.visit?.cursor, clock.now())
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
    const close = () => { closeRefusedNotice = undefined; notice.remove(); if (!exited) host.focus() }
    closeRefusedNotice = close
    notice.setAttribute('role', 'status')
    notice.setAttribute('data-officeview-show-notice', '')
    notice.textContent = 'Full screen is not available here; presenting in the window. '
    const retry = owner.createElement('button')
    retry.type = 'button'
    retry.textContent = 'Click for full screen'
    retry.onclick = () => {
      void retryFullscreen()
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
  const firstVisit = driver?.visit
  if (firstVisit) cb.onSlideChange?.(firstVisit.index, firstVisit.cursor?.step ?? 0)
  return session
}
