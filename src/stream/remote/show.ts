/**
 * The remote show session (streaming Stage 3 piece 3, design §4.3; parent
 * §8.2, §12a.5, §12a.7, §12a.8). Lazy: loaded by the remote `present()` start
 * only; nothing eager imports this module.
 *
 * A streamed deck is shown slide by slide from the server-marked bitmaps the
 * show fetcher (show-fetch.ts) keeps inside the `'unit'` rate class. No local
 * rendering, no animations, transitions or media: every slide is its final
 * image, fitted (letterboxed) to the audience surface. Hidden units are never
 * shown, counted, prefetched or asked for notes.
 *
 * Index space: everything the host and Presenter View see is a *streamed*
 * index (0-based over the shown slides, hidden excluded): `slideIndex`,
 * `goTo()`, `onSlideChange`, `state.slide`, `state.order` (0..n-1),
 * `slideCount`, `slideImage(i)`. The streamed → unit map stays private.
 * The viewer callbacks `currentSlide`/`scrollToSlide` speak the viewer's page
 * (unit) index.
 *
 * Notes (presenter role with `can.notes` only): one server-drawn image for the
 * current slide, requested once the slide has settled, painted only into the
 * Presenter View notes pane. A client-side `rate-limited` refusal is retried
 * once when the fetcher has nothing in flight (R11). With `can.notes` off no
 * notes request is ever sent and `notes()` is empty.
 *
 * `doc.ended` (revocation, expiry, any terminal end) exits the show; nothing
 * paints after exit. The show never closes the document: the viewer owns it.
 */
import { preparedPlacement, SlideShowError, type ShowStart, type SlideShowSession } from '../../components/present'
import { createDiagnosticSink } from '../../slideshow/diagnostics'
import { createCursorHider, type CursorHider, type CursorTimer } from '../../slideshow/player/cursor'
import { createGestureState, gestureToInput, keyToInput, MODIFIER_KEYS } from '../../slideshow/player/keys'
import { nextVisible, previousVisible } from '../../slideshow/player/order'
import type { SessionController, SessionState, ShowInput } from '../../slideshow/player/session'
import { createAudienceSurface, END_OF_SHOW, type AudienceSurface } from '../../slideshow/player/surface'
import type { ImageNotes } from '../../slideshow/presenter/notes'
import { mountPresenterView, type PresenterSource } from '../../slideshow/presenter/PresenterView'
import { createPresenterPlacement, type PresenterPlacement } from '../../slideshow/presenter/placement'
import { createElapsedTimer, type ElapsedTimer } from '../../slideshow/presenter/timer'
import { adoptPresenterWindow, type PresenterWindow } from '../../slideshow/presenter/window'
import type { DiagnosticSink, Lease, SlideShowWorkspace, TerminalReserve } from '../../slideshow/types'
import { admitShow, backingFor, createSlideShowWorkspace, surfaceBytes, type Backing, type SlideShowProfile } from '../../slideshow/workspace'
import { StreamSessionError } from './errors'
import { bucketFor, type RemoteOfficeDoc } from './remote-doc'
import { createShowFetcher, type ImageBitmapLike, type RemoteErrorCode, type ShowFetchClock, type ShowFetcher } from './show-fetch'

/** What the remote `present()` start hands the lazy show: the local start contract over a streamed deck. */
export interface RemoteShowStart extends Omit<ShowStart, 'doc' | 'handle' | 'source' | 'isCurrent'> {
  doc: RemoteOfficeDoc
  /** Defaults to "the start's signal has not been aborted". */
  isCurrent?(): boolean
  /** The deck's show settings are not streamed: a host may ask for looping (or kiosk, which always loops). */
  loop?: boolean
  kiosk?: boolean
}

/** How long the current slide must stay current before its notes are asked for (no notes-cancel on the server, R9). */
export const NOTES_SETTLE_MS = 300
/** How long the post-exit focus restore waits for the browser to leave full screen (as the local session). */
const RESTORE_LISTENER_MS = 5000
/** The server's notes image bounds (session.ts): 720 CSS px × scale wide, at most 8,192 device px tall. */
const NOTES_WIDTH_CSS = 720, NOTES_MAX_PX = 8192
/** Presenter View's timer and clock refresh. */
const PRESENTER_TICK_MS = 1000

/** @internal Seams for tests: the fetcher clock, the presenter refresh, wall timers and the workspace. */
export const remoteShowDeps = {
  /** The fetcher's (and the notes settle delay's) clock. */
  clock: (win: Window | null): Required<ShowFetchClock> => {
    const w = win ?? globalThis
    return {
      now: () => (win?.performance ?? globalThis.performance).now(),
      wake: (ms, cb) => { const id = w.setTimeout(cb, ms); return () => w.clearTimeout(id as number) },
    }
  },
  /** A repeating wall timer; returns the stop. */
  every: (win: Window | null, ms: number, cb: () => void): (() => void) => {
    const w = win ?? globalThis
    const id = w.setInterval(cb, ms)
    return () => w.clearInterval(id)
  },
  timer: (win: Window | null): CursorTimer => ({
    set: (cb, ms) => (win ?? globalThis).setTimeout(cb, ms),
    clear: (id) => (win ?? globalThis).clearTimeout(id as number),
  }),
  wall: (): number => Date.now(),
  workspace: (profile: SlideShowProfile): SlideShowWorkspace => createSlideShowWorkspace(profile),
}

const NO_NOTES: readonly never[] = []
const HIDDEN = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip-path:inset(50%);white-space:nowrap;border:0'
const zero = (c: { width: number; height: number }) => { c.width = 0; c.height = 0 }

/** Draws `img` fitted into `canvas` (letterboxed on black). */
function fitInto(canvas: HTMLCanvasElement, img: ImageBitmapLike): boolean {
  const ctx = canvas.width > 0 && canvas.height > 0 ? canvas.getContext('2d') : null
  if (!ctx || !(img.width > 0) || !(img.height > 0)) return false
  const k = Math.min(canvas.width / img.width, canvas.height / img.height)
  const w = img.width * k, h = img.height * k
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.drawImage(img as unknown as CanvasImageSource, (canvas.width - w) / 2, (canvas.height - h) / 2, w, h)
  ctx.restore()
  return true
}

/** A neutral frame with one line of text (loading, or a slide that could not be shown). */
function textFrame(canvas: HTMLCanvasElement, text: string): void {
  const ctx = canvas.width > 0 && canvas.height > 0 ? canvas.getContext('2d') : null
  if (!ctx) return
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.fillStyle = '#1f1f1f'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
  ctx.fillStyle = '#d0d0d0'
  ctx.font = `${Math.max(12, Math.round(canvas.height / 24))}px sans-serif`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'middle'
  ctx.fillText(text, canvas.width / 2, canvas.height / 2)
  ctx.restore()
}

export async function runRemoteShow(req: RemoteShowStart): Promise<SessionController> {
  const start = { isCurrent: () => !req.signal.aborted, ...req }
  const { host, callbacks: cb, signal, doc } = start
  const owner = host.ownerDocument
  const win = owner.defaultView
  const list = doc.manifest.units.list
  /** Streamed index → unit index: the shown slides, hidden units left out. */
  const unitOf: number[] = []
  list.forEach((u, i) => { if (!u.hidden) unitOf.push(i) })
  const count = unitOf.length
  const order: readonly number[] = unitOf.map((_, i) => i)
  const kiosk = !!start.kiosk
  const loop = !!start.loop || kiosk

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
  if (!start.isCurrent()) throw new SlideShowError('cancelled')

  // The output surface is admitted before anything is shown (as the local show, spec 7).
  const profile: SlideShowProfile = win?.matchMedia?.('(pointer: coarse)').matches ? 'mobile' : 'desktop'
  const workspace = remoteShowDeps.workspace(profile)
  const first = list[unitOf[0] ?? 0] ?? { w: 1, h: 1 }
  const cssSize = () => ({ w: host.clientWidth || win?.innerWidth || first.w, h: host.clientHeight || win?.innerHeight || first.h })
  const admitted = admitShow(workspace, cssSize(), win?.devicePixelRatio ?? 1, sink)
  if (!admitted.ok) throw new SlideShowError('load-failed', 'The slideshow workspace cannot hold the output surface')
  const terminal: TerminalReserve = admitted.terminal
  let backing: Backing = admitted.backing

  // State
  let slide: number | 'end' = 'end'
  let lastShown = 0
  let screen: SessionState['screen'] = 'slide'
  let buffer = ''
  let paused = false
  let exited = false
  let started = false
  let resolveEnded!: () => void
  const ended = new Promise<void>((resolve) => { resolveEnded = resolve })

  let surface: AudienceSurface | undefined
  let live: HTMLElement | undefined
  let frameText: HTMLElement | undefined
  let cursorHider: CursorHider | undefined
  let presenter: SlideShowSession['presenter'] = start.popup === undefined ? 'none' : 'swap'
  let presenterWindow: PresenterWindow | undefined
  let view: { dispose(): void } | undefined
  let swapped: { dispose(): void } | undefined
  let placement: PresenterPlacement | undefined
  let timer: ElapsedTimer | undefined
  let stopTick: (() => void) | undefined
  let stopManifest: (() => void) | undefined
  const listeners = new AbortController()
  const viewSubs = new Set<() => void>()
  const frameSubs = new Set<() => void>()
  const clock = remoteShowDeps.clock(win)

  // ---------------------------------------------------------------------
  // Fetching: the show fetcher over RemoteOfficeDoc.paint (R2, R16)
  // ---------------------------------------------------------------------

  /** Unit requests the io has open; notes retries wait for none (R11). */
  let open = 0
  let idle: Array<() => void> = []
  const settledOne = () => {
    if (--open > 0) return
    const run = idle
    idle = []
    for (const f of run) f()
  }
  /**
   * Retained slide surfaces are byte-admitted (spec 7): each bitmap's bytes are leased from the show's
   * workspace before its request goes out and held until it is evicted, replaced or disposed. The cache
   * holds what a share of the room left after the audience output fits at the largest bucket; with less
   * room, fewer neighbours are prefetched (the wanted set never exceeds the cache, so nothing thrashes).
   */
  const maxBucket = Math.max(1, ...doc.manifest.recorder.scalesBuckets)
  const destBytes = (unit: number, bucket: number) => { const u = list[unit] ?? first; return surfaceBytes(Math.ceil(u.w * bucket), Math.ceil(u.h * bucket)) }
  const slots = Math.floor((workspace.limit - workspace.used()) / 2 / destBytes(unitOf[0] ?? 0, maxBucket))
  const maxCached = Math.max(1, Math.min(8, slots))
  /** Neighbours prefetched besides the target: at most 3 (2 next + 1 previous), always fewer than the cache holds. */
  const neighbours = Math.max(0, Math.min(3, maxCached - 2))
  const fetcher: ShowFetcher = createShowFetcher({
    async paint(unit, scale, dpr, sig, charge) {
      // The destination's bytes first: no room is a typed refusal, and nothing is requested.
      const lease = workspace.lease(destBytes(unit, scale * dpr), 'remote-show:slide')
      if (!lease) throw new StreamSessionError('workspace-full', 'the slideshow workspace has no room for this slide')
      open++
      // R2: a fresh canvas per bitmap; `close` frees it and its bytes.
      const canvas = owner.createElement('canvas')
      zero(canvas)
      const free = () => { zero(canvas); lease.release() }
      // An abandoned request gives its bytes back now, not when its late result arrives.
      sig.addEventListener('abort', free, { once: true })
      try {
        await doc.paint(unit, scale, dpr, canvas, sig, { onSend: charge })
        // The doc returns without painting after an abort or the end: nothing to keep.
        if (sig.aborted || !(canvas.width > 0)) { free(); throw new StreamSessionError('closed', 'nothing painted') }
        sig.removeEventListener('abort', free)
        // The raster may be smaller than the bucket surface: keep only what it holds.
        lease.shrink(lease.bytes - surfaceBytes(canvas.width, canvas.height))
        return Object.assign(canvas, { close: free })
      } catch (e) {
        free()
        throw e
      } finally {
        settledOne()
      }
    },
  }, clock, { maxCached })
  /** Units whose last request failed (cleared once a bitmap lands). */
  const failed = new Map<number, RemoteErrorCode>()
  /** Bumped on every content change Presenter View shows (bitmaps, errors, notes). */
  let content = 0

  /** The bucket for the audience backing: the slide fitted to it, as a server bucket (so the key moves only when the bucket does). */
  const applyBucket = () => {
    const fit = Math.min(backing.w / first.w, backing.h / first.h)
    fetcher.setBucket(bucketFor(fit, 1, doc.manifest.recorder.scalesBuckets), 1)
  }

  /** Current + 2 next + 1 previous (§12a.7), target first. */
  const want = () => {
    if (slide === 'end') { fetcher.want(count ? [unitOf[count - 1]!] : []); return }
    const p = slide
    const near = [p + 1, p + 2].map((q) => (q < count ? q : loop ? q % count : -1))
    const wanted = [p, ...[...near, p - 1].filter((q) => q >= 0 && q < count && q !== p).slice(0, neighbours)]
    fetcher.want(wanted.map((q) => unitOf[q]!))
  }

  // ---------------------------------------------------------------------
  // The audience frame
  // ---------------------------------------------------------------------

  function draw(): void {
    if (exited || !surface) return
    const canvas = surface.canvas
    let kind: 'image' | 'loading' | 'error' | 'end' = 'end'
    let text = ''
    if (slide !== 'end') {
      const unit = unitOf[slide]!
      const bmp = fetcher.cached(unit)
      const code = failed.get(unit)
      if (bmp && fitInto(canvas, bmp)) kind = 'image'
      else {
        kind = code ? 'error' : 'loading'
        text = code ? `Slide ${slide + 1} could not be shown (${code})` : `Loading slide ${slide + 1}`
        textFrame(canvas, text)
      }
    }
    surface.root.setAttribute('data-officeview-show-frame', kind)
    if (frameText && frameText.textContent !== text) frameText.textContent = text
    afterFrame()
  }

  /** Each subscriber still subscribed, stopping once one of them ends the show. */
  function notify(set: Set<() => void>, onThrow?: (f: () => void, e: unknown) => void): void {
    for (const f of [...set]) {
      if (exited) return
      if (!set.has(f)) continue
      if (!onThrow) { f(); continue }
      try { f() } catch (e) { onThrow(f, e) }
    }
  }

  function afterFrame(): void {
    if (exited) return
    notify(frameSubs, (f, e) => {
        frameSubs.delete(f)
        sink.add({ code: 'presenter-paint', message: `Presenter View stopped mirroring the audience (${e instanceof Error ? e.message : String(e)})` })
    })
  }

  fetcher.onReady((unit) => {
    if (exited) return
    failed.delete(unit)
    content++
    if (slide !== 'end' && unitOf[slide] === unit) draw()
    syncPresenter()
  })
  fetcher.onError((unit, code) => {
    if (exited) return
    // A rate refusal pauses the fetcher, which re-issues the unit: not a failed slide.
    if (code === 'rate-limited') return
    failed.set(unit, code)
    content++
    if (slide !== 'end' && unitOf[slide] === unit) draw()
    syncPresenter()
  })

  // ---------------------------------------------------------------------
  // Notes (presenter, can.notes only)
  // ---------------------------------------------------------------------

  let notes: ImageNotes | undefined
  let notesUnit: number | undefined
  let notesImage: HTMLCanvasElement | undefined
  /** The notes image's bytes and its Presenter View copy's (the pane canvas), held while the notes are shown. */
  let notesLease: Lease | undefined
  /**
   * The one pane canvas holding a copy of the notes image (Presenter View shows one pane at a time). Its bytes are
   * half of `notesLease`; it is destroyed (zeroed) before they are refunded, and when a new pane takes the copy.
   */
  let notesPane: HTMLCanvasElement | undefined
  let notesAbort: AbortController | undefined
  let notesWait: (() => void) | undefined
  /** One per followed slide: a retry queued for an earlier visit never runs. */
  let notesGen = 0
  const notesOn = () => doc.manifest.can.notes && presenter !== 'none'
  const setNotes = (state: ImageNotes['state'], image?: HTMLCanvasElement) => {
    notes = {
      kind: 'image',
      state,
      paint(c) {
        if (exited || !image || !(image.width > 0) || image !== notesImage) return
        if (notesPane && notesPane !== c) zero(notesPane)
        notesPane = c
        c.width = image.width; c.height = image.height
        const ctx = c.getContext('2d')
        ctx?.drawImage(image, 0, 0)
      },
    }
    content++
    syncPresenter()
  }
  const dropNotes = () => {
    notesGen++
    notesWait?.(); notesWait = undefined
    notesAbort?.abort(); notesAbort = undefined
    if (notesImage) zero(notesImage)
    notesImage = undefined
    if (notesPane) zero(notesPane)
    notesPane = undefined
    notesLease?.release()
    notesLease = undefined
    notesUnit = undefined
    if (notes) { notes = undefined; content++ }
  }
  /** The current slide's notes: asked for once it has been current NOTES_SETTLE_MS (the server has no notes-cancel). */
  function followNotes(): void {
    const unit = slide === 'end' ? undefined : unitOf[slide]
    if (exited || !notesOn() || unit === undefined) { dropNotes(); syncPresenter(); return }
    if (notesUnit === unit) return
    dropNotes()
    notesUnit = unit
    setNotes('loading')
    const gen = notesGen
    notesWait = clock.wake(NOTES_SETTLE_MS, () => { notesWait = undefined; requestNotes(unit, gen, true) })
  }
  function requestNotes(unit: number, gen: number, retry: boolean): void {
    if (exited || gen !== notesGen || notesUnit !== unit || !notesOn()) return
    const dpr = win?.devicePixelRatio ?? 1
    // The worst-case image (720 × scale wide, up to 8,192 px tall) is reserved before the request, then
    // trimmed to the delivered image plus its pane copy; no room is an error in the pane, nothing sent.
    const notesScale = bucketFor(1, dpr, doc.manifest.recorder.scalesBuckets.filter((b) => b === 1 || b === 2).concat(2))
    const lease = workspace.lease(surfaceBytes(NOTES_WIDTH_CSS * notesScale, NOTES_MAX_PX), 'remote-show:notes')
    if (!lease) { sink.add({ code: 'workspace', message: 'The slideshow workspace has no room for the notes image' }); setNotes('error'); return }
    const ctl = new AbortController()
    notesAbort = ctl
    const image = owner.createElement('canvas')
    zero(image)
    const stale = () => exited || notesAbort !== ctl || ctl.signal.aborted
    const free = () => { zero(image); lease.release() }
    ctl.signal.addEventListener('abort', free, { once: true })
    doc.paintNotes(unit, dpr, image, ctl.signal).then(() => {
      if (stale() || !(image.width > 0)) { free(); if (!stale()) { notesAbort = undefined; setNotes('error') } return }
      notesAbort = undefined
      ctl.signal.removeEventListener('abort', free)
      const bytes = surfaceBytes(image.width, image.height)
      if (lease.bytes >= 2 * bytes) lease.shrink(lease.bytes - 2 * bytes)
      else if (!lease.grow(2 * bytes - lease.bytes)) {
        free()
        sink.add({ code: 'workspace', message: 'The slideshow workspace has no room for the notes image' })
        setNotes('error')
        return
      }
      notesImage = image
      notesLease = lease
      setNotes('ready', image)
    }, (e: unknown) => {
      free()
      if (stale()) return
      notesAbort = undefined
      const code = e instanceof StreamSessionError ? e.code : undefined
      if (code === 'rate-limited' && retry) {
        // R11: once, after the fetcher's unit requests have drained (their leases free the room).
        // The wake lets the session's serial inbox release the leases first; a unit issued meanwhile is waited for too.
        const again = (): void => {
          if (gen !== notesGen) return
          if (open > 0) { idle.push(again); return }
          notesWait = clock.wake(0, () => { notesWait = undefined; if (open > 0) again(); else requestNotes(unit, gen, false) })
        }
        again()
        return
      }
      if (!doc.manifest.can.notes) { dropNotes(); syncPresenter(); return }
      setNotes('error')
    })
  }

  // ---------------------------------------------------------------------
  // Presenter View's change feed
  // ---------------------------------------------------------------------

  let viewVersion = 0
  let viewKey = ''
  function syncPresenter(force = false): void {
    if (exited || !viewSubs.size) return
    const t = clock.now()
    const key = [slide, screen, paused, presenter, swapped ? 1 : 0, sink.list().length, content,
      timer ? Math.floor(timer.elapsed(t) / 1000) : 0, timer?.paused, Math.floor(remoteShowDeps.wall() / 60000)].join('|')
    if (!force && key === viewKey) return
    viewKey = key
    viewVersion++
    notify(viewSubs)
  }

  // ---------------------------------------------------------------------
  // Navigation
  // ---------------------------------------------------------------------

  const announce = () => {
    if (!live) return
    live.textContent = slide === 'end' ? END_OF_SHOW : `Slide ${slide + 1} of ${count}`
  }

  function enter(p: number | 'end'): void {
    slide = p
    screen = 'slide'
    if (p !== 'end') lastShown = p
    surface?.show({ slide: p, count, screen })
    if (p === 'end') surface?.root.removeAttribute('data-unit')
    else surface?.root.setAttribute('data-unit', String(unitOf[p]))
    want()
    draw()
    announce()
    if (started && p !== 'end') cb.onSlideChange?.(p, 0)
    followNotes()
    syncPresenter()
  }

  function next(): void {
    if (exited) return
    if (slide === 'end') { exit(); return }
    const after = nextVisible(order, slide)
    enter(after ?? (loop ? order[0] : undefined) ?? 'end')
  }
  function previous(): void {
    if (exited) return
    const to = previousVisible(order, slide === 'end' ? Infinity : slide)
    if (to !== undefined) enter(to)
  }
  function goTo(i: number): void {
    if (exited) return
    if (!Number.isInteger(i) || i < 0 || i >= count) {
      sink.add({ code: 'invalid-slide', message: `goTo(${i}) is outside the shown slides; ignored.` })
      return
    }
    enter(i)
  }
  const firstSlide = () => { if (count) enter(0) }
  const lastSlide = () => { if (count) enter(count - 1) }

  const release = () => {
    listeners.abort()
    stopManifest?.()
    stopTick?.()
    stopTick = undefined
    dropNotes()
    fetcher.dispose()
    view?.dispose()
    swapped?.dispose()
    placement?.dispose()
    viewSubs.clear()
    frameSubs.clear()
    presenterWindow?.close()
    cursorHider?.dispose()
    surface?.dispose()
    terminal.release()
  }

  function exit(): void {
    if (exited) return
    exited = true
    release()
    host.replaceChildren()
    start.cleanup()
    const page = unitOf[lastShown] ?? 0
    cb.scrollToSlide(page)
    cb.restoreFocus()
    // As the local session: focus is restored again once the browser has left full screen (bounded).
    if (!fullscreenSettled || start.token.owns(start.platform.fullscreenElement(owner))) {
      const done = new AbortController()
      const t = remoteShowDeps.timer(win)
      const id = t.set(() => done.abort(), RESTORE_LISTENER_MS)
      done.signal.addEventListener('abort', () => t.clear(id))
      owner.addEventListener('fullscreenchange', () => {
        if (start.platform.fullscreenElement(owner)) return
        done.abort()
        cb.scrollToSlide(page)
        cb.restoreFocus()
      }, { signal: done.signal })
      void start.fullscreen.then((outcome) => { if (outcome === 'refused') done.abort() })
    }
    try { cb.onEnd?.() } finally { resolveEnded() }
  }

  /** Single-screen (spec 5.1): Alt+F5 or the toggle swaps the surface between the audience and Presenter View. */
  function toggleView(): void {
    if (exited || !started || presenter === 'window' || !surface) return
    presenter = 'swap'
    if (swapped) {
      swapped.dispose()
      swapped = undefined
      surface.root.hidden = false
      if (!host.hasAttribute('tabindex')) host.tabIndex = -1
      host.focus()
    } else {
      const box = owner.createElement('div')
      box.setAttribute('data-officeview-presenter-swap', '')
      if (!policy.print) box.setAttribute('data-ov-print', 'false')
      host.append(box)
      const mounted = mountPresenterView(box, { session, host: 'swap', doc: owner })
      swapped = { dispose: () => { mounted.dispose(); box.remove() } }
      surface.root.hidden = true
    }
    startTick()
    followNotes()
    syncPresenter(true)
  }

  /** Presenter View's timer and wall clock move without a frame loop: a 1 Hz refresh while it can be shown. */
  function startTick(): void {
    if (stopTick || exited || presenter === 'none') return
    stopTick = remoteShowDeps.every(win, PRESENTER_TICK_MS, () => syncPresenter())
  }

  function dispatch(input: ShowInput): void { run(input, false) }

  /** Audience or presenter input, in the local session's precedence (kiosk gates the audience only). */
  function run(input: ShowInput, presenterInput: boolean): void {
    if (exited || !started) return
    if (input.kind === 'exit') { exit(); return }
    if (kiosk && !presenterInput) return
    if (input.kind === 'digit') { buffer += String(input.d); return }
    let target: ShowInput = input
    if (input.kind === 'enter' && buffer) target = { kind: 'goto', index: Number(buffer) - 1 }
    if (target.kind === 'goto') {
      buffer = ''
      screen = 'slide'
      surface?.show({ slide, count, screen })
      goTo(target.index)
      return
    }
    buffer = ''
    if (screen !== 'slide') {
      screen = 'slide'
      surface?.show({ slide, count, screen })
      syncPresenter()
      return
    }
    switch (target.kind) {
      // A finished bitmap has no triggers or links: a click advances.
      case 'click': case 'next': case 'enter': next(); break
      case 'previous': previous(); break
      case 'first': firstSlide(); break
      case 'last': lastSlide(); break
      case 'black': case 'white': screen = target.kind; surface?.show({ slide, count, screen }); syncPresenter(); break
      case 'pause': paused = !paused; syncPresenter(); break
    }
  }

  const policy = { copy: false, print: host.parentElement?.getAttribute('data-ov-print') !== 'false' }
  let presenterKey: (e: KeyboardEvent) => void = () => {}
  let surfaceKind: 'fullscreen' | 'in-window' = 'in-window'

  const session: SessionController = {
    get slideIndex() { return lastShown },
    buildIndex: 0,
    get surface() { return surfaceKind },
    get presenter() { return presenter },
    get state(): SessionState {
      return { slide, build: 0, screen, buffer, paused, clockPaused: paused, kiosk, loop, order, exited }
    },
    stage: undefined,
    cursor: undefined,
    next,
    previous,
    goTo,
    exit,
    dispatch,
    tick: () => syncPresenter(),
    activate: () => {},
    ended,
    get presenterSource() { return source },
  }

  const source: PresenterSource = {
    subscribe(f) {
      viewSubs.add(f)
      return () => { viewSubs.delete(f) }
    },
    get version() { return viewVersion },
    slideCount: count,
    // R18: never called for an image-backed source (slideImage is present).
    previewOf() { throw new Error('a streamed show has no slide model') },
    previewOptions: { showAnimation: false, reducedMotion: true, sink },
    notes: () => (doc.manifest.can.notes && notes ? notes : NO_NOTES),
    diagnostics: () => sink.list(),
    sink,
    get timer() { return timer ??= createElapsedTimer(clock.now()) },
    now: () => clock.now(),
    wall: () => remoteShowDeps.wall(),
    slideTime: () => 0,
    get workspace() { return workspace },
    ...(start.watermark ? { watermark: start.watermark } : {}),
    protection: policy,
    audienceFrame: () => surface?.canvas,
    onFrame(f) {
      frameSubs.add(f)
      return () => { frameSubs.delete(f) }
    },
    /** Cache-only (R18): the streamed slide's bitmap, fitted; never a fetch. */
    slideImage(i, canvas) {
      const unit = Number.isInteger(i) ? unitOf[i] : undefined
      const bmp = unit === undefined ? undefined : fetcher.cached(unit)
      return !!bmp && fitInto(canvas, bmp)
    },
    hidden: () => false,
    key: (e) => presenterKey(e),
    input: (i) => run(i, true),
    toggleView,
    get placement() { return placement },
    touch: () => syncPresenter(true),
  }

  // ---------------------------------------------------------------------
  // Surface, input and start
  // ---------------------------------------------------------------------

  const retryFullscreen: typeof start.retryFullscreen = (o) => start.retryFullscreen(o).then((r) => {
    if (r === 'entered' && !exited && start.isCurrent()) { surfaceKind = 'fullscreen'; closeRefusedNotice?.() }
    return r
  })
  let closeRefusedNotice: (() => void) | undefined
  try {
    surface = createAudienceSurface(host, backing, (e) => dispatch({ kind: 'click', x: e.clientX, y: e.clientY }))
    if (!policy.print) surface.root.setAttribute('data-ov-print', 'false')
    live = owner.createElement('div')
    live.setAttribute('data-officeview-show-live', '')
    live.setAttribute('aria-live', 'polite')
    live.style.cssText = HIDDEN
    frameText = owner.createElement('p')
    frameText.setAttribute('data-officeview-show-frame-text', '')
    frameText.setAttribute('role', 'status')
    frameText.style.cssText = HIDDEN
    surface.root.append(live, frameText)
    cursorHider = createCursorHider(surface.root, remoteShowDeps.timer(win))
    const on = { signal: listeners.signal }
    const gestures = createGestureState(() => clock.now())
    const keyInput = (e: KeyboardEvent, presenterInput = false) => {
      const input = keyToInput(e, buffer)
      if (input) {
        e.preventDefault()
        run(input, presenterInput)
      } else if (screen !== 'slide' && !MODIFIER_KEYS.has(e.key) && e.key !== 'Tab') {
        e.preventDefault()
        run({ kind: screen }, presenterInput)
      }
    }
    const swapKey = (e: KeyboardEvent) => {
      if (e.key !== 'F5' || !e.altKey || e.ctrlKey || e.metaKey) return false
      e.preventDefault()
      toggleView()
      return true
    }
    host.addEventListener('keydown', (e) => {
      if ((e.target as Element | null)?.closest?.('[data-officeview-presenter-view]')) { presenterKey(e); return }
      if (swapKey(e)) return
      if (e.target !== host && (e.target as Element | null)?.closest?.('button, input, textarea, select')) return
      keyInput(e)
    }, on)
    presenterKey = (e: KeyboardEvent) => {
      if (swapKey(e)) return
      const t = e.target as Element | null
      if (e.key !== 'Escape' && (t?.closest?.('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [data-officeview-presenter-notes]') || ((e.key === 'Enter' || e.key === ' ') && t?.closest?.('button')))) return
      keyInput(e, true)
    }
    const gesture = (e: Event) => {
      if ((e.target as Element | null)?.closest?.('[data-officeview-presenter-view]')) return
      const input = gestureToInput(e as WheelEvent | TouchEvent, gestures)
      if (e.type === 'wheel' && e.cancelable) e.preventDefault()
      if (input) dispatch(input)
    }
    for (const type of ['wheel', 'touchstart', 'touchend', 'touchcancel']) host.addEventListener(type, gesture, { ...on, passive: type !== 'wheel' })
    // Resize, full-screen change, DPR change: a new backing (when the reserve admits it) and its bucket.
    const resized = () => {
      if (exited || !surface) return
      const css = cssSize()
      if (!(css.w > 0 && css.h > 0)) return
      const next = backingFor(css, win?.devicePixelRatio ?? 1)
      if (next.w === backing.w && next.h === backing.h) return
      if (!terminal.resize(surfaceBytes(next.w, next.h))) {
        sink.add({ code: 'workspace', message: 'The new output size does not fit the slideshow workspace; the show keeps its size' })
        return
      }
      backing = next
      surface.canvas.width = next.w
      surface.canvas.height = next.h
      applyBucket()
      draw()
    }
    win?.addEventListener('resize', resized, on)
    owner.addEventListener('fullscreenchange', resized, on)
    const watchDpr = () => {
      const q = win?.matchMedia?.(`(resolution: ${win.devicePixelRatio ?? 1}dppx)`)
      q?.addEventListener?.('change', () => { resized(); watchDpr() }, { ...on, once: true })
    }
    watchDpr()
    // Revocation, expiry or any terminal end exits the show; a resume may turn notes off (or on).
    void doc.ended.then(() => { if (!exited) exit() })
    stopManifest = doc.onManifestChange(() => { if (!exited) followNotes() })

    // 'beginning' is the first shown slide; 'current' the viewer's page, or the next shown one after a hidden page.
    let from: number | 'end'
    if (start.from === 'current') {
      const page = cb.currentSlide()
      const k = unitOf.findIndex((u) => u >= page)
      from = k >= 0 ? k : count ? count - 1 : 'end'
    } else if (start.from === 'beginning') from = count ? 0 : 'end'
    else from = Number.isInteger(start.from) && start.from >= 0 && start.from < count ? start.from : count ? 0 : 'end'
    applyBucket()
    enter(from)
    if (!start.isCurrent()) throw new SlideShowError('cancelled')
    timer = createElapsedTimer(clock.now())
    const popup = start.popup
    if (popup === null || popup?.closed) {
      sink.add({ code: 'popup-blocked', message: 'The presenter window was blocked; presenting single-screen.' })
    } else if (popup) {
      try {
        const opener = win ?? popup.opener
        const pw = presenterWindow = adoptPresenterWindow(popup, opener, sink)
        const theme = host.parentElement?.getAttribute('data-ov-theme')
        if (theme) pw.root.setAttribute('data-ov-theme', theme)
        if (!policy.print) pw.root.setAttribute('data-ov-print', 'false')
        placement = createPresenterPlacement({
          popup, opener, host, retryFullscreen, fullscreen: start.fullscreen,
          prepared: preparedPlacement.get(owner), sink, changed: () => syncPresenter(true),
        })
        presenter = 'window'
        view = mountPresenterView(pw.root, { session, host: 'window', doc: pw.doc })
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
    if (presenter === 'swap') toggleView()
    // Presenter View's timer and wall clock move without a frame loop.
    startTick()
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
    const close = () => { closeRefusedNotice = undefined; notice.remove(); if (!exited) host.focus() }
    closeRefusedNotice = close
    notice.setAttribute('role', 'status')
    notice.setAttribute('data-officeview-show-notice', '')
    notice.textContent = 'Full screen is not available here; presenting in the window. '
    const retry = owner.createElement('button')
    retry.type = 'button'
    retry.textContent = 'Click for full screen'
    retry.onclick = () => { void retryFullscreen() }
    const dismiss = owner.createElement('button')
    dismiss.type = 'button'
    dismiss.textContent = 'Dismiss'
    dismiss.onclick = close
    notice.append(retry, dismiss)
    host.append(notice)
  })

  cb.onStart?.()
  if (slide !== 'end') cb.onSlideChange?.(slide, 0)
  return session
}
