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
 *
 * Piece 4 (Task 11): navigation is the shared show driver's (`createShowDriver`)
 * over the plan-backed source (`createPlanSource`). A deck with an animated
 * slide (`anim`, D5) on a session that serves show plans plays on `PlanStage`:
 * each animated slide is fetched as one bundle through the same paced fetcher
 * (+2 next, +1 previous), its composited steps' keyframe tracks just in time
 * (D4), its Morph pair on demand (the sequential predecessor's prefetched,
 * D11). A slide refused a bundle (`still`, `show-budget`, `show-limits`,
 * `show-plan-timeout`) plays its piece-3 still through the unit path; a failed
 * bundle shows the typed error frame. Any other deck keeps piece 3's still
 * stage: every slide is its final image, drawn as before. Links never fire
 * (`link-withheld`, §5.1); the show never renders locally.
 */
import { preparedPlacement, SlideShowError, type ShowStart, type SlideShowSession } from '../../components/present'
import { createDiagnosticSink } from '../../slideshow/diagnostics'
import { createCursorHider, type CursorHider, type CursorTimer } from '../../slideshow/player/cursor'
import { createShowDriver, type DriverHooks, type DriverStage, type ShowDriver, type SlideSource } from '../../slideshow/player/driver'
import { createGestureState, gestureToInput, keyToInput, MODIFIER_KEYS } from '../../slideshow/player/keys'
import { runLink } from '../../slideshow/player/links'
import type { SlideCursor } from '../../slideshow/engine/steps'
import { nextVisible } from '../../slideshow/player/order'
import type { SessionClock, SessionController, SessionState, ShowInput } from '../../slideshow/player/session'
import { createAudienceSurface, END_OF_SHOW, type AudienceSurface } from '../../slideshow/player/surface'
import type { ImageNotes } from '../../slideshow/presenter/notes'
import { mountPresenterView, type PresenterSource } from '../../slideshow/presenter/PresenterView'
import { createPresenterPlacement, type PresenterPlacement } from '../../slideshow/presenter/placement'
import { createElapsedTimer, type ElapsedTimer } from '../../slideshow/presenter/timer'
import { adoptPresenterWindow, type PresenterWindow } from '../../slideshow/presenter/window'
import type { DiagnosticSink, Lease, SlideShowWorkspace, TerminalReserve } from '../../slideshow/types'
import { admitShow, backingFor, createSlideShowWorkspace, surfaceBytes, type Backing, type SlideShowProfile } from '../../slideshow/workspace'
import { SHOW_PLAN_LIMITS } from '../show-plan'
import { StreamSessionError } from './errors'
import { createPlanSource } from './plan-source'
import { createPlanStage, type PlanSlideData, type PlanStage, type PlanTrack } from './plan-stage'
import { bucketFor, type RemoteOfficeDoc } from './remote-doc'
import { decodeBundle, decodeMorph, decodeTrack, type ShowBundle, type ShowMorph, type ShowTrack } from './show-bundle'
import { BACKOFF_MS, createShowFetcher, type ImageBitmapLike, type RemoteErrorCode, type ShowFetchClock, type ShowFetcher } from './show-fetch'

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
  /** Piece 4: the per-frame driver of an animated show (requestAnimationFrame); returns the stop. */
  frames: (win: Window | null, frame: () => void): (() => void) => {
    if (!win || typeof win.requestAnimationFrame !== 'function') return () => {}
    let on = true
    let id = 0
    const loop = () => { if (!on) return; frame(); id = win.requestAnimationFrame(loop) }
    id = win.requestAnimationFrame(loop)
    return () => { on = false; win.cancelAnimationFrame(id) }
  },
  /** Piece 4: decodes a show asset's PNG. */
  decode: (png: Uint8Array): Promise<ImageBitmapLike> => createImageBitmap(new Blob([png as BlobPart], { type: 'image/png' })),
  /** Piece 4: `'canvas2d'` starts PlanStage on Canvas 2D (tests, the forced path). */
  backend: undefined as 'canvas2d' | undefined,
  /** Piece 4: the per-show byte budget (spec §5: 128 MiB of plan records, assets and track frames). */
  get showBytes(): number { return SHOW_PLAN_LIMITS.showDecoded },
}

/** A show refusal that plays the piece-3 still (Task 10), or a transient one that is not a failed slide. */
const STILL_CODES: ReadonlySet<string> = new Set(['still', 'show-budget', 'show-limits', 'show-plan-timeout'])
const TRANSIENT: ReadonlySet<string> = new Set(['rate-limited', 'admission'])

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
  const motion: 'full' | 'reduced' = (start.motion === 'auto' ? !!win?.matchMedia?.('(prefers-reduced-motion: reduce)').matches : start.motion === 'reduced') ? 'reduced' : 'full'
  /** Piece 4: the session's show requests (absent: piece 3's still show). */
  const show = doc.session
  /** The deck plays on PlanStage when a shown slide is animated and the session serves show plans (D5). */
  const planned = !!show?.showPlan && unitOf.some((u) => !!list[u]!.anim)
  const animated = (unit: number) => planned && !!list[unit]?.anim

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
  /** The click checkpoint on screen (the driver's cursor step). */
  let build = 0
  let driver: ShowDriver | undefined
  let planStage: PlanStage | undefined
  let stopFrames: (() => void) | undefined
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
  /**
   * The show clock (fix round 1): what the driver and PlanStage play by — builds, transitions, tracks, Morph and the
   * after-timer advance. It stops while the show is paused; transport pacing and deadlines keep `clock`.
   */
  let pausedAt: number | undefined
  let pausedFor = 0
  const showClock: SessionClock = {
    now: () => (pausedAt ?? clock.now()) - pausedFor,
    get paused() { return pausedAt !== undefined },
    pause() { if (pausedAt === undefined) pausedAt = clock.now() },
    resume() { if (pausedAt === undefined) return; pausedFor += clock.now() - pausedAt; pausedAt = undefined },
  }
  const showBytes = remoteShowDeps.showBytes

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
    ...(planned ? { bundle: (unit: number, scale: number, dpr: number, m: 'full' | 'reduced', sig: AbortSignal, charge: () => void) => fetchBundle(unit, scale * dpr, m, sig, charge) } : {}),
    // R28: the session's one unit send budget (it outlives this show and stamps every charged send).
    ...(doc.unitBudget ? { budget: doc.unitBudget } : {}),
  }, clock, { maxCached, animated, motion, showBytes, extraBytes: () => extraBytes() })
  /** The show scale of a bucket: the server takes 1 or 2 only (Task 10). */
  const showScale = (bucket: number): 1 | 2 => (bucket >= 2 && doc.manifest.recorder.scalesBuckets.includes(2) ? 2 : 1)
  /** The scale each bundle was fetched at (its tracks and Morph pairs are asked at it). */
  const scaleOf = new WeakMap<ShowBundle, 1 | 2>()
  /** One slide bundle (D1): the plan, the entry frame, the layers, decoded and charged to the workspace. */
  async function fetchBundle(unit: number, bucket: number, m: 'full' | 'reduced', sig: AbortSignal, charge: () => void): Promise<ShowBundle> {
    open++
    try {
      const scale = showScale(bucket)
      const reply = await show!.showPlan!(unit, scale, m, sig, { workspace, onSend: charge })
      // The fetch's signal owns the decode too (finding 2): an abort mid-decode gives everything back at once.
      const b = await decodeBundle(unit, reply, workspace, remoteShowDeps.decode, sig)
      if (sig.aborted || exited) { b.release(); throw new StreamSessionError('closed', 'the bundle is no longer wanted') }
      const p = pinnable(b)
      scaleOf.set(p, scale)
      return p
    } finally {
      settledOne()
    }
  }
  /**
   * A bundle the fetcher may evict or replace (a bucket change) while the slide it belongs to is on screen: its
   * release waits until the slide's lease lets go of it (pinned), so the stage never loses the bitmaps it draws.
   */
  const pins = new WeakMap<ShowBundle, { n: number; dead: boolean; release(): void }>()
  /** Bytes of bundles the fetcher let go of that a slide still pins: still held, so still in the show budget. */
  let pinnedBytes = 0
  function pinnable(b: ShowBundle): ShowBundle {
    const p: ShowBundle = {
      ...b,
      release() {
        const h = pins.get(p)!
        if (h.dead) return
        h.dead = true
        if (!h.n) h.release()
        else pinnedBytes += b.decodedBytes
      },
    }
    pins.set(p, { n: 0, dead: false, release: () => b.release() })
    return p
  }
  const pin = (b: ShowBundle) => { const h = pins.get(b); if (h) h.n++ }
  const unpin = (b: ShowBundle) => {
    const h = pins.get(b)
    if (!h) return
    h.n = Math.max(0, h.n - 1)
    if (!h.n && h.dead) { pinnedBytes -= b.decodedBytes; h.release() }
  }
  /** Streamed index → the bundle the slide's live lease pins (what the stage draws). */
  const held = new Map<number, ShowBundle>()
  const bundleFor = (i: number): ShowBundle | undefined => held.get(i) ?? (unitOf[i] !== undefined && animated(unitOf[i]!) ? fetcher.bundle(unitOf[i]!) : undefined)
  /** Animated units refused a bundle with a still code: they play their piece-3 still. */
  const stillOnly = new Set<number>()

  // ---------------------------------------------------------------------
  // Keyframe tracks (D4) and Morph pairs (D11): held for the slide on screen; their bytes count against the
  // show budget (the fetcher evicts bundles for them). Each is one charged send, asked only with budget room.
  // ---------------------------------------------------------------------

  interface TrackSlot { index: number; ctl: AbortController; state: 'pending' | ShowTrack | 'failed' }
  const tracks = new Map<string, TrackSlot>()
  interface PairSlot { key: string; from: number; to: number; fromStep: number; ctl: AbortController; result?: ShowMorph; promise: Promise<ShowMorph | undefined> }
  const pairs = new Map<string, PairSlot>()
  /** Decoded keyframe-track frames held by the D23 windows (step tracks and composited Morph pairs). */
  let frameBytes = 0
  /**
   * The show's byte owner for decoded track frames (fix round 1): each frame entering a window is admitted against
   * the show budget — records, assets, encoded tracks, Morph pairs and every frame held — after the least recently
   * used non-current bundles are evicted for it; no room refuses the frame (the window holds the previous picture,
   * D23). Leases of any other kind go to the workspace as they are.
   */
  const frameOwner: SlideShowWorkspace = {
    get limit() { return workspace.limit },
    used: () => workspace.used(),
    reserveTerminal: (bytes) => workspace.reserveTerminal(bytes),
    lease(bytes, tag) {
      if (tag !== 'plan:track-frame') return workspace.lease(bytes, tag)
      if (exited || !fetcher.makeRoom(bytes)) return undefined
      const l = workspace.lease(bytes, tag)
      if (!l) return undefined
      frameBytes += bytes
      let held = bytes
      return {
        get tag() { return l.tag },
        get bytes() { return l.bytes },
        grow: () => false,
        shrink(n: number) { const k = Math.min(Math.max(0, n), held); held -= k; frameBytes -= k; l.shrink(k) },
        transfer: (t: string) => l.transfer(t),
        release() { frameBytes -= held; held = 0; l.release() },
      }
    },
  }
  function extraBytes(): number {
    let n = frameBytes + pinnedBytes
    for (const t of tracks.values()) if (typeof t.state === 'object') n += t.state.bytes
    for (const p of pairs.values()) n += p.result?.bytes ?? 0
    return n
  }
  const overBudget = () => fetcher.bundleBytes() + extraBytes() > showBytes
  function dropTrack(key: string): void {
    const t = tracks.get(key)
    if (!t) return
    tracks.delete(key)
    t.ctl.abort()
    if (typeof t.state === 'object') t.state.dispose()
  }
  function dropPair(p: PairSlot): void {
    if (pairs.get(p.key) === p) pairs.delete(p.key)
    p.ctl.abort()
    p.result?.release()
    p.result = undefined
  }
  /** A slide change: the tracks of every other slide go, and every Morph pair not into `p`. */
  function dropExtras(p: number | 'end' | undefined): void {
    for (const [k, t] of [...tracks]) if (t.index !== p) dropTrack(k)
    for (const q of [...pairs.values()]) if (q.to !== p) dropPair(q)
  }
  const codeOf = (e: unknown) => (e instanceof StreamSessionError ? e.code : undefined)

  /** The keyframe track of composited step `k` of streamed slide `i` (on screen), when the budget has room (else `step-deferred`). */
  function requestTrack(i: number, k: number): void {
    const key = `${i}:${k}`
    if (exited || tracks.has(key) || !show?.showTrack) return
    const b = held.get(i)
    const step = b?.plan.steps[k]
    if (!b || !step || step.eligibility !== 'composited' || !step.track || !fetcher.room('track')) return
    const unit = unitOf[i]!
    const count = step.track.count
    const slot: TrackSlot = { index: i, ctl: new AbortController(), state: 'pending' }
    tracks.set(key, slot)
    const live = () => tracks.get(key) === slot && !exited
    show.showTrack(unit, k, scaleOf.get(b) ?? 1, motion, slot.ctl.signal, { workspace }).then((reply) => {
      if (!live()) { reply.release(); return }
      try {
        const t = decodeTrack(unit, count, reply, frameOwner, remoteShowDeps.decode)
        if (!live()) { t.dispose(); return }
        slot.state = t
        if (overBudget()) {
          t.dispose()
          slot.state = 'failed'
          sink.add({ code: 'show-budget', slideIndex: i, effect: `step:${k}`, message: `The keyframe track of step ${k + 1} does not fit the show budget; the step cuts to its end` })
        }
      } catch (e) {
        slot.state = 'failed'
        sink.add({ code: 'show-track-refused', slideIndex: i, effect: `step:${k}`, message: `The keyframe track of step ${k + 1} was refused (${codeOf(e) ?? (e as Error).message})` })
      }
    }, (e: unknown) => {
      if (!live()) return
      slot.state = 'failed'
      if (codeOf(e) === 'rate-limited') fetcher.paused()
    })
  }
  /** D4: on entering a slide, the tracks of its first two composited steps (from the step entered at). */
  function tracksOnEntry(i: number, at: number): void {
    const b = held.get(i)
    if (!b) return
    let n = 0
    for (let k = Math.max(0, at - 1); k < b.plan.steps.length && n < 2; k++) {
      if (b.plan.steps[k]!.eligibility !== 'composited') continue
      requestTrack(i, k)
      n++
    }
  }

  /** The Morph pair `(from → to)` left at `fromStep` (streamed indices): asked once, kept until the show leaves `to`. */
  function requestPair(from: number, to: number, fromStep: number): PairSlot | undefined {
    const key = `${from}>${to}@${fromStep}`
    const have = pairs.get(key)
    if (have) return have
    // D13: under reduced motion a Morph plays as a fade; nothing is asked (and nothing charged).
    if (exited || !show?.showMorph || motion === 'reduced' || !fetcher.room('morph')) return undefined
    const uf = unitOf[from], ut = unitOf[to]
    if (uf === undefined || ut === undefined) return undefined
    const target = bundleFor(to)
    const scale = (target && scaleOf.get(target)) ?? 1
    const slot = { key, from, to, fromStep, ctl: new AbortController() } as PairSlot
    const live = () => pairs.get(key) === slot && !exited
    slot.promise = (async () => {
      try {
        const reply = await show.showMorph!(uf, ut, fromStep, scale, motion, slot.ctl.signal, { workspace })
        if (!live()) { reply.release(); return undefined }
        // The pair's controller owns the decode too (finding 2): `dropPair` and the end give everything back at once.
        const m = await decodeMorph({ from: uf, to: ut, fromStep }, reply, frameOwner, remoteShowDeps.decode, slot.ctl.signal)
        if (!live()) { m.release(); return undefined }
        slot.result = m
        if (overBudget()) {
          dropPair(slot)
          sink.add({ code: 'show-budget', slideIndex: to, effect: 'morph', message: 'The Morph pair does not fit the show budget' })
          return undefined
        }
        return m
      } catch (e) {
        if (pairs.get(key) === slot) pairs.delete(key)
        const code = codeOf(e)
        // The server's own fade (D11): play it.
        if (code === 'morph-fallback') return undefined
        if (code === 'rate-limited') fetcher.paused()
        throw e
      }
    })()
    // A pair nobody waits for (a prefetch dropped) must not reject unhandled.
    slot.promise.catch(() => {})
    pairs.set(key, slot)
    return slot
  }
  /** PlanStage's Morph request: a pair asked (or prefetched) at another step is dropped and this one asked for. */
  function morphFor(from: number, to: number, fromStep: number): Promise<{ record: ShowMorph['record']; assets: ShowMorph['assets']; track?: PlanTrack } | undefined> {
    for (const p of [...pairs.values()]) if (p.to === to && (p.from !== from || p.fromStep !== fromStep)) dropPair(p)
    const slot = requestPair(from, to, fromStep)
    // No room, or no Morph service: the fade plays at the deadline (`morph-late`).
    if (!slot) return Promise.reject(new StreamSessionError('rate-limited', 'no room for the Morph pair'))
    return slot.promise.then((m) => m && { record: m.record, assets: m.assets, ...(m.track ? { track: m.track } : {}) })
  }
  /** D11: the sequential successor's pair is prefetched with it, at the step this slide is left at by Next (its last). */
  function prefetchMorph(): void {
    if (exited || !planned || motion === 'reduced' || slide === 'end') return
    const v = driver?.visit
    if (!v || v.index !== slide || !v.timeline) return
    const q = nextVisible(order, slide) ?? (loop ? order[0] : undefined)
    if (q === undefined || q === slide) return
    const b = fetcher.bundle(unitOf[q]!)
    if (b?.plan.show.transition?.element !== 'morph') return
    requestPair(slide, q, v.timeline.steps)
  }

  /** PlanStage's view of streamed slide `i`: its pinned bundle (or the cached one), its tracks. */
  function planData(i: number): PlanSlideData | undefined {
    const b = bundleFor(i)
    if (!b) return undefined
    return {
      plan: b.plan, assets: b.assets,
      track: (k) => { const t = tracks.get(`${i}:${k}`); return !t || t.state === 'failed' ? undefined : t.state },
    }
  }

  /** Acquires waiting for their animated unit to settle (a bundle, a still refusal or a failure). */
  const waiters = new Map<number, Set<() => void>>()
  const settledFor = (unit: number) => !animated(unit) || stillOnly.has(unit) || failed.has(unit) || !!fetcher.bundle(unit)
  function settle(unit: number, sig: AbortSignal): Promise<void> {
    if (settledFor(unit)) return Promise.resolve()
    return new Promise<void>((resolve, reject) => {
      let set = waiters.get(unit)
      if (!set) waiters.set(unit, set = new Set())
      const s0 = set
      const done = () => { s0.delete(done); sig.removeEventListener('abort', abort); resolve() }
      const abort = () => { s0.delete(done); reject(sig.reason ?? new Error('aborted')) }
      s0.add(done)
      sig.addEventListener('abort', abort, { once: true })
    })
  }
  const wake = (unit: number) => { const set = waiters.get(unit); if (set && settledFor(unit)) for (const f of [...set]) f() }

  /** When each failed unit failed: a later visit past the fetcher's back-off re-asks for it (an animated slide then waits for it). */
  const failedAt = new Map<number, number>()
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

  /** The stage box (PlanStage's compositor canvases), hidden while a text frame is shown on the output canvas. */
  const stageBox = () => surface?.root.querySelector<HTMLElement>('[data-officeview-show-stage]') ?? undefined

  /**
   * One audience frame. On PlanStage the driver ticks (the stage draws); a slide with nothing to show yet gets the
   * piece-3 loading or typed error frame over it. On the still stage the slide's bitmap is drawn as piece 3 drew it.
   * `tick: false` from inside the driver's own calls (a slide change, a slide becoming drawable): the next frame
   * draws the stage, so the driver is never re-entered.
   */
  function render(tick = true): void {
    if (exited || !surface) return
    const canvas = surface.canvas
    let kind: 'image' | 'loading' | 'error' | 'end' = 'end'
    let text = ''
    let bmp: ImageBitmapLike | undefined
    if (slide !== 'end') {
      const unit = unitOf[slide]!
      bmp = planStage ? undefined : fetcher.cached(unit)
      const code = failed.get(unit)
      // The still stage draws the bitmap here (piece 3); PlanStage draws on its tick.
      if (planStage ? !!held.get(slide) || !!fetcher.cached(unit) : !!bmp && fitInto(canvas, bmp)) kind = 'image'
      else {
        kind = code ? 'error' : 'loading'
        text = code ? `Slide ${slide + 1} could not be shown (${code})` : `Loading slide ${slide + 1}`
      }
    }
    if (planStage) {
      // The output canvas under the stage's compositor: black again once a text frame has gone (the letterbox).
      if (kind !== 'loading' && kind !== 'error' && textShown) {
        const ctx = canvas.getContext('2d')
        if (ctx) { ctx.save(); ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.fillStyle = '#000'; ctx.fillRect(0, 0, canvas.width, canvas.height); ctx.restore() }
        textShown = false
      }
      if (tick) driver?.tick()
      if (exited) return
      const box = stageBox()
      if (box) box.hidden = kind === 'loading' || kind === 'error'
    }
    if (kind === 'loading' || kind === 'error') { textFrame(canvas, text); textShown = true }
    surface.root.setAttribute('data-officeview-show-frame', kind)
    if (frameText && frameText.textContent !== text) frameText.textContent = text
    afterFrame()
    // A build settling ends what Next is (the last build's settle point leaves the deck) with no other event.
    if (planned && !exited) syncPresenter()
  }
  /** A loading or error frame is on the output canvas (PlanStage clears it once the slide shows). */
  let textShown = false

  /** Each subscriber still subscribed, stopping once one of them ends the show. */
  function notify(set: Set<() => void>, onThrow?: (f: () => void, e: unknown) => void): void {
    for (const f of [...set]) {
      if (exited) return
      if (!set.has(f)) continue
      if (!onThrow) { f(); continue }
      try { f() } catch (e) { onThrow(f, e) }
    }
  }

  /** A blank screen (or its end) is a new audience frame: Presenter View's Current
   * preview repaints from frame notifications, as after the local session's render(). */
  function setScreen(next: 'slide' | 'black' | 'white'): void {
    screen = next
    surface?.show({ slide, count, screen })
    afterFrame()
    syncPresenter()
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
    wake(unit)
    if (slide !== 'end' && unitOf[slide] === unit) render()
    prefetchMorph()
    syncPresenter()
  })
  fetcher.onError((unit, code) => {
    if (exited) return
    // A rate refusal pauses the fetcher, which re-issues the unit (as `admission` re-issues it): not a failed slide.
    if (TRANSIENT.has(code)) return
    if (STILL_CODES.has(code) && animated(unit)) {
      // The fetcher re-asks for the piece-3 still; the slide plays it.
      stillOnly.add(unit)
      const i = unitOf.indexOf(unit)
      sink.add({ code, ...(i >= 0 ? { slideIndex: i } : {}), message: `Slide ${i + 1} plays as its still image (${code})` })
      wake(unit)
      return
    }
    failed.set(unit, code)
    failedAt.set(unit, clock.now())
    content++
    wake(unit)
    if (slide !== 'end' && unitOf[slide] === unit) render()
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
    // The build and the cursor's history move Next (the next checkpoint), so they are part of the key.
    const key = [slide, build, driver?.visit?.cursor?.events.length ?? -1, planned && nextStep() ? 1 : 0, screen, paused, presenter, swapped ? 1 : 0, sink.list().length, content,
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

  /** A slide change (the driver's, before the slide's model is in): what piece 3's enter did, but navigation. */
  function entered(p: number | 'end'): void {
    slide = p
    screen = 'slide'
    build = 0
    if (p !== 'end') lastShown = p
    surface?.show({ slide: p, count, screen })
    if (p === 'end') surface?.root.removeAttribute('data-unit')
    else surface?.root.setAttribute('data-unit', String(unitOf[p]))
    dropExtras(p)
    // An animated slide whose bundle failed long enough ago is asked for again by `want`: its visit waits for it
    // (shown loading) rather than playing static. A still unit keeps piece 3's error frame until its bitmap lands.
    const unit = p === 'end' ? undefined : unitOf[p]
    if (unit !== undefined && animated(unit) && failed.has(unit) && clock.now() - (failedAt.get(unit) ?? 0) >= BACKOFF_MS) { failed.delete(unit); failedAt.delete(unit) }
    want()
    render(false)
    announce()
    followNotes()
    syncPresenter()
  }

  function next(): void {
    if (exited) return
    driver?.next()
  }
  function previous(): void {
    if (exited) return
    driver?.previous()
  }
  function goTo(i: number): void {
    if (exited) return
    if (!Number.isInteger(i) || i < 0 || i >= count) {
      sink.add({ code: 'invalid-slide', message: `goTo(${i}) is outside the shown slides; ignored.` })
      return
    }
    driver?.goTo(i)
  }
  const firstSlide = () => { if (!exited) driver?.first() }
  const lastSlide = () => { if (!exited) driver?.last() }

  const release = () => {
    listeners.abort()
    stopManifest?.()
    stopTick?.()
    stopTick = undefined
    stopFrames?.()
    stopFrames = undefined
    askingAgain?.(); askingAgain = undefined
    dropNotes()
    // The driver lets go of the slide on screen (its pinned bundle), the stage of everything it uploaded.
    driver?.dispose()
    planStage?.dispose()
    dropExtras(undefined)
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
      // Restore first, notifying on its own: an invalid index navigates nowhere (and draws nothing).
      if (screen !== 'slide') setScreen('slide')
      goTo(target.index)
      return
    }
    buffer = ''
    if (screen !== 'slide') { setScreen('slide'); return }
    switch (target.kind) {
      // A click goes through the driver: a plan's trigger or link region takes it, else it advances.
      case 'click': driver?.click(target.x, target.y); break
      case 'next': case 'enter': next(); break
      case 'previous': previous(); break
      case 'first': firstSlide(); break
      case 'last': lastSlide(); break
      case 'black': case 'white': setScreen(target.kind); break
      // The show clock stops with the show: builds, transitions, tracks and the after-timer advance hold.
      case 'pause': paused = !paused; if (paused) showClock.pause(); else showClock.resume(); syncPresenter(); break
    }
  }

  const policy = { copy: false, print: host.parentElement?.getAttribute('data-ov-print') !== 'false' }
  let presenterKey: (e: KeyboardEvent) => void = () => {}
  let surfaceKind: 'fullscreen' | 'in-window' = 'in-window'

  const session: SessionController = {
    get slideIndex() { return lastShown },
    get buildIndex() { return build },
    get surface() { return surfaceKind },
    get presenter() { return presenter },
    get state(): SessionState {
      return { slide, build, screen, buffer, paused, clockPaused: paused, kiosk, loop, order, exited }
    },
    stage: undefined,
    get cursor() { return driver?.visit?.cursor },
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

  /**
   * The checkpoint the next click settles on, `undefined` when it leaves the slide (the local `nextCheckpoint`'s
   * rule, on the visit's own timeline and the stage's slide time).
   */
  function nextStep(): { step: number; cursor: SlideCursor; bundle: ShowBundle } | undefined {
    const v = driver?.visit
    const b = slide === 'end' ? undefined : held.get(slide)
    if (!planStage || !b || !v || v.index !== slide || !v.cursor || !v.timeline) return undefined
    const tl = v.timeline, c = v.cursor
    try {
      const t = Math.max(0, planStage.slideTime(showClock.now()))
      if (c.step >= tl.steps && c.mode === 'live' && tl.busyUntil(c) <= t) return undefined
      const last = c.events[c.events.length - 1]
      const after = tl.next(c, Math.max(last ? last.t : 0, t))
      return after === c ? undefined : { step: after.step, cursor: tl.checkpoint(after.step), bundle: b }
    } catch {
      // A history the timeline cannot replay (a text build's fanned history): the slide advance.
      return undefined
    }
  }

  /** A frame asked of a track window lands later: Presenter View is told, once, to paint Next again. */
  let askingAgain: (() => void) | undefined
  function askAgain(): void {
    if (askingAgain) return
    askingAgain = clock.wake(50, () => { askingAgain = undefined; if (exited) return; content++; syncPresenter() })
  }

  /** Presenter View's Next for an animated deck (§5.2); `false` when nothing can be shown yet. */
  function nextPreview(canvas: HTMLCanvasElement): boolean {
    if (exited || slide === 'end' || !planStage) return false
    const after = nextVisible(order, slide) ?? (loop ? order[0] : undefined)
    const entry = () => after !== undefined && !!source.slideImage?.(after, canvas)
    const n = nextStep()
    if (!n) return entry()
    // A composited step's track not here yet is the budget's business (`step-deferred`), not a refused lease: the
    // entry frame stands in, and the next change of the view asks again.
    if (n.bundle.plan.steps[n.step - 1]?.eligibility === 'composited') {
      const track = planData(slide)?.track(n.step - 1)
      if (track === undefined || track === 'pending') return entry()
      if (!track.frame(track.count - 1)) {
        // The step's end frame is not in the track's decode window. Another step's track is not playing, so its
        // window can move there; the one the audience plays (the settle point) is left alone.
        if (n.step !== driver?.visit?.cursor?.step) { track.want(track.count - 1); askAgain() }
        return entry()
      }
    }
    let ok = false
    try { ok = planStage.composeAt(slide, n.cursor, canvas) } catch { ok = false }
    if (ok) return true
    sink.add({ code: 'preview-reduced', slideIndex: slide, effect: `step:${n.step}`, message: `Next could not be composed for slide ${slide + 1} (no room in the show workspace); it shows the entry frame` })
    return entry()
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
    // PlanStage's compositor canvas while it is shown, else the output canvas.
    audienceFrame: () => {
      const box = planStage ? stageBox() : undefined
      const gl = box && !box.hidden ? [...box.querySelectorAll<HTMLCanvasElement>('canvas')].pop() : undefined
      return gl ?? surface?.canvas
    },
    onFrame(f) {
      frameSubs.add(f)
      return () => { frameSubs.delete(f) }
    },
    /** Cache-only (R18): the streamed slide's bitmap (an animated slide's entry frame, D15), fitted; never a fetch. */
    slideImage(i, canvas) {
      const unit = Number.isInteger(i) ? unitOf[i] : undefined
      const b = unit === undefined ? undefined : bundleFor(i)
      const bmp = b ? b.assets.bitmap(b.plan.entry) : unit === undefined ? undefined : fetcher.cached(unit)
      return !!bmp && fitInto(canvas, bmp)
    },
    // Piece 4 §5.2: only an animated deck composes Next; an unanimated one keeps the piece-3 still path.
    ...(planned ? { nextImage: (canvas: HTMLCanvasElement) => nextPreview(canvas), nextBuild: () => !exited && !!nextStep() } : {}),
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
    // The driver over the plan-backed source, on PlanStage (an animated deck) or piece 3's still stage.
    const planSource = createPlanSource({ manifest: doc.manifest, unitOf, bundles: bundleFor, loop, kiosk })
    const slides: SlideSource = {
      count,
      settings: planSource.settings,
      hidden: planSource.hidden,
      async acquire(i, sig) {
        const unit = unitOf[i]!
        await settle(unit, sig)
        // The bundle is pinned before the source reads it, for as long as the slide holds its lease.
        const b = animated(unit) && !stillOnly.has(unit) ? fetcher.bundle(unit) : undefined
        if (b) { pin(b); held.set(i, b) }
        const unhold = () => { if (b) { if (held.get(i) === b) held.delete(i); unpin(b) } }
        let h: Awaited<ReturnType<SlideSource['acquire']>>
        try { h = await planSource.acquire(i, sig) } catch (e) { unhold(); throw e }
        // The text fan (S2-8, D22) is charged to the show workspace for the visit's life.
        const charged: Lease[] = []
        let done = false
        return {
          slide: h.slide,
          charge(bytes: number, tag: string) {
            if (done) return false
            const l = workspace.lease(bytes, tag)
            if (!l) return false
            charged.push(l)
            return true
          },
          release() {
            if (done) return
            done = true
            h.release()
            for (const l of charged.splice(0)) l.release()
            unhold()
          },
        }
      },
    }
    if (planned) {
      planStage = createPlanStage({
        root: surface.root, output: surface.canvas, workspace, sink, reduced: motion === 'reduced',
        showNow: () => showClock.now(),
        data: planData,
        still: (i) => (unitOf[i] === undefined ? undefined : fetcher.cached(unitOf[i]!)),
        morph: morphFor,
        ...(remoteShowDeps.backend ? { force: remoteShowDeps.backend } : {}),
      })
    }
    /** Piece 3's still stage: the slide's bitmap is drawn by `render`; nothing animates. */
    const stillStage: DriverStage = {
      prepare() {}, async enter() {}, frame: () => undefined, flattenCurrent() {}, leave() {},
      hitTest: () => undefined, slideTime: (t) => t,
    }
    const hooks: DriverHooks = {
      onSlideChange(i, b) {
        if (exited) return
        build = b
        if (started) cb.onSlideChange?.(i, b)
        // D4: starting step k asks for step k + 1's track (and k's, if it is still missing).
        if (b > 0 && driver?.visit?.index === i) { requestTrack(i, b - 1); requestTrack(i, b) }
        syncPresenter()
      },
      onVisit(v) {
        if (exited) return
        if (!v) { entered(driver?.slide ?? 'end'); return }
        if (v.cursor) tracksOnEntry(v.index, v.cursor.step)
        prefetchMorph()
        render(false)
        // The visit's cursor now exists: what Next is (a build or the next slide) can be told.
        syncPresenter()
      },
      onCursor() { syncPresenter() },
      runLink(link, v) {
        runLink(link, { goTo, next, previous, first: firstSlide, last: lastSlide, exit }, win ?? (globalThis as unknown as Window), { sink, slideIndex: v.index })
      },
      onEnd() { exit() },
      // A blanked screen holds the after-timer auto-advance.
      held: () => screen !== 'slide',
    }
    driver = createShowDriver({ source: slides, stage: planStage ?? stillStage, clock: showClock, sink, reducedMotion: motion === 'reduced', kiosk, hooks })
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
      if (planStage) planStage.resize({ w: next.w, h: next.h }, 1)
      else {
        surface.canvas.width = next.w
        surface.canvas.height = next.h
      }
      applyBucket()
      render()
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
    void driver.start(from)
    if (!start.isCurrent()) throw new SlideShowError('cancelled')
    // An animated deck draws every frame (transitions, builds, tracks, Morph).
    if (planStage) stopFrames = remoteShowDeps.frames(win, () => render())
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
