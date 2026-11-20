/**
 * Piece 5 Task 11 (spec §4.4, §5; plan D1, D4, D12, D16, D20, D22): a streamed show's audio. Lazy: imported by the
 * remote show (show.ts) only.
 *
 * - **The player.** The unchanged media player (`createMediaPlayer`) on the audience surface, `clocked` (§5.2: late
 *   bytes join the node at its clock offset), and the transition sound (`createTransitionSound`) over the same
 *   opener, on the show clock (late bytes join at the elapsed offset; a finished sound never plays). Audio plays only in the audience window (Presenter View mirrors no element, as locally). The driver
 *   never takes an end from it: a streamed slide's audio ends on the show clock (the source's `mediaClock`), so a
 *   refusal or failure after the plan was sent raises the player's typed notice and keeps the clock end (D4).
 * - **Video (D22).** The player is never handed a video node: `enter` gives it the slide without video nodes and
 *   without blanked commands (`cmd: ''`, which the server ships for every command it does not play), so it opens
 *   nothing, raises no notice and creates no element for them. The driver ends a video node as piece 4 does.
 * - **The opener.** `open(rec)` maps a plan-source record (D16) to its server unit and plan record: a record the
 *   plan ships refused answers its typed reason with nothing sent (§5.3 then ends it as the local fallback does);
 *   otherwise `openRemoteMediaLease` over the session, `unit` always sent. Requests wait for the show's
 *   `media-show` answer. A `session-rotated` refusal (a drop: the server's implicit media-end) waits for the resume,
 *   then the request is sent again.
 * - **Request timing (§4.4).** A slide's media is requested when it becomes the next or the current slide
 *   (`want`), the current slide's first; the session paces the sends (its `media` log: 6 per 60 s, look-ahead at most
 *   4 of them, Task 10 ruling 3). A delivery asked for ahead is handed to the player (or the transition sound) that
 *   opens the record on that slide, and is asked for once per stay in the window. One that leaves the window unused
 *   is released (aborted while pending). Under room pressure delivered, unused look-ahead goes first (`evict`).
 * - **The transport (D1).** `start(i)` sends `media-show` for the starting slide, `moved(i)` a `media-cursor` on each
 *   slide change, `end()` the `media-end` and releases every lease. After a resume `media-show` and the cursor are
 *   sent again for the slide on screen (the session does not).
 */
import { createMediaPlayer, type MediaPlayer } from '../../slideshow/media/player'
import { createTransitionSound, type TransitionSoundController } from '../../slideshow/media/sounds'
import type { MediaLeaseResult } from '../../slideshow/media/lease'
import type { DriverMedia } from '../../slideshow/player/driver'
import type { DiagnosticSink, MediaRecord, ShowSlide, SlideShowWorkspace, TimeNode } from '../../slideshow/types'
import type { PlanMediaRecord } from '../show-plan'
import { openRemoteMediaLease } from './media-lease'
import { remoteMediaRecord, remoteRecordOf } from './plan-source'
import type { MediaShowAnswer, StreamSession } from './session'

export type ShowMediaTransport = Pick<StreamSession, 'mediaShow' | 'mediaCursor' | 'mediaRequest' | 'mediaEnd'>

export interface ShowMediaOptions {
  transport: ShowMediaTransport
  /** The audience surface root: every media element lives in its document (the audience window). */
  root: HTMLElement
  win: Window
  workspace: SlideShowWorkspace
  sink: DiagnosticSink
  /** Streamed index → server unit. */
  unitOf: readonly number[]
  /** The plan audio records of streamed slide `i`, when its bundle is in (undefined: none yet, or no plan). */
  records(i: number): readonly PlanMediaRecord[] | undefined
  /** The shown visit's slide, when `i` is on screen (the transition sound's metadata). */
  slide(i: number): ShowSlide | undefined
  /** Calls `cb` after a resume replaced the session's manifest; returns the unsubscribe. */
  onResume(cb: () => void): () => void
  /** The show clock (the driver's; it stops while the show is paused): a late transition sound joins on it (F3). */
  now(): number
}

export interface ShowMedia {
  /** The driver's `hooks.media`: the player's controller over the slide without video nodes or blanked commands. */
  enter(slide: ShowSlide, now: () => number): DriverMedia
  readonly sounds: TransitionSoundController
  readonly player: MediaPlayer
  /** `media-show` for the starting streamed slide. */
  start(i: number): void
  /** The slide on screen changed (`'end'`: the end-of-show slide): `media-cursor`, then the request window. */
  moved(i: number | 'end', next: number | undefined): void
  /** Re-checks the request window (a bundle came in). */
  want(): void
  setPaused(paused: boolean): void
  /** `media-end`; every pending request aborted, every lease released. Idempotent. */
  end(): void
}

const MAX_DEPTH = 256
/**
 * Refusals of a delivery asked for ahead (for the next slide) that an open on the slide, once it is current, asks
 * again for once: the session's look-ahead share (`media-busy` at the 256 MiB look-ahead, `rate-limited`), not the
 * media's. A refusal of a request made for the current slide is final (D4: the notice, and the clock end).
 */
const RETRY_AHEAD: ReadonlySet<string> = new Set(['media-busy', 'rate-limited', 'resource-limit'])
/** A drop is retried after the resume, at most this many times per request. */
const ROTATIONS = 3

interface Entry {
  key: string
  unit: number
  ctl: AbortController
  promise: Promise<MediaLeaseResult>
  result?: MediaLeaseResult
  taken: boolean
  /** Asked for while its slide was the next one, not the current. */
  ahead: boolean
}

const aborted = (signal: AbortSignal) => signal.reason ?? Object.assign(new Error('The media request was aborted.'), { name: 'AbortError' })

/** The slide as the player sees it: no video node (D22) and no blanked command. Leaves only; ids unchanged. */
function audioOnly(slide: ShowSlide): ShowSlide {
  const timing = slide.timing
  if (!timing || 'refused' in timing) return slide
  let changed = false
  const walk = (n: TimeNode, depth: number): TimeNode | undefined => {
    if (!('children' in n)) {
      if (n.kind === 'video' || (n.kind === 'cmd' && n.cmd === '')) { changed = true; return undefined }
      return n
    }
    if (depth > MAX_DEPTH) return n
    const children = n.children.map((c) => walk(c, depth + 1)).filter((c): c is TimeNode => c !== undefined)
    return children.length === n.children.length && children.every((c, i) => c === n.children[i]) ? n : { ...n, children }
  }
  const root = walk(timing.root, 0)
  return changed && root ? { ...slide, timing: { ...timing, root } } : slide
}

export function createShowMedia(o: ShowMediaOptions): ShowMedia {
  const { transport, workspace: ws, win } = o
  let ended = false
  /** The slide on screen (streamed), for the cursor and the resume. */
  let at: number | undefined
  let window: { cur: number; next: number | undefined } | undefined
  /** Look-ahead deliveries by `unit|mediaId`, until taken or out of the window. */
  const entries = new Map<string, Entry>()
  /** Keys requested (taken or opened) during the current stay in the window: never asked for again there. */
  const served = new Set<string>()
  /** The `media-show` answer requests wait for. */
  let shown: Promise<MediaShowAnswer> | undefined
  let resumeWaiters: Array<() => void> = []
  const all = new Set<AbortController>()

  const keyOf = (unit: number, mediaId: string) => `${unit}|${mediaId}`
  const inWindow = (unit: number) => !!window && (o.unitOf[window.cur] === unit || (window.next !== undefined && o.unitOf[window.next] === unit))

  const resumed = (signal: AbortSignal) => new Promise<void>((resolve, reject) => {
    if (signal.aborted) { reject(aborted(signal)); return }
    const done = () => { signal.removeEventListener('abort', stop); resolve() }
    const stop = () => { resumeWaiters = resumeWaiters.filter((f) => f !== done); reject(aborted(signal)) }
    resumeWaiters.push(done)
    signal.addEventListener('abort', stop, { once: true })
  })
  const stopResume = o.onResume(() => {
    if (ended) return
    // The server's media for this sid ended with the old socket: the show starts again where the viewer is.
    if (at !== undefined) {
      const unit = o.unitOf[at]!
      shown = transport.mediaShow(unit)
      transport.mediaCursor(unit)
    }
    const run = resumeWaiters
    resumeWaiters = []
    for (const f of run) f()
  })

  /** Ruling 4's `evict`: delivered look-ahead no one has taken, other slides' first. */
  function evict(): boolean {
    const cur = window ? o.unitOf[window.cur] : undefined
    const idle = [...entries.values()].filter((e) => e.result && !('refused' in e.result) && !e.taken)
    const pick = idle.filter((e) => e.unit !== cur)
    const go = pick.length ? pick : idle
    for (const e of go) drop(e)
    return go.length > 0
  }

  async function fetchLease(rec: MediaRecord, unit: number, signal: AbortSignal): Promise<MediaLeaseResult> {
    for (let turn = 0; ; turn++) {
      if (ended) return { refused: 'media-unavailable' }
      const answer = await (shown ?? Promise.resolve<MediaShowAnswer>({ refused: 'media-unavailable' }))
      if (signal.aborted) throw aborted(signal)
      if ('refused' in answer && answer.refused === 'session-rotated' && turn < ROTATIONS) { await resumed(signal); continue }
      if ('refused' in answer) return { refused: answer.refused }
      const r = await openRemoteMediaLease(transport, rec, ws, win, signal, { unit, evict })
      if ('refused' in r && r.refused === 'session-rotated' && turn < ROTATIONS && !ended) { await resumed(signal); continue }
      return r
    }
  }

  function track(ctl: AbortController) { all.add(ctl); return () => all.delete(ctl) }

  /** Starts a look-ahead request for `r` on server unit `unit`. */
  function prefetch(r: PlanMediaRecord, unit: number): void {
    const key = keyOf(unit, r.mediaId)
    const ctl = new AbortController()
    const untrack = track(ctl)
    const ahead = !window || o.unitOf[window.cur] !== unit
    const e: Entry = { key, unit, ctl, taken: false, ahead, promise: undefined as never }
    e.promise = fetchLease(remoteMediaRecord(r, unit), unit, ctl.signal).then((res) => {
      untrack()
      e.result = res
      // Out of the window (or the show ended) before it came: nobody will take it.
      if (!e.taken && (entries.get(key) !== e || ended) && !('refused' in res)) res.release()
      return res
    }, (err: unknown) => { untrack(); throw err })
    e.promise.catch(() => {})
    entries.set(key, e)
    served.add(key)
  }

  function drop(e: Entry): void {
    if (entries.get(e.key) === e) entries.delete(e.key)
    if (e.taken) return
    if (e.result) { if (!('refused' in e.result)) e.result.release() } else e.ctl.abort()
  }

  async function open(rec: MediaRecord, signal: AbortSignal): Promise<MediaLeaseResult> {
    if (signal.aborted) throw aborted(signal)
    const info = remoteRecordOf(rec)
    if (!info || ended) return { refused: 'media-unavailable' }
    // §5.3: what the plan refused is never asked for; the player raises its notice and the clock ends it.
    if (info.record.refused) return { refused: info.record.refused }
    const key = keyOf(info.unit, info.record.mediaId)
    const e = entries.get(key)
    if (e && !e.taken) {
      e.taken = true
      entries.delete(key)
      const cut = () => e.ctl.abort(signal.reason)
      signal.addEventListener('abort', cut, { once: true })
      let res: MediaLeaseResult
      try { res = await e.promise } finally { signal.removeEventListener('abort', cut) }
      if (signal.aborted) { if (!('refused' in res)) res.release(); throw aborted(signal) }
      if (!('refused' in res) || !e.ahead || !RETRY_AHEAD.has(res.refused)) return res
    }
    if (inWindow(info.unit)) served.add(key)
    const ctl = new AbortController()
    const untrack = track(ctl)
    const cut = () => ctl.abort(signal.reason)
    signal.addEventListener('abort', cut, { once: true })
    try { return await fetchLease(rec, info.unit, ctl.signal) } finally { untrack(); signal.removeEventListener('abort', cut) }
  }

  const player = createMediaPlayer({
    root: o.root,
    sink: o.sink,
    clocked: true,
    // F1: the plan's duration, so a late join folds onto the clip and an elapsed one never plays.
    durationOf: (rec) => remoteRecordOf(rec)?.record.durationMs,
    open,
    relieve: () => sounds.release(),
    stopTransitionSound: (onlyLooping) => sounds.release(onlyLooping),
  })
  const sounds: TransitionSoundController = createTransitionSound({
    root: o.root, sink: o.sink, slide: o.slide, open: (rec, s) => player.open(rec, s),
    // F3: late bytes join at the elapsed offset, and a finished non-looping sound never plays.
    clock: { now: o.now, durationOf: (rec) => remoteRecordOf(rec)?.record.durationMs },
  })
  const stripped = new WeakMap<ShowSlide, ShowSlide>()

  function want(): void {
    if (ended || !window) return
    const slides = [window.cur, ...(window.next !== undefined && window.next !== window.cur ? [window.next] : [])]
    // Out of the window: what was asked for ahead and never taken goes; its key may be asked for again later.
    for (const e of [...entries.values()]) if (!inWindow(e.unit)) drop(e)
    for (const k of [...served]) if (!inWindow(Number(k.slice(0, k.indexOf('|'))))) served.delete(k)
    // The current slide's first (the session also puts them first).
    for (const i of slides) {
      const unit = o.unitOf[i]
      if (unit === undefined) continue
      for (const r of o.records(i) ?? []) {
        if (r.refused) continue
        const key = keyOf(unit, r.mediaId)
        if (served.has(key) || entries.has(key)) continue
        prefetch(r, unit)
      }
    }
  }

  return {
    player,
    sounds,
    enter(slide, now) {
      let s = stripped.get(slide)
      if (!s) stripped.set(slide, s = audioOnly(slide))
      return player.enter(s, now)
    },
    start(i) {
      if (ended) return
      const unit = o.unitOf[i]
      if (unit === undefined) return
      at = i
      shown = transport.mediaShow(unit)
    },
    moved(i, next) {
      if (ended) return
      if (i === 'end') { window = undefined; want(); for (const e of [...entries.values()]) drop(e); served.clear(); return }
      const first = at === undefined
      if (at !== i) {
        at = i
        if (!first && shown) transport.mediaCursor(o.unitOf[i]!)
        // A show that opened on the end-of-show slide starts its media where the viewer first lands.
        if (!shown) shown = transport.mediaShow(o.unitOf[i]!)
      }
      window = { cur: i, next }
      want()
    },
    want,
    setPaused(paused) { player.setPaused(paused); sounds.setPaused(paused) },
    end() {
      if (ended) return
      ended = true
      stopResume()
      window = undefined
      for (const e of [...entries.values()]) drop(e)
      entries.clear()
      served.clear()
      for (const c of [...all]) c.abort()
      // The elements let go of their URLs before their leases (S2-6): sounds, then the player.
      sounds.dispose()
      player.dispose()
      if (shown) transport.mediaEnd()
      for (const f of resumeWaiters.splice(0)) f()
    },
  }
}
