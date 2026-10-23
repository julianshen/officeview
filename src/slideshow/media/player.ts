/**
 * Media playback (slideshow plan S2-7; spec 2.1 goal 7, 6.1 `p:cmd` /
 * `p:audio` / `p:video`, 5.2 one audio output, 6.5 codecs, 7 cross-slide
 * sound, autoplay, media end, the failed-playback escape and retry).
 *
 * One `MediaPlayer` per show owns every media element; `enter(slide)` gives
 * the shown slide a `MediaController` that follows its timeline:
 *
 * - **Following the timeline.** `sync(snapshot, t)` runs once per frame. A
 *   media node that becomes active starts a new activation; a node whose own
 *   timing ends it (its `dur`, an end condition) stops its element. A `cmd`
 *   behaviour runs once when it begins: the 1 ms `playFrom(0.0)` PowerPoint
 *   writes has usually ended by the frame that samples it, so "begun" is
 *   active, frozen or ended - not merely active.
 * - **Elements.** Bytes are read only when a node plays (a lease per node,
 *   S2-6), after `canPlayType` accepts the part's content type; an
 *   undecodable format (WMV, AVI) reads nothing, leaves the poster the stage
 *   already paints, and is diagnosed (spec 6.5). Elements are created in the
 *   audience surface's own document (realm-safe) and are the only media
 *   elements of the show: the presenter mirror copies frames (`drawFrame`),
 *   never a second element or a second audio output (spec 5.2). Video sits in
 *   a DOM overlay at the shape's position on the slide (`box`), diagnosed
 *   once per slide: effects on the shape do not move the moving picture.
 * - **Media end.** A successful clip ends at its real end (or earlier by the
 *   node's own timing, which the scheduler already applies). A refusal or
 *   failure - the lease refused, `play()` rejected, a media error - gives a
 *   diagnosed synthetic end: at the explicit finite `dur` from the node's
 *   begin, on the show clock, else at once. Each node's end fires at most
 *   once per activation; a later "Click to play" (`retry`) plays the media
 *   only and never re-arms successors.
 * - **Autoplay.** Every `play()` handles rejection. On rejection or a media
 *   error the poster stays with a "Click to play" button that calls `play()`
 *   inside its own click (transient activation) and never reaches the
 *   surface's advance handler. The show itself continues.
 * - **Slide change.** `leave()` stops everything except media whose
 *   `numSld` spans further slides (`retain(numSld)`, one `leave()` per slide
 *   exit); an element lets go of its URL (`src` removed, `load()`) before
 *   its lease is released (the S2-6 contract).
 * - **Pressure.** A lease refused for the workspace first stops and releases
 *   cross-slide media (spec 7), then retries once.
 * - **Pause.** A stopped show clock pauses playing media and resumes it.
 * - **Effect sounds** (S2-10). A `p:audio` on a `p:sndTgt` - the sound of an
 *   animation effect, in its `subTnLst`, begun with the effect - is followed
 *   like shape audio: its part is the slide relationship the target names,
 *   read through the same lease and pressure rule, played in a hidden
 *   element of the audience document (the one audio output). There is no
 *   poster, so a refusal is diagnosed and the show goes on without "Click to
 *   play" (the transition-sound convention), with the same synthetic end.
 *   "Loop until next sound" is `repeatCount="indefinite"` (`loop`); "Stop
 *   previous sound" is `p:cmd type="evt" cmd="onstopaudio"`, which stops
 *   every audio node of the slide and the transition sound. A new effect
 *   sound stops a looping one, and a looping transition sound. Diagnostics name `behaviour <id>`. A
 *   built-in sound (`builtIn:<name>`, not in the package), or media on no
 *   shape and no sound, has no part: diagnosed `media-missing`.
 */
import type { SchedulerEntry, SchedulerSnapshot } from '../engine/scheduler'
import type { AudioBehaviour, CmdBehaviour, DiagnosticSink, MediaRecord, ShowSlide, TimeNode, VideoBehaviour } from '../types'
import type { MediaLease, MediaLeaseResult } from './lease'

/** A shape's box in CSS px relative to the audience surface root. */
export interface MediaBox { left: number; top: number; width: number; height: number }

export interface MediaPlayerDeps {
  /** The audience surface root: elements and buttons are created in its document. */
  root: HTMLElement
  sink: DiagnosticSink
  /** Opens one media part's lease (S2-6 `openMediaLease` bound to the show's handle, workspace and window). */
  open(rec: MediaRecord, signal: AbortSignal): Promise<MediaLeaseResult>
  /** The shape's box on the shown slide (`Stage.mediaBox`); `undefined` hides the overlay. */
  box?(spid: string): MediaBox | undefined
  /** Frees other workspace holders (the looping transition sound) when a lease is refused; true when it freed something. */
  relieve?(): boolean
  /**
   * Stops the transition sound (S2-10): `onlyLooping` when an effect sound starts ("loop until next sound"),
   * any when a "Stop previous sound" command runs. True when one was playing.
   */
  stopTransitionSound?(onlyLooping: boolean): boolean
  /** @internal Element factory (tests); default: `root.ownerDocument.createElement`. */
  create?(kind: 'audio' | 'video'): HTMLMediaElement
}

export interface MediaController {
  play(nodeId: number): void
  pause(nodeId: number): void
  stop(nodeId: number): void
  onEnd(cb: (nodeId: number, synthetic: boolean, t: number) => void): void
  retry(nodeId: number): void
  /** Additive: the slide has media nodes or media commands (otherwise `sync` is never needed). */
  readonly wants: boolean
  /** Additive: follows the timeline at slide time `t`; `quiet` records what has begun without acting (a checkpoint entry). */
  sync(snap: SchedulerSnapshot, t: number, quiet?: boolean): void
  /** Additive: the slide is left. */
  leave(): void
}

export interface MediaPlayer {
  enter(slide: ShowSlide, now: () => number): MediaController
  /** Opens a lease with the pressure rule: on a workspace refusal, cross-slide media goes first, then one retry. */
  open(rec: MediaRecord, signal: AbortSignal): Promise<MediaLeaseResult>
  setPaused(paused: boolean): void
  /** The presenter mirror: copies the current frame of the playing video on `spid` (no second element, no audio). */
  drawFrame(spid: string, ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): boolean
  /** Stops and releases media carried from earlier slides; true when anything was released. */
  dropCrossSlide(): boolean
  /** Re-places video overlays (resize). */
  layout(): void
  dispose(): void
}

type MediaNode = AudioBehaviour | VideoBehaviour

interface Live {
  spid: string
  kind: 'audio' | 'video'
  el: HTMLMediaElement
  lease: MediaLease
  /** Meant to be playing (a paused show clock pauses the element but keeps this). */
  playing: boolean
  slide: number
  /** A refused play(): the owning controller's failure path, or a diagnostic once the media is carried past its slide. */
  onReject?: (e: unknown) => void
}

interface NodeState {
  node: MediaNode
  id: number
  /** The target shape's spid; for an effect sound (no shape), the diagnostic label `behaviour <id>`. */
  spid: string
  /** An effect sound's `p:sndTgt` (`r:embed` relationship id, or `builtIn:<name>`): no shape, no poster, no overlay. */
  sound?: string
  rec?: MediaRecord
  activation: number
  ended: boolean
  begin: number
  /** A pending synthetic end (explicit finite `dur`), on the show clock. */
  due?: number
  live?: Live
  opening: boolean
  /** Bumped when a pending open is abandoned (a looping sound superseded while its lease is still opening): a late lease is released, never attached. */
  openGen: number
  /** Undecodable or missing: play commands are no-ops (spec 6.5). */
  unplayable: boolean
  want: boolean
  from?: number
  button?: HTMLButtonElement
  iterations: number
  /** Begun (active or frozen) at the last sync. */
  seen: boolean
  /** Observed active, so leaving `active` means its own timing ended it. */
  wasActive: boolean
}

const MAX_DEPTH = 128

function finiteDur(n: TimeNode): number | undefined {
  const d = n.ctn.durMs
  return typeof d === 'number' && Number.isFinite(d) && d >= 0 ? d : undefined
}

/** When the entry's current activation began, from its linear progress (media and commands carry no easing). */
function beginOf(e: SchedulerEntry | undefined, t: number): number {
  if (!e) return t
  const d = finiteDur(e.node)
  const c = e.node.ctn
  if (d === undefined || d === 0 || c.accel || c.decel || c.autoRev) return t
  return Math.max(0, t - (e.iteration + e.progress) * d)
}

function isMedia(n: TimeNode): n is MediaNode {
  return n.kind === 'audio' || n.kind === 'video'
}

const what = (st: NodeState) => st.sound !== undefined ? 'effect sound' : st.node.kind
const stays = (st: NodeState) => st.sound !== undefined ? 'the effect runs silent' : 'the poster stays'
function missing(st: NodeState): string {
  if (st.sound === undefined) return `The ${st.node.kind} on shape ${st.spid} has no media part`
  if (st.sound.startsWith('builtIn:')) return `The effect sound "${st.sound.slice(8)}" is a built-in sound with no media part in the file; the effect runs silent`
  if (st.sound === '') return `The ${st.node.kind} targets no shape and no sound; nothing to play`
  return `The effect sound (${st.sound}) has no media part; the effect runs silent`
}

export function createMediaPlayer(deps: MediaPlayerDeps): MediaPlayer {
  const doc = deps.root.ownerDocument
  const create = deps.create ?? ((kind) => doc.createElement(kind))
  /** Media carried across slide exits (`numSld`), oldest first. */
  const carried = new Set<Live>()
  let current: { lives: () => Iterable<Live>; layout: () => void } | undefined
  let paused = false
  let layer: HTMLElement | undefined

  const mount = (el: HTMLElement) => {
    if (!layer) {
      layer = doc.createElement('div')
      layer.setAttribute('data-officeview-show-media-layer', '')
      // Over the compositor, under the black/white screen and the end-of-show slide.
      const end = deps.root.querySelector(':scope > [data-officeview-show-end]')
      const blank = end?.previousElementSibling
      if (blank && blank.parentNode === deps.root) deps.root.insertBefore(layer, blank)
      else deps.root.append(layer)
    }
    layer.append(el)
  }

  /** The element lets go of its URL before the lease is released (S2-6). `force` ignores the retention span. */
  function teardown(live: Live, force: boolean): void {
    carried.delete(live)
    live.playing = false
    try { live.el.pause() } catch { /* a detached realm */ }
    live.el.removeAttribute('src')
    try { live.el.load() } catch { /* nothing to load */ }
    live.el.remove()
    if (force) live.lease.release()
    else live.lease.leave()
  }

  /** One slide exit for a live element: kept while its span reaches past this slide, else stopped and released. */
  function exitSlide(live: Live): void {
    if (live.lease.spans > 1) {
      live.lease.leave()
      carried.add(live)
      live.onReject = (e) => {
        live.playing = false
        if ((e as { name?: unknown } | undefined)?.name === 'AbortError') return
        deps.sink.add({ code: 'media-autoplay', slideIndex: live.slide, effect: live.spid, message: `The browser refused to resume this ${live.kind} (${String((e as { name?: unknown } | undefined)?.name ?? e)}); the show continues` })
      }
    } else teardown(live, false)
  }

  function dropCrossSlide(): boolean {
    if (carried.size === 0) return false
    for (const live of [...carried]) teardown(live, true)
    return true
  }

  async function open(rec: MediaRecord, signal: AbortSignal): Promise<MediaLeaseResult> {
    const first = await deps.open(rec, signal)
    if (!('refused' in first) || first.refused !== 'resource-limit') return first
    // Spec 7: cross-slide media is stopped and released first, then the read is tried once more.
    const freed = dropCrossSlide()
    if (!(deps.relieve?.() || freed)) return first
    return deps.open(rec, signal)
  }

  function resume(live: Live, onReject: (e: unknown) => void, onStart?: () => void): void {
    live.playing = true
    if (paused) return
    let p: Promise<void> | undefined
    try {
      p = live.el.play() as Promise<void> | undefined
    } catch (e) {
      onReject(e)
      return
    }
    // Old engines return undefined: treat as started.
    if (p && typeof p.then === 'function') p.then(onStart, onReject)
    else onStart?.()
  }

  function enter(slide: ShowSlide, now: () => number): MediaController {
    const nodes = new Map<number, NodeState>()
    const bySpid = new Map<string, NodeState>()
    const cmds: CmdBehaviour[] = []
    const timing = slide.timing
    if (timing && !('refused' in timing)) {
      const walk = (n: TimeNode, depth: number) => {
        if (depth > MAX_DEPTH) return
        if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
        if (isMedia(n)) {
          const t = n.target
          const onShape = 'spid' in t
          // An effect sound names its part by relationship id; media on no shape and no sound has none (media-missing).
          const sound = onShape ? undefined : 'snd' in t ? t.snd : ''
          const spid = onShape ? t.spid : `behaviour ${n.ctn.id}`
          const rec = onShape ? slide.media.find((m) => m.spid === spid) : sound ? slide.media.find((m) => m.rId === sound) : undefined
          const st: NodeState = {
            node: n, id: n.ctn.id, spid, ...(sound !== undefined ? { sound } : {}), rec,
            activation: 0, ended: false, begin: 0, opening: false, openGen: 0, unplayable: false, want: false, iterations: 0, seen: false, wasActive: false,
          }
          nodes.set(st.id, st)
          if (onShape && !bySpid.has(spid)) bySpid.set(spid, st)
        } else if (n.kind === 'cmd') cmds.push(n)
      }
      walk(timing.root, 0)
    }
    const abort = new AbortController()
    let left = false
    let endCb: ((nodeId: number, synthetic: boolean, t: number) => void) | undefined
    const begun = new Set<number>()

    const diagnose = (code: string, st: NodeState | { spid: string }, message: string) =>
      deps.sink.add({ code, slideIndex: slide.index, effect: st.spid, message })

    function endNow(st: NodeState, synthetic: boolean, t = now()): void {
      if (st.ended || left) return
      st.ended = true
      st.due = undefined
      endCb?.(st.id, synthetic, t)
    }

    function removeButton(st: NodeState): void {
      st.button?.remove()
      st.button = undefined
    }

    function showButton(st: NodeState): void {
      if (st.button || left) return
      const b = doc.createElement('button')
      b.type = 'button'
      b.textContent = 'Click to play'
      b.setAttribute('data-officeview-show-play', '')
      const box = deps.box?.(st.spid)
      if (box) {
        b.style.left = `${box.left + box.width / 2}px`
        b.style.top = `${box.top + box.height / 2}px`
      }
      b.addEventListener('click', (e) => {
        // The media's click: never a slide advance, and play() runs inside this activation.
        e.stopPropagation()
        retry(st.id)
      })
      st.button = b
      mount(b)
    }

    /** Refusal or failure: diagnosed, the poster stays, and the timeline gets its synthetic end (spec 7 escape). */
    function fail(st: NodeState, code: string, message: string, offerRetry: boolean): void {
      if (left) return
      diagnose(code, st, message)
      // An effect sound has no poster to click: diagnosed only (the transition-sound convention).
      if (offerRetry && st.live && st.sound === undefined) showButton(st)
      if (st.ended) return
      const d = finiteDur(st.node)
      if (d === undefined) { endNow(st, true); return }
      st.due = st.begin + d
      if (now() >= st.due) endNow(st, true, st.due)
    }

    function rejected(st: NodeState, live: Live, e: unknown): void {
      if (left || st.live !== live) return
      const name = (e as { name?: unknown } | undefined)?.name
      // A play() interrupted by our own pause() or a new source is not a failure.
      if (name === 'AbortError') return
      live.playing = false
      if (name === 'NotSupportedError') { st.unplayable = true; fail(st, 'media-codec', `The browser cannot decode this ${what(st)} (${st.rec?.contentType}); ${stays(st)}`, false); return }
      fail(st, 'media-autoplay', st.sound !== undefined
        ? `The browser refused to play this effect sound (${String(name ?? e)}); the effect runs silent and the show continues`
        : `The browser refused to play this ${st.node.kind} (${String(name ?? e)}); "Click to play" is offered and the show continues`, true)
    }

    function doPlay(st: NodeState): void {
      const live = st.live
      if (!live) return
      // A successful start clears the affordance; a rejection brings it back.
      resume(live, (e) => rejected(st, live, e), () => { if (st.live === live && live.playing) removeButton(st) })
    }

    function place(st: NodeState): void {
      const live = st.live
      if (!live || live.kind !== 'video') return
      const box = deps.box?.(st.spid)
      live.el.hidden = !box
      if (!box) return
      const s = live.el.style
      s.left = `${box.left}px`
      s.top = `${box.top}px`
      s.width = `${box.width}px`
      s.height = `${box.height}px`
    }

    function attach(st: NodeState, el: HTMLMediaElement, lease: MediaLease): void {
      const n = st.node
      const live: Live = { spid: st.spid, kind: n.kind, el, lease, playing: false, slide: slide.index }
      live.onReject = (e) => rejected(st, live, e)
      el.setAttribute('data-officeview-show-media', n.kind)
      if (st.sound !== undefined) {
        el.setAttribute('data-officeview-show-effect-sound', String(st.id))
        el.hidden = true
      } else el.setAttribute('data-officeview-show-media-spid', st.spid)
      el.preload = 'auto'
      if (n.kind === 'video') (el as HTMLVideoElement).playsInline = true
      el.loop = n.ctn.repeatCount === 'indefinite'
      if (n.volume !== undefined) el.volume = n.volume
      if (n.muted) el.muted = true
      if (n.numSld !== undefined && n.numSld > 1) lease.retain(n.numSld)
      el.addEventListener('ended', () => {
        if (st.live !== live) return
        if (st.iterations > 0) {
          st.iterations--
          el.currentTime = 0
          doPlay(st)
          return
        }
        live.playing = false
        endNow(st, false)
      })
      el.addEventListener('error', () => {
        if (st.live !== live) return
        live.playing = false
        const code = el.error?.code
        if (code === 4) { st.unplayable = true; fail(st, 'media-codec', `The browser cannot decode this ${what(st)} (${st.rec?.contentType}); ${stays(st)}`, false); return }
        fail(st, 'media-error', `The ${what(st)} failed while playing (media error ${code ?? 'unknown'}); ${st.sound !== undefined ? 'the show continues' : '"Click to play" is offered and the show continues'}`, true)
      })
      el.src = lease.url
      st.live = live
      if (n.kind === 'video') {
        diagnose('media-overlay', st, 'Video plays in an overlay at the shape\'s position; effects on the shape do not move the moving picture')
        if ((n as VideoBehaviour).fullScrn) diagnose('media-fullscreen', st, 'Full-screen video plays at its shape\'s position')
        place(st)
      }
      mount(el)
    }

    function start(st: NodeState, begin: number, from?: number): void {
      if (left) return
      st.activation++
      st.ended = false
      st.begin = begin
      st.due = undefined
      st.want = true
      removeButton(st)
      const rc = st.node.ctn.repeatCount
      st.iterations = typeof rc === 'number' && Number.isFinite(rc) && rc > 1 ? Math.ceil(rc) - 1 : 0
      if (st.sound !== undefined) {
        // "Loop until next sound": this sound is the next one.
        for (const o of nodes.values()) {
          if (o === st || o.sound === undefined || o.activation === 0 || o.ended) continue
          if (o.live?.el.loop) stop(o.id)
          else if (o.opening && o.node.ctn.repeatCount === 'indefinite') {
            // Still opening: the late lease must never attach or play (PR #19 review).
            o.openGen++
            o.opening = false
            o.want = false
            o.from = undefined
            endNow(o, false)
          }
        }
        deps.stopTransitionSound?.(true)
      }
      if (st.unplayable) {
        fail(st, st.rec ? 'media-codec' : 'media-missing', st.rec ? `The browser cannot decode this ${what(st)} (${st.rec.contentType}); ${stays(st)}` : missing(st), false)
        return
      }
      if (st.live) {
        if (from !== undefined) st.live.el.currentTime = from
        doPlay(st)
        return
      }
      if (st.opening) { st.from = from; return }
      const rec = st.rec
      if (!rec) {
        st.unplayable = true
        fail(st, 'media-missing', missing(st), false)
        return
      }
      const el = create(st.node.kind)
      if (el.canPlayType(rec.contentType) === '') {
        st.unplayable = true
        fail(st, 'media-codec', `The browser cannot decode this ${what(st)} (${rec.contentType}); ${stays(st)}`, false)
        return
      }
      st.opening = true
      st.from = from
      const gen = st.openGen
      open(rec, abort.signal).then((res) => {
        if (gen !== st.openGen) { if (!('refused' in res)) res.release(); return }
        st.opening = false
        if ('refused' in res) {
          if (res.refused === 'media-limit' || res.refused === 'media-empty') st.unplayable = true
          fail(st, 'media-refused', `The ${what(st)} was not read (${res.refused}); ${stays(st)}`, false)
          return
        }
        if (left) { res.release(); return }
        attach(st, el, res)
        if (st.from !== undefined) el.currentTime = st.from
        st.from = undefined
        if (st.want) doPlay(st)
      }, () => {
        // Only an abort rejects (the slide was left): nothing was charged.
        if (gen === st.openGen) st.opening = false
      })
    }

    function stopElement(st: NodeState): void {
      st.want = false
      removeButton(st)
      const live = st.live
      if (!live) return
      live.playing = false
      live.el.pause()
    }

    function play(nodeId: number): void {
      const st = nodes.get(nodeId)
      if (st) start(st, now())
    }

    function pause(nodeId: number): void {
      const st = nodes.get(nodeId)
      if (!st) return
      st.want = false
      if (st.live) { st.live.playing = false; st.live.el.pause() }
    }

    function stop(nodeId: number): void {
      const st = nodes.get(nodeId)
      if (!st) return
      stopElement(st)
      if (st.live) st.live.el.currentTime = 0
      // A stopped clip has ended: whatever waits on it runs (once per activation).
      if (st.activation > 0) endNow(st, false)
    }

    /** Plays the media only: the activation, its end and its successors are untouched (spec 7 retry). */
    function retry(nodeId: number): void {
      const st = nodes.get(nodeId)
      const live = st?.live
      if (!st || !live || left) return
      st.want = true
      live.playing = true
      let p: Promise<void> | undefined
      try {
        p = live.el.play() as Promise<void> | undefined
      } catch (e) {
        live.playing = false
        diagnose('media-autoplay', st, `The browser refused to play this ${st.node.kind} again (${String((e as { name?: unknown })?.name ?? e)})`)
        return
      }
      Promise.resolve(p).then(() => { if (st.live === live) removeButton(st) }, (e) => {
        if (st.live !== live) return
        live.playing = false
        if ((e as { name?: unknown })?.name !== 'AbortError') diagnose('media-autoplay', st, `The browser refused to play this ${st.node.kind} again (${String((e as { name?: unknown })?.name ?? e)})`)
      })
    }

    /** A running or paused activation resumes; anything else starts a new one. */
    function resumeOrStart(st: NodeState, t: number, from?: number): void {
      if (st.live && st.activation > 0 && !st.ended) {
        if (from !== undefined) st.live.el.currentTime = from
        st.want = true
        doPlay(st)
      } else start(st, t, from)
    }

    function command(c: CmdBehaviour, t: number): void {
      const text = c.cmd.trim()
      if (c.cmdType === 'evt' && text.toLowerCase() === 'onstopaudio') {
        // "Stop previous sound": every audio of the slide, and the transition sound.
        for (const st of nodes.values()) if (st.node.kind === 'audio') stop(st.id)
        deps.stopTransitionSound?.(false)
        return
      }
      const spid = 'spid' in c.target ? c.target.spid : undefined
      const st = spid === undefined ? undefined : bySpid.get(spid)
      if (c.cmdType !== 'call') { diagnose('media-command', { spid: spid ?? String(c.ctn.id) }, `cmd type="${c.cmdType}" cmd="${text}" is not supported`); return }
      if (!st) { diagnose('media-command', { spid: spid ?? String(c.ctn.id) }, `cmd "${text}" targets no media node`); return }
      const from = /^playFrom\(\s*([0-9]+(?:\.[0-9]+)?)\s*\)$/.exec(text)
      if (from) { resumeOrStart(st, t, Number(from[1])); return }
      switch (text) {
        case 'play': case 'resume': resumeOrStart(st, t); return
        case 'pause': pause(st.id); return
        case 'stop': stop(st.id); return
        case 'togglePause':
          if (st.live?.playing) pause(st.id)
          else resumeOrStart(st, t)
          return
        default:
          diagnose('media-command', st, `cmd "${text}" is not supported`)
      }
    }

    function fireDue(t: number): void {
      for (const st of nodes.values()) if (st.due !== undefined && t >= st.due && !st.ended) endNow(st, true, st.due)
    }

    function sync(snap: SchedulerSnapshot, t: number, quiet = false): void {
      if (left) return
      fireDue(t)
      if (nodes.size === 0 && cmds.length === 0) return
      const active = new Map<number, SchedulerEntry>()
      for (const e of snap.active) active.set(e.node.ctn.id, e)
      const frozen = new Set<number>()
      for (const e of snap.frozen) frozen.add(e.node.ctn.id)
      for (const st of nodes.values()) {
        const e = active.get(st.id)
        const seen = !!e || frozen.has(st.id)
        // Begun: active, or already settled when first sampled (reduced motion enters each step settled) - it plays to its clip end.
        if (seen && !st.seen && !quiet) start(st, beginOf(e, t))
        else if (!e && st.wasActive && !st.ended) {
          // The node's own timing ended it (dur, an end condition, a Next that settles its step): the clip stops with it.
          stopElement(st)
          st.ended = true
          st.due = undefined
        }
        st.seen = seen
        st.wasActive = !!e
      }
      if (cmds.length) {
        const seen = new Set<number>(active.keys())
        for (const e of snap.frozen) seen.add(e.node.ctn.id)
        for (const c of cmds) {
          const id = c.ctn.id
          const now = seen.has(id) || snap.ended.has(id)
          if (now && !begun.has(id)) {
            begun.add(id)
            if (!quiet) command(c, beginOf(active.get(id), t))
          } else if (!now) begun.delete(id)
        }
      }
      fireDue(t)
    }

    function leave(): void {
      if (left) return
      left = true
      endCb = undefined
      abort.abort()
      if (current === view) current = undefined
      // Media carried from earlier slides loses one more slide of span first.
      for (const live of [...carried]) exitSlide(live)
      for (const st of nodes.values()) {
        removeButton(st)
        st.due = undefined
        const live = st.live
        st.live = undefined
        if (live) exitSlide(live)
      }
    }

    const view = {
      lives: function* () { for (const st of nodes.values()) if (st.live) yield st.live },
      layout: () => { for (const st of nodes.values()) place(st) },
    }
    current = view

    return {
      play, pause, stop, retry,
      onEnd(cb) { endCb = cb },
      get wants() { return nodes.size > 0 || cmds.length > 0 },
      sync,
      leave,
    }
  }

  function all(): Live[] {
    return [...carried, ...(current ? current.lives() : [])]
  }

  return {
    enter,
    open,
    setPaused(p) {
      if (p === paused) return
      paused = p
      for (const live of all()) {
        if (!live.playing) continue
        if (p) { live.el.pause(); continue }
        // A resume without a fresh gesture can be refused: the same failure path as any play() (diagnosed, "Click to play").
        const refused = (e: unknown) => { if (live.onReject) live.onReject(e); else live.playing = false }
        try {
          const r = live.el.play() as Promise<void> | undefined
          if (r && typeof r.then === 'function') r.then(undefined, refused)
        } catch (e) { refused(e) }
      }
    },
    drawFrame(spid, ctx, x, y, w, h) {
      const live = all().find((l) => l.spid === spid && l.kind === 'video')
      if (!live) return false
      try {
        ctx.drawImage(live.el as HTMLVideoElement, x, y, w, h)
        return true
      } catch {
        return false
      }
    },
    dropCrossSlide,
    layout() { current?.layout() },
    dispose() {
      dropCrossSlide()
      for (const live of current ? [...current.lives()] : []) teardown(live, true)
      current = undefined
      layer?.remove()
      layer = undefined
    },
  }
}
