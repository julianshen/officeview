/**
 * Piece 5 (spec §5; plan D4, D15, D18, D22, X10, X14): media ends on the show clock.
 *
 * In a streamed show a media node's end is not its element's `ended` event: it is computed from the plan's
 * `durationMs` on the show clock, so the server classifier and the client driver evaluate the same histories with
 * ends at the same times. This module is that computation, DOM-free and package-free, shared by both:
 *
 * - `mediaEndOf`: one audio node's (or the transition sound's) end, from its record, its begin and the media events
 *   after it (rules 1–9 below).
 * - `computeMediaEnds`: the loop over a slide history: every audio activation whose computed end is due by `t` gets
 *   its `mediaEnd` event, earliest first, re-evaluated after each (an end can begin another node).
 *
 * Video nodes never reach it (D22): they end as piece 4 ends them (`syntheticEnds`), which is sound only because the
 * classifier stills every slide on which a video end can change what is drawn. Local shows keep ending media on the
 * element's `ended` event; nothing here runs for them.
 *
 * `MediaClockRecord` is the structural subset of the plan's `PlanMediaRecord` it needs (X14: no `src/slideshow` →
 * `src/stream` import edge). `spid` (additive to the brief) names the shape a shape-audio record plays on: a node
 * finds its record by its target, as the local player does (`rId === snd`, or `spid`), never by reading a
 * `mediaId`'s text, so the lookup is the same before and after the per-session binding renames ids.
 */
import { parseMediaCommand, type MediaCommand } from '../media/commands'
import type { AudioBehaviour, TimeNode } from '../types'
import type { SlideCursor, SlideTimeline } from './steps'

export interface MediaClockRecord { mediaId: string; kind: 'audio'; durationMs: number; refused?: string; spid?: string }
export interface MediaClock { record(nodeId: number): MediaClockRecord | undefined }
/** The node id that stands for the slide's transition sound (it is no timing node). */
export const TRANSITION_SOUND = -1
export type MediaEvent =
  | { at: number; type: 'cmd'; nodeId: number; cmd: MediaCommand } // a fired cmd node (its target media node), onstopaudio included (X10)
  | { at: number; type: 'sound-start'; nodeId: number } // a later effect sound or transition sound began
  | { at: number; type: 'slide-exit' }
/** `open`: no event in `events` ends it yet; it ends at slide exit. */
export type MediaEnd = { at: number } | { open: true }

const finiteDur = (n: { ctn: { durMs: number | 'indefinite' | undefined } }): number | undefined => {
  const d = n.ctn.durMs
  return typeof d === 'number' && Number.isFinite(d) && d >= 0 ? d : undefined
}

/**
 * When `node`, begun at `beganAt` (show clock, ms), ends (spec §5, rules 1–9). `events` may hold events of other
 * nodes and events before the begin: only those at or after `beganAt` count, in time order (stable).
 *
 * 1. A playing record ends at `beganAt + durationMs`.
 * 2. `playFrom(x)` is an absolute seek, as the player's `currentTime = x` is: the node's position in its current
 *    iteration becomes `x·1000` (clamped at `durationMs`), so `durationMs − x·1000` remains of it, and it plays.
 * 3. `pause` and `resume` (`play`, `togglePause` by the clock's state) stop and restart the node's clock.
 * 4. `repeatCount = N` plays `N × durationMs`; an indefinite repeat, a loop (the transition sound's) or "until next
 *    sound" is open until an event below ends it.
 * 5. A finite `dur` bounds the result: the earlier of `beganAt + dur` and the media end.
 * 6. A plan-refused record, or none, ends at the finite `dur`, else at once (§5.3). Only plan-refused records: a
 *    failure after the plan was sent keeps the clock end (D4).
 * 7. `stop` on the node ends it at the command. 8. `onstopaudio` ends every audio node and the transition sound.
 * 9. A sound-start (another sound, or a later one) ends an `untilNextSound` node and a looping transition sound.
 *
 * And slide exit ends anything still playing.
 */
export function mediaEndOf(node: AudioBehaviour, rec: MediaClockRecord | undefined, beganAt: number, events: readonly MediaEvent[], loop: { transitionLoop?: boolean; untilNextSound?: boolean } = {}): MediaEnd {
  const self = node.ctn.id
  const dur = finiteDur(node)
  const bound = dur === undefined ? Infinity : beganAt + dur
  const later = events.map((e, i) => ({ e, i })).filter(x => x.e.at >= beganAt).sort((a, b) => a.e.at - b.e.at || a.i - b.i).map(x => x.e)
  /** An event that ends the node whatever its clock (rules 7–9, slide exit). */
  const ends = (e: MediaEvent): boolean => {
    if (e.type === 'slide-exit') return true
    if (e.type === 'sound-start') return (!!loop.untilNextSound || !!loop.transitionLoop) && (e.nodeId !== self || e.at > beganAt)
    return e.cmd.verb === 'onstopaudio' || (e.nodeId === self && e.cmd.verb === 'stop')
  }
  const done = (at: number): MediaEnd => ({ at: Math.min(at, bound) })
  // Rule 6: the local fallback.
  if (!rec || rec.refused !== undefined) {
    const first = later.find(ends)
    return done(first ? Math.min(first.at, dur === undefined ? beganAt : bound) : dur === undefined ? beganAt : bound)
  }
  const { total, d } = extent(node, rec.durationMs, loop)
  const h: Head = { pos: 0, running: true, at: beganAt }
  for (const e of later) {
    // The clock runs out before this event.
    const out = advance(h, e.at, total)
    if (out !== undefined) return done(out)
    if (ends(e)) return done(e.at)
    if (e.type === 'cmd' && e.nodeId === self) command(h, e.cmd, d, total)
  }
  if (h.running && Number.isFinite(total)) return done(h.at + total - h.pos)
  return Number.isFinite(bound) ? { at: bound } : { open: true }
}

/** A node's own playback clock (rules 1–4): `pos` ms played into its playback (across repeats), at show time `at`. */
interface Head { pos: number; running: boolean; at: number }

/** `total`: the playback's length (`Infinity` when open); `d`: one iteration's. */
function extent(node: Pick<AudioBehaviour, 'ctn'>, durationMs: number | undefined, loop: { transitionLoop?: boolean; untilNextSound?: boolean }): { total: number; d: number | undefined } {
  const rc = node.ctn.repeatCount
  const open = rc === 'indefinite' || !!loop.transitionLoop || !!loop.untilNextSound
  const times = typeof rc === 'number' && Number.isFinite(rc) && rc > 0 ? rc : 1
  const d = durationMs !== undefined && Number.isFinite(durationMs) ? Math.max(0, durationMs) : undefined
  return { total: open || d === undefined ? Infinity : d * times, d }
}

/** Runs the clock to `to`; the time it ran out, when that is at or before `to` (the head is left as it was then). */
function advance(h: Head, to: number, total: number): number | undefined {
  if (h.running) {
    if (h.at + (total - h.pos) <= to) return h.at + (total - h.pos)
    h.pos += to - h.at
  }
  h.at = to
  return undefined
}

/** Rules 2–3: one of the node's own commands. */
function command(h: Head, c: MediaCommand, d: number | undefined, total: number): void {
  if (c.verb === 'pause') h.running = false
  else if (c.verb === 'resume' || c.verb === 'play') h.running = true
  else if (c.verb === 'togglePause') h.running = !h.running
  else if (c.verb === 'playFrom') {
    // An absolute seek within the current iteration (the player keeps its iteration count).
    const x = Math.max(0, c.seconds * 1000)
    if (d === undefined) h.pos = x
    else {
      const iter = d > 0 ? Math.floor(h.pos / d) : 0
      h.pos = iter * d + Math.min(x, d)
    }
    // Never past the playback's end (a fractional repeat's last pass is shorter than `d`).
    h.pos = Math.min(h.pos, total)
    h.running = true
  }
}

/**
 * Where `node`'s playback is at show time `t`, on the clock `mediaEndOf` runs (rules 1–4): its begin, `durationMs`
 * (undefined: unknown, never runs out) and its own commands among `events` (others are ignored, as are stops and
 * slide exits: the caller ends the node on those). The streamed player seeks late bytes here (spec §5.2), so the
 * element and the computed end cannot disagree.
 *
 * - `pos`: ms into the playback, across repeats (`offset` and `iteration` fold it onto one clip).
 * - `done`: the clock ran out by `t` (it never plays).
 */
export function mediaPlayheadAt(node: Pick<AudioBehaviour, 'ctn'>, durationMs: number | undefined, beganAt: number, events: readonly MediaEvent[], t: number, loop: { transitionLoop?: boolean; untilNextSound?: boolean } = {}): { pos: number; offset: number; iteration: number; running: boolean; done: boolean } {
  const self = node.ctn.id
  const { total, d } = extent(node, durationMs, loop)
  const h: Head = { pos: 0, running: true, at: beganAt }
  const later = events.map((e, i) => ({ e, i })).filter(x => x.e.at >= beganAt && x.e.at <= t).sort((a, b) => a.e.at - b.e.at || a.i - b.i).map(x => x.e)
  let done = false
  for (const e of later) {
    if (advance(h, e.at, total) !== undefined) { done = true; break }
    if (e.type === 'cmd' && e.nodeId === self) command(h, e.cmd, d, total)
  }
  if (!done && t >= h.at && advance(h, t, total) !== undefined) done = true
  if (done) h.pos = total
  const iteration = d !== undefined && d > 0 && Number.isFinite(h.pos) ? Math.floor(h.pos / d) : 0
  return { pos: h.pos, offset: d !== undefined && d > 0 ? h.pos - iteration * d : h.pos, iteration, running: h.running, done }
}

/** The most computed ends one `settle` call records (a guard: past it the call stops and says so). */
export const MAX_SETTLED_ENDS = 1024
const MAX_DEPTH = 256

type Walked = { audio: Map<number, AudioBehaviour>; sounds: Set<number>; bySpid: Map<string, number> }
function walkMedia(roots: ReadonlyArray<TimeNode | undefined>): Walked {
  const audio = new Map<number, AudioBehaviour>(), sounds = new Set<number>(), bySpid = new Map<string, number>()
  const walk = (n: TimeNode, depth: number) => {
    if (depth > MAX_DEPTH) return
    if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
    if (n.kind !== 'audio' && n.kind !== 'video') return
    // The first media node on a shape is the one its commands drive (the local player's `bySpid`).
    if ('spid' in n.target && !bySpid.has(n.target.spid)) bySpid.set(n.target.spid, n.ctn.id)
    if (n.kind !== 'audio') return
    if (!audio.has(n.ctn.id)) audio.set(n.ctn.id, n)
    if ('snd' in n.target) sounds.add(n.ctn.id)
  }
  for (const r of roots) if (r) walk(r, 0)
  return { audio, sounds, bySpid }
}

/**
 * Each audio node's record, found by its target as the local player finds its part: an effect sound by `mediaId ===
 * snd` (the plan renames the sound to its record), a shape's audio by the record's `spid`. `roots`: every tree the
 * slide is scheduled on (the streamed timing and the text fan-out), so either tree's ids resolve.
 */
export function mediaClockOf(roots: ReadonlyArray<TimeNode | undefined>, records: readonly MediaClockRecord[]): MediaClock {
  const byNode = new Map<number, MediaClockRecord>()
  for (const [id, n] of walkMedia(roots).audio) {
    const t = n.target
    const rec = 'snd' in t ? records.find(r => r.mediaId === t.snd) : 'spid' in t ? records.find(r => r.spid === t.spid) : undefined
    if (rec) byNode.set(id, rec)
  }
  return { record: (id) => byNode.get(id) }
}

/** One recorded end (`SlideMediaEnds.settle`). */
export interface SettledEnd { nodeId: number; at: number }

/** A slide's audio on the show clock (`slideMediaEnds`). */
export interface SlideMediaEnds {
  /** The slide has an audio node: `settle` can record something. */
  readonly audio: boolean
  /**
   * `cursor` with every computed audio end due by `t` recorded (`timeline.mediaEnd(…, at, false)`), earliest first
   * (ties by node id), each found again after the last one was recorded, as an end can begin another node. A
   * checkpoint (or a static cursor) is returned as it is: a settled checkpoint carries no media end, on either side.
   * An end due before the history's last event is recorded at that event's time (the timeline takes no event out
   * of order); it is recorded only while its activation is still open then. `truncated`: `MAX_SETTLED_ENDS` ran out, or the scheduler's
   * history window did (nothing more is recorded).
   */
  settle(timeline: SlideTimeline, cursor: SlideCursor, t: number): { cursor: SlideCursor; ends: SettledEnd[]; truncated: boolean }
}

/**
 * The shared loop (spec §5.4: the classifier and the client evaluate the same histories, with ends at the same
 * computed times). Begins are the scheduler's exact activation begins (`SlideTimeline.activations`), never a frame's
 * observation. Media events: each fired `cmd` that parses (`parseMediaCommand`), aimed at the first media node on
 * its shape, `onstopaudio` at the slide; each effect sound's begin is a `sound-start` (a shape's audio is not a
 * "sound", as locally). An effect sound with an indefinite repeat loops until the next sound (`untilNextSound`).
 */
export function slideMediaEnds(roots: ReadonlyArray<TimeNode | undefined>, clock: MediaClock): SlideMediaEnds {
  const { audio, sounds, bySpid } = walkMedia(roots)
  return {
    audio: audio.size > 0,
    settle(timeline, cursor, t) {
      const ends: SettledEnd[] = []
      if (audio.size === 0 || cursor.mode === 'checkpoint' || cursor.mode === 'static' || !(t >= 0)) return { cursor, ends, truncated: false }
      let cur = cursor
      for (let n = 0; n < MAX_SETTLED_ENDS; n++) {
        const got = timeline.activations(cur, t)
        // Activations past the scheduler's history window are unknown: nothing more is recorded, on either side.
        if (got.truncated) return { cursor: cur, ends, truncated: true }
        const acts = got.activations
        const events: MediaEvent[] = []
        for (const a of acts) {
          const node = a.node
          if (node.kind === 'cmd') {
            const c = parseMediaCommand(node.cmdType, node.cmd)
            if (!c) continue
            const target = c.verb === 'onstopaudio' ? node.ctn.id : 'spid' in node.target ? bySpid.get(node.target.spid) : undefined
            if (target !== undefined) events.push({ at: a.begin, type: 'cmd', nodeId: target, cmd: c })
          } else if (node.kind === 'audio' && sounds.has(node.ctn.id)) events.push({ at: a.begin, type: 'sound-start', nodeId: node.ctn.id })
        }
        const last = cur.events.length ? cur.events[cur.events.length - 1]!.t : 0
        let best: SettledEnd | undefined
        for (const a of acts) {
          const id = a.node.ctn.id
          const node = audio.get(id)
          if (a.node.kind !== 'audio' || !node) continue
          if (cur.events.some(e => e.kind === 'mediaEnd' && e.nodeId === id && e.t >= a.begin && e.t <= a.end)) continue
          const end = mediaEndOf(a.node, clock.record(id), a.begin, events, { untilNextSound: sounds.has(id) && a.node.ctn.repeatCount === 'indefinite' })
          if (!('at' in end) || !(end.at <= t)) continue
          const at = Math.max(end.at, last)
          if (!(at < a.end)) continue
          if (!best || at < best.at || (at === best.at && id < best.nodeId)) best = { nodeId: id, at }
        }
        if (!best) return { cursor: cur, ends, truncated: false }
        cur = timeline.mediaEnd(cur, best.nodeId, best.at, false)
        ends.push(best)
      }
      return { cursor: cur, ends, truncated: true }
    },
  }
}
