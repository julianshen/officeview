/**
 * The navigation layer over the S1-9 scheduler (slideshow spec 6.2 click
 * steps, navigation rules and slide states; 4.4 `showAnimation`; 4.5
 * reduced motion).
 *
 * A cursor is `(step, event history, mode)`; render timing is
 * `scheduler.evaluate(events, t)` on a working copy of the tree, so a cursor
 * is a value and two histories at the same step and time stay distinct.
 *
 * Click steps are the main sequence's children that wait on a bare
 * `delay="indefinite"` (spec 6.2: each `clickEffect` group starts a step;
 * `withEffect`/`afterEffect` items are already nested in that group by the
 * file, with their own begin delays). Other children start on their own
 * (an `onBegin` first group) and belong to the step before them.
 *
 * Settling (the "Next during an active effect" rule): the working tree forces
 * `nextAc="seek"` on the main sequence and puts an empty waiting `par`
 * (a settle point) before every click group and after the last one. A
 * `next` that begins a settle point makes the scheduler seek every step still
 * running (finite parts to their end, indefinite ones keep running), so
 * entering step k+1 is two `next` events (settle point, then the group) and
 * Next during the last step is one (the final settle point). The scheduler's
 * own seek is therefore the only settling mechanism, and `events` stays the
 * only history. Settle points never reach a snapshot (their ids are removed
 * from `ended`).
 *
 * Checkpoint n replays steps 0..n−1 settled, with no trigger events, at the
 * time every finite effect of the initial state has ended; still-active
 * effects that never end are then taken at the end of their first iteration.
 * The result does not depend on `t`. Leaving a checkpoint (Next or a trigger)
 * replays its history at the new time, so earlier steps stay settled.
 *
 * Reduced motion: each step is entered settled (the settle point after it is
 * begun at once), and any effect still running (a trigger's) is shown at its
 * end state; indefinite ones are held at their first iteration's end.
 *
 * Static (no timing, refused timing, `showAnimation="0"`): no click steps and
 * an empty snapshot, so the scrolling viewer's paint is shown, no exit is
 * applied, and the first Next leaves the slide.
 */
import type { CTn, ContainerNode, DiagnosticSink, ShowEvent, TimeNode, TimingTree } from '../types'
import { easeProgress } from './ease'
import { createScheduler, type Scheduler, type SchedulerEntry, type SchedulerSnapshot } from './scheduler'

export interface SlideCursor {
  readonly step: number
  readonly events: readonly ShowEvent[]
  readonly mode: 'static' | 'initial' | 'checkpoint' | 'live'
}

export interface SlideTimeline {
  /** Click steps on the main sequence (0 for a static slide: one Next leaves it). */
  readonly steps: number
  stateAt(cursor: SlideCursor, t: number): SchedulerSnapshot
  /**
   * Starts the next click step, or settles the last one while a finite main
   * sequence effect still runs. Returns the **same cursor object** when the
   * slide has nothing left: the caller then leaves the slide.
   */
  next(cursor: SlideCursor, t: number): SlideCursor
  /** The settled state after `n` click steps (0 ≤ n ≤ steps, an integer; otherwise `RangeError`). */
  checkpoint(n: number): SlideCursor
  /** `undefined` when `spid` is not a trigger shape (or the slide is static). */
  trigger(cursor: SlideCursor, spid: string, t: number): SlideCursor | undefined
  triggers(): readonly string[]
  /** Additive: the slide on entry, before any click (spec 6.2 "initial"). */
  initial(): SlideCursor
  /**
   * Additive: when the main sequence's last finite effect for this history
   * ends (`-Infinity` when none runs or will run). Indefinite effects,
   * interactive sequences and media never count (spec 4.4).
   */
  busyUntil(cursor: SlideCursor): number
}

export interface SlideTimelineOptions {
  showAnimation: boolean
  reducedMotion: boolean
  /** Passed to the scheduler, so its refusals are diagnosed. */
  sink?: DiagnosticSink
  slideIndex?: number
}

const INF = Number.POSITIVE_INFINITY
const NO_EVENTS: readonly ShowEvent[] = Object.freeze([])
const EMPTY: SchedulerSnapshot = Object.freeze({ active: Object.freeze([]), frozen: Object.freeze([]), ended: new Set<number>() })
const STATIC: SlideCursor = Object.freeze({ step: 0, events: NO_EVENTS, mode: 'static' })

function isContainer(n: TimeNode): n is ContainerNode {
  return n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'
}

function walk(n: TimeNode, f: (n: TimeNode) => void): void {
  f(n)
  if (isContainer(n)) for (const c of n.children) walk(c, f)
}

/** A main sequence child that consumes a `next`: it waits on a bare indefinite delay and nothing else can begin it. */
function isClickGroup(n: TimeNode): boolean {
  const s = n.ctn.start
  return s.length > 0 && s.every((c) => c.delayMs === 'indefinite') && s.some((c) => c.evt === undefined && c.tn === undefined)
}

function finite(v: number | 'indefinite' | undefined): number | undefined {
  return typeof v === 'number' && Number.isFinite(v) && v >= 0 ? v : undefined
}

function showsFill(c: CTn): boolean {
  if (c.fill === 'remove') return false
  if (c.fill !== undefined) return true
  return c.durMs === undefined && c.end.length === 0 && c.repeatCount === undefined && c.repeatDurMs === undefined
}

/**
 * A snapshot with every still-active behaviour at its settled state: a
 * finite one at the end of its active duration (frozen, or gone without
 * fill); a never-ending one with a finite simple duration at the end of its
 * first iteration; one without a simple duration unchanged.
 */
function settle(snap: SchedulerSnapshot): SchedulerSnapshot {
  if (snap.active.length === 0) return snap
  const active: SchedulerEntry[] = []
  const frozen: SchedulerEntry[] = [...snap.frozen]
  let moved = false
  for (const e of snap.active) {
    const c = e.node.ctn
    const d = finite(c.durMs)
    if (d === undefined) { active.push(e); continue }
    const L = c.autoRev ? 2 * d : d
    const rc = c.repeatCount === 'indefinite' ? INF : typeof c.repeatCount === 'number' && c.repeatCount > 0 ? c.repeatCount : undefined
    const rd = c.repeatDurMs === 'indefinite' ? INF : finite(c.repeatDurMs)
    const ad = rc === undefined && rd === undefined ? L : Math.min(rc === undefined ? INF : L === 0 ? 0 : rc * L, rd ?? INF)
    if (!Number.isFinite(ad)) {
      active.push({ ...e, progress: easeProgress(1, c.accel, c.decel, c.autoRev), iteration: 0 })
      continue
    }
    moved = true
    if (!showsFill(c)) continue
    let k = 0
    let p = 1
    if (L > 0) {
      k = Math.floor(ad / L)
      p = ad / L - k
      if (k > 0 && p === 0) { k -= 1; p = 1 }
    }
    frozen.push({ ...e, progress: easeProgress(p, c.accel, c.decel, c.autoRev), iteration: k })
  }
  if (moved) frozen.sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
  return { active, frozen, ended: snap.ended }
}

function nexts(events: readonly ShowEvent[]): number {
  let c = 0
  for (const e of events) if (e.kind === 'next') c++
  return c
}

function checkTime(t: number, events: readonly ShowEvent[]): void {
  if (!Number.isFinite(t) || t < 0) throw new RangeError(`slide timeline: time ${t}`)
  const last = events[events.length - 1]
  if (last && t < last.t) throw new RangeError(`slide timeline: time ${t} is before the last event (${last.t})`)
}

const staticTimeline: SlideTimeline = Object.freeze({
  steps: 0,
  stateAt: () => EMPTY,
  next: (cursor: SlideCursor) => cursor,
  checkpoint(n: number) {
    if (n !== 0) throw new RangeError(`slide timeline: checkpoint ${n} of 0`)
    return STATIC
  },
  trigger: () => undefined,
  triggers: () => [],
  initial: () => STATIC,
  busyUntil: () => -INF,
})

export function createSlideTimeline(tree: TimingTree | { refused: string } | undefined, options: SlideTimelineOptions): SlideTimeline {
  if (!tree || 'refused' in tree || !options.showAnimation) return staticTimeline
  const reduced = options.reducedMotion

  // Working tree: the main sequence with forced seek and settle points. Nodes off its path are shared, never mutated.
  let main: ContainerNode | undefined
  let minId = 0
  const triggerList: string[] = []
  walk(tree.root, (n) => {
    if (n.ctn.id < minId) minId = n.ctn.id
    if (!isContainer(n)) return
    if (!main && n.kind === 'seq' && n.ctn.nodeType === 'mainSeq') main = n
    if (n.ctn.nodeType === 'interactiveSeq') {
      for (const c of n.ctn.start) {
        const spid = c.evt === 'onClick' && c.target && 'spid' in c.target ? c.target.spid : undefined
        if (spid !== undefined && !triggerList.includes(spid)) triggerList.push(spid)
      }
    }
  })
  const triggerSet = new Set(triggerList)
  const sentinels = new Set<number>()
  let steps = 0
  let root = tree.root
  let workMain: ContainerNode | undefined
  if (main) {
    const point = (): ContainerNode => {
      const id = --minId
      sentinels.add(id)
      return { kind: 'par', ctn: { id, durMs: 0, accel: 0, decel: 0, autoRev: false, fill: 'remove', start: [{ delayMs: 'indefinite' }], end: [] }, children: [] }
    }
    const children: TimeNode[] = []
    for (const kid of main.children) {
      if (isClickGroup(kid)) {
        children.push(point())
        steps++
      }
      children.push(kid)
    }
    children.push(point())
    const m = main
    workMain = { ...m, seq: { concurrent: m.seq?.concurrent ?? false, nextAc: 'seek', prevAc: m.seq?.prevAc ?? 'none' }, children }
    const wm = workMain
    const rebuild = (n: TimeNode): TimeNode => {
      if (n === m) return wm
      if (!isContainer(n)) return n
      let changed = false
      const kids = n.children.map((c) => {
        const r = rebuild(c)
        if (r !== c) changed = true
        return r
      })
      return changed ? { ...n, children: kids } : n
    }
    root = rebuild(tree.root)
  }
  const work: TimingTree = { root, build: tree.build }
  const scheduler: Scheduler = createScheduler(work, { sink: options.sink, slideIndex: options.slideIndex })

  /** The latest finite end under `node` for this history: a container that never ends is searched below. */
  const finiteEnd = (node: TimeNode, events: readonly ShowEvent[]): number => {
    if (node.kind === 'audio' || node.kind === 'video') return -INF
    const e = scheduler.endTimeOf(node.ctn.id, events)
    if (typeof e === 'number') return e
    if (e !== 'indefinite' || !isContainer(node)) return -INF
    let best = -INF
    for (const c of node.children) {
      const v = finiteEnd(c, events)
      if (v > best) best = v
    }
    return best
  }

  const busy = new WeakMap<readonly ShowEvent[], number>()
  const busyUntil = (cursor: SlideCursor): number => {
    if (cursor.mode === 'static' || cursor.mode === 'checkpoint' || !workMain) return -INF
    let v = busy.get(cursor.events)
    if (v === undefined) {
      v = finiteEnd(workMain, cursor.events)
      busy.set(cursor.events, v)
    }
    return v
  }

  /** Every finite effect of the initial state has ended by then (checkpoint replay time). */
  let horizon: number | undefined
  const settledTime = (): number => {
    if (horizon === undefined) {
      const h = finiteEnd(work.root, NO_EVENTS)
      horizon = h > 0 ? h : 0
    }
    return horizon
  }

  const withoutSentinels = (snap: SchedulerSnapshot): SchedulerSnapshot => {
    let hit = false
    for (const id of sentinels) if (snap.ended.has(id)) { hit = true; break }
    if (!hit) return snap
    const ended = new Set(snap.ended)
    for (const id of sentinels) ended.delete(id)
    return { active: snap.active, frozen: snap.frozen, ended }
  }

  const repeat = (count: number, t: number): ShowEvent[] => Array.from({ length: count }, () => ({ kind: 'next' as const, t }))

  const checkpoints = new Map<number, { cursor: SlideCursor; snap?: SchedulerSnapshot }>()
  const checkpoint = (n: number): SlideCursor => {
    if (!Number.isInteger(n) || n < 0 || n > steps) throw new RangeError(`slide timeline: checkpoint ${n} of ${steps}`)
    let c = checkpoints.get(n)
    if (!c) {
      // Settle points S0..Sn and groups 1..n: 2n + 1 nexts.
      c = { cursor: Object.freeze({ step: n, events: Object.freeze(repeat(2 * n + 1, settledTime())), mode: 'checkpoint' as const }) }
      checkpoints.set(n, c)
    }
    return c.cursor
  }

  /** The history a live cursor continues from: a checkpoint is replayed at `t`. */
  const base = (cursor: SlideCursor, t: number): readonly ShowEvent[] => {
    if (cursor.mode === 'checkpoint') {
      if (!Number.isFinite(t) || t < 0) throw new RangeError(`slide timeline: time ${t}`)
      return repeat(nexts(cursor.events), t)
    }
    checkTime(t, cursor.events)
    return cursor.events
  }

  const initialCursor: SlideCursor = Object.freeze({ step: 0, events: NO_EVENTS, mode: 'initial' })

  return {
    steps,
    initial: () => initialCursor,
    triggers: () => triggerList,
    busyUntil,
    checkpoint,
    stateAt(cursor, t) {
      if (cursor.mode === 'static') return EMPTY
      if (cursor.mode === 'checkpoint') {
        const c = checkpoints.get(cursor.step)
        if (c && c.cursor === cursor) {
          c.snap ??= withoutSentinels(settle(scheduler.evaluate(cursor.events, settledTime())))
          return c.snap
        }
        return withoutSentinels(settle(scheduler.evaluate(cursor.events, settledTime())))
      }
      const snap = withoutSentinels(scheduler.evaluate(cursor.events, t))
      return reduced ? settle(snap) : snap
    },
    next(cursor, t) {
      if (cursor.mode === 'static') return cursor
      const k = cursor.step
      if (k >= steps) {
        if (reduced || cursor.mode === 'checkpoint') return cursor
        checkTime(t, cursor.events)
        // Next during the last step: begin the final settle point while a finite effect runs.
        if (nexts(cursor.events) !== 2 * steps || !(busyUntil(cursor) > t)) return cursor
        return { step: k, events: [...cursor.events, ...repeat(1, t)], mode: 'live' }
      }
      const events = base(cursor, t)
      const want = 2 * (k + 1) + (reduced ? 1 : 0)
      return { step: k + 1, events: [...events, ...repeat(Math.max(0, want - nexts(events)), t)], mode: 'live' }
    },
    trigger(cursor, spid, t) {
      if (cursor.mode === 'static' || !triggerSet.has(spid)) return undefined
      const events = base(cursor, t)
      return { step: cursor.step, events: [...events, { kind: 'trigger', spid, t }], mode: 'live' }
    },
  }
}
