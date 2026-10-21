// Pinned snapshot of ss/s1-9@89c16a5 src/slideshow/engine/scheduler.ts (S1-8 reader->scheduler contract test only).
// Not product code: S1-10 replaces the import with src/slideshow/engine/scheduler once S1-9 merges.
/**
 * The event-driven timing scheduler (slideshow spec 6.2): render timing at
 * show time `t` is a pure function of `(tree, event history, t)`.
 *
 * Model (ECMA-376 Part 1 §19.5, SMIL timing):
 * - Every node lives in its parent's local time ("frame"). A node's begin
 *   instances come from its start conditions; restart rules turn them into
 *   active intervals `[begin, end)`.
 * - Simple duration `d` (`dur`; containers without one take the implicit
 *   end of their children, or `endSync`), iteration length `L = d` (`2d`
 *   with `autoRev`), active duration from `repeatCount`/`repeatDur`, then cut
 *   by the earliest end condition and by the parent's end.
 * - `par` children start with the parent, `seq` children chain, `excl`
 *   children pause the running sibling until they end.
 * - A `mainSeq` advances on `next` events and an `interactiveSeq` on its
 *   trigger's clicks: each one begins the next child waiting on
 *   `delay="indefinite"`. With `nextAc="seek"` the steps still running are
 *   settled at that instant (finite parts jump to their end, indefinite ones
 *   keep running); without `concurrent` they are cut.
 * - Ended nodes show their fill (`hold`, `freeze`, `transition`; SMIL `auto`
 *   when absent) only while their parent is active or itself showing fill.
 *
 * One time base: a frame maps absolute show time to its local time through
 * every ancestor (pauses, seeks, repeat iteration, accel/decel, autoRev).
 * Sampling, event times, cross-container references and `endTimeOf` all
 * use that same mapping (`toLocal`, `toAbsFirst`, `lastAbs`).
 *
 * Resolution is memoised recursion. Long acyclic reference chains never
 * hit a depth ceiling: past `MAX_RESOLVE` nested steps the deepest pending
 * step is deferred, resolved first from the top, and the walk retried.
 * A real reference cycle is broken where it closes (the closing reference
 * never fires) and diagnosed once per node; results that depended on the
 * broken edge are not memoised, so the rest of the graph still resolves.
 *
 * Hostile input (spec 7): missing ids and non-finite or negative times are
 * ignored. Nothing here reads a clock.
 */
import type { BehaviourNode, CTn, Condition, ContainerNode, DiagnosticSink, ShowEvent, TimeNode, TimingTree } from '../../../src/slideshow/types'
import { easeProgress, invertEase } from './ease'

/** One behaviour's timing in a snapshot. `order` is its document (pre-order) index: the composition order. */
export interface SchedulerEntry { readonly node: BehaviourNode; readonly progress: number; readonly iteration: number; readonly order?: number }

/**
 * `progress` is local simple time after accel/decel/autoRev, 0–1;
 * `iteration` counts from 0 (a frozen entry keeps its last iteration).
 * Both lists are in document order. `ended` holds the ids of every node,
 * container or behaviour, whose current activation has ended.
 */
export interface SchedulerSnapshot {
  active: ReadonlyArray<SchedulerEntry>
  frozen: ReadonlyArray<SchedulerEntry>
  ended: ReadonlySet<number>
}

export interface Scheduler {
  evaluate(events: readonly ShowEvent[], t: number): SchedulerSnapshot
  /**
   * Absolute time after which the node's last activation is never active
   * again (in a reversing parent, the reverse replay counts); `undefined`
   * when it never begins or the id is unknown.
   */
  endTimeOf(nodeId: number, events: readonly ShowEvent[]): number | 'indefinite' | undefined
}

/** Optional diagnostics: `timing-cycle`, `timing-forward-reference`, `timing-too-deep`, each once per node. */
export interface SchedulerOptions { sink?: DiagnosticSink; slideIndex?: number }

const INF = Number.POSITIVE_INFINITY
/** Tree nesting the scheduler walks (the reader enforces the same limit). */
const MAX_DEPTH = 64
/** Nested resolution steps before the deepest one is deferred and resolved first. */
const MAX_RESOLVE = 256
/** Deferral rounds per top-level query before resolution gives up (diagnosed). */
const MAX_DRIVE = 10_000
/** Activations per node (restart storms). */
const MAX_INTERVALS = 10_000
/** Interval computations per query after which provisional results are memoised anyway. */
const WORK_BUDGET = 1_000_000
/** Ancestor activations and iterations kept as syncbase history per reference (the most recent). */
const MAX_HISTORY = 1_000

interface Interval {
  begin: number
  /** Active end, local to the frame; `INF` when indefinite. */
  end: number
  d: number
  L: number
  /** Local spans during which an `excl` sibling paused it. */
  pauses: Array<[number, number]>
  /** `nextAc="seek"` jumps: from local time `at`, elapsed time gains `delta`. */
  seeks: Array<{ at: number; delta: number }>
  event?: ShowEvent
}

/** How a container's children frame relates to the frame the container lives in. */
interface Link { iv: Interval; k: number; ctn: CTn }

interface Frame {
  /** Memo identity: equal keys mean the same time mapping. */
  readonly key: string
  /** The frame the owner lives in; undefined for the root node's own frame (local = absolute). */
  readonly parent: Frame | undefined
  /** The container whose children live in this frame. */
  readonly owner: ContainerNode | undefined
  /** Sampled frames: the owner's interval and iteration (full mapping). */
  readonly link?: Link
  /** Provisional frames, used while the owner's own duration is being resolved: linear from this begin. */
  readonly pre?: number
  /** Local sample time. */
  readonly tau: number
  /** The owner has ended at the sample: still-active children are cut. */
  readonly frozen: boolean
  /** The owner is active, or ended and showing its fill. */
  readonly visible: boolean
  /** The owner repeats forever and is sampled at infinity. */
  readonly unbounded: boolean
  readonly beganBy?: ShowEvent
}

interface Inst { t: number; event?: ShowEvent }

interface Sample { status: 'idle' | 'active' | 'ended'; visible: boolean; progress: number; iteration: number; child?: Frame }

const IDLE: Sample = { status: 'idle', visible: false, progress: 0, iteration: 0 }

interface Index {
  root: TimeNode
  list: TimeNode[]
  order: Map<TimeNode, number>
  parent: Map<TimeNode, ContainerNode>
  kids: Map<ContainerNode, TimeNode[]>
  byId: Map<number, TimeNode>
}

interface Group {
  owner: ContainerNode
  frame: Frame
  kids: TimeNode[]
  raw: Map<TimeNode, Interval[]>
  adjusted?: Map<TimeNode, Interval[]>
  next: number
  seqBusy: boolean
  queue: Inst[]
  qi: number
}

type Report = (code: 'timing-cycle' | 'timing-forward-reference' | 'timing-too-deep', node: TimeNode) => void

/** Thrown to unwind a too-deep resolution; never escapes a query. */
const DEFER: unique symbol = Symbol('scheduler.defer')

function isContainer(n: TimeNode): n is ContainerNode {
  return n.kind === 'par' || n.kind === 'seq' || n.kind === 'excl'
}

/** A time value: finite non-negative ms, `'indefinite'`, or absent/invalid. */
function time(v: number | 'indefinite' | undefined): number | 'indefinite' | undefined {
  if (v === 'indefinite') return v
  return typeof v === 'number' && Number.isFinite(v) && v >= 0 ? v : undefined
}

function count(v: number | 'indefinite' | undefined): number | undefined {
  if (v === 'indefinite') return INF
  return typeof v === 'number' && v > 0 ? v : undefined
}

function showsFill(ctn: CTn): boolean {
  if (ctn.fill === 'remove') return false
  if (ctn.fill !== undefined) return true // hold, freeze, transition
  // SMIL fill="auto": freeze only when no dur, end or repeat is given.
  return ctn.durMs === undefined && ctn.end.length === 0 && ctn.repeatCount === undefined && ctn.repeatDurMs === undefined
}

/** A seq child waiting on `delay="indefinite"` for its sequence to advance. */
function waits(n: TimeNode): boolean {
  return n.ctn.start.some((c) => c.delayMs === 'indefinite' && c.evt === undefined && c.tn === undefined)
}

// ---------------------------------------------------------------------------
// Time mapping
// ---------------------------------------------------------------------------

/** Elapsed active time of an interval at parent-local `tau` (pauses hold it, seeks jump it). */
function activeElapsed(iv: Interval, tau: number): number {
  if (!(tau > iv.begin)) return 0
  if (tau === INF) return INF
  let e = tau - iv.begin
  for (const [s, x] of iv.pauses) {
    if (s >= tau) break
    e -= Math.min(x, tau) - s
  }
  for (const sk of iv.seeks) if (tau >= sk.at) e += sk.delta
  return e > 0 ? e : 0
}

/** The earliest parent-local time at which `activeElapsed` reaches `e` (`INF` if never). */
function elapsedInverse(iv: Interval, e: number): number {
  if (!(e > 0)) return iv.begin
  if (!Number.isFinite(e)) return INF
  const marks: Array<{ at: number; until?: number; delta?: number }> = []
  for (const [s, x] of iv.pauses) marks.push({ at: s, until: x })
  for (const sk of iv.seeks) marks.push({ at: sk.at, delta: sk.delta })
  marks.sort((a, b) => a.at - b.at)
  let cur = iv.begin
  let acc = 0
  for (const m of marks) {
    if (m.at > cur) {
      if (acc + (m.at - cur) >= e) return cur + (e - acc)
      acc += m.at - cur
      cur = m.at
    }
    if (m.delta !== undefined) {
      acc += m.delta
      if (acc >= e) return cur
    } else if (m.until !== undefined && m.until > cur) {
      cur = m.until
      if (!Number.isFinite(cur)) return INF
    }
  }
  return cur + (e - acc)
}

function position(elapsed: number, L: number, ended: boolean): { k: number; p: number } {
  if (!Number.isFinite(L)) return { k: 0, p: 0 }
  if (L <= 0) return { k: 0, p: 1 }
  if (!Number.isFinite(elapsed)) return { k: 0, p: 0 }
  let k = Math.floor(elapsed / L)
  let p = (elapsed - k * L) / L
  // The frozen value at an iteration boundary is the end of the last iteration.
  if (ended && k > 0 && p === 0) {
    k -= 1
    p = 1
  }
  return { k, p }
}

/** Children-local time for the owner's elapsed time `e` in iteration `link.k`. */
function localFromElapsed(link: Link, e: number, lenient: boolean): number | undefined {
  const { iv, k, ctn } = link
  if (!Number.isFinite(iv.L)) return e
  const w = e - k * iv.L
  if (w < 0) return lenient ? w : undefined
  if (w > iv.L) return lenient ? iv.d : undefined
  if (iv.L <= 0 || iv.d <= 0) return 0
  return easeProgress(w / iv.L, ctn.accel, ctn.decel, ctn.autoRev) * iv.d
}

/**
 * Absolute time to frame-local time. Strict mode drops times outside the
 * frame's current activation and iteration (events from an earlier
 * iteration are reset, SMIL); lenient mode extrapolates times before it.
 */
function toLocal(frame: Frame, abs: number, lenient = false): number | undefined {
  if (!frame.parent) return abs >= 0 || lenient ? abs : undefined
  const pl = toLocal(frame.parent, abs, lenient)
  if (pl === undefined) return undefined
  if (frame.pre !== undefined) {
    const v = pl - frame.pre
    return v >= 0 || lenient ? v : undefined
  }
  const link = frame.link!
  if (pl < link.iv.begin) return lenient ? pl - link.iv.begin : undefined
  if (pl >= link.iv.end && !lenient) return undefined
  return localFromElapsed(link, activeElapsed(link.iv, pl), lenient)
}

/** The first absolute time at which the frame's local time reaches `t` (`INF` if never). */
function toAbsFirst(frame: Frame, t: number): number {
  if (!frame.parent) return t
  if (frame.pre !== undefined) return toAbsFirst(frame.parent, t + frame.pre)
  const { iv, k, ctn } = frame.link!
  let e: number
  if (!Number.isFinite(iv.L)) e = t
  else if (iv.L <= 0 || iv.d <= 0) e = t <= 0 ? k * iv.L : INF
  else if (t > iv.d) e = INF
  else {
    const q = invertEase(Math.max(0, t) / iv.d, ctn.accel, ctn.decel)
    e = k * iv.L + (ctn.autoRev ? q / 2 : q) * iv.L
  }
  const pl = elapsedInverse(iv, e)
  if (!(pl <= iv.end)) return INF
  return toAbsFirst(frame.parent, pl)
}

/**
 * The absolute time at which a node active on local `[lo, hi)` stops being
 * active for the last time in the frame's current iteration, clipped by
 * every ancestor's end. A reversing ancestor replays its children, so the
 * last exit is its reverse crossing of `lo`.
 */
function lastAbs(frame: Frame, lo: number, hi: number): number {
  if (!frame.parent) return hi
  if (frame.pre !== undefined) return lastAbs(frame.parent, lo + frame.pre, hi + frame.pre)
  const { iv, k, ctn } = frame.link!
  let elo: number
  let ehi: number
  if (!Number.isFinite(iv.L)) {
    elo = lo
    ehi = hi
  } else if (iv.L <= 0 || iv.d <= 0) {
    elo = k * iv.L
    ehi = lo <= 0 && hi > 0 ? INF : k * iv.L
  } else {
    const base = k * iv.L
    const unit = (v: number): number => (v > 0 ? (v < iv.d ? v / iv.d : 1) : 0)
    const qlo = invertEase(unit(lo), ctn.accel, ctn.decel)
    const qhi = hi >= iv.d ? undefined : invertEase(unit(hi), ctn.accel, ctn.decel)
    if (ctn.autoRev) {
      elo = base + (qhi === undefined ? qlo / 2 : 1 - qhi / 2) * iv.L
      ehi = base + (lo <= 0 ? 1 : 1 - qlo / 2) * iv.L
    } else {
      elo = base + qlo * iv.L
      ehi = base + (qhi === undefined ? 1 : qhi) * iv.L
    }
  }
  return lastAbs(frame.parent, elapsedInverse(iv, elo), Math.min(elapsedInverse(iv, ehi), iv.end))
}

// ---------------------------------------------------------------------------

function eventMatcher(c: Condition): ((e: ShowEvent) => boolean) | undefined {
  const spid = c.target && 'spid' in c.target ? c.target.spid : undefined
  switch (c.evt) {
    case 'onClick':
      // A click on a shape is a trigger; a click on the slide (or no target) is "next".
      return spid !== undefined ? (e) => e.kind === 'trigger' && e.spid === spid : (e) => e.kind === 'next'
    case 'onNext':
      return (e) => e.kind === 'next'
    case 'onPrev':
      return (e) => e.kind === 'previous'
    case 'onMouseOver':
      return (e) => e.kind === 'mouseOver' && (spid === undefined || e.spid === spid)
    case 'onMouseOut':
      return (e) => e.kind === 'mouseOut' && (spid === undefined || e.spid === spid)
    case 'onStopAudio':
      return (e) => e.kind === 'stopAudio'
    default:
      return undefined
  }
}

/** Excl: a child beginning while a sibling runs pauses it until the child ends. */
function pauseSiblings(kids: readonly TimeNode[], raw: ReadonlyMap<TimeNode, readonly Interval[]>): Map<TimeNode, Interval[]> {
  interface Slot { iv: Interval; rank: number; pausedAt?: number }
  const out = new Map<TimeNode, Interval[]>()
  const slots: Slot[] = []
  kids.forEach((k, rank) => {
    const copies = (raw.get(k) ?? []).map((iv) => ({ ...iv, pauses: [...iv.pauses], seeks: [...iv.seeks] }))
    out.set(k, copies)
    for (const iv of copies) slots.push({ iv, rank })
  })
  slots.sort((a, b) => a.iv.begin - b.iv.begin || a.rank - b.rank)
  const stack: Slot[] = []
  const settle = (limit: number): void => {
    while (stack.length > 0) {
      const top = stack[stack.length - 1]
      if (!(top.iv.end <= limit)) return
      stack.pop()
      const below = stack[stack.length - 1]
      if (below && below.pausedAt !== undefined) {
        const t = top.iv.end
        below.iv.pauses.push([below.pausedAt, t])
        below.iv.end += t - below.pausedAt
        below.pausedAt = undefined
      }
    }
  }
  for (const s of slots) {
    settle(s.iv.begin)
    const top = stack[stack.length - 1]
    if (top) top.pausedAt = s.iv.begin
    stack.push(s)
  }
  settle(INF)
  return out
}

function buildIndex(tree: TimingTree): Index {
  const ix: Index = { root: tree.root, list: [], order: new Map(), parent: new Map(), kids: new Map(), byId: new Map() }
  const stack: Array<[TimeNode, ContainerNode | undefined, number]> = [[tree.root, undefined, 0]]
  while (stack.length > 0) {
    const [n, p, depth] = stack.pop()!
    if (ix.order.has(n)) continue // a node object reachable twice is walked once
    ix.order.set(n, ix.list.length)
    ix.list.push(n)
    if (p) {
      ix.parent.set(n, p)
      ix.kids.get(p)!.push(n)
    }
    if (!ix.byId.has(n.ctn.id)) ix.byId.set(n.ctn.id, n) // duplicate ids: the first wins
    if (isContainer(n)) {
      ix.kids.set(n, [])
      if (depth < MAX_DEPTH) for (let i = n.children.length - 1; i >= 0; i--) stack.push([n.children[i], n, depth + 1])
    }
  }
  return ix
}

/** One query: memo tables live for a single `(events, t)`. */
class Run {
  readonly rootFrame: Frame
  private readonly groups = new Map<string, Group>()
  private rootIvs: Interval[] | undefined
  private readonly states = new Map<TimeNode, Sample>()
  private readonly refs = new Map<string, number[]>()
  private readonly ivIds = new WeakMap<Interval, number>()
  private ivIdNext = 0
  // Resolution bookkeeping: in-progress keys with their entry depth, the
  // shallowest in-progress step the current step leaned on, and deferral.
  private readonly busy = new Map<unknown, number>()
  private depth = 0
  private lowHit = INF
  private work = 0
  private deferred: (() => unknown) | undefined
  private giveUp = false

  constructor(private readonly ix: Index, private readonly events: readonly ShowEvent[], t: number, private readonly report: Report) {
    this.rootFrame = { key: 'r', parent: undefined, owner: undefined, tau: t, frozen: false, visible: true, unbounded: false }
  }

  /** Runs a top-level query, resolving deferred deep steps first until it completes. */
  drive<T>(fn: () => T): T {
    const jobs: Array<() => unknown> = []
    for (let round = 0; ; round++) {
      if (round > MAX_DRIVE) this.giveUp = true
      try {
        if (jobs.length > 0) {
          jobs[jobs.length - 1]()
          jobs.pop()
          continue
        }
        return fn()
      } catch (e) {
        if (e !== DEFER || !this.deferred) throw e
        jobs.push(this.deferred)
        this.deferred = undefined
      }
    }
  }

  /**
   * One memoised resolution step. A re-entered key is a reference cycle:
   * it answers `hit()`, is diagnosed, and taints every step above it up to
   * the cycle's entry so their provisional results are not memoised.
   */
  private guarded<T>(key: unknown, hit: () => T, store: (v: T) => void, fn: () => T, node: TimeNode | undefined, cycle = true): T {
    const entered = this.busy.get(key)
    if (entered !== undefined) {
      if (cycle) {
        if (entered < this.lowHit) this.lowHit = entered
        if (node) this.report('timing-cycle', node)
      }
      return hit()
    }
    if (this.depth >= MAX_RESOLVE) {
      if (this.giveUp) {
        if (node) this.report('timing-too-deep', node)
        return hit()
      }
      this.deferred = () => this.guarded(key, hit, store, fn, node, cycle)
      throw DEFER
    }
    const mine = ++this.depth
    const saved = this.lowHit
    this.lowHit = INF
    this.busy.set(key, mine)
    try {
      const v = fn()
      const low = this.lowHit
      if (low >= mine || this.work > WORK_BUDGET) store(v)
      this.lowHit = low >= mine ? saved : Math.min(saved, low)
      return v
    } catch (e) {
      this.lowHit = saved
      throw e
    } finally {
      this.busy.delete(key)
      this.depth--
    }
  }

  // -- sampling ---------------------------------------------------------------

  stateOf(node: TimeNode): Sample | undefined {
    const have = this.states.get(node)
    if (have) return have
    return this.guarded<Sample | undefined>(node, () => undefined, (v) => { if (v) this.states.set(node, v) }, () => {
      const f = this.currentFrame(node)
      return f ? this.sample(node, f) : IDLE
    }, node)
  }

  currentFrame(node: TimeNode): Frame | undefined {
    if (node === this.ix.root) return this.rootFrame
    const p = this.ix.parent.get(node)
    return p ? this.stateOf(p)?.child : undefined
  }

  private ivId(iv: Interval): number {
    let id = this.ivIds.get(iv)
    if (id === undefined) this.ivIds.set(iv, (id = this.ivIdNext++))
    return id
  }

  private sample(node: TimeNode, frame: Frame): Sample {
    const ivs = this.intervalsIn(node, frame)
    const tau = frame.tau
    let iv: Interval | undefined
    for (let i = ivs.length - 1; i >= 0; i--) {
      if (ivs[i].begin <= tau) {
        iv = ivs[i]
        break
      }
    }
    if (!iv) return IDLE
    const ended = tau >= iv.end || frame.frozen
    const elapsed = activeElapsed(iv, tau >= iv.end ? iv.end : tau)
    const { k, p } = position(elapsed, iv.L, ended)
    const ctn = node.ctn
    const progress = easeProgress(p, ctn.accel, ctn.decel, ctn.autoRev)
    const visible = frame.visible && (!ended || showsFill(ctn))
    let child: Frame | undefined
    if (isContainer(node)) {
      const finiteL = Number.isFinite(iv.L)
      const unbounded = frame.unbounded || (!Number.isFinite(elapsed) && finiteL && iv.L > 0)
      // Children see the owner's eased (and mirrored) simple time: the same map as `localFromElapsed`.
      const tau2 = unbounded ? INF : finiteL ? (iv.d > 0 ? progress * iv.d : 0) : elapsed
      child = {
        key: `${frame.key}/${this.ivId(iv)}.${k}`,
        parent: frame,
        owner: node,
        link: { iv, k, ctn },
        tau: tau2,
        frozen: ended,
        visible,
        unbounded,
        beganBy: k === 0 ? iv.event : undefined,
      }
    }
    return { status: ended ? 'ended' : 'active', visible, progress, iteration: k, child }
  }

  // -- intervals --------------------------------------------------------------

  intervalsIn(node: TimeNode, frame: Frame): readonly Interval[] {
    if (node === this.ix.root) {
      if (this.rootIvs) return this.rootIvs
      return this.guarded('root', () => [], (v) => { this.rootIvs = v }, () => this.own(node, frame, 0), node)
    }
    const owner = frame.owner
    if (!owner || this.ix.parent.get(node) !== owner) return []
    const g = this.group(owner, frame)
    if (owner.kind === 'seq') {
      this.seqUpTo(g, g.kids.length - 1) // a later advance can still settle an earlier step
      const got = g.raw.get(node)
      if (!got) {
        this.report('timing-forward-reference', node)
        return []
      }
      return got
    }
    if (owner.kind === 'excl') return this.exclChild(g, node)
    return this.parChild(g, node)
  }

  private group(owner: ContainerNode, frame: Frame): Group {
    let g = this.groups.get(frame.key)
    if (!g) {
      g = {
        owner,
        frame,
        kids: this.ix.kids.get(owner) ?? [],
        raw: new Map(),
        next: 0,
        seqBusy: false,
        queue: owner.kind === 'seq' ? this.advanceQueue(owner, frame) : [],
        qi: 0,
      }
      this.groups.set(frame.key, g)
    }
    return g
  }

  private parChild(g: Group, node: TimeNode): Interval[] {
    const have = g.raw.get(node)
    if (have) return have
    return this.guarded(`${g.frame.key}|${this.ix.order.get(node)}`, () => [], (v) => { g.raw.set(node, v) }, () => this.own(node, g.frame, 0), node)
  }

  private exclChild(g: Group, node: TimeNode): readonly Interval[] {
    if (g.adjusted) return g.adjusted.get(node) ?? []
    // While the group is being adjusted, siblings see each other's raw intervals.
    const m = this.guarded<Map<TimeNode, Interval[]> | undefined>(`${g.frame.key}|excl`, () => undefined, (v) => { g.adjusted = v }, () => {
      const raw = new Map<TimeNode, Interval[]>()
      for (const k of g.kids) raw.set(k, this.parChild(g, k))
      return pauseSiblings(g.kids, raw)
    }, undefined, false)
    return m ? m.get(node) ?? [] : this.parChild(g, node)
  }

  /** The events that advance a sequence: `next` for the main sequence, its trigger's clicks for an interactive one. */
  private advanceQueue(owner: ContainerNode, frame: Frame): Inst[] {
    let match: (e: ShowEvent) => boolean
    if (owner.ctn.nodeType === 'mainSeq') match = (e) => e.kind === 'next'
    else if (owner.ctn.nodeType === 'interactiveSeq') {
      const spids = new Set<string>()
      for (const c of owner.ctn.start) if (c.evt === 'onClick' && c.target && 'spid' in c.target) spids.add(c.target.spid)
      match = (e) => e.kind === 'trigger' && spids.has(e.spid)
    } else return []
    const out: Inst[] = []
    for (const e of this.events) {
      if (e === frame.beganBy || !match(e)) continue
      const t = toLocal(frame, e.t)
      if (t !== undefined) out.push({ t, event: e })
    }
    return out
  }

  private seqUpTo(g: Group, idx: number): void {
    if (g.seqBusy) return
    g.seqBusy = true
    try {
      while (g.next <= idx) {
        this.seqStep(g, g.next)
        g.next++
      }
    } finally {
      g.seqBusy = false
    }
  }

  /** One seq child; transactional, so a deferral retries it from scratch. */
  private seqStep(g: Group, j: number): void {
    const qi = g.qi
    try {
      const kid = g.kids[j]
      const prevList = j > 0 ? g.raw.get(g.kids[j - 1]) ?? [] : []
      const prev = prevList[prevList.length - 1]
      const base = j === 0 ? 0 : prev && Number.isFinite(prev.end) ? prev.end : undefined
      let insts = this.startInstances(kid, g.frame, base)
      let advanced = false
      if (insts.length === 0 && (j === 0 || prev) && waits(kid)) {
        const from = prev ? prev.begin : 0
        while (g.qi < g.queue.length) {
          const q = g.queue[g.qi++]
          if (q.t >= from) {
            insts = [q]
            advanced = true
            break
          }
        }
      }
      const ivs = this.build(kid, g.frame, insts)
      const first = ivs[0]
      const commits: Array<() => void> = []
      if (first) {
        const opts = g.owner.seq
        for (let i = 0; i < j; i++) {
          const list = g.raw.get(g.kids[i])
          const last = list?.[list.length - 1]
          if (!last || !(last.begin <= first.begin && first.begin < last.end)) continue
          if (advanced && opts?.nextAc === 'seek') {
            const c = this.seekPlan(g.kids[i], last, first.begin, g.frame)
            if (c) commits.push(c)
          } else if (!opts?.concurrent) commits.push(() => { last.end = first.begin })
        }
      }
      g.raw.set(kid, ivs)
      for (const c of commits) c()
    } catch (e) {
      g.qi = qi
      throw e
    }
  }

  /** Settling a running step at local time `s`: finite parts jump to their end, indefinite ones keep running. */
  private seekPlan(node: TimeNode, iv: Interval, s: number, frame: Frame): (() => void) | undefined {
    let settle: number
    if (Number.isFinite(iv.end)) settle = activeElapsed(iv, iv.end)
    else if (isContainer(node) && !Number.isFinite(iv.L)) {
      const cf = this.preFrame(node, frame, iv.begin, iv.event)
      settle = 0
      for (const c of this.ix.kids.get(node) ?? []) {
        const l = this.intervalsIn(c, cf)
        const e = l[l.length - 1]?.end
        if (e !== undefined && Number.isFinite(e) && e > settle) settle = e
      }
    } else return undefined
    const delta = settle - activeElapsed(iv, s)
    if (!(delta > 0)) return undefined
    return () => {
      iv.seeks.push({ at: s, delta })
      if (Number.isFinite(iv.end)) iv.end = Math.max(s, iv.end - delta)
    }
  }

  private own(node: TimeNode, frame: Frame, base: number | undefined): Interval[] {
    return this.build(node, frame, this.startInstances(node, frame, base))
  }

  private startInstances(node: TimeNode, frame: Frame, base: number | undefined): Inst[] {
    const conds = node.ctn.start
    if (conds.length === 0) return base === undefined ? [] : [{ t: base }]
    return this.instances(conds, base, frame, node)
  }

  /** Restart rules turn sorted begin instances into active intervals. */
  private build(node: TimeNode, frame: Frame, insts: Inst[]): Interval[] {
    const sorted = insts.filter((i) => Number.isFinite(i.t)).sort((a, b) => a.t - b.t)
    const restart = node.ctn.restart ?? 'always'
    const out: Interval[] = []
    for (const inst of sorted) {
      const last = out[out.length - 1]
      if (last) {
        if (inst.t <= last.begin) continue
        if (restart === 'never') break
        if (inst.t < last.end) {
          if (restart !== 'always') continue
          last.end = inst.t
        }
      }
      out.push(this.interval(node, frame, inst))
      if (out.length >= MAX_INTERVALS) break
    }
    return out
  }

  /** The children frame used while a container's own duration is resolved (linear in its elapsed time). */
  private preFrame(node: ContainerNode, frame: Frame, begin: number, event: ShowEvent | undefined): Frame {
    return { key: `${frame.key}/${this.ix.order.get(node)}p${begin}`, parent: frame, owner: node, pre: begin, tau: INF, frozen: false, visible: false, unbounded: false, beganBy: event }
  }

  private interval(node: TimeNode, frame: Frame, inst: Inst): Interval {
    this.work++
    const b = inst.t
    const ctn = node.ctn
    const dur = time(ctn.durMs)
    let d: number
    if (isContainer(node) && (ctn.endSync || dur === undefined)) {
      const cf = this.preFrame(node, frame, b, inst.event)
      const implicit = ctn.endSync ? this.endSync(node, ctn.endSync, cf) : this.implicitEnd(node, cf)
      d = dur === undefined ? implicit : Math.min(dur === 'indefinite' ? INF : dur, implicit)
    } else d = dur === undefined || dur === 'indefinite' ? INF : dur
    const L = ctn.autoRev ? d * 2 : d
    const rc = count(ctn.repeatCount)
    const rdv = time(ctn.repeatDurMs)
    const rd = rdv === undefined ? undefined : rdv === 'indefinite' ? INF : rdv
    let ad: number
    if (rc === undefined && rd === undefined) ad = L
    else ad = Math.min(rc === undefined ? INF : L === 0 ? 0 : rc * L, rd ?? INF)
    let end = b + ad
    for (const e of this.instances(ctn.end, 0, frame, node)) if (e.t > b && e.t < end) end = e.t
    if (node.kind === 'audio' || node.kind === 'video') {
      for (const e of this.events) {
        if (e.kind !== 'mediaEnd' || e.nodeId !== ctn.id) continue
        const t = toLocal(frame, e.t)
        if (t !== undefined && t > b && t < end) end = t
      }
    }
    return { begin: b, end, d, L, pauses: [], seeks: [], event: inst.event }
  }

  private implicitEnd(node: ContainerNode, cf: Frame): number {
    let end = 0
    for (const c of this.ix.kids.get(node) ?? []) {
      const l = this.intervalsIn(c, cf)
      const last = l[l.length - 1]
      if (last && last.end > end) end = last.end
    }
    return end
  }

  private endSync(node: ContainerNode, c: Condition, cf: Frame): number {
    const extra = typeof c.delayMs === 'number' && Number.isFinite(c.delayMs) && c.delayMs > 0 ? c.delayMs : c.delayMs === 'indefinite' ? INF : 0
    const kids = this.ix.kids.get(node) ?? []
    let t: number
    if (c.tn !== undefined) {
      const ends = this.refTimes(c.tn, 'end', cf, node)
      t = ends.length > 0 ? Math.min(...ends) : INF
    } else if (c.rtn === 'first') {
      t = INF
      for (const k of kids) {
        const first = this.intervalsIn(k, cf)[0]
        if (first && first.end < t) t = first.end
      }
    } else if (c.rtn === 'all') {
      t = 0
      for (const k of kids) {
        const l = this.intervalsIn(k, cf)
        const last = l[l.length - 1]
        if (!last) return INF // a child that never begins holds it open
        if (last.end > t) t = last.end
      }
    } else if (c.rtn === 'none') t = INF
    else t = this.implicitEnd(node, cf)
    return t + extra
  }

  private instances(conds: readonly Condition[], base: number | undefined, frame: Frame, node: TimeNode): Inst[] {
    const out: Inst[] = []
    for (const c of conds) {
      if (typeof c.delayMs !== 'number' || !Number.isFinite(c.delayMs)) continue
      const dl = Math.max(0, c.delayMs)
      switch (c.evt) {
        case undefined:
          if (c.tn === undefined && base !== undefined) out.push({ t: base + dl })
          break
        case 'onBegin':
        case 'begin':
          if (c.tn === undefined) out.push({ t: dl }) // the parent's begin
          else for (const t of this.refTimes(c.tn, 'begin', frame, node)) out.push({ t: t + dl })
          break
        case 'onEnd':
        case 'end':
          if (c.tn !== undefined) for (const t of this.refTimes(c.tn, 'end', frame, node)) out.push({ t: t + dl })
          break
        default: {
          const match = eventMatcher(c)
          if (!match) break
          for (const e of this.events) {
            if (!match(e)) continue
            const t = toLocal(frame, e.t)
            if (t !== undefined) out.push({ t: t + dl, event: e })
          }
        }
      }
    }
    return out
  }

  /** Begin or end times of node `tn`, local to `frame`. */
  private refTimes(tn: number, which: 'begin' | 'end', frame: Frame, from: TimeNode): number[] {
    const target = this.ix.byId.get(tn)
    if (!target || target === from) return []
    for (let f: Frame | undefined = frame; f; f = f.parent) {
      if (f.owner !== target) continue
      // An ancestor: its begin is its frame's local 0; its end lies beyond its children.
      if (which !== 'begin') return []
      if (f === frame) return [0]
      const l = toLocal(frame, toAbsFirst(f, 0), true)
      return l !== undefined && Number.isFinite(l) ? [l] : []
    }
    // Memoised like every other step, so a deferred reference is not re-deferred on retry.
    const key = `${frame.key}|${this.ix.order.get(target)}|${which}`
    const have = this.refs.get(key)
    if (have) return have
    return this.guarded(`ref|${key}`, () => [], (v) => { this.refs.set(key, v) }, () => this.refTimesOf(target, which, frame), target)
  }

  private refTimesOf(target: TimeNode, which: 'begin' | 'end', frame: Frame): number[] {
    const out: number[] = []
    for (const g of this.historyFor(target, frame)) {
      for (const iv of this.intervalsIn(target, g)) {
        if (g === frame) {
          const t = which === 'begin' ? iv.begin : iv.end
          if (Number.isFinite(t)) out.push(t)
          continue
        }
        // Through absolute time, with both frames' full mappings. "End" is the last exit, as `endTimeOf`.
        const abs = which === 'begin' ? toAbsFirst(g, iv.begin) : lastAbs(g, iv.begin, iv.end)
        if (!Number.isFinite(abs)) continue
        let l = toLocal(frame, abs)
        if (l === undefined) {
          const early = toLocal(frame, abs, true) // before this frame began: a negative instance (SMIL)
          if (early !== undefined && early < 0) l = early
        }
        if (l !== undefined) out.push(l)
      }
    }
    return out
  }

  /**
   * Every frame `target` has lived in up to the sample, oldest first, as
   * seen from `frame`: from their lowest common ancestor's frame (in
   * `frame`'s own chain) down to the target's parent, taking at each level
   * every activation and every completed repeat iteration, not only the
   * current one. A repeat resets a container's children for replay, but
   * what they already did stays in the syncbase history. The shared chain
   * keeps both sides on one time base, and no ancestor's state is needed.
   * Ancestors shared with `frame` contribute only their current iteration
   * (SMIL: a repeat resets its own subtree's references).
   */
  private historyFor(target: TimeNode, frame: Frame): Frame[] {
    const ancestors = new Set<TimeNode>()
    for (let p = this.ix.parent.get(target); p; p = this.ix.parent.get(p)) ancestors.add(p)
    let f: Frame | undefined = frame
    while (f && !(f.owner && ancestors.has(f.owner))) f = f.parent
    if (!f) return []
    const path: ContainerNode[] = []
    for (let p = this.ix.parent.get(target); p && p !== f.owner; p = this.ix.parent.get(p)) path.push(p)
    let frames: Frame[] = [f]
    for (let i = path.length - 1; i >= 0 && frames.length > 0; i--) {
      const next: Frame[] = []
      for (const g of frames) next.push(...this.historyChildren(path[i], g))
      frames = next.length > MAX_HISTORY ? next.slice(next.length - MAX_HISTORY) : next
    }
    return frames
  }

  /** Children frames of container `c` in `g`, oldest first: earlier activations and iterations, then the sampled one. */
  private historyChildren(c: ContainerNode, g: Frame): Frame[] {
    const current = this.sample(c, g).child
    if (current?.unbounded) return [current]
    const out: Frame[] = []
    for (const iv of this.intervalsIn(c, g)) {
      if (!(iv.begin <= g.tau)) break
      const finiteL = Number.isFinite(iv.L) && iv.L > 0
      let last: number
      if (current?.link?.iv === iv) last = current.link.k - 1
      else {
        const done = activeElapsed(iv, iv.end)
        last = finiteL && Number.isFinite(done) ? position(done, iv.L, true).k : 0
      }
      for (let k = Math.max(0, last - MAX_HISTORY + 1); k <= last; k++) {
        out.push({
          key: `${g.key}/${this.ivId(iv)}.${k}`,
          parent: g,
          owner: c,
          link: { iv, k, ctn: c.ctn },
          // A completed iteration reached every local time up to d.
          tau: finiteL ? iv.d : activeElapsed(iv, iv.end),
          frozen: true,
          visible: false,
          unbounded: false,
          beganBy: k === 0 ? iv.event : undefined,
        })
      }
    }
    if (current) out.push(current)
    return out.length > MAX_HISTORY ? out.slice(out.length - MAX_HISTORY) : out
  }

  /** `endTimeOf`, inside `drive`: the end of the last activation, in whichever ancestor iteration it happened. */
  endOf(node: TimeNode): number | 'indefinite' | undefined {
    let frames: Frame[]
    if (node === this.ix.root) frames = [this.rootFrame]
    else {
      const top = this.stateOf(this.ix.root)?.child
      if (!top) return undefined
      frames = this.historyFor(node, top)
    }
    for (let i = frames.length - 1; i >= 0; i--) {
      const f = frames[i]
      const ivs = this.intervalsIn(node, f)
      const last = ivs[ivs.length - 1]
      if (!last) continue
      if (f.unbounded) return 'indefinite'
      const end = lastAbs(f, last.begin, last.end)
      return Number.isFinite(end) ? end : 'indefinite'
    }
    return undefined
  }
}

/** Finite events up to `t`, in time order (stable, so equal inputs give equal outputs). */
function prepare(events: readonly ShowEvent[], t: number): ShowEvent[] {
  return events.filter((e) => Number.isFinite(e.t) && e.t <= t).sort((a, b) => a.t - b.t)
}

export function createScheduler(tree: TimingTree, options: SchedulerOptions = {}): Scheduler {
  const ix = buildIndex(tree)
  const reported = new Set<string>()
  const report: Report = (code, node) => {
    const key = `${code}|${ix.order.get(node)}`
    if (!options.sink || reported.has(key)) return
    reported.add(key)
    const message =
      code === 'timing-cycle'
        ? `time node ${node.ctn.id} is part of a reference cycle; the reference that closes it never fires`
        : code === 'timing-forward-reference'
          ? `time node ${node.ctn.id} is referenced before its sequence reaches it; the reference never fires`
          : `time node ${node.ctn.id} is behind a reference chain too deep to resolve`
    options.sink.add({ code, slideIndex: options.slideIndex, effect: `node ${node.ctn.id}`, message })
  }
  return {
    evaluate(events, t) {
      const active: SchedulerEntry[] = []
      const frozen: SchedulerEntry[] = []
      const ended = new Set<number>()
      if (Number.isNaN(t)) return { active, frozen, ended }
      const run = new Run(ix, prepare(events, t), t, report)
      for (const node of ix.list) {
        const s = run.drive(() => run.stateOf(node))
        if (!s || s.status === 'idle') continue
        if (s.status === 'ended') ended.add(node.ctn.id)
        if (isContainer(node) || !s.visible) continue
        const entry: SchedulerEntry = { node, progress: s.progress, iteration: s.iteration, order: ix.order.get(node) }
        ;(s.status === 'active' ? active : frozen).push(entry)
      }
      return { active, frozen, ended }
    },
    endTimeOf(nodeId, events) {
      const node = ix.byId.get(nodeId)
      if (!node) return undefined
      const run = new Run(ix, prepare(events, INF), INF, report)
      return run.drive(() => run.endOf(node))
    },
  }
}
