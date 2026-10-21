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
 * Hostile input (spec 7): reference cycles, self references, missing ids
 * and over-deep resolution chains resolve to "never begins"; non-finite or
 * negative times are ignored. Nothing here reads a clock.
 */
import type { BehaviourNode, CTn, Condition, ContainerNode, ShowEvent, TimeNode, TimingTree } from '../types'
import { easeProgress } from './ease'

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
  /** Absolute end of the node's last activation; `undefined` when it never begins or the id is unknown. */
  endTimeOf(nodeId: number, events: readonly ShowEvent[]): number | 'indefinite' | undefined
}

const INF = Number.POSITIVE_INFINITY
/** Tree nesting the scheduler walks (the reader enforces the same limit). */
const MAX_DEPTH = 64
/** Nested reference resolution; deeper chains resolve to "never begins". */
const MAX_RESOLVE = 256
/** Activations per node (restart storms). */
const MAX_INTERVALS = 10_000

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

interface Frame {
  /** The container whose children live in this frame (undefined for the root's own frame). */
  readonly owner: ContainerNode | undefined
  /** Absolute show time of local 0 (the owner's current iteration start). */
  readonly origin: number
  /** Local sample time. */
  readonly tau: number
  /** The owner has ended at the sample: still-active children are cut. */
  readonly frozen: boolean
  /** The owner is active, or ended and showing its fill. */
  readonly visible: boolean
  /** The owner repeats forever and is sampled at infinity. */
  readonly unbounded: boolean
  /** Absolute end of the owner's activation (children are clipped to it). */
  readonly endAbs: number
  /** Absolute begin of each ancestor's current activation. */
  readonly ancestors: ReadonlyMap<TimeNode, number>
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
  busy: Set<TimeNode>
  adjusted?: Map<TimeNode, Interval[]>
  adjusting: boolean
  next: number
  seqBusy: boolean
  queue: Inst[]
  qi: number
}

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
function pauseSiblings(kids: readonly TimeNode[], raw: ReadonlyMap<TimeNode, Interval[]>): Map<TimeNode, Interval[]> {
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

/** One evaluation: memo tables live for a single `(events, t)`. */
class Run {
  readonly rootFrame: Frame
  private readonly groups = new Map<ContainerNode, Map<number, Group>>()
  private rootIvs: Interval[] | undefined
  private readonly states = new Map<TimeNode, Sample>()
  private readonly busyStates = new Set<TimeNode>()
  private depth = 0

  constructor(private readonly ix: Index, private readonly events: readonly ShowEvent[], t: number) {
    this.rootFrame = { owner: undefined, origin: 0, tau: t, frozen: false, visible: true, unbounded: false, endAbs: INF, ancestors: new Map() }
  }

  // -- sampling ---------------------------------------------------------------

  stateOf(node: TimeNode): Sample | undefined {
    const have = this.states.get(node)
    if (have) return have
    if (this.busyStates.has(node) || this.depth >= MAX_RESOLVE) return undefined
    this.busyStates.add(node)
    this.depth++
    const f = this.currentFrame(node)
    const s = f ? this.sample(node, f) : IDLE
    this.depth--
    this.busyStates.delete(node)
    this.states.set(node, s)
    return s
  }

  currentFrame(node: TimeNode): Frame | undefined {
    if (node === this.ix.root) return this.rootFrame
    const p = this.ix.parent.get(node)
    return p ? this.stateOf(p)?.child : undefined
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
      const origin = frame.origin + iv.begin + (finiteL && k > 0 ? k * iv.L : 0)
      // Children see the parent's eased (and mirrored) simple time.
      const tau2 = unbounded ? INF : finiteL ? (iv.d > 0 ? progress * iv.d : 0) : elapsed
      child = {
        owner: node,
        origin,
        tau: tau2,
        frozen: ended,
        visible,
        unbounded,
        endAbs: Math.min(frame.endAbs, frame.origin + iv.end),
        ancestors: new Map(frame.ancestors).set(node, origin),
        beganBy: k === 0 ? iv.event : undefined,
      }
    }
    return { status: ended ? 'ended' : 'active', visible, progress, iteration: k, child }
  }

  // -- intervals --------------------------------------------------------------

  intervalsIn(node: TimeNode, frame: Frame): readonly Interval[] {
    if (node === this.ix.root) {
      if (!this.rootIvs) {
        this.rootIvs = [] // cycle guard
        this.rootIvs = this.own(node, frame, 0)
      }
      return this.rootIvs
    }
    const owner = frame.owner
    if (!owner || this.ix.parent.get(node) !== owner) return []
    const g = this.group(owner, frame)
    if (owner.kind === 'seq') {
      this.seqUpTo(g, g.kids.length - 1) // a later advance can still settle an earlier step
      return g.raw.get(node) ?? []
    }
    if (owner.kind === 'excl') return this.exclChild(g, node)
    return this.parChild(g, node)
  }

  private group(owner: ContainerNode, frame: Frame): Group {
    let byOrigin = this.groups.get(owner)
    if (!byOrigin) this.groups.set(owner, (byOrigin = new Map()))
    let g = byOrigin.get(frame.origin)
    if (!g) {
      g = {
        owner,
        frame,
        kids: this.ix.kids.get(owner) ?? [],
        raw: new Map(),
        busy: new Set(),
        adjusting: false,
        next: 0,
        seqBusy: false,
        queue: owner.kind === 'seq' ? this.advanceQueue(owner, frame) : [],
        qi: 0,
      }
      byOrigin.set(frame.origin, g)
    }
    return g
  }

  private parChild(g: Group, node: TimeNode): Interval[] {
    const have = g.raw.get(node)
    if (have) return have
    if (g.busy.has(node)) return []
    g.busy.add(node)
    const ivs = this.own(node, g.frame, 0)
    g.busy.delete(node)
    g.raw.set(node, ivs)
    return ivs
  }

  private exclChild(g: Group, node: TimeNode): readonly Interval[] {
    if (g.adjusted) return g.adjusted.get(node) ?? []
    if (g.adjusting) return this.parChild(g, node)
    g.adjusting = true
    for (const k of g.kids) this.parChild(g, k)
    g.adjusted = pauseSiblings(g.kids, g.raw)
    g.adjusting = false
    return g.adjusted.get(node) ?? []
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
    for (const e of this.events) if (e !== frame.beganBy && e.t >= frame.origin && match(e)) out.push({ t: e.t - frame.origin, event: e })
    return out
  }

  private seqUpTo(g: Group, idx: number): void {
    if (g.seqBusy) return
    g.seqBusy = true
    while (g.next <= idx) this.seqStep(g, g.next++)
    g.seqBusy = false
  }

  private seqStep(g: Group, j: number): void {
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
    g.raw.set(kid, ivs)
    const first = ivs[0]
    if (!first) return
    const opts = g.owner.seq
    for (let i = 0; i < j; i++) {
      const list = g.raw.get(g.kids[i])
      const last = list?.[list.length - 1]
      if (!last || !(last.begin <= first.begin && first.begin < last.end)) continue
      if (advanced && opts?.nextAc === 'seek') this.seek(g.kids[i], last, first.begin, g.frame)
      else if (!opts?.concurrent) last.end = first.begin
    }
  }

  /** Settles a running step at local time `s`: finite parts jump to their end, indefinite ones keep running. */
  private seek(node: TimeNode, iv: Interval, s: number, frame: Frame): void {
    let settle: number
    if (Number.isFinite(iv.end)) settle = activeElapsed(iv, iv.end)
    else if (isContainer(node) && !Number.isFinite(iv.L)) {
      const cf = this.childFrame(node, frame, iv.begin, iv.event)
      settle = 0
      for (const c of this.ix.kids.get(node) ?? []) {
        const l = this.intervalsIn(c, cf)
        const e = l[l.length - 1]?.end
        if (e !== undefined && Number.isFinite(e) && e > settle) settle = e
      }
    } else return
    const delta = settle - activeElapsed(iv, s)
    if (!(delta > 0)) return
    iv.seeks.push({ at: s, delta })
    if (Number.isFinite(iv.end)) iv.end = Math.max(s, iv.end - delta)
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

  private childFrame(node: ContainerNode, frame: Frame, begin: number, event: ShowEvent | undefined): Frame {
    const origin = frame.origin + begin
    return { owner: node, origin, tau: INF, frozen: false, visible: false, unbounded: false, endAbs: INF, ancestors: new Map(frame.ancestors).set(node, origin), beganBy: event }
  }

  private interval(node: TimeNode, frame: Frame, inst: Inst): Interval {
    const b = inst.t
    const ctn = node.ctn
    const dur = time(ctn.durMs)
    let d: number
    if (isContainer(node) && (ctn.endSync || dur === undefined)) {
      const cf = this.childFrame(node, frame, b, inst.event)
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
        if (e.kind !== 'mediaEnd' || e.nodeId !== ctn.id || e.t < frame.origin) continue
        const t = e.t - frame.origin
        if (t > b && t < end) end = t
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
          for (const e of this.events) if (e.t >= frame.origin && match(e)) out.push({ t: e.t - frame.origin + dl, event: e })
        }
      }
    }
    return out
  }

  /** Begin or end times of node `tn`, local to `frame`. */
  private refTimes(tn: number, which: 'begin' | 'end', frame: Frame, from: TimeNode): number[] {
    const target = this.ix.byId.get(tn)
    if (!target || target === from) return []
    const anc = frame.ancestors.get(target)
    if (anc !== undefined) return which === 'begin' ? [anc - frame.origin] : []
    if (this.depth >= MAX_RESOLVE) return []
    this.depth++
    const out = this.refTimesOf(target, which, frame)
    this.depth--
    return out
  }

  private refTimesOf(target: TimeNode, which: 'begin' | 'end', frame: Frame): number[] {
    let ivs: readonly Interval[]
    let shift = 0
    if (frame.owner !== undefined && this.ix.parent.get(target) === frame.owner) ivs = this.intervalsIn(target, frame)
    else {
      const f = this.currentFrame(target)
      if (!f) return []
      ivs = this.intervalsIn(target, f)
      shift = f.origin - frame.origin
    }
    const out: number[] = []
    for (const iv of ivs) {
      const t = which === 'begin' ? iv.begin : iv.end
      if (Number.isFinite(t)) out.push(t + shift)
    }
    return out
  }
}

/** Finite events up to `t`, in time order (stable, so equal inputs give equal outputs). */
function prepare(events: readonly ShowEvent[], t: number): ShowEvent[] {
  return events.filter((e) => Number.isFinite(e.t) && e.t <= t).sort((a, b) => a.t - b.t)
}

export function createScheduler(tree: TimingTree): Scheduler {
  const ix = buildIndex(tree)
  return {
    evaluate(events, t) {
      const active: SchedulerEntry[] = []
      const frozen: SchedulerEntry[] = []
      const ended = new Set<number>()
      if (Number.isNaN(t)) return { active, frozen, ended }
      const run = new Run(ix, prepare(events, t), t)
      for (const node of ix.list) {
        const s = run.stateOf(node)
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
      const run = new Run(ix, prepare(events, INF), INF)
      const f = run.currentFrame(node)
      if (!f) return undefined
      const ivs = run.intervalsIn(node, f)
      const last = ivs[ivs.length - 1]
      if (!last) return undefined
      if (f.unbounded) return 'indefinite'
      const end = Math.min(f.origin + last.end, f.endAbs)
      return Number.isFinite(end) ? end : 'indefinite'
    },
  }
}
