/**
 * Slide residency and the fallback ladder (slideshow spec 7).
 *
 * The current, previous and next slides are resident; the rest are evicted.
 * A resident slide holds its metadata handle (S1-7: eviction removes the
 * reference, so the store can refund it) and one layer lease, reserved
 * **before** its layers are built. The layer builder allocates through a
 * sub-owner carved from that lease, so it never reaches the ledger directly.
 *
 * When a slide does not fit: cross-slide media is stopped, then neighbours are
 * evicted (only for the current slide), then the ladder steps down — reduced
 * resolution, a flattened slide, the terminal fallback — each step diagnosed.
 * The terminal fallback allocates nothing: it paints into the permanent
 * reserve taken at show start, which is never lent.
 */
import type { AllocationOwner, DiagnosticSink, Lease, MetadataHandle, SlideMetadataStore, SlideShowWorkspace, TerminalReserve } from './types'
import { backingFor, createSubOwner, surfaceBytes, type Backing, type ClosableSubOwner, type ResizeTicket, type SlideShowTerminalReserve } from './workspace'

export type LadderLevel = 'full' | 'reduced' | 'flattened' | 'terminal'
const LADDER: readonly LadderLevel[] = ['full', 'reduced', 'flattened', 'terminal']

export interface Residency {
  /** Makes `slide` resident at `level` or the nearest lower rung that fits; rejects `'cancelled'` if the slide is left first. */
  ensure(slide: number, level: LadderLevel): Promise<LadderLevel>
  /** Drops a slide's layers and metadata. The current slide keeps its metadata while it is shown. */
  evict(slide: number): void
  /** Stops cross-slide media, then drops every neighbour's metadata and layers; never the current slide's. */
  onPressure(): void
  /**
   * Every new output takes its reserve before the old is released (old + new at the peak, smaller included);
   * if it cannot, the old backing store is kept with a diagnostic. On `'resized'` both stay charged until
   * `commitResize`; a resize while one is pending keeps the old store.
   */
  resize(css: { w: number; h: number }, dpr: number): 'resized' | 'kept-old'
  /** The new output surface has replaced the old one: the old output is refunded. */
  commitResize(): void
  /** Moves the window to `slide - 1 .. slide + 1`, evicting everything outside it. */
  show(slide: number): void
  /** The slides currently holding (or acquiring) residency, ascending. */
  resident(): number[]
  /** The sub-owner a resident slide's layers allocate from; `undefined` if it holds no layer lease. */
  owner(slide: number): AllocationOwner | undefined
  readonly backing: Backing
  /** Releases every slide, the current one included. The terminal reserve stays with its owner. */
  dispose(): void
}

export interface ResidencyDeps {
  sink?: DiagnosticSink
  /** The admitted backing store (from `admitShow`). */
  backing?: Backing
  /** Bytes a slide's layers need at a rung, reserved before they are built (the stage prices its render plan). */
  cost?: (slide: number, level: Exclude<LadderLevel, 'terminal'>, backing: Backing) => number
  /** Builds a slide's layers, allocating only through `owner`. */
  build?: (slide: number, level: Exclude<LadderLevel, 'terminal'>, owner: AllocationOwner) => Promise<void>
  /** Stops and releases media playing across slides (first relief under pressure). */
  stopCrossSlideMedia?: () => void
}

interface Entry {
  slide: number
  abort: AbortController
  acquiring?: Promise<MetadataHandle>
  handle?: MetadataHandle
  lease?: Lease
  /** The rung the slide's layers are built at now; retired generations settle on their own. */
  gen?: Generation
  dead: boolean
}

/**
 * One rung's reservation for one slide: a workspace lease and the sub-owner its layers allocate from. A rung
 * change makes a new generation; the old one is retired, and its charge follows its live allocations down,
 * refunded only when the last holder (an in-flight build or an outstanding layer lease) lets go.
 */
interface Generation {
  level: Exclude<LadderLevel, 'terminal'>
  lease: Lease
  owner: ClosableSubOwner
  builds: number
  retired: boolean
}

const cancelled = () => new Error('cancelled')

/**
 * A conservative default price when the stage injects none: full keeps the
 * layers plus their upload peak and a transition surface (3 x output), reduced
 * the same at DPR 1, flattened one full-slide frame and its texture at DPR 1.
 */
function defaultCost(level: Exclude<LadderLevel, 'terminal'>, backing: Backing): number {
  const out = surfaceBytes(backing.w, backing.h)
  const scale = backing.dpr > 1 ? backing.dpr : 1
  const atOne = surfaceBytes(backing.w / scale, backing.h / scale)
  return level === 'full' ? 3 * out : level === 'reduced' ? 3 * atOne : 2 * atOne
}

export function createResidency(ws: SlideShowWorkspace, store: SlideMetadataStore, terminal: TerminalReserve, deps: ResidencyDeps = {}): Residency {
  const entries = new Map<number, Entry>()
  let current: number | undefined
  // Without an admitted backing, price against a DPR 1 surface of the reserve's output bytes.
  let backing: Backing = deps.backing ?? { w: Math.ceil(terminal.outputBytes / 4), h: 1, dpr: 1 }
  const sink = deps.sink
  const price = (slide: number, level: Exclude<LadderLevel, 'terminal'>) => (deps.cost ? deps.cost(slide, level, backing) : defaultCost(level, backing))
  const inWindow = (slide: number) => current !== undefined && Math.abs(slide - current) <= 1

  let pendingResize: ResizeTicket | undefined

  /** A retired generation keeps exactly what is still held: everything while a build runs, then its live leases. */
  const settle = (g: Generation) => {
    if (!g.retired || g.builds > 0) return
    const live = g.owner.used()
    const extra = g.lease.bytes - live
    if (extra > 0) g.lease.shrink(extra)
    if (live === 0) g.lease.release()
  }
  const retire = (g: Generation) => {
    if (g.retired) return
    g.retired = true
    g.owner.close()
    settle(g)
  }
  const idle = (g: Generation) => g.builds === 0 && g.owner.used() === 0
  const newGeneration = (slide: number, level: Generation['level'], bytes: number): Generation | undefined => {
    const lease = ws.lease(bytes, `slide:${slide}:${level}`)
    if (!lease) return undefined
    const g: Generation = { level, lease, owner: undefined!, builds: 0, retired: false }
    g.owner = createSubOwner(bytes, () => settle(g))
    return g
  }
  const releaseLayers = (entry: Entry) => {
    if (entry.gen) retire(entry.gen)
    entry.gen = undefined
  }
  /** Ends an entry: metadata goes now; layers are retired and refunded as their holders let go. */
  const drop = (entry: Entry) => {
    entry.dead = true
    if (entries.get(entry.slide) === entry) entries.delete(entry.slide)
    entry.abort.abort()
    entry.handle?.release()
    entry.handle = undefined
    releaseLayers(entry)
    store.evict(entry.slide)
  }

  const evictNeighbours = (keep: number) => {
    let freed = false
    for (const entry of [...entries.values()].sort((a, b) => a.slide - b.slide)) {
      if (entry.slide === current || entry.slide === keep) continue
      drop(entry)
      freed = true
    }
    // The store may hold neighbour entries residency never acquired (another consumer's): drop their cache too.
    for (const slide of store.resident()) if (slide !== current && slide !== keep) { store.evict(slide); freed = true }
    return freed
  }

  const tryRung = (entry: Entry, rung: Generation['level'], bytes: number): boolean => {
    const old = entry.gen
    if (old && old.level === rung) {
      // Same rung: grow in place if needed; never shrink under live allocations.
      const held = old.lease.bytes
      if (bytes > held) {
        if (!old.lease.grow(bytes - held)) return false
        old.owner.setCapacity(bytes)
      }
      return true
    }
    // An idle old rung holds nothing live: refund it first, so the change needs no peak.
    if (old && idle(old)) releaseLayers(entry)
    const next = newGeneration(entry.slide, rung, bytes)
    if (!next) return false
    releaseLayers(entry)
    entry.gen = next
    return true
  }

  const fit = (entry: Entry, level: LadderLevel): LadderLevel => {
    const isCurrent = entry.slide === current
    let relief = isCurrent ? 0 : 2 // 0: media not stopped yet, 1: neighbours not evicted yet, 2: nothing left
    for (let i = LADDER.indexOf(level); i < LADDER.length; i++) {
      const rung = LADDER[i]!
      if (rung === 'terminal') {
        releaseLayers(entry)
        if (level !== 'terminal' && isCurrent) sink?.add({ code: 'ladder-terminal', slideIndex: entry.slide, message: 'the slide is painted directly into the terminal reserve; effects and slide changes become cuts' })
        return 'terminal'
      }
      const bytes = price(entry.slide, rung)
      for (;;) {
        if (tryRung(entry, rung, bytes)) return rung
        if (relief === 0) { relief = 1; if (deps.stopCrossSlideMedia) { deps.stopCrossSlideMedia(); continue } }
        if (relief === 1) { relief = 2; if (evictNeighbours(entry.slide)) continue }
        break
      }
      const next = LADDER[i + 1]!
      // Only a shown slide is diagnosed: a neighbour's prefetch is rebuilt when it becomes current.
      if (next !== 'terminal' && isCurrent) {
        sink?.add({
          code: `ladder-${next}`,
          slideIndex: entry.slide,
          message: next === 'reduced' ? 'the slide does not fit: its layers are rebuilt at reduced resolution' : 'the slide does not fit: it is flattened into one layer per checkpoint, with cuts and a crossfade',
        })
      }
    }
    return 'terminal'
  }

  const residency: Residency = {
    get backing() { return backing },
    async ensure(slide, level) {
      if (!LADDER.includes(level)) throw new RangeError(`unknown ladder level: ${String(level)}`)
      if (current === undefined) current = slide
      if (!inWindow(slide)) throw cancelled()
      let entry = entries.get(slide)
      if (!entry) {
        entry = { slide, abort: new AbortController(), dead: false }
        entries.set(slide, entry)
      }
      const mine = entry
      if (!mine.handle) {
        mine.acquiring ??= store.acquire(slide, mine.abort.signal)
        let handle: MetadataHandle
        try {
          handle = await mine.acquiring
        } catch (error) {
          // A failed or aborted read never poisons the slide: the next ensure starts afresh.
          mine.acquiring = undefined
          if (!mine.dead && entries.get(slide) === mine) entries.delete(slide)
          throw mine.dead ? cancelled() : error
        }
        // Completed after the slide was left: the reference is dropped, the result discarded.
        if (mine.dead) { handle.release(); throw cancelled() }
        if (mine.handle && mine.handle !== handle) handle.release()
        else mine.handle = handle
      }
      const got = fit(mine, level)
      const g = mine.gen
      if (got === 'terminal' || !deps.build || !g) return got
      g.builds++
      try {
        await deps.build(slide, got, g.owner)
      } finally {
        g.builds--
        // Superseded or left while building: what the build no longer holds is refunded now.
        settle(g)
      }
      if (mine.dead || mine.gen !== g) throw cancelled()
      return got
    },
    evict(slide) {
      const entry = entries.get(slide)
      if (!entry) { if (slide !== current) store.evict(slide); return }
      if (slide === current) { releaseLayers(entry); return }
      drop(entry)
    },
    onPressure() {
      deps.stopCrossSlideMedia?.()
      evictNeighbours(current ?? Number.NaN)
    },
    resize(css, dpr) {
      if (!(Number.isFinite(css.w) && Number.isFinite(css.h) && css.w > 0 && css.h > 0)) {
        sink?.add({ code: 'resize-kept-old', message: 'the new output size is not presentable; the old backing store is kept' })
        return 'kept-old'
      }
      const keep = () => {
        sink?.add({ code: 'resize-kept-old', message: 'the workspace cannot hold the new reserve next to the old one; the old backing store is kept and scaled' })
        return 'kept-old' as const
      }
      if (pendingResize) return keep()
      const next = backingFor(css, dpr)
      const bytes = surfaceBytes(next.w, next.h)
      const begin = (terminal as Partial<SlideShowTerminalReserve>).beginResize
      if (begin) {
        const ticket = begin.call(terminal, bytes)
        if (!ticket) return keep()
        pendingResize = ticket
      } else if (!terminal.resize(bytes)) return keep()
      backing = next
      return 'resized'
    },
    commitResize() {
      const ticket = pendingResize
      pendingResize = undefined
      ticket?.commit()
    },
    show(slide) {
      current = slide
      for (const entry of [...entries.values()]) if (!inWindow(entry.slide)) drop(entry)
    },
    resident: () => [...entries.keys()].sort((a, b) => a - b),
    owner: slide => entries.get(slide)?.gen?.owner,
    dispose() {
      for (const entry of [...entries.values()]) drop(entry)
    },
  }
  return residency
}
