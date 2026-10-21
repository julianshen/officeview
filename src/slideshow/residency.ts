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
import { backingFor, createSubOwner, surfaceBytes, type Backing, type ClosableSubOwner } from './workspace'

export type LadderLevel = 'full' | 'reduced' | 'flattened' | 'terminal'
const LADDER: readonly LadderLevel[] = ['full', 'reduced', 'flattened', 'terminal']

export interface Residency {
  /** Makes `slide` resident at `level` or the nearest lower rung that fits; rejects `'cancelled'` if the slide is left first. */
  ensure(slide: number, level: LadderLevel): Promise<LadderLevel>
  /** Drops a slide's layers and metadata. The current slide keeps its metadata while it is shown. */
  evict(slide: number): void
  /** Stops cross-slide media, then drops every neighbour's metadata and layers; never the current slide's. */
  onPressure(): void
  /** A larger output takes its reserve before the old is released; if it cannot, the old backing store is kept. */
  resize(css: { w: number; h: number }, dpr: number): 'resized' | 'kept-old'
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
  owner?: ClosableSubOwner
  building: number
  dead: boolean
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

  const releaseLayers = (entry: Entry) => {
    entry.owner?.close()
    entry.owner = undefined
    entry.lease?.release()
    entry.lease = undefined
  }
  /** Ends an entry: metadata and layers go now, unless a build is still allocating (then on its completion). */
  const drop = (entry: Entry) => {
    entry.dead = true
    if (entries.get(entry.slide) === entry) entries.delete(entry.slide)
    entry.abort.abort()
    entry.handle?.release()
    entry.handle = undefined
    if (entry.building === 0) releaseLayers(entry)
    else entry.owner?.close()
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

  const tryLease = (entry: Entry, bytes: number): boolean => {
    if (!entry.lease) {
      const lease = ws.lease(bytes, `slide:${entry.slide}:layers`)
      if (!lease) return false
      entry.lease = lease
      entry.owner = createSubOwner(bytes)
      return true
    }
    const held = entry.lease.bytes
    if (bytes > held && !entry.lease.grow(bytes - held)) return false
    if (bytes < held) entry.lease.shrink(held - bytes)
    entry.owner!.setCapacity(bytes)
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
        if (tryLease(entry, bytes)) return rung
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
        entry = { slide, abort: new AbortController(), building: 0, dead: false }
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
      if (got === 'terminal' || !deps.build || !mine.owner) return got
      mine.building++
      try {
        await deps.build(slide, got, mine.owner)
      } finally {
        mine.building--
        // The build finished after its slide was left: its reservation is released now.
        if (mine.dead && mine.building === 0) releaseLayers(mine)
      }
      if (mine.dead) throw cancelled()
      return got
    },
    evict(slide) {
      const entry = entries.get(slide)
      if (!entry) { if (slide !== current) store.evict(slide); return }
      if (slide === current) { if (entry.building === 0) releaseLayers(entry); return }
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
      const next = backingFor(css, dpr)
      if (!terminal.resize(surfaceBytes(next.w, next.h))) {
        sink?.add({ code: 'resize-kept-old', message: 'the workspace cannot hold a larger reserve; the old backing store is kept and scaled up' })
        return 'kept-old'
      }
      backing = next
      return 'resized'
    },
    show(slide) {
      current = slide
      for (const entry of [...entries.values()]) if (!inWindow(entry.slide)) drop(entry)
    },
    resident: () => [...entries.keys()].sort((a, b) => a - b),
    owner: slide => entries.get(slide)?.owner,
    dispose() {
      for (const entry of [...entries.values()]) drop(entry)
    },
  }
  return residency
}
