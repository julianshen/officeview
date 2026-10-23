/**
 * The presenter's next-checkpoint preview and mirror (slideshow spec 5.2
 * "Current" and "Next"; plan S2-4).
 *
 * **Next** is a static preview of the settled end state of the next click
 * checkpoint: the next build of the current slide, or the next visible
 * slide's first checkpoint. It is replayed from the slide's checkpoint
 * timeline (its own scheduler state; the live never is touched) by asking
 * that timeline's `next` at the caller's clock time: the scheduler answers
 * with the same cursor when the slide has nothing left, so identity decides
 * slide-advance against next-in-slide — the S1-10 contract, kept. On the
 * last step the only Next left is the main sequence's settle point, which
 * begins only while a finite effect still runs; that judgment belongs to the
 * show clock, so `nextCheckpoint` takes it as its required trailing argument (the
 * same `clock.now()` the session passes). The returned position is the
 * settled cursor `timeline.checkpoint(step)`, whose snapshot applies every
 * zero-duration effect and holds never-ending ones at their first
 * iteration's end (S1-10 checkpoint semantics, verbatim). Computing it is
 * pure: the live cursor, its event history, media and the store are never
 * touched.
 *
 * **Current** is mirrored by blitting the already-rendered audience frame,
 * so the animation in progress shows without painting a second timeline.
 *
 * Where the timing comes from. A slide's timing tree is read lazily by the
 * metadata store, so the synchronous `nextCheckpoint` cannot parse it. Call
 * `primePreview` for a slide first (it acquires the metadata handle, builds
 * the timeline once and holds it); until then the slide answers the honest
 * slide-advance position, like the session does while metadata loads.
 * Concurrent primes of one slide share one in-flight read, and releasing
 * while a prime is in flight returns the fresh handle instead of pinning it.
 * Holds are per slide: `releasePreview(model, slide)` hands back the slide
 * the preview has left, the whole-show form drops every hold, and
 * `createPresenterSurfaces` takes an optional `onClose` that runs on
 * `dispose()` (last, guarded) so the hold's lifetime ends with the surfaces
 * it belongs to. Releases free the preview's handles only — the store keeps
 * its normal cached entry (spec 7: the shown slide's metadata is never
 * dropped while it is shown) and refunds it under its own pressure rules.
 * Additively to the brief, S2-5 reads
 * `previewTimeline`/`previewHandle` to paint the primed slide.
 */
import { isPackageError } from '../../core/package-error'
import { nextVisible } from '../player/order'
import { createSlideTimeline, type SlideCursor, type SlideTimeline, type SlideTimelineOptions } from '../engine/steps'
import type { PptxSlide } from '../../pptx/types'
import type { AllocationOwner, DiagnosticSink, Lease, MetadataHandle, RenderState, ShowEvent, ShowModel, ShowSlide, SlideMetadataStore, SlideShowWorkspace, TimeNode, TimingTree } from '../types'
import { backingFor, surfaceBytes, type Backing } from '../workspace'
import { paintDirect, type DirectAssets } from '../render/direct'
import { slideSize } from '../render/scene'

export interface PreviewOptions extends SlideTimelineOptions {
  /**
   * The stage's expanded text timing for slide `index` (S2-8 `Stage.textTiming`): by-word and by-letter builds
   * fanned out over their items, with the source pre-order. The preview schedules the same tree and order as the
   * live session, so Next shows the build the click will play (PR #19 review 2). `undefined`: the slide's own timing.
   */
  textTiming?(index: number, meta: ShowSlide, charge: (bytes: number, tag: string) => boolean): { tree: TimingTree; orderOf: ReadonlyMap<TimeNode, number> } | undefined
}

export interface PreviewPosition {
  slide: number
  cursor: SlideCursor
}

export interface PreviewTarget {
  slide: PptxSlide
  /** The checked render state: a checkpoint settle for Next, or the live stage's state. */
  state: RenderState
  /** The paint's device-pixel backing (the leased surface's `width`/`height`). */
  backing: { w: number; h: number }
}

// ---------------------------------------------------------------------------
// The preview's timing hold
// ---------------------------------------------------------------------------

interface Held {
  handle: MetadataHandle
  timeline: SlideTimeline
}

const heldByStore = new WeakMap<SlideMetadataStore, Map<number, Held>>()
const primesInFlight = new WeakMap<SlideMetadataStore, Map<number, Promise<boolean>>>()
/** Bumped by every whole-show release, so a prime started before it is not stored after it. */
const holdGeneration = new WeakMap<SlideMetadataStore, number>()

/**
 * Acquires `slide`'s metadata and builds its preview timeline, holding the
 * handle so its repeated `nextCheckpoint` calls stay synchronous. `true` when
 * the timeline is available (a static slide included: its timeline answers
 * the slide advance). A store admission refusal (`resource-limit`, any other
 * package error this late) is diagnosed and answered `false` — the slide keeps
 * its un-primed answer; cancellation and closed-store errors and programming
 * errors are rethrown.
 */
export async function primePreview(model: ShowModel, slide: number, options: PreviewOptions, signal?: AbortSignal): Promise<boolean> {
  const store = model.store
  const held = heldByStore.get(store)
  if (held?.has(slide)) return true
  const inflight = primesInFlight.get(store)?.get(slide)
  if (inflight) return inflight
  const gen = holdGeneration.get(store) ?? 0
  const job = prime(model, slide, options, gen, signal).finally(() => {
    const m = primesInFlight.get(store)
    if (!m) return
    m.delete(slide)
    if (m.size === 0) primesInFlight.delete(store)
  })
  const map = primesInFlight.get(store) ?? new Map<number, Promise<boolean>>()
  map.set(slide, job)
  primesInFlight.set(store, map)
  return job
}

async function prime(model: ShowModel, slide: number, options: PreviewOptions, gen: number, signal?: AbortSignal): Promise<boolean> {
  const store = model.store
  let h: MetadataHandle
  try {
    h = await store.acquire(slide, signal)
  } catch (e) {
    // The store's own contract: cancellation leaves as the rejection the
    // caller armed (`cancelled`, or the abort reason); anything else that is
    // a package refusal is reported and answered `false`. A code-less error
    // is a programming error and is never swallowed.
    if (!isPackageError(e) || e.code === 'cancelled') throw e
    options.sink?.add({
      code: 'preview-prime', slideIndex: slide,
      message: `The preview could not read the slide's metadata (${e.message}); it shows the slide-advance position until the workspace has room`,
    })
    return false
  }
  // A refused timing tree arrives as `{ refused: string }` and createSlideTimeline
  // answers the static timeline for it; building never throws (S1-10).
  // Text by word or letter: the expanded tree and order the session schedules, charged to this hold's handle.
  const { textTiming, ...timelineOptions } = options
  const timing = h.slide.timing
  const text = timing && !('refused' in timing) && textTiming ? textTiming(slide, h.slide, (bytes, tag) => h.attach(bytes, tag)) : undefined
  const timeline = createSlideTimeline(text?.tree ?? timing, { ...timelineOptions, slideIndex: slide, ...(text ? { orderOf: text.orderOf } : {}) })
  const held = heldByStore.get(store)
  if (held?.has(slide)) { h.release(); return true } // a concurrent prime stored one first
  if ((holdGeneration.get(store) ?? 0) !== gen) {
    // Released while this prime was in flight: the fresh handle comes back,
    // and the store keeps its normal cached entry (spec 7).
    h.release()
    return false
  }
  const map = held ?? new Map<number, Held>()
  map.set(slide, { handle: h, timeline })
  heldByStore.set(store, map)
  return true
}

/** The primed slide's timeline, or `undefined` before `primePreview` resolved for it. */
export function previewTimeline(model: ShowModel, slide: number): SlideTimeline | undefined {
  return heldByStore.get(model.store)?.get(slide)?.timeline
}

/** The primed slide's held handle (its slide object, builds and media records ride on it). */
export function previewHandle(model: ShowModel, slide: number): MetadataHandle | undefined {
  return heldByStore.get(model.store)?.get(slide)?.handle
}

/**
 * Releases the preview's timing holds: one slide's (`slide` given — the
 * presenter calls this when its preview moves on), or every held slide of the
 * show. The store keeps its normal cached entry (spec 7: the shown slide's
 * metadata is never dropped while it is shown) and refunds it under its own
 * pressure rules; only the preview's handle is released. A prime still in
 * flight returns its fresh handle instead of pinning it (the store's refcount
 * covers the concurrency).
 */
export function releasePreview(model: ShowModel, slide?: number): void {
  const store = model.store
  const held = heldByStore.get(store)
  if (slide === undefined) {
    if (held) {
      for (const { handle } of held.values()) handle.release()
    }
    // Even with no holds yet, a whole-show release pulls the rug under any prime
    // still in flight: its fresh handle comes back instead of being stored.
    heldByStore.delete(store)
    holdGeneration.set(store, (holdGeneration.get(store) ?? 0) + 1)
    return
  }
  if (!held) return
  const entry = held.get(slide)
  if (!entry) return
  held.delete(slide)
  entry.handle.release()
}

// ---------------------------------------------------------------------------
// nextCheckpoint
// ---------------------------------------------------------------------------

const NO_EVENTS: readonly ShowEvent[] = Object.freeze([])
export const END_CURSOR: SlideCursor = Object.freeze({ step: 0, events: NO_EVENTS, mode: 'checkpoint' })

/**
 * Where the next click lands, as a position the preview paints: the next
 * build of `slide`, or the next visible slide's settled checkpoint 0, or
 * `'end'` of the visible order (a looping show restarts at the order's first
 * slide, like the session does). A slide that has not been primed (or whose
 * metadata could not be read) answers the slide advance.
 *
 * `t` is the show clock reading the click is judged at — the same
 * `clock.now()` the session passes to its own `next`; on the last step a
 * running finite effect means the next click is the settle point, so the
 * position stays on the slide, settled (`timeline.checkpoint(cursor.step)`),
 * while anything over ends the slide. The cursor is never changed: the replay
 * runs on the primed slide's own timeline, the live one's history stays where
 * it is.
 */
export function nextCheckpoint(model: ShowModel, order: readonly number[], slide: number, cursor: SlideCursor, t: number): PreviewPosition | 'end' {
  const timeline = previewTimeline(model, slide)
  if (!timeline) return advance(model, order, slide)
  if (!Number.isFinite(t) || t < 0) throw new RangeError(`slide timeline: time ${t}`)
  const last = cursor.events[cursor.events.length - 1]
  // The last step's only remaining Next is the settle point (S1-10): with the
  // finite effects over, the next click leaves the slide — the same predicate
  // the session's own next() judges at this clock reading.
  if (cursor.step >= timeline.steps && cursor.mode === 'live' && timeline.busyUntil(cursor) <= t) {
    return advance(model, order, slide)
  }
  // The step timeline answers the same cursor when nothing is left in the
  // slide: identity is the leave-the-slide test (S1-10 contract).
  const after = timeline.next(cursor, Math.max(last ? last.t : 0, t))
  if (after === cursor) return advance(model, order, slide)
  return { slide, cursor: timeline.checkpoint(after.step) }
}

function advance(model: ShowModel, order: readonly number[], slide: number): PreviewPosition | 'end' {
  const settings = model.index.settings
  const target = nextVisible(order, slide) ?? (settings.loop || settings.type === 'kiosk' ? order[0] : undefined)
  if (target === undefined) return 'end'
  // The primed target's own checkpoint 0 is the honest cursor; the frozen
  // stand-in replays identically on any timeline (a checkpoint cursor derives
  // its base from `nexts(events)`, so zero nexts at horizon is checkpoint 0).
  const tl = previewTimeline(model, target)
  return { slide: target, cursor: tl ? tl.checkpoint(0) : END_CURSOR }
}

// ---------------------------------------------------------------------------
// Paint
// ---------------------------------------------------------------------------

/**
 * Paints the preview letterboxed into the canvas: black bands, the slide at
 * its largest fit, then the assets (paintDirect draws the watermark last).
 * Effect-free: the state was computed elsewhere, so the live scheduler,
 * trigger history and media are never touched — but not resource-free: image
 * or layer work is admitted against `assets.scratch` (the presenter lease
 * owner; never the terminal reserve).
 */
export function paintPreview(ctx: CanvasRenderingContext2D, target: PreviewTarget, assets: DirectAssets): void {
  const { w, h } = slideSize(target.slide)
  const scale = Math.min(target.backing.w / w, target.backing.h / h)
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, target.backing.w, target.backing.h)
  ctx.translate((target.backing.w - w * scale) / 2, (target.backing.h - h * scale) / 2)
  paintDirect(ctx, target.slide, target.state, { ...assets, scale })
  ctx.restore()
}

/**
 * A fit-to-pane copy of the rendered audience frame: the mirror shows the
 * animation in progress from the same scheduler state without painting a
 * second one. The backing is cleared first, so bands reveal nothing stale
 * after the source's aspect changes; smoothing stays on at high quality for
 * the downscale.
 */
export function mirrorFrame(source: HTMLCanvasElement, ctx: CanvasRenderingContext2D): { x: number; y: number; w: number; h: number } | undefined {
  const backing = { w: ctx.canvas.width, h: ctx.canvas.height }
  if (!source.width || !source.height || !backing.w || !backing.h) return undefined
  const scale = Math.min(backing.w / source.width, backing.h / source.height)
  const dest = { x: (backing.w - source.width * scale) / 2, y: (backing.h - source.height * scale) / 2, w: source.width * scale, h: source.height * scale }
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.clearRect(0, 0, backing.w, backing.h)
  ctx.imageSmoothingEnabled = true
  ctx.imageSmoothingQuality = 'high'
  ctx.drawImage(source, dest.x, dest.y, dest.w, dest.h)
  ctx.restore()
  // Where the frame landed: video overlays are composited into this rectangle (PR #19 review 2).
  return dest
}

// ---------------------------------------------------------------------------
// Presenter surfaces: leased canvases for the presenter window
// ---------------------------------------------------------------------------

export interface PresenterSurfaces {
  readonly mirror: HTMLCanvasElement
  readonly preview: HTMLCanvasElement
  /** The "See all slides" thumbnail canvas for deck index `i`, created on first request. */
  grid(i: number): HTMLCanvasElement
  /** Frees the thumbnail for deck index `i` (S2-5 calls it as its grid scrolls); `false` when none is held, so a cap slot opens for the next request. */
  releaseGrid(i: number): boolean
  /** Re-cuts mirror and preview for the new pane; `false` leaves the surfaces as they came out. */
  resize(css: { w: number; h: number }, dpr: number): boolean
  /** Unloads the created nodes, releases every presenter lease, and runs the close callback once. */
  dispose(): void
  /** The presenter lease owner: every lease it grants is a `presenter:*` change, never the terminal reserve. */
  readonly scratch: AllocationOwner
}

/** The presenter window's assumed pane until the first `resize` carries the real one. */
const OPEN_CSS = { w: 480, h: 270 }
/** Grid thumbnails keep a fixed pane and DPR regardless of the mirror's size. */
const GRID_CSS = { w: 320, h: 180 }
const GRID_DPR = 1
/** The all-slides grid's thumbnail cap: 48 panes at DPR 1 bound it to ~11 MiB of the shared ledger. */
const GRID_MAX_THUMBS = 48
const PLACEHOLDER_ATTR = 'data-officeview-show-placeholder'
const SURFACE_ATTR = 'data-officeview-show-surface'
const REFUSAL_CODE = 'presenter-surface'
/** DPR steps tried after the capped one (the workspace admission ladder). */
const DPR_STEPS = [1.75, 1.5, 1.25, 1]

/**
 * The presenter window's surfaces (S2-5 drops them into its layout). Every
 * surface is leased from the workspace under a `presenter:*` tag before its
 * canvas is created, at a capped backing store that steps down the DPR ladder
 * (each step down is diagnosed) and, when nothing is left, stands down to a
 * zero-size placeholder with a diagnostic. A resize admits old + new leases
 * at the peak, so the replacement is whole before the old is released, and
 * the replaced canvas zeroed. `dispose()` releases every lease this object
 * granted — surfaces and scratch — refuses later ones, and runs `onClose`
 * once, which drops the preview's timing hold with the surfaces (pass
 * `() => releasePreview(model)`).
 */
export function createPresenterSurfaces(doc: Document, workspace: SlideShowWorkspace, sink?: DiagnosticSink, onClose?: () => void): PresenterSurfaces {
  const tell = (message: string, effect?: string) => {
    sink?.add({ code: REFUSAL_CODE, effect, message })
    return message
  }

  const placeholder = (name: string, message: string): HTMLCanvasElement => {
    // A zero-size canvas allocates no bitmap: the stopgap charges nothing.
    const c = doc.createElement('canvas')
    c.width = 0
    c.height = 0
    c.setAttribute(PLACEHOLDER_ATTR, 'presenter-surface')
    c.setAttribute(SURFACE_ATTR, name)
    c.textContent = message
    return c
  }

  const canvasFor = (name: string, backing: Backing): HTMLCanvasElement => {
    const c = doc.createElement('canvas')
    c.width = backing.w
    c.height = backing.h
    c.setAttribute(SURFACE_ATTR, name)
    return c
  }

  /** Backing candidates for a CSS pane: the capped DPR first, then lower ones, deduplicated. */
  const ladder = (css: { w: number; h: number }, dpr: number): Backing[] => {
    const first = backingFor(css, dpr)
    const candidates = [first, ...DPR_STEPS.filter(d => d < first.dpr).map(d => backingFor(css, d))]
    const seen = new Set<string>()
    return candidates.filter((b) => {
      const key = `${b.w}x${b.h}`
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })
  }

  /** Admits a fresh set of one-surface leases, all-or-nothing, while the old ones stay charged. */
  const admit = (tags: readonly string[], backing: Backing): Lease[] | undefined => {
    const leases: Lease[] = []
    for (const tag of tags) {
      const lease = workspace.lease(surfaceBytes(backing.w, backing.h), tag)
      if (!lease) {
        for (const l of leases) l.release()
        return undefined
      }
      leases.push(lease)
    }
    return leases
  }

  let live = true
  const rawDpr = doc.defaultView?.devicePixelRatio ?? 1
  const openDpr = Number.isFinite(rawDpr) && rawDpr > 0 ? rawDpr : 1
  if (openDpr !== rawDpr) tell(`This window reports device pixel ratio ${String(rawDpr)}; Presenter View uses 1`, 'dpr-clamped')

  const nodes = new Map<string, { canvas: HTMLCanvasElement; lease?: Lease }>()
  const gridCache = new Map<number, HTMLCanvasElement>()
  const scratchLive = new Set<Lease>()

  /** A wrapped work lease: dispose tracks it, and the wrapper stays delegating as it grows or shrinks. */
  const wrapped = (raw: Lease): Lease => ({
    tag: raw.tag,
    get bytes() { return raw.bytes },
    grow(more) { return raw.grow(more) },
    shrink(less) { raw.shrink(less) },
    transfer(next) {
      const moved = raw.transfer(next)
      scratchLive.delete(raw)
      scratchLive.add(moved)
      return wrapped(moved)
    },
    release() {
      scratchLive.delete(raw)
      raw.release()
    },
  })

  /** The DPR actually granted against the first candidate: a step down is diagnosed (spec 7). */
  const stepDown = (granted: Backing, first: Backing): void => {
    if (granted.dpr === first.dpr) return
    tell(
      `Presenter View runs at device pixel ratio ${String(granted.dpr)} instead of ${String(first.dpr)}: the workspace could not hold the larger backing store`,
      `dpr-stepdown:${String(granted.dpr)}`,
    )
  }

  const open = (): void => {
    const candidates = ladder(OPEN_CSS, openDpr)
    for (const backing of candidates) {
      const leases = admit(['presenter:mirror', 'presenter:preview'], backing)
      if (!leases) continue
      nodes.set('mirror', { canvas: canvasFor('presenter', backing), lease: leases[0]! })
      nodes.set('preview', { canvas: canvasFor('presenter', backing), lease: leases[1]! })
      stepDown(backing, candidates[0]!)
      return
    }
    // Nothing fits, even at DPR 1: the placeholder stands in, zero-cost, diagnosed.
    tell('Presenter View could not lease its surfaces; placeholders stand in until the workspace has room. Playback on the audience surface is unaffected')
    nodes.set('mirror', { canvas: placeholder('presenter', 'The Presenter View surface is a placeholder: the workspace could not hold the mirror at any device pixel ratio') })
    nodes.set('preview', { canvas: placeholder('presenter', 'The Presenter View surface is a placeholder: the workspace could not hold the preview at any device pixel ratio') })
  }
  open()

  return {
    get mirror() { return nodes.get('mirror')!.canvas },
    get preview() { return nodes.get('preview')!.canvas },
    get scratch() {
      // Work leases are presenter charges on the same ledger, tagged apart.
      return {
        lease: (bytes: number, tag: string) => {
          if (!live) {
            tell('Presenter View work lease refused: the surfaces are disposed', 'after-dispose')
            return undefined
          }
          const raw = workspace.lease(bytes, tag.startsWith('presenter:') ? tag : `presenter:${tag}`)
          if (!raw) return undefined
          scratchLive.add(raw)
          return wrapped(raw)
        },
      }
    },

    grid(i) {
      if (!live) {
        return placeholder('presenter-grid', tell('Presenter View is disposed: no all-slides thumbnail is leased', 'after-dispose'))
      }
      if (!Number.isInteger(i) || i < 0) {
        return placeholder('presenter-grid', tell(`Grid requested for slide index ${String(i)}, which is not a slide; a placeholder stands in`))
      }
      const cached = gridCache.get(i)
      if (cached) return cached
      if (gridCache.size >= GRID_MAX_THUMBS) {
        return placeholder('presenter-grid', tell(`Presenter View caps its all-slides thumbnails at ${String(GRID_MAX_THUMBS)}; slide ${String(i)} beyond the cap shows a placeholder until releaseGrid frees a slot`))
      }
      // Thumbnails are fixed at DPR 1: a 300-slide deck costs ~11 MiB, not the
      // workspace the audience surface needs.
      const backing = backingFor(GRID_CSS, GRID_DPR)
      const lease = workspace.lease(surfaceBytes(backing.w, backing.h), `presenter:grid:${String(i)}`)
      if (!lease) {
        // Not cached: the thumbnail retries when the workspace has room again.
        return placeholder('presenter-grid', tell(`Presenter View could not lease the all-slides thumbnail for slide ${String(i)}`))
      }
      const g = canvasFor('presenter-grid', backing)
      nodes.set(`presenter:grid:${String(i)}`, { canvas: g, lease })
      gridCache.set(i, g)
      return g
    },

    releaseGrid(i) {
      const tag = `presenter:grid:${String(i)}`
      const entry = nodes.get(tag)
      gridCache.delete(i)
      if (!entry) return false
      entry.lease?.release()
      entry.canvas.width = 0
      entry.canvas.height = 0
      entry.canvas.remove()
      nodes.delete(tag)
      return true
    },

    resize(nextCss, dpr) {
      if (!live) return false
      if (!Number.isFinite(nextCss.w) || !Number.isFinite(nextCss.h) || nextCss.w <= 0 || nextCss.h <= 0) {
        tell('Presenter View was resized to a size it cannot use; the surfaces are unchanged', 'invalid-resize')
        return false
      }
      if (!Number.isFinite(dpr) || dpr <= 0) {
        tell('Presenter View was resized with an invalid device pixel ratio; the surfaces are unchanged', 'invalid-resize')
        return false
      }
      const candidates = ladder(nextCss, dpr)
      for (const backing of candidates) {
        const leases = admit(['presenter:mirror', 'presenter:preview'], backing)
        if (!leases) continue
        const nextMirror = canvasFor('presenter', backing)
        const nextPreview = canvasFor('presenter', backing)
        // Swap first, then release and zero the replaced bitmaps: the replacement is whole
        // before the old charge goes back (spec 7).
        for (const [tag, entry] of nodes) {
          if (tag !== 'mirror' && tag !== 'preview') continue
          entry.lease?.release()
          entry.canvas.width = 0
          entry.canvas.height = 0
        }
        nodes.set('mirror', { canvas: nextMirror, lease: leases[0]! })
        nodes.set('preview', { canvas: nextPreview, lease: leases[1]! })
        stepDown(backing, candidates[0]!)
        return true
      }
      tell(`Presenter View could not lease surfaces of ${String(Math.ceil(nextCss.w))}x${String(Math.ceil(nextCss.h))} device px at any DPR; keeping what it holds`)
      return false
    },

    dispose() {
      if (!live) return
      live = false
      try {
        for (const raw of scratchLive) raw.release()
        scratchLive.clear()
        for (const [tag, entry] of nodes) {
          entry.lease?.release()
          // The bitmap is handed back with the lease, so nothing is retained uncharged.
          entry.canvas.width = 0
          entry.canvas.height = 0
          entry.canvas.remove()
          // Grid thumbnails are gone; mirror and preview stay as safe zero-size
          // canvases, so a pending frame's accessor gets a placeholder, not a throw.
          if (tag.startsWith('presenter:grid:')) nodes.delete(tag)
        }
        gridCache.clear()
      } finally {
        // Last and guarded: a throwing close callback cannot stop the teardown
        // it belongs to (the release above has already run either way).
        onClose?.()
      }
    },
  }
}
