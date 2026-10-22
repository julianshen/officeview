/**
 * The presenter's next-checkpoint preview and mirror (slideshow spec 5.2
 * "Current" and "Next"; plan S2-4).
 *
 * **Next** is a static preview of the settled end state of the next click
 * checkpoint: the next build of the current slide, or the next visible
 * slide's first checkpoint. It is replayed from the slide's checkpoint
 * timeline (its own scheduler state; the live never is touched) by asking
 * that timeline's `next` at the last event time: the scheduler answers with
 * the same cursor when the slide has nothing left, so identity decides
 * slide-advance against next-in-slide — the S1-10 contract, kept. The
 * returned position is the settled cursor `timeline.checkpoint(step)`, whose
 * snapshot applies every zero-duration effect and holds never-ending ones at
 * their first iteration's end (S1-10 checkpoint semantics, verbatim).
 *
 * **Current** is mirrored by blitting the already-rendered audience frame,
 * so the animation in progress shows without painting a second timeline.
 *
 * Where the timing comes from. A slide's timing tree is read lazily by the
 * metadata store, so the synchronous `nextCheckpoint` cannot parse it. Call
 * `primePreview` for a slide first (it acquires the metadata handle, builds
 * the timeline once and holds it); until then the slide answers the honest
 * slide-advance position, like the session does while metadata loads.
 * `releasePreview` frees the held handles when the preview changes for good
 * or the presenter closes. Additively to the brief, S2-5 reads
 * `previewTimeline`/`previewHandle` to paint the primed slide.
 */
import { nextVisible } from '../player/order'
import { createSlideTimeline, type SlideCursor, type SlideTimeline, type SlideTimelineOptions } from '../engine/steps'
import type { PptxSlide } from '../../pptx/types'
import type { AllocationOwner, DiagnosticSink, Lease, MetadataHandle, RenderState, ShowEvent, ShowModel, SlideMetadataStore, SlideShowWorkspace } from '../types'
import { backingFor, surfaceBytes, type Backing } from '../workspace'
import { paintDirect, type DirectAssets } from '../render/direct'
import { slideSize } from '../render/scene'

export interface PreviewOptions extends SlideTimelineOptions {}

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

/**
 * Acquires `slide`'s metadata and builds its preview timeline, holding the
 * handle so its repeated `nextCheckpoint` calls stay synchronous. `true` when
 * the timeline is available (a static slide included: its timeline answers
 * the slide advance); `false` sets the slide back to the un-primed answer(s)
 * and diagnoses, never silently. `store.acquire` rejects on a closed store or
 * cancellation.
 */
export async function primePreview(model: ShowModel, slide: number, options: PreviewOptions, signal?: AbortSignal): Promise<boolean> {
  const store = model.store
  const held = heldByStore.get(store)
  if (held?.has(slide)) return true
  let h: MetadataHandle
  try {
    h = await store.acquire(slide, signal)
  } catch (e) {
    if ((e as { code?: unknown })?.code !== undefined) throw e
    options.sink?.add({
      code: 'preview-prime', slideIndex: slide,
      message: `The preview could not read the slide's metadata (${e instanceof Error ? e.message : String(e)}); it previews as an un-primed slide until the store has room`,
    })
    return false
  }
  // A refused timing tree arrives as `{ refused: string }` and createSlideTimeline
  // answers the static timeline for it; building never throws (S1-10).
  const timeline = createSlideTimeline(h.slide.timing, { ...options, slideIndex: slide })
  const map = heldByStore.get(store) ?? new Map<number, Held>()
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

/** Releases every held preview handle of this show; the presenter calls it when the preview changes for good or closes. */
export function releasePreview(model: ShowModel): void {
  const store = model.store
  const held = heldByStore.get(store)
  if (!held) return
  for (const { handle } of held.values()) handle.release()
  heldByStore.delete(store)
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
 * metadata could not be read) answers the slide advance. The cursor is never
 * changed: the replay runs on the primed slide's own timeline, the live one's
 * history stays where it is.
 */
export function nextCheckpoint(model: ShowModel, order: readonly number[], slide: number, cursor: SlideCursor): PreviewPosition | 'end' {
  const timeline = previewTimeline(model, slide)
  if (!timeline) return advance(model, order, slide)
  const last = cursor.events[cursor.events.length - 1]
  // The step timeline answers the same cursor when nothing is left in the
  // slide: identity is the leave-the-slide test (S1-10 contract).
  const after = timeline.next(cursor, last ? last.t : 0)
  if (after === cursor) return advance(model, order, slide)
  return { slide, cursor: timeline.checkpoint(after.step) }
}

function advance(model: ShowModel, order: readonly number[], slide: number): PreviewPosition | 'end' {
  const settings = model.index.settings
  const target = nextVisible(order, slide) ?? (settings.loop || settings.type === 'kiosk' ? order[0] : undefined)
  if (target === undefined) return 'end'
  return { slide: target, cursor: END_CURSOR }
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
 * second one. Bands stay transparent like the backing they blit from.
 */
export function mirrorFrame(source: HTMLCanvasElement, ctx: CanvasRenderingContext2D): void {
  const backing = { w: ctx.canvas.width, h: ctx.canvas.height }
  if (!source.width || !source.height || !backing.w || !backing.h) return
  const scale = Math.min(backing.w / source.width, backing.h / source.height)
  ctx.save()
  ctx.setTransform(1, 0, 0, 1, 0, 0)
  ctx.drawImage(
    source,
    (backing.w - source.width * scale) / 2,
    (backing.h - source.height * scale) / 2,
    source.width * scale,
    source.height * scale,
  )
  ctx.restore()
}

// ---------------------------------------------------------------------------
// Presenter surfaces: leased canvases for the presenter window
// ---------------------------------------------------------------------------

export interface PresenterSurfaces {
  readonly mirror: HTMLCanvasElement
  readonly preview: HTMLCanvasElement
  /** The "See all slides" thumbnail canvas for deck index `i`, created on first request. */
  grid(i: number): HTMLCanvasElement
  /** Re-cuts mirror and preview for the new pane; `false` leaves the surfaces as they came out. */
  resize(css: { w: number; h: number }, dpr: number): boolean
  /** Unloads the created nodes and releases every presenter lease. */
  dispose(): void
  /** The presenter lease owner: every lease it grants is a `presenter:*` change, never the terminal reserve. */
  readonly scratch: AllocationOwner
}

/** The presenter window's assumed pane until the first `resize` carries the real one. */
const OPEN_CSS = { w: 480, h: 270 }
/** Grid thumbnails keep a fixed pane regardless of the mirror's size. */
const GRID_CSS = { w: 320, h: 180 }
const PLACEHOLDER_ATTR = 'data-officeview-show-placeholder'
const SURFACE_ATTR = 'data-officeview-show-surface'
const REFUSAL_CODE = 'presenter-surface'
/** DPR steps tried after the capped one (the workspace admission ladder). */
const DPR_STEPS = [1.75, 1.5, 1.25, 1]

/**
 * The presenter window's surfaces (S2-5 drops them into its layout). Every
 * surface is leased from the workspace under a `presenter:*` tag before its
 * canvas is created, at a capped backing store that steps down the DPR
 * ladder and, when nothing is left, stands down to a zero-size placeholder
 * with a diagnostic. A resize admits old + new leases at the peak, so the
 * replacement is whole before the old is released.
 */
export function createPresenterSurfaces(doc: Document, workspace: SlideShowWorkspace, sink?: DiagnosticSink): PresenterSurfaces {
  const tell = (message: string) => {
    sink?.add({ code: REFUSAL_CODE, message })
    return message
  }

  const placeholder = (name: string, message: string, tag?: string): HTMLCanvasElement => {
    // A zero-size canvas allocates no bitmap: the stopgap charges nothing.
    const c = doc.createElement('canvas')
    c.width = 0
    c.height = 0
    c.setAttribute(PLACEHOLDER_ATTR, 'presenter-surface')
    c.setAttribute(SURFACE_ATTR, name)
    if (tag) c.setAttribute('data-officeview-show-diagnostics', tag)
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
  const openDpr = doc.defaultView?.devicePixelRatio ?? 1
  if (!Number.isFinite(openDpr) || openDpr <= 0) throw new RangeError('presenter surfaces: device pixel ratio')

  const nodes = new Map<string, { canvas: HTMLCanvasElement; lease?: Lease }>()
  const gridCache = new Map<number, HTMLCanvasElement>()

  const open = (): void => {
    for (const backing of ladder(OPEN_CSS, openDpr)) {
      const leases = admit(['presenter:mirror', 'presenter:preview'], backing)
      if (!leases) continue
      const mirror = canvasFor('presenter', backing)
      const preview = canvasFor('presenter', backing)
      nodes.set('mirror', { canvas: mirror, lease: leases[0]! })
      nodes.set('preview', { canvas: preview, lease: leases[1]! })
      return
    }
    // Nothing fits, even at DPR 1: the placeholder stands in, zero-cost, diagnosed.
    const message = 'Presenter View could not lease its surfaces; placeholders stand in until the workspace has room. Playback on the audience surface is unaffected'
    nodes.set('mirror', { canvas: placeholder('presenter', message) })
    nodes.set('preview', { canvas: placeholder('presenter', message) })
  }
  open()

  return {
    get mirror() { return nodes.get('mirror')!.canvas },
    get preview() { return nodes.get('preview')!.canvas },
    get scratch() {
      // Work leases are presenter charges on the same ledger, tagged apart.
      return { lease: (bytes: number, tag: string) => workspace.lease(bytes, tag.startsWith('presenter:') ? tag : `presenter:${tag}`) }
    },

    grid(i) {
      if (!Number.isInteger(i) || i < 0) {
        return placeholder('presenter-grid', tell(`Grid requested for slide index ${String(i)}, which is not a slide; a placeholder stands in`), 'invalid-index')
      }
      const cached = gridCache.get(i)
      if (cached) return cached
      const tag = `presenter:grid:${String(i)}`
      for (const backing of ladder(GRID_CSS, openDpr)) {
        const lease = workspace.lease(surfaceBytes(backing.w, backing.h), tag)
        if (!lease) continue
        const g = canvasFor('presenter-grid', backing)
        nodes.set(tag, { canvas: g, lease })
        gridCache.set(i, g)
        return g
      }
      const g = placeholder('presenter-grid', tell(`Presenter View could not lease the all-slides thumbnail for slide ${String(i)}`), 'workspace-refused')
      gridCache.set(i, g)
      nodes.set(tag, { canvas: g })
      return g
    },

    resize(nextCss, dpr) {
      if (!live) return false
      if (!Number.isFinite(nextCss.w) || !Number.isFinite(nextCss.h) || nextCss.w <= 0 || nextCss.h <= 0) {
        tell('Presenter View was resized to a size it cannot use; the surfaces are unchanged')
        return false
      }
      if (!Number.isFinite(dpr) || dpr <= 0) {
        tell('Presenter View was resized with an invalid device pixel ratio; the surfaces are unchanged')
        return false
      }
      for (const backing of ladder(nextCss, dpr)) {
        const leases = admit(['presenter:mirror', 'presenter:preview'], backing)
        if (!leases) continue
        const mirrorLease = leases[0]!
        const previewLease = leases[1]!
        const nextMirror = canvasFor('presenter', backing)
        const nextPreview = canvasFor('presenter', backing)
        // Swap first, then release: the replacement is whole before the old charge goes back.
        nodes.get('mirror')?.lease?.release()
        nodes.get('preview')?.lease?.release()
        nodes.set('mirror', { canvas: nextMirror, lease: mirrorLease })
        nodes.set('preview', { canvas: nextPreview, lease: previewLease })
        return true
      }
      tell(`Presenter View could not lease surfaces of ${String(Math.ceil(nextCss.w))}x${String(Math.ceil(nextCss.h))} device px at any DPR; keeping what it holds`)
      return false
    },

    dispose() {
      if (!live) return
      live = false
      for (const [tag, entry] of nodes) {
        entry.lease?.release()
        entry.canvas.remove()
        nodes.delete(tag)
      }
      gridCache.clear()
    },
  }
}
