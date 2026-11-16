/**
 * The show-plan bundle build (piece 4 spec §5, §7; plan D1, D3, D5, D6, D7, D12, D13, D14, D16, D22).
 *
 * One animated slide becomes a bundle: the plan record (`ShowPlanRecord`), the entry frame (the slide before
 * step 1, the transition's to-frame) and one asset per layer the classifier keeps, in z order. Every asset
 * carries the visible mark; Layer-1 is applied per session at send time (`markLayerRaster` with each layer's
 * `rest`). Steps the layers cannot prove play `composited`, as keyframe tracks rendered on demand
 * (`tracks.ts`).
 *
 * The order inside one build:
 * 1. Strip what never plays streamed (`cmd` payloads, made inert in place, Task 14 ruling; the transition sound, piece 5) BEFORE preparing, so
 *    the server classifies the tree the client will play.
 * 2. D3 limits before any raster: over 64 steps (or an iterate over 128 keys, Task 4 ruling) the slide is its
 *    piece-3 still (`show-limits`).
 * 3. Rasterise (growth layers above the bucket, D12), classify (Task 5), mask (`layer × mask`), burn the
 *    visible mark at each layer's rest (D8), check the budget from `w·h·4` before any PNG is encoded.
 * 4. Validate the marks at every step boundary (§7, D7, D9) on the decoded shipped PNGs; failure plays every
 *    step composited.
 *
 * Caches (D6): L1 (no mark) holds the preparation, the unmarked rasters and the classification per
 * `strictMarks`; L2 (L1 + mark settings + `strictMarks`) holds the finished bundle and its verdict.
 */
import { createHash } from 'node:crypto'
import { showCanvasBackend } from './canvas'
import type { PptxDocument, PptxPackageHandle } from '../../../src/pptx/types'
import { createDiagnosticSink } from '../../../src/slideshow/diagnostics'
import { createSlideShowWorkspace } from '../../../src/slideshow/workspace'
import { SHOW_BUILD_BYTES, showBuildGrant, showMorphGrant } from './grant'
import { createSlideTimeline, type SlideTimeline } from '../../../src/slideshow/engine/steps'
import { hitRegionsOf, layoutTextBases, prepareSlide, rasterPreparedLayer, runRegionsOf, type PreparedSlide } from '../../../src/slideshow/prepare'
import type { PlanLayer } from '../../../src/slideshow/render/plan'
import { slideSize, textRangeBox } from '../../../src/slideshow/render/scene'
import { createRng } from '../../../src/slideshow/rng'
import { targetKey } from '../../../src/slideshow/target'
import type { DiagnosticSink, MetadataHandle, CanvasFactory, LayerBitmap, Lease, ShowIndex, ShowSlide, SlideMetadataStore, SlideShowWorkspace, TimeNode, TimingTree, TransitionSpec } from '../../../src/slideshow/types'
import { encodeShowMessage, PlanRejected, SHOW_DIAGNOSTIC_CODES, SHOW_PLAN_ITERATE_ITEMS, SHOW_PLAN_LIMITS, type PlanContainerRecord, type PlanLayerRecord, type PlanStepRecord, type Rect, type ShowPlanRecord } from '../../../src/stream/show-plan'
import { LruBytes, showPlanL1Key, showPlanL2Key } from '../cache'
import { basis } from '../mark/basis'
import { extendedCodeword } from '../mark/bch'
import { detect } from '../mark/detect'
import { markLayerRaster } from '../mark/mark-raster'
import { messageFor } from '../mark/message'
import { DEFAULT_PARAMS } from '../mark/params'
import { decodePng, encodePng } from '../mark/png'
import { markOptions, paintMark, type MarkSpec } from '../mark/visible'
import { tauFor } from '../detect/tau'
import { classifiedRasters, createSlideComposer, drawsAt, maskedRaster, slideStateAt, type DrawnLayer } from './compose'
import { allocateFrames, classifySlide, stepDurations, type Classification } from './eligibility'
import { burnVisibleMark, validateMarks } from './marks'

export interface ShowPlanJob {
  unitIndex: number
  scale: 1 | 2
  motion: 'full' | 'reduced'
  mark: MarkSpec
  strictMarks: boolean
  /**
   * Additive: a Layer-1 basis key for the §7 build-time validation. Controller ruling (Task 6 follow-up): a
   * fixed validation key (derived from the document key, unit and bucket) is accepted and is the default. The
   * validation proves that a page basis cropped to each layer's rest survives composition; the session's own
   * Layer-1 is applied per session at send time (Task 7, `markLayerRaster` with `asset.rest`), never here.
   */
  kBasis?: Uint8Array
}
/**
 * `rest` is the asset's top-left at its rest on the unit page, in page device px at the bucket, and `rest.scale`
 * the asset px per page device px there (additive to the brief: 1 for a layer at the bucket that rests where it
 * is painted; a growth layer, D12, or one resting scaled differs). Together they are the Layer-1 crop
 * (`markLayerRaster({ crop: { x, y, pageW, pageH, scale } })`). The entry frame rests at (0, 0), scale 1.
 */
export interface ShowAsset { tag: 2 | 3; assetId: number; w: number; h: number; png: Uint8Array; rest: { x: number; y: number; scale: number } }
/**
 * `assets`: the entry frame first, then the layers in z order (D1). `decodedBytes`: Σ `w·h·4` of the assets plus
 * the plan record's bytes.
 */
export interface BuiltShowPlan { plan: ShowPlanRecord; assets: ShowAsset[]; decodedBytes: number }
/**
 * Task 6 decision. `show-limits` (over 64 steps, an iterate over 128 keys, a record the codec would refuse),
 * `show-budget` (over 128 MiB, or one asset over 64 MiB), `show-plan-timeout` (over `SHOW_PLAN_DEADLINE_MS`).
 * `still`: the classifier could not prove the slide (`Classification.fallback`, e.g. unproven autoplay); it plays
 * its piece-3 still with the classifier's `diagnostics`. Every refusal plays the still; none ships a layer.
 */
export type ShowRefusal = { refused: 'show-budget' | 'show-limits' | 'show-plan-timeout' | 'still'; diagnostics: Array<{ code: string; step?: number }> }

export const SHOW_BUNDLE_BUDGET = SHOW_PLAN_LIMITS.showDecoded
export { SHOW_BUILD_BYTES, showBuildGrant, showMorphGrant }
export const SHOW_PLAN_DEADLINE_MS = 10_000
/** The entry frame's asset id; layer `i` (z order) is `1 + i`; track frames start at `TRACK_ASSET_BASE`. */
export const ENTRY_ASSET_ID = 0
export const TRACK_ASSET_BASE = 0x10000
const MAX_UNIT_PX = 8192
const L1_BYTES = 256 * 1024 * 1024
const L2_BYTES = 256 * 1024 * 1024

/** Test seams (never set in production). Any seam bypasses the L2 cache. */
export interface ShowBuildOptions {
  /** The bundle budget (default `SHOW_BUNDLE_BUDGET`). */
  budget?: number
  /** Replaces the layer composite of boundary `step` before validation (forced mismatch). */
  compose?: (step: number, composite: ImageData) => ImageData
  deadlineMs?: number
  now?: () => number
  /**
   * Production (not a seam): the bytes admission granted this job (`showBuildGrant`; the pool sends it). Absent,
   * the grant `onShow` would admit for the slide at the bucket.
   */
  grant?: number
  /** Test seam: the per-job ledger the preparation leases from (default `showLedger`). */
  ledger?: (grant: number) => SlideShowWorkspace
}

/** A finite per-job ledger of `grant` bytes (no terminal reserve): every preparation surface leases from it first. */
export const showLedger = (grant: number): SlideShowWorkspace => createSlideShowWorkspace('desktop', { limit: grant, scratch: 0 })

/**
 * The job's ledger: the given grant (a finite, non-negative number, at most `SHOW_BUILD_BYTES`), else what `onShow`
 * admits for it: the slide's `showBuildGrant`, or a Morph pair's `showMorphGrant` (`unit` = [from, to]).
 */
export function jobLedger(show: ShowDoc, unit: number | readonly [number, number], scale: number, opts: Pick<ShowBuildOptions, 'grant' | 'ledger'>): SlideShowWorkspace {
  const given = opts.grant
  const size = (u: number) => show.units[u] ?? { w: 0, h: 0 }
  const grant = given !== undefined && Number.isFinite(given) && given >= 0 ? Math.min(given, SHOW_BUILD_BYTES)
    : typeof unit === 'number' ? showBuildGrant(size(unit), scale) : showMorphGrant(size(unit[0]), size(unit[1]), scale)
  return (opts.ledger ?? showLedger)(grant)
}

interface L1Entry {
  /** Set when the slide cannot play from a plan whatever the marks: its refusal. */
  refusal?: ShowRefusal
  ps?: PreparedSlide
  timeline?: SlideTimeline
  record?: Omit<ShowPlanRecord, 'layers' | 'steps' | 'entry' | 'diagnostics'>
  extraDiagnostics?: Array<{ code: string }>
  rasters?: Map<PlanLayer, LayerBitmap>
  classes: Map<boolean, Classification>
  /**
   * Task 6 follow-up (controller ruling): a refusal found after preparation (`show-plan-timeout`, `still`,
   * `show-budget`, `show-limits`), per `strictMarks`, cached under the L1 key so a repeat request for the
   * slide, bucket and motion plays its still at once, whatever its mark.
   */
  refusals: Map<boolean, ShowRefusal>
}
interface Stats { l1: { hits: number; misses: number }; l2: { hits: number; misses: number } }

/** What `DocumentHost` keeps per open pptx for the show (the projected package: never the original). */
export interface ShowDoc {
  docRef: string
  docKey: string
  contentPolicyId: string
  fingerprint: string
  doc: PptxDocument
  handle: PptxPackageHandle
  index: ShowIndex
  store: SlideMetadataStore
  /** Acquires slide `i`'s metadata with its layout and master artwork in place for the resolver (as local mode). */
  acquire(i: number): Promise<MetadataHandle>
  /** The show reader's diagnostics (index, metadata, underlay). */
  sink: DiagnosticSink
  units: ReadonlyArray<{ w: number; h: number }>
  hidden: readonly number[]
  /** Loaded by the host before a build (async reads); the build itself is synchronous. */
  metas: Map<number, { slide: ShowSlide; charge: (bytes: number, tag: string) => boolean }>
  images?: ReadonlyArray<CanvasImageSource | undefined>
  l1: LruBytes<L1Entry>
  l2: LruBytes<BuiltShowPlan | ShowRefusal>
  stats: Stats
}

/** The `OpenDoc` view the build needs: its show inputs. */
export interface ShowHost { show?: ShowDoc }

export function createShowDoc(init: Omit<ShowDoc, 'metas' | 'l1' | 'l2' | 'stats'>): ShowDoc {
  return { ...init, metas: new Map(), l1: new LruBytes(L1_BYTES), l2: new LruBytes(L2_BYTES), stats: { l1: { hits: 0, misses: 0 }, l2: { hits: 0, misses: 0 } } }
}

class Timeout extends Error { constructor() { super('show-plan-timeout') } }

const factory: CanvasFactory = (w, h) => showCanvasBackend().createCanvas(Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))) as never
const resolveFont = (f: string) => f
const measureCtx = () => showCanvasBackend().createCanvas(1, 1).getContext('2d') as unknown as CanvasRenderingContext2D
const D21 = new Set<string>(SHOW_DIAGNOSTIC_CODES)
const MAX_DIAGNOSTICS = 1024

// ---------------------------------------------------------------------------
// What never plays streamed (D16)
// ---------------------------------------------------------------------------

/**
 * The tree with every `cmd` leaf made inert (Task 14 controller ruling, replacing D16's drop): the node keeps its
 * place, `cmdType` and timing (begin, duration, fill, restart), so the step and click structure is local mode's —
 * a long media call still holds its group, and a → during it seeks as locally — while its command payload (an OLE
 * verb, a media call, an event name) never leaves the server (`cmd: ''`).
 */
export function inertCommands(tree: TimingTree): { tree: TimingTree; stripped: number } {
  let stripped = 0
  const walk = (n: TimeNode, depth: number): TimeNode => {
    if (!('children' in n)) { if (n.kind === 'cmd') { stripped++; return { ...n, cmd: '' } } return n }
    if (depth > 256) return n
    const children = n.children.map(c => walk(c, depth + 1))
    return children.every((c, i) => c === n.children[i]) ? n : { ...n, children }
  }
  const root = walk(tree.root, 0)
  return { tree: root === tree.root ? tree : { ...tree, root }, stripped }
}

function streamedMeta(meta: ShowSlide): { meta: ShowSlide; cmdDropped: boolean } {
  let timing = meta.timing, cmdDropped = false
  if (timing && !('refused' in timing)) {
    const d = inertCommands(timing)
    timing = d.tree
    cmdDropped = d.stripped > 0
  }
  let transition: TransitionSpec | undefined = meta.transition
  if (transition?.sound) { const { sound: _s, ...rest } = transition; transition = rest }
  return { meta: { ...meta, timing, ...(transition ? { transition } : {}), notes: [], links: meta.links, media: [] }, cmdDropped }
}

// ---------------------------------------------------------------------------
// Growth (D12)
// ---------------------------------------------------------------------------

/** The largest `animScale` factor each target key reaches (from, to or by; at least 1). */
function growthTargets(tree: TimingTree | undefined): Map<string, number> {
  const out = new Map<string, number>()
  if (!tree) return out
  const walk = (n: TimeNode, depth: number) => {
    if (depth > 256) return
    if ('children' in n) { for (const c of n.children) walk(c, depth + 1); return }
    if (n.kind !== 'animScale' || !('spid' in n.target)) return
    const f = Math.max(1, ...[n.by, n.to, n.from].flatMap(p => (p ? [Math.abs(p.x), Math.abs(p.y)] : [])).filter(Number.isFinite))
    const k = targetKey(n.target)
    out.set(k, Math.max(out.get(k) ?? 1, f))
  }
  walk(tree.root, 0)
  return out
}

/** Each raster layer's growth factor: its own target's, or an enclosing group's (sub-layers move with it). */
function growthOf(ps: PreparedSlide, targets: ReadonlyMap<string, number>): Map<PlanLayer, number> {
  const out = new Map<PlanLayer, number>()
  if (targets.size === 0) return out
  const own = (key: string) => {
    let g = 1
    for (const [k, f] of targets) if (key === k || key.startsWith(`${k}/`)) g = Math.max(g, f)
    return g
  }
  const walk = (layers: readonly PlanLayer[], inherited: number) => {
    for (const l of layers) {
      const g = l.kind === 'target' ? Math.max(inherited, own(l.key)) : inherited
      if (l.kind === 'target' && l.sub) { walk(l.sub, g); continue }
      if (g > 1) out.set(l, g)
    }
  }
  walk(ps.plan.layers, 1)
  return out
}

// ---------------------------------------------------------------------------
// Preparation (L1)
// ---------------------------------------------------------------------------

const clampRect = (r: Rect, W: number, H: number): Rect => {
  const x = Math.max(-4 * W, Math.min(5 * W, r.x)), y = Math.max(-4 * H, Math.min(5 * H, r.y))
  return { x, y, w: Math.max(0, Math.min(5 * W - x, r.w)), h: Math.max(0, Math.min(5 * H - y, r.h)) }
}
const union = (rs: readonly Rect[]): Rect => {
  const x0 = Math.min(...rs.map(r => r.x)), y0 = Math.min(...rs.map(r => r.y))
  return { x: x0, y: y0, w: Math.max(...rs.map(r => r.x + r.w)) - x0, h: Math.max(...rs.map(r => r.y + r.h)) - y0 }
}

/** Each main-sequence click step's trigger: its first effect's node type (click, with previous, after previous). */
function stepTriggers(tree: TimingTree | undefined, steps: number): PlanStepRecord['trigger'][] {
  const all = Array.from({ length: steps }, (): PlanStepRecord['trigger'] => 'click')
  if (!tree) return all
  let main: TimeNode | undefined
  const find = (n: TimeNode, depth: number) => {
    if (main || depth > 64 || !('children' in n)) return
    if (n.ctn.nodeType === 'mainSeq') { main = n; return }
    for (const c of n.children) find(c, depth + 1)
  }
  find(tree.root, 0)
  if (!main || !('children' in main)) return all
  const groups = main.children.filter(c => c.ctn.start.length > 0 && c.ctn.start.every(s => s.delayMs === 'indefinite'))
  if (groups.length !== steps) return all
  const first = (n: TimeNode, depth: number): string | undefined => {
    if (depth > 64) return undefined
    const t = n.ctn.nodeType
    if (t === 'clickEffect' || t === 'withEffect' || t === 'afterEffect') return t
    if ('children' in n) for (const c of n.children) { const f = first(c, depth + 1); if (f) return f }
    return undefined
  }
  return groups.map(g => { const t = first(g, 0); return t === 'withEffect' ? 'with' : t === 'afterEffect' ? 'after' : 'click' })
}

/**
 * Additive (Task 7): slide `unit` prepared as the client will play it (the streamed metadata, text bases laid out)
 * and its timeline under `motion`, with no raster. The bundle build (`prepareL1`) and the Morph pair build
 * (`morph.ts`) share it, so both classify the same show.
 */
export function prepareShowSlide(show: ShowDoc, unit: number, motion: 'full' | 'reduced'): { ps: PreparedSlide; timeline: SlideTimeline; cmdDropped: boolean; sink: DiagnosticSink } {
  const loaded = show.metas.get(unit)
  if (!loaded) throw new Error(`show: slide ${unit} metadata not loaded`)
  const { meta, cmdDropped } = streamedMeta(loaded.slide)
  const slide = show.doc.slides[unit]!
  const sink = createDiagnosticSink()
  const ps = prepareSlide({ index: unit, slide, meta, watermark: false, charge: loaded.charge, sink })
  layoutTextBases(ps, measureCtx(), resolveFont)
  const tt = ps.textTiming
  const settings = show.index.settings
  const timeline = createSlideTimeline(tt?.tree ?? meta.timing, { showAnimation: settings.showAnimation, reducedMotion: motion === 'reduced', ...(tt ? { orderOf: tt.orderOf } : {}) })
  return { ps, timeline, cmdDropped, sink }
}

function prepareL1(show: ShowDoc, job: ShowPlanJob, ws: SlideShowWorkspace): L1Entry {
  const unit = job.unitIndex
  const { ps, timeline, cmdDropped, sink } = prepareShowSlide(show, unit, job.motion)
  const meta = ps.meta, slide = ps.slide
  const measure = measureCtx()
  const tt = ps.textTiming
  const settings = show.index.settings
  const classes = new Map<boolean, Classification>()
  if (timeline.steps > SHOW_PLAN_LIMITS.steps) return { refusal: { refused: 'show-limits', diagnostics: [{ code: 'show-limits' }] }, classes, refusals: new Map() }
  // D22: the text items the client fans out itself; over 128 key strings the slide plays its still (Task 4 ruling).
  const keys: Array<[string, string[]]> = tt ? [...tt.items].map(([k, items]) => [k, [...items]]) : []
  const groups: Array<[string, { wd?: string[][]; lt?: string[][] }]> = tt?.groups ? [...tt.groups].map(([k, g]) => [k, { ...(g.wd ? { wd: g.wd.map(x => [...x]) } : {}), ...(g.lt ? { lt: g.lt.map(x => [...x]) } : {}) }]) : []
  const iterateItems = keys.reduce((n, [, xs]) => n + 1 + xs.length, 0) + groups.reduce((n, [, g]) => n + 1 + [...(g.wd ?? []), ...(g.lt ?? [])].reduce((m, l) => m + l.length, 0), 0)
  if (iterateItems > SHOW_PLAN_ITERATE_ITEMS) return { refusal: { refused: 'show-limits', diagnostics: [{ code: 'show-limits' }] }, classes, refusals: new Map() }

  const { w: W, h: H } = slideSize(slide)
  const scale = job.scale
  // Rasters: every layer at the bucket; a growth target above it (D12), within maxUnitPx and the 64 MiB asset cap.
  const growth = growthOf(ps, growthTargets(tt?.tree ?? (meta.timing && !('refused' in meta.timing) ? meta.timing : undefined)))
  const rasters = new Map<PlanLayer, LayerBitmap>()
  // PR #78 round 13: every raster leases its w·h·4 from the job's ledger before its canvas exists; one that does
  // not fit releases what was leased and refuses the slide to its still (never a layer left out).
  const ctx = (s: number) => ({ scale: s, images: show.images ?? [], resolveFont, canvas: factory, workspace: ws, maxTextureSize: MAX_UNIT_PX, sink })
  const leases: Lease[] = []
  for (const layer of ps.layers) {
    if (layer.kind === 'watermark') continue
    let r = rasterPreparedLayer(ps, layer, ctx(scale))
    if ('refused' in r) {
      if (r.refused === 'empty') continue
      for (const l of leases) l.release()
      return { refusal: { refused: 'show-budget', diagnostics: [{ code: 'show-budget' }] }, classes, refusals: new Map() }
    }
    const g = growth.get(layer)
    if (g && g > 1) {
      const px = Math.max(r.canvas.width, r.canvas.height)
      const fit = Math.min(g, MAX_UNIT_PX / px, Math.sqrt(SHOW_PLAN_LIMITS.assetDecoded / (r.canvas.width * r.canvas.height * 4)))
      if (fit > 1 + 1e-9) {
        // Leased beside the bucket raster (both are held at the peak); one that does not fit keeps the bucket raster.
        const grown = rasterPreparedLayer(ps, layer, ctx(scale * fit))
        if (!('refused' in grown)) { r.lease.release(); r = grown }
      }
    }
    leases.push(r.lease)
    rasters.set(layer, { canvas: r.canvas, bounds: r.bounds, scale: r.scale, bytes: r.bytes })
  }

  // Hit regions: triggers (D14) and links (§5.1), as painted rectangles, clamped to the record's bounds.
  const triggerSpids = [...timeline.triggers()]
  const linkSpids = [...new Set(meta.links.map(l => l.spid))]
  const regions = hitRegionsOf(ps, [...new Set([...triggerSpids, ...linkSpids])], measure, resolveFont)
  const rectsOf = (spid: string) => (regions.get(spid) ?? []).map(r => clampRect(r, W, H)).filter(r => r.w > 0 && r.h > 0)
  const triggers = triggerSpids.slice(0, SHOW_PLAN_LIMITS.triggers).map(spid => {
    const rects = rectsOf(spid)
    return { spid, rects: rects.length > SHOW_PLAN_LIMITS.rectsPerStep ? [union(rects)] : rects }
  })
  // A link on a text run covers only that run's laid-out line boxes (PR #78 round 6); one that cannot be located
  // gets no region, so it never takes the clicks of other runs or of unlinked text.
  const runLinks = meta.links.flatMap(l => (l.run ? [{ spid: l.spid, run: l.run }] : []))
  const runRegions = runLinks.length ? runRegionsOf(ps, runLinks, measure, resolveFont) : new Map<string, Rect[]>()
  const linkRectsOf = (link: (typeof meta.links)[number]) => (link.run ? runRegions.get(`${link.spid}|${link.run.p}|${link.run.r}`) ?? [] : rectsOf(link.spid)).map(r => clampRect(r, W, H)).filter(r => r.w > 0 && r.h > 0)
  // The hit test takes the last region holding a point: a shape's own link goes before its runs' links, so a run wins
  // over its shape (as locally). Shapes keep the order of their first link.
  const bySpid = new Map<string, Array<(typeof meta.links)[number]>>()
  for (const link of meta.links) { const l = bySpid.get(link.spid) ?? []; l.push(link); bySpid.set(link.spid, l) }
  const ordered = [...bySpid.values()].flatMap(ls => [...ls.filter(l => !l.run), ...ls.filter(l => l.run)])
  let withheld = false
  const linkRegions: ShowPlanRecord['linkRegions'] = []
  for (const link of ordered) {
    let action: ShowPlanRecord['linkRegions'][number]['action']
    const a = link.action
    if (a.kind === 'slide') {
      const u = a.index
      if (!Number.isInteger(u) || u < 0 || u >= show.units.length || show.hidden.includes(u)) action = { kind: 'withheld' }
      else action = { kind: 'slide', index: u - show.hidden.filter(h => h < u).length }
    } else if (a.kind === 'jump') action = { kind: 'jump', verb: a.verb }
    else action = { kind: 'withheld' }
    if (action.kind === 'withheld') withheld = true
    // At the cap a shape's link keeps the rects that fit (as before); a run link is dropped whole, never cut to part
    // of its text.
    const rects = linkRectsOf(link)
    if (link.run && linkRegions.length + rects.length > SHOW_PLAN_LIMITS.links) continue
    for (const rect of rects) if (linkRegions.length < SHOW_PLAN_LIMITS.links) linkRegions.push({ rect, action })
  }

  const base = [...ps.base].filter(([k]) => ps.animated.has(k)).map(([k, g]): [string, typeof g] => [k, { x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot }])
  const record: L1Entry['record'] = {
    v: 1,
    slide: { index: unit, w: W, h: H },
    options: { showAnimation: settings.showAnimation, useTimings: settings.useTimings },
    show: { transition: meta.transition, timing: meta.timing, builds: meta.builds },
    base,
    triggers,
    linkRegions,
    iterate: { keys, groups },
  }
  const extraDiagnostics = [...(cmdDropped ? [{ code: 'cmd-dropped' }] : []), ...(withheld ? [{ code: 'link-withheld' }] : [])]
  return { ps, timeline, record, extraDiagnostics, rasters, classes, refusals: new Map() }
}

const L1_OVERHEAD = 64 * 1024
const rastersBytes = (e: L1Entry) => [...(e.rasters?.values() ?? [])].reduce((n, r) => n + r.canvas.width * r.canvas.height * 4, 0) + L1_OVERHEAD

// ---------------------------------------------------------------------------
// Frames
// ---------------------------------------------------------------------------

const rgbaOf = (canvas: LayerBitmap['canvas']) => (canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, canvas.width, canvas.height).data

/** `rgba` (W × H device px) with the page's visible mark painted once over it, at `scale` device px per slide px. */
export function markFrame(rgba: Uint8ClampedArray, W: number, H: number, scale: number, page: { w: number; h: number }, mark: MarkSpec): Uint8ClampedArray {
  const opts = markOptions(mark)
  if (!opts) return rgba
  const c = showCanvasBackend().createCanvas(W, H)
  const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  const img = ctx.createImageData(W, H)
  img.data.set(rgba)
  ctx.putImageData(img, 0, 0)
  ctx.setTransform(scale, 0, 0, scale, 0, 0)
  paintMark(ctx, { widthPx: page.w, heightPx: page.h }, opts)
  return ctx.getImageData(0, 0, W, H).data
}

/** A canvas-backed bitmap from straight RGBA, at `bounds` (slide px). */
function bitmapOf(w: number, h: number, rgba: Uint8ClampedArray, bounds: Rect, scale: number): LayerBitmap {
  const canvas = showCanvasBackend().createCanvas(w, h) as unknown as HTMLCanvasElement
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  const img = ctx.createImageData(w, h)
  img.data.set(rgba)
  ctx.putImageData(img, 0, 0)
  return { canvas, bounds: { ...bounds }, scale, bytes: w * h * 4 }
}

/** Attributes that only show, hide, fade or recolour a target: they never change its settled transform. */
const NON_GEOMETRIC = new Set(['style.visibility', 'style.opacity', 'fillcolor', 'stroke.color', 'style.color', 'fill.on', 'stroke.on', 'fill.type', 'fill.opacity', 'stroke.opacity'])

/**
 * Fix round 3 (controller ruling): whether any interactive (trigger) sequence of `tree` holds a behaviour that can
 * change a layer's settled transform: a motion path, scale or rotation, or a set/anim of any attribute other than
 * visibility, opacity and colour. Trigger histories (any order, overlap, repetition) cannot be enumerated, so a
 * trigger slide is proven only when its triggers never move a layer: its rests are then exactly the main-boundary
 * rests, which D8 and the mark validation prove.
 */
export function triggersMoveGeometry(tree: TimingTree | undefined): boolean {
  if (!tree) return false
  const geometric = (n: TimeNode, depth: number): boolean => {
    if (depth > 256) return true
    if ('children' in n) return n.children.some(c => geometric(c, depth + 1))
    if (n.kind === 'animMotion' || n.kind === 'animRot' || n.kind === 'animScale') return true
    if (n.kind === 'anim' || n.kind === 'set') return n.attrNames.some(a => !NON_GEOMETRIC.has(a))
    return false
  }
  const interactive = (n: TimeNode): boolean => n.ctn.nodeType === 'interactiveSeq' || n.ctn.start.some(c => c.evt !== undefined && c.target !== undefined)
  const walk = (n: TimeNode, depth: number): boolean => {
    if (depth > 256) return true
    if (interactive(n)) return geometric(n, depth)
    return 'children' in n && n.children.some(c => walk(c, depth + 1))
  }
  return walk(tree.root, 0)
}

/** The draws of each step boundary `0..steps` (settled checkpoints). */
function boundaryDraws(ps: PreparedSlide, timeline: SlideTimeline): DrawnLayer[][] {
  const rng = createRng(0x5eed + ps.index), sink = createDiagnosticSink(), measure = measureCtx()
  return Array.from({ length: timeline.steps + 1 }, (_, k) => drawsAt(ps, slideStateAt(ps, timeline, timeline.checkpoint(k), 0, rng, sink), { measure, resolveFont }))
}

/**
 * PR #78 round 7: the slide as it arrives, `timeline.initial()` at time 0, the state the local stage draws its
 * incoming transition frame from (`stateOf(p, initial, 0)`). With autoplay it differs from `checkpoint(0)`, which
 * settles the pre-click effects (an auto fade is at progress 1 there, 0 here).
 */
function arrivalDraws(ps: PreparedSlide, timeline: SlideTimeline): DrawnLayer[] {
  const rng = createRng(0x5eed + ps.index), sink = createDiagnosticSink()
  return drawsAt(ps, slideStateAt(ps, timeline, timeline.initial(), 0, rng, sink), { measure: measureCtx(), resolveFont })
}

// ---------------------------------------------------------------------------
// The build
// ---------------------------------------------------------------------------

/** The cached L1 entry for `job` (computed on a miss). */
const l1KeyOf = (show: ShowDoc, job: ShowPlanJob) => showPlanL1Key({ docRef: show.docRef, contentPolicyId: show.contentPolicyId, layoutFingerprint: show.fingerprint, scaleBucket: job.scale, motion: job.motion, unitId: job.unitIndex })

/**
 * The cached L1 entry for `job` (computed on a miss), its rasters charged to `ws`: leased one by one as they are
 * drawn on a miss, or all at once on a hit (the cached rasters count toward this job's grant). A hit whose
 * rasters do not fit `ws` is refused (`show-budget`) without touching the cache.
 */
function l1For(show: ShowDoc, job: ShowPlanJob, ws: SlideShowWorkspace): L1Entry {
  const key = l1KeyOf(show, job)
  const hit = show.l1.get(key)
  if (hit) {
    show.stats.l1.hits++
    if (hit.rasters && !ws.lease(rastersBytes(hit) - L1_OVERHEAD, 'l1:rasters')) return { refusal: { refused: 'show-budget', diagnostics: [{ code: 'show-budget' }] }, classes: new Map(), refusals: new Map() }
    return hit
  }
  show.stats.l1.misses++
  const entry = prepareL1(show, job, ws)
  show.l1.put(key, entry, rastersBytes(entry))
  return entry
}

/** The preparation a keyframe track composes from (the slide's L1 entry: its timeline and original rasters). */
export function trackInputs(show: ShowDoc, job: ShowPlanJob, ws: SlideShowWorkspace = jobLedger(show, job.unitIndex, job.scale, {})): { ps: PreparedSlide; timeline: SlideTimeline; rasters: Map<PlanLayer, LayerBitmap> } {
  const e = l1For(show, job, ws)
  if (e.refusal || !e.ps || !e.timeline || !e.rasters) throw new Error('show: the slide has no plan')
  return { ps: e.ps, timeline: e.timeline, rasters: e.rasters }
}

function classificationOf(e: L1Entry, job: ShowPlanJob): Classification {
  const have = e.classes.get(job.strictMarks)
  if (have) return have
  const c = classifySlide({ ps: e.ps!, timeline: e.timeline!, scale: job.scale, rasters: e.rasters!, strictMarks: job.strictMarks, triggers: e.timeline!.triggers(), resolveFont })
  e.classes.set(job.strictMarks, c)
  return c
}

const markSettingsOf = (m: MarkSpec) => ({ visible: m.kind, text: m.kind === 'viewer' ? m.text : m.kind === 'fixed' ? 'fixed' : '', forensic: false as const })
/** The §7 validation code's message key (shared with the Morph pair build, `morph.ts`). */
export const VALIDATION_MESSAGE_KEY = createHash('sha256').update('officeview show-plan validation message@1').digest()

export function buildShowPlan(doc: ShowHost, job: ShowPlanJob, opts: ShowBuildOptions = {}): BuiltShowPlan | ShowRefusal {
  const show = doc.show
  if (!show) throw new Error('show: not a deck')
  const now = opts.now ?? (() => performance.now())
  const start = now(), deadline = opts.deadlineMs ?? SHOW_PLAN_DEADLINE_MS
  const check = () => { if (now() - start > deadline) throw new Timeout() }
  const seamed = opts.budget !== undefined || !!opts.compose
  const ws = jobLedger(show, job.unitIndex, job.scale, opts)
  let e: L1Entry | undefined
  try {
    const l2key = showPlanL2Key(l1KeyOf(show, job), markSettingsOf(job.mark), job.strictMarks, job.kBasis ? Buffer.from(job.kBasis).toString('hex') : '')
    if (!seamed) {
      const hit = show.l2.get(l2key)
      if (hit) { show.stats.l2.hits++; return hit }
    }
    e = l1For(show, job, ws)
    if (e.refusal) return e.refusal
    const refused = seamed ? undefined : e.refusals.get(job.strictMarks)
    if (refused) return refused
    show.stats.l2.misses++
    check()
    const out = assemble(show, e, job, opts, check, ws)
    if (seamed) return out
    if ('refused' in out) e.refusals.set(job.strictMarks, out)
    else show.l2.put(l2key, out, out.assets.reduce((n, a) => n + a.png.length, 64 * 1024))
    return out
  } catch (err) {
    if (!(err instanceof Timeout)) throw err
    const out: ShowRefusal = { refused: 'show-plan-timeout', diagnostics: [{ code: 'show-plan-timeout' }] }
    if (e && !seamed) e.refusals.set(job.strictMarks, out)
    return out
  }
}

function assemble(show: ShowDoc, e: L1Entry, job: ShowPlanJob, opts: ShowBuildOptions, check: () => void, ws: SlideShowWorkspace): BuiltShowPlan | ShowRefusal {
  const ps = e.ps!, timeline = e.timeline!, rasters = e.rasters!, record = e.record!
  // Fix round 3 (controller ruling): a trigger that can move a layer leaves its rests unprovable (trigger histories
  // cannot be enumerated): the slide plays its still, before any classification work.
  if (timeline.triggers().length > 0 && triggersMoveGeometry(ps.meta.timing && !('refused' in ps.meta.timing) ? ps.meta.timing : undefined)) {
    return { refused: 'still', diagnostics: [{ code: 'step-composited' }, { code: 'trigger-cut' }, { code: 'mark-rest' }] }
  }
  const c = classificationOf(e, job)
  check()
  if (c.fallback === 'still') {
    const codes = new Set(c.diagnostics.map(d => d.code))
    return { refused: codes.has('show-limits') ? 'show-limits' : codes.has('show-budget') ? 'show-budget' : 'still', diagnostics: c.diagnostics.filter(d => D21.has(d.code)) }
  }
  const scale = job.scale
  const { w: W, h: H } = record.slide
  const composer = createSlideComposer(ps, rasters, scale)
  const PW = composer.width, PH = composer.height
  const page = { w: W, h: H }
  const budget = opts.budget ?? SHOW_BUNDLE_BUDGET
  try {
    const n = timeline.steps
    const boundaries = boundaryDraws(ps, timeline)
    const init = timeline.initial(), init0 = timeline.stateAt(init, 0)
    const autoplay = timeline.busyUntil(init) > 0 || timeline.keyTimes(init, 0, Number.POSITIVE_INFINITY).length > 0 || init0.active.length > 0 || init0.frozen.length > 0
    // The entry frame: the slide as it arrives (what the incoming transition shows; round 7: with autoplay, the
    // initial state at time 0, not the settled checkpoint 0, which stays the boundary the rests are validated at),
    // composed from the original rasters, marked once (as a unit).
    const entryRgba = markFrame(composer.compose(autoplay ? arrivalDraws(ps, timeline) : boundaries[0]!), PW, PH, scale, page, job.mark)
    check()
    // Fix round 1 (controller ruling): a trigger needs the layers that draw its end state whether it plays from
    // layers or is cut (§6a: a cut jumps to its end state), and so does autoplay. Task 5 lets a trigger slide play
    // from layers only with zero drops and masks, so its layers are complete.
    const interactive = c.triggers.size > 0 || autoplay
    let steps: Array<{ eligibility: 'layers' | 'composited'; reason?: string; frames?: number }> = c.steps
    let composited = !(steps.some(s => s.eligibility === 'layers') || interactive)
    const diagnostics: Array<{ code: string; step?: number }> = [...c.diagnostics]

    /** Every step composited (validation failure, a non-translation rest): D3 frames over all steps. */
    const allComposited = (why: 'mark-rest' | 'marks') => {
      const durs = stepDurations(timeline)
      const frames = allocateFrames(durs.map(d => ({ composited: true, durMs: d })))
      steps = steps.map((s, k) => ({ eligibility: 'composited' as const, reason: s.eligibility === 'composited' ? s.reason : why, frames: frames[k]!.frames }))
      for (let k = 0; k < n; k++) {
        if (!diagnostics.some(d => d.code === 'step-composited' && d.step === k)) diagnostics.push({ code: 'step-composited', step: k })
        if (frames[k]!.truncated && !diagnostics.some(d => d.code === 'step-truncated' && d.step === k)) diagnostics.push({ code: 'step-truncated', step: k })
      }
      composited = true
    }

    /**
     * Fix round 1: after the rest rule or mark validation withdraws the layers, a slide with a trigger or autoplay
     * cannot draw them from keyframe tracks (which carry main steps only): it plays its still, never a plan whose
     * triggered or autoplayed content was removed. Zero main steps make no difference.
     */
    const stillIfInteractive = (): ShowRefusal | undefined => {
      if (!interactive) return undefined
      const diags = diagnostics.filter(d => D21.has(d.code))
      if (!diags.some(d => d.code === 'step-composited')) diags.push({ code: 'step-composited' })
      return { refused: 'still', diagnostics: diags.slice(0, MAX_DIAGNOSTICS) }
    }

    // The layers that ship: kept by the classifier, masked (`layer × mask`), at their rest (D8).
    interface Shipped { layer: PlanLayer; index: number; bmp: LayerBitmap; rest: Rect }
    let shipped: Shipped[] = []
    let classified: Map<PlanLayer, LayerBitmap> = new Map()
    if (!composited) {
      // PR #78 round 13: the masked copies lease first (the validation reference draws from these same copies).
      const maskBytes = [...c.masks.keys()].reduce((n, l) => { const b = rasters.get(l); return n + (b && !c.drop.has(l) ? b.canvas.width * b.canvas.height * 4 : 0) }, 0)
      if (maskBytes > 0 && !ws.lease(maskBytes, 'masks')) return { refused: 'show-budget', diagnostics: [{ code: 'show-budget' }] }
      classified = classifiedRasters(rasters, c.drop, c.masks)
      ps.layers.forEach((layer, index) => {
        const bmp = classified.get(layer)
        if (!bmp) return
        let rest: Rect = { ...bmp.bounds }
        for (const list of boundaries) {
          const d = list.find(x => x.layer === layer)
          if (!d || !d.state.visible || !(d.state.opacity > 0)) continue
          const m = stateMatrixOf(d.state)
          // An axis-aligned scale and translation (a grown layer held at its growth) maps the box to a box; a rotation or skew does not.
          if (Math.abs(m[1]) > 1e-6 || Math.abs(m[2]) > 1e-6 || !(m[0] > 1e-6) || !(m[3] > 1e-6)) rest = { x: NaN, y: NaN, w: 0, h: 0 }
          else rest = { x: m[0] * bmp.bounds.x + m[4], y: m[3] * bmp.bounds.y + m[5], w: m[0] * bmp.bounds.w, h: m[3] * bmp.bounds.h }
          break
        }
        shipped.push({ layer, index, bmp, rest })
      })
      // A rest that is not a pure translation cannot carry a page-anchored mark: fail closed.
      if (shipped.some(s => !Number.isFinite(s.rest.x))) {
        allComposited('mark-rest'); diagnostics.push({ code: 'mark-rest' }); shipped = []
        const still = stillIfInteractive()
        if (still) return still
      }
    }

    const layerRecords = (list: readonly Shipped[]): Array<PlanLayerRecord | PlanContainerRecord> => {
      const { parentOf, groups } = groupTree(ps)
      const measure = measureCtx()
      const out: Array<PlanLayerRecord | PlanContainerRecord> = []
      const emitted = new Set<number>()
      /** A group's container record (and its own enclosing groups'), before the first child naming it. */
      const container = (id: number, z: number) => {
        if (emitted.has(id)) return
        const g = groups.get(id)!
        if (g.parent !== undefined) container(g.parent, z)
        emitted.add(id)
        out.push({ id, key: g.key, z: z - 0.5, container: true, ...(g.parent !== undefined ? { parent: g.parent } : {}) })
      }
      for (const s of list) {
        const p = parentOf.get(s.layer)
        if (p !== undefined) container(p, s.index)
        out.push(raster(s))
      }
      return out
      function raster(s: Shipped): PlanLayerRecord {
        const l = s.layer
        const r: PlanLayerRecord = { id: s.index, bounds: { ...s.bmp.bounds }, z: s.index, assetId: 1 + s.index, decodedBytes: s.bmp.canvas.width * s.bmp.canvas.height * 4, scale: s.bmp.scale }
        if (l.kind === 'target') {
          r.key = l.key
          if (l.part) r.part = l.part
          if (l.part === 'range') {
            const b = textRangeBox(l.nodes, ps.slide, resolveFont, measure)
            if (b) r.maskBox = clampRect({ x: b.minX, y: b.minY, w: b.maxX - b.minX, h: b.maxY - b.minY }, W, H)
          }
        }
        const parent = parentOf.get(l)
        if (parent !== undefined) r.parent = parent
        if (l.kind === 'static' && l.background) r.background = true
        return r
      }
    }
    const planOf = (list: readonly Shipped[]): ShowPlanRecord => {
      const triggerOf = stepTriggers(ps.meta.timing && !('refused' in ps.meta.timing) ? ps.meta.timing : undefined, n)
      const stepRecords: PlanStepRecord[] = steps.map((s, k) => ({
        trigger: triggerOf[k]!, eligibility: s.eligibility,
        ...(s.eligibility === 'composited' && s.reason ? { reason: s.reason } : {}),
        ...(s.eligibility === 'composited' ? { track: { fps: 15 as const, count: Math.max(1, s.frames ?? 1) } } : {}),
      }))
      const diags = [...diagnostics, ...(e.extraDiagnostics ?? [])].filter(d => D21.has(d.code)).slice(0, MAX_DIAGNOSTICS)
      // Each trigger the classifier cut (§6a) says so, so the client animates only the others.
      const triggers = record.triggers.map(t => (c.triggers.get(t.spid) === 'cut' ? { ...t, cut: true as const } : t))
      return { ...record, triggers, layers: layerRecords(list), steps: stepRecords, entry: ENTRY_ASSET_ID, diagnostics: diags }
    }
    const planBytes = (plan: ShowPlanRecord): number | undefined => {
      try { return encodeShowMessage({ tag: 1, plan }).length - 1 } catch (err) { if (err instanceof PlanRejected) return undefined; throw err }
    }
    const limits = (): ShowRefusal => ({ refused: 'show-limits', diagnostics: [{ code: 'show-limits' }] })
    const overBudget = (): ShowRefusal => ({ refused: 'show-budget', diagnostics: [{ code: 'show-budget' }] })
    const sizeOf = (list: readonly Shipped[], plan: ShowPlanRecord) => {
      const bytes = planBytes(plan)
      if (bytes === undefined) return undefined
      return { bytes: list.reduce((a, s) => a + s.bmp.canvas.width * s.bmp.canvas.height * 4, 0) + PW * PH * 4 + bytes, maxAsset: Math.max(PW * PH * 4, ...list.map(s => s.bmp.canvas.width * s.bmp.canvas.height * 4)) }
    }

    // The budget, from w·h·4, before any layer PNG is encoded.
    let plan = planOf(shipped)
    let size = sizeOf(shipped, plan)
    if (!size) return limits()
    if (size.bytes > budget || size.maxAsset > SHOW_PLAN_LIMITS.assetDecoded) return overBudget()
    check()

    // Visible mark at rest, then encode: what the client receives.
    let layerAssets: ShowAsset[] = shipped.map(s => {
      const marked = burnVisibleMark(s.bmp, s.rest, job.mark, page)
      const w = marked.canvas.width, h = marked.canvas.height
      return { tag: 2, assetId: 1 + s.index, w, h, png: encodePng(w, h, rgbaOf(marked.canvas)), rest: { x: Math.round(s.rest.x * scale), y: Math.round(s.rest.y * scale), scale: w / (s.rest.w * scale) } }
    })
    check()

    // §7 validation at every boundary, on the decoded shipped PNGs with the validation code's Layer-1 applied.
    if (shipped.length) {
      const ok = validate(ps, rasters, classified, c, shipped.map((s, i) => ({ layer: s.layer, bounds: s.bmp.bounds, scale: s.bmp.scale, asset: layerAssets[i]! })), boundaries, { scale, PW, PH, page, job, docKey: show.docKey, compose: opts.compose, check })
      if (!ok) {
        allComposited('marks')
        const still = stillIfInteractive()
        if (still) return still
        shipped = []
        layerAssets = []
        plan = planOf(shipped)
        size = sizeOf(shipped, plan)
        if (!size) return limits()
        if (size.bytes > budget) return overBudget()
      }
    }
    const entry: ShowAsset = { tag: 3, assetId: ENTRY_ASSET_ID, w: PW, h: PH, png: encodePng(PW, PH, entryRgba), rest: { x: 0, y: 0, scale: 1 } }
    return { plan, assets: [entry, ...layerAssets], decodedBytes: size.bytes }
  } finally {
    composer.dispose()
  }
}

/**
 * The group targets with sub-layers (container records, numbered after the raster layers): each one's key and
 * enclosing group, and each raster sub-layer's immediate group.
 */
function groupTree(ps: PreparedSlide): { parentOf: Map<PlanLayer, number>; groups: Map<number, { key: string; parent?: number }> } {
  const parentOf = new Map<PlanLayer, number>()
  const groups = new Map<number, { key: string; parent?: number }>()
  let next = ps.layers.length
  const walk = (layers: readonly PlanLayer[], parent: number | undefined, depth: number) => {
    if (depth > 64) return
    for (const l of layers) {
      if (l.kind === 'target' && l.sub) {
        const id = next++
        groups.set(id, { key: l.key, ...(parent !== undefined ? { parent } : {}) })
        walk(l.sub, id, depth + 1)
        continue
      }
      if (parent !== undefined) parentOf.set(l, parent)
    }
  }
  walk(ps.plan.layers, undefined, 0)
  return { parentOf, groups }
}

function stateMatrixOf(s: DrawnLayer['state']): number[] {
  const t = s.transform.every(Number.isFinite) ? s.transform : [1, 0, 0, 1, 0, 0]
  const dx = s.motion && Number.isFinite(s.motion.dx) ? s.motion.dx : 0, dy = s.motion && Number.isFinite(s.motion.dy) ? s.motion.dy : 0
  return [t[0], t[1], t[2], t[3], t[4] + dx, t[5] + dy]
}

interface ValidateEnv {
  scale: number; PW: number; PH: number; page: { w: number; h: number }; job: ShowPlanJob; docKey: string
  compose?: ShowBuildOptions['compose']; check: () => void
}

/**
 * §7: the composite of the shipped layers (decoded from their PNGs, Layer-1 applied with the validation code at
 * each rest crop) at every boundary, against the classified slide composed once, marked visibly once and with
 * Layer-1 once; and the validation code attributed on the composite (D7).
 */
function validate(ps: PreparedSlide, rasters: ReadonlyMap<PlanLayer, LayerBitmap>, classified: ReadonlyMap<PlanLayer, LayerBitmap>, c: Classification, shipped: ReadonlyArray<{ layer: PlanLayer; bounds: Rect; scale: number; asset: ShowAsset }>, boundaries: DrawnLayer[][], env: ValidateEnv): boolean {
  const { scale, PW, PH, page, job } = env
  const kBasis = job.kBasis ?? createHash('sha256').update(`officeview show-plan validation basis@1\n${env.docKey}\n${job.unitIndex}\n${scale}`).digest()
  const codes = [1, 2, 3, 4].map(t => extendedCodeword(messageFor(VALIDATION_MESSAGE_KEY, 0, t, 0)))
  const code = codes[0]!
  let chips: ReturnType<typeof basis> | undefined
  try { chips = basis(kBasis, PW, PH, DEFAULT_PARAMS) } catch { chips = undefined }
  const sent = new Map<PlanLayer, LayerBitmap>(), visible = new Map<PlanLayer, LayerBitmap>()
  for (const s of shipped) {
    const v = decodePng(s.asset.png)
    visible.set(s.layer, bitmapOf(v.w, v.h, v.rgba, s.bounds, s.scale))
    const r = markLayerRaster({ png: s.asset.png, w: s.asset.w, h: s.asset.h, kBasis, code, crop: { x: s.asset.rest.x, y: s.asset.rest.y, pageW: PW, pageH: PH, scale: s.asset.rest.scale } })
    const d = decodePng(r.png)
    sent.set(s.layer, bitmapOf(d.w, d.h, d.rgba, s.bounds, s.scale))
  }
  env.check()
  const layered = createSlideComposer(ps, sent, scale)
  // The visible mark's own contribution (controller ruling): the shipped layers with their visible mark only.
  const visibleOnly = createSlideComposer(ps, visible, scale)
  const reference = createSlideComposer(ps, rasters, scale)
  for (const l of c.drop) reference.classify(l, null)
  // The shipped masked copies (`classifiedRasters`, the same `maskedRaster`), not a second copy of each (PR #78 round 13).
  for (const [l, m] of c.masks) { const b = rasters.get(l); if (b) reference.classify(l, classified.get(l) ?? maskedRaster(b, m)) }
  const tau = tauFor(codes.length)
  const visibleRefs = new Map<number, Uint8ClampedArray>()
  const image = (data: Uint8ClampedArray) => ({ width: PW, height: PH, data, colorSpace: 'srgb' }) as unknown as ImageData
  try {
    const visibleRef = (k: number) => {
      let v = visibleRefs.get(k)
      if (!v) { v = markFrame(new Uint8ClampedArray(reference.compose(boundaries[k]!, undefined, true)), PW, PH, scale, page, job.mark); visibleRefs.set(k, v) }
      return v
    }
    const verdict = validateMarks({ plan: { steps: boundaries.slice(1) } },
      (k) => {
        env.check()
        const v = visibleRef(k)
        const r = markLayerRaster({ png: encodePng(PW, PH, v), w: PW, h: PH, kBasis, code, crop: { x: 0, y: 0, pageW: PW, pageH: PH } })
        return image(decodePng(r.png).rgba)
      },
      (k) => {
        const composite = image(new Uint8ClampedArray(layered.compose(boundaries[k]!)))
        return env.compose ? env.compose(k, composite) : composite
      },
      {
        ...(chips ? {
          attributes: (k: number, composite: ImageData) => {
            const bmp = (d: Uint8ClampedArray) => ({ width: PW, height: PH, data: d })
            const decision = detect(bmp(new Uint8ClampedArray(composite.data)), bmp(visibleRef(k)), chips!, codes, DEFAULT_PARAMS, tau)
            return decision.kind === 'attributed' && decision.index === 0
          },
        } : {}),
        contribution: {
          composeMarked: (k) => image(new Uint8ClampedArray(visibleOnly.compose(boundaries[k]!))),
          // The shipped layers unmarked are the classified rasters: the reference's own unmarked composite.
          composeUnmarked: (k) => image(new Uint8ClampedArray(reference.compose(boundaries[k]!, undefined, true))),
          referenceMarked: (k) => image(visibleRef(k)),
          referenceUnmarked: (k) => image(new Uint8ClampedArray(reference.compose(boundaries[k]!, undefined, true))),
        },
      })
    if (layered.takeDiagnostics().length || visibleOnly.takeDiagnostics().length) return false
    return verdict.ok
  } finally {
    layered.dispose()
    visibleOnly.dispose()
    reference.dispose()
  }
}

