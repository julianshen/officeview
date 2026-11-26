/**
 * The pool's per-worker render core (§9.2/§9.4): parse once per
 * (docRef, contentPolicyId, layoutFingerprint), paint a unit at a scale
 * bucket, then either encode a final raster with the visible mark composited
 * or record commands (relaxed sessions). Deterministic: no clock, no
 * randomness; the mark text is a parameter (§14).
 */
import { ObserverBudgetExceeded, occlusionObserver, type OcclusionEvidence } from './occlusion'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { installCanvasBackend, type CanvasBackend } from '../../src/core/canvas-backend'
import { OfficePackage } from '../../src/core/zip'
import type { WatermarkOptions } from '../../src/core/watermark'
import { FIXED_MARK_TEXT, markOptions, paintMark, paintMarkRegion, type MarkSpec } from './mark/visible'
import { detectAndParse } from '../../src/components/loadOfficeSource'
import { parseXlsx } from '../../src/xlsx/parse'
import { prepareSheetAssets } from '../../src/components/xlsx/sheetAssets'
import { clockEpochMs, type SheetClock } from './sheet/clock'
import { buildStreamLayout, LayoutUnavailable, type ServerSheetLayout } from './sheet/layout'
import { LayoutTooLarge } from '../../src/stream/sheet-wire'
import { buildSheetSnapshot, servableSheet, type ServableSheet, type SheetSnapshot } from './sheet/snapshot'
import { TILE_PALETTE_COLORS, tileJobProblem, tileLimit, type RenderedTile, type TileJob } from './sheet/tile'
import { sessionSizesOf, type WireSizes } from './sheet/view-state'
import { sheetTextView, type TextView } from './sheet/text-view'
import { answerSheetText, type SheetTextJob, type SheetTextReply } from './sheet/text-ops'
import { keyDocument, sheetKeyView, type KeyView } from './sheet/key-view'
import { openSheetViewPort, type DerivedCache, type DerivedView, type SheetViewJob, type SheetViewPort, type SheetViewReply } from './sheet/view-ops'
import { canonicalView, canonicalViewHash, IDENTITY_VIEW, isIdentityView, type CanonicalView } from './sheet/view-spec'
import type { FormatBudget } from '../../src/xlsx/view/formatWorkspace'
import { TILE_PX, TILE_ZOOM_STEPS } from '../../src/stream/sheet-wire'
import { renderSheetViewport } from '../../src/xlsx/view/renderViewport'
import { tileLogicalBox } from '../../src/xlsx/view/tileCache'
import { LayoutCache } from '../../src/xlsx/view/textLayout'
import { buildSheetGeometry, frameBudgetFor, planViewport, type SheetGeometryResult } from '../../src/xlsx/view/viewport'
import type { LogicalViewport } from '../../src/xlsx/view/contracts'
import { getPaintables, staticSheetSpec, type PaintableArray } from '../../src/render/paint'
import type { XlsxDocument } from '../../src/xlsx/types'
import { createRecorder, type ImageSnapshot } from '../../src/stream/record/recorder'
import type { Asset } from '../../src/stream/ops'
import { UNIT_RASTER_DECODED_CAP } from '../../src/stream/raster'
import { serverBackend } from './backend'
import { facesUsed, loadPack, pinFonts } from './fonts'
import { admit, projectPackage, type PolicyDiagnostic } from './policy'
import { hiddenSlideUnits } from './policy/pptx'
import { pptxPackageHandle } from '../../src/pptx/parse'
import { drawNotes, NOTES_WIDTH_CSS, planNotes, readSlideNotes } from './notes'
import type { PptxDocument } from '../../src/pptx/types'
import { decodeImageAsset } from '../../src/core/images'
import { createDiagnosticSink } from '../../src/slideshow/diagnostics'
import { readShowIndex } from '../../src/slideshow/read/deck'
import { shapeResolver } from '../../src/slideshow/read/resolve'
import { createSlideMetadataStore } from '../../src/slideshow/read/store'
import { createUnderlayReader } from '../../src/slideshow/read/underlay'
import { readSlideTiming } from '../../src/slideshow/read/timing'
import type { Lease, SlideShowWorkspace } from '../../src/slideshow/types'
import { gateBackend } from './show/canvas'
import { buildShowPlan, createShowDoc, type BuiltShowPlan, type ShowBuildOptions, type ShowDoc, type ShowPlanJob, type ShowRefusal } from './show/build'
import { hasAudioUse, NO_MEDIA_CONTEXT, slideAudioUses, type SlideMediaContext } from './media/uses'
import { mediaIdentity } from './media/bind'
import { renderTrack } from './show/tracks'
import { buildMorphPlan, MorphInvalid, morphReference, type BuiltMorph, type MorphJob } from './show/morph'
import { asAudioUse, createDocMediaTable, localProbeRunner, MediaPartRefused, writePart, type DocMediaEntry, type DocMediaTable, type MediaTableSeams, type ProbeRunner } from './media/probe'
import { attachMediaTemp, type MediaTemp } from './media/temp'
import type { FfmpegInfo } from './media/ffmpeg'
import { hiddenShapeIds } from './policy/pptx'
import { readTree } from './policy/tree'
import { readRelationships } from '../../src/slideshow/read/deck'
import { mediaShapes } from '../../src/slideshow/read/media'
import { parseXml } from '../../src/core/xml'
import { answerText, buildServedUnit, OBSERVER_BUDGET, prepareWorkspace, recordSpans, ServedCache, TEXT_RECORD_BUDGET, TextBudgetExceeded, TextRefusal, type ServedUnit, type TextJob, type TextReply } from './text/served'

export const MAX_UNIT_PX = 8192
export const SCALE_BUCKETS = [1, 2] as const
/** Stage 4b: sized snapshot geometries kept per document (LRU). */
export const SIZED_GEOMETRIES = 8
/** Stage 4b: text views kept per document (LRU, by `(sheetIndex, clock, sha256(sizes))`), released on eviction and `drop`. */
export const TEXT_VIEWS = 16
/** Stage 4b: the formatting bytes (CF text scopes) one document's text views may hold together; past it a scope is
 * unavailable, and its sqrefs are blank in the view (P4, fail closed). */
export const TEXT_VIEW_FORMAT_BYTES = 64 * 1024 * 1024
/** Stage 4c (P3): `sheet-view` ports kept per document (LRU, by `(sheetIndex, clock)`), disposed on eviction and `drop`. */
export const SHEET_VIEW_PORTS = 4
/** Stage 4c (P3): derived views kept per document (LRU, by `(sheetIndex, clock, canonical hash)`), dropped on `drop`. */
export const DERIVED_VIEWS = 8
/** One cached `sheet-view` port: `uses` counts the jobs holding it; one evicted while held is disposed by its last user. */
interface SheetViewEntry { port: Promise<{ port: SheetViewPort; keyView: KeyView }>; uses: number; evicted: boolean }
/** The canonical hash of session sizes, shared by the sized-geometry and text-view keys. */
const sizesHash = (sizes: WireSizes) => createHash('sha256').update(JSON.stringify(sizes.map((e) => [e.axis, e.index, e.px]))).digest('hex')
/** The canonical hash of the identity view, the view part of every identity-view geometry and text-view key. */
const IDENTITY_HASH = canonicalViewHash(IDENTITY_VIEW)
/** One cached text view: `uses` counts the jobs holding it; an entry evicted while held is released by its last user. */
interface TextViewEntry { view: Promise<TextView>; uses: number; evicted: boolean }
export { FIXED_MARK_TEXT, markOptions, paintMark, type MarkSpec }
export type Projection = 'final-raster' | 'commands'

export const sizeCeil = (v: number) => Math.ceil(v - 1e-9)
/** A unit surface is allowed only within BOTH the per-side cap and the §14
 * UNIT_RASTER decoded cap (w × h × 4 ≤ 64 MiB) — checked before any surface
 * is allocated. Larger units need tiles (Stage 4) and are refused today. */
export const unitSurfaceAllowed = (w: number, h: number) => Math.max(w, h) <= MAX_UNIT_PX && w * h * 4 <= UNIT_RASTER_DECODED_CAP
export interface UnitSize { w: number; h: number }
/** `evidence`: the §7.7 occlusion evidence of the paint pass (server-side only). */
export type RenderedUnit = (
  | { kind: 'raster'; w: number; h: number; png: Uint8Array; fallback?: string }
  | { kind: 'commands'; w: number; h: number; ops: Uint8Array; opCount: number; assets: Asset[]; faces: string[] }
) & { evidence?: OcclusionEvidence }
/** `wrap` may interpose on the context the painter receives (the visible mark is painted outside it). */
export interface PaintObserver { beforeUnit(ctx: CanvasRenderingContext2D, unitIndex: number): void; afterUnit(ctx: CanvasRenderingContext2D, unitIndex: number): void; wrap?(ctx: CanvasRenderingContext2D): CanvasRenderingContext2D }
export class RenderRefusal extends Error { constructor(readonly code: 'unsupported-format' | 'unit-range' | 'scale' | 'oversize' | 'unknown-doc' | 'bad-job' | 'grant' | 'media-part' | 'layout-too-large', m?: string) { super(m ?? code); this.name = 'RenderRefusal' } }

/** §7.7: every paint pass that serves a unit records occlusion evidence. */
function observe(ctx: CanvasRenderingContext2D, w: number, h: number, scale: number, observers: readonly PaintObserver[], maxObserverBytes?: number) {
  const occ = occlusionObserver(w, h, maxObserverBytes, scale)
  let painted = occ.wrap(ctx)
  for (const o of observers) if (o.wrap) painted = o.wrap(painted)
  return { painted, evidence: () => occ.evidence(), runDetails: () => occ.runDetails(), exceeded: () => occ.exceeded() }
}

const root = resolve(import.meta.dir, '../..')
const versionOf = (p: string) => (JSON.parse(readFileSync(p, 'utf8')) as { version: string }).version

/** §9.3 layoutFingerprint inputs: officeview + backend versions, the font set, locale, policy. */
export function layoutFingerprint(contentPolicyId: string): string {
  const pack = loadPack()
  const h = createHash('sha256')
  h.update(`officeview@${versionOf(resolve(root, 'package.json'))}\n`)
  h.update(`@napi-rs/canvas@${versionOf(resolve(root, 'node_modules/@napi-rs/canvas/package.json'))}\n`)
  h.update(`pack@${pack.version}:${pack.faces.map((f) => `${f.id}=${f.hash96}`).join(',')}\n`)
  // The alias table changes layout too, and the show path resolves fonts pack-only (Task 14 ruling).
  h.update(`aliases=${Object.entries(pack.aliases).sort().map(([k, v]) => `${k}>${v}`).join(',')};show=pack-only/1\n`)
  h.update(`locale=en-US\npolicy=${contentPolicyId}\n`)
  return h.digest('hex')
}

/**
 * The `DocumentHost` entry for an open document. `notes`: a deck's ORIGINAL package (the projection drops notes
 * slides) and unit i's slide part. `show` (piece 4): a deck's show inputs over the PROJECTED package (never the
 * original, which holds withheld content), its metadata store and the show-plan caches. `anim` (D5): the shown
 * units with timing, a transition or both.
 */
export interface OpenDoc { key: string; docRef: string; show?: ShowDoc; /** Piece 5 (D21): the probe table, over the projected package. */ media: DocMediaTable; anim: number[]; morph: number[]; /** Piece 5 (D19): the shown units whose plan ships an audio record or a transition sound. */ mediaUnits: number[]; /** Absent until built: an XLSX entry opened first in tiles mode builds them on its first units use (R5). */ paintables?: PaintableArray; /** The single in-flight lazy build of `paintables`. */ paintablesBuild?: Promise<PaintableArray>; /** Stage 4a (R5): an XLSX entry's projected parse (engine values), retained for the lazy units paintables. */ parsed?: XlsxDocument; /** Stage 4a (F1): the projected package's saved-values parse (`valuePolicy 'cached'`), built by a tiles-mode open; every sheet snapshot reads it, never `parsed`. */ savedParse?: XlsxDocument; /** The single in-flight build of `savedParse` for a tiles open joining a units-created entry. */ savedParseBuild?: Promise<XlsxDocument>; /** Sheet snapshots by `sheetIndex|day|timeZone`, one promise each, released on `drop`. */ snapshots?: Map<string, Promise<SheetSnapshot>>; /** Each saved-parse sheet's coverage and blanked copy, computed once (clock-free) and shared by every snapshot's view. */ servable?: Map<number, ServableSheet>; /** One bounded text-layout cache per snapshot key, shared by its tiles (a layout is a pure function of its key). */ tileLayouts?: Map<string, LayoutCache>; /** Stage 4b: the snapshot geometries at session sizes, by `snapshot key|sha256(sizes)`, at most `SIZED_GEOMETRIES` (LRU), released on `drop`. */ sizedGeometries?: Map<string, SheetGeometryResult>; /** Stage 4b: text views by `snapshot key|sha256(sizes)`, at most `TEXT_VIEWS` (LRU). */ textViews?: Map<string, TextViewEntry>; /** Stage 4b: the formatting budget every text view of the entry reserves from. */ textBudget?: FormatBudget; /** Stage 4c: `sheet-view` ports by `sheetIndex|day|timeZone`, at most `SHEET_VIEW_PORTS` (LRU). */ viewPorts?: Map<string, SheetViewEntry>; /** Stage 4c: derived views by `port key|canonical hash`, at most `DERIVED_VIEWS` (LRU). */ derivedViews?: Map<string, DerivedView>; /** Stage 4a: an XLSX entry's projected (visible) sheet names in tab order. */ sheets?: readonly string[]; units: UnitSize[]; hidden: number[]; raw: Array<{ widthPx: number; heightPx: number }>; fingerprint: string; contentPolicyId: string; diagnostics: PolicyDiagnostic[]; format: 'docx' | 'xlsx' | 'pptx'; notes?: { pkg: OfficePackage; slideParts: readonly string[] } }
/** Piece 5 (D22): a unit's AUDIO uses (an audio node, a `p:sndTgt` sound, a transition sound) and their probed entries.
 * Video nodes are never listed. `spid`: the shape an audio node plays on (none for a sound). */
/** Fix round 1: derived from `slideAudioUses`, the same definition as the plan's records. `identity`: what the plan's
 * `mediaId` hashes (`mediaIdentity`); `placeholder`: its placeholder in the cached plan. A part-less use (refused
 * `media-linked` or `media-unreadable`) has a synthetic `k` (≥ the part count) and `partPath: ''`. */
export interface MediaUnitRefs { unit: number; hidden: boolean; refs: Array<{ k: number; partPath: string; spid?: string; hiddenShape: boolean; entry: DocMediaEntry; identity: string; placeholder: string }> }
/** A notes job (Stage 3c): `draw: false` plans only (size, truncation, characters drawn); `true` draws with the visible mark. */
export interface NotesJob { unitIndex: number; scale: number; mark: MarkSpec; draw: boolean }
export interface NotesPlan { kind: 'notes-plan'; w: number; h: number; truncated: boolean; chars: number }
export type NotesRaster = RenderedUnit & { kind: 'raster'; truncated: boolean; chars: number }
/** `hidden`: unit indices of hidden slides (pptx under `'full'` only; elsewhere they are not units). */
/** `anim` (piece 4 D5): shown units with timing, a transition or both (a plan is requested only for these).
 * `morph` (Task 7): the shown units entered by a Morph transition (a `show-morph` names one as its `to`). */
/** `media` (piece 5, D19): the shown units whose plan ships an audio record or a transition sound (structural: from
 * the slide's metadata and package, no probe), so the manifest marks them `media` and a show asks for their plan. */
export interface OpenResult { docKey: string; units: UnitSize[]; hidden: number[]; anim: number[]; morph: number[]; media: number[]; layoutFingerprint: string; contentPolicyId: string; format: 'docx' | 'xlsx' | 'pptx'; diagnostics: PolicyDiagnostic[]; /** 0 until the paintables exist (a tiles-mode XLSX open builds none). */ fontDiagnostics: number; /** XLSX: the projected (visible) sheet names in tab order. */ sheets?: readonly string[] }
/** Stage 4a (R5): how an XLSX open treats its static per-sheet units. `units` builds them at open (today);
 * `tiles` builds none, sizing the units from `staticSheetSpec`, and the first units use builds them. */
export type SheetMode = 'units' | 'tiles'
/** A units path reached an entry whose paintables are not built yet: `ensurePaintables` first, then retry. */
export class PaintablesUnbuilt extends Error { constructor() { super('paintables not built'); this.name = 'PaintablesUnbuilt' } }

const registerFont = () => { throw new Error('stream server: embedded fonts are not registered in Stage 1') }

function readImage(source: CanvasImageSource, backend: CanvasBackend): ImageSnapshot {
  const image = source as unknown as { width: number; height: number; naturalWidth?: number; naturalHeight?: number }
  const w = image.naturalWidth ?? image.width, h = image.naturalHeight ?? image.height
  const c = backend.createCanvas(w, h), ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.drawImage(source, 0, 0)
  return { w, h, rgba: new Uint8Array(ctx.getImageData(0, 0, w, h).data) }
}

/** Metadata-free PNG (§7.7): keep the critical chunks (IHDR, PLTE, IDAT,
 * IEND) and tRNS; drop every other ancillary chunk (sRGB, sBIT, text, time…).
 * Browsers treat an untagged PNG as sRGB, so pixels are unchanged. */
export function criticalChunksOnly(png: Uint8Array): Uint8Array {
  const keep = new Set(['IHDR', 'PLTE', 'IDAT', 'IEND', 'tRNS'])
  const v = new DataView(png.buffer, png.byteOffset, png.byteLength), parts: Uint8Array[] = [png.subarray(0, 8)]
  for (let at = 8; at + 12 <= png.length;) {
    const len = v.getUint32(at), end = at + 12 + len
    if (keep.has(String.fromCharCode(...png.subarray(at + 4, at + 8)))) parts.push(png.subarray(at, end))
    at = end
  }
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0
  for (const p of parts) { out.set(p, o); o += p.length }
  return out
}

/** Hidden slides of a deck opened under `'full'`, read from the original
 * package (there the projection is the identity). A slide list that does not
 * line up with the units refuses the open: never guess which unit is hidden. */
export async function hiddenUnits(doc: PptxDocument, units: number): Promise<number[]> {
  const handle = pptxPackageHandle(doc)
  if (!handle || handle.slideParts.length !== units) throw new RenderRefusal('unsupported-format', 'slide list does not match the units')
  return hiddenSlideUnits(handle)
}

/** An unbounded lease ledger for the show metadata store: its own 64 MiB aggregate bounds what it retains. */
const SHOW_META_WORKSPACE: SlideShowWorkspace = {
  limit: Number.POSITIVE_INFINITY, used: () => 0, reserveTerminal: () => undefined,
  lease: (bytes, tag) => { const l: Lease = { tag, bytes, grow: () => true, shrink() {}, transfer: () => l, release() {} }; return l },
}

/** A deck's show inputs over its projected package (piece 4): the show index, the metadata store, the caches. */
async function openShow(docRef: string, docKey: string, contentPolicyId: string, fingerprint: string, doc: PptxDocument, handle: NonNullable<ReturnType<typeof pptxPackageHandle>>, units: UnitSize[], hidden: number[]): Promise<ShowDoc> {
  const sink = createDiagnosticSink()
  const index = await readShowIndex(doc, handle, sink)
  // The resolver sees the slide's layout and master artwork, exactly as local mode's (`createUnderlayReader`).
  const underlay = createUnderlayReader(handle, index, sink)
  const store = createSlideMetadataStore({ doc, handle, index, workspace: SHOW_META_WORKSPACE, sink, timing: readSlideTiming, resolver: (slide) => shapeResolver(slide, underlay.underlays.get(doc.slides.indexOf(slide)) ?? []) })
  return createShowDoc({ docRef, docKey, contentPolicyId, fingerprint, doc, handle, index, store, acquire: (i) => underlay.acquire(store, i), sink, units, hidden })
}

/** Ruling B1 (Task 6 fix round 1): the static size of the refusal unit (`REFUSAL_SPEC`, src/render/paint.ts). */
export const TILES_REFUSAL_UNIT: Readonly<UnitSize> = Object.freeze({ w: 480, h: 96 })
/** Ruling B1: what a tiles-mode XLSX open advertises. A unit over MAX_UNIT_PX (it could never be rendered) is listed as
 * the refusal unit, as an `exportLimits` sheet is, so the manifest validates; its print still refuses (`render`
 * reads the entry's own size and refuses it `oversize`). Units mode returns the entry's sizes unchanged. */
export function tilesUnits(units: readonly UnitSize[]): UnitSize[] {
  return units.map((u) => unitSurfaceAllowed(u.w, u.h) ? u : { ...TILES_REFUSAL_UNIT })
}

/** Marks `e` evicted; it is released now if no job holds it, else by the last job that does. */
function evictTextView(e: TextViewEntry): void {
  e.evicted = true
  if (e.uses === 0) e.view.then((v) => v.release(), () => undefined)
}
/** Marks `e` evicted; its port is disposed (and its key view released) now if no job holds it, else by the last job. */
function evictViewPort(e: SheetViewEntry): void {
  e.evicted = true
  if (e.uses === 0) void closeViewPort(e)
}
function closeViewPort(e: SheetViewEntry): Promise<void> {
  return e.port.then(async ({ port, keyView }) => { try { await port.dispose() } finally { keyView.release() } }, () => undefined)
}
/** A formatting budget that refuses (throws) past `max` bytes held. */
function boundedBudget(max: number): FormatBudget {
  let held = 0
  return {
    reserve(bytes) { if (held + bytes > max) throw new RangeError('sheet text: formatting exceeds its budget'); held += bytes },
    release(bytes) { held -= bytes },
  }
}

export class DocumentHost {
  readonly backend: CanvasBackend
  private docs = new Map<string, OpenDoc>()
  private served: ServedCache
  parses = 0
  pages = 0
  /** TEST-visible: static paintable sets built (an eager open's, or a lazy build's). */
  paintablesBuilt = 0
  /** Reads a 'full' deck's hidden units after its paintables exist (TEST-ONLY override: fault injection). */
  hiddenReader: (doc: PptxDocument, paintables: PaintableArray) => Promise<number[]> = (doc, paintables) => hiddenUnits(doc, paintables.length)
  constructor(backend = serverBackend(), servedCacheBytes?: number) { this.backend = backend; installCanvasBackend(gateBackend(backend)); this.served = new ServedCache(servedCacheBytes) }

  async open(docRef: string, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full', o: { sheetMode?: SheetMode } = {}): Promise<OpenResult> {
    const sheetMode = o.sheetMode ?? 'units'
    const contentPolicyId = contentPolicy === 'full' ? 'full@1' : 'visible-only@1'
    const fingerprint = layoutFingerprint(contentPolicyId)
    const key = `${docRef}|${contentPolicyId}|${fingerprint}`
    let doc = this.docs.get(key)
    if (!doc) {
      const admitted = admit(bytes)
      if (!admitted.ok) throw new RenderRefusal('unsupported-format')
      const original = await OfficePackage.load(bytes)
      const projected = await projectPackage(original, admitted.format, contentPolicy)
      const parsed = await detectAndParse(projected.pkg)
      // R5: a tiles-mode XLSX open builds no paintables; its units are sized from the sheet metrics alone.
      const lazy = admitted.format === 'xlsx' && sheetMode === 'tiles'
      // F1: tiles paint saved values (S4 §6.2 step 7), so a tiles open also parses the projected package under 'cached'.
      const savedParse = lazy ? await parseXlsx(projected.pkg, { valuePolicy: 'cached' }) : undefined
      const xlsx = admitted.format === 'xlsx' ? parsed as XlsxDocument : undefined
      const paintables = lazy ? undefined : await getPaintables(parsed, { registerFont })
      this.parses++
      if (paintables) this.paintablesBuilt++
      const raw = paintables ? paintables.map((p) => ({ widthPx: p.spec.widthPx, heightPx: p.spec.heightPx })) : xlsx!.sheets.map(staticSheetSpec)
      let hidden: number[] = []
      // The lease is ours until the doc is cached: a refusal here must release it.
      if (contentPolicy === 'full' && admitted.format === 'pptx') {
        try { hidden = await this.hiddenReader(parsed as PptxDocument, paintables!) } catch (e) { paintables!.dispose(); throw e }
      }
      // Projected slide parts keep their original paths; a list that does not line up with the units serves no notes.
      const handle = admitted.format === 'pptx' ? pptxPackageHandle(parsed as PptxDocument) : undefined
      const notes = handle && handle.slideParts.length === raw.length ? { pkg: original, slideParts: [...handle.slideParts] } : undefined
      // Projection read every original entry through the cache: the retained package keeps the archive, not those reads.
      original.dropCache()
      const units = raw.map((r) => ({ w: sizeCeil(Math.max(1, r.widthPx)), h: sizeCeil(Math.max(1, r.heightPx)) }))
      let show: ShowDoc | undefined
      const anim: number[] = [], morph: number[] = [], mediaUnits: number[] = []
      if (handle && handle.slideParts.length === raw.length) {
        try {
          show = await openShow(docRef, key, projected.contentPolicyId, fingerprint, parsed as PptxDocument, handle, units, hidden)
          for (let i = 0; i < units.length; i++) {
            if (hidden.includes(i)) continue
            const h = await show.acquire(i)
            try {
              if ((h.slide.timing && !('refused' in h.slide.timing)) || h.slide.transition) anim.push(i)
              if (h.slide.transition?.element === 'morph') morph.push(i)
              // D19: structural (no probe; any k will do): does the plan ship a record or a transition sound?
              // A slide whose package cannot be read for it is not flagged (media only; its animation is untouched).
              if (hasAudioUse(h.slide)) {
                try {
                  const planned = slideAudioUses(h.slide, i, await slideMediaContext(show, i, projected.contentPolicyId), () => 0, 0)
                  if (planned.records.some((r) => !r.hiddenShape)) mediaUnits.push(i)
                } catch { /* not flagged */ }
              }
            } finally { h.release() }
          }
        } catch {
          // No readable show index or metadata: every slide stays a piece-3 still (fail closed).
          show?.store.close()
          show = undefined
          anim.length = 0
          morph.length = 0
          mediaUnits.length = 0
        }
      }
      const opened = { key, docRef, ...(paintables ? { paintables } : {}), ...(xlsx ? { parsed: xlsx, sheets: Object.freeze(xlsx.sheets.map((s) => s.name)) } : {}), ...(savedParse ? { savedParse } : {}), raw, hidden, anim, morph, mediaUnits, ...(notes ? { notes } : {}), ...(show ? { show } : {}), units, fingerprint, contentPolicyId: projected.contentPolicyId, diagnostics: projected.diagnostics, format: admitted.format } as Omit<OpenDoc, 'media'> as OpenDoc
      opened.media = this.mediaTableFor(opened)
      if (show) show.media = opened.media
      doc = opened
      this.docs.set(key, doc)
    }
    // A units-mode open stays eager (R5), also when it joins an entry a tiles open created.
    if (sheetMode === 'units') await this.buildPaintables(doc)
    // A tiles open joining a units-created entry: the entry kept no projected package, so re-project these bytes (F1).
    else if (doc.format === 'xlsx' && !doc.savedParse) await this.buildSavedParse(doc, bytes, contentPolicy)
    return { docKey: key, units: sheetMode === 'tiles' && doc.format === 'xlsx' ? tilesUnits(doc.units) : doc.units, hidden: doc.hidden, anim: [...doc.anim], morph: [...doc.morph], media: [...doc.mediaUnits], layoutFingerprint: doc.fingerprint, contentPolicyId: doc.contentPolicyId, format: doc.format, diagnostics: doc.diagnostics, fontDiagnostics: doc.paintables?.fontDiagnostics?.length ?? 0, ...(doc.sheets ? { sheets: [...doc.sheets] } : {}) }
  }

  /** Builds the entry's static paintables if a tiles-mode open left them out (R5); every units path awaits this first. */
  async ensurePaintables(docKey: string): Promise<void> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    await this.buildPaintables(doc)
  }

  /** The entry's single build: concurrent first uses share one promise; a failed build is retried by the next use.
   * A build that settles after its entry was dropped disposes what it built. */
  private buildPaintables(doc: OpenDoc): Promise<PaintableArray> {
    if (doc.paintables) return Promise.resolve(doc.paintables)
    doc.paintablesBuild ??= getPaintables(doc.parsed!, { registerFont }).then((p) => {
      this.paintablesBuilt++
      doc.paintablesBuild = undefined
      if (this.docs.get(doc.key) !== doc) { p.dispose(); throw new RenderRefusal('unknown-doc') }
      doc.paintables = p
      return p
    }, (e: unknown) => { doc.paintablesBuild = undefined; throw e })
    return doc.paintablesBuild
  }

  /** The entry's single saved-values parse build (F1), from the bytes a joining tiles open brought: the same projection
   * as the entry's own. A build that settles after its entry was dropped is discarded. Not counted in `parses`. */
  private buildSavedParse(doc: OpenDoc, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full'): Promise<XlsxDocument> {
    if (doc.savedParse) return Promise.resolve(doc.savedParse)
    doc.savedParseBuild ??= (async () => {
      const original = await OfficePackage.load(bytes)
      const projected = await projectPackage(original, 'xlsx', contentPolicy)
      return parseXlsx(projected.pkg, { valuePolicy: 'cached' })
    })().then((saved) => {
      doc.savedParseBuild = undefined
      if (this.docs.get(doc.key) !== doc) throw new RenderRefusal('unknown-doc')
      doc.savedParse = saved
      return saved
    }, (e: unknown) => { doc.savedParseBuild = undefined; throw e })
    return doc.savedParseBuild
  }

  /**
   * Stage 4a (S4 §6.2): the servable snapshot of sheet `sheetIndex` (an index into the entry's projected sheets, as
   * `OpenResult.sheets`) under the session clock. Built from `savedParse` only; one promise per
   * `(docKey, sheetIndex, day, timeZone)`, released on `drop`. Callers never release it themselves.
   */
  sheetSnapshot(docKey: string, sheetIndex: number, clock: SheetClock): Promise<SheetSnapshot> {
    const doc = this.docs.get(docKey)
    if (!doc) return Promise.reject(new RenderRefusal('unknown-doc'))
    const saved = doc.savedParse
    if (doc.format !== 'xlsx' || !saved) return Promise.reject(new RenderRefusal('bad-job', 'no saved-values parse: the entry was not opened in tiles mode'))
    if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= saved.sheets.length) return Promise.reject(new RenderRefusal('bad-job', 'sheet index'))
    const key = `${sheetIndex}|${clock.day}|${clock.timeZone}`
    const snapshots = doc.snapshots ??= new Map()
    let built = snapshots.get(key)
    if (!built) {
      const contentPolicy = doc.contentPolicyId === 'full@1' ? 'full' as const : 'visible-only' as const
      const cfClock = { epochNowMs: clockEpochMs(clock), timeZone: clock.timeZone }
      const servable = (i: number): ServableSheet => this.servableOf(doc, saved, i)
      built = buildSheetSnapshot(saved, sheetIndex, contentPolicy, clock, () => {
        // The lease reads `doc.sheets[sheetIndex]`, and CF's cross-sheet rules read the other sheets: every slot holds its
        // servable (blanked) sheet, so no sheet's covered cell reaches any rule (owner 11.1). Chart contexts stay shared.
        return prepareSheetAssets(this.servableWorkbook(doc, saved), { sheetIndex, registerFont, cfClock, chartOwner: saved })
      }, servable).then((snap) => {
        // Dropped while building: nothing will release it later.
        if (this.docs.get(docKey) !== doc) { snap.release(); throw new RenderRefusal('unknown-doc') }
        return snap
      })
      snapshots.set(key, built)
      built.catch(() => { if (snapshots.get(key) === built) snapshots.delete(key) })
    }
    return built
  }

  /** Each saved-parse sheet's coverage and blanked copy, computed once (clock-free). */
  private servableOf(doc: OpenDoc, saved: XlsxDocument, i: number): ServableSheet {
    const cache = doc.servable ??= new Map()
    let s = cache.get(i)
    if (!s) cache.set(i, s = servableSheet(saved.sheets[i]!, doc.contentPolicyId === 'full@1' ? 'full' : 'visible-only'))
    return s
  }
  /** The servable workbook: every slot its servable (blanked) sheet (`saved` itself when none is blanked). */
  private servableWorkbook(doc: OpenDoc, saved: XlsxDocument): XlsxDocument {
    const sheets = saved.sheets.map((_, i) => this.servableOf(doc, saved, i).sheet)
    return sheets.every((s, i) => s === saved.sheets[i]) ? saved : { ...saved, sheets }
  }

  /**
   * Stage 4b (P2, P9): one `sheet-text` op over the text view of its sheet at the job's session sizes and (Stage 4c,
   * P9) sort and filter view, under the session clock (no clock is read). Views are built from the cached snapshots
   * (`sheetTextView`) and cached per `(sheetIndex, clock, sha256(sizes), canonical view hash)`, at most `TEXT_VIEWS` per document (LRU), released on eviction and on `drop`;
   * a view a running job holds is released by that job when it ends. Fit measures and lays out on the server backend's canvases.
   */
  async sheetText(docKey: string, clock: SheetClock, job: SheetTextJob): Promise<SheetTextReply> {
    const held: TextViewEntry[] = []
    try {
      return await answerSheetText(job, {
        viewOf: (sheetIndex, sizes, view) => {
          const e = this.textViewEntry(docKey, sheetIndex, clock, sizes, view)
          e.uses++; held.push(e)
          return e.view
        },
        // Find's text document is built over the servable workbook, as the views are.
        workbook: () => {
          const doc = this.docs.get(docKey), saved = doc?.savedParse
          if (!doc || !saved) throw new RenderRefusal('unknown-doc')
          return this.servableWorkbook(doc, saved)
        },
        createCanvas: () => this.backend.createCanvas(1, 1),
      })
    } finally {
      for (const e of held) if (--e.uses === 0 && e.evicted) e.view.then((v) => v.release(), () => undefined)
    }
  }

  /**
   * Stage 4c (P3, P5, P6): one `sheet-view` job (derive, sort, filter, value list, range detection) through the local
   * engine over the key view of its sheet under the session clock. One port per `(sheetIndex, clock)`, at most
   * `SHEET_VIEW_PORTS` per document (LRU); derived views per `(sheetIndex, clock, canonical hash)`, at most
   * `DERIVED_VIEWS` per document (LRU). Both are released on `drop`. `contentPolicy` must be the entry's own.
   */
  sheetView(docKey: string, clock: SheetClock, contentPolicy: 'visible-only' | 'full', job: SheetViewJob): Promise<SheetViewReply> {
    return this.viewJob(docKey, clock, contentPolicy, job)
  }

  /**
   * Stage 4c (P9): the derived view of `view` on sheet `sheetIndex` under `clock`, from the derived-view cache or by a
   * `derive` job on the sheet's port (the entry's own content policy). A replay the engine refuses fails (P3).
   */
  private async derivedView(docKey: string, sheetIndex: number, clock: SheetClock, view: CanonicalView): Promise<DerivedView> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    let derived: DerivedView | undefined
    const r = await this.viewJob(docKey, clock, doc.contentPolicyId === 'full@1' ? 'full' : 'visible-only', { op: 'derive', sheetIndex, view }, (d) => { derived = d })
    if (r.op !== 'view' || !('order' in r) || !derived) throw new RenderRefusal('bad-job', 'sheet view: no derived view')
    return derived
  }

  /** One `sheet-view` job; `seen` gets each derived view the job reads from or puts in the cache. */
  private async viewJob(docKey: string, clock: SheetClock, contentPolicy: 'visible-only' | 'full', job: SheetViewJob, seen?: (d: DerivedView) => void): Promise<SheetViewReply> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if ((doc.contentPolicyId === 'full@1') !== (contentPolicy === 'full')) throw new RenderRefusal('bad-job', 'sheet-view: content policy')
    const portKey = `${job.sheetIndex}|${clock.day}|${clock.timeZone}`
    const e = this.viewPortEntry(doc, docKey, portKey, job.sheetIndex, clock, contentPolicy)
    e.uses++
    try {
      const { port } = await e.port
      const derived = doc.derivedViews ??= new Map()
      const cache: DerivedCache = {
        get: (view) => {
          const key = `${portKey}|${canonicalViewHash(view)}`
          const hit = derived.get(key)
          if (hit) { derived.delete(key); derived.set(key, hit); seen?.(hit) }
          return hit
        },
        put: (d) => {
          seen?.(d)
          if (this.docs.get(docKey) !== doc) return
          const key = `${portKey}|${canonicalViewHash(d.view)}`
          derived.delete(key)
          derived.set(key, d)
          while (derived.size > DERIVED_VIEWS) derived.delete(derived.keys().next().value!)
        },
      }
      return await port.job(job, cache)
    } finally {
      if (--e.uses === 0 && e.evicted) void closeViewPort(e)
    }
  }

  /** The cached `sheet-view` port of `(sheetIndex, clock)`, built on a miss (a failed build is not kept). */
  private viewPortEntry(doc: OpenDoc, docKey: string, key: string, sheetIndex: number, clock: SheetClock, policy: 'visible-only' | 'full'): SheetViewEntry {
    const ports = doc.viewPorts ??= new Map()
    const hit = ports.get(key)
    if (hit) { ports.delete(key); ports.set(key, hit); return hit }
    const budget = doc.textBudget ??= boundedBudget(TEXT_VIEW_FORMAT_BYTES)
    const entry: SheetViewEntry = { uses: 0, evicted: false, port: this.sheetSnapshot(docKey, sheetIndex, clock).then(async (snap) => {
      const saved = doc.savedParse
      if (!saved || this.docs.get(docKey) !== doc) throw new RenderRefusal('unknown-doc')
      const workbook = this.servableWorkbook(doc, saved)
      const keyView = await sheetKeyView(snap, saved.sheets[sheetIndex]!, { workbook, clock, budget })
      let port: SheetViewPort
      try {
        port = await openSheetViewPort({ sheetIndex, policy, clock, doc: keyDocument(workbook, new Map([[sheetIndex, keyView]])), geometry: snap.geometry })
      } catch (e) { keyView.release(); throw e }
      // Dropped while building: nothing else will dispose it.
      if (this.docs.get(docKey) !== doc) {
        try { await port.dispose() } finally { keyView.release() }
        throw new RenderRefusal('unknown-doc')
      }
      return { port, keyView }
    }) }
    ports.set(key, entry)
    entry.port.catch(() => { if (ports.get(key) === entry) ports.delete(key) })
    while (ports.size > SHEET_VIEW_PORTS) {
      const [oldKey, old] = ports.entries().next().value!
      ports.delete(oldKey)
      evictViewPort(old)
    }
    return entry
  }

  /** The cached text view entry of `(sheetIndex, clock, sizes, view)`, built on a miss (a failed build is not kept).
   * Stage 4c (P9): a view other than the identity view is built over its derived order and visibility. */
  private textViewEntry(docKey: string, sheetIndex: number, clock: SheetClock, sizes: WireSizes, canonical: CanonicalView = IDENTITY_VIEW): TextViewEntry {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    const views = doc.textViews ??= new Map()
    const identity = isIdentityView(canonicalView(canonical))
    const key = `${sheetIndex}|${clock.day}|${clock.timeZone}|${sizesHash(sizes)}|${identity ? IDENTITY_HASH : canonicalViewHash(canonical)}`
    const hit = views.get(key)
    if (hit) { views.delete(key); views.set(key, hit); return hit }
    const budget = doc.textBudget ??= boundedBudget(TEXT_VIEW_FORMAT_BYTES)
    const entry: TextViewEntry = { uses: 0, evicted: false, view: this.sheetSnapshot(docKey, sheetIndex, clock).then(async (snap) => {
      const saved = doc.savedParse
      if (!saved || this.docs.get(docKey) !== doc) throw new RenderRefusal('unknown-doc')
      const derived = identity ? undefined : await this.derivedView(docKey, sheetIndex, clock, canonical)
      if (this.docs.get(docKey) !== doc) throw new RenderRefusal('unknown-doc')
      const view = await sheetTextView(snap, saved.sheets[sheetIndex]!, sessionSizesOf(sizes), { workbook: this.servableWorkbook(doc, saved), clock, budget, ...(derived ? { order: derived.order, visibility: derived.visibility } : {}) })
      // Dropped while building: nothing else will release it.
      if (this.docs.get(docKey) !== doc) { view.release(); throw new RenderRefusal('unknown-doc') }
      return view
    }) }
    views.set(key, entry)
    entry.view.catch(() => { if (views.get(key) === entry) views.delete(key) })
    while (views.size > TEXT_VIEWS) {
      const [oldKey, old] = views.entries().next().value!
      views.delete(oldKey)
      evictTextView(old)
    }
    return entry
  }

  /**
   * Stage 4a (S4 §6.4, F23): the `sheet-layout` of sheet `sheetIndex` (wire `sheetId` = its decimal index among the
   * visible sheets, R12), from its servable snapshot under the session clock and the same parse's declared filters.
   * A layout too fragmented to fit any answer is `RenderRefusal('layout-too-large')`; a diagnostic geometry `bad-job`.
   * The caller encodes (and applies the R14 caps to) the result.
   */
  async sheetLayout(docKey: string, sheetIndex: number, clock: SheetClock): Promise<ServerSheetLayout> {
    const snap = await this.sheetSnapshot(docKey, sheetIndex, clock)
    const saved = this.docs.get(docKey)?.savedParse
    if (!saved) throw new RenderRefusal('unknown-doc')
    const sheet = saved.sheets[sheetIndex]!
    const declared = saved.declaredFilters?.find((d) => sheet.sourcePartPath !== undefined ? d.partPath === sheet.sourcePartPath : d.sheetId === sheet.sheetId)
    try {
      const layout: ServerSheetLayout = buildStreamLayout(snap, String(sheetIndex), declared)
      // Stage 4b: server-only (never encoded), for the `coverage-unknown` refusal of `sheet-resize`.
      if (snap.coverage !== 'off' && snap.coverage.kind === 'unknown') layout.coverageUnknown = true
      return layout
    } catch (e) {
      if (e instanceof LayoutTooLarge) throw new RenderRefusal('layout-too-large', e.message)
      if (e instanceof LayoutUnavailable) throw new RenderRefusal('bad-job', e.message)
      throw e
    }
  }

  /**
   * Stage 4a (S4 §6.5): one 512² device-px tile of sheet `job.sheetIndex`, painted from its servable snapshot under
   * `job.clock` only (covered cells blank, saved values, F16/F19), then the visible mark in global sheet device space
   * (`paintMarkRegion`), encoded metadata-free. Text layouts are cached per snapshot (`OpenDoc.tileLayouts`). The logical box is `(col, row)·512/s` with `s = zoom·dpr`; the plan is
   * profile `'D'` against a fresh per-tile `frameBudgetFor('D')`. An edge tile is a full 512² surface, blank past the
   * sheet's extent (R15). A blank tile is marked too (owner ruling). Refused `bad-job`: a malformed job, a sheet out of
   * range, `zoomStep` outside 0..10, `dpr` outside {1, 2}, an unknown palette, a `viewGen` other than 0 without `sizes` or `view` (4a, F5),
   * malformed `sizes` (4b: the tile is painted over the snapshot's geometry at those sizes, `sizedGeometry`; Stage 4c, P9:
   * in the derived order and visibility of a `view` other than the identity view), a malformed `view`, a negative index
   * or one past `tileLimit` (F8), or a sheet whose geometry is a diagnostic (it has no layout either). Reads no clock.
   */
  async renderTile(docKey: string, job: TileJob): Promise<RenderedTile> {
    const problem = tileJobProblem(job)
    if (problem) throw new RenderRefusal('bad-job', `tile: ${problem}`)
    const snap = await this.sheetSnapshot(docKey, job.sheetIndex, job.clock)
    if (snap.geometry.kind !== 'ready') throw new RenderRefusal('bad-job', 'tile: sheet geometry unavailable')
    // Stage 4c (P9): a view other than the identity view paints over its derived order and visibility.
    const derived = job.view && !isIdentityView(canonicalView(job.view)) ? await this.derivedView(docKey, job.sheetIndex, job.clock, job.view) : undefined
    const geometry = derived || job.sizes?.length ? this.sizedGeometry(docKey, snap, job.sizes ?? [], derived) : snap.geometry
    if (geometry.kind !== 'ready') throw new RenderRefusal('bad-job', 'tile: sheet geometry unavailable')
    const zoom = TILE_ZOOM_STEPS[job.zoomStep]!, dpr = job.dpr, s = zoom * dpr
    const m = geometry.metrics
    if (job.col >= tileLimit(m.cols.totalCss(), zoom, dpr) || job.row >= tileLimit(m.rows.totalCss(), zoom, dpr)) throw new RenderRefusal('bad-job', 'tile: index past the grid')
    const canvas = this.backend.createCanvas(TILE_PX, TILE_PX)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    const viewport: LogicalViewport = { box: tileLogicalBox(job.col, job.row, zoom, dpr), zoom, deviceScale: dpr }
    const doc = this.docs.get(docKey)
    // Dropped while the snapshot was awaited.
    if (!doc) throw new RenderRefusal('unknown-doc')
    const layouts = doc.tileLayouts ??= new Map()
    const layoutKey = `${job.sheetIndex}|${job.clock.day}|${job.clock.timeZone}`
    let cache = layouts.get(layoutKey)
    if (!cache) layouts.set(layoutKey, cache = new LayoutCache())
    const frameBudget = frameBudgetFor('D')
    const plan = planViewport(geometry, viewport, 'D', { frameBudget })
    ctx.setTransform(s, 0, 0, s, 0, 0)
    renderSheetViewport(snap.sheet, ctx, geometry, viewport, snap.assets, { colors: TILE_PALETTE_COLORS[job.palette], preparedPlan: plan, profile: 'D', cache, frameBudget })
    const mark = markOptions(job.mark)
    if (mark) {
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      paintMarkRegion(ctx, mark, { x: job.col * TILE_PX, y: job.row * TILE_PX, w: TILE_PX, h: TILE_PX, scale: s })
    }
    this.pages++
    const png = criticalChunksOnly(new Uint8Array((canvas as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png')))
    return { png, w: TILE_PX, h: TILE_PX, palette: job.palette }
  }

  /**
   * Stage 4b (P5): the snapshot's geometry at the session's `sizes` (canonical): `buildSheetGeometry` over the
   * snapshot's sheet with those sizes, so its drawings are placed by them too (`sizedDrawings`). Stage 4c (P9): with a
   * derived view, in its order and visibility. Cached per `(snapshot key, sha256(sizes), canonical view hash)`, at most
   * `SIZED_GEOMETRIES` per document (LRU), released on `drop`.
   */
  private sizedGeometry(docKey: string, snap: SheetSnapshot, sizes: WireSizes, derived?: DerivedView): SheetGeometryResult {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    const cache = doc.sizedGeometries ??= new Map()
    const key = `${snap.sheetIndex}|${snap.clock.day}|${snap.clock.timeZone}|${sizesHash(sizes)}|${derived ? canonicalViewHash(derived.view) : IDENTITY_HASH}`
    let g = cache.get(key)
    if (g) { cache.delete(key); cache.set(key, g); return g }
    g = buildSheetGeometry(snap.sheet, { sizes: sessionSizesOf(sizes), ...(derived ? { order: derived.order, visibility: derived.visibility } : {}) })
    cache.set(key, g)
    while (cache.size > SIZED_GEOMETRIES) cache.delete(cache.keys().next().value!)
    return g
  }

  /** The static paintables every units path paints from (synchronous: `ensurePaintables` builds them). */
  private paintablesOf(doc: OpenDoc): PaintableArray {
    if (!doc.paintables) throw new PaintablesUnbuilt()
    return doc.paintables
  }

  has(docKey: string) { return this.docs.has(docKey) }
  drop(docKey: string) {
    const d = this.docs.get(docKey)
    if (d) {
      this.droppedProbes += d.media.probes; d.paintables?.dispose(); d.show?.store.close(); this.docs.delete(docKey)
      for (const built of d.snapshots?.values() ?? []) built.then((snap) => snap.release(), () => undefined)
      d.snapshots?.clear()
      d.tileLayouts?.clear()
      d.sizedGeometries?.clear()
      for (const e of d.textViews?.values() ?? []) evictTextView(e)
      d.textViews?.clear()
      for (const e of d.viewPorts?.values() ?? []) evictViewPort(e)
      d.viewPorts?.clear()
      d.derivedViews?.clear()
    }
    this.served.drop(docKey)
  }
  get docCount() { return this.docs.size }

  render(docKey: string, job: { unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: readonly string[] }, observers: readonly PaintObserver[] = []): RenderedUnit {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!Number.isInteger(job.unitIndex) || job.unitIndex < 0 || job.unitIndex >= doc.units.length) throw new RenderRefusal('unit-range')
    if (!(SCALE_BUCKETS as readonly number[]).includes(job.scale)) throw new RenderRefusal('scale')
    const r = doc.raw[job.unitIndex]!
    const w = sizeCeil(Math.max(1, r.widthPx * job.scale)), h = sizeCeil(Math.max(1, r.heightPx * job.scale))
    if (!unitSurfaceAllowed(w, h)) throw new RenderRefusal('oversize')
    const paintable = this.paintablesOf(doc)[job.unitIndex]!
    const page = { widthPx: r.widthPx, heightPx: r.heightPx }
    const mark = markOptions(job.mark)
    this.pages++
    if (job.projection === 'commands') {
      const cmd = this.recordCommands(paintable, job.unitIndex, w, h, job.scale, page, mark, observers)
      if (cmd.ok) {
        const use = facesUsed(cmd.ops)
        if (use.unpinned.length === 0 && !use.faces.some((f) => job.missingFaces?.includes(f))) return { kind: 'commands', w, h, ops: cmd.ops, opCount: cmd.opCount, assets: cmd.assets, faces: use.faces, evidence: cmd.evidence }
        return { ...this.raster(paintable, job.unitIndex, w, h, job.scale, page, mark, observers), fallback: 'face-missing' }
      }
      return { ...this.raster(paintable, job.unitIndex, w, h, job.scale, page, mark, observers), fallback: `recorder:${cmd.reason}` }
    }
    return this.raster(paintable, job.unitIndex, w, h, job.scale, page, mark, observers)
  }

  /** Piece 5 (D21): ffmpeg and the media temp directory, from the worker init; undefined = media off. */
  /** `runner`: how a probe is run (a render worker's asks the pool; default this thread, fix round 2). */
  private mediaConfig?: { info: FfmpegInfo; dir: string; temp: MediaTemp; runner?: ProbeRunner }
  /** TEST-ONLY seams for the probe table (a clock, a held probe, a run observer); empty in production. */
  mediaSeams: MediaTableSeams = {}
  /** Sets ffmpeg and the media dir (a view over the pool's directory: a worker never claims one). Documents opened
   * afterwards get a table over them; with no config every entry is `media-unavailable`. */
  configureMedia(media: { info: FfmpegInfo; dir: string; prefix?: string } | undefined, runner?: ProbeRunner): void {
    this.mediaConfig = media ? { info: media.info, dir: media.dir, temp: attachMediaTemp(media.dir, media.prefix), ...(runner ? { runner } : {}) } : undefined
  }
  /** With no media config: a temp that is never used (no ffmpeg means no probe stages a file). */
  private static readonly NO_TEMP: MediaTemp = {
    dir: '', create: () => Promise.reject(new Error('media off')), adopt: () => { throw new Error('media off') }, reserve: () => { throw new Error('media off') },
    live: () => 0, bytes: () => 0, remove: () => Promise.resolve(),
  }
  private mediaTableFor(doc: OpenDoc): DocMediaTable {
    const m = this.mediaConfig
    return createDocMediaTable(doc, m?.info, m?.temp ?? DocumentHost.NO_TEMP, this.mediaSeams, m?.runner ?? (m?.info && localProbeRunner(m.info)))
  }
  /** The document's probe table (tests, stats). */
  mediaTable(docKey: string): DocMediaTable {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    return doc.media
  }
  /** Probes run, summed over open documents (worker stats). */
  get mediaProbes(): number { let n = 0; for (const d of this.docs.values()) n += d.media.probes; return n + this.droppedProbes }
  private droppedProbes = 0

  /**
   * Piece 5 (D21, D22): unit `unit`'s audio uses and their probed entries. Loads the slide (and probes in k order up
   * to its highest k) first. A hidden slide lists nothing and probes nothing.
   */
  async mediaRefs(docKey: string, unit: number): Promise<MediaUnitRefs> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!Number.isInteger(unit) || unit < 0 || unit >= doc.units.length) throw new RenderRefusal('unit-range')
    const hidden = doc.hidden.includes(unit)
    if (!doc.show || hidden) return { unit, hidden, refs: [] }
    await this.loadShow(doc, unit)
    const show = doc.show
    const slide = show.metas.get(unit)!.slide
    const table = doc.media
    const kOf = (partPath: string) => { const e = table.peek(partPath); if (!e) throw new Error('media: part not probed'); return e.k }
    const uses = slideAudioUses(slide, unit, show.mediaCtx.get(unit)!, kOf, table.partCount)
    const refs: MediaUnitRefs['refs'] = uses.records.map((r) => ({
      k: r.ref.k, partPath: r.partPath ?? '', ...(r.spid !== undefined ? { spid: r.spid } : {}), hiddenShape: r.hiddenShape,
      // A use the slide refuses (`slideAudioUses`: no part, a link, a chain on a video end) is refused here too.
      entry: r.partPath !== undefined && r.refused === undefined ? asAudioUse(table.peek(r.partPath)!) : { k: r.ref.k, partPath: '', refused: r.refused ?? 'media-unreadable' },
      identity: mediaIdentity(r.ref), placeholder: r.placeholder,
    }))
    return { unit, hidden, refs }
  }

  /**
   * Piece 5 (D21): writes the projected part to a fresh 0600 file directly in the media dir (`dir` must be the one
   * configured), so part bytes never cross the worker boundary. Only a media part of the projection is written.
   */
  async mediaPart(docKey: string, partPath: string, dir: string, name?: string): Promise<{ path: string; bytes: number; sourceHash: string }> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    const m = this.mediaConfig
    if (!m || typeof dir !== 'string' || dir !== m.dir) throw new RenderRefusal('media-part', 'not the media directory')
    if (typeof partPath !== 'string' || !(await doc.media.parts()).includes(partPath)) throw new RenderRefusal('media-part', 'not a media part of the projection')
    try {
      const r = await writePart(doc, partPath, dir, name)
      await this.afterPartForTests?.()
      return r
    } catch (e) {
      if (e instanceof MediaPartRefused) throw new RenderRefusal('media-part', e.reason)
      throw e
    }
  }

  /** TEST-ONLY (set by the worker from a `testHooks` init): runs once `mediaPart` has written its file. */
  afterPartForTests?: () => Promise<void>
  /** TEST-ONLY seams for the show-plan build (a forced budget or composite mismatch, a ledger); empty in production. */
  showOptions: ShowBuildOptions = {}
  /** The seams plus the job's admitted grant (PR #78 round 13), when the pool sent one. */
  private showJobOptions(grant?: number): ShowBuildOptions { return grant === undefined ? this.showOptions : { ...this.showOptions, grant } }

  /**
   * Piece 4: the show-plan bundle of one shown slide (`buildShowPlan`). The slide's metadata and the deck's
   * images are read here (async), then the build runs synchronously on them. A hidden or out-of-range unit
   * has no plan (`unit-range`).
   */
  async showPlan(docKey: string, job: ShowPlanJob, grant?: number): Promise<BuiltShowPlan | ShowRefusal> {
    const doc = this.showDoc(docKey, job)
    await this.loadShow(doc, job.unitIndex)
    this.pages++
    return buildShowPlan(doc, job, this.showJobOptions(grant))
  }

  /** Piece 4: the keyframe track of one `composited` step (`renderTrack`); `[]` for a step that plays from layers. */
  async showTrack(docKey: string, job: ShowPlanJob & { step: number }, grant?: number): Promise<Array<{ assetId: number; w: number; h: number; png: Uint8Array }>> {
    const doc = this.showDoc(docKey, job)
    if (!Number.isInteger(job.step) || job.step < 0) throw new RenderRefusal('bad-job')
    await this.loadShow(doc, job.unitIndex)
    this.pages++
    return renderTrack(doc, job, this.showJobOptions(grant))
  }

  /**
   * Piece 4 Task 7: one Morph pair (`buildMorphPlan`): `from` and `to` shown slides of this deck, `to` entered by
   * Morph, `fromStep` within the from-slide's steps (checked here, on the slide as prepared: `unit-range`).
   */
  async showMorph(docKey: string, job: MorphJob, grant?: number): Promise<BuiltMorph> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!doc.show) throw new RenderRefusal('unsupported-format', 'no show for this document')
    if (!job || typeof job !== 'object' || (job.motion !== 'full' && job.motion !== 'reduced') || !job.mark || typeof job.mark !== 'object' || (job.strictMarks !== undefined && typeof job.strictMarks !== 'boolean')) throw new RenderRefusal('bad-job')
    if (job.scale !== 1 && job.scale !== 2) throw new RenderRefusal('scale')
    const shown = (u: unknown) => Number.isInteger(u) && (u as number) >= 0 && (u as number) < doc.units.length && !doc.hidden.includes(u as number)
    if (!shown(job.from) || !shown(job.to) || job.from === job.to || !doc.morph.includes(job.to) || !Number.isInteger(job.fromStep) || job.fromStep < 0) throw new RenderRefusal('unit-range')
    await this.loadShow(doc, job.from)
    await this.loadShow(doc, job.to)
    this.pages++
    try { return buildMorphPlan(doc, job, this.showJobOptions(grant)) } catch (e) { if (e instanceof MorphInvalid) throw new RenderRefusal('unit-range', e.message); throw e }
  }

  /** TEST-ONLY (Task 7 fix round 1): the Morph pair composed from the server's own rasters, marks off (`morphReference`). */
  async showMorphReference(docKey: string, job: MorphJob, progresses: readonly number[]) {
    const doc = this.docs.get(docKey)
    if (!doc?.show) throw new RenderRefusal('unknown-doc')
    await this.loadShow(doc, job.from)
    await this.loadShow(doc, job.to)
    return morphReference(doc, job, progresses)
  }

  /** Show-plan cache counters, summed over open documents (L1: preparation; L2: marked bundles). */
  showCacheStats() {
    const out = { l1: { hits: 0, misses: 0 }, l2: { hits: 0, misses: 0 } }
    for (const d of this.docs.values()) {
      if (!d.show) continue
      out.l1.hits += d.show.stats.l1.hits; out.l1.misses += d.show.stats.l1.misses
      out.l2.hits += d.show.stats.l2.hits; out.l2.misses += d.show.stats.l2.misses
    }
    return out
  }

  private showDoc(docKey: string, job: ShowPlanJob): OpenDoc {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!doc.show) throw new RenderRefusal('unsupported-format', 'no show for this document')
    if (!job || typeof job !== 'object' || (job.motion !== 'full' && job.motion !== 'reduced') || typeof job.strictMarks !== 'boolean' || !job.mark || typeof job.mark !== 'object') throw new RenderRefusal('bad-job')
    if (!Number.isInteger(job.unitIndex) || job.unitIndex < 0 || job.unitIndex >= doc.units.length || doc.hidden.includes(job.unitIndex)) throw new RenderRefusal('unit-range')
    if (job.scale !== 1 && job.scale !== 2) throw new RenderRefusal('scale')
    if (job.kBasis !== undefined && !(job.kBasis instanceof Uint8Array && job.kBasis.length === 32)) throw new RenderRefusal('bad-job')
    if (job.media !== undefined && typeof job.media !== 'boolean') throw new RenderRefusal('bad-job')
    return doc
  }

  /** The slide's metadata and the deck's images (async reads), then the media probes up to the slide's highest k
   * (D21), so the synchronous build reads `doc.media.peek`. */
  private async loadShow(doc: OpenDoc, unit: number): Promise<void> {
    const show = doc.show!
    if (!show.metas.has(unit)) {
      const h = await show.acquire(unit)
      // The handle is held for the document's life: the slide's text-build working copy is charged to it.
      show.metas.set(unit, { slide: h.slide, charge: (bytes, tag) => h.attach(bytes, tag) })
    }
    if (!show.images) {
      const backend = this.backend
      show.images = await Promise.all(show.doc.images.map((a) => decodeImageAsset(a as never, (b, mime) => backend.decodeImage(b, mime) as Promise<CanvasImageSource>).catch(() => undefined)))
    }
    await doc.media.load(unit)
    // Piece 5: the slide's hidden shapes and external media links, which the synchronous build cannot read.
    if (!show.mediaCtx.has(unit)) show.mediaCtx.set(unit, hasAudioUse(show.metas.get(unit)!.slide) ? await slideMediaContext(show, unit, doc.contentPolicyId) : NO_MEDIA_CONTEXT)
  }

  /** Stage 3c notes: unit i's notes slide read from the original package
   * (uncached afterwards: C4, every request reads again), planned or drawn.
   * A deck whose slide list does not line up with its units has no notes. */
  async notes(docKey: string, job: NotesJob): Promise<NotesPlan | NotesRaster> {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!doc.notes) throw new RenderRefusal('unsupported-format', 'no notes for this document')
    if (!Number.isInteger(job.unitIndex) || job.unitIndex < 0 || job.unitIndex >= doc.notes.slideParts.length) throw new RenderRefusal('unit-range')
    if (job.scale !== 1 && job.scale !== 2) throw new RenderRefusal('scale')
    const scale = job.scale
    let notes
    try { notes = await readSlideNotes(doc.notes.pkg, doc.notes.slideParts[job.unitIndex]!, job.unitIndex) } finally { doc.notes.pkg.dropCache() }
    if (!job.draw) return { kind: 'notes-plan', ...planNotes(notes, scale, this.backend) }
    const d = drawNotes(notes.paragraphs, scale, this.backend, notes.truncated)
    if (!unitSurfaceAllowed(d.w, d.h)) throw new RenderRefusal('oversize')
    const mark = markOptions(job.mark)
    if (mark) paintMark(d.ctx, { widthPx: NOTES_WIDTH_CSS, heightPx: d.h / scale }, mark)
    this.pages++
    const png = criticalChunksOnly(new Uint8Array((d.canvas as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png')))
    return { kind: 'raster', w: d.w, h: d.h, png, truncated: d.truncated, chars: d.chars }
  }

  /**
   * Stage 3a text jobs. `prepare` records the unit's spans (paint order), runs
   * one visual pass at scale 1 on a real canvas of the unit's size through the
   * occlusion observer (no visible mark), joins, sanitises and caches the
   * served unit. Other ops answer from the cache or `{ missing }`.
   * Notes are never indexed: units are the document's paintables only.
   */
  /** `grant`: the admission bytes the pool reserved for a prepare; a prepare
   * whose workspace exceeds it is refused before any canvas is created. */
  text(docKey: string, job: TextJob, grant = Number.POSITIVE_INFINITY): TextReply {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!job || typeof job !== 'object') throw new RenderRefusal('bad-job')
    if (job.op === 'prepare') {
      const u = job.unit
      if (!Number.isInteger(u) || u < 0 || u >= doc.units.length) throw new RenderRefusal('unit-range')
      const unit = this.served.get(docKey, u) ?? this.prepare(doc, u, grant)
      this.served.put(docKey, u, unit)
      return { op: 'prepare', served: unit.served, withheld: unit.withheld, bytes: unit.bytes, causes: { ...unit.causes, occluded: { ...unit.causes.occluded } } }
    }
    if (!['search', 'select', 'copy', 'a11y'].includes(job.op)) throw new RenderRefusal('bad-job')
    try {
      return answerText(job, doc.units.length, (u) => this.served.get(docKey, u))
    } catch (err) {
      if (err instanceof TextRefusal) throw new RenderRefusal(err.code)
      throw err
    }
  }

  /** Test/ops: the cached served unit, without touching LRU order. */
  servedUnit(docKey: string, unit: number): ServedUnit | undefined { return this.served.peek(docKey, unit) }
  servedBytes(docKey: string): number { return this.served.bytesOf(docKey) }
  /** The served-index cache bound per document (R6 default). */
  setServedCacheBytes(limit: number): void { this.served.limit = limit }

  private prepare(doc: OpenDoc, index: number, grant: number): ServedUnit {
    const r = doc.raw[index]!
    const w = sizeCeil(Math.max(1, r.widthPx)), h = sizeCeil(Math.max(1, r.heightPx))
    if (!unitSurfaceAllowed(w, h)) throw new RenderRefusal('oversize')
    if (!(prepareWorkspace(w, h) <= grant)) throw new RenderRefusal('grant', `prepare needs ${prepareWorkspace(w, h)} bytes, granted ${grant}`)
    const paintable = this.paintablesOf(doc)[index]!
    // Bounded text work: running caps on recorded spans and on the visual
    // pass's observer records, both inside TEXT_PREPARE_ALLOWANCE.
    let spans
    try {
      spans = recordSpans((ctx) => paintable.paint(ctx), this.backend.createCanvas(4, 4).getContext('2d') as unknown as CanvasRenderingContext2D, TEXT_RECORD_BUDGET)
    } catch (err) {
      if (err instanceof TextBudgetExceeded) throw new RenderRefusal('grant', err.message)
      throw err
    }
    this.pages++
    let pass
    try {
      pass = this.paintPass(paintable, index, w, h, 1, [], OBSERVER_BUDGET)
    } catch (err) {
      if (err instanceof ObserverBudgetExceeded) throw new RenderRefusal('grant', err.message)
      throw err
    }
    if (pass.exceeded()) throw new RenderRefusal('grant', 'occlusion observer record budget exceeded')
    return buildServedUnit(spans, pass.runDetails(), 1)
  }

  /** The observed paint of one unit (shared by the raster render and the text visual pass). */
  private paintPass(paintable: PaintableArray[number], index: number, w: number, h: number, scale: number, observers: readonly PaintObserver[], maxObserverBytes?: number) {
    const canvas = this.backend.createCanvas(w, h)
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    ctx.scale(scale, scale)
    const { painted, evidence, runDetails, exceeded } = observe(ctx, w, h, scale, observers, maxObserverBytes)
    for (const o of observers) o.beforeUnit(painted, index)
    paintable.paint(painted)
    for (const o of observers) o.afterUnit(painted, index)
    return { canvas, ctx, evidence, runDetails, exceeded }
  }

  private raster(paintable: PaintableArray[number], index: number, w: number, h: number, scale: number, page: { widthPx: number; heightPx: number }, mark: WatermarkOptions | undefined, observers: readonly PaintObserver[]): RenderedUnit & { kind: 'raster' } {
    const { canvas, ctx, evidence } = this.paintPass(paintable, index, w, h, scale, observers)
    if (mark) paintMark(ctx, page, mark)
    const png = criticalChunksOnly(new Uint8Array((canvas as unknown as { toBuffer(m: 'image/png'): Buffer }).toBuffer('image/png')))
    return { kind: 'raster', w, h, png, evidence: evidence() }
  }

  private recordCommands(paintable: PaintableArray[number], index: number, w: number, h: number, scale: number, page: { widthPx: number; heightPx: number }, mark: WatermarkOptions | undefined, observers: readonly PaintObserver[]) {
    const backing = this.backend.createCanvas(w, h).getContext('2d') as unknown as CanvasRenderingContext2D
    const rec = createRecorder(backing, { scale, readImage: (s) => readImage(s, this.backend), partialAlphaSnapshots: 'allow', patternSources: 'snapshot' })
    // Font pinning sits OUTSIDE the recorder so recorded fonts carry pack names.
    const ctx = pinFonts(rec.ctx, loadPack())
    ctx.scale(scale, scale)
    const { painted, evidence } = observe(ctx, w, h, scale, observers)
    for (const o of observers) o.beforeUnit(painted, index)
    paintable.paint(painted)
    for (const o of observers) o.afterUnit(painted, index)
    if (mark) paintMark(ctx, page, mark)
    const out = rec.finish()
    if (out.discriminator === 3) return { ok: false as const, reason: out.reason }
    return { ok: true as const, ops: out.ops, opCount: out.opCount, assets: out.assets, evidence: evidence() }
  }
}

const MEDIA_REL_TYPE = /\/(audio|video|media)$/
/**
 * Piece 5: what the build needs of slide `unit`'s package: its hidden shapes (under `full`; the projection drops them
 * otherwise, as `mediaRefs` reads them), its external media relationships and the shapes they play on.
 */
async function slideMediaContext(show: ShowDoc, unit: number, contentPolicyId: string): Promise<SlideMediaContext> {
  const part = show.handle.slideParts[unit]!
  let xml: string | undefined
  const slideXml = async () => (xml ??= (await show.handle.pkg.entry(part)?.async('string')) ?? '')
  const rels = await readRelationships(show.handle, part).catch(() => [])
  const linkedRIds = new Set(rels.filter((r) => r.external && MEDIA_REL_TYPE.test(r.type)).map((r) => r.id))
  const linkedSpids = new Set<string>()
  if (linkedRIds.size) for (const [id, spids] of mediaShapes(parseXml(await slideXml()))) if (linkedRIds.has(id)) for (const sp of spids) linkedSpids.add(sp)
  const hiddenSpids = contentPolicyId.startsWith('full') ? new Set<string>(hiddenShapeIds(readTree(await slideXml()))) : new Set<string>()
  return { hiddenSpids, linkedRIds, linkedSpids }
}
