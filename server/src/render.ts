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
import { FIXED_MARK_TEXT, markOptions, paintMark, type MarkSpec } from './mark/visible'
import { detectAndParse } from '../../src/components/loadOfficeSource'
import { getPaintables, type PaintableArray } from '../../src/render/paint'
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
import { asAudioUse, createDocMediaTable, MediaPartRefused, writePart, type DocMediaEntry, type DocMediaTable, type MediaTableSeams } from './media/probe'
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
export class RenderRefusal extends Error { constructor(readonly code: 'unsupported-format' | 'unit-range' | 'scale' | 'oversize' | 'unknown-doc' | 'bad-job' | 'grant' | 'media-part', m?: string) { super(m ?? code); this.name = 'RenderRefusal' } }

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
export interface OpenDoc { key: string; docRef: string; show?: ShowDoc; /** Piece 5 (D21): the probe table, over the projected package. */ media: DocMediaTable; anim: number[]; morph: number[]; /** Piece 5 (D19): the shown units whose plan ships an audio record or a transition sound. */ mediaUnits: number[]; paintables: PaintableArray; units: UnitSize[]; hidden: number[]; raw: Array<{ widthPx: number; heightPx: number }>; fingerprint: string; contentPolicyId: string; diagnostics: PolicyDiagnostic[]; format: 'docx' | 'xlsx' | 'pptx'; notes?: { pkg: OfficePackage; slideParts: readonly string[] } }
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
export interface OpenResult { docKey: string; units: UnitSize[]; hidden: number[]; anim: number[]; morph: number[]; media: number[]; layoutFingerprint: string; contentPolicyId: string; format: 'docx' | 'xlsx' | 'pptx'; diagnostics: PolicyDiagnostic[]; fontDiagnostics: number }

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

export class DocumentHost {
  readonly backend: CanvasBackend
  private docs = new Map<string, OpenDoc>()
  private served: ServedCache
  parses = 0
  pages = 0
  /** Reads a 'full' deck's hidden units after its paintables exist (TEST-ONLY override: fault injection). */
  hiddenReader: (doc: PptxDocument, paintables: PaintableArray) => Promise<number[]> = (doc, paintables) => hiddenUnits(doc, paintables.length)
  constructor(backend = serverBackend(), servedCacheBytes?: number) { this.backend = backend; installCanvasBackend(gateBackend(backend)); this.served = new ServedCache(servedCacheBytes) }

  async open(docRef: string, bytes: Uint8Array, contentPolicy: 'visible-only' | 'full'): Promise<OpenResult> {
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
      const paintables = await getPaintables(parsed, { registerFont })
      this.parses++
      const raw = paintables.map((p) => ({ widthPx: p.spec.widthPx, heightPx: p.spec.heightPx }))
      let hidden: number[] = []
      // The lease is ours until the doc is cached: a refusal here must release it.
      if (contentPolicy === 'full' && admitted.format === 'pptx') {
        try { hidden = await this.hiddenReader(parsed as PptxDocument, paintables) } catch (e) { paintables.dispose(); throw e }
      }
      // Projected slide parts keep their original paths; a list that does not line up with the units serves no notes.
      const handle = admitted.format === 'pptx' ? pptxPackageHandle(parsed as PptxDocument) : undefined
      const notes = handle && handle.slideParts.length === paintables.length ? { pkg: original, slideParts: [...handle.slideParts] } : undefined
      // Projection read every original entry through the cache: the retained package keeps the archive, not those reads.
      original.dropCache()
      const units = raw.map((r) => ({ w: sizeCeil(Math.max(1, r.widthPx)), h: sizeCeil(Math.max(1, r.heightPx)) }))
      let show: ShowDoc | undefined
      const anim: number[] = [], morph: number[] = [], mediaUnits: number[] = []
      if (handle && handle.slideParts.length === paintables.length) {
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
      const opened = { key, docRef, paintables, raw, hidden, anim, morph, mediaUnits, ...(notes ? { notes } : {}), ...(show ? { show } : {}), units, fingerprint, contentPolicyId: projected.contentPolicyId, diagnostics: projected.diagnostics, format: admitted.format } as Omit<OpenDoc, 'media'> as OpenDoc
      opened.media = this.mediaTableFor(opened)
      if (show) show.media = opened.media
      doc = opened
      this.docs.set(key, doc)
    }
    return { docKey: key, units: doc.units, hidden: doc.hidden, anim: [...doc.anim], morph: [...doc.morph], media: [...doc.mediaUnits], layoutFingerprint: doc.fingerprint, contentPolicyId: doc.contentPolicyId, format: doc.format, diagnostics: doc.diagnostics, fontDiagnostics: doc.paintables.fontDiagnostics?.length ?? 0 }
  }

  has(docKey: string) { return this.docs.has(docKey) }
  drop(docKey: string) { const d = this.docs.get(docKey); if (d) { this.droppedProbes += d.media.probes; d.paintables.dispose(); d.show?.store.close(); this.docs.delete(docKey) } this.served.drop(docKey) }
  get docCount() { return this.docs.size }

  render(docKey: string, job: { unitIndex: number; scale: number; projection: Projection; mark: MarkSpec; missingFaces?: readonly string[] }, observers: readonly PaintObserver[] = []): RenderedUnit {
    const doc = this.docs.get(docKey)
    if (!doc) throw new RenderRefusal('unknown-doc')
    if (!Number.isInteger(job.unitIndex) || job.unitIndex < 0 || job.unitIndex >= doc.paintables.length) throw new RenderRefusal('unit-range')
    if (!(SCALE_BUCKETS as readonly number[]).includes(job.scale)) throw new RenderRefusal('scale')
    const r = doc.raw[job.unitIndex]!, paintable = doc.paintables[job.unitIndex]!
    const w = sizeCeil(Math.max(1, r.widthPx * job.scale)), h = sizeCeil(Math.max(1, r.heightPx * job.scale))
    if (!unitSurfaceAllowed(w, h)) throw new RenderRefusal('oversize')
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
  private mediaConfig?: { info: FfmpegInfo; dir: string; temp: MediaTemp }
  /** TEST-ONLY seams for the probe table (a clock, a held probe, a run observer); empty in production. */
  mediaSeams: MediaTableSeams = {}
  /** Sets ffmpeg and the media dir (a view over the pool's directory: a worker never claims one). Documents opened
   * afterwards get a table over them; with no config every entry is `media-unavailable`. */
  configureMedia(media: { info: FfmpegInfo; dir: string; prefix?: string } | undefined): void {
    this.mediaConfig = media ? { info: media.info, dir: media.dir, temp: attachMediaTemp(media.dir, media.prefix) } : undefined
  }
  /** With no media config: a temp that is never used (no ffmpeg means no probe stages a file). */
  private static readonly NO_TEMP: MediaTemp = {
    dir: '', create: () => Promise.reject(new Error('media off')), adopt: () => { throw new Error('media off') }, reserve: () => { throw new Error('media off') },
    live: () => 0, bytes: () => 0, remove: () => Promise.resolve(),
  }
  private mediaTableFor(doc: OpenDoc): DocMediaTable {
    const m = this.mediaConfig
    return createDocMediaTable(doc, m?.info, m?.temp ?? DocumentHost.NO_TEMP, this.mediaSeams)
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
      if (!Number.isInteger(u) || u < 0 || u >= doc.paintables.length) throw new RenderRefusal('unit-range')
      const unit = this.served.get(docKey, u) ?? this.prepare(doc, u, grant)
      this.served.put(docKey, u, unit)
      return { op: 'prepare', served: unit.served, withheld: unit.withheld, bytes: unit.bytes, causes: { ...unit.causes, occluded: { ...unit.causes.occluded } } }
    }
    if (!['search', 'select', 'copy', 'a11y'].includes(job.op)) throw new RenderRefusal('bad-job')
    try {
      return answerText(job, doc.paintables.length, (u) => this.served.get(docKey, u))
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
    const r = doc.raw[index]!, paintable = doc.paintables[index]!
    const w = sizeCeil(Math.max(1, r.widthPx)), h = sizeCeil(Math.max(1, r.heightPx))
    if (!unitSurfaceAllowed(w, h)) throw new RenderRefusal('oversize')
    if (!(prepareWorkspace(w, h) <= grant)) throw new RenderRefusal('grant', `prepare needs ${prepareWorkspace(w, h)} bytes, granted ${grant}`)
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
