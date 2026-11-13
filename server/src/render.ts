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
import { paintWatermark, setWatermarkMeasurer, type WatermarkOptions } from '../../src/core/watermark'
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
import { answerText, buildServedUnit, OBSERVER_BUDGET, prepareWorkspace, recordSpans, ServedCache, TEXT_RECORD_BUDGET, TextBudgetExceeded, TextRefusal, type ServedUnit, type TextJob, type TextReply } from './text/served'

export const MAX_UNIT_PX = 8192
export const SCALE_BUCKETS = [1, 2] as const
export const FIXED_MARK_TEXT = 'officeview · protected'
export type MarkSpec = { kind: 'fixed' } | { kind: 'viewer'; text: string } | { kind: 'off' }
export type Projection = 'final-raster' | 'commands'

/** Visible-mark styling [call]: mid-grey reads on light and dark pages alike. */
export function markOptions(mark: MarkSpec): WatermarkOptions | undefined {
  if (mark.kind === 'off') return undefined
  return { text: mark.kind === 'fixed' ? FIXED_MARK_TEXT : mark.text, color: '#808080', opacity: 0.18 }
}

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
export class RenderRefusal extends Error { constructor(readonly code: 'unsupported-format' | 'unit-range' | 'scale' | 'oversize' | 'unknown-doc' | 'bad-job' | 'grant', m?: string) { super(m ?? code); this.name = 'RenderRefusal' } }

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
  h.update(`locale=en-US\npolicy=${contentPolicyId}\n`)
  return h.digest('hex')
}

/** `notes`: a deck's ORIGINAL package (the projection drops notes slides) and unit i's slide part. */
interface OpenDoc { key: string; paintables: PaintableArray; units: UnitSize[]; hidden: number[]; raw: Array<{ widthPx: number; heightPx: number }>; fingerprint: string; contentPolicyId: string; diagnostics: PolicyDiagnostic[]; format: 'docx' | 'xlsx' | 'pptx'; notes?: { pkg: OfficePackage; slideParts: readonly string[] } }
/** A notes job (Stage 3c): `draw: false` plans only (size, truncation, characters drawn); `true` draws with the visible mark. */
export interface NotesJob { unitIndex: number; scale: number; mark: MarkSpec; draw: boolean }
export interface NotesPlan { kind: 'notes-plan'; w: number; h: number; truncated: boolean; chars: number }
export type NotesRaster = RenderedUnit & { kind: 'raster'; truncated: boolean; chars: number }
/** `hidden`: unit indices of hidden slides (pptx under `'full'` only; elsewhere they are not units). */
export interface OpenResult { docKey: string; units: UnitSize[]; hidden: number[]; layoutFingerprint: string; contentPolicyId: string; format: 'docx' | 'xlsx' | 'pptx'; diagnostics: PolicyDiagnostic[]; fontDiagnostics: number }

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

/** Paints the visible mark from a clean watermark module state.
 *
 * `src/core/watermark.ts` keeps a process-global text-width cache keyed on the
 * ROUNDED point size, but (with no measurer registered, as here) stores a
 * width computed from the UNROUNDED pixel size. Two pages whose sizes round to
 * the same pt (a 960-px deck and a 955-px sheet are both 50 pt) then share one
 * entry, and whichever painted first sets the tile step of the other: the
 * canonical a long-lived worker serves would depend on what it rendered
 * before, and a fresh trace process would refuse the reference (F1, Stage 2b).
 * `setWatermarkMeasurer(null)` is the module's only reset (measurer stays
 * null, cache cleared), so every mark is laid out from its own page alone. */
function paintMark(ctx: CanvasRenderingContext2D, page: { widthPx: number; heightPx: number }, mark: WatermarkOptions): void {
  setWatermarkMeasurer(null)
  paintWatermark(ctx, page, mark)
}

/** Hidden slides of a deck opened under `'full'`, read from the original
 * package (there the projection is the identity). A slide list that does not
 * line up with the units refuses the open: never guess which unit is hidden. */
export async function hiddenUnits(doc: PptxDocument, units: number): Promise<number[]> {
  const handle = pptxPackageHandle(doc)
  if (!handle || handle.slideParts.length !== units) throw new RenderRefusal('unsupported-format', 'slide list does not match the units')
  return hiddenSlideUnits(handle)
}

export class DocumentHost {
  readonly backend: CanvasBackend
  private docs = new Map<string, OpenDoc>()
  private served: ServedCache
  parses = 0
  pages = 0
  /** Reads a 'full' deck's hidden units after its paintables exist (TEST-ONLY override: fault injection). */
  hiddenReader: (doc: PptxDocument, paintables: PaintableArray) => Promise<number[]> = (doc, paintables) => hiddenUnits(doc, paintables.length)
  constructor(backend = serverBackend(), servedCacheBytes?: number) { this.backend = backend; installCanvasBackend(backend); this.served = new ServedCache(servedCacheBytes) }

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
      doc = { key, paintables, raw, hidden, ...(notes ? { notes } : {}), units: raw.map((r) => ({ w: sizeCeil(Math.max(1, r.widthPx)), h: sizeCeil(Math.max(1, r.heightPx)) })), fingerprint, contentPolicyId: projected.contentPolicyId, diagnostics: projected.diagnostics, format: admitted.format }
      this.docs.set(key, doc)
    }
    return { docKey: key, units: doc.units, hidden: doc.hidden, layoutFingerprint: doc.fingerprint, contentPolicyId: doc.contentPolicyId, format: doc.format, diagnostics: doc.diagnostics, fontDiagnostics: doc.paintables.fontDiagnostics?.length ?? 0 }
  }

  has(docKey: string) { return this.docs.has(docKey) }
  drop(docKey: string) { const d = this.docs.get(docKey); if (d) { d.paintables.dispose(); this.docs.delete(docKey) } this.served.drop(docKey) }
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
    let paragraphs
    try { paragraphs = await readSlideNotes(doc.notes.pkg, doc.notes.slideParts[job.unitIndex]!, job.unitIndex) } finally { doc.notes.pkg.dropCache() }
    if (!job.draw) return { kind: 'notes-plan', ...planNotes(paragraphs, scale, this.backend) }
    const d = drawNotes(paragraphs, scale, this.backend)
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
