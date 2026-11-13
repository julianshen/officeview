/**
 * §8.3 `RemoteOfficeDoc`: the provider a remote `<OfficeDoc>` paints from.
 * Replay is 1:1 at a scale bucket; zoom re-requests. Text services (Stage 3a)
 * are server-answered: the client never indexes text. A service the manifest
 * does not offer rejects locally with `text-unavailable`, announced once.
 */
import { canvasBackend, type CanvasBackend, type CanvasLike } from '../../core/canvas-backend'
import type { HighlightRect } from '../../core/overlay'
import type { Manifest } from '../manifest'
import { replayUnit } from '../replay'
import { StreamSessionError } from './errors'
import type { PrintedPage } from './print-wire'
import type { NotesLease, PrintOptions, StreamNotice, UnitLease, UnitPayload } from './session'
import type { RemoteRange, TextMessage, TextRequest } from './text-wire'

export type { RemoteRange } from './text-wire'
export type { PrintedPage } from './print-wire'
export type { PrintOptions } from './session'

export interface SessionLike {
  readonly manifest: Manifest
  unit(index: number, bucket: number, signal?: AbortSignal): Promise<UnitLease>
  /** One text request; resolves on the op's terminal message (see StreamSession.text). */
  text(request: TextRequest, on: (m: TextMessage) => void, signal?: AbortSignal): Promise<void>
  /** One checked print job (see StreamSession.print); absent ⇒ print is unavailable. */
  print?(from: number, to: number, opts: PrintOptions): Promise<PrintedPage[]>
  /** One checked notes image (see StreamSession.notes); absent ⇒ notes are unavailable. */
  notes?(unit: number, scale: 1 | 2, signal?: AbortSignal): Promise<NotesLease>
  close(): void
  readonly ended?: Promise<StreamSessionError | undefined>
  /** Calls `cb` after a resume has replaced `manifest`; returns the unsubscribe. */
  onManifestChange?(cb: () => void): () => void
  /** True once the session has ended (a synchronous barrier; `ended` settles later). */
  isEnded?(): boolean
}

export interface RemoteOfficeDoc {
  readonly kind: 'remote'
  readonly manifest: Manifest
  /** CSS-pixel unit sizes at scale 1. */
  units(): Promise<Array<{ widthPx: number; heightPx: number }>>
  /** Paints the unit at the bucket for `scale × dpr`; the canvas is resized to the bucket surface. */
  paint(unit: number, scale: 'fit' | number, dpr: number, canvas: CanvasLike, signal?: AbortSignal): Promise<void>
  /**
   * Server search. Each `onHits` call carries EVERY hit found so far (replace,
   * don't append); after a resume the search is re-issued and `onHits([])`
   * announces that the earlier hits are void. Rejects on abort, refusal or end.
   */
  search(query: string, o?: { caseSensitive?: boolean; signal?: AbortSignal; onHits?: (h: RemoteSearchHit[]) => void }): Promise<{ total: number; truncated: boolean }>
  /** The rects of a range. A newer select rejects a pending older one with
   * code 'closed' (a drag issuing many selects keeps only the newest answer). */
  select(range: RemoteRange, signal?: AbortSignal): Promise<{ units: RemoteSearchHit[]; paragraphs: number }>
  /** The range's text; rejects StreamSessionError with `limit`/`requested` on a cap. */
  copy(range: RemoteRange): Promise<string>
  /** Accessible paragraphs of a unit (one paragraph in 'focused' mode). */
  a11y(unit: number, paragraph?: number): Promise<Array<{ rect: HighlightRect; text?: string }>>
  /** Server-drawn print pages for units `from..to` (inclusive), every page
   * checked against the job (see StreamSession.print for the rejections). */
  print(range: { from: number; to: number }, opts?: PrintOptions): Promise<PrintedPage[]>
  /** Paints the presenter's server-drawn, marked notes image of `unit` at the
   * notes scale for `dpr` (1 or 2); the canvas is resized to the image. Rejects
   * `policy-refused` without sending when `can.notes` is off; otherwise as
   * StreamSession.notes. Nothing is painted after an abort, a close or a refusal. */
  paintNotes(unit: number, dpr: number, canvas: CanvasLike, signal?: AbortSignal): Promise<void>
  /** Calls `cb` after a resume has replaced `manifest` (R3: `can.copy`, `caps` and
   * `can.textA11y` may change); returns the unsubscribe. */
  onManifestChange(cb: () => void): () => void
  close(): void
  /** Settles once when the session ends: its typed error, or undefined after close(). */
  readonly ended: Promise<StreamSessionError | undefined>
}

/** Rects of one unit, in page px at scale 1 relative to the unit. */
export interface RemoteSearchHit { unit: number; rects: HighlightRect[] }

/** Flat `[x, y, w, h, …]` (wire) → HighlightRects. */
const toRects = (flat: readonly number[]): HighlightRect[] => {
  const out: HighlightRect[] = []
  for (let i = 0; i + 3 < flat.length; i += 4) out.push({ x: flat[i]!, y: flat[i + 1]!, width: flat[i + 2]!, height: flat[i + 3]! })
  return out
}

/** Smallest bucket ≥ scale × dpr, else the largest (§8.3). */
export function bucketFor(scale: number, dpr: number, buckets: readonly number[]): number {
  const want = scale * dpr, sorted = [...buckets].sort((a, b) => a - b)
  return sorted.find((b) => b >= want - 1e-9) ?? sorted[sorted.length - 1]!
}

export function createRemoteDoc(session: SessionLike, o: { backend?: CanvasBackend; onNotice?: (n: StreamNotice) => void } = {}): RemoteOfficeDoc {
  let closed = false, textNoticed = false
  const backend = () => o.backend ?? canvasBackend()
  const textUnavailable = (): Promise<never> => {
    if (!textNoticed) { textNoticed = true; o.onNotice?.({ code: 'text-unavailable' }) }
    return Promise.reject(new StreamSessionError('text-unavailable', 'text unavailable in this viewing mode'))
  }
  const can = () => session.manifest.can
  /** Draws into an offscreen surface; the visible canvas is touched only on success. */
  async function render(index: number, bucket: number, payload: UnitPayload): Promise<CanvasLike | { image: CanvasImageSource; w: number; h: number }> {
    const b = backend()
    if (payload.kind === 'raster') return { image: await b.decodeImage(payload.png, 'image/png'), w: payload.w, h: payload.h }
    const dims = session.manifest.units.list[index]!
    const surface = b.createCanvas(dims.w * bucket, dims.h * bucket)
    await replayUnit(surface.getContext('2d')!, payload, b)
    return surface
  }
  const doc: RemoteOfficeDoc & { session: SessionLike } = {
    kind: 'remote',
    session,
    get manifest() { return session.manifest },
    async units() { return session.manifest.units.list.map((u) => ({ widthPx: u.w, heightPx: u.h })) },
    async paint(unit, scale, dpr, canvas, signal) {
      if (closed) throw new StreamSessionError('closed')
      const bucket = bucketFor(scale === 'fit' ? 1 : scale, dpr, session.manifest.recorder.scalesBuckets)
      let out: Awaited<ReturnType<typeof render>> | undefined
      // §7.6/§10: a unit that fails to replay is re-requested once, then reported.
      for (let attempt = 0; attempt < 2 && !out; attempt++) {
        const payload = await session.unit(unit, bucket, signal)
        // The admission lease covers decode/replay; release it once that work is done.
        try { out = await render(unit, bucket, payload) } catch { /* retry once */ } finally { payload.release() }
      }
      if (!out) { o.onNotice?.({ code: 'unit-failed', unit }); throw new StreamSessionError('unit-failed') }
      try {
        // Read synchronously: `ended.then` may not have run yet when the render resolves.
        if (closed || session.isEnded?.() || signal?.aborted) return
        const w = 'image' in out ? out.w : out.width, h = 'image' in out ? out.h : out.height
        canvas.width = w; canvas.height = h
        const ctx = canvas.getContext('2d')!
        ctx.setTransform(1, 0, 0, 1, 0, 0)
        ctx.clearRect(0, 0, w, h)
        ctx.drawImage(('image' in out ? out.image : out) as CanvasImageSource, 0, 0)
      } finally {
        // A decoded raster's ImageBitmap frees its native surface now, not at GC.
        if ('image' in out) (out.image as { close?: () => void }).close?.()
      }
    },
    async search(query, so = {}) {
      if (!can().search) return textUnavailable()
      let hits: RemoteSearchHit[] = [], done = { total: 0, truncated: false }
      await session.text({ type: 'search', query, ...(so.caseSensitive !== undefined ? { caseSensitive: so.caseSensitive } : {}) }, (m) => {
        if (m.type === 'restart') { if (hits.length) { hits = []; so.onHits?.(hits) } }
        else if (m.type === 'search-result') { hits = [...hits, ...m.matches.map((x) => ({ unit: x.unit, rects: toRects(x.rects) }))]; so.onHits?.(hits) }
        else if (m.type === 'search-done') done = { total: m.total, truncated: m.truncated }
      }, so.signal)
      return done
    },
    async select(range, signal) {
      if (!can().copy) return textUnavailable()
      // The same unit may repeat across chunks: its rects concatenate.
      let units = new Map<number, number[]>(), paragraphs = 0
      await session.text({ type: 'select', range }, (m) => {
        if (m.type === 'restart') units = new Map()
        else if (m.type === 'select-result') {
          paragraphs = m.paragraphs
          for (const u of m.units) {
            const rects = units.get(u.unit) ?? []
            for (const n of u.rects) rects.push(n)
            units.set(u.unit, rects)
          }
        }
      }, signal)
      return { units: [...units].map(([unit, rects]) => ({ unit, rects: toRects(rects) })), paragraphs }
    },
    async copy(range) {
      if (!can().copy) return textUnavailable()
      const parts: string[] = []
      await session.text({ type: 'copy', range }, (m) => { if (m.type === 'copy-result') parts.push(m.text) })
      return parts.join('')
    },
    async a11y(unit, paragraph) {
      if (!can().textA11y) return textUnavailable()
      const out: Array<{ rect: HighlightRect; text?: string }> = []
      await session.text({ type: 'a11y', unit, ...(paragraph !== undefined ? { paragraph } : {}) }, (m) => {
        if (m.type === 'a11y-result') for (const p of m.paragraphs) out.push({ rect: toRects(p.rect)[0]!, ...(p.text !== undefined ? { text: p.text } : {}) })
      })
      return out
    },
    async print(range, po = {}) {
      if (closed) throw new StreamSessionError('closed')
      if (!session.print) throw new StreamSessionError('policy-refused', 'print unavailable')
      return session.print(range.from, range.to, po)
    },
    async paintNotes(unit, dpr, canvas, signal) {
      if (closed) throw new StreamSessionError('closed')
      if (!can().notes || !session.notes) throw new StreamSessionError('policy-refused', 'notes unavailable')
      // The server takes scale 1 or 2 only; the notes are 720 × scale device px wide.
      const scale = bucketFor(1, dpr, session.manifest.recorder.scalesBuckets.filter((b) => b === 1 || b === 2)) as 1 | 2 | undefined
      if (scale === undefined) throw new StreamSessionError('policy-refused', 'no notes scale')
      const lease = await session.notes(unit, scale, signal)
      let image: CanvasImageSource
      // The admission lease covers the decode; release it once that work is done.
      try { image = await backend().decodeImage(lease.png, 'image/png') } finally { lease.release() }
      try {
        // Read synchronously: `ended.then` may not have run yet when the decode resolves.
        if (closed || session.isEnded?.() || signal?.aborted || !can().notes) return
        canvas.width = lease.w; canvas.height = lease.h
        const ctx = canvas.getContext('2d')!
        ctx.setTransform(1, 0, 0, 1, 0, 0)
        ctx.clearRect(0, 0, lease.w, lease.h)
        ctx.drawImage(image, 0, 0)
      } finally {
        // An ImageBitmap's native surface is freed now, not at GC.
        (image as { close?: () => void }).close?.()
      }
    },
    onManifestChange(cb) { return session.onManifestChange?.(cb) ?? (() => {}) },
    close() { if (closed) return; closed = true; session.close() },
    ended: session.ended ?? new Promise(() => {}),
  }
  // A terminal end invalidates every pending paint: nothing commits pixels after it.
  void doc.ended.then((e) => { closed = true; if (e) o.onNotice?.({ code: e.code }) })
  return doc
}
