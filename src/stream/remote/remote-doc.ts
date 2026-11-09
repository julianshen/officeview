/**
 * §8.3 `RemoteOfficeDoc`: the provider a remote `<OfficeDoc>` paints from.
 * Replay is 1:1 at a scale bucket; zoom re-requests. Stage 1 has no semantic
 * services: search/copy reject and "text unavailable" is announced once.
 */
import { canvasBackend, type CanvasBackend, type CanvasLike } from '../../core/canvas-backend'
import type { Manifest } from '../manifest'
import { replayUnit } from '../replay'
import { StreamSessionError } from './errors'
import type { StreamNotice, UnitPayload } from './session'

export interface SessionLike {
  readonly manifest: Manifest
  unit(index: number, bucket: number, signal?: AbortSignal): Promise<UnitPayload>
  close(): void
}

export interface RemoteOfficeDoc {
  readonly kind: 'remote'
  readonly manifest: Manifest
  /** CSS-pixel unit sizes at scale 1. */
  units(): Promise<Array<{ widthPx: number; heightPx: number }>>
  /** Paints the unit at the bucket for `scale × dpr`; the canvas is resized to the bucket surface. */
  paint(unit: number, scale: 'fit' | number, dpr: number, canvas: CanvasLike, signal?: AbortSignal): Promise<void>
  search(query: string): Promise<never>
  copy(range: unknown): Promise<never>
  close(): void
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
    return Promise.reject(new StreamSessionError('policy-refused', 'text unavailable in this viewing mode'))
  }
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
        try { out = await render(unit, bucket, payload) } catch { /* retry once */ }
      }
      if (!out) { o.onNotice?.({ code: 'unit-failed', unit }); throw new StreamSessionError('unit-failed') }
      if (closed || signal?.aborted) return
      const w = 'image' in out ? out.w : out.width, h = 'image' in out ? out.h : out.height
      canvas.width = w; canvas.height = h
      const ctx = canvas.getContext('2d')!
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.clearRect(0, 0, w, h)
      ctx.drawImage(('image' in out ? out.image : out) as CanvasImageSource, 0, 0)
    },
    search: textUnavailable,
    copy: textUnavailable,
    close() { if (closed) return; closed = true; session.close() },
  }
  return doc
}
