/**
 * Worker render-to-bitmap protocol (main-safe: no worker globals touched here).
 *
 * A worker parses source bytes, lays out and paints one page/sheet/slide onto
 * an OffscreenCanvas and transfers the ImageBitmap back; the main thread only
 * presents it. Deterministic: same bytes + same browser build = same bitmap
 * (PNG encoding is a pure function of pixels).
 */
import type { PaintableArray, PaintOptions } from '../render/paint'

export type WorkerRenderKind = 'render-page' | 'render-sheet' | 'render-slide'
export type WorkerRenderFormat = 'docx' | 'xlsx' | 'pptx'
export interface WorkerRenderRequest {
  id: number
  kind: WorkerRenderKind
  format: WorkerRenderFormat
  /** Transferred (zero-copy) on postMessage; read here as a view, never copied. */
  source: ArrayBuffer
  index: number
  options?: PaintOptions
}
export type WorkerRenderResponse =
  | { id: number; ok: true; bitmap: ImageBitmap; width: number; height: number }
  | { id: number; ok: false; error: string }
/** Minimal 2D canvas surface (OffscreenCanvas in production, node-canvas in tests). */
export interface WorkerCanvas {
  getContext(kind: '2d'): CanvasRenderingContext2D | null
  width: number
  height: number
}
export interface WorkerRenderEnv {
  createCanvas(width: number, height: number): Promise<WorkerCanvas>
  toBitmap(canvas: WorkerCanvas): Promise<ImageBitmap>
  parse(format: WorkerRenderFormat, source: Uint8Array): Promise<unknown>
  paintables(doc: never, options?: PaintOptions): Promise<PaintableArray>
}
const KINDS: ReadonlySet<string> = new Set(['render-page', 'render-sheet', 'render-slide'])
const FORMATS: ReadonlySet<string> = new Set(['docx', 'xlsx', 'pptx'])
export async function handleWorkerRequest(req: WorkerRenderRequest, env: WorkerRenderEnv): Promise<WorkerRenderResponse> {
  try {
    if (!req || !KINDS.has(req.kind) || !FORMATS.has(req.format)) {
      return { id: req?.id ?? -1, ok: false, error: `unsupported worker request kind/format` }
    }
    if (!Number.isInteger(req.index) || req.index < 0 || !(req.source instanceof ArrayBuffer) || !req.source.byteLength) {
      return { id: req.id, ok: false, error: `invalid worker request (index/source)` }
    }
    const doc = await env.parse(req.format, new Uint8Array(req.source))
    const paintables = await env.paintables(doc as never, req.options)
    try {
      const page = paintables[req.index]
      if (!page) return { id: req.id, ok: false, error: `page index ${req.index} out of range (${paintables.length})` }
      const width = Math.max(1, Math.ceil(page.spec.widthPx)), height = Math.max(1, Math.ceil(page.spec.heightPx))
      const canvas = await env.createCanvas(width, height)
      const ctx = canvas.getContext('2d')
      if (!ctx) return { id: req.id, ok: false, error: `2d context unavailable` }
      page.paint(ctx as never)
      const bitmap = await env.toBitmap(canvas)
      return { id: req.id, ok: true, bitmap, width, height }
    } finally {
      paintables.dispose()
    }
  } catch (error) {
    return { id: (req as WorkerRenderRequest)?.id ?? -1, ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}
/** True only where a real worker + OffscreenCanvas 2D pipeline exists. */
export function isWorkerRenderSupported(): boolean {
  return typeof Worker !== 'undefined' && typeof OffscreenCanvas !== 'undefined'
}
type PendingEntry = { resolve: (bitmap: ImageBitmap) => void; reject: (error: Error) => void }
const pendingByWorker = new WeakMap<Worker, Map<number, PendingEntry>>()
function pendingFor(worker: Worker): Map<number, PendingEntry> {
  let map = pendingByWorker.get(worker)
  if (!map) {
    map = new Map()
    pendingByWorker.set(worker, map)
    const previous = worker.onmessage
    worker.onmessage = (ev: MessageEvent) => {
      if (typeof previous === 'function') previous.call(worker, ev)
      const res = ev?.data as WorkerRenderResponse | undefined
      if (!res || typeof res.id !== 'number') return
      const entry = pendingByWorker.get(worker)?.get(res.id)
      if (!entry) return
      pendingByWorker.get(worker)?.delete(res.id)
      if (res.ok) entry.resolve(res.bitmap)
      else entry.reject(new Error(res.error))
    }
  }
  return map
}
/**
 * Render one page through a worker and receive its bitmap (transferred).
 * Falls back: callers check isWorkerRenderSupported() and otherwise paint
 * the same getPaintables page on the main thread — identical code path.
 */
export function renderPageBitmap(worker: Worker, req: WorkerRenderRequest): Promise<ImageBitmap> {
  const map = pendingFor(worker)
  return new Promise<ImageBitmap>((resolve, reject) => {
    map.set(req.id, { resolve, reject })
    try {
      worker.postMessage(req, req.source ? [req.source] : [])
    } catch (error) {
      map.delete(req.id)
      reject(error instanceof Error ? error : new Error(String(error)))
    }
  })
}
