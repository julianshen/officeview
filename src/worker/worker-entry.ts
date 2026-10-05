/**
 * Dedicated-worker entry (built to dist/worker.js). Parses, lays out and
 * paints inside the worker; the main thread only presents bitmaps.
 *
 * Hosts: `new Worker(new URL('officeview/dist/worker.js', import.meta.url),
 * { type: 'module' })`, then drive it with renderPageBitmap() from the main
 * entry. No DOM is touched here: measurement uses OffscreenCanvas and font
 * registration uses the worker FontFaceSet adapter.
 */
import { OfficePackage } from '../core/zip'
import { parseDocx } from '../docx/parse'
import { parseXlsx } from '../xlsx/parse'
import { parsePptx } from '../pptx/parse'
import { getPaintables, type PaintOptions } from '../render/paint'
import { workerRegister } from '../core/fonts/register'
import {
  handleWorkerRequest,
  type WorkerCanvas,
  type WorkerRenderFormat,
  type WorkerRenderRequest,
  type WorkerRenderResponse,
} from './office-worker'

async function parse(format: WorkerRenderFormat, source: Uint8Array): Promise<unknown> {
  const pkg = await OfficePackage.load(source)
  if (format === 'xlsx') return parseXlsx(pkg)
  if (format === 'pptx') return parsePptx(pkg)
  return parseDocx(pkg)
}
const productionEnv = {
  async createCanvas(width: number, height: number): Promise<WorkerCanvas> {
    return new OffscreenCanvas(width, height) as unknown as WorkerCanvas
  },
  async toBitmap(canvas: WorkerCanvas): Promise<ImageBitmap> {
    return (canvas as unknown as OffscreenCanvas).transferToImageBitmap()
  },
  parse,
  async paintables(doc: never, options?: PaintOptions) {
    return getPaintables(doc as never, { ...options, registerFont: workerRegister })
  },
}
const scope = globalThis as unknown as {
  onmessage: ((ev: MessageEvent) => void) | null
  postMessage: (message: unknown, transfer: Transferable[]) => void
}
if (typeof window === 'undefined' && typeof scope.postMessage === 'function') {
  scope.onmessage = (ev: MessageEvent) => {
    const req = ev.data as WorkerRenderRequest
    handleWorkerRequest(req, productionEnv).then((res: WorkerRenderResponse) => {
      scope.postMessage(res, res.ok ? [res.bitmap] : [])
    })
  }
}
