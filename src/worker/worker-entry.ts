/**
 * Dedicated-worker entry (built to dist/worker.js). Parses, lays out and
 * paints inside the worker; the main thread only presents bitmaps.
 *
 * Hosts: `new Worker(new URL('officeview/dist/worker.js', import.meta.url),
 * { type: 'module' })`, then drive it with renderPageBitmap() from the main
 * entry. No DOM is touched here: measurement uses OffscreenCanvas and font
 * registration uses the worker FontFaceSet adapter.
 *
 * The same worker also hosts the resident XLSX viewer kernel (plan S4.1):
 * viewer commands (string requestId + viewer kind) route to one long-lived
 * ViewerKernel that parses each document once; everything else stays on the
 * legacy one-shot render path.
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
import { ViewerKernel } from './viewerKernel'
import { isViewerCommand } from './viewerProtocol'

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
type Post = (message: unknown, transfer: Transferable[]) => void
/** Thin message router shared by the worker scope and tests. */
export function createWorkerMessageHandler(
  post: Post,
  kernel: ViewerKernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)) }),
  legacyEnv: Parameters<typeof handleWorkerRequest>[1] = productionEnv,
): (ev: { data: unknown }) => void {
  return ev => {
    const data = ev.data
    if (isViewerCommand(data)) {
      void kernel.handle(data, (reply, transfer) => post(reply, transfer))
      return
    }
    handleWorkerRequest(data as WorkerRenderRequest, legacyEnv).then((res: WorkerRenderResponse) => {
      post(res, res.ok ? [res.bitmap] : [])
    })
  }
}
const scope = globalThis as unknown as {
  onmessage: ((ev: MessageEvent) => void) | null
  postMessage: (message: unknown, transfer: Transferable[]) => void
}
if (typeof window === 'undefined' && typeof scope.postMessage === 'function') {
  scope.onmessage = createWorkerMessageHandler((message, transfer) => scope.postMessage(message, transfer))
}
