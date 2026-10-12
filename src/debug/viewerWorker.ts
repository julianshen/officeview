/**
 * Dev-server worker entry for the e2e page: the shipped dist/worker.js source
 * (src/worker/worker-entry.ts) loaded from inside src/debug, so the harness
 * server's locked-down fs.allow serves it through the module graph.
 */
// Capability denial (S10.3) runs before the entry module evaluates.
import './workerCaps'
import { createWorkerMessageHandler } from '../worker/worker-entry'
// The kernel imports its large-source opener on demand; named here for the same reason.
import '../worker/viewerLargeSource'
import { OfficePackage } from '../core/zip'
import { parseXlsx } from '../xlsx/parse'
import { ViewerKernel } from '../worker/viewerKernel'
import { FORCE_PAGED_LIMITS, workerNameSegments } from './forcePaged'

// S13 `?paged=1`: the worker repeats admission on the blob it is sent, so it
// gets the same lowered object cap as the page, through the kernel's existing
// importLimits env. The entry's own handler (installed on import) is replaced
// by an identical one around that kernel; nothing else differs.
const scope = globalThis as unknown as {
  name?: string
  onmessage: ((ev: MessageEvent) => void) | null
  postMessage: (message: unknown, transfer: Transferable[]) => void
}
if (workerNameSegments(scope.name).includes('paged')) {
  const kernel = new ViewerKernel({ parse: async bytes => parseXlsx(await OfficePackage.load(bytes)), importLimits: FORCE_PAGED_LIMITS })
  scope.onmessage = createWorkerMessageHandler((message, transfer) => scope.postMessage(message, transfer), kernel)
}
