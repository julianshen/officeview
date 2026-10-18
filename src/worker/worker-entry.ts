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
//
// The entry is only a bootstrap: it imports nothing the rest of the worker graph
// shares and loads the router (workerMessages.ts) on demand. WebKit does not put a
// module worker's top-level script in the worker's module map, so if any lazy chunk
// imported the entry chunk (which a bundler does when the entry holds modules the
// lazy chunk also needs, e.g. a host's Vite re-bundle of this package) WebKit would
// evaluate the entry again and its bootstrap would replace `onmessage` with a fresh,
// empty kernel ("Unknown documentId"). Messages that arrive before the router loads
// are queued and replayed in order; a failed load is rethrown so the host's Worker
// receives an `error` event instead of waiting forever.
type Handler = (ev: MessageEvent) => void
const scope = globalThis as unknown as {
  onmessage: Handler | null
  postMessage: (message: unknown, transfer: Transferable[]) => void
}
if (typeof window === 'undefined' && typeof scope.postMessage === 'function') {
  let handler: Handler | null = null
  const queued: MessageEvent[] = []
  scope.onmessage = ev => { if (handler) handler(ev); else queued.push(ev) }
  import('./workerMessages').then(
    ({ createWorkerMessageHandler }) => {
      const ready = handler = createWorkerMessageHandler((message, transfer) => scope.postMessage(message, transfer))
      for (const ev of queued.splice(0)) ready(ev)
    },
    (error: unknown) => { setTimeout(() => { throw error }) },
  )
}
export {}
