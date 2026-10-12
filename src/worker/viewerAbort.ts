import type { Page, SourceFailure, SourceStore } from '../xlsx/view/contracts'

const deadlines = new WeakMap<AbortSignal, AbortSignal>()
export function bindQueryDeadline(signal: AbortSignal, deadline: AbortSignal): void { deadlines.set(signal, deadline) }
export function awaitQueryDeadline<T>(work: Promise<T>, signal: AbortSignal): Promise<T> {
  const deadline = deadlines.get(signal)
  return deadline ? awaitAbortable(awaitAbortable(work, signal), deadline) : awaitAbortable(work, signal)
}

/** Abort settles the owner without waiting for a provider that ignores its signal.
 * The late fulfillment/rejection remains observed, but cannot publish a result. */
export function awaitAbortable<T>(work: Promise<T>, signal: AbortSignal, interrupt: () => boolean = () => true): Promise<T> {
  return new Promise((resolve, reject) => {
    const abort = () => { if (interrupt()) reject(new DOMException('Query cancelled', 'AbortError')) }
    if (signal.aborted) abort()
    else signal.addEventListener('abort', abort, { once: true })
    work.then(value => { signal.removeEventListener('abort', abort); resolve(value) },
      error => { signal.removeEventListener('abort', abort); reject(error) })
  })
}

/** Keep the source's typed cancellation contract while unwinding session finally
 * blocks (draft charges, prepared leases, scheduler ports) promptly on abort. */
export function interruptibleStore(store: SourceStore): SourceStore {
  const error: SourceFailure = { code: 'cancelled', retryable: true, message: 'Query cancelled' }
  const page = (): Page<never> => ({ items: [], stamp: store.stamp, completeness: 'incomplete', error })
  const read = async <T>(work: () => Promise<T>, signal: AbortSignal, cancelled: () => T): Promise<T> => {
    if (signal.aborted) return cancelled()
    try { return await awaitQueryDeadline(work(), signal) }
    catch (cause) { if (signal.aborted) return cancelled(); throw cause }
  }
  return {
    get stamp() { return store.stamp },
    get extent() { return store.extent },
    get valuePolicy() { return store.valuePolicy },
    readCell: (at, options) => read(() => store.readCell(at, options), options.signal, () => ({ kind: 'error', code: 'cancelled', retryable: true })),
    readRows: (range, cursor, options) => read(() => store.readRows(range, cursor, options), options.signal, page),
    queryColumn: (request, options) => read(() => store.queryColumn(request, options), options.signal, page),
    queryOverview: (request, options) => read(() => store.queryOverview(request, options), options.signal, page),
    queryStructural: (rect, options) => read(() => store.queryStructural(rect, options), options.signal, () => ({ stamp: store.stamp, completeness: 'incomplete', merges: [], tables: [], spills: [], drawings: [], missing: [rect], error })),
    ensureComplete: (rect, options) => read(() => store.ensureComplete(rect, options), options.signal, () => ({ range: rect, stamp: store.stamp, completeness: 'incomplete', missing: [rect], error })),
    async *scanPopulated(request, options) {
      const iterator = store.scanPopulated(request, options)[Symbol.asyncIterator]()
      try {
        for (;;) {
          const next = await read(() => iterator.next(), options.signal, () => ({ done: false as const, value: page() }))
          if (next.done) return
          yield next.value
          if (options.signal.aborted) return
        }
      } finally { void iterator.return?.().catch(() => {}) }
    },
    ...(store.scanOccupied ? { async *scanOccupied(request, options) {
      const iterator = store.scanOccupied!(request, options)[Symbol.asyncIterator]()
      try {
        for (;;) {
          const next = await read(() => iterator.next(), options.signal, () => ({ done: false as const, value: page() }))
          if (next.done) return
          yield next.value
          if (options.signal.aborted) return
        }
      } finally { void iterator.return?.().catch(() => {}) }
    } } satisfies Pick<SourceStore, 'scanOccupied'> : {}),
    ...(store.settle ? { settle: (payload: object) => store.settle!(payload) } : {}),
    release: () => store.release(),
  }
}
