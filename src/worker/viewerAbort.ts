import type { Page, SourceFailure, SourceStore } from '../xlsx/view/contracts'
import { unavailableDrawings } from '../xlsx/view/drawingPayload'

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
export function interruptibleStore(store: SourceStore, observeWork?: (work: Promise<unknown>) => void): SourceStore {
  const watch = <T>(work: Promise<T>): Promise<T> => { observeWork?.(work); return work }
  const error: SourceFailure = { code: 'cancelled', retryable: true, message: 'Query cancelled' }
  const draw = store.readDrawings
  const page = (): Page<never> => ({ items: [], stamp: store.stamp, completeness: 'incomplete', error })
  const read = async <T>(work: () => Promise<T>, signal: AbortSignal, cancelled: () => T): Promise<T> => {
    if (signal.aborted) return cancelled()
    try { return await awaitQueryDeadline(watch(work()), signal) }
    catch (cause) { if (signal.aborted) return cancelled(); throw cause }
  }
  const facade: SourceStore = {
    get stamp() { return store.stamp },
    get extent() { return store.extent },
    get valuePolicy() { return store.valuePolicy },
    readCell: (at, options) => read(() => store.readCell(at, options), options.signal, () => ({ kind: 'error', code: 'cancelled', retryable: true })),
    readRows: (range, cursor, options) => read(() => store.readRows(range, cursor, options), options.signal, page),
    ...(draw ? {
      readDrawings: (options, profile, memory) => read(
        () => draw.call(store, options, profile, memory),
        options.signal,
        () => unavailableDrawings(error.message, error),
      ),
    } satisfies Pick<SourceStore, 'readDrawings'> : {}),
    ...(store.readRegionStyles ? { readRegionStyles: (options, grantBytes, memory) => read(
      () => store.readRegionStyles!(options, grantBytes, memory),
      options.signal,
      () => ({ sheetId: store.stamp.sheetId, stamp: store.stamp, tables: [], cells: new Int32Array(), diagnostics: [{ kind: 'unsupported' as const, feature: 'table-style-read-failed', message: error.message }] }),
    ) } satisfies Pick<SourceStore, 'readRegionStyles'> : {}),
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
      } finally { const ending = iterator.return?.(); if (ending) void watch(ending).catch(() => {}) }
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
      } finally { const ending = iterator.return?.(); if (ending) void watch(ending).catch(() => {}) }
    } } satisfies Pick<SourceStore, 'scanOccupied'> : {}),
    ...(store.settle ? { settle: (payload: object) => store.settle!(payload) } : {}),
    release: () => store.release(),
  }
  if (observeWork) {
    // Observe the facade's queued iterator cleanup as well as the raw provider;
    // return() can still be unwinding after the public request has cancelled.
    const observed = <T>(iterable: AsyncIterable<T>): AsyncIterable<T> => ({
      [Symbol.asyncIterator]() {
        const iterator = iterable[Symbol.asyncIterator]()
        return {
          next: () => watch(iterator.next()),
          ...(iterator.return ? { return: (value?: unknown) => watch(iterator.return!(value)) } : {}),
        }
      },
    })
    const populated = facade.scanPopulated.bind(facade)
    facade.scanPopulated = (...args) => observed(populated(...args))
    if (facade.scanOccupied) {
      const occupied = facade.scanOccupied.bind(facade)
      facade.scanOccupied = (...args) => observed(occupied(...args))
    }
  }
  return facade
}
