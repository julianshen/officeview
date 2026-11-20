/**
 * The process's worker start gate: a workaround for two Bun 1.3.10 crashes (reproduced on Linux, see
 * .superpowers/sdd/2026-11-17-streaming-s3-p5-media/bun-segv-report.md).
 *
 * 1. `terminate()` on a worker that is still loading its module graph: a transpile job on Bun's shared thread pool
 *    writes into the terminated worker's freed source-map table (SIGSEGV in `SavedSourceMap.putValue`).
 * 2. Two workers loading at once, after earlier workers were terminated: the resolver's process-wide directory cache
 *    is corrupted (SIGSEGV in `BSSMap.getOrPut` from `new Worker` or `VirtualMachine.initWorker`).
 *
 * So, in this process, one worker loads at a time (the next `new Worker` waits until the previous one has loaded), and
 * a terminate that arrives during a load is carried out once the load ends. "Loaded" is the worker's first message
 * (the pool's workers answer their `init`, posted first); a load also ends when the worker closes. Nothing else ends
 * a load: time passing is no evidence that the module graph stopped loading.
 *
 * Startup-failure policy (bounded for the caller, safe for the runtime; Round 1, R1/R2):
 * - `boundMs` after the start was requested, `ready` resolves `'stalled'` if the load has not ended. The caller fails
 *   the work waiting on it (the pool: `RenderFailed('worker start timed out')`) instead of waiting longer.
 * - A stalled worker is NOT released and NOT terminated: it keeps the gate (no other worker starts loading beside it)
 *   and a requested terminate stays deferred, because forcibly terminating an unacknowledged loader is exactly the
 *   crashing operation. If it answers later it becomes usable (`loaded`) and the gate opens; if it closes, the gate
 *   opens and the caller replaces it. A loader that never answers and never closes keeps the gate for the life of the
 *   process: every later start then reports `'stalled'` at its bound and its work fails with a controlled error.
 * - A `new Worker` that throws is a closed start: the gate opens for the next worker, `ready` resolves `'closed'` and
 *   `close` is reported (on a microtask), so the caller's waits end.
 *
 * Messages posted before the worker exists are queued and sent, in order, when it starts; a worker terminated before it
 * started never spawns and reports `close` all the same. Every worker this process starts must go through here, and
 * must answer its first message promptly (or it holds the gate).
 */

export const WORKER_LOAD_BOUND_MS = 10_000

/** `loaded`: answered its first message; `closed`: closed (or never started) before that; `stalled`: the bound
 * passed first (it may still load later: check `loaded`). */
export type WorkerReadiness = 'loaded' | 'closed' | 'stalled'
export interface WorkerEvents { message(e: MessageEvent): void; close(): void }
export interface StartedWorker {
  post(message: unknown, transfer?: Transferable[]): void
  /** Idempotent. Deferred while the worker loads (however long that takes); a worker not yet started never starts. */
  terminate(): void
  readonly loaded: boolean
  /** Resolves once: loaded, closed (or cancelled, or a failed construction), or stalled at the bound. */
  readonly ready: Promise<WorkerReadiness>
}
export interface StartOptions {
  /** The startup bound, measured from the request (default `WORKER_LOAD_BOUND_MS`). */
  boundMs?: number
  /** TEST-ONLY seam: builds the worker (default `new Worker(url, options)`). */
  create?: (url: string, options: WorkerOptions) => Worker
}

const stats = { started: 0, loaded: 0, loading: 0, maxLoading: 0, cancelled: 0, failed: 0, stalled: 0 }
/** TEST-ONLY, read-only: counters of this process's gate (`maxLoading`: the most workers loading at once). */
export function workerStartStats(): Readonly<typeof stats> { return { ...stats } }

/** Resolves when the previously started worker's load has ended (loaded or closed), or it was cancelled. */
let gate: Promise<void> = Promise.resolve()

export function startWorker(url: string, options: WorkerOptions, on: WorkerEvents, o: StartOptions = {}): StartedWorker {
  const create = o.create ?? ((u: string, opts: WorkerOptions) => new Worker(u, opts))
  let worker: Worker | undefined
  let loaded = false, loading = false, ended = false, wantTerminate = false, terminated = false, cancelled = false, closed = false
  let queue: Array<[unknown, Transferable[] | undefined]> = []
  let open!: () => void
  let resolveReady!: (r: WorkerReadiness) => void
  let readied = false
  const ready = new Promise<WorkerReadiness>((r) => { resolveReady = r })
  const markReady = (r: WorkerReadiness) => { if (!readied) { readied = true; clearTimeout(timer); resolveReady(r) } }
  const prev = gate
  gate = new Promise<void>((r) => { open = r })
  const timer = setTimeout(() => { if (!readied) { stats.stalled++; markReady('stalled') } }, o.boundMs ?? WORKER_LOAD_BOUND_MS)
  ;(timer as { unref?: () => void }).unref?.()

  const close = () => { if (!closed) { closed = true; on.close() } }
  const kill = () => { if (worker && !terminated) { terminated = true; worker.terminate() } }
  /** The load is over (loaded or closed): the next worker may start, a pending terminate goes. */
  const endLoad = (r: 'loaded' | 'closed') => {
    if (ended) return
    ended = true
    if (loading) { loading = false; stats.loading-- }
    open()
    markReady(r)
    if (wantTerminate) kill()
  }

  void prev.then(() => {
    if (cancelled) { stats.cancelled++; open(); return }
    let w: Worker
    try { w = create(url, options) } catch {
      // R2: nothing was created, so nothing is loading: a closed start.
      stats.failed++
      ended = true; cancelled = true; queue = []
      open(); markReady('closed'); queueMicrotask(close)
      return
    }
    stats.started++
    stats.loading++; loading = true
    stats.maxLoading = Math.max(stats.maxLoading, stats.loading)
    worker = w
    w.onmessage = (e: MessageEvent) => {
      if (!loaded) { loaded = true; stats.loaded++; endLoad('loaded') }
      on.message(e)
    }
    w.addEventListener('close', () => { endLoad('closed'); close() })
    for (const [m, t] of queue) w.postMessage(m, t ?? [])
    queue = []
  })

  return {
    post(message, transfer) {
      if (cancelled || terminated || closed) return
      if (worker) worker.postMessage(message, transfer ?? []); else queue.push([message, transfer])
    },
    terminate() {
      if (wantTerminate || cancelled) return
      wantTerminate = true
      if (!worker) { cancelled = true; queue = []; markReady('closed'); queueMicrotask(close); return }
      if (ended) kill()
    },
    get loaded() { return loaded },
    ready,
  }
}
