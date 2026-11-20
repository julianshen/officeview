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
 * (the pool's workers answer their `init`, posted first) or its close. A load that never ends is given up on after
 * `boundMs`: the gate opens for the next worker and a pending terminate goes through.
 *
 * Messages posted before the worker exists are queued and sent, in order, when it starts; a worker terminated before it
 * started never spawns and reports `close` all the same. Every worker this process starts must go through here.
 */

export const WORKER_LOAD_BOUND_MS = 10_000

export interface WorkerEvents { message(e: MessageEvent): void; close(): void }
export interface StartedWorker {
  post(message: unknown, transfer?: Transferable[]): void
  /** Idempotent. Deferred while the worker loads (at most the bound); a worker not yet started never starts. */
  terminate(): void
  readonly loaded: boolean
  /** Resolves when the load is over: loaded, closed, cancelled, or past the bound. A job's deadline starts after it. */
  readonly ready: Promise<void>
}

const stats = { started: 0, loaded: 0, loading: 0, maxLoading: 0, cancelled: 0, terminatedLoading: 0 }
/** TEST-ONLY, read-only: counters of this process's gate. `terminatedLoading` counts terminates of a worker still
 * loading (only past the bound); `maxLoading` the most workers loading at once. */
export function workerStartStats(): Readonly<typeof stats> { return { ...stats } }

/** Resolves when the previously started worker has loaded (or closed, or passed its bound). */
let gate: Promise<void> = Promise.resolve()

export function startWorker(url: string, options: WorkerOptions, on: WorkerEvents, boundMs = WORKER_LOAD_BOUND_MS): StartedWorker {
  let worker: Worker | undefined
  let loaded = false, loading = false, ended = false, wantTerminate = false, terminated = false, cancelled = false, closed = false
  let queue: Array<[unknown, Transferable[] | undefined]> = []
  let timer: ReturnType<typeof setTimeout> | undefined
  let open!: () => void
  let markReady!: () => void
  const ready = new Promise<void>((r) => { markReady = r })
  const prev = gate
  gate = new Promise<void>((r) => { open = r })

  const close = () => { if (!closed) { closed = true; on.close() } }
  const kill = () => { if (worker && !terminated) { terminated = true; if (loading) stats.terminatedLoading++; worker.terminate() } }
  /** The load is over (loaded, closed or past the bound): the next worker may start, a pending terminate goes. */
  const endLoad = () => {
    if (ended) return
    ended = true
    if (timer) clearTimeout(timer)
    open()
    markReady()
    if (wantTerminate) kill()
  }
  const settle = () => { if (loading) { loading = false; stats.loading-- } }

  void prev.then(() => {
    if (cancelled) { stats.cancelled++; open(); return }
    stats.started++
    stats.loading++; loading = true
    stats.maxLoading = Math.max(stats.maxLoading, stats.loading)
    const w = new Worker(url, options)
    worker = w
    w.onmessage = (e: MessageEvent) => {
      if (!loaded) { loaded = true; stats.loaded++; settle(); endLoad() }
      on.message(e)
    }
    w.addEventListener('close', () => { settle(); endLoad(); close() })
    timer = setTimeout(() => { endLoad(); settle() }, boundMs)
    ;(timer as { unref?: () => void }).unref?.()
    for (const [m, t] of queue) w.postMessage(m, t ?? [])
    queue = []
  })

  return {
    post(message, transfer) {
      if (cancelled || terminated) return
      if (worker) worker.postMessage(message, transfer ?? []); else queue.push([message, transfer])
    },
    terminate() {
      if (wantTerminate || cancelled) return
      wantTerminate = true
      if (!worker) { cancelled = true; queue = []; markReady(); queueMicrotask(close); return }
      if (ended) kill()
    },
    get loaded() { return loaded },
    ready,
  }
}
