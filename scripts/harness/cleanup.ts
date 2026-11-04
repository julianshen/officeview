/**
 * Once-only teardown shared by the in-band `finally` path and OS signals.
 * The manual harnesses start things that outlive a killed process (VoiceOver
 * keeps speaking; safaridriver and the loopback servers keep their ports), so
 * Ctrl+C / SIGTERM / SIGHUP must run the same teardown before exiting.
 *
 * The teardown is bounded: after `deadlineMs` the run settles anyway (and a
 * signal exits anyway), so one wedged dependency cannot hold the process.
 * Order matters inside the teardown: the VoiceOver check stops VoiceOver
 * first, with its own short per-step timeouts, so it is attempted and
 * verified well inside the deadline.
 */

export type TeardownSignal = 'SIGINT' | 'SIGTERM' | 'SIGHUP'
const SIGNALS: readonly TeardownSignal[] = ['SIGINT', 'SIGTERM', 'SIGHUP']
/** Default overall teardown deadline. */
export const TEARDOWN_DEADLINE_MS = 10_000
/** Conventional shell exit status for death by signal: 128 + signal number. */
export const SIGNAL_EXIT_CODES: Record<TeardownSignal, number> = { SIGHUP: 129, SIGINT: 130, SIGTERM: 143 }

export interface SignalHost {
  once(signal: TeardownSignal, listener: () => void): unknown
  off(signal: TeardownSignal, listener: () => void): unknown
  exit(code: number): never | void
}

export interface Teardown {
  /** Runs the teardown at most once; later callers await the same run. */
  run(): Promise<void>
  /** Removes the signal handlers (after an in-band teardown). */
  dispose(): void
}

export function installTeardown(teardown: () => Promise<void>, host: SignalHost = process as unknown as SignalHost, log: (message: string) => void = message => console.error(message), deadlineMs = TEARDOWN_DEADLINE_MS): Teardown {
  let running: Promise<void> | null = null
  const bounded = () => {
    let timer: ReturnType<typeof setTimeout> | undefined
    const deadline = new Promise<void>(resolve => {
      timer = setTimeout(() => { log(`teardown did not finish within ${deadlineMs} ms; continuing without it`); resolve() }, deadlineMs)
    })
    const work = teardown().catch(error => log(`teardown failed: ${error instanceof Error ? error.message : String(error)}`))
    return Promise.race([work, deadline]).finally(() => clearTimeout(timer))
  }
  const run = () => running ??= bounded()
  const handlers = new Map<TeardownSignal, () => void>()
  const dispose = () => { for (const [signal, handler] of handlers) host.off(signal, handler); handlers.clear() }
  for (const signal of SIGNALS) {
    const handler = () => {
      log(`${signal} received: cleaning up before exit`)
      dispose()
      void run().then(() => host.exit(SIGNAL_EXIT_CODES[signal]))
    }
    handlers.set(signal, handler)
    host.once(signal, handler)
  }
  return { run, dispose }
}
