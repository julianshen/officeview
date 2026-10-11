/**
 * Page side of the S8.4 kernel scenarios: drives the kernel worker and records
 * three separate series per iteration (spec §12.6):
 *
 * - `kernel` — the kernel call alone, timed inside the worker;
 * - `query-completion` — the worker session operation up to its atomic commit,
 *   timed inside the worker; `query-roundtrip` adds the page→worker→page
 *   message and structured-clone cost of the reply;
 * - `input-to-present` — page input mark → query reply → stand-in viewport and
 *   status written to the DOM → double requestAnimationFrame (harness.ts).
 *
 * Kernel and query runs alternate so input-to-present never includes the
 * separate kernel run. Warmup iterations run the same path and are counted,
 * not recorded.
 */
import type { ViewerHarness } from './harness'
import type { KernelReply, KernelRequest } from './kernelProtocol'
import type { QuerySample, ScenarioInfo } from './kernelScenarios'

export interface KernelChannel {
  postMessage(message: KernelRequest): void
  onmessage: ((event: MessageEvent<KernelReply>) => void) | null
}

export interface KernelBench {
  readonly info: ScenarioInfo
  readonly notes: readonly string[]
  run(warmup: number, repeat: number): Promise<void>
}

/** Pause between runs so one run's garbage collection does not land in the next. */
const SETTLE_MS = 30

export function unitOf(name: string): string {
  if (name.endsWith('-bytes')) return 'bytes'
  if (name.endsWith('-rows') || name.endsWith('-values') || name.endsWith('-admitted')) return 'count'
  return 'ms'
}

export async function startKernelBench(
  harness: ViewerHarness,
  channel: KernelChannel,
  scenario: string,
  rows: number | null,
  profile: 'D' | 'M',
  present: (sample: QuerySample) => void,
  settle: () => Promise<void> = () => new Promise(resolve => setTimeout(resolve, SETTLE_MS)),
): Promise<KernelBench> {
  let nextId = 1
  const waiting = new Map<number, { resolve: (reply: KernelReply) => void; reject: (error: Error) => void }>()
  channel.onmessage = (event) => {
    const reply = event.data
    const entry = waiting.get(reply.id)
    if (!entry) return
    waiting.delete(reply.id)
    if (reply.type === 'error') entry.reject(new Error(`kernel worker: ${reply.message}`))
    else entry.resolve(reply)
  }
  function call<T extends KernelReply['type']>(request: KernelRequest extends infer R ? R extends KernelRequest ? Omit<R, 'id'> : never : never, expect: T): Promise<Extract<KernelReply, { type: T }>> {
    const id = nextId++
    return new Promise((resolve, reject) => {
      waiting.set(id, {
        resolve: (reply) => (reply.type === expect ? resolve(reply as Extract<KernelReply, { type: T }>) : reject(new Error(`expected ${expect}, got ${reply.type}`))),
        reject,
      })
      channel.postMessage({ ...request, id } as KernelRequest)
    })
  }

  const ready = await call({ type: 'setup', scenario, rows, profile }, 'ready')
  for (const [name, value] of Object.entries(ready.prep.metrics)) harness.recordOneShot(`prep:${name}`, unitOf(name), 'operation', value)

  function record(name: string, unit: string, value: number, warmup: boolean): void {
    if (warmup) harness.recordWarmup(name, unit, 'operation')
    else harness.recordOneShot(name, unit, 'operation', value)
  }

  return {
    info: ready.info,
    notes: ready.prep.notes,
    async run(warmup: number, repeat: number): Promise<void> {
      for (let i = 0; i < warmup + repeat; i++) {
        const isWarmup = i < warmup
        const kernel = (await call({ type: 'kernel' }, 'kernel')).sample
        record('kernel', 'ms', kernel.ms, isWarmup)
        for (const [name, value] of Object.entries(kernel.extras)) record(name, unitOf(name), value, isWarmup)
        await settle()

        const label = `query-${i}`
        harness.markInput(label)
        const sent = performance.now()
        const query = (await call({ type: 'query' }, 'query')).sample
        const roundtrip = performance.now() - sent
        present(query)
        await harness.markPresent(label, isWarmup)
        record('query-completion', 'ms', query.ms, isWarmup)
        record('query-roundtrip', 'ms', roundtrip, isWarmup)
        for (const [name, value] of Object.entries(query.extras)) record(`query:${name}`, unitOf(name), value, isWarmup)
        await settle()
      }
    },
  }
}
