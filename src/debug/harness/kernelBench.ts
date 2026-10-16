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
import type { QuerySample, SampleRefusal, ScenarioInfo } from './kernelScenarios'
import type { ScenarioRefusal } from './types'
import { SCENARIOS } from './scenarios'
import { WORKER_BUDGETS } from '../../xlsx/view/tileCache'

export interface KernelChannel {
  postMessage(message: KernelRequest): void
  onmessage: ((event: MessageEvent<KernelReply>) => void) | null
}

export interface KernelBench {
  readonly info: ScenarioInfo
  readonly notes: readonly string[]
  /** Declared design refusals observed by run(); empty when every run completed. */
  readonly refusals: readonly ScenarioRefusal[]
  run(warmup: number, repeat: number): Promise<void>
}

/** A refusal is "by design" only when the scenario declares it for this
 * profile AND it names exactly that profile's worker budget; anything else
 * stays a harness failure. */
export function designRefusal(info: ScenarioInfo, refusal: SampleRefusal): ScenarioRefusal {
  const reason = SCENARIOS[info.id].designRefusal?.[info.profile]
  if (reason === undefined || refusal.limitBytes !== WORKER_BUDGETS[info.profile]) {
    throw new Error(`${info.id} ${refusal.boundary} refused on profile ${info.profile} (${refusal.kind} ${refusal.requiredBytes} B > ${refusal.limitBytes} B): not a declared design refusal`)
  }
  return { outcome: 'refused-by-design', ...refusal, reason }
}

/** Pause between runs so one run's garbage collection does not land in the next. */
const SETTLE_MS = 30

export function unitOf(name: string): string {
  if (name.endsWith('-bytes')) return 'bytes'
  if (name.endsWith('-rows') || name.endsWith('-values') || name.endsWith('-admitted') || name.endsWith('-reads')) return 'count'
  return 'ms'
}

export async function startKernelBench(
  harness: ViewerHarness,
  channel: KernelChannel,
  scenario: string,
  rows: number | null,
  profile: 'D' | 'M',
  admitted: boolean,
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

  const ready = await call({ type: 'setup', scenario, rows, profile, admitted }, 'ready')
  for (const [name, value] of Object.entries(ready.prep.metrics)) harness.recordOneShot(`prep:${name}`, unitOf(name), 'operation', value)

  function record(name: string, unit: string, value: number, warmup: boolean): void {
    if (warmup) harness.recordWarmup(name, unit, 'operation')
    else harness.recordOneShot(name, unit, 'operation', value)
  }

  const refusals: ScenarioRefusal[] = []
  return {
    info: ready.info,
    notes: ready.prep.notes,
    refusals,
    async run(warmup: number, repeat: number): Promise<void> {
      for (let i = 0; i < warmup + repeat; i++) {
        const isWarmup = i < warmup
        const kernel = (await call({ type: 'kernel' }, 'kernel')).sample
        // A refused run is recorded once as an outcome, never as a timing
        // sample; the query still runs once so the product path's refusal is
        // recorded too. The scenario's admission does not change between
        // iterations, so the remaining iterations are not run.
        if (kernel.refusal) refusals.push(designRefusal(ready.info, kernel.refusal))
        else {
          record('kernel', 'ms', kernel.ms, isWarmup)
          for (const [name, value] of Object.entries(kernel.extras)) record(name, unitOf(name), value, isWarmup)
        }
        await settle()

        const label = `query-${i}`
        const sent = performance.now()
        if (!refusals.length) harness.markInput(label)
        const query = (await call({ type: 'query' }, 'query')).sample
        if (query.refusal) {
          const refusal = designRefusal(ready.info, query.refusal)
          if (!refusals.length) await harness.markPresent(label, true)
          refusals.push(refusal)
        } else if (refusals.length) {
          throw new Error(`${ready.info.id}: the kernel was refused but the session query committed; the outcome is inconsistent`)
        } else {
          const roundtrip = performance.now() - sent
          present(query)
          await harness.markPresent(label, isWarmup)
          record('query-completion', 'ms', query.ms, isWarmup)
          record('query-roundtrip', 'ms', roundtrip, isWarmup)
          for (const [name, value] of Object.entries(query.extras)) record(`query:${name}`, unitOf(name), value, isWarmup)
        }
        await settle()
        if (refusals.length) {
          for (const refusal of refusals) harness.recordScenarioRefusal(refusal)
          return
        }
      }
    },
  }
}
