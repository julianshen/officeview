/**
 * Module Worker entry for the S8.4 kernel scenarios (dev harness only). The
 * kernels run here, off the page's main thread, as they do in the viewer
 * worker; the page measures the message round-trip and presentation.
 */
import { createScenarioRunner, type ScenarioRunner } from './kernelScenarios'
import { isScenarioId } from './scenarios'
import type { KernelRequest, KernelReply } from './kernelProtocol'

let runner: ScenarioRunner | null = null
const post = (reply: KernelReply): void => (self as unknown as Worker).postMessage(reply)

self.onmessage = async (event: MessageEvent<KernelRequest>) => {
  const request = event.data
  try {
    switch (request.type) {
      case 'setup': {
        if (runner) throw new Error('scenario already set up; one scenario per worker')
        if (!isScenarioId(request.scenario)) throw new Error(`unknown scenario ${request.scenario}`)
        if (request.profile !== 'D' && request.profile !== 'M') throw new Error(`unknown profile ${String(request.profile)}`)
        runner = createScenarioRunner(request.scenario, request.rows ?? undefined, request.profile)
        const prep = await runner.prepare()
        post({ type: 'ready', id: request.id, info: runner.info, prep })
        break
      }
      case 'kernel': {
        if (!runner) throw new Error('kernel before setup')
        post({ type: 'kernel', id: request.id, sample: await runner.kernel() })
        break
      }
      case 'query': {
        if (!runner) throw new Error('query before setup')
        post({ type: 'query', id: request.id, sample: await runner.query() })
        break
      }
    }
  } catch (error) {
    post({ type: 'error', id: request.id, message: error instanceof Error ? `${error.message}\n${error.stack ?? ''}` : String(error) })
  }
}
