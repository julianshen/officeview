/** S8.6 Node/Bun iteration diagnostic, never browser acceptance evidence.
 * bun build scripts/bench-sort-kernel.ts --target=node --outfile=/tmp/iv-s86/bench.mjs
 * node /tmp/iv-s86/bench.mjs > /tmp/iv-s86/bench.jsonl
 */
import { createScenarioRunner } from '../src/debug/harness/kernelScenarios'

console.log(JSON.stringify({ kind: 'environment', runtime: process.versions, platform: process.platform, arch: process.arch, samples: 3, warmup: 1 }))
for (const id of ['sort-100k-8key', 'sort-1m-3key', 'sort-1m-8key', 'text-sort-1m-natural'] as const) {
  const runner = createScenarioRunner(id)
  const prep = await runner.prepare()
  await runner.kernel()
  await runner.query()
  const kernels: number[] = [], queries: number[] = []
  let extras: Record<string, number> = {}
  for (let sample = 0; sample < 3; sample++) {
    const kernel = await runner.kernel()
    kernels.push(kernel.ms)
    extras = kernel.extras
    queries.push((await runner.query()).ms)
  }
  console.log(JSON.stringify({ kind: 'measurement', scenario: runner.info, prep, kernels, queries, extras }))
}
