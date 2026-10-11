/** S8.4 harness scenarios at reduced scale: every runner prepares, runs the
 * delivered kernel and commits a session view with the expected semantics.
 * No timing assertions (spec §12.6); the real-browser bench measures time. */
import { describe, expect, it } from 'vitest'
import { createScenarioRunner, VIEWPORT_ROWS, type QuerySample } from '../src/debug/harness/kernelScenarios'
import { startKernelBench, unitOf, type KernelChannel } from '../src/debug/harness/kernelBench'
import type { KernelReply, KernelRequest } from '../src/debug/harness/kernelProtocol'
import { createViewerHarness } from '../src/debug/harness/harness'
import { makeEnvironment } from './helpers/harnessResult'
import { SCENARIOS, SCENARIO_IDS, isScenarioId } from '../src/debug/harness/scenarios'

describe('S8.4 scenario catalog', () => {
  it('maps each scenario to a frozen-profile threshold and an F2/F4 fixture', () => {
    expect(SCENARIO_IDS).toHaveLength(7)
    expect(new Set(SCENARIO_IDS.map(id => SCENARIOS[id].threshold))).toEqual(new Set([
      'numeric-sort-100k', 'numeric-sort-1m', 'numeric-filter-1m', 'natural-text-sort-1m', 'cold-value-index-100k',
    ]))
    expect(SCENARIOS['sort-100k-1key'].rows - 2).toBe(100_000)
    expect(SCENARIOS['sort-1m-8key'].rows).toBe(1_048_576)
    expect(isScenarioId('sort-1m-8key')).toBe(true)
    expect(isScenarioId('toString')).toBe(false)
  })
})

describe('S8.4 scenario runners (reduced scale)', () => {
  it('numeric sort: kernel completes and the session commits rows ordered by date with source-ID ties', async () => {
    const runner = createScenarioRunner('sort-100k-1key', 502)
    const prep = await runner.prepare()
    expect(prep.notes.join(' ')).toMatch(/prepareColumns\(1\) -> complete/)
    const kernel = await runner.kernel()
    expect(kernel.ms).toBeGreaterThanOrEqual(0)
    expect(kernel.extras['kernel-peak-bytes']).toBeGreaterThan(0)
    const query = await runner.query()
    expect(query.total).toBe(1_048_576)
    expect(query.viewport[0]).toEqual(['id', 'date', 'flag', 'name', 'amount', 'ratio', 'code', 'note'])
    const body = query.viewport.slice(1).map(row => [Number(row[1]), Number(row[0])])
    for (let i = 1; i < body.length; i++) {
      expect(body[i][0] > body[i - 1][0] || (body[i][0] === body[i - 1][0] && body[i][1] > body[i - 1][1])).toBe(true)
    }
    // A second query replaces the committed view (owner advances).
    expect((await runner.query()).viewport).toEqual(query.viewport)
  })

  it('eight-key sort compares every lane, lowest cardinality first', async () => {
    const runner = createScenarioRunner('sort-1m-8key', 202)
    await runner.prepare()
    await runner.kernel()
    const query = await runner.query()
    const body = query.viewport.slice(1)
    expect(body.length).toBe(VIEWPORT_ROWS - 1)
    // flag (booleans) ascending: FALSE before TRUE; then name within flag.
    expect(body[0][2]).toBe('false')
    for (let i = 1; i < body.length; i++) {
      if (body[i][2] === body[i - 1][2]) expect(body[i][3].localeCompare(body[i - 1][3])).toBeGreaterThanOrEqual(0)
    }
  })

  it('numeric filter: kernel and session agree on the visible count (about half the body)', async () => {
    const runner = createScenarioRunner('filter-1m-numeric', 1002)
    const prep = await runner.prepare()
    expect(prep.notes[0]).toMatch(/classified number/)
    const kernel = await runner.kernel()
    const query = await runner.query()
    expect(query.visible).toBe(kernel.extras['kernel-visible-rows'])
    // Every row outside the 1,000-row body stays visible; the body keeps amounts > 750 (row * 1.5).
    const bodyVisible = query.visible - (1_048_576 - 1000)
    expect(bodyVisible).toBeGreaterThan(400)
    expect(bodyVisible).toBeLessThan(600)
    expect(Number(query.viewport[1][4])).toBeGreaterThan(750)
  })

  it('natural text sort orders item-2 before item-10 with the host collator', async () => {
    const runner = createScenarioRunner('text-sort-1m-natural', 400)
    await runner.prepare()
    await runner.kernel()
    const query = await runner.query()
    const values = query.viewport.map(row => row[0])
    const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'accent', caseFirst: 'upper' })
    for (let i = 1; i < values.length; i++) expect(collator.compare(values[i - 1].replace('…', ''), values[i].replace('…', ''))).toBeLessThanOrEqual(0)
  })

  it('cold value index builds the complete domain and returns a bounded first page', async () => {
    const runner = createScenarioRunner('value-index-100k-cold', 300)
    await runner.prepare()
    const kernel = await runner.kernel()
    expect(kernel.extras['index-unique-values']).toBe(300)
    const query = await runner.query()
    expect(query.total).toBe(300)
    expect(query.viewport.length).toBe(VIEWPORT_ROWS)
  })
})

/** In-process stand-in for the module worker (same request/reply protocol). */
function inProcessChannel(): KernelChannel {
  let runner: ReturnType<typeof createScenarioRunner> | null = null
  const channel: KernelChannel = {
    onmessage: null,
    postMessage(request: KernelRequest) {
      void (async () => {
        let reply: KernelReply
        try {
          if (request.type === 'setup') {
            runner = createScenarioRunner(request.scenario as never, request.rows ?? undefined)
            reply = { type: 'ready', id: request.id, info: runner.info, prep: await runner.prepare() }
          } else if (request.type === 'kernel') reply = { type: 'kernel', id: request.id, sample: await runner!.kernel() }
          else reply = { type: 'query', id: request.id, sample: await runner!.query() }
        } catch (error) {
          reply = { type: 'error', id: request.id, message: (error as Error).message }
        }
        channel.onmessage?.({ data: reply } as MessageEvent<KernelReply>)
      })()
    },
  }
  return channel
}

describe('startKernelBench (page side)', () => {
  it('records kernel, query-completion, query-roundtrip and input-to-present separately, counting warmup', async () => {
    let t = 0
    const harness = createViewerHarness({ clock: () => (t += 5), raf: cb => setTimeout(cb, 0), environment: () => makeEnvironment(), timeoutMs: 2000 })
    const presented: QuerySample[] = []
    const bench = await startKernelBench(harness, inProcessChannel(), 'sort-100k-1key', 302, q => presented.push(q), async () => {})
    expect(bench.info).toMatchObject({ fixture: 'F2', rows: 302, bodyRows: 300 })
    await bench.run(1, 3)
    expect(presented).toHaveLength(4)
    for (const name of ['kernel', 'query-completion', 'query-roundtrip', 'input-to-present', 'kernel-peak-bytes']) {
      expect(harness.sampleCount(name), name).toBe(3)
      expect(harness.warmupCount(name), name).toBe(1)
    }
    expect(harness.sampleCount('prep:prepare-columns')).toBe(1)
    const result = await harness.exportResult()
    expect(result.samples.find(s => s.name === 'kernel-peak-bytes')?.unit).toBe('bytes')
    expect(result.samples.find(s => s.name === 'prep:prepared-columns-admitted')?.values).toEqual([1])
  })

  it('surfaces worker errors instead of recording a sample', async () => {
    const harness = createViewerHarness({ environment: () => makeEnvironment() })
    await expect(startKernelBench(harness, inProcessChannel(), 'sort-100k-1key', 1, () => {})).rejects.toThrow(/kernel worker: invalid scenario row count/)
  })

  it('derives units from series names', () => {
    expect(unitOf('kernel-peak-bytes')).toBe('bytes')
    expect(unitOf('index-unique-values')).toBe('count')
    expect(unitOf('prepare-columns')).toBe('ms')
  })
})
