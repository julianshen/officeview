/** Mobile M admission of the 1M natural-text sort (S8.4 text-sort-1m-natural),
 * at FULL scale, in node with the M budget profile (no simulator needed).
 *
 * The iPad run reported "requiredBytes 268,435,526 > limitBytes 268,435,456"
 * (70 B over). That figure is the worker aggregate at the FIRST charge that
 * crossed the 256 MiB budget, not the operation's peak: the sort halts about
 * 11% into its key read. The exact peaks below prove the real shortfall, so
 * the harness records the refusal as by design instead of failing. */
import { describe, expect, it } from 'vitest'
import { createFixture, type FixtureCell } from './helpers/xlsx-viewer-fixtures'
import { createScenarioRunner } from '../src/debug/harness/kernelScenarios'
import { startKernelBench, type KernelChannel } from '../src/debug/harness/kernelBench'
import type { KernelReply, KernelRequest } from '../src/debug/harness/kernelProtocol'
import { createViewerHarness } from '../src/debug/harness/harness'
import { validateHarnessResult } from '../src/debug/harness/result'
import { sortRowOrder } from '../src/xlsx/view/sort'
import { WORKER_BUDGETS } from '../src/xlsx/view/tileCache'
import { makeEnvironment } from './helpers/harnessResult'
import { SCENARIOS } from '../src/debug/harness/scenarios'
import type { HarnessEnvironment } from '../src/debug/harness/types'

const ROWS = 1_000_000
const M = 268_435_456
const D = 536_870_912
/** Resident F4 cells: 1M × 64 B cell objects + Σ(2·UTF-16 + 40 B) strings. */
const RESIDENT_F4 = 239_052_798
const F4_STRINGS = 175_052_798
const LONG = 120_000

function inProcessChannel(): KernelChannel {
  let runner: ReturnType<typeof createScenarioRunner> | null = null
  const channel: KernelChannel = {
    onmessage: null,
    postMessage(request: KernelRequest) {
      void (async () => {
        let reply: KernelReply
        try {
          if (request.type === 'setup') {
            runner = createScenarioRunner(request.scenario as never, request.rows ?? undefined, request.profile, request.admitted)
            reply = { type: 'ready', id: request.id, info: runner.info, prep: await runner.prepare() }
          } else if (request.type === 'kernel') reply = { type: 'kernel', id: request.id, sample: await runner!.kernel() }
          else reply = { type: 'query', id: request.id, sample: await runner!.query() }
        } catch (error) {
          reply = { type: 'error', id: request.id, message: error instanceof Error ? error.message : String(error) }
        }
        channel.onmessage?.({ data: reply } as MessageEvent<KernelReply>)
      })()
    },
  }
  return channel
}

describe('M profile: 1M natural text sort (full scale)', () => {
  it('uses the spec M worker budget (256 MiB)', () => {
    expect(WORKER_BUDGETS.M).toBe(M)
    expect(WORKER_BUDGETS.D).toBe(D)
  })

  it('exact kernel peaks: neither the resident nor the cold source path fits 256 MiB', async () => {
    const handle = createFixture('F4', { rowCount: ROWS })
    const cells: FixtureCell[] = new Array(ROWS)
    let resident = 0, strings = 0
    for (let row = 0; row < ROWS; row++) {
      const cell = handle.cellAt(row, 0)!
      cells[row] = cell
      const text = (cell.value as string).length * 2 + 40
      resident += 64 + text
      strings += text
    }
    expect(resident).toBe(RESIDENT_F4)
    expect(strings).toBe(F4_STRINGS)
    const store = handle.createStore({ read: { cellAt: (row, col) => (col === 0 ? cells[row] ?? null : null) } })
    const spec = { range: { r0: 0, r1: ROWS - 1, c0: 0, c1: 0, headerRows: 0 as const, totalsRows: 0 as const }, keys: [{ col: 0, order: 'asc' as const }], caseSensitive: false }
    const unbounded = { locale: 'en', maxWorkerBytes: 4 * 1024 ** 3, residentWorkerBytes: 0 }

    // Resident source owns the strings: the dictionary charges a 16 B slot each.
    const owned = await sortRowOrder(store, spec, { ...unbounded, sourceRetainsStrings: true })
    expect(owned.kind).toBe('complete')
    if (owned.kind !== 'complete') return
    expect(owned.ledger).toMatchObject({
      fixedOverheadBytes: 1_536, tagsBytes: 1_000_000, numbersBytes: 8_000_000, idsBytes: 4_000_000,
      dictionaryBytes: 16_000_000, dictionaryMapPeakBytes: 128_000_000, ranksBytes: 4_000_000, textKeyBytes: 12_000_000,
      indexBytes: 4_000_000, scratchBytes: 4_000_000, orderBytes: 4_000_000, inverseBytes: 4_000_000,
      peakBytes: 158_622_016,
    })
    // Peak is the end of the key read: lanes + ids + slots + 128 B/entry intern
    // map (157,001,536 B) plus the last in-flight source page.
    expect(RESIDENT_F4 + owned.ledger.peakBytes).toBe(397_674_814)
    expect(RESIDENT_F4 + owned.ledger.peakBytes - M).toBe(129_239_358)
    expect(RESIDENT_F4 + owned.ledger.peakBytes).toBeLessThanOrEqual(D)

    // Cold/paged source: nothing else retains the strings, so the dictionary owns copies.
    const cold = await sortRowOrder(store, spec, unbounded)
    expect(cold.kind).toBe('complete')
    if (cold.kind !== 'complete') return
    expect(cold.ledger.dictionaryBytes).toBe(F4_STRINGS)
    expect(cold.ledger.peakBytes).toBe(317_379_338)
    expect(cold.ledger.peakBytes - M).toBe(48_943_882)
  }, LONG)

  it('admitted M run: kernel and session return the typed worker-budget refusal (first crossing charge)', async () => {
    const runner = createScenarioRunner('text-sort-1m-natural', undefined, 'M', true)
    const prep = await runner.prepare()
    expect(prep.metrics['resident-source-bytes']).toBe(RESIDENT_F4)
    const kernel = await runner.kernel()
    expect(kernel.refusal).toEqual({ boundary: 'kernel', kind: 'resource-limit', requiredBytes: 268_435_526, limitBytes: M })
    const query = await runner.query()
    expect(query.refusal).toEqual({ boundary: 'query', kind: 'quota', requiredBytes: 268_435_498, limitBytes: M })
  }, LONG)

  it('unadmitted M run (nothing resident charged) is refused too', async () => {
    const runner = createScenarioRunner('text-sort-1m-natural', undefined, 'M', false)
    await runner.prepare()
    expect((await runner.kernel()).refusal).toEqual({ boundary: 'kernel', kind: 'resource-limit', requiredBytes: 268_436_558, limitBytes: M })
    expect((await runner.query()).refusal).toEqual({ boundary: 'query', kind: 'quota', requiredBytes: 268_435_472, limitBytes: M })
  }, LONG)

  it('the bench records "refused (by design)" once, with no timing samples, and the result validates', async () => {
    let fixtureScenario: HarnessEnvironment['scenario']
    const harness = createViewerHarness({
      raf: cb => setTimeout(cb, 0), timeoutMs: 5000,
      environment: fixture => makeEnvironment({ profile: 'mobile', viewport: { width: 1024, height: 768 }, ...(fixture?.scenario ? { scenario: fixtureScenario = fixture.scenario } : {}), note: fixture?.note ?? 'x' }),
    })
    harness.setFixture({ adapter: 'synthetic-fixture-store', bytes: 0, rows: ROWS, cols: 1, scenario: { id: 'text-sort-1m-natural', kind: 'sort', keys: [0], rows: ROWS, admitted: true }, note: 'long scenario note '.repeat(60) }, null)
    const bench = await startKernelBench(harness, inProcessChannel(), 'text-sort-1m-natural', null, 'M', true, () => {}, async () => {})
    await bench.run(3, 30)
    expect(bench.refusals).toHaveLength(2)
    expect(harness.sampleCount('kernel')).toBe(0)
    expect(harness.warmupCount('kernel')).toBe(0)
    expect(harness.sampleCount('query-completion')).toBe(0)
    const result = await harness.exportResult()
    expect(fixtureScenario?.refusals).toEqual([
      { outcome: 'refused-by-design', boundary: 'kernel', kind: 'resource-limit', requiredBytes: 268_435_526, limitBytes: M, reason: SCENARIOS['text-sort-1m-natural'].designRefusal!.M },
      { outcome: 'refused-by-design', boundary: 'query', kind: 'quota', requiredBytes: 268_435_498, limitBytes: M, reason: SCENARIOS['text-sort-1m-natural'].designRefusal!.M },
    ])
    expect(result.environment.note).toMatch(/^REFUSED \(by design\): kernel resource-limit 268435526 B > 268435456 B; query quota 268435498 B > 268435456 B/)
    expect(result.environment.note.length).toBeLessThanOrEqual(1024)
    const validation = validateHarnessResult(result)
    expect(validation.ok ? [] : validation.errors).toEqual([])
  }, LONG)
})

describe('design-refusal bookkeeping', () => {
  function refusingChannel(profile: 'D' | 'M', limitBytes: number): KernelChannel {
    const channel: KernelChannel = {
      onmessage: null,
      postMessage(request: KernelRequest) {
        const info = { id: 'text-sort-1m-natural' as const, profile, admitted: true, keys: [0], fixture: 'F4' as const, rows: ROWS, cols: 1, bodyRows: ROWS, description: '' }
        const refusal = { kind: 'resource-limit' as const, requiredBytes: limitBytes + 1, limitBytes }
        const reply: KernelReply = request.type === 'setup' ? { type: 'ready', id: request.id, info, prep: { metrics: {}, notes: [] } }
          : request.type === 'kernel' ? { type: 'kernel', id: request.id, sample: { ms: 1, extras: {}, refusal: { boundary: 'kernel', ...refusal } } }
            : { type: 'query', id: request.id, sample: { ms: 1, visible: 0, total: 0, viewport: [], extras: {}, refusal: { boundary: 'query', ...refusal, kind: 'quota' } } }
        queueMicrotask(() => channel.onmessage?.({ data: reply } as MessageEvent<KernelReply>))
      },
    }
    return channel
  }

  it('a refusal the scenario does not declare for the profile is a failure, not "by design"', async () => {
    const harness = createViewerHarness({ environment: () => makeEnvironment() })
    const bench = await startKernelBench(harness, refusingChannel('D', D), 'text-sort-1m-natural', null, 'D', true, () => {}, async () => {})
    await expect(bench.run(0, 1)).rejects.toThrow(/refused.*not a declared design refusal/)
  })

  it('an M refusal against some other limit than the M worker budget is a failure', async () => {
    const harness = createViewerHarness({ environment: () => makeEnvironment() })
    const bench = await startKernelBench(harness, refusingChannel('M', 33_554_432), 'text-sort-1m-natural', null, 'M', true, () => {}, async () => {})
    await expect(bench.run(0, 1)).rejects.toThrow(/not a declared design refusal/)
  })

  it('the result validator accepts declared refusals and rejects malformed ones', async () => {
    const scenario = { id: 'text-sort-1m-natural', kind: 'sort' as const, keys: [0], rows: ROWS, admitted: true }
    const good = { outcome: 'refused-by-design' as const, boundary: 'kernel' as const, kind: 'resource-limit' as const, requiredBytes: M + 70, limitBytes: M, reason: 'r' }
    const build = async (refusals: unknown) => {
      const harness = createViewerHarness({ environment: () => makeEnvironment({ scenario: { ...scenario, refusals } as never }) })
      // finalizeResult refuses to emit an invalid result at all.
      try { return validateHarnessResult(await harness.exportResult()) } catch (error) { return { ok: false, errors: [String(error)] } }
    }
    expect((await build([good])).ok).toBe(true)
    for (const bad of [[{ ...good, outcome: 'refused' }], [{ ...good, requiredBytes: M }], [{ ...good, boundary: 'x' }], [{ ...good, extra: 1 }], [], 'x']) {
      expect((await build(bad)).ok, JSON.stringify(bad)).toBe(false)
    }
  })
})
