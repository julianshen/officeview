/**
 * The worker start gate (Bun 1.3.10 workaround, see `src/worker-start.ts`): one worker loads at a time in this
 * process, and a worker is never terminated while it loads.
 */
import { afterEach, describe, expect, it } from 'bun:test'
import { resolve } from 'node:path'
import { Pool } from '../src/pool'
import { startWorker, workerStartStats } from '../src/worker-start'
import { FakeWorker, fakeFactory } from './fixtures/fake-worker'

const SLOW = resolve(import.meta.dir, 'fixtures/workers/slow-load.ts')
const on = () => {
  const got: unknown[] = []
  let closedAt = 0
  const closed = new Promise<void>((r) => { (got as unknown as { r: () => void }).r = r })
  return { got, closed, get closedAt() { return closedAt }, events: { message: (e: MessageEvent) => { got.push(e.data) }, close: () => { closedAt = performance.now(); (got as unknown as { r: () => void }).r() } } }
}
const until = async (f: () => boolean, ms = 5000) => { const end = Date.now() + ms; while (!f()) { if (Date.now() > end) throw new Error('timed out'); await Bun.sleep(10) } }

describe('worker start gate (Bun 1.3.10: no concurrent loads, no terminate mid-load)', () => {
  // No fake may keep holding the process-wide gate after its test.
  afterEach(async () => { await FakeWorker.release(); FakeWorker.reset() })

  it('workers started together load one at a time, and messages posted before the start arrive in order', async () => {
    const before = workerStartStats()
    const a = on(), b = on()
    const wa = startWorker(SLOW, {}, a.events), wb = startWorker(SLOW, {}, b.events)
    wa.post(1); wa.post(2); wb.post(3)
    await until(() => a.got.length === 2 && b.got.length === 1)
    expect(a.got).toEqual([1, 2])
    expect(b.got).toEqual([3])
    const s = workerStartStats()
    expect(s.started - before.started).toBe(2)
    expect(s.maxLoading).toBe(1)
    wa.terminate(); wb.terminate()
    await Promise.all([a.closed, b.closed])
  }, 20_000)

  it('a terminate during the load waits for the load to finish; close follows', async () => {
    const a = on()
    const t0 = performance.now()
    const w = startWorker(SLOW, {}, a.events)
    w.post('hi')
    await Bun.sleep(20)
    w.terminate()
    await a.closed
    expect(a.closedAt - t0).toBeGreaterThanOrEqual(240) // not before the 250 ms load
  }, 20_000)

  it('a worker terminated before its start never spawns, and still reports close', async () => {
    const a = on(), b = on()
    const before = workerStartStats()
    const wa = startWorker(SLOW, {}, a.events)
    wa.post('loaded?') // its answer is its load signal
    const wb = startWorker(SLOW, {}, b.events) // queued behind wa's load
    wb.post('dropped')
    wb.terminate()
    await b.closed
    wa.terminate()
    await a.closed
    const s = workerStartStats()
    expect(s.started - before.started).toBe(1)
    expect(s.cancelled - before.cancelled).toBe(1)
    expect(b.got).toEqual([])
  }, 20_000)

  it('the bound with no terminate: the loader stays alive and keeps the gate; ready reports stalled; the next worker starts only once it answers (R1)', async () => {
    FakeWorker.scripts = [{ init: 'manual' }, { init: 0 }]
    const wa = startWorker('a', {}, on().events, { boundMs: 50, create: fakeFactory })
    wa.post({ type: 'init', id: 1 })
    const wb = startWorker('b', {}, on().events, { boundMs: 5000, create: fakeFactory })
    wb.post({ type: 'init', id: 2 })
    expect(await wa.ready).toBe('stalled')
    await Bun.sleep(100)
    // Real states: one worker constructed, alive, unacknowledged, never terminated; the next one not constructed.
    expect(FakeWorker.all).toHaveLength(1)
    expect(FakeWorker.all[0]!).toMatchObject({ acked: false, closed: false, terminateCalls: 0 })
    expect(wa.loaded).toBe(false)
    FakeWorker.all[0]!.ack()
    await until(() => FakeWorker.all.length === 2 && FakeWorker.all[1]!.acked)
    expect(wa.loaded).toBe(true)
    expect(await wb.ready).toBe('loaded')
    wa.terminate(); wb.terminate()
    await until(() => FakeWorker.all.every((w) => w.closed))
    expect(FakeWorker.all.map((w) => w.terminatedUnacked)).toEqual([false, false])
  })

  it('a terminate requested during a load is never carried out at the bound: it waits for the answer (R1)', async () => {
    FakeWorker.scripts = [{ init: 'manual' }]
    const a = on()
    const w = startWorker('a', {}, a.events, { boundMs: 30, create: fakeFactory })
    w.post({ type: 'init', id: 1 })
    await until(() => FakeWorker.all.length === 1)
    w.terminate()
    await Bun.sleep(100)
    expect(FakeWorker.all[0]!).toMatchObject({ acked: false, closed: false, terminateCalls: 0 })
    FakeWorker.all[0]!.ack()
    await a.closed
    expect(FakeWorker.all[0]!).toMatchObject({ terminateCalls: 1, terminatedUnacked: false })
  })

  it('a worker constructor that throws: the start reports close and ready closed, and a queued successor starts (R2)', async () => {
    FakeWorker.scripts = [{ throws: true }, { init: 0 }]
    let unhandled = 0
    const count = () => { unhandled++ }
    process.on('unhandledRejection', count)
    try {
      const a = on(), b = on()
      const wa = startWorker('a', {}, a.events, { create: fakeFactory })
      wa.post({ type: 'init', id: 1 })
      const wb = startWorker('b', {}, b.events, { create: fakeFactory })
      wb.post({ type: 'init', id: 2 })
      await a.closed
      expect(await wa.ready).toBe('closed')
      expect(await wb.ready).toBe('loaded')
      expect(FakeWorker.all).toHaveLength(1)
      expect(b.got).toHaveLength(1)
      wb.terminate()
      await b.closed
      await Bun.sleep(10)
      expect(unhandled).toBe(0)
      expect(workerStartStats().loading).toBe(0)
    } finally { process.off('unhandledRejection', count) }
  })

  it('a pool disposed while its workers start: one loads at a time, the rest never spawn, none is terminated mid-load', async () => {
    const before = workerStartStats()
    const pool = new Pool({ workers: 3 })
    await Bun.sleep(0) // the first worker has started loading
    // Not awaited here: the gate's counters are read while the terminate is still deferred; awaited below.
    const closed = pool.dispose()
    await until(() => workerStartStats().loading === 0, 15_000)
    const s = workerStartStats()
    expect(s.started - before.started).toBe(1)
    expect(s.cancelled - before.cancelled).toBe(2)
    expect(s.maxLoading).toBe(1)
    await closed
  }, 20_000)
})
