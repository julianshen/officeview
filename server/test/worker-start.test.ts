/**
 * The worker start gate (Bun 1.3.10 workaround, see `src/worker-start.ts`): one worker loads at a time in this
 * process, and a worker is never terminated while it loads.
 */
import { describe, expect, it } from 'bun:test'
import { resolve } from 'node:path'
import { Pool } from '../src/pool'
import { startWorker, workerStartStats } from '../src/worker-start'

const SLOW = resolve(import.meta.dir, 'fixtures/workers/slow-load.ts')
const NEVER = resolve(import.meta.dir, 'fixtures/workers/never-load.ts')
const on = () => {
  const got: unknown[] = []
  let closedAt = 0
  const closed = new Promise<void>((r) => { (got as unknown as { r: () => void }).r = r })
  return { got, closed, get closedAt() { return closedAt }, events: { message: (e: MessageEvent) => { got.push(e.data) }, close: () => { closedAt = performance.now(); (got as unknown as { r: () => void }).r() } } }
}
const until = async (f: () => boolean, ms = 5000) => { const end = Date.now() + ms; while (!f()) { if (Date.now() > end) throw new Error('timed out'); await Bun.sleep(10) } }

describe('worker start gate (Bun 1.3.10: no concurrent loads, no terminate mid-load)', () => {
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
    expect(workerStartStats().terminatedLoading - before.terminatedLoading).toBe(0)
  }, 20_000)

  it('a terminate during the load waits for the load to finish; close follows', async () => {
    const a = on()
    const before = workerStartStats()
    const t0 = performance.now()
    const w = startWorker(SLOW, {}, a.events)
    w.post('hi')
    await Bun.sleep(20)
    w.terminate()
    await a.closed
    expect(a.closedAt - t0).toBeGreaterThanOrEqual(240) // not before the 250 ms load
    expect(workerStartStats().terminatedLoading - before.terminatedLoading).toBe(0)
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

  it('a load that never finishes is ended at the bound: the gate opens and the terminate goes through', async () => {
    const a = on(), b = on()
    const before = workerStartStats()
    const wa = startWorker(NEVER, {}, a.events, 300)
    const wb = startWorker(SLOW, {}, b.events)
    wb.post('next')
    await until(() => workerStartStats().started - before.started === 1) // wa spawned and loading; wb waits its turn
    const t0 = performance.now()
    wa.terminate() // deferred: wa is loading
    await Bun.sleep(100)
    expect(workerStartStats().started - before.started).toBe(1) // the gate is still held
    await a.closed
    expect(a.closedAt - t0).toBeGreaterThanOrEqual(150) // not before the 300 ms bound (spawned before t0)
    expect(workerStartStats().terminatedLoading - before.terminatedLoading).toBe(1) // ended past the bound, still loading
    await until(() => b.got.length === 1) // the gate opened for wb
    wb.terminate()
    await b.closed
  }, 20_000)

  it('a pool disposed while its workers start: one loads at a time, the rest never spawn, none is terminated mid-load', async () => {
    const before = workerStartStats()
    const pool = new Pool({ workers: 3 })
    await Bun.sleep(0) // the first worker has started loading
    pool.dispose()
    await until(() => workerStartStats().loading === 0, 15_000)
    const s = workerStartStats()
    expect(s.started - before.started).toBe(1)
    expect(s.cancelled - before.cancelled).toBe(2)
    expect(s.maxLoading).toBe(1)
    expect(s.terminatedLoading - before.terminatedLoading).toBe(0)
  }, 20_000)
})
