/**
 * The pool's worker-startup paths (Bun 1.3.10 start gate, `src/worker-start.ts`), on controllable stand-in workers:
 * a job waits for its worker's load before its deadline starts; a worker that dies while loading gets one fresh
 * replacement, and a second death, a dispose during the replacement's start or a stalled start fails the job with
 * `RenderFailed`, posting nothing and keeping no pending entry.
 */
import { afterEach, describe, expect, it } from 'bun:test'
import { rejection, rejectionMessage } from './fixtures/rejection'
import { DEFAULT_CAPS } from '../src/admission'
import { Pool, RenderFailed, type EncodeJob, type PoolOptions } from '../src/pool'
import { FakeWorker, fakeFactory } from './fixtures/fake-worker'

const job: EncodeJob = { png: new Uint8Array(1), w: 1, h: 1, kBasis: new Uint8Array(1), code: new Int8Array(1) }
const until = async (f: () => boolean, ms = 5000) => { const end = Date.now() + ms; while (!f()) { if (Date.now() > end) throw new Error('timed out'); await Bun.sleep(5) } }
let pool: Pool | undefined
const mk = (o: PoolOptions = {}) => (pool = new Pool({ workers: 1, testHooks: true, workerFactoryForTests: fakeFactory, ...o }))
const pending = (p: Pool) => (p as unknown as { slots: Array<{ pending: Map<number, unknown> }> }).slots[0]!.pending.size
/** Messages a fake was sent beyond its init. */
const jobsSent = (w: FakeWorker) => w.received.filter((m) => (m as { type?: string }).type !== 'init').length

// The dispose's wait is awaited after `release()`: a fake still loading (manual init) closes only once released.
afterEach(async () => { const closed = pool?.dispose(); pool = undefined; await FakeWorker.release(); await closed; FakeWorker.reset() })

describe('pool worker startup (start gate)', () => {
  it("a job's deadline starts once its worker has loaded: a 300 ms load does not time out a 100 ms job", async () => {
    FakeWorker.scripts = [{ init: 300 }]
    const p = mk({ caps: { ...DEFAULT_CAPS, renderDeadlineMs: 100 } })
    expect(await p.encode('s', job, 16)).toBeNull()
    expect(p.stats().retired).toBe(0)
  })

  it('a worker that dies while loading, then a replacement that dies too: RenderFailed, nothing posted, no pending entry (R3)', async () => {
    FakeWorker.scripts = [{ die: true }, { die: true }]
    const p = mk()
    expect(await rejection(p.encode('s', job, 16))).toBeInstanceOf(RenderFailed)
    expect(FakeWorker.all).toHaveLength(2)
    expect(FakeWorker.all.map(jobsSent)).toEqual([0, 0])
    expect(pending(p)).toBe(0)
    expect(p.stats()).toMatchObject({ reserved: 0, running: 0 })
  })

  it("a dispose during the replacement's start: RenderFailed, the job is never posted, and the replacement is terminated only once it has loaded (R3)", async () => {
    FakeWorker.scripts = [{ die: true }, { init: 'manual' }]
    const p = mk()
    const r = p.encode('s', job, 16)
    await until(() => FakeWorker.all.length === 2)
    // Not awaited here: the replacement is still loading (the terminate is deferred until its ack); awaited below.
    const closed = p.dispose()
    const fresh = FakeWorker.all[1]!
    expect(fresh.terminateCalls).toBe(0) // still loading: deferred
    fresh.ack()
    expect(await rejectionMessage(r)).toContain('pool disposed')
    await until(() => fresh.closed)
    expect(fresh).toMatchObject({ terminateCalls: 1, terminatedUnacked: false })
    expect(jobsSent(fresh)).toBe(0)
    expect(pending(p)).toBe(0)
    await closed
  })

  it('a start that passes the load bound fails its jobs with RenderFailed and is never terminated while loading; once it answers, the slot works (R1)', async () => {
    FakeWorker.scripts = [{ init: 'manual' }]
    const p = mk({ workerLoadBoundMsForTests: 80 })
    expect(await rejectionMessage(p.encode('s', job, 16))).toMatch(/worker start timed out/)
    expect(await rejectionMessage(p.encode('s', job, 16))).toMatch(/worker start timed out/) // at once, not another bound
    const w = FakeWorker.all[0]!
    expect(w).toMatchObject({ acked: false, closed: false, terminateCalls: 0 })
    expect(jobsSent(w)).toBe(0)
    expect(pending(p)).toBe(0)
    w.ack()
    await until(() => w.acked)
    expect(await p.encode('s', job, 16)).toBeNull()
    expect(FakeWorker.all).toHaveLength(1)
  })

  it("dispose() resolves only once every worker it terminated has closed: no worker of the pool is left alive", async () => {
    const p = mk({ workers: 2 })
    await until(() => FakeWorker.all.length === 2 && FakeWorker.all.every((w) => w.acked))
    await p.dispose()
    expect(FakeWorker.all.map((w) => ({ terminateCalls: w.terminateCalls, closed: w.closed }))).toEqual([{ terminateCalls: 1, closed: true }, { terminateCalls: 1, closed: true }])
  })

  it("a dispose during a load: its promise waits for the deferred terminate's close (it does not settle early, nor hang)", async () => {
    FakeWorker.scripts = [{ init: 'manual' }]
    const p = mk()
    await until(() => FakeWorker.all.length === 1)
    const w = FakeWorker.all[0]!
    let settled = false
    const d = p.dispose().then(() => { settled = true })
    await Bun.sleep(20)
    expect({ settled, terminateCalls: w.terminateCalls, closed: w.closed }).toEqual({ settled: false, terminateCalls: 0, closed: false })
    w.ack()
    await d
    expect({ settled, terminateCalls: w.terminateCalls, closed: w.closed, terminatedUnacked: w.terminatedUnacked }).toEqual({ settled: true, terminateCalls: 1, closed: true, terminatedUnacked: false })
  })

  it('a second dispose() returns the same wait', async () => {
    const p = mk()
    await until(() => FakeWorker.all[0]?.acked === true)
    const a = p.dispose(), b = p.dispose()
    expect(b).toBe(a)
    await a
    expect(FakeWorker.all[0]!.closed).toBe(true)
  })
})
