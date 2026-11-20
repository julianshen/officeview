/**
 * The pool's worker-startup paths (Bun 1.3.10 start gate, `src/worker-start.ts`), on controllable stand-in workers:
 * a job waits for its worker's load before its deadline starts; a worker that dies while loading gets one fresh
 * replacement, and a second death, a dispose during the replacement's start or a stalled start fails the job with
 * `RenderFailed`, posting nothing and keeping no pending entry.
 */
import { afterEach, describe, expect, it } from 'bun:test'
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

afterEach(async () => { pool?.dispose(); pool = undefined; await FakeWorker.release(); FakeWorker.reset() })

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
    await expect(p.encode('s', job, 16)).rejects.toBeInstanceOf(RenderFailed)
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
    p.dispose()
    const fresh = FakeWorker.all[1]!
    expect(fresh.terminateCalls).toBe(0) // still loading: deferred
    fresh.ack()
    await expect(r).rejects.toThrow('pool disposed')
    await until(() => fresh.closed)
    expect(fresh).toMatchObject({ terminateCalls: 1, terminatedUnacked: false })
    expect(jobsSent(fresh)).toBe(0)
    expect(pending(p)).toBe(0)
  })

  it('a start that passes the load bound fails its jobs with RenderFailed and is never terminated while loading; once it answers, the slot works (R1)', async () => {
    FakeWorker.scripts = [{ init: 'manual' }]
    const p = mk({ workerLoadBoundMsForTests: 80 })
    await expect(p.encode('s', job, 16)).rejects.toThrow(/worker start timed out/)
    await expect(p.encode('s', job, 16)).rejects.toThrow(/worker start timed out/) // at once, not another bound
    const w = FakeWorker.all[0]!
    expect(w).toMatchObject({ acked: false, closed: false, terminateCalls: 0 })
    expect(jobsSent(w)).toBe(0)
    expect(pending(p)).toBe(0)
    w.ack()
    await until(() => w.acked)
    expect(await p.encode('s', job, 16)).toBeNull()
    expect(FakeWorker.all).toHaveLength(1)
  })
})
