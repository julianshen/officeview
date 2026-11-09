/**
 * T12 (Stage-0 scope call 2): a SYNTHETIC admission probe. There is no server
 * at Stage 0, so this is an in-process model of the pool's §14 / §9.2
 * reservation and queue rules, driven by 1,000 concurrent requests over 40
 * sessions plus 5 rogue sessions bursting 100 each. Admitted jobs run a REAL
 * @napi-rs/canvas paint of a corpus unit in one of 4 worker threads.
 *
 *   bun run s0:napi scripts/stream-s0/t12-admission.ts
 *
 * It proves the caps are enforceable and measures peak RSS and latency; it does
 * not test a server.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { resolveBackend, writeResult } from './results'

export { AdmissionModel, DEFAULT_CAPS } from '../../server/src/admission'
export type { Caps, Job, JobState, RequestResult } from '../../server/src/admission'
import { AdmissionModel, DEFAULT_CAPS, type Job } from '../../server/src/admission'

// ---------------------------------------------------------------------------
// Driver (real paints in worker threads)

interface WorkerReply { id: number; ms: number; backend: string; error?: string }

async function main(): Promise<number> {
  if (resolveBackend() !== '@napi-rs/canvas') {
    console.error('[t12] refused: run with `bun run s0:napi scripts/stream-s0/t12-admission.ts`')
    return 2
  }
  const WORKERS = 4
  const preload = join(fileURLToPath(new URL('.', import.meta.url)), 'napi-preload.ts')
  const workerUrl = new URL('./t12-worker.ts', import.meta.url).href
  const workers = Array.from({ length: WORKERS }, () => new Worker(workerUrl, { preload: [preload] } as WorkerOptions))
  const ready = await Promise.all(workers.map((w) => new Promise<{ backend: string; units: Array<{ key: string; bytes: number }> }>((resolve, reject) => {
    w.onmessage = (e) => resolve(e.data)
    w.onerror = (e) => reject(e)
  })))
  const backends = ready.map((r) => r.backend)
  if (!backends.every((b) => b === '@napi-rs/canvas')) throw new Error(`worker backends ${backends.join(',')}`)
  const units = ready[0]!.units
  const model = new AdmissionModel()
  const t0 = performance.now()
  const now = () => performance.now() - t0
  const latency = new Map<number, number>() // job id → request time
  const done: number[] = []
  const failures: string[] = []
  let rssPeak = process.memoryUsage().rss
  const sampler = setInterval(() => { rssPeak = Math.max(rssPeak, process.memoryUsage().rss) }, 25)

  const idle: Worker[] = [...workers]
  const waiting: Job[] = [] // started (charged) jobs not yet on a worker
  const unitOf = new Map<number, string>()
  const inflight = new Map<number, Job>()
  let resolveAll!: () => void
  const allDone = new Promise<void>((r) => { resolveAll = r })
  let outstanding = 0

  const dispatch = () => {
    while (idle.length > 0 && waiting.length > 0) {
      const job = waiting.shift()!
      const w = idle.pop()!
      inflight.set(job.id, job)
      ;(w as Worker & { jobId?: number }).jobId = job.id
      w.postMessage({ id: job.id, key: unitOf.get(job.id) })
    }
  }
  for (const w of workers) {
    w.onmessage = (e: MessageEvent<WorkerReply>) => {
      const reply = e.data
      const job = inflight.get(reply.id)!
      inflight.delete(reply.id)
      if (reply.error) failures.push(reply.error)
      done.push(now() - latency.get(job.id)!)
      for (const started of model.finish(job, now())) waiting.push(started)
      idle.push(w)
      outstanding--
      dispatch()
      if (outstanding === 0) resolveAll()
    }
  }
  const deadlineTimer = setInterval(() => model.expire(now()), 100)

  const sessions: Array<{ id: string; count: number }> = [
    ...Array.from({ length: 40 }, (_, i) => ({ id: `s${i}`, count: 25 })),
    ...Array.from({ length: 5 }, (_, i) => ({ id: `rogue${i}`, count: 100 })),
  ]
  const results = { admitted: 0, queued: 0, refused: 0 }
  // All requests arrive "concurrently": interleave sessions round-robin at t≈0.
  const maxCount = Math.max(...sessions.map((s) => s.count))
  let n = 0
  for (let k = 0; k < maxCount; k++) {
    for (const s of sessions) {
      if (k >= s.count) continue
      const unit = units[n++ % units.length]!
      const r = model.request(s.id, unit.bytes, now())
      results[r.status]++
      if (r.status === 'refused') continue
      latency.set(r.job.id, now())
      unitOf.set(r.job.id, unit.key)
      outstanding++
      if (r.status === 'admitted') waiting.push(r.job)
    }
  }
  // Queued jobs carry their unit key too (assigned above); start dispatching.
  dispatch()
  if (outstanding > 0) await allDone
  clearInterval(sampler)
  clearInterval(deadlineTimer)
  for (const w of workers) w.terminate()

  done.sort((a, b) => a - b)
  const pct = (p: number) => done[Math.min(done.length - 1, Math.floor(done.length * p))] ?? 0
  const caps = DEFAULT_CAPS
  const within = model.peakReserved <= caps.renderReservationBytes
    && model.peakSessionRunning <= caps.perSessionOutstanding
    && model.peakSessionQueued <= caps.perSessionQueue
    && model.peakSessionReserved <= caps.sessionTransientBytes
    && model.peakProcessQueued <= caps.processQueuedUnits
  const green = within && failures.length === 0
  writeResult('t12', {
    status: green ? 'GREEN' : 'RED',
    synthetic: true,
    caps,
    requests: { total: n, ...results, refusalsByReason: model.refusals },
    peak: {
      reservedBytes: model.peakReserved,
      sessionReservedBytes: model.peakSessionReserved,
      sessionRunning: model.peakSessionRunning,
      sessionQueued: model.peakSessionQueued,
      processQueued: model.peakProcessQueued,
      rssBytes: rssPeak,
    },
    timeouts: model.timeouts,
    paintFailures: failures.slice(0, 10),
    latencyMs: { p50: pct(0.5), p95: pct(0.95), max: done[done.length - 1] ?? 0, completed: done.length },
    workers: { count: WORKERS, backends },
    unitsUsed: units.length,
    note: 'Synthetic in-process model of §14/§9.2 admission (no server at Stage 0); paints are real @napi-rs/canvas renders in worker threads. Latency is request→finish wall clock and depends on machine load.',
  })
  console.log(`[t12] ${green ? 'GREEN' : 'RED'} requests ${n} admitted ${results.admitted} queued ${results.queued} refused ${results.refused} · peak reserved ${(model.peakReserved / 2 ** 20).toFixed(1)} MiB · session running≤${model.peakSessionRunning} queued≤${model.peakSessionQueued} · rss ${(rssPeak / 2 ** 20).toFixed(0)} MiB · p95 ${pct(0.95).toFixed(0)} ms · timeouts ${model.timeouts}`)
  return green ? 0 : 3
}

if (import.meta.main) process.exit(await main())
