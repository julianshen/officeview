import { describe, expect, it } from 'vitest'
import { AdmissionModel, DEFAULT_CAPS } from '../../scripts/stream-s0/t12-admission'

const MiB = 2 ** 20

describe('T12 admission model (§14 / §9.2)', () => {
  it('33rd request refused: 8 running + 32 queued, the next is refused', () => {
    const m = new AdmissionModel()
    for (let i = 0; i < 8; i++) expect(m.request('s', 1000).status).toBe('admitted')
    for (let i = 0; i < 32; i++) expect(m.request('s', 1000).status).toBe('queued')
    const r = m.request('s', 1000)
    expect(r).toEqual({ status: 'refused', reason: 'session-queue' })
    // another session is unaffected
    expect(m.request('t', 1000).status).toBe('admitted')
  })

  it('reservation is released on finish and the queue is pumped FIFO', () => {
    const m = new AdmissionModel()
    const jobs = Array.from({ length: 8 }, () => m.request('s', 10 * MiB))
    const q1 = m.request('s', 3 * MiB)
    const q2 = m.request('s', 4 * MiB)
    expect(m.reserved).toBe(80 * MiB)
    const first = jobs[0]!
    if (first.status !== 'admitted' || q1.status !== 'queued' || q2.status !== 'queued') throw new Error('setup')
    const started = m.finish(first.job)
    expect(started.map((j) => j.id)).toEqual([q1.job.id]) // only one slot freed
    expect(m.reserved).toBe(73 * MiB) // −10 MiB +3 MiB: the promoted job's own size
    expect(m.peakReserved).toBe(80 * MiB)
  })

  it('session transient cap (96 MiB) queues work even with free slots', () => {
    const m = new AdmissionModel()
    expect(m.request('s', 60 * MiB).status).toBe('admitted')
    expect(m.request('s', 40 * MiB).status).toBe('queued') // 100 MiB > 96 MiB
    expect(m.request('s', 97 * MiB)).toEqual({ status: 'refused', reason: 'too-large' })
  })

  it('process render reservation (512 MiB) is never exceeded across sessions', () => {
    const m = new AdmissionModel()
    for (let i = 0; i < 20; i++) m.request(`s${i}`, 90 * MiB)
    expect(m.reserved).toBeLessThanOrEqual(DEFAULT_CAPS.renderReservationBytes)
    expect(m.runningCount).toBe(5) // 5 × 90 = 450; a 6th would be 540
  })

  it('deadline exceeded job keeps its charge until its worker drains', () => {
    const m = new AdmissionModel()
    const r = m.request('s', 50 * MiB, 0)
    if (r.status !== 'admitted') throw new Error('setup')
    const retired = m.expire(DEFAULT_CAPS.renderDeadlineMs + 1)
    expect(retired.map((j) => j.id)).toEqual([r.job.id])
    expect(r.job.state).toBe('retiring')
    expect(m.reserved).toBe(50 * MiB) // not refunded at the deadline
    m.finish(r.job, 5000) // worker finally returns
    expect(m.reserved).toBe(0)
    expect(m.timeouts).toBe(1)
  })
})
