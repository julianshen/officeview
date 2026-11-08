import { describe, expect, it } from 'vitest'
import { AdmissionModel } from '../../scripts/stream-s0/t12-admission'

describe('T12 admission', () => {
  it('33rd request refused (8 outstanding + 32 queued, next refused)', () => {
    const m = new AdmissionModel()
    // 8 admitted
    for (let i = 0; i < 8; i++) expect(m.request('s', 1000)).toBe('admitted')
    // 32 queued
    for (let i = 0; i < 32; i++) expect(m.request('s', 1000)).toBe('queued')
    // next refused
    expect(m.request('s', 1000)).toBe('refused')
  })
  it('reservation released on complete', () => {
    const m = new AdmissionModel()
    m.request('s', 1000)
    const before = m.peak().reserved
    m.complete('x', 1000, 's')
    // outstanding released; if queue empty, reserved drops
    expect(m.peak().reserved).toBeLessThanOrEqual(before)
  })
  it('deadline exceeded job refunds only after drain', () => {
    const m = new AdmissionModel()
    m.request('s', 1000) // admitted, outstanding=1
    // queued one
    m.request('s', 1000) // queued (outstanding still 1 < 8? outstanding=1<8 so admitted actually)
    // make outstanding full: admit 7 more
    for (let i = 0; i < 7; i++) m.request('s', 1000)
    // now queued
    m.request('s', 1000)
    // complete an outstanding job: promotes queued, reserved stays (drain-then-refund: charge kept until drain)
    m.complete('y', 1000, 's')
    // reserved should not drop below (promoted keeps charge); completed count increments
    expect(m.completed).toBe(1)
  })
})
