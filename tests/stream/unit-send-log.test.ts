/** R28: the session's unit send log (pure, injected clock). */
import { describe, expect, it } from 'vitest'
import { UnitSendLog } from '../../src/stream/remote/unit-budget'

describe('UnitSendLog (pure)', () => {
  it('admits cap sends per span; a stamp frees its slot 1 ms after the span', () => {
    let t = 0
    const log = new UnitSendLog(() => t, 3, 1_000)
    log.record(); t = 50; log.record(); t = 100; log.record()
    expect(log.room()).toBe(0)
    expect(log.waitMs()).toBe(901) // the t = 0 stamp frees its slot at 1_001
    expect(log.room(0, 4)).toBe(1) // a higher cap (re-sends) still has room
    t = 1_001
    expect(log.room()).toBe(1)
    expect(log.room(1)).toBe(0) // one send already spoken for
    expect(log.waitMs(1)).toBe(50) // the t = 50 stamp frees the next slot at 1_051
  })
  it('a refusal the log cannot explain blocks every send for a whole span', () => {
    let t = 0
    const log = new UnitSendLog(() => t, 48, 62_000)
    log.record()
    t = 10; log.refused()
    expect(log.room()).toBe(0)
    expect(log.waitMs()).toBe(62_000)
    t = 62_009; expect(log.room()).toBe(0)
    t = 62_010; expect(log.room()).toBe(48)
  })
  it('a refusal the log explains (≥ the server limit) waits for log room and ≥ 2 s', () => {
    let t = 0
    const log = new UnitSendLog(() => t, 48, 62_000)
    for (let i = 0; i < 60; i++) log.record()
    t = 5; log.refused()
    expect(log.room(0, 61)).toBe(0) // blocked 2 s even for a cap with room
    t = 2_005; expect(log.room(0, 61)).toBe(1)
    expect(log.room()).toBe(0) // the log itself is full until the stamps expire
    expect(log.waitMs()).toBe(62_001 - 2_005)
  })
})
