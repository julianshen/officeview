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

  it('R29: a send by reqId holds its slot until answered, then a whole span from the answer', () => {
    let t = 0
    const log = new UnitSendLog(() => t, 2, 1_000)
    log.record(1); log.record(2)
    t = 50_000 // unanswered sends never expire on time alone
    expect(log.room()).toBe(0)
    expect(log.waitMs()).toBe(Infinity)
    log.answered(2)
    expect(log.has(2)).toBe(true)
    expect(log.waitMs()).toBe(1_001)
    t = 51_001
    expect(log.room()).toBe(1)
    expect(log.has(2)).toBe(false)
    log.answered(1); log.answered(1) // only the first reply counts
    t = 52_001; expect(log.room()).toBe(1)
    t = 52_002; expect(log.room()).toBe(2)
  })
  it('R29: settleAll (resume) takes every unanswered send as received now', () => {
    let t = 0
    const log = new UnitSendLog(() => t, 1, 1_000)
    log.record(7)
    t = 9_000; log.settleAll()
    expect(log.waitMs()).toBe(1_001)
    t = 10_001; expect(log.room()).toBe(1)
  })
  it('R29: a table past its bound is treated as full', () => {
    const t = 0
    const log = new UnitSendLog(() => t, 1_000, 1_000)
    for (let i = 1; i <= 256; i++) log.record(i)
    expect(log.room(0, 10_000)).toBe(0)
  })
})
