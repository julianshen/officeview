/**
 * Held lifetime charges (spec §12.4): a charge with explicit owners ends at its
 * last holder, never twice, and stays a reachability charge until first held.
 */
import { describe, expect, it } from 'vitest'
import { MainMemoryLedger } from '../src/xlsx/view/mainLedger'

describe('held lifetime charges (MainMemoryLedger.chargeWhileHeld)', () => {
  it('ends at the last holder, never twice, and is charged again for a later holder', () => {
    const ledger = new MainMemoryLedger({ register() {} })
    const target = {}
    ledger.chargeWhileHeld('viewBuffers', [target], 100)
    const a = ledger.hold(target)!, b = ledger.hold(target)!
    a.release(); a.release()
    expect(ledger.bytes, 'a second release of the same hold is ignored').toBe(100)
    b.release()
    expect(ledger.bytes).toBe(0)
    expect(ledger.hasLifetime(target)).toBe(false)
    b.release()
    expect(ledger.bytes).toBe(0)
    const late = ledger.hold(target)!
    expect(ledger.bytes, 'a holder arriving after the last release holds the bytes again').toBe(100)
    expect(ledger.hasLifetime(target)).toBe(true)
    late.release()
    expect(ledger.bytes).toBe(0)
  })
  it('stays a reachability charge until first held; unknown targets have no hold', () => {
    const tokens: string[] = []
    const ledger = new MainMemoryLedger({ register(_target, token) { tokens.push(token) } })
    const target = {}
    ledger.chargeWhileHeld('geometry', [target], 64)
    expect(ledger.hold({})).toBeNull()
    expect(ledger.bytes).toBe(64)
    ledger.collected(tokens[0])
    expect(ledger.bytes, 'collection still ends an unheld charge').toBe(0)
  })
  it('a settle of a held target ends it once; its holders then release nothing more', () => {
    const ledger = new MainMemoryLedger({ register() {} })
    const target = {}
    ledger.chargeWhileHeld('viewBuffers', [target], 10)
    const hold = ledger.hold(target)!
    ledger.settle(target)
    expect(ledger.bytes).toBe(0)
    ledger.set('viewBuffers', 'other', 5)
    hold.release()
    expect(ledger.bytes).toBe(5)
  })
})

