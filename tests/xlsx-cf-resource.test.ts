// @vitest-environment node
/**
 * Task 11 fix round 1, finding 2: the resource audit must observe real
 * allocations, reads, peaks and work — never disconnected counters or
 * literals. These tests pin the tracking budget, the pressure-sample gate,
 * and deliberate-leak detection before the audit functions use them.
 */
import { describe, expect, it } from 'vitest'
import { checkPressureSample, createTrackingBudget } from '../scripts/cf-refs/audit'

describe('tracking budget observes real holds', () => {
  it('tracks net zero and peak max-held across reserve/release', () => {
    const { budget, net, peak } = createTrackingBudget()
    expect(net()).toBe(0)
    expect(peak()).toBe(0)
    budget.reserve(100)
    budget.reserve(50)
    expect(net()).toBe(150)
    expect(peak()).toBe(150)
    budget.release(30)
    expect(net()).toBe(120)
    expect(peak()).toBe(150)
    budget.release(120)
    expect(net()).toBe(0)
    expect(peak()).toBe(150)
  })

  it('retain/commit moves through held and releases to zero', () => {
    const { budget, net, peak } = createTrackingBudget()
    let committed = false
    budget.retain?.([], 80, () => { committed = true })
    expect(committed).toBe(true)
    expect(net()).toBe(80)
    expect(peak()).toBe(80)
    budget.release(80)
    expect(net()).toBe(0)
  })
})

describe('pressure-sample gate fails leaks and tautologies', () => {
  it('accepts a balanced admitted sample with observed peak and results', () => {
    expect(() => checkPressureSample({
      name: 'ok', outcome: 'admitted', peakBytes: 120, netBytes: 0,
      sourceReads: 5, cellsWithResults: 3, wallMs: 1,
    }, { expectOutcome: 'admitted', requireWork: true })).not.toThrow()
  })

  it('rejects a leaked sample with net bytes held', () => {
    expect(() => checkPressureSample({
      name: 'leak', outcome: 'admitted', peakBytes: 120, netBytes: 64,
      sourceReads: 5, cellsWithResults: 3, wallMs: 1,
    }, { expectOutcome: 'admitted', requireWork: true })).toThrow(/netBytes/)
  })

  it('rejects an admitted sample that observed no work', () => {
    expect(() => checkPressureSample({
      name: 'tautology', outcome: 'admitted', peakBytes: 0, netBytes: 0,
      sourceReads: 0, cellsWithResults: 0, wallMs: 1,
    }, { expectOutcome: 'admitted', requireWork: true })).toThrow(/peakBytes/)
  })

  it('accepts an admitted no-match scope with observed work but zero outcomes', () => {
    expect(() => checkPressureSample({
      name: 'no-match', outcome: 'admitted', peakBytes: 1394142, netBytes: 0,
      sourceReads: 10, cellsWithResults: 0, wallMs: 1,
    }, { expectOutcome: 'admitted', requireWork: true })).not.toThrow()
  })

  it('accepts a refused sample with zero net and zero reads', () => {
    expect(() => checkPressureSample({
      name: 'refused', outcome: 'refused', peakBytes: 0, netBytes: 0,
      sourceReads: 0, cellsWithResults: 0, wallMs: 1, reason: 'estimator refused',
    }, { expectOutcome: 'refused' })).not.toThrow()
  })

  it('rejects a wrong-outcome sample', () => {
    expect(() => checkPressureSample({
      name: 'wrong', outcome: 'admitted', peakBytes: 10, netBytes: 0,
      sourceReads: 1, cellsWithResults: 1, wallMs: 1,
    }, { expectOutcome: 'refused' })).toThrow(/outcome/)
  })
})

describe('pressure functions observe real work (no literals)', () => {
  it('a leaked scope fails the gate; release passes', async () => {
    const { prepareCfTextScope } = await import('../src/xlsx/cf/text')
    const { numericSheet, cellIsRule } = await import('../scripts/cf-refs/audit')
    const sheet = numericSheet(10, cellIsRule(0, { firstRow: 0, lastRow: 9, firstCol: 0, lastCol: 0 }) as Record<string, unknown>)
    const { budget, net, peak } = createTrackingBudget()
    const scope = await prepareCfTextScope(sheet, {
      rects: [{ r0: 0, r1: 9, c0: 0, c1: 0 }],
      signal: new AbortController().signal,
      locale: 'en-US',
      budget,
    })
    expect(scope).toBeDefined()
    // Deliberate leak: no dispose. The gate must catch the held bytes.
    expect(net()).toBeGreaterThan(0)
    expect(() => checkPressureSample({
      name: 'leak', outcome: 'admitted', peakBytes: peak(), netBytes: net(),
      sourceReads: 10, cellsWithResults: 5, wallMs: 1,
    }, { expectOutcome: 'admitted' })).toThrow(/netBytes/)
    scope?.dispose()
    expect(net()).toBe(0)
    expect(() => checkPressureSample({
      name: 'balanced', outcome: 'admitted', peakBytes: peak(), netBytes: net(),
      sourceReads: 10, cellsWithResults: 5, wallMs: 1,
    }, { expectOutcome: 'admitted' })).not.toThrow()
  })

  it('whole-column refuses with zero enumeration', async () => {
    const { pressureWholeColumn } = await import('../scripts/cf-refs/audit')
    const s = await pressureWholeColumn()
    expect(s.outcome).toBe('refused')
    expect(s.sourceReads).toBe(0)
    expect(s.netBytes).toBe(0)
    expect(s.cellsWithResults).toBe(0)
    expect(() => checkPressureSample(s, { expectOutcome: 'refused' })).not.toThrow()
  })

  it('shared rules admit with observed visits and outcomes', async () => {
    const { pressureSharedRules } = await import('../scripts/cf-refs/audit')
    const s = await pressureSharedRules()
    expect(s.outcome).toBe('admitted')
    expect(s.sourceReads).toBeGreaterThan(0)
    expect(s.peakBytes).toBeGreaterThan(0)
    expect(s.cellsWithResults).toBeGreaterThan(0)
    expect(s.netBytes).toBe(0)
    expect(() => checkPressureSample(s, { expectOutcome: 'admitted' })).not.toThrow()
  })

  it('statistics visits every stored cell via the product counter', async () => {
    const { pressureStatisticsVisits } = await import('../scripts/cf-refs/audit')
    const s = await pressureStatisticsVisits()
    expect(s.outcome).toBe('admitted')
    expect(s.sourceReads).toBe(200)
    expect(s.peakBytes).toBeGreaterThan(0)
    expect(s.netBytes).toBe(0)
    expect(() => checkPressureSample(s, { expectOutcome: 'admitted' })).not.toThrow()
  }, 60000)

  it('paged cycles observe real allocation with stable retention and abort safety', async () => {
    const { pressureCycles } = await import('../scripts/cf-refs/audit')
    const s = await pressureCycles(3)
    expect(s.cycles).toBe(3)
    expect(s.reads).toBe(12)
    expect(s.maxPeakBytes).toBeGreaterThan(0)
    // Stability (not zero): retained shared definitions are legitimate and
    // constant; growth would throw inside pressureCycles.
  }, 120000)
})
