// @vitest-environment node
/**
 * Ruling S3-I1-UNITS (owner, 2026-11-14): a CF preparation's per-call work is capped in deterministic work units
 * (`CF_LIMITS.activeUnits`), never in wall time. The clock drives the cooperative yield only, so the result, the
 * refusal and the unit count of every budget-driven path are the same under a frozen clock and under one that
 * advances 1,000 ms per read.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CF_LIMITS, createCfMemory, createCfWorkBudget, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { createCfService } from '../src/xlsx/cf/service'
import { createResidentCfSource } from '../src/xlsx/cf/statistics'
import { prepareCfPaint, residentSource } from '../src/xlsx/cf/render'
import { prepareCfTextScope } from '../src/xlsx/cf/text'
import { cfRulesVersion, prepareCfRemote } from '../src/xlsx/view/pagedSource'
import type { CfRule, CfStamp, SourceRect } from '../src/xlsx/cf/model'
import type { XlsxCell, XlsxSheet } from '../src/xlsx/types'
import type { FormatBudget } from '../src/xlsx/view/formatWorkspace'
import { barRule, barSheet, withBars } from './helpers/cf-bars'

const frozen = (): (() => number) => () => 0
const racing = (): (() => number) => { let t = 0; return () => (t += 1000) }
/** Every default budget (per-call, bounds, statistics) reads `performance.now()`: drive them all from `clock`. */
async function underClock<T>(clock: () => number, run: () => Promise<T>): Promise<T> {
  const spy = vi.spyOn(performance, 'now').mockImplementation(clock)
  try { return await run() } finally { spy.mockRestore() }
}
afterEach(() => { vi.restoreAllMocks() })

const ROWS = 60
const rect = (firstRow: number, lastRow: number): SourceRect => ({ firstRow, lastRow, firstCol: 0, lastCol: 0 })
const cells = (n: number): XlsxCell[] => Array.from({ length: n }, (_, r) => ({ ref: `A${r + 1}`, row: r, col: 0, value: (r * 37) % 101, styleIndex: 0, hasCachedValue: true }) as XlsxCell)
function rules(rows: number): CfRule[] {
  const all = rect(0, rows - 1), base = { dialect: 'base', formulas: [] } as const
  return [
    { id: 1, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [all], writes: new Set(['font.bold']), dxf: { 'font.bold': true },
      definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 50 }], ...base } },
    { id: 2, priority: 2, documentOrder: 1, stopIfTrue: false, ranges: [all], writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00' },
      definition: { kind: 'top10', rank: 5, percent: false, bottom: false, ...base } },
    { id: 3, priority: 3, documentOrder: 2, stopIfTrue: false, ranges: [all], writes: new Set(['colorScale', 'fill']),
      definition: { kind: 'colorScale', points: [{ bound: { type: 'min' }, argb: 0xfff8696b }, { bound: { type: 'max' }, argb: 0xff63be7b }], ...base } },
    { ...barRule(all, {}, 4) },
  ] as CfRule[]
}
function sheet(rows = ROWS): XlsxSheet {
  const all = cells(rows)
  return {
    name: 'U', sheetId: 'u', rows: all.map((cell, r) => ({ index: r, cells: [cell] })), cols: [{ min: 0, max: 0, widthChars: 200 / 7 }],
    merges: [], mergeRanges: [], conditionalFormatting: { diagnostics: [], rules: rules(rows) },
  } as unknown as XlsxSheet
}
const stampOf = (s: XlsxSheet): CfStamp => ({ owner: 'units', sheetId: s.sheetId!, workbookValueGen: 1, cfGen: cfRulesVersion(s.conditionalFormatting!), viewGen: 1, structureGen: 1, clockGen: 1 })
const formatBudget = { reserve: () => {}, release: () => {}, retain: (_t: unknown, _b: unknown, commit: () => void) => { commit() }, settle: () => {} } as unknown as FormatBudget

/** One `prepareCf` call through the service; returns every cell's outcome, or the refusal, and the per-call units. */
async function serviceRun(source: 'render' | 'statistics', budget?: () => CfWorkBudget) {
  const s = sheet(), stamp = stampOf(s), memory = createCfMemory()
  const service = createCfService(s.conditionalFormatting!, memory, { evidence: { allows: () => true }, ...(budget ? { budget } : {}) })
  const src = source === 'render' ? residentSource(s, stamp, memory, 'engine') : createResidentCfSource(s.rows.flatMap(r => r.cells), s.sheetId!, stamp, memory)
  try {
    const lease = await service.prepareCf({ stamp, sourceRects: [rect(0, ROWS - 1)], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, src)
    const out = Array.from({ length: ROWS }, (_, r) => lease.value.at(r, 0))
    lease.release()
    return JSON.stringify(out, (_k, v) => v instanceof Set ? [...v] : v)
  } catch (error) { return `refused ${(error as { code?: string }).code ?? (error as Error).message}` } finally { service.dispose() }
}

describe('S3-I1-UNITS: a per-call budget refuses on work units, never on wall time', () => {
  it('a budget whose clock advances 1,000 ms per read admits the same work as a frozen one', () => {
    for (const now of [frozen(), racing()]) {
      const budget = createCfWorkBudget({ now })
      for (let i = 0; i < 300; i++) {
        const spent = budget.spend(1)
        expect(spent).not.toBe('refused')
        if (spent === 'yield') budget.resume()
      }
      expect(budget.units).toBe(300)
    }
  })

  it('a call at exactly CF_LIMITS.activeUnits admits; one unit more refuses, under either clock', () => {
    expect(Number.isSafeInteger(CF_LIMITS.activeUnits) && CF_LIMITS.activeUnits > 0).toBe(true)
    for (const now of [frozen(), racing()]) {
      const budget = createCfWorkBudget({ now })
      let left = CF_LIMITS.activeUnits
      while (left > 0) {
        const n = Math.min(left, 128), spent = budget.spend(n)
        expect(spent).not.toBe('refused')
        if (spent === 'yield') budget.resume()
        left -= n
      }
      expect(budget.units).toBe(CF_LIMITS.activeUnits)
      expect(budget.spend(1)).toBe('refused')
      expect(budget.spend(0)).toBe('refused') // a refusal is final
    }
  })

  it('units spent while paused (before resume) still count, so the count cannot depend on when the clock yields', () => {
    const budget = createCfWorkBudget({ now: racing() })
    expect(budget.spend(5)).toBe('yield')
    expect(budget.spend(7)).toBe('yield')
    expect(budget.units).toBe(12)
  })

  it('a whole-scan budget keeps its own CF_LIMITS.workUnits cap, separate from the per-call cap', () => {
    const scan = createCfWorkBudget({ now: frozen(), scan: true })
    for (let n = 0; n < 7812; n++) { expect(scan.spend(128)).toBe('yield'); scan.resume() }
    expect(scan.spend(64)).toBe('continue'); expect(scan.units).toBe(CF_LIMITS.workUnits)
    expect(scan.spend(1)).toBe('refused')
    const capped = createCfWorkBudget({ now: frozen(), scan: true, maxUnits: 10 })
    expect(capped.spend(10)).toBe('continue'); expect(capped.spend(1)).toBe('refused')
  })
})

describe('S3-I1-UNITS: every budget-driven path gives the same result under a frozen and a racing clock', () => {
  it('prepareCf (default per-call, bounds and statistics budgets) over both resident sources', async () => {
    await withBars(async () => {
      for (const source of ['render', 'statistics'] as const) {
        const still = await underClock(frozen(), () => serviceRun(source))
        const fast = await underClock(racing(), () => serviceRun(source))
        expect(still.startsWith('refused'), source).toBe(false)
        expect(fast, source).toBe(still)
      }
    })
  })

  it('the per-call unit count is the same under both clocks', async () => {
    await withBars(async () => {
      for (const source of ['render', 'statistics'] as const) {
        const counts: number[] = []
        for (const now of [frozen(), racing()]) {
          let budget: CfWorkBudget | undefined
          expect(await serviceRun(source, () => (budget = createCfWorkBudget({ now })))).not.toMatch(/^refused/)
          counts.push(budget!.units)
        }
        expect(counts[0], source).toBeGreaterThan(0)
        expect(counts[1], source).toBe(counts[0])
      }
    })
  })

  it('the text path (multi-chunk), the paint path and the paged wire path', async () => {
    await withBars(async () => {
      const rows = 2_500
      const text = async () => {
        // A multi-chunk bar scope (1,000-cell chunks sharing one statistics scan), as the text-stats tests use.
        const scope = await prepareCfTextScope(barSheet(cells(rows).map(c => c.value as number)), { rects: [{ r0: 0, r1: rows - 1, c0: 0, c1: 0 }], signal: new AbortController().signal, locale: 'en-US', budget: formatBudget })
        if (!scope) return 'unavailable'
        const out = Array.from({ length: rows }, (_, r) => [scope.describe(r, 0), scope.effectiveStyle(r, 0, undefined)])
        scope.dispose()
        return JSON.stringify(out)
      }
      const paint = async () => {
        const batch = await prepareCfPaint(sheet())
        if (!batch) return 'unavailable'
        const out = Array.from({ length: ROWS }, (_, r) => [batch.barAt(r, 0), batch.layer.styleAt(r, 0, undefined)])
        batch.dispose()
        return JSON.stringify(out)
      }
      const wire = async () => {
        const s = sheet(), st = stampOf(s), memory = createCfMemory()
        const reply = await prepareCfRemote({ rects: [{ r0: 0, r1: ROWS - 1, c0: 0, c1: 0 }], stamp: { ...st }, grantBytes: 1 << 22 }, s.sheetId!, residentSource(s, st, memory, 'engine'), s.conditionalFormatting!, memory)
        return JSON.stringify(reply)
      }
      for (const [name, run] of [['text', text], ['paint', paint], ['wire', wire]] as const) {
        const still = await underClock(frozen(), run)
        const fast = await underClock(racing(), run)
        expect(still, name).not.toBe('unavailable')
        // Each path must actually paint or describe CF here, so the comparison is not two empty results.
        expect(still, name).toMatch(name === 'paint' ? /fillColor/ : /dataBar/)
        expect(fast, name).toBe(still)
      }
    })
  }, 60_000)
})

describe('S3-I1-UNITS: the default per-call budget refuses deterministically past CF_LIMITS.activeUnits', () => {
  // A colour scale checks its scope against every sheet merge: ranges x merges units (service blockScope), so the
  // merge count moves the call's units one for one. Merges far from the rule leave the result unchanged.
  const merges = (n: number): SourceRect[] => Array.from({ length: n }, (_, i) => ({ firstRow: 1000 + i * 2, lastRow: 1001 + i * 2, firstCol: 5, lastCol: 6 }))
  const scaleSheet = (): XlsxSheet => ({ ...sheet(10), conditionalFormatting: { diagnostics: [], rules: [rules(10)[2]!] } } as unknown as XlsxSheet)
  async function run(mergeCount: number, budget?: () => CfWorkBudget): Promise<string> {
    const s = scaleSheet(), stamp = stampOf(s), memory = createCfMemory()
    const service = createCfService(s.conditionalFormatting!, memory, { evidence: { allows: () => true }, merges: merges(mergeCount), ...(budget ? { budget } : {}) })
    try {
      const lease = await service.prepareCf({ stamp, sourceRects: [rect(0, 9)], signal: new AbortController().signal, grantBytes: CF_LIMITS.scratchBytes }, residentSource(s, stamp, memory, 'engine'))
      const out = JSON.stringify(Array.from({ length: 10 }, (_, r) => lease.value.at(r, 0)))
      lease.release()
      return out
    } catch (error) { return `refused ${(error as { code?: string }).code}: ${(error as Error).message}` } finally { service.dispose() }
  }

  it('the call at exactly the cap admits and one unit over refuses with the existing code', async () => {
    // Count the call's units with one merge on an unlimited scan budget, then size the merges to land on the cap.
    let counter: CfWorkBudget | undefined
    const reference = await run(1, () => (counter = createCfWorkBudget({ now: frozen(), scan: true, maxUnits: 10_000_000 })))
    expect(reference).not.toMatch(/^refused/)
    const atCap = 1 + CF_LIMITS.activeUnits - counter!.units
    expect(atCap).toBeGreaterThan(1)
    for (const clock of [frozen(), racing()]) {
      expect(await underClock(clock, () => run(atCap))).toBe(reference)
      expect(await underClock(clock, () => run(atCap + 1))).toBe('refused resource-limit: CF preparation exceeds its admitted batch')
    }
  }, 60_000)
})
