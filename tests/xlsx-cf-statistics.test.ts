import { describe, expect, it } from 'vitest'
import { PackageError } from '../src/core/package-error'
import { parseXml } from '../src/core/xml'
import type {
  CfDefinition,
  CfLease,
  CfLiteral,
  CfRule,
  CfStamp,
  CfValuePage,
  SourceRect,
} from '../src/xlsx/cf/model'
import type { XlsxCell } from '../src/xlsx/types'
import { createCfMemory, createCfWorkBudget, CF_LIMITS, type CfWorkBudget } from '../src/xlsx/cf/limits'
import { parseCfSheet } from '../src/xlsx/cf/parse'
import {
  cfValueKey,
  createCfStatistics,
  createResidentCfSource,
  estimateStatisticsBytes,
  type CfRangeStatistics,
} from '../src/xlsx/cf/statistics'
import { evaluatePredicate, type CfEvidenceGate } from '../src/xlsx/cf/predicates'

const SHEET = 'sheet1'
const stamp = (gen: number, sheetId = SHEET): CfStamp => ({
  owner: 't5', sheetId, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1,
})
const rect = (firstRow: number, lastRow: number, firstCol = 0, lastCol = 0): SourceRect => ({ firstRow, lastRow, firstCol, lastCol })
const colName = (c: number): string => {
  let n = c + 1, s = ''
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26) }
  return s
}
function rule<D extends object>(definition: D, ranges: readonly SourceRect[] = [rect(0, 0)]): CfRule {
  return {
    id: 0, priority: 1, documentOrder: 0, ranges, stopIfTrue: false,
    writes: new Set(['ALL']),
    definition: { dialect: 'base', formulas: [], ...definition } as unknown as CfDefinition,
  }
}
const cell = (row: number, col: number, value: XlsxCell['value'], extra?: Partial<XlsxCell>): XlsxCell => ({
  ref: `${colName(col)}${row + 1}`, row, col, value, styleIndex: 0, ...extra,
})
const frozenBudget = (extra?: { signal?: AbortSignal }): CfWorkBudget => createCfWorkBudget({ now: () => 0, ...extra })

function makeCounter() {
  return { cells: 0, pages: 0, bounds: 0 }
}

/** Array-backed counting source: one page per requested rect. Stands in for the resident path. */
function arraySource(cells: XlsxCell[], counter = makeCounter(), sheetId = SHEET) {
  const memory = createCfMemory({})
  const inner = createResidentCfSource(cells, sheetId, stamp(1, sheetId), memory)
  return {
    memory,
    counter,
    source: {
      bounds: (id: string) => { counter.bounds++; return inner.bounds(id) },
      stamp: (id: string) => inner.stamp(id),
      pages: async function* (id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
        for await (const page of inner.pages(id, rects, budget, signal)) {
          counter.pages++
          counter.cells += page.value.values.length
          yield page
        }
      },
    },
  }
}

/** Generator source: values computed on the fly, chunked into small pages. Stands in for paged delivery. */
function generatorSource(
  count: number,
  valueAt: (i: number) => CfLiteral,
  counter = makeCounter(),
  opts: { pageSize?: number; abortAt?: number; aborter?: AbortController; memory?: ReturnType<typeof createCfMemory> } = {},
) {
  const memory = opts.memory ?? createCfMemory({})
  const pageSize = opts.pageSize ?? 64
  const literalBytes = (v: CfLiteral): number =>
    v.kind === 'string' ? 16 + v.value.length * 2 : v.kind === 'error' ? 16 + v.value.length * 2 : 8
  return {
    memory,
    counter,
    source: {
      bounds: (_id: string) => { counter.bounds++; return { firstRow: 0, lastRow: count - 1, firstCol: 0, lastCol: 0 } },
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
        for (const r of rects) {
          const rows: number[] = []
          for (let i = r.firstRow; i <= r.lastRow; i++) rows.push(i)
          for (let s = 0; s < rows.length; s += pageSize) {
            if (signal.aborted) throw new PackageError('cancelled', 't5 probe source cancelled')
            const chunk = rows.slice(s, s + pageSize)
            const values = chunk.map(valueAt)
            const spent = budget.spend(chunk.length)
            if (spent === 'refused') {
              if (signal.aborted) throw new PackageError('cancelled', 't5 probe source cancelled')
              throw new PackageError('resource-limit', 't5 probe source budget exhausted', undefined, 'memory-budget')
            }
            if (spent === 'yield') {
              await Promise.resolve()
              if (signal.aborted) throw new PackageError('cancelled', 't5 probe source cancelled')
              budget.resume()
            }
            counter.pages++
            counter.cells += chunk.length
            if (opts.abortAt !== undefined && counter.cells >= opts.abortAt) opts.aborter?.abort()
            const bytes = 128 + chunk.length * 4 + chunk.length * 2 + values.reduce((n, v) => n + literalBytes(v), 0)
            const hold = memory.reserve(bytes)
            if (hold.kind === 'unresolved') throw new PackageError('resource-limit', 't5 probe page refused', undefined, 'memory-budget')
            yield hold.adopt<CfValuePage>({ rows: new Uint32Array(chunk), cols: new Uint16Array(chunk.length), values }, bytes)
          }
        }
      },
    },
  }
}

describe('CF statistics preparation', () => {
  it('one scan serves 100 rules on the same 50-row A:A scope', async () => {
    // Exact mean 2 with manifest values: twenty-four 0s, two 2s, twenty-four 4s.
    const pattern = [...Array.from({ length: 24 }, () => 0), 2, 2, ...Array.from({ length: 24 }, () => 4)]
    const cells: XlsxCell[] = pattern.map((v, r) => cell(r, 0, v))
    const counter = makeCounter()
    const { memory, source } = arraySource(cells, counter)
    const stats = createCfStatistics(memory)
    const scope = [rect(0, 49)]
    const out = await stats.prepareStatistics(Array.from({ length: 100 }, () => scope), stamp(1), source, frozenBudget())
    expect(out).toHaveLength(100)
    // One scan only: every stored cell visited once, plus bounded metadata work.
    expect(counter.cells).toBe(50)
    expect(counter.pages).toBeLessThanOrEqual(2)
    const first = out[0]!
    expect(first.kind).toBe('ready')
    if (first.kind !== 'ready') throw new Error('expected ready statistics')
    const statsValue = first.value.value
    expect(statsValue.numericCount).toBe(50)
    expect(statsValue.sum).toBe(100)
    expect(statsValue.mean).toBe(2)
    expect(statsValue.min).toBe(0)
    expect(statsValue.max).toBe(4)
    const lease = first.value
    lease.release()
    for (const outcome of out.slice(1)) if (outcome.kind === 'ready') outcome.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('disjoint and overlapping scopes scan separately without cross-scope dedup', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 75; r++) cells.push(cell(r, 0, r))
    const counter = makeCounter()
    const { source } = arraySource(cells, counter)
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const disjoint = await stats.prepareStatistics([[rect(0, 49)], [rect(50, 74)]], stamp(1), source, frozenBudget())
    expect(counter.cells).toBe(75)
    for (const o of disjoint) { expect(o.kind).toBe('ready'); if (o.kind === 'ready') o.value.release() }
    const overlapping = await stats.prepareStatistics([[rect(0, 49)], [rect(24, 74)]], stamp(2), source, frozenBudget())
    // Fresh generation: both scopes scan fully (50 + 51), proving overlapping scopes share nothing.
    expect(counter.cells).toBe(75 + 50 + 51)
    for (const o of overlapping) { expect(o.kind).toBe('ready'); if (o.kind === 'ready') o.value.release() }
    expect(memory.currentBytes).toBe(0)
  })

  it('hidden rows follow the stated policy and key separately', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 10; r++) cells.push(cell(r, 0, 1))
    const { source } = arraySource(cells)
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const hidden = [rect(0, 4)]
    const included = await stats.prepareStatistics([[rect(0, 9)]], stamp(1), source, frozenBudget(), { hidden: 'include', hiddenRows: hidden })
    const excluded = await stats.prepareStatistics([[rect(0, 9)]], stamp(1), source, frozenBudget(), { hidden: 'exclude', hiddenRows: hidden })
    expect(included[0]!.kind).toBe('ready')
    expect(excluded[0]!.kind).toBe('ready')
    if (included[0]!.kind !== 'ready' || excluded[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    // Different policies are different cache entries even over identical cells.
    expect(included[0]!.value.value.numericCount).toBe(10)
    expect(excluded[0]!.value.value.numericCount).toBe(5)
    expect(stats.size).toBe(2)
    included[0]!.value.release()
    excluded[0]!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('whole-column scope clamps to populated bounds instead of scanning a million rows', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 50; r++) cells.push(cell(r, 0, 2))
    const counter = makeCounter()
    const { source } = arraySource(cells, counter)
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const out = await stats.prepareStatistics([[rect(0, 1048575)]], stamp(1), source, frozenBudget())
    expect(out[0]!.kind).toBe('ready')
    expect(counter.cells).toBeLessThanOrEqual(50)
    if (out[0]!.kind === 'ready') {
      expect(out[0]!.value.value.numericCount).toBe(50)
      out[0]!.value.release()
    }
    expect(memory.currentBytes).toBe(0)
  })

  it('unknown bounds withhold instead of returning provisional statistics', async () => {
    const counter = makeCounter()
    const src = generatorSource(10, () => ({ kind: 'number', value: 1 }), counter)
    const boundless = { ...src.source, bounds: (_id: string) => { counter.bounds++; return undefined } }
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const out = await stats.prepareStatistics([[rect(0, 9)]], stamp(1), boundless, frozenBudget())
    expect(out[0]).toMatchObject({ kind: 'unresolved', pendingSheetIds: [SHEET] })
    expect(counter.cells).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('reservation precedes typed arrays: requiredBytes-1 refuses before any builder runs', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 200; r++) cells.push(cell(r, 0, r % 7))
    const counter = makeCounter()
    const probed = arraySource(cells, counter)
    const required = estimateStatisticsBytes(200, false)
    expect(required).toBeGreaterThan(200 * 8)
    const tight = createCfMemory({ limitBytes: required - 1 })
    const stats = createCfStatistics(tight)
    const out = await stats.prepareStatistics([[rect(0, 199)]], stamp(1), probed.source, frozenBudget())
    expect(out[0]!.kind).toBe('unresolved')
    // Refusal happened before any array was built or any cell visited.
    expect(counter.cells).toBe(0)
    expect(tight.currentBytes).toBe(0)
  })

  it('current bytes include pinned pages plus statistics workspace', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 50; r++) cells.push(cell(r, 0, [0, 2, 4][r % 3]!))
    const memory = createCfMemory({})
    // Pins and stats share one account, as in a Task 6 evaluation pass.
    const shared = createResidentCfSource(cells, SHEET, stamp(1), memory)
    const pins: CfLease<CfValuePage>[] = []
    for await (const page of shared.pages(SHEET, [rect(0, 49)], frozenBudget(), new AbortController().signal)) pins.push(page)
    const pinned = memory.currentBytes
    expect(pinned).toBeGreaterThan(0)
    const stats = createCfStatistics(memory)
    const out = await stats.prepareStatistics([[rect(0, 49)]], stamp(1), shared, frozenBudget())
    expect(out[0]!.kind).toBe('ready')
    if (out[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    const lease = out[0]!.value
    expect(lease.bytes).toBeGreaterThan(0)
    // Exact accounting: pins plus exactly one statistics lease (the cache holds no separate charge).
    expect(memory.currentBytes).toBe(pinned + lease.bytes)
    lease.release()
    for (const p of pins) p.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('generations and sheets do not alias cache keys', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 10; r++) cells.push(cell(r, 0, 1))
    const counter = makeCounter()
    const { memory, source } = arraySource(cells, counter)
    // Same values served under a different sheet identity: no aliasing.
    const other = arraySource(cells, counter, 'other')
    const stats = createCfStatistics(memory)
    const first = await stats.prepareStatistics([[rect(0, 9)]], stamp(1), source, frozenBudget())
    const second = await stats.prepareStatistics([[rect(0, 9)]], stamp(1), source, frozenBudget())
    expect(counter.cells).toBe(10)
    const third = await stats.prepareStatistics([[rect(0, 9)]], stamp(2), source, frozenBudget())
    expect(counter.cells).toBe(20)
    const fourth = await stats.prepareStatistics([[rect(0, 9)]], stamp(1, 'other'), other.source, frozenBudget())
    expect(counter.cells).toBe(30)
    for (const out of [first, second, third, fourth]) for (const o of out) if (o.kind === 'ready') o.value.release()
    expect(other.memory.currentBytes).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('warm cache hits loop zero times', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 30; r++) cells.push(cell(r, 0, r))
    const counter = makeCounter()
    const { source } = arraySource(cells, counter)
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const one = await stats.prepareStatistics([[rect(0, 29)]], stamp(1), source, frozenBudget())
    expect(counter.cells).toBe(30)
    const two = await stats.prepareStatistics([[rect(0, 29)]], stamp(1), source, frozenBudget())
    expect(counter.cells).toBe(30)
    for (const out of [one, two]) for (const o of out) if (o.kind === 'ready') o.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('evicted but leased statistics remain charged', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 10; r++) cells.push(cell(r, 0, r))
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory, { maxEntries: 4 })
    const first = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), source, frozenBudget())
    expect(first[0]!.kind).toBe('ready')
    if (first[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    const held = first[0]!.value
    const heldBytes = held.bytes
    expect(heldBytes).toBeGreaterThan(0)
    // Hold every filler lease so six live keys force eviction of the oldest.
    const fillers: CfLease<CfRangeStatistics>[] = []
    for (let c = 1; c <= 5; c++) {
      const out = await stats.prepareStatistics([[rect(c, c)]], stamp(1), source, frozenBudget())
      expect(out[0]!.kind).toBe('ready')
      if (out[0]!.kind === 'ready') fillers.push(out[0]!.value)
    }
    expect(stats.size).toBeLessThanOrEqual(4)
    // The evicted entry's lease still owns its bytes.
    expect(memory.currentBytes).toBeGreaterThanOrEqual(heldBytes)
    held.release()
    for (const f of fillers) f.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('dropping every lease returns the account to zero', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 10; r++) cells.push(cell(r, 0, r))
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const out = await stats.prepareStatistics([[rect(0, 9)], [rect(0, 4)]], stamp(1), source, frozenBudget())
    expect(memory.currentBytes).toBeGreaterThan(0)
    for (const o of out) if (o.kind === 'ready') o.value.release()
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('cancellation during a scan rejects without caching anything', async () => {
    const controller = new AbortController()
    const counter = makeCounter()
    const src = generatorSource(50000, (i) => ({ kind: 'number', value: i % 11 }), counter, {
      pageSize: 64, abortAt: 500, aborter: controller,
    })
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    await expect(
      stats.prepareStatistics([[rect(0, 49999)]], stamp(1), src.source, frozenBudget(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: 'cancelled' })
    // Observed promptly by the checkpoint, long before the 50k span completed.
    expect(counter.cells).toBeLessThan(50000)
    expect(counter.cells).toBeGreaterThan(0)
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('shared work exhaustion refuses past the limit without completing the scan', async () => {
    const span = 550000
    const counter = makeCounter()
    const src = generatorSource(span, (i) => ({ kind: 'number', value: i % 7 }), counter, { pageSize: 1024 })
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const out = await stats.prepareStatistics([[rect(0, span - 1)]], stamp(1), src.source, budget)
    expect(out[0]).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(budget.units).toBeLessThanOrEqual(CF_LIMITS.workUnits + 127)
    // Refused mid-scan after real work (about half the span at ~2 units/cell), not up front.
    expect(counter.cells).toBeLessThan(span)
    expect(counter.cells).toBeGreaterThan(100000)
    expect(memory.currentBytes).toBe(0)
    // Nothing was cached: a fresh budget refuses again instead of hitting.
    const again = await stats.prepareStatistics([[rect(0, span - 1)]], stamp(1), src.source, frozenBudget())
    expect(again[0]).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(stats.size).toBe(0)
  })

  it('ten thousand whole-column rules refuse in the parser without scanning', async () => {
    const rules = Array.from({ length: 10000 }, (_, i) => `<cfRule type="expression" priority="${i + 1}"/>`).join('')
    const ws = parseXml(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A:A">${rules}</conditionalFormatting></worksheet>`)
    const memory = createCfMemory({})
    const lease = parseCfSheet(ws, { colors: {}, dxfNodes: [] }, memory, {})
    expect(lease.value.rules).toHaveLength(0)
    expect(lease.value.unresolvedSheet).toMatchObject({ writes: 'ALL', reason: 'cf-resource-limit' })
    const before = memory.currentBytes
    lease.release()
    expect(memory.currentBytes).toBeLessThanOrEqual(before)
    expect(memory.currentBytes).toBe(0)
  })

  it('prepared statistics drive predicate decisions at oracle coordinates', async () => {
    // Stage1 above-average matrix: values 0,2,4 with mean 2.
    const cells: XlsxCell[] = [cell(1, 1, 0), cell(2, 1, 2), cell(3, 1, 4)]
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const gate: CfEvidenceGate = { allows: () => false }
    const prepared = await stats.prepareStatistics([[rect(1, 3, 1, 1)]], stamp(1), source, frozenBudget())
    expect(prepared[0]!.kind).toBe('ready')
    if (prepared[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    const st = prepared[0]!.value.value
    const above = (above: boolean, equal: boolean) =>
      rule({ kind: 'aboveAverage', above, equal }, [rect(1, 3, 1, 1)])
    // B3 (2 is not strictly above mean 2), B4 (4 > 2), C2 (0 < 2): manifest stage1-above-average-stddev.
    expect(evaluatePredicate(above(true, false), { row: 2, col: 1 }, { kind: 'number', value: 2 }, st, gate))
      .toMatchObject({ kind: 'no-match' })
    expect(evaluatePredicate(above(true, false), { row: 3, col: 1 }, { kind: 'number', value: 4 }, st, gate))
      .toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(above(false, false), { row: 1, col: 1 }, { kind: 'number', value: 0 }, st, gate))
      .toMatchObject({ kind: 'ready', value: true })
    prepared[0]!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('cfValueKey separates kinds and folds text case', () => {
    expect(cfValueKey({ kind: 'number', value: 1 })).toBe(cfValueKey({ kind: 'number', value: 1 }))
    expect(cfValueKey({ kind: 'number', value: 1 })).not.toBe(cfValueKey({ kind: 'string', value: '1' }))
    expect(cfValueKey({ kind: 'string', value: 'A' })).toBe(cfValueKey({ kind: 'string', value: 'a' }))
    expect(cfValueKey({ kind: 'string', value: 'a' })).not.toBe(cfValueKey({ kind: 'string', value: 'á' }))
    expect(cfValueKey({ kind: 'boolean', value: true })).not.toBe(cfValueKey({ kind: 'number', value: 1 }))
    expect(cfValueKey({ kind: 'blank' })).toBeUndefined()
    expect(cfValueKey({ kind: 'error', value: '#N/A' })).toBeUndefined()
  })
})
