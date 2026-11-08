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
import { measurePayloadBytes } from '../src/xlsx/view/payloadBytes'
import { parseCfSheet } from '../src/xlsx/cf/parse'
import {
  cfValueKey,
  createCfStatistics,
  createResidentCfSource,
  estimateStatisticsInitialBytes,
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
function arraySource(
  cells: XlsxCell[],
  counter = makeCounter(),
  sheetId = SHEET,
  live?: { gen: number; owner: string; structureGen: number },
) {
  const memory = createCfMemory({})
  const liveStamp = (): CfStamp => ({
    owner: live?.owner ?? 't5', sheetId, workbookValueGen: live?.gen ?? 1,
    cfGen: 1, viewGen: 1, structureGen: live?.structureGen ?? 1, clockGen: 1,
  })
  const inner = createResidentCfSource(cells, sheetId, liveStamp(), memory)
  return {
    memory,
    counter,
    source: {
      bounds: (id: string) => { counter.bounds++; return inner.bounds(id) },
      stamp: (_id: string) => liveStamp(),
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
      bounds: async (_id: string) => { counter.bounds++; return { firstRow: 0, lastRow: count - 1, firstCol: 0, lastCol: 0 } },
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
    const live = { gen: 1, owner: 't5', structureGen: 1 }
    const { memory, source } = arraySource(cells, counter, SHEET, live)
    const stats = createCfStatistics(memory)
    const req = (gen: number): CfStamp => ({
      owner: 't5', sheetId: SHEET, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1,
    })
    const disjoint = await stats.prepareStatistics([[rect(0, 49)], [rect(50, 74)]], req(1), source, frozenBudget())
    expect(counter.cells).toBe(75)
    for (const o of disjoint) { expect(o.kind).toBe('ready'); if (o.kind === 'ready') o.value.release() }
    // Fresh generation on the live source: both scopes scan fully (50 + 51).
    live.gen = 2
    const overlapping = await stats.prepareStatistics([[rect(0, 49)], [rect(24, 74)]], req(2), source, frozenBudget())
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
    expect(included[0]!.value.value.hasHiddenRows).toBe(true)
    expect(excluded[0]!.value.value.hasHiddenRows).toBe(true)
    expect(stats.size).toBe(2)
    included[0]!.value.release()
    excluded[0]!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('identifies when statistics scopes do not intersect hidden rows', async () => {
    const cells: XlsxCell[] = []
    for (let r = 0; r < 10; r++) cells.push(cell(r, 0, 1))
    const { source } = arraySource(cells)
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const hidden = [rect(0, 4)]
    const disjoint = await stats.prepareStatistics([[rect(5, 9)]], stamp(1), source, frozenBudget(), { hidden: 'include', hiddenRows: hidden })
    expect(disjoint[0]!.kind).toBe('ready')
    if (disjoint[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(disjoint[0]!.value.value.hasHiddenRows).toBe(false)
    disjoint[0]!.value.release()
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
    const boundless = { ...src.source, bounds: async (_id: string) => { counter.bounds++; return undefined } }
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
    // The initial workspace chunk (span-scaled) is admitted before anything is built.
    const required = estimateStatisticsInitialBytes(200)
    expect(required).toBe(256 + 200 * 16 + 256)
    const tight = createCfMemory({ limitBytes: required - 1 })
    const stats = createCfStatistics(tight)
    const out = await stats.prepareStatistics([[rect(0, 199)]], stamp(1), probed.source, frozenBudget())
    expect(out[0]!.kind).toBe('unresolved')
    // Refusal happened before any scan or array was built.
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
    const live = { gen: 1, owner: 't5', structureGen: 1 }
    const { memory, source } = arraySource(cells, counter, SHEET, live)
    // The second sheet serves identical values under its own identity.
    const otherLive = { gen: 1, owner: 't5', structureGen: 1 }
    const other = arraySource(cells, counter, 'other', otherLive)
    const stats = createCfStatistics(memory)
    const req = (gen: number, sheetId = SHEET): CfStamp => ({
      owner: 't5', sheetId, workbookValueGen: gen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1,
    })
    const first = await stats.prepareStatistics([[rect(0, 9)]], req(1), source, frozenBudget())
    const second = await stats.prepareStatistics([[rect(0, 9)]], req(1), source, frozenBudget())
    expect(counter.cells).toBe(10)
    // A new value generation rescans instead of aliasing the old statistics.
    live.gen = 2
    const third = await stats.prepareStatistics([[rect(0, 9)]], req(2), source, frozenBudget())
    expect(counter.cells).toBe(20)
    // A different sheet identity rescans even with identical values.
    const fourth = await stats.prepareStatistics([[rect(0, 9)]], req(1, 'other'), other.source, frozenBudget())
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

  it('shared work refuses past the budget without completing the later scope', async () => {
    // Two distinct admitted scopes share one budget: shuffled values force full
    // comparison counts, so the first completes and the second is refused mid-sort.
    const cells: XlsxCell[] = []
    for (let r = 0; r < 20000; r++) cells.push(cell(r, 0, (r * 7919) % 20000))
    const counter = makeCounter()
    const { memory, source } = arraySource(cells, counter)
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const out = await stats.prepareStatistics([[rect(0, 9999)], [rect(10000, 19999)]], stamp(1), source, budget)
    expect(out).toHaveLength(2)
    expect(out[0]!.kind).toBe('ready')
    expect(out[1]).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(budget.units).toBeLessThanOrEqual(CF_LIMITS.workUnits + 127)
    expect(budget.units).toBeGreaterThan(800000)
    expect(stats.size).toBe(1)
    for (const o of out) if (o.kind === 'ready') o.value.release()
    expect(memory.currentBytes).toBe(0)
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

  it('frequencies count every kind across rectangles and deduplicate stored cells', async () => {
    // Reviewer probe: [1,1,2] must count n:1 twice (not an empty map).
    const numCells = [cell(0, 0, 1), cell(1, 0, 1), cell(2, 0, 2)]
    const { memory: m1, source: s1 } = arraySource(numCells)
    const stats1 = createCfStatistics(m1)
    const [numOut] = await stats1.prepareStatistics([[rect(0, 2)]], stamp(1), s1, frozenBudget())
    expect(numOut!.kind).toBe('ready')
    if (numOut!.kind !== 'ready') throw new Error('expected ready statistics')
    expect([...numOut!.value.value.frequencies.entries()].sort()).toEqual([['n:1', 2], ['n:2', 1]])
    expect(numOut!.value.value.storedCells).toBe(3)
    const allowGate: CfEvidenceGate = { allows: () => true }
    const dup = rule({ kind: 'predicate', type: 'duplicateValues' })
    expect(evaluatePredicate(dup, { row: 0, col: 0 }, { kind: 'number', value: 1 }, numOut!.value.value, allowGate))
      .toMatchObject({ kind: 'ready', value: true })
    expect(evaluatePredicate(dup, { row: 2, col: 0 }, { kind: 'number', value: 2 }, numOut!.value.value, allowGate))
      .toMatchObject({ kind: 'no-match' })
    numOut!.value.release()
    // Review-final-r1 note: Excel counts cell values across each overlapping rectangle
    // in the sqref list (e.g. A1:A2 A1 counts 'a' twice, so 'a' matches duplicateValues).
    const strCells = [cell(0, 0, 'a'), cell(1, 0, 'b')]
    const { memory: m2, source: s2 } = arraySource(strCells)
    const stats2 = createCfStatistics(m2)
    const [strOut] = await stats2.prepareStatistics([[rect(0, 1), rect(0, 0)]], stamp(1), s2, frozenBudget())
    expect(strOut!.kind).toBe('ready')
    if (strOut!.kind !== 'ready') throw new Error('expected ready statistics')
    expect([...strOut!.value.value.frequencies.entries()].sort()).toEqual([['s:a', 2], ['s:b', 1]])
    expect(evaluatePredicate(dup, { row: 0, col: 0 }, { kind: 'string', value: 'a' }, strOut!.value.value, allowGate))
      .toMatchObject({ kind: 'ready', value: true })
    strOut!.value.release()
    expect(m1.currentBytes).toBe(0)
    expect(m2.currentBytes).toBe(0)
  })

  it('translated small spreads keep full precision (Neumaier + two-pass variance)', async () => {
    const cells = [cell(0, 0, 1e12), cell(1, 0, 1e12 + 1), cell(2, 0, 1e12 + 2)]
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const [out] = await stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    const st = out!.value.value
    expect(st.mean).toBe(1000000000001)
    expect(st.stddevPop).toBeCloseTo(Math.sqrt(2 / 3), 12)
    out!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('cancellation sums exactly and overflow withholds instead of inventing means', async () => {
    const cells = [cell(0, 0, 1e16), cell(1, 0, 1), cell(2, 0, -1e16)]
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const [out] = await stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(out!.value.value.sum).toBe(1)
    expect(out!.value.value.mean).toBeCloseTo(1 / 3, 15)
    out!.value.release()
    // Overflow: [1e308, 1e308] sums to Infinity, so the mean is withheld;
    // ranking does not need a mean and still works.
    const big = [cell(0, 0, 1e308), cell(1, 0, 1e308)]
    const { memory: m2, source: s2 } = arraySource(big)
    const stats2 = createCfStatistics(m2)
    const [bigOut] = await stats2.prepareStatistics([[rect(0, 1)]], stamp(1), s2, frozenBudget())
    expect(bigOut!.kind).toBe('ready')
    if (bigOut!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(bigOut!.value.value.mean).toBeUndefined()
    const allowGate: CfEvidenceGate = { allows: () => true }
    const above = rule({ kind: 'aboveAverage', above: true, equal: false }, [rect(0, 1)])
    expect(evaluatePredicate(above, { row: 0, col: 0 }, { kind: 'number', value: 1e308 }, bigOut!.value.value, allowGate))
      .toMatchObject({ kind: 'no-match' })
    const top = rule({ kind: 'top10', rank: 1, bottom: false, percent: false }, [rect(0, 1)])
    expect(evaluatePredicate(top, { row: 0, col: 0 }, { kind: 'number', value: 1e308 }, bigOut!.value.value, allowGate))
      .toMatchObject({ kind: 'ready', value: true })
    bigOut!.value.release()
    expect(memory.currentBytes).toBe(0)
    expect(m2.currentBytes).toBe(0)
  })

  it('string storage is admitted before the key is built', async () => {
    // Reviewer probe: on a 300 B account, a 10,000-char key must be refused
    // before toLowerCase runs, with the admission already covering it.
    const memory = createCfMemory({ limitBytes: 300 })
    const stats = createCfStatistics(memory)
    const lower = String.prototype.toLowerCase
    const calls: Array<{ length: number; charged: number }> = []
    String.prototype.toLowerCase = function (this: string) {
      calls.push({ length: this.length, charged: memory.currentBytes })
      return lower.call(this)
    }
    const src = {
      bounds: async (_id: string) => rect(0, 0),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'string', value: 'A'.repeat(10000) } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    try {
      const [out] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), src, frozenBudget())
      expect(out!.kind).toBe('unresolved')
    } finally {
      String.prototype.toLowerCase = lower
    }
    // No key was ever built unadmitted: every call observed admitted room.
    for (const call of calls) expect(call.charged).toBeGreaterThan(call.length)
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('the adopted candidate survives the async handoff with its charge', async () => {
    // Reviewer probe: a microtask observing mid-handoff state must never see
    // 0 charged bytes while the result is still alive. Only post-adoption
    // releases are observed: transient workspace holds (union sweep, pages)
    // are acquired and released before any result exists.
    const memory = createCfMemory({})
    const reserve = memory.reserve.bind(memory)
    const observed: number[] = []
    let adopted = false
    memory.reserve = ((bytes: number) => {
      const hold = reserve(bytes)
      if (hold.kind === 'reserved') {
        const adopt = hold.adopt.bind(hold)
        hold.adopt = ((value: unknown, retained?: number) => {
          adopted = true
          return adopt(value as never, retained)
        }) as typeof hold.adopt
        const release = hold.release.bind(hold)
        hold.release = () => {
          release()
          if (adopted) void queueMicrotask(() => observed.push(memory.currentBytes))
        }
      }
      return hold
    }) as typeof memory.reserve
    const stats = createCfStatistics(memory)
    const cells = [cell(0, 0, 1), cell(1, 0, 2), cell(2, 0, 3)]
    const probed = arraySource(cells)
    const [out] = await stats.prepareStatistics([[rect(0, 2)]], stamp(1), probed.source, frozenBudget())
    expect(out!.kind).toBe('ready')
    await Promise.resolve()
    await new Promise(resolve => setTimeout(resolve, 0))
    // The retained result is adopted (charged) before any release runs.
    expect(memory.currentBytes).toBeGreaterThan(0)
    expect(observed.every(v => v > 0)).toBe(true)
    if (out!.kind === 'ready') out!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('sparse scopes do not size numeric arrays by dense rectangle area', async () => {
    // Reviewer probe: 600,000-row span with 2 stored numbers must stay ready
    // with a small charge, far below the 8 MiB session cap.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const src = {
      bounds: async (_id: string) => rect(0, 599999),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0, 599999]), cols: new Uint16Array([0, 0]), values: [{ kind: 'number', value: 1 }, { kind: 'number', value: 2 }] as const },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    const [out] = await stats.prepareStatistics([[rect(0, 599999)]], stamp(1), src, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(out!.value.value.numericCount).toBe(2)
    expect(out!.value.bytes).toBeLessThan(1024 * 1024)
    expect(memory.peakBytes).toBeLessThan(8 * 1024 ** 2)
    out!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('statistics exceeding the session cap withhold instead of publishing', async () => {
    // Over the 8 MiB dedicated cap: refused. Just under it: ready.
    const bigString = (i: number): CfLiteral => ({ kind: 'string', value: `k${i}-` + 'x'.repeat(4000) })
    const over = generatorSource(2500, bigString, makeCounter(), { pageSize: 100 })
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const [refused] = await stats.prepareStatistics([[rect(0, 2499)]], stamp(1), over.source, frozenBudget())
    expect(refused!).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
    const under = generatorSource(800, bigString, makeCounter(), { pageSize: 100 })
    const [ready] = await stats.prepareStatistics([[rect(0, 799)]], stamp(1), under.source, frozenBudget())
    expect(ready!.kind).toBe('ready')
    if (ready!.kind === 'ready') {
      expect(ready!.value.bytes).toBeLessThanOrEqual(8 * 1024 ** 2)
      ready!.value.release()
    }
    expect(memory.currentBytes).toBe(0)
  })

  it('one shared charge serves 100 identical scopes instead of 100 adoptions', async () => {
    // Reviewer probe shape (many identical scopes, roomy account): a shared
    // scan and sort serves every rule from one charge instead of refusing most.
    const cells = Array.from({ length: 10000 }, (_, i) => cell(i, 0, i))
    const { source } = arraySource(cells)
    const account = createCfMemory({ limitBytes: 16 * 1024 ** 2 })
    const stats = createCfStatistics(account)
    const scopes = Array.from({ length: 100 }, () => [rect(0, 9999)])
    const results = await stats.prepareStatistics(scopes, stamp(1), source, frozenBudget())
    expect(results.filter(r => r.kind === 'ready')).toHaveLength(100)
    if (results[0]!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(account.currentBytes).toBeLessThan(2 * results[0]!.value.bytes)
    for (const r of results) if (r.kind === 'ready') r.value.release()
    expect(account.currentBytes).toBe(0)
  })

  it('concurrent builds coalesce to one shared graph', async () => {
    const cells = [cell(0, 0, 1), cell(1, 0, 2), cell(2, 0, 3)]
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const [[a], [b]] = await Promise.all([
      stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget()),
      stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget()),
    ])
    expect(a!.kind).toBe('ready')
    expect(b!.kind).toBe('ready')
    if (a!.kind !== 'ready' || b!.kind !== 'ready') throw new Error('expected ready statistics')
    // One shared graph, not two distinct builds.
    expect(a!.value.value).toBe(b!.value.value)
    a!.value.release()
    // Releasing one lease keeps the shared graph charged for the other.
    expect(memory.currentBytes).toBeGreaterThan(0)
    const [hit] = await stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget())
    expect(hit!.kind).toBe('ready')
    // The shared graph is still the same object: no uncharged replacement.
    if (hit!.kind === 'ready') expect(hit!.value.value).toBe(b!.value.value)
    b!.value.release()
    if (hit!.kind === 'ready') hit!.value.release()
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
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

  it('aborting after the first scope unwinds unreturned leases', async () => {
    // Reviewer probe: two scopes, abort during the second. The throw must not
    // leak the first scope's prepared lease.
    const controller = new AbortController()
    let calls = 0
    const cells = [cell(0, 0, 1), cell(1, 0, 2)]
    const { memory, source } = arraySource(cells)
    const aborting = {
      ...source,
      pages: async function* (id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
        for await (const page of source.pages(id, rects, budget, signal)) {
          calls++
          if (calls === 2) controller.abort()
          yield page
        }
      },
    }
    const stats = createCfStatistics(memory)
    await expect(
      stats.prepareStatistics([[rect(0, 0)], [rect(1, 1)]], stamp(1), aborting, frozenBudget(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: 'cancelled' })
    // The in-flight build observes the abort at its next checkpoint and
    // releases its workspace; wait for that cooperative shutdown.
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('a provider error on a later scope unwinds earlier leases and rethrows', async () => {
    // A faulty provider still releases the page it holds when it throws.
    const cells = [cell(0, 0, 1), cell(1, 0, 2)]
    const { memory, source } = arraySource(cells)
    let calls = 0
    const failing = {
      ...source,
      pages: async function* (id: string, rects: readonly SourceRect[], budget: CfWorkBudget, signal: AbortSignal) {
        let held: CfLease<CfValuePage> | undefined
        try {
          for await (const page of source.pages(id, rects, budget, signal)) {
            calls++
            held = page
            if (calls === 2) throw new Error('t5 probe provider failure')
            yield page
            held = undefined
          }
        } finally {
          if (held) {
            try { held.release() } catch { /* release-only path */ }
            held = undefined
          }
        }
      },
    }
    const stats = createCfStatistics(memory)
    await expect(
      stats.prepareStatistics([[rect(0, 0)], [rect(1, 1)]], stamp(1), failing, frozenBudget()),
    ).rejects.toThrow('t5 probe provider failure')
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('a mid-read generation change withholds instead of publishing stale statistics', async () => {
    // Reviewer probe: the source advances gen 1 -> 2 while pages are read.
    let current = stamp(1)
    const src = {
      bounds: async (_id: string) => rect(0, 1),
      stamp: (_id: string) => current,
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 1 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
        current = stamp(2)
        yield {
          value: { rows: new Uint32Array([1]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 2 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const [out] = await stats.prepareStatistics([[rect(0, 1)]], stamp(1), src, frozenBudget())
    expect(out!).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('a new owner and structure generation does not reuse old statistics', async () => {
    // Reviewer probe: same sheet and value generation, different owner and
    // structure identity, must rescan rather than alias the old entry.
    const cells = [cell(0, 0, 99)]
    const counter = makeCounter()
    const live = { gen: 1, owner: 'one', structureGen: 1 }
    const { memory, source } = arraySource(cells, counter, SHEET, live)
    const stats = createCfStatistics(memory)
    const req = (owner: string, structureGen: number): CfStamp => ({
      owner, sheetId: SHEET, workbookValueGen: 1, cfGen: 1, viewGen: 1, structureGen, clockGen: 1,
    })
    const [first] = await stats.prepareStatistics([[rect(0, 0)]], req('one', 1), source, frozenBudget())
    expect(first!.kind).toBe('ready')
    if (first!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(first!.value.value.sum).toBe(99)
    live.owner = 'two'
    live.structureGen = 2
    const [second] = await stats.prepareStatistics([[rect(0, 0)]], req('two', 2), source, frozenBudget())
    expect(second!.kind).toBe('ready')
    if (second!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(second!.value.value.sum).toBe(99)
    expect(counter.cells).toBe(2)
    first!.value.release()
    second!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('a host yield lets a queued cancellation land mid-scan', async () => {
    // Reviewer probe: with real macrotask yields, an already-queued timer abort
    // is observed instead of completing the whole preparation first.
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), 0)
    let timerObserved = false
    timer.unref?.()
    const cells = Array.from({ length: 20000 }, (_, i) => cell(i, 0, i % 7))
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    try {
      await expect(
        stats.prepareStatistics([[rect(0, 19999)]], stamp(1), source, frozenBudget(), { signal: controller.signal }),
      ).rejects.toMatchObject({ code: 'cancelled' })
      timerObserved = true
    } finally {
      clearTimeout(timer)
    }
    expect(timerObserved).toBe(true)
    // The in-flight build observes the abort at its next checkpoint and
    // releases its workspace; wait for that cooperative shutdown.
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('declared-union and scanned-populated counts stay distinct', async () => {
    // Reviewer finding 11: A:A declares 1,048,576 cells; 50 populated rows scan 50.
    const cells: XlsxCell[] = []
    for (let r = 0; r < 50; r++) cells.push(cell(r, 0, r))
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const [out] = await stats.prepareStatistics([[rect(0, 1048575)]], stamp(1), source, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(out!.value.value.declaredCells).toBe(1048576)
    expect(out!.value.value.spanCells).toBe(50)
    expect(out!.value.value.storedCells).toBe(50)
    out!.value.release()
    // Overlapping scopes: union area is deduplicated, stored cells counted once.
    const [overlap] = await stats.prepareStatistics([[rect(0, 10), rect(5, 15)]], stamp(1), source, frozenBudget())
    expect(overlap!.kind).toBe('ready')
    if (overlap!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(overlap!.value.value.declaredCells).toBe(16)
    expect(overlap!.value.value.spanCells).toBe(16)
    expect(overlap!.value.value.storedCells).toBe(16)
    overlap!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('statistics over the session cap withhold with peak at most the cap', async () => {
    // Round-2 finding 1: one scope of 5,000 identical 1,000-character strings
    // must withhold (not publish with a 10 MiB peak over the 8 MiB cap).
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const src = {
      bounds: async (_id: string) => rect(0, 4999),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: {
            rows: Uint32Array.from({ length: 5000 }, (_, i) => i),
            cols: new Uint16Array(5000),
            values: Array.from({ length: 5000 }, () => ({ kind: 'string', value: 'a'.repeat(1000) }) as const),
          },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    const [out] = await stats.prepareStatistics([[rect(0, 4999)]], stamp(1), src, frozenBudget())
    expect(out!).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
    // The session cap binds the peak: never mind the small retained map,
    // the admitted workspace must stay within the dedicated cap.
    expect(memory.peakBytes).toBeLessThanOrEqual(8 * 1024 ** 2)
  })

  it('numeric growth admits the replacement peak while old buffers are live', async () => {
    // Round-2 finding 2: with 1,025 identical numbers, growth builds two
    // 5,120-element arrays while both old 1,024-element arrays are still
    // referenced across the awaited copies. Every construction must observe
    // admitted charge covering all live numeric buffers.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const Seen = globalThis.Float64Array
    let constructed = 0
    const events: Array<{ live: number; charge: number }> = []
    class Instrumented extends Seen {
      constructor(n?: number) {
        super(n as number)
        if (typeof n === 'number') {
          constructed += this.byteLength
          events.push({ live: constructed, charge: memory.currentBytes })
        }
      }
    }
    ;(globalThis as unknown as { Float64Array: unknown }).Float64Array = Instrumented
    try {
      const src = {
        bounds: async (_id: string) => rect(0, 1024),
        stamp: (_id: string) => stamp(1),
        pages: async function* (_id: string) {
          yield {
            value: {
              rows: Uint32Array.from({ length: 1025 }, (_, i) => i),
              cols: new Uint16Array(1025),
              values: Array.from({ length: 1025 }, () => ({ kind: 'number', value: 1 }) as const),
            },
            bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
          }
        },
      }
      const [out] = await stats.prepareStatistics([[rect(0, 1024)]], stamp(1), src, frozenBudget())
      expect(out!.kind).toBe('ready')
      if (out!.kind === 'ready') out!.value.release()
    } finally {
      ;(globalThis as unknown as { Float64Array: unknown }).Float64Array = Seen
    }
    expect(events.length).toBeGreaterThanOrEqual(4)
    for (const e of events) expect(e.charge).toBeGreaterThanOrEqual(e.live)
    expect(memory.currentBytes).toBe(0)
  })

  it('overlap dedup entries are admitted per distinct cell', async () => {
    // Round-2 finding 2: the overlap-dedup Set grows without admission.
    // 1,000 distinct numbers need ~16 KiB workspace + ~128 KiB frequency
    // keys. The overlapping twin adds ~64 KiB of Set entries, so a 180 KiB
    // shared account fits the single rect but withholds the overlap.
    const cells = Array.from({ length: 1000 }, (_, i) => cell(i, 0, i))
    const singleMemory = createCfMemory({ limitBytes: 180 * 1024 })
    const singleInner = createResidentCfSource(cells.map(c => ({ ...c })), SHEET, stamp(1), singleMemory)
    const singleSource = {
      bounds: (id: string) => singleInner.bounds(id),
      stamp: (_id: string) => stamp(1),
      pages: singleInner.pages.bind(singleInner),
    }
    const statsSingle = createCfStatistics(singleMemory)
    const [readyOut] = await statsSingle.prepareStatistics([[rect(0, 999)]], stamp(1), singleSource, frozenBudget())
    expect(readyOut!.kind).toBe('ready')
    if (readyOut!.kind === 'ready') readyOut!.value.release()
    expect(singleMemory.currentBytes).toBe(0)
    // Same cells under two identical rects: the dedup Set charges force a
    // withhold instead of an overrun on the same tight account.
    const tight = createCfMemory({ limitBytes: 180 * 1024 })
    const inner = createResidentCfSource(cells.map(c => ({ ...c })), SHEET, stamp(1), tight)
    const wrapped = {
      bounds: (id: string) => inner.bounds(id),
      stamp: (_id: string) => stamp(1),
      pages: inner.pages.bind(inner),
    }
    const stats = createCfStatistics(tight)
    const [refused] = await stats.prepareStatistics([[rect(0, 999), rect(0, 999)]], stamp(1), wrapped, frozenBudget())
    expect(refused!).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(tight.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('resident pages admit the full modeled payload before building', async () => {
    // Round-2 finding 3: a normal 64-number page was admitted at 1,152 B
    // while the payload model reports 6,872 B. The charge must cover the
    // measured page, for numeric, boolean and string pages alike.
    const check = async (cells: XlsxCell[]): Promise<void> => {
      const memory = createCfMemory({})
      const source = createResidentCfSource(cells, SHEET, stamp(1), memory)
      for await (const page of source.pages(SHEET, [rect(0, cells.length - 1)], frozenBudget(), new AbortController().signal)) {
        const model = measurePayloadBytes(page.value)
        expect(typeof model).toBe('number')
        // Admitted charge covers the modeled retained payload (plus bounded staging).
        expect(page.bytes).toBeGreaterThanOrEqual(model!)
        expect(page.bytes - model!).toBeLessThan(1024)
        page.release()
      }
      expect(memory.currentBytes).toBe(0)
    }
    await check(Array.from({ length: 64 }, (_, i) => cell(i, 0, 1)))
    await check(Array.from({ length: 64 }, (_, i) => cell(i, 0, i % 2 === 0)))
    await check(Array.from({ length: 64 }, (_, i) => cell(i, 0, `k${i}-` + 'x'.repeat(20))))
  })

  it('aborting between adoption and publication releases the candidate', async () => {
    // Round-2 finding 4: abort the sole caller right after hold.adopt and
    // no consumer or cleanup may keep the adopted charge.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const controller = new AbortController()
    const reserve = memory.reserve.bind(memory)
    memory.reserve = ((bytes: number) => {
      const hold = reserve(bytes)
      if (hold.kind === 'reserved') {
        const adopt = hold.adopt.bind(hold)
        hold.adopt = ((value: unknown, retained?: number) => {
          const lease = adopt(value as never, retained)
          if ((value as { sortedAsc?: unknown }).sortedAsc !== undefined) controller.abort()
          return lease
        }) as typeof hold.adopt
      }
      return hold
    }) as typeof memory.reserve
    const src = {
      bounds: async (_id: string) => rect(0, 0),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 1 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    await expect(
      stats.prepareStatistics([[rect(0, 0)]], stamp(1), src, frozenBudget(), { signal: controller.signal }),
    ).rejects.toMatchObject({ code: 'cancelled' })
    await new Promise(resolve => setTimeout(resolve, 10))
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
    stats.dispose()
    expect(memory.currentBytes).toBe(0)
  })

  it('an aborting waiter does not disturb the shared job for others', async () => {
    // The shared build continues for a live waiter after one waiter aborts;
    // the aborted call unwinds while the other still resolves ready.
    const cells = [cell(0, 0, 5), cell(1, 0, 6), cell(2, 0, 7)]
    const { memory, source } = arraySource(cells)
    const stats = createCfStatistics(memory)
    const firstController = new AbortController()
    const a = stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget(), { signal: firstController.signal })
    const b = stats.prepareStatistics([[rect(0, 2)]], stamp(1), source, frozenBudget())
    firstController.abort()
    await expect(a).rejects.toMatchObject({ code: 'cancelled' })
    const [ready] = await b
    expect(ready!.kind).toBe('ready')
    if (ready!.kind === 'ready') ready!.value.release()
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('a generation change during post-scan work withholds at publication', async () => {
    // Round-2 finding 5: the source flips gen 1 -> 2 after the pre-sort
    // stamp check (queued microtask); the final publication gate must see it.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    let current = stamp(1)
    let armed = false
    const src = {
      bounds: async (_id: string) => rect(0, 999),
      stamp: (_id: string) => {
        if (armed) queueMicrotask(() => { current = stamp(2) })
        return current
      },
      pages: async function* (_id: string) {
        yield {
          value: {
            rows: Uint32Array.from({ length: 1000 }, (_, i) => i),
            cols: new Uint16Array(1000),
            values: Array.from({ length: 1000 }, (_, i) => ({ kind: 'number', value: i }) as const),
          },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
        armed = true
      },
    }
    const [out] = await stats.prepareStatistics([[rect(0, 999)]], stamp(1), src, frozenBudget())
    expect(out!).toMatchObject({ kind: 'unresolved', reason: 'cf-stale-generation' })
    expect(current.workbookValueGen).toBe(2)
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('warm cache hits do not rescan bounds', async () => {
    // Round-2 finding 6: a cache hit must not rescan the sheet (3,001 row
    // reads on a 1,000-cell source); bounds are memoized per generation.
    let reads = 0
    const cells: XlsxCell[] = Array.from({ length: 1000 }, (_, i) => ({
      get row() { reads++; return i },
      col: 0, ref: `A${i + 1}`, value: 1, styleIndex: 0,
    } as XlsxCell))
    const counter = makeCounter()
    const { memory, source } = arraySource(cells, counter)
    const stats = createCfStatistics(memory)
    const [first] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), source, frozenBudget())
    expect(first!.kind).toBe('ready')
    const boundsCalls = counter.bounds
    reads = 0
    const [hit] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), source, frozenBudget())
    expect(hit!.kind).toBe('ready')
    if (first!.kind === 'ready') first!.value.release()
    if (hit!.kind === 'ready') hit!.value.release()
    expect(counter.bounds).toBe(boundsCalls)
    expect(reads).toBe(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('rejected cells spend shared budget in the resident adapter', async () => {
    // Round-2 finding 6: a query excluding all stored cells must still
    // charge the examination work instead of completing with zero units.
    let reads = 0
    const cells: XlsxCell[] = Array.from({ length: 10000 }, (_, i) => ({
      get row() { reads++; return i < 5000 ? 0 : 10000 },
      col: i % 5000, ref: 'A1', value: 1, styleIndex: 0,
    } as unknown as XlsxCell))
    const budget = frozenBudget()
    const source = createResidentCfSource(cells, SHEET, stamp(1), createCfMemory({}))
    let yielded = 0
    for await (const page of source.pages(SHEET, [{ firstRow: 1, lastRow: 9999, firstCol: 0, lastCol: 4999 }], budget, new AbortController().signal)) {
      yielded++
      page.release()
    }
    expect(yielded).toBe(0)
    expect(budget.units).toBe(10000)
    // Filtering must examine each cell (short-circuit comparisons read row
    // at most twice), but every examination is charged exactly once.
    expect(reads).toBeLessThanOrEqual(20000)
  })

  it('tiny preparations flush their budget remainder', async () => {
    // Round-2 finding 6: a one-number preparation must count its tail
    // instead of completing with zero statistics units.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const src = {
      bounds: async (_id: string) => rect(0, 0),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 1 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    const [out] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), src, budget)
    expect(out!.kind).toBe('ready')
    if (out!.kind === 'ready') out!.value.release()
    expect(budget.units).toBeGreaterThan(0)
    expect(memory.currentBytes).toBe(0)
  })

  it('union traversal across many rectangles is budgeted', async () => {
    // Round-2 finding 6: coordinate work per scope must count. Same single
    // cell under 65 identical rects must cost clearly more than under one.
    // A bare source (no adapter spends) isolates build-level accounting:
    // pre-fix the gap is only the 64 extra cell visits; the counted union
    // adds ~132 more (65 edges + slabs, times declared and span).
    const bare = () => ({
      bounds: async (_id: string) => rect(0, 0),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 1 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    })
    const oneBudget = frozenBudget()
    {
      const memory = createCfMemory({})
      const stats = createCfStatistics(memory)
      const [out] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), bare(), oneBudget)
      expect(out!.kind).toBe('ready')
      if (out!.kind === 'ready') out!.value.release()
      expect(memory.currentBytes).toBe(0)
    }
    const manyBudget = frozenBudget()
    {
      const memory = createCfMemory({})
      const stats = createCfStatistics(memory)
      const scopes = [Array.from({ length: 65 }, () => rect(0, 0))]
      const [out] = await stats.prepareStatistics(scopes, stamp(1), bare(), manyBudget)
      expect(out!.kind).toBe('ready')
      if (out!.kind === 'ready') {
        expect(out!.value.value.declaredCells).toBe(1)
        out!.value.release()
      }
      expect(memory.currentBytes).toBe(0)
    }
    // Union + dedup work over 65 rects dwarfs the single-rect baseline.
    expect(manyBudget.units - oneBudget.units).toBeGreaterThanOrEqual(128)
  })

  it('union of more than 64 identical ranges is exact', async () => {
    // Round-2 finding 7: sixty-five identical A1 ranges must report the
    // true one-cell union instead of the naive sum (65).
    const { memory, source } = arraySource([cell(0, 0, 1)])
    const stats = createCfStatistics(memory)
    const scopes = [Array.from({ length: 65 }, () => rect(0, 0))]
    const [out] = await stats.prepareStatistics(scopes, stamp(1), source, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(out!.value.value.declaredCells).toBe(1)
    expect(out!.value.value.spanCells).toBe(1)
    expect(out!.value.value.storedCells).toBe(1)
    out!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('overlapping explicit blanks deduplicate in stored counts', async () => {
    // Round-2 finding 7: explicit blanks bypassed the overlap set and were
    // counted once per overlapping rect.
    const { memory, source } = arraySource([cell(0, 0, null)])
    const stats = createCfStatistics(memory)
    const [out] = await stats.prepareStatistics([[rect(0, 0), rect(0, 0)]], stamp(1), source, frozenBudget())
    expect(out!.kind).toBe('ready')
    if (out!.kind !== 'ready') throw new Error('expected ready statistics')
    expect(out!.value.value.storedCells).toBe(1)
    out!.value.release()
    expect(memory.currentBytes).toBe(0)
  })

  it('bounds follow owner and structure generations instead of stale geometry', async () => {
    // Round-3 finding 1: the bounds memo keyed only sheet+generation, so a
    // structure change or owner replacement reused the old A1 bounds and
    // reported numericCount 0 for A11's value 99.
    const run = async (change: 'structure' | 'owner'): Promise<void> => {
      const memory = createCfMemory({})
      const stats = createCfStatistics(memory)
      const firstSource = {
        bounds: async (_id: string) => rect(0, 0),
        stamp: (_id: string) => stamp(1),
        pages: async function* (_id: string) {
          yield {
            value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 1 } as const] },
            bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
          }
        },
      }
      const [first] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), firstSource, frozenBudget())
      expect(first!.kind).toBe('ready')
      if (first!.kind === 'ready') first!.value.release()
      const next: CfStamp = { ...stamp(1), structureGen: 2, ...(change === 'owner' ? { owner: 'new' } : {}) }
      const secondSource = {
        bounds: async (_id: string) => rect(0, 10),
        stamp: (_id: string) => next,
        pages: async function* (_id: string) {
          yield {
            value: { rows: new Uint32Array([10]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 99 } as const] },
            bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
          }
        },
      }
      const [second] = await stats.prepareStatistics([[rect(10, 10)]], next, secondSource, frozenBudget())
      expect(second!.kind).toBe('ready')
      if (second!.kind !== 'ready') throw new Error('expected ready statistics after generation change')
      expect(second!.value.value.numericCount).toBe(1)
      expect(second!.value.value.sum).toBe(99)
      second!.value.release()
      expect(memory.currentBytes).toBe(0)
      expect(stats.size).toBe(0)
    }
    await run('structure')
    await run('owner')
  })

  it('a pending bounds answer is not memoized past its generation', async () => {
    // Round-3 finding 1: an undefined bounds (importing source) was cached,
    // so a later complete source at the same generation stayed pending.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const pendingSource = {
      bounds: async (_id: string) => undefined,
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {},
    }
    const [pending] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), pendingSource, frozenBudget())
    expect(pending!).toMatchObject({ kind: 'unresolved', reason: 'cf-pending-data' })
    const completeSource = {
      bounds: async (_id: string) => rect(0, 0),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield {
          value: { rows: new Uint32Array([0]), cols: new Uint16Array([0]), values: [{ kind: 'number', value: 7 } as const] },
          bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); },
        }
      },
    }
    const [ready] = await stats.prepareStatistics([[rect(0, 0)]], stamp(1), completeSource, frozenBudget())
    expect(ready!.kind).toBe('ready')
    if (ready!.kind !== 'ready') throw new Error('expected ready statistics once bounds arrive')
    expect(ready!.value.value.sum).toBe(7)
    ready!.value.release()
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('inner union examinations are budgeted per visit', async () => {
    // Round-3 finding 2: one coordinate slab scanned 256 rectangles with a
    // single unit. 64 shifted rectangles (257-span, 1-column shift) force
    // ~8k examinations per union; the shared budget must count them.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const rects = Array.from({ length: 64 }, (_, i) => ({ firstRow: i, lastRow: i + 256, firstCol: i, lastCol: i + 256 }))
    const src = {
      bounds: async (_id: string) => ({ firstRow: 0, lastRow: 319, firstCol: 0, lastCol: 319 }),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] as CfLiteral[] }, bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); } }
      },
    }
    const [out] = await stats.prepareStatistics([rects], stamp(1), src, budget)
    expect(out!.kind).toBe('ready')
    if (out!.kind === 'ready') out!.value.release()
    // Two unions (declared + span) of ~127 slabs × ~64 examinations each.
    expect(budget.units).toBeGreaterThanOrEqual(10000)
    expect(memory.currentBytes).toBe(0)
  })

  it('four heavy union scopes exceed the shared work budget instead of hiding it', async () => {
    // Round-3 finding 2: 4 scopes × 256 shifted rectangles cost 1,046,528
    // inner examinations but reported only 7,160 units. Counted honestly
    // they exceed the 1,000,000-unit plan bound, so later scopes withhold.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const scopes = Array.from({ length: 4 }, (_, group) =>
      Array.from({ length: 256 }, (_, i) => ({ firstRow: group * 600 + i, lastRow: group * 600 + i + 256, firstCol: i, lastCol: i + 256 })),
    )
    const src = {
      bounds: async (_id: string) => ({ firstRow: 0, lastRow: 3000, firstCol: 0, lastCol: 512 }),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] as CfLiteral[] }, bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); } }
      },
    }
    const outs = await stats.prepareStatistics(scopes, stamp(1), src, budget)
    expect(outs.filter(o => o.kind === 'ready').length).toBeLessThan(4)
    expect(outs.filter(o => o.kind === 'unresolved').length).toBeGreaterThanOrEqual(1)
    for (const o of outs) {
      if (o.kind === 'ready') o.value.release()
      else expect(o).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    }
    expect(budget.units).toBeGreaterThan(900000)
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('a zero-byte account refuses before touching the source', async () => {
    // Round-3 finding 2: a cold 10,000-cell preparation performed 30,001 row
    // reads before refusing on an exhausted account. The session probes for
    // minimum room first, so refusal costs zero reads and zero units.
    let reads = 0
    const cells: XlsxCell[] = Array.from({ length: 10000 }, (_, i) => ({
      get row() { reads++; return i },
      col: 0, ref: `A${i + 1}`, value: 1, styleIndex: 0,
    } as unknown as XlsxCell))
    const budget = frozenBudget()
    const source = createResidentCfSource(cells, SHEET, stamp(1), createCfMemory({}))
    const session = createCfStatistics(createCfMemory({ limitBytes: 0 }))
    const [out] = await session.prepareStatistics([[rect(0, 0)]], stamp(1), source, budget)
    expect(out!).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    expect(reads).toBe(0)
    expect(budget.units).toBe(0)
  })

  it('union sort comparisons and merges are counted per interval', async () => {
    // Round-4 finding 1: two scopes of 256 shuffled rectangles ran 1.6M
    // sort-comparator calls while the budget recorded only examinations.
    // Counting comparisons and merge iterations withholds the over-budget
    // scope instead of publishing past the plan bound.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const budget = frozenBudget()
    const shuffled = (group: number) => Array.from({ length: 256 }, (_, i) => ({
      firstRow: group * 1000 + ((i * 149) % 257),
      lastRow: group * 1000 + ((i * 149) % 257) + 256,
      firstCol: i,
      lastCol: i + 256,
    }))
    const src = {
      bounds: async (_id: string) => ({ firstRow: 0, lastRow: 3000, firstCol: 0, lastCol: 512 }),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] as CfLiteral[] }, bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); } }
      },
    }
    const outs = await stats.prepareStatistics([shuffled(0)], stamp(1), src, budget)
    expect(outs.filter(o => o.kind === 'unresolved').length).toBeGreaterThanOrEqual(1)
    for (const o of outs) {
      if (o.kind === 'ready') o.value.release()
      else expect(o).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    }
    // Counted examinations plus comparisons and merges reach the bound.
    expect(budget.units).toBeGreaterThan(900000)
    expect(memory.currentBytes).toBe(0)
    expect(stats.size).toBe(0)
  })

  it('resident bounds count actual examined cells and observe cancellation', async () => {
    // Round-3 finding 2: populated bounds must spend for every examined cell
    // and abort promptly, instead of scanning silently. Now genuinely
    // cooperative: yields to the host at checkpoints.
    const cells: XlsxCell[] = Array.from({ length: 1000 }, (_, i) => cell(i, 0, 1))
    const source = createResidentCfSource(cells, SHEET, stamp(1), createCfMemory({}))
    const budget = frozenBudget()
    const bounds = await source.bounds(SHEET, budget, new AbortController().signal)
    expect(bounds).toMatchObject({ firstRow: 0, lastRow: 999 })
    expect(budget.units).toBe(1000)
    const aborted = new AbortController()
    aborted.abort()
    await expect(source.bounds(SHEET, frozenBudget(), aborted.signal)).rejects.toMatchObject({ code: 'cancelled' })
  })

  it('cold bounds yield to the host so queued cancellations land mid-scan', async () => {    // Round-4 finding 2: a timer abort queued before bounds must be observed
    // mid-scan (not after a complete 10,000-row sweep). The abort fires on a
    // row read so the timing is deterministic, not timer-flaky.
    const controller = new AbortController()
    let reads = 0
    const cells: XlsxCell[] = Array.from({ length: 10000 }, (_, i) => ({
      get row() { reads++; if (reads === 500) controller.abort(); return i },
      col: 0, ref: `A${i + 1}`, value: 1, styleIndex: 0,
    } as unknown as XlsxCell))
    const budget = frozenBudget()
    const source = createResidentCfSource(cells, SHEET, stamp(1), createCfMemory({}))
    await expect(source.bounds(SHEET, budget, controller.signal)).rejects.toMatchObject({ code: 'cancelled' })
    expect(reads).toBeLessThan(1000)
    expect(reads).toBeGreaterThan(0)
    expect(budget.units).toBeGreaterThan(0)
  })

  it('sweep workspace admits the simultaneous peak tracked in the session cap', async () => {
    // Round-4 finding 3: for 256 rectangles the sweep hold must cover the
    // clamped objects and array (32,784 B), the cleaned reference array
    // (2,064 B), and the simultaneous sweep structures (~34,880 B) —
    // 69,728 B total. The clamp hold lives through the sweeps and is
    // tracked in the cap, so the admitted high-water mark covers the peak.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    const rectCount = 256
    const rects = Array.from({ length: rectCount }, (_, i) => ({ firstRow: i, lastRow: i + 256, firstCol: i, lastCol: i + 256 }))
    const src = {
      bounds: async (_id: string) => ({ firstRow: 0, lastRow: 512, firstCol: 0, lastCol: 512 }),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] as CfLiteral[] }, bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); } }
      },
    }
    const [out] = await stats.prepareStatistics([rects], stamp(1), src, frozenBudget())
    expect(out!.kind).toBe('ready')
    expect(stats.size).toBe(1)
    if (out!.kind === 'ready') out!.value.release()
    // Simultaneous modeled peak: clamped objects and array, cleaned array,
    // edge set, coordinate list, and interval arrays at the widest slab.
    const clamped = rectCount * (24 + 4 * 16 + 4 * 8) + (16 + 8 * rectCount)
    const cleaned = 16 + 8 * rectCount
    const sweep = 32 + 24 * 512 + 16 + 16 * 512 + 16 + 56 * 256
    expect(memory.peakBytes).toBeGreaterThanOrEqual(clamped + cleaned + sweep)
    expect(memory.currentBytes).toBe(0)
    stats.dispose()
    expect(stats.size).toBe(0)
  })

  it('union sort comparisons and merge iterations are spent before running', async () => {
    // Round-5 finding 1: comparisons and merges must be spent before/during
    // execution, never as post-hoc bulk charges. Two shuffled 256-rectangle
    // scopes exceed the 1 M plan bound once honestly counted, so at least
    // one withholds instead of publishing over-budget work.
    const memory = createCfMemory({})
    const stats = createCfStatistics(memory)
    let attempted = 0
    const inner = frozenBudget()
    const budget: CfWorkBudget = {
      get units() { return inner.units },
      spend(n: number) { attempted += n; return inner.spend(n) },
      resume() { inner.resume() },
    }
    const shuffled = (group: number) => Array.from({ length: 256 }, (_, i) => ({
      firstRow: group * 1000 + ((i * 149) % 257),
      lastRow: group * 1000 + ((i * 149) % 257) + 256,
      firstCol: i,
      lastCol: i + 256,
    }))
    const src = {
      bounds: async (_id: string) => ({ firstRow: 0, lastRow: 3000, firstCol: 0, lastCol: 512 }),
      stamp: (_id: string) => stamp(1),
      pages: async function* (_id: string) {
        yield { value: { rows: new Uint32Array(0), cols: new Uint16Array(0), values: [] as CfLiteral[] }, bytes: 0, release() {}, transfer() { throw new Error('t5 probe'); } }
      },
    }
    const outs = await stats.prepareStatistics([shuffled(0), shuffled(1)], stamp(1), src, budget)
    expect(outs.filter(o => o.kind === 'ready').length).toBeLessThan(2)
    for (const o of outs) {
      if (o.kind === 'ready') o.value.release()
      else expect(o).toMatchObject({ kind: 'unresolved', reason: 'cf-resource-limit' })
    }
    // The comparisons and merges actually flowed through the spender...
    expect(attempted).toBeGreaterThan(900000)
    // ...while performed work respects the bound, with only a batched tail
    // refused (no hidden post-work bulk escapes counting).
    expect(budget.units).toBeLessThanOrEqual(CF_LIMITS.workUnits + 127)
    expect(attempted - budget.units).toBeLessThan(512)
    expect(memory.currentBytes).toBe(0)
  })
})
