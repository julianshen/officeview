import { expect, it } from 'vitest'
import { createCfMemory, createCfWorkBudget } from '../src/xlsx/cf/limits'
import { createResidentCfSource } from '../src/xlsx/cf/statistics'
import type { CfPrepared, CfStamp, CfSheetRules, CfValueSource } from '../src/xlsx/cf/model'
import * as cacheApi from '../src/xlsx/cf/cache'
import * as serviceApi from '../src/xlsx/cf/service'
const stamp = (value = 1): CfStamp => ({ owner: 'owner', sheetId: 's', workbookValueGen: value, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 })
const rect = { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }
const rules: CfSheetRules = { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, ranges: [rect], stopIfTrue: false, writes: new Set(['fill']), dxf: { 'fill.pattern': 'solid', 'fill.fg': 'FF00FF00' }, definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] } }] }
const request = (s = stamp(), signal = new AbortController().signal) => ({ stamp: s, sourceRects: [rect], signal, grantBytes: 1_000_000 })
const budget = () => createCfWorkBudget({ now: () => 0 })
function deferred() { let resolve!: () => void; const promise = new Promise<void>(r => { resolve = r }); return { promise, resolve } }
function resident(value: number, s = stamp()) { return createResidentCfSource([{ ref: 'A1', row: 0, col: 0, value, styleIndex: 0 }], 's', s, createCfMemory()) }
it('current negative cell never paints positive previous-generation fill, including late completion', async () => {
  const { createCfService } = serviceApi; const memory = createCfMemory(); const service = createCfService(rules, memory, { budget })
  const positive = await service.prepareCf(request(), resident(1))
  expect(positive.value.at(0, 0)).toMatchObject({ kind: 'ready', value: { properties: { 'fill.fg': 'FF00FF00' } } })
  const gate = deferred(); const started = deferred(); const next = stamp(2); const base = resident(-1, next)
  const source: CfValueSource = { ...base, async *pages(...args) { started.resolve(); await gate.promise; yield* base.pages(...args) } }
  const pending = service.prepareCf(request(next), source); await started.promise
  expect(positive.value.at(0, 0).kind).toBe('unresolved'); expect(service.at(next, 0, 0).kind).toBe('unresolved')
  gate.resolve(); const negative = await pending
  expect(negative.value.at(0, 0)).toMatchObject({ kind: 'ready', value: { properties: {} } })
  positive.release(); negative.release(); service.dispose(); expect(memory.currentBytes).toBe(0)
})
for (const key of ['owner', 'cfGen', 'viewGen', 'structureGen', 'clockGen'] as const) it(`invalidates the full ${key} stamp`, async () => {
  const { createCfService } = serviceApi; const memory = createCfMemory(); const service = createCfService(rules, memory, { budget })
  const first = await service.prepareCf(request(), resident(1)); const changed = { ...stamp(), [key]: key === 'owner' ? 'other' : 2 }
  const second = await service.prepareCf(request(changed), resident(-1, changed))
  expect(first.value.at(0, 0).kind).toBe('unresolved'); first.release(); second.release(); service.dispose(); expect(memory.currentBytes).toBe(0)
})
it('deduplicates computation but does not refund live producer memory on cancellation or close', async () => {
  const { createCfCache } = cacheApi; const memory = createCfMemory(); const cache = createCfCache(memory); const gate = deferred(), entered = deferred(); let visits = 0, disposed = 0
  const producer = async () => { visits++; const hold = memory.reserve(1024); if (hold.kind !== 'reserved') throw Error('refused'); entered.resolve(); await gate.promise; return hold.adopt<CfPrepared>({ stamp: stamp(), at: () => ({ kind: 'no-match' }), dispose: () => { disposed++ } }) }
  const c = new AbortController(); const a = cache.prepare('same', stamp(), c.signal, producer); const aError = a.catch((e: any) => e.code)
  const b = cache.prepare('same', stamp(), new AbortController().signal, producer); const bError = b.catch((e: any) => e.code)
  await entered.promise; expect(visits).toBe(1); c.abort(); expect(await aError).toBe('cancelled'); expect(memory.currentBytes).toBeGreaterThanOrEqual(1024)
  cache.dispose(); expect(await bError).toBe('cancelled'); expect(memory.currentBytes).toBeGreaterThanOrEqual(1024)
  gate.resolve(); await new Promise(r => setTimeout(r, 10)); expect(disposed).toBe(1); expect(memory.currentBytes).toBe(0)
})
it('one cancelled waiter does not cancel its peer; ready cache and consumer leases hold one graph', async () => {
  const { createCfCache } = cacheApi; const memory = createCfMemory(), cache = createCfCache(memory), gate = deferred(); let visits = 0
  const producer = async () => { visits++; const hold = memory.reserve(512); if (hold.kind !== 'reserved') throw Error(); await gate.promise; return hold.adopt<CfPrepared>({ stamp: stamp(), at: () => ({ kind: 'no-match' }), dispose() {} }) }
  const controller = new AbortController(); const first = cache.prepare('x', stamp(), controller.signal, producer).catch((e: any) => e.code)
  const second = cache.prepare('x', stamp(), new AbortController().signal, producer); controller.abort(); expect(await first).toBe('cancelled'); gate.resolve()
  const lease = await second; expect(visits).toBe(1); cache.dispose(); expect(memory.currentBytes).toBeGreaterThanOrEqual(512); lease.release(); expect(memory.currentBytes).toBe(0)
})
it('late old-generation producers cannot publish and their losing results dispose', async () => {
  const { createCfCache } = cacheApi; const m = createCfMemory(), cache = createCfCache(m), gate = deferred(); let disposed = 0
  const old = cache.prepare('x', stamp(), new AbortController().signal, async () => { const h = m.reserve(512); if (h.kind !== 'reserved') throw Error(); await gate.promise; return h.adopt<CfPrepared>({ stamp: stamp(), at: () => ({ kind: 'no-match' }), dispose() { disposed++ } }) }).catch((e: any) => e.code)
  cache.activate(stamp(2)); expect(await old).toBe('cancelled'); expect(m.currentBytes).toBeGreaterThanOrEqual(512); gate.resolve(); await new Promise(r => setTimeout(r, 10)); expect(disposed).toBe(1); expect(m.currentBytes).toBe(0); cache.dispose()
})
it('refuses before source visits on zero grant, oversized batch and missing sheet scope', async () => {
  const { createCfService } = serviceApi; const memory = createCfMemory(); let visits = 0; const source = resident(1); const basePages = source.pages; source.pages = async function* (...a) { visits++; yield* basePages(...a) }
  const service = createCfService(rules, memory, { budget }); await expect(service.prepareCf({ ...request(), grantBytes: 0 }, source)).rejects.toMatchObject({ code: 'resource-limit' })
  await expect(service.prepareCf({ ...request(), sourceRects: [{ ...rect, lastRow: 5000 }] }, source)).rejects.toMatchObject({ code: 'resource-limit' })
  const unsafe = createCfService({ ...rules, unresolvedSheet: { reason: 'cf-over-cap', writes: 'ALL' } }, memory, { budget }); const held = await unsafe.prepareCf(request(), source)
  expect(held.value.at(0, 0)).toMatchObject({ kind: 'unresolved', reason: 'cf-over-cap' }); expect(visits).toBe(0); held.release(); unsafe.dispose(); service.dispose(); expect(memory.currentBytes).toBe(0)
})
it('resident and paged-shaped source snapshots agree, cover blanks and are immutable', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory(); const s = createCfService(rules, m, { budget }); const base = resident(1); let visits = 0
  const paged: CfValueSource = { bounds: (...a) => base.bounds(...a), stamp: (...a) => base.stamp(...a), async *pages(...a) { visits++; yield* base.pages(...a) } }
  const first = await s.prepareCf(request(), paged), second = await s.prepareCf(request(), paged)
  expect(first.value.at(0, 0)).toEqual(second.value.at(0, 0)); expect(visits).toBe(1)
  const state = first.value.at(0, 0); if (state.kind === 'ready') expect(Object.isFrozen(state.value.properties)).toBe(true)
  expect(first.value.at(10, 10).kind).toBe('unresolved'); first.release(); second.release(); s.dispose(); expect(m.currentBytes).toBe(0)
})
it('unmeasured positive precedence conflicts refuse; a reviewed evidence gate enables composition', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory(); const paired = { ...rules, rules: [rules.rules[0]!, { ...rules.rules[0]!, id: 1, priority: 2, documentOrder: 1, dxf: { 'fill.fg': 'FFFF0000' } }] }
  const unmeasured = createCfService(paired, m, { budget, evidence: { allows: () => false } }); const withheld = await unmeasured.prepareCf(request(), resident(1))
  expect(withheld.value.at(0, 0)).toMatchObject({ kind: 'unresolved', reason: 'cf-unmeasured-capability' }); withheld.release(); unmeasured.dispose()
  const measured = createCfService(paired, m, { budget, evidence: { allows: () => true } }); const ready = await measured.prepareCf(request(), resident(1))
  expect(ready.value.at(0, 0)).toMatchObject({ kind: 'ready', value: { properties: { 'fill.fg': 'FF00FF00' } } }); ready.release(); measured.dispose(); expect(m.currentBytes).toBe(0)
})
it('a refused cold preparation invalidates old handles before it visits any new source', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory(), s = createCfService(rules, m, { budget }); const first = await s.prepareCf(request(), resident(1))
  await expect(s.prepareCf({ ...request(stamp(2)), grantBytes: 0 }, resident(-1, stamp(2)))).rejects.toMatchObject({ code: 'resource-limit' })
  expect(first.value.at(0, 0).kind).toBe('unresolved'); first.release(); s.dispose(); expect(m.currentBytes).toBe(0)
})
it('a shared statistics scope is scanned once for many rules and source lookup uses numeric cell IDs', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory(); const range = { ...rect, lastRow: 49 }; const many: CfSheetRules = { diagnostics: [], rules: Array.from({ length: 100 }, (_,i) => ({ ...rules.rules[0]!, id: i, priority: i+1, documentOrder: i, ranges: [range], definition: { kind: 'aboveAverage', above: true, equal: false, dialect: 'base', formulas: [] } })) }
  const source = createResidentCfSource(Array.from({ length: 50 }, (_,row) => ({ ref: 'A'+(row+1), row, col: 0, value: row, styleIndex: 0 })), 's', stamp(), m)
  const pages = source.pages; let visits = 0; source.pages = async function* (...a) { for await (const page of pages(...a)) { visits += page.value.rows.length; yield page } }
  const s = createCfService(many, m, { budget, evidence: { allows: () => true } }); const lease = await s.prepareCf({ ...request(), sourceRects: [range], grantBytes: 8*1024**2 }, source)
  expect(visits).toBe(100); expect(lease.value.at(49, 0).kind).toBe('ready'); expect(m.peakBytes).toBeLessThan(8*1024**2); lease.release(); s.dispose(); expect(m.currentBytes).toBe(0)
})
it('admits predicate string workspace before case folding a leased source value', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory({ limitBytes: 1024*1024 }); const s = createCfService({ ...rules, rules: [{ ...rules.rules[0]!, definition: { kind: 'text', type: 'containsText', needle: 'X', dialect: 'base', formulas: [] } }] }, m, { budget })
  const text = 'X'.repeat(600000); const source = createResidentCfSource([{ ref: 'A1', row: 0, col: 0, value: text, styleIndex: 0 }], 's', stamp(), createCfMemory())
  const lower = String.prototype.toLowerCase; let largeCalls = 0
  String.prototype.toLowerCase = function() { if (this.length === text.length) largeCalls++; return lower.call(this) }
  try { const lease = await s.prepareCf(request(), source); expect(largeCalls).toBe(0); expect(lease.value.at(0,0).kind).toBe('unresolved'); lease.release() } finally { String.prototype.toLowerCase = lower; s.dispose() }
  expect(m.currentBytes).toBe(0)
})
it('does not construct a request key before admission on an exhausted account', async () => {
  const { createCfService } = serviceApi; const m = createCfMemory({ limitBytes: 0 }), s = createCfService(rules, m, { budget }), req = request(), source = resident(1)
  const map = Array.prototype.map; let builds = 0
  Array.prototype.map = function(this: unknown[], ...args: Parameters<typeof map>) { if (this === req.sourceRects) builds++; return map.apply(this, args) } as typeof map
  try { await expect(s.prepareCf(req, source)).rejects.toMatchObject({ code: 'resource-limit' }); expect(builds).toBe(0) } finally { Array.prototype.map = map; s.dispose() }
})
it('breaks equal-priority ties by document order, favoring later document order rules', async () => {
  const { createCfService } = serviceApi
  const memory = createCfMemory()
  const tiedRules: CfSheetRules = {
    diagnostics: [],
    rules: [
      {
        id: 0,
        priority: 7,
        documentOrder: 0,
        ranges: [rect],
        stopIfTrue: false,
        writes: new Set(['fill']),
        dxf: { 'fill.pattern': 'solid', 'fill.fg': 'FFFF0000' },
        definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 10 }], dialect: 'base', formulas: [] },
      },
      {
        id: 1,
        priority: 7,
        documentOrder: 1,
        ranges: [rect],
        stopIfTrue: false,
        writes: new Set(['fill']),
        dxf: { 'fill.pattern': 'solid', 'fill.fg': 'FF00FF00' },
        definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
      },
    ],
  }
  const service = createCfService(tiedRules, memory, { budget, evidence: { allows: () => true } })
  const prepared = await service.prepareCf(request(), resident(25))
  expect(prepared.value.at(0, 0)).toMatchObject({ kind: 'ready', value: { properties: { 'fill.fg': 'FF00FF00' } } })
  prepared.release()
  service.dispose()
  expect(memory.currentBytes).toBe(0)
})

it('CF sqref on A1:A5000 on a 3-row sheet does not extend extent or prepare 5000 rows (I3)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { ObjectXlsxSheetAdapter } = await import('../src/xlsx/view/source')
  const sheet: any = {
    name: 'Sheet1',
    sheetId: 's1',
    rows: [
      { index: 0, cells: [{ col: 0, row: 0, value: 10, valueIsError: false, styleIndex: 0 }] },
      { index: 1, cells: [{ col: 0, row: 1, value: 20, valueIsError: false, styleIndex: 0 }] },
      { index: 2, cells: [{ col: 0, row: 2, value: 30, valueIsError: false, styleIndex: 0 }] },
    ],
    cols: [],
    merges: [],
    mergeRanges: [],
    conditionalFormatting: {
      rules: [
        {
          id: 1,
          priority: 1,
          documentOrder: 0,
          ranges: [{ firstRow: 0, lastRow: 4999, firstCol: 0, lastCol: 0 }],
          stopIfTrue: false,
          writes: new Set(['fill']),
          dxf: { 'fill.pattern': 'solid', 'fill.fg': 'FFFF0000' },
          definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 0 }], dialect: 'base', formulas: [] },
        },
      ],
    },
  }
  const store = new ObjectXlsxSheetAdapter({ docId: 'doc1', sheetId: 'Sheet1', valueGen: 1 }, sheet, [], 262144)
  expect(store.extent.used).toEqual({ r0: 0, r1: 2, c0: 0, c1: 0 })
  const paint = await prepareCfPaint(sheet)
  expect(paint).toBeDefined()
  expect(paint!.layer.rects).toEqual([{ minRow: 0, maxRow: 2, minCol: 0, maxCol: 0 }])
  paint!.dispose()

  const { computeMetrics } = await import('../src/xlsx/render')
  const { getPaintables } = await import('../src/render/paint')

  const m = computeMetrics(sheet)
  expect(m.exportLimits).toBeUndefined()
  const pages = await getPaintables({ sheets: [sheet] })
  expect(pages[0].spec.heightPx).toBeLessThan(1000)
  expect(pages[0].spec).not.toEqual({ widthPx: 480, heightPx: 96 })

  // Also verify blank-matching rules never extend static extent or trigger refusal
  sheet.conditionalFormatting.rules[0].definition = { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] }
  const mBlank = computeMetrics(sheet)
  expect(mBlank.exportLimits).toBeUndefined()
  const pagesBlank = await getPaintables({ sheets: [sheet] })
  expect(pagesBlank[0].spec.heightPx).toBeLessThan(1000)
  expect(pagesBlank[0].spec).not.toEqual({ widthPx: 480, heightPx: 96 })
})

it('range statistics are cached across preparations with zero rescan visits (I5)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  let a2Scans = 0
  const rows: XlsxRow[] = [
    { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] },
  ]
  for (let r = 1; r < 50; r++) {
    const cell: any = { row: r, col: 0, ref: `A${r + 1}`, styleIndex: 0 }
    Object.defineProperty(cell, 'value', { get() { a2Scans++; return r } })
    rows.push({ index: r, cells: [cell] })
  }
  const sheet: any = {
    name: 'S', sheetId: '1', merges: [], mergeRanges: [], cols: [],
    rows,
    conditionalFormatting: {
      diagnostics: [],
      rules: [{
        id: 0, priority: 1, documentOrder: 0, stopIfTrue: false,
        ranges: [{ firstRow: 0, lastRow: 49, firstCol: 0, lastCol: 0 }],
        writes: new Set(['fill']),
        dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00' },
        definition: { kind: 'top10', rank: 5, percent: false, bottom: false, dialect: 'base', formulas: [] },
      }],
    },
  }
  // Prepare only row 0 (A1)
  const req = { sourceRects: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }] }
  const first = await prepareCfPaint(sheet, { request: req as any })
  expect(first).toBeDefined()
  first!.dispose()
  const firstScans = a2Scans
  expect(firstScans).toBe(49) // All 49 rows scanned for statistics on first preparation

  // Second preparation of row 0 must reuse cached statistics and perform ZERO rescan visits on rows 1..49
  const second = await prepareCfPaint(sheet, { request: req as any })
  expect(second).toBeDefined()
  second!.dispose()
  const rescanVisits = a2Scans - firstScans
  expect(rescanVisits).toBe(0)
})

it('a 3000x20 sheet with one rule keeps its CF without whole-sheet refusal (I4)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const rows: XlsxRow[] = []
  for (let r = 0; r < 3000; r++) {
    const cells: XlsxCell[] = []
    for (let c = 0; c < 20; c++) {
      cells.push({ row: r, col: c, ref: '', value: (r + c) % 10, styleIndex: 0 })
    }
    rows.push({ index: r, cells })
  }
  const sheet: XlsxSheet = {
    name: 'S', sheetId: '1', merges: [], mergeRanges: [], cols: [],
    rows,
    conditionalFormatting: {
      diagnostics: [],
      rules: [{
        id: 0, priority: 1, documentOrder: 0, stopIfTrue: false,
        ranges: [{ firstRow: 0, lastRow: 2999, firstCol: 0, lastCol: 19 }],
        writes: new Set(['fill']),
        dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00' },
        definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 5 }], dialect: 'base', formulas: [] },
      }],
    },
  }
  const batch = await prepareCfPaint(sheet)
  expect(batch).toBeDefined()
  expect(sheet.diagnostics?.some(d => d.feature === 'cf-paint-unavailable')).not.toBe(true)
  expect(batch?.layer.styleAt(0, 6, undefined)?.fillColor).toBe('FF00FF00')
  batch?.dispose()
})

it('going over work budget refuses per range, never for the whole sheet (I4)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const sheet: XlsxSheet = {
    name: 'S', sheetId: '1', merges: [], mergeRanges: [], cols: [],
    rows: [
      { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }, { row: 0, col: 1, ref: 'B1', value: 10, styleIndex: 0 }] },
    ],
    conditionalFormatting: {
      diagnostics: [],
      rules: [
        {
          id: 0, priority: 1, documentOrder: 0, stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }],
          writes: new Set(['fill']),
          dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FF00FF00' },
          definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 5 }], dialect: 'base', formulas: [] },
        },
        {
          id: 1, priority: 2, documentOrder: 1, stopIfTrue: false,
          ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 1, lastCol: 1 }],
          writes: new Set(['fill']),
          dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFF0000' },
          definition: { kind: 'cellIs', operator: 'greaterThan', operands: [{ kind: 'number', value: 5 }], dialect: 'base', formulas: [] },
        },
      ],
    },
  }
  const batch = await prepareCfPaint(sheet)
  expect(batch).toBeDefined()
  expect(batch?.layer.styleAt(0, 0, undefined)?.fillColor).toBe('FF00FF00')
  batch?.dispose()
})

it('merged sub-cells do not take anchor value in statistics, agreeing in resident and paged paths (N2)', async () => {
  const { prepareCfPaint } = await import('../src/xlsx/cf/render')
  const { createObjectWorkbookSource } = await import('../src/xlsx/view/source')
  const makeSheet = (): XlsxSheet => ({
    name: 'S', sheetId: '1', cols: [],
    merges: ['A2:A3'],
    mergeRanges: [{ minRow: 1, maxRow: 2, minCol: 0, maxCol: 0 }],
    rows: [
      { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 1, styleIndex: 0 }] },
      { index: 1, cells: [{ row: 1, col: 0, ref: 'A2', value: 5, styleIndex: 0 }] },
      { index: 3, cells: [{ row: 3, col: 0, ref: 'A4', value: 2, styleIndex: 0 }] },
    ],
    conditionalFormatting: {
      diagnostics: [],
      rules: [{
        id: 0, priority: 1, documentOrder: 0, stopIfTrue: false,
        ranges: [{ firstRow: 0, lastRow: 3, firstCol: 0, lastCol: 0 }],
        writes: new Set(['fill']),
        dxf: { 'fill.pattern': 'solid', 'fill.bg': 'FFFF0000' },
        definition: { kind: 'predicate', type: 'duplicateValues', dialect: 'base', formulas: [] },
      }],
    },
  })

  // 1. Resident path
  const sheetResident = makeSheet()
  const batchResident = await prepareCfPaint(sheetResident)
  expect(batchResident).toBeDefined()
  // A2 has value 5 and is not a duplicate of itself (merged A2:A3 subcell does not count as 5 in statistics)
  const residentStyle = batchResident?.layer.styleAt(1, 0, undefined)
  expect(residentStyle?.fillColor).toBeUndefined()
  batchResident?.dispose()

  // 2. Paged path (storeCfSource)
  const sheetPaged = makeSheet()
  const store = createObjectWorkbookSource({ sheets: [sheetPaged] }, 'paged-test').store('1')!
  const batchPaged = await prepareCfPaint(sheetPaged, { store })
  expect(batchPaged).toBeDefined()
  const pagedStyle = batchPaged?.layer.styleAt(1, 0, undefined)
  expect(pagedStyle?.fillColor).toBeUndefined()
  batchPaged?.dispose()
})
