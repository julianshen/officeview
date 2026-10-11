/** S1.1b import limits: the quota ledger (design note §4.3; ledger parts of
 * RED tests 4 and 7) and the admission gates G1/G2 (RED test 19). */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { bytesRangeSource, openPart, readZipInventory, type RangeSource, type ZipLimits } from '../src/core/zip-stream'
import { PackageError } from '../src/core/package-error'
import { createQuotaLedger, POOLS, quotaLimitsFor, type QuotaLimits } from '../src/xlsx/storage/quota'
import { buildZip, type BuildEntry } from './helpers/zip-builder'

const MiB = 1024 * 1024
const LIMITS: ZipLimits = {
  maxEntries: 4096, maxCentralDirectoryBytes: 4 * MiB, maxNameBytes: 1024,
  maxPartExpandedBytes: 2048 * MiB, maxTotalExpandedBytes: 4096 * MiB, maxRatio: 2000,
}
const LEDGER_LIMITS: QuotaLimits = { ...quotaLimitsFor('D'), expandedQueue: 8 * MiB }

function allZero(current: Record<string, number>): boolean {
  return POOLS.every((pool) => current[pool] === 0)
}

async function single(entry: BuildEntry): Promise<{ src: RangeSource & { reads: number }; entry: Awaited<ReturnType<typeof readZipInventory>>['entries'][number] }> {
  const bytes = buildZip([entry]).bytes
  const inner = bytesRangeSource(bytes)
  const src = { size: inner.size, reads: 0, read(a: number, b: number, s?: AbortSignal) { src.reads++; return inner.read(a, b, s) } }
  const inv = await readZipInventory(src, LIMITS)
  return { src, entry: inv.entries[0] }
}

describe('quota ledger', () => {
  it('tracks current and peak per pool, releases idempotently and types the overflow', () => {
    const ledger = createQuotaLedger({ ...LEDGER_LIMITS, decodedBuffers: 100, pages: 50 }, { maxTotalExpandedBytes: 1000 })
    const a = ledger.reserve('decodedBuffers', 60)
    const b = ledger.reserve('decodedBuffers', 40)
    expect(ledger.snapshot().current.decodedBuffers).toBe(100)
    let err: unknown
    try { ledger.reserve('decodedBuffers', 1) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('resource-limit')
    try { ledger.reserve('pages', 51) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('quota')
    a.release()
    a.release()
    expect(ledger.snapshot().current.decodedBuffers).toBe(40)
    expect(ledger.snapshot().peak.decodedBuffers).toBe(100)
    b.release()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })

  it('charges expanded bytes against the total and holds only the latest step in the queue', () => {
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 200_000 })
    ledger.charge(65_536)
    ledger.charge(65_536)
    expect(ledger.snapshot().current.expandedQueue).toBe(65_536)
    expect(ledger.snapshot().totalExpanded).toBe(131_072)
    const part = ledger.partBudget('xl/a.xml', 10)
    part.charge(10)
    expect(() => part.charge(1)).toThrow(PackageError)
    expect(() => ledger.charge(200_000)).toThrow(/expanded total/)
    ledger.drainQueue()
    expect(ledger.snapshot().current.expandedQueue).toBe(0)
  })

  it('releaseAll zeroes every pool and fails later charges with cancelled (AC-STORE-02)', () => {
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 1e9 })
    ledger.reserve('pages', 1000)
    ledger.reserve('sharedStrings', 10)
    ledger.charge(5)
    ledger.releaseAll()
    const snap = ledger.snapshot()
    expect(allZero(snap.current)).toBe(true)
    expect(snap.released).toBe(true)
    let err: unknown
    try { ledger.charge(1) } catch (e) { err = e }
    expect((err as PackageError).code).toBe('cancelled')
  })
})

describe('ledger through openPart (ledger parts of RED tests 4 and 7)', () => {
  it('a 64 MiB zero bomb under the 8 MiB ledger stops at limit + one output buffer; the ledger returns to zero', async () => {
    const size = 64 * MiB
    const raw = new Uint8Array(deflateRawSync(new Uint8Array(size)))
    const { src, entry } = await single({ name: 'bomb.xml', raw, uncompressedSize: size, crc: 0 })
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 8 * MiB })
    let yielded = 0
    let err: unknown
    try {
      for await (const chunk of openPart(src, entry, { signal: new AbortController().signal, budget: ledger })) yielded += chunk.length
    } catch (e) { err = e }
    expect((err as PackageError).code).toBe('resource-limit')
    expect(yielded).toBeLessThanOrEqual(8 * MiB + 65_536)
    expect(ledger.snapshot().peak.expandedQueue).toBeLessThanOrEqual(65_536)
    expect(ledger.snapshot().totalExpanded).toBeLessThanOrEqual(8 * MiB)
    ledger.drainQueue()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })

  it('cancellation mid-part surfaces cancelled, stops reads and releaseAll leaves every pool at zero', async () => {
    const data = new Uint8Array(3 * MiB)
    for (let i = 0, x = 7; i < data.length; i++) { x = (x * 1103515245 + 12345) >>> 0; data[i] = x >>> 24 }
    const { src, entry } = await single({ name: 'p.xml', data })
    const ledger = createQuotaLedger(LEDGER_LIMITS, { maxTotalExpandedBytes: 1e9 })
    const held = ledger.reserve('pages', 4 * MiB)
    void held
    const ac = new AbortController()
    let pulls = 0
    let readsAtAbort = 0
    let err: unknown
    try {
      for await (const chunk of openPart(src, entry, { signal: ac.signal, budget: ledger, rangeReadBytes: 65_536 })) {
        void chunk
        if (++pulls === 3) { ac.abort(); readsAtAbort = src.reads }
      }
    } catch (e) { err = e }
    expect((err as PackageError).code).toBe('cancelled')
    expect(src.reads).toBe(readsAtAbort)
    expect(ledger.snapshot().current.pages).toBe(4 * MiB)
    ledger.releaseAll()
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// Admission gates (RED test 19)

import { admit, admissionTotal, classifyPart, inventoryLimits, predictedPeak, verifyObjectCandidate, type ObjectDecision } from '../src/xlsx/storage/admission'
import { importLimitsFor, type ImportLimitOverrides } from '../src/xlsx/storage/importSession'
import { countingEvents, newAdmissionCounters } from '../src/xlsx/storage/sheetRecords'
import { tokenizeBytes } from '../src/core/xml-stream'
import type { ProfileId } from '../src/xlsx/view/contracts'
import { formulaRows, metaRows, repeat, SS_NS, valueCells, xlsxBytes, xlsxEntries, type PackageSpec } from './helpers/xlsx-package'

async function g2(bytes: Uint8Array, profile: ProfileId, overrides: ImportLimitOverrides = {}): Promise<ObjectDecision | { mode: 'g1'; result: ReturnType<typeof admit> }> {
  const limits = importLimitsFor(profile, overrides)
  const src = bytesRangeSource(bytes)
  const inv = await readZipInventory(src, inventoryLimits(limits))
  const gate1 = admit(profile, bytes.length, inv, limits)
  if (gate1.mode !== 'object-candidate') return { mode: 'g1', result: gate1 }
  return verifyObjectCandidate(src, inv, limits, new AbortController().signal, { profile })
}

function decision(d: Awaited<ReturnType<typeof g2>>): 'object' | 'paged' | 'refuse' {
  if (d.mode === 'g1') return d.result.mode === 'paged' ? 'paged' : d.result.mode === 'refuse' ? 'refuse' : 'object'
  return d.mode
}

describe('part classification', () => {
  it('tokenizes only .xml/.rels; media are excluded; drawing-family parts are flagged', () => {
    expect(classifyPart('xl/worksheets/sheet1.xml')).toEqual({ cls: 'xml', drawing: false })
    expect(classifyPart('xl/drawings/_rels/drawing1.xml.rels')).toEqual({ cls: 'xml', drawing: true })
    expect(classifyPart('xl/charts/chart1.xml')).toEqual({ cls: 'xml', drawing: true })
    expect(classifyPart('xl/drawings/vmlDrawing1.vml')).toEqual({ cls: 'binary', drawing: true })
    expect(classifyPart('xl/printerSettings/printerSettings1.bin')).toEqual({ cls: 'binary', drawing: false })
    expect(classifyPart('xl/media/image1.png').cls).toBe('media')
    expect(classifyPart('docProps/thumbnail.jpeg').cls).toBe('media')
  })
})

describe('how the measured shapes fall when every guard is applied (§4.3 table)', () => {
  const shapes: Array<[string, () => PackageSpec, 'object' | 'paged', 'object' | 'paged']> = [
    ['empty sheet', () => ({}), 'object', 'object'],
    ['16,384 cols', () => ({ beforeSheetData: `<cols>${repeat(16384, (i) => `<col min="${i + 1}" max="${i + 1}" width="9" customWidth="1"/>`)}</cols>` }), 'object', 'object'],
    ['31,250 metadata-only rows', () => ({ sheetData: metaRows(31_250) }), 'object', 'object'],
    ['30k unknown elements with 16 attributes', () => ({ afterSheetData: `<extLst>${repeat(30_000, () => `<x ${repeat(16, (k) => `a${k}="1" `)}/>`)}</extLst>` }), 'object', 'paged'],
    ['80k merges', () => ({ afterSheetData: `<mergeCells>${repeat(80_000, (i) => `<mergeCell ref="A${2 * i + 1}:B${2 * i + 2}"/>`)}</mergeCells>` }), 'paged', 'paged'],
    ['30k cellXfs records', () => ({ styles: `<cellXfs>${repeat(30_000, () => '<xf numFmtId="0"/>')}</cellXfs>` }), 'paged', 'paged'],
    ['one 2 M-character inline string', () => ({ sheetData: `<row r="1"><c r="A1" t="inlineStr"><is><t>${'x'.repeat(2_000_000)}</t></is></c></row>` }), 'paged', 'paged'],
    ['100k shared strings', () => ({ sharedStrings: Array.from({ length: 100_000 }, (_, i) => `s${i}`) }), 'paged', 'paged'],
    ['125k metadata-only rows (counterexample)', () => ({ sheetData: metaRows(125_000) }), 'paged', 'paged'],
  ]
  for (const [name, spec, onD, onM] of shapes) {
    it(`${name}: ${onD} on D, ${onM} on M`, async () => {
      const bytes = xlsxBytes(spec())
      expect(decision(await g2(bytes, 'D'))).toBe(onD)
      expect(decision(await g2(bytes, 'M'))).toBe(onM)
    })
  }

  it('the 125k-row counterexample predicts about 287.9 MiB and trips the envelope on both profiles', async () => {
    const d = await g2(xlsxBytes({ sheetData: metaRows(125_000) }), 'D', { object: { rows: 1e9 } })
    expect(d.mode).toBe('paged')
    if (d.mode !== 'paged') return
    expect(d.reason).toMatch(/predicted peak/)
  })

  it('80k merges, 30k style records and a 2 M-character string pass the envelope on D but are paged by their secondary cap', async () => {
    const reasons = await Promise.all([
      g2(xlsxBytes({ afterSheetData: `<mergeCells>${repeat(80_000, (i) => `<mergeCell ref="A${2 * i + 1}:B${2 * i + 2}"/>`)}</mergeCells>` }), 'D'),
      g2(xlsxBytes({ styles: `<cellXfs>${repeat(30_000, () => '<xf numFmtId="0"/>')}</cellXfs>` }), 'D'),
      g2(xlsxBytes({ sheetData: `<row r="1"><c r="A1" t="inlineStr"><is><t>${'x'.repeat(2_000_000)}</t></is></c></row>` }), 'D'),
    ])
    expect(reasons.map((r) => (r.mode === 'paged' ? r.reason.split(' ')[0] : r.mode))).toEqual(['merges', 'style', 'string'])
  })
})

describe('boundaries at cap - 1, cap, cap + 1 (D and M)', () => {
  // The envelope is zeroed so the secondary cap under test is the binding guard.
  const isolate: ImportLimitOverrides = { object: { envelope: { constantBytes: 0, perElement: 0, perAttribute: 0, perTextChar: 0 } } }
  const caps: Array<[string, (n: number) => PackageSpec, (p: ProfileId) => number]> = [
    ['rows', (n) => ({ sheetData: metaRows(n) }), (p) => importLimitsFor(p).object.rows],
    ['cells', (n) => ({ sheetData: valueCells(n) }), (p) => importLimitsFor(p).object.cells],
    ['unique strings', (n) => ({ sharedStrings: Array.from({ length: n }, () => 'a') }), (p) => importLimitsFor(p).object.uniqueStrings],
    ['merges', (n) => ({ afterSheetData: `<mergeCells>${repeat(n, (i) => `<mergeCell ref="A${2 * i + 1}:A${2 * i + 2}"/>`)}</mergeCells>` }), (p) => importLimitsFor(p).object.merges],
    ['style records', (n) => ({ styles: `<cellXfs>${repeat(n - 0, () => '<xf/>')}</cellXfs>` }), (p) => importLimitsFor(p).object.styleRecords],
    ['longest string', (n) => ({ sharedStrings: ['y'.repeat(n)] }), (p) => importLimitsFor(p).inlineStringSoftBytes],
  ]
  for (const profile of ['D', 'M'] as const) {
    for (const [name, spec, cap] of caps) {
      it(`${name} on ${profile}`, async () => {
        const c = cap(profile)
        const got = await Promise.all([c - 1, c, c + 1].map(async (n) => decision(await g2(xlsxBytes(spec(n)), profile, isolate))))
        expect(got).toEqual(['object', 'object', 'paged'])
      })
    }
  }

  for (const profile of ['D', 'M'] as const) {
    it(`the admission inequality on ${profile}: budget - 1, equal, + 1`, async () => {
      const bytes = xlsxBytes({ sheetData: formulaRows(30, (i) => `A${i + 1}+1`) })
      const probe = await g2(bytes, profile, { object: { budgetBytes: 2 ** 40 } })
      expect(probe.mode).toBe('object')
      if (probe.mode !== 'object') return
      const limits = importLimitsFor(profile)
      const need = predictedPeak(probe.counters, limits) + 30 * 2048 + limits.object.evaluation.maxArrayCells * 2000
      expect(admissionTotal(probe.counters, limits)).toBe(need)
      const policy = async (budgetBytes: number) => {
        const d = await g2(bytes, profile, { object: { budgetBytes } })
        return d.mode === 'object' ? d.valuePolicy : d.mode
      }
      // the evaluation reserve decides engine vs saved values (decision 2) ...
      expect([await policy(need + 1), await policy(need), await policy(need - 1)]).toEqual(['engine', 'engine', 'cached'])
      // ... and the element envelope alone decides object vs paged
      const peak = predictedPeak(probe.counters, limits)
      expect([await policy(peak + 1), await policy(peak), await policy(peak - 1)]).toEqual(['cached', 'cached', 'paged'])
    })
  }

  it('G1 declared caps: compressed bytes, entries, non-drawing total, one non-drawing part, one drawing part', async () => {
    const limits = importLimitsFor('D')
    const big = (name: string, size: number): BuildEntry => ({ name, data: new Uint8Array(size), method: 0 })
    const inv = async (entries: BuildEntry[]) => readZipInventory(bytesRangeSource(buildZip(entries).bytes), inventoryLimits(limits))
    const base = xlsxEntries({})
    const small = await inv(base)
    expect(admit('D', limits.object.compressedBytes, small, limits).mode).toBe('object-candidate')
    expect(admit('D', limits.object.compressedBytes + 1, small, limits).mode).toBe('paged')
    expect(admit('D', limits.paged.compressedBytes + 1, small, limits).mode).toBe('refuse')
    const many = await inv([...base, ...Array.from({ length: limits.object.entries + 1 - base.length }, (_, i) => big(`xl/x${i}.bin`, 1))])
    expect(admit('D', 1000, many, limits)).toMatchObject({ mode: 'paged', reason: expect.stringMatching(/entry count/) })
    const part = await inv([...base, big('xl/big.bin', limits.object.nonDrawingPartBytes + 1)])
    expect(admit('D', 1000, part, limits)).toMatchObject({ mode: 'paged', reason: expect.stringMatching(/xl\/big.bin/) })
    const drawingOk = await inv([...base, big('xl/drawings/drawing1.xml', limits.object.drawingPartBytes)])
    expect(admit('D', 1000, drawingOk, limits).mode).toBe('object-candidate')
    const drawing = await inv([...base, big('xl/drawings/drawing1.xml', limits.object.drawingPartBytes + 1)])
    expect(admit('D', 1000, drawing, limits)).toMatchObject({ mode: 'paged', reason: expect.stringMatching(/drawing part/) })
    const total = await inv([...base, big('xl/a.bin', 3 * MiB), big('xl/b.bin', 3 * MiB)])
    expect(admit('D', 1000, total, limits)).toMatchObject({ mode: 'paged', reason: expect.stringMatching(/non-drawing parts total/) })
    const media = await inv([...base, big('xl/media/image1.png', 20 * MiB)])
    expect(admit('D', 1000, media, limits).mode).toBe('object-candidate')
  })

  it('refusals name the cap, the value and the part', async () => {
    const limits = importLimitsFor('M')
    const inv = await readZipInventory(bytesRangeSource(xlsxBytes({})), inventoryLimits(limits))
    const res = admit('M', limits.paged.compressedBytes + 5, inv, limits)
    expect(res.mode).toBe('refuse')
    if (res.mode !== 'refuse') return
    expect(res.failure.code).toBe('resource-limit')
    expect(res.failure.message).toContain(String(limits.paged.compressedBytes + 5))
    expect(res.failure.message).toContain(String(limits.paged.compressedBytes))
  })
})

describe('formula evaluation reserve (round-4 table)', () => {
  const cases: Array<[string, () => PackageSpec, number]> = [
    ['20 trivial uncached', () => ({ sheetData: formulaRows(20, (i) => `A${i + 1}+1`) }), 20],
    ['200 trivial uncached', () => ({ sheetData: formulaRows(200, (i) => `A${i + 1}+1`) }), 200],
    ['2,000 trivial uncached', () => ({ sheetData: formulaRows(2000, (i) => `A${i + 1}+1`) }), 2000],
    ['20 volatile (cached NOW())', () => ({ sheetData: formulaRows(20, () => 'NOW()', '', true) }), 20],
    ['20 cached under fullCalcOnLoad', () => ({ sheetData: formulaRows(20, (i) => `A${i + 1}+1`, '', true), fullCalcOnLoad: true }), 20],
    ['20 cached ca="1"', () => ({ sheetData: formulaRows(20, (i) => `A${i + 1}+1`, ' ca="1"', true) }), 20],
  ]
  for (const [name, spec, n] of cases) {
    it(`${name} stays object (engine) on D and M with N = ${n}`, async () => {
      for (const profile of ['D', 'M'] as const) {
        const d = await g2(xlsxBytes(spec()), profile)
        expect(d.mode).toBe('object')
        if (d.mode !== 'object') return
        expect(d.counters.loadEvaluatedFormulas).toBe(n)
        expect(d.valuePolicy).toBe('engine')
        const limits = importLimitsFor(profile)
        expect(d.evaluation.maxArrayCells).toBe(limits.object.evaluation.maxArrayCells)
        expect(d.evaluation.maxMaterializedCells).toBe(Math.floor((limits.object.budgetBytes - d.predictedPeak - n * 2048) / 2000))
        // invariant (rule 4): accounted allocation stays within the object budget
        expect(d.predictedPeak + n * 2048 + d.evaluation.maxMaterializedCells * 2000).toBeLessThanOrEqual(limits.object.budgetBytes)
        expect(d.admissionTotal).toBeLessThanOrEqual(limits.object.budgetBytes)
      }
    })
  }

  it('cached, non-volatile formulas are not load-evaluated', async () => {
    const d = await g2(xlsxBytes({ sheetData: formulaRows(50, (i) => `A${i + 1}+1`, '', true) }), 'D')
    expect(d.mode === 'object' && d.counters.loadEvaluatedFormulas).toBe(0)
  })

  it('20,000 trivial uncached formulas: saved values on D (the evaluation reserve does not fit), paged on M (envelope)', async () => {
    const bytes = xlsxBytes({ sheetData: formulaRows(20_000, (i) => `A${i + 1}+1`) })
    const d = await g2(bytes, 'D', { object: { rows: 1e9, cells: 1e9 } })
    expect(d.mode === 'object' && d.valuePolicy).toBe('cached')
    const m = await g2(bytes, 'M', { object: { rows: 1e9, cells: 1e9 } })
    expect(m.mode).toBe('paged')
  })
})

describe('compression ratio is a diagnostic, not a guard (orchestrator decision 4)', () => {
  it('a part deflated far beyond 200:1 with declared and actual sizes inside the caps stays object and is diagnosed', async () => {
    const bytes = xlsxBytes({ afterSheetData: `<extLst><ext uri="x">${'a'.repeat(1_500_000)}</ext></extLst>` })
    const limits = importLimitsFor('D')
    const inv = await readZipInventory(bytesRangeSource(bytes), inventoryLimits(limits))
    const sheet = inv.find('xl/worksheets/sheet1.xml')!
    expect(sheet.uncompressedSize / sheet.compressedSize).toBeGreaterThan(200)
    expect(admit('D', bytes.length, inv, limits).mode).toBe('object-candidate')
    const d = await g2(bytes, 'D')
    expect(d.mode).toBe('object')
    expect(d.mode === 'object' && d.counters.highRatioParts).toEqual(['xl/worksheets/sheet1.xml'])
  })
})

describe('verification pass bounds', () => {
  it('lying declared sizes (1 MiB declared, inflating far beyond) are stopped by actual counting', async () => {
    const huge = '<row r="1"/>'.repeat(400_000)
    const entries = xlsxEntries({ sheetData: huge })
    const layout = buildZip(entries)
    const forged = layout.bytes.slice()
    const sheetIndex = entries.findIndex((e) => e.name === 'xl/worksheets/sheet1.xml')
    new DataView(forged.buffer).setUint32(layout.entries[sheetIndex].centralOffset + 24, 1 * MiB, true)
    const d = await g2(forged, 'D', { object: { rows: 1e9 } })
    expect(d.mode).toBe('refuse') // the first byte beyond the declared size is a parse-error before it is yielded
    if (d.mode === 'refuse') expect(d.failure.code).toBe('parse-error')
  })

  it('the ledger peak stays within one output buffer and expanded accounting equals the declared sizes', async () => {
    const spec = { sheetData: metaRows(20_000) }
    const entries = xlsxEntries(spec)
    const bytes = buildZip(entries).bytes
    const limits = importLimitsFor('D')
    const src = bytesRangeSource(bytes)
    const inv = await readZipInventory(src, inventoryLimits(limits))
    const ledger = createQuotaLedger(quotaLimitsFor('D'), { maxTotalExpandedBytes: 1e9 })
    const d = await verifyObjectCandidate(src, inv, limits, new AbortController().signal, { ledger })
    expect(d.mode).toBe('object')
    const declared = inv.entries.reduce((n, e) => n + e.uncompressedSize, 0)
    expect(ledger.snapshot().totalExpanded).toBe(declared)
    expect(d.mode === 'object' && d.counters.expandedBytes).toBe(declared)
    expect(ledger.snapshot().peak.expandedQueue).toBeLessThanOrEqual(65_536)
    expect(allZero(ledger.snapshot().current)).toBe(true)
  })

  it('small files: sheetData or si inside mc:AlternateContent stays object (legacy rendering) and is recorded for a diagnostic (decision 5)', async () => {
    const mc = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    const wrapped = xlsxEntries({}).map((e) => e.name === 'xl/worksheets/sheet1.xml'
      ? { ...e, data: `<worksheet xmlns="${SS_NS}" ${mc}><mc:AlternateContent><mc:Choice Requires="x"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></mc:Choice></mc:AlternateContent></worksheet>` }
      : e)
    const d = await g2(buildZip(wrapped).bytes, 'D')
    expect(d.mode === 'object' && d.counters.alternateContentData).toEqual(['xl/worksheets/sheet1.xml'])

    const si = xlsxEntries({ sharedStrings: ['a'] }).map((e) => e.name === 'xl/sharedStrings.xml'
      ? { ...e, data: `<sst xmlns="${SS_NS}" ${mc}><mc:AlternateContent><mc:Choice Requires="x"><si><t>a</t></si></mc:Choice></mc:AlternateContent></sst>` }
      : e)
    const s = await g2(buildZip(si).bytes, 'M')
    expect(s.mode === 'object' && s.counters.alternateContentData).toEqual(['xl/sharedStrings.xml'])

    const merges = xlsxEntries({ afterSheetData: `<mc:AlternateContent ${mc}><mc:Choice Requires="x"><mergeCells><mergeCell ref="A1:B2"/></mergeCells></mc:Choice></mc:AlternateContent>` })
    const m = await g2(buildZip(merges).bytes, 'D')
    expect(m.mode).toBe('object')
    if (m.mode !== 'object') return
    expect(m.counters.alternateContentSubtrees).toBe(1)
    expect(m.counters.structuralIncomplete).toEqual(['merges'])
  })

  it('large mode keeps the rule: data inside mc:AlternateContent is a typed unsupported failure', async () => {
    const counters = newAdmissionCounters()
    const events = countingEvents(counters, 'xl/worksheets/sheet1.xml', { check() {}, alternateContentData: 'refuse' })
    const xml = `<worksheet xmlns:mc="m"><mc:AlternateContent><mc:Choice><sheetData/></mc:Choice></mc:AlternateContent></worksheet>`
    let err: unknown
    try {
      await tokenizeBytes((async function* () { yield new TextEncoder().encode(xml) })(), events)
    } catch (e) { err = e }
    expect((err as PackageError).code).toBe('unsupported')
  })

  it('tokenizer grammar failures (DOCTYPE, malformed XML) fall back to the legacy parser; bytes are still verified (decision 3)', async () => {
    const dtd = xlsxEntries({}).map((e) => e.name === 'xl/styles.xml' ? { ...e, data: '<!DOCTYPE x><styleSheet/>' } : e)
    const t = await g2(buildZip(dtd).bytes, 'D')
    expect(t.mode).toBe('object')
    if (t.mode !== 'object') return
    expect(t.legacyFallback).toMatchObject({ part: 'xl/styles.xml' })
    const declared = (await readZipInventory(bytesRangeSource(buildZip(dtd).bytes), inventoryLimits(importLimitsFor('D')))).entries.reduce((n, e) => n + e.uncompressedSize, 0)
    expect(t.counters.expandedBytes).toBe(declared)

    const bad = xlsxEntries({}).map((e) => e.name === 'xl/styles.xml' ? { ...e, data: '<styleSheet><a></styleSheet>' } : e)
    const d = await g2(buildZip(bad).bytes, 'D')
    expect(d.mode === 'object' && d.legacyFallback?.part).toBe('xl/styles.xml')
  })

  it('byte-level violations stay refusals even after a grammar fallback (CRC mismatch)', async () => {
    const entries = xlsxEntries({}).map((e) => e.name === 'xl/styles.xml' ? { ...e, data: '<!DOCTYPE x><styleSheet/>' } : e)
    const layout = buildZip(entries)
    const forged = layout.bytes.slice()
    const ws = entries.findIndex((e) => e.name === 'xl/worksheets/sheet1.xml')
    const at = layout.entries[ws].centralOffset + 16
    new DataView(forged.buffer).setUint32(at, (new DataView(forged.buffer).getUint32(at, true) ^ 1) >>> 0, true)
    const d = await g2(forged, 'D')
    expect(d.mode === 'refuse' && d.failure.code).toBe('parse-error')
  })

  it('binary parts are byte-counted, not tokenized', async () => {
    const d = await g2(xlsxBytes({ extra: [{ name: 'xl/printerSettings/printerSettings1.bin', data: new Uint8Array([0, 1, 2, 0x3c, 0xff]) }] }), 'D')
    expect(d.mode).toBe('object')
  })

  it('counters start at zero', () => {
    expect(newAdmissionCounters().elements).toBe(0)
  })
})
