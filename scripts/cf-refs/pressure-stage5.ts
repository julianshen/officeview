/**
 * Task 39: the Stage 5 resource scenarios `audit --stage 5` runs (formula batches over synthetic sources). Each returns
 * what it observed; the audit and tests/xlsx-cf-stage5.test.ts check the bounds. Every scenario disposes what it
 * prepared and reports the memory still charged (`netBytes`), which must be 0.
 *
 * The batch helpers come from tests/helpers/cf-formula.ts (the same compile and batch wiring the Task 34-37 tests use).
 */
// functions.ts first: the engine import cycle (functions.ts <-> evaluator.ts <-> cond.ts) needs FUNCTIONS defined first.
import '../../src/xlsx/formula/functions'
import { batchOf, formulaRule, literalOf, type CellInput, type FakeSource } from '../../tests/helpers/cf-formula'
import { createCfMemory, createCfUnitCeiling, createCfWorkBudget, CF_LIMITS, type CfMemory, type CfWorkBudget } from '../../src/xlsx/cf/limits'
import { createCfStatistics } from '../../src/xlsx/cf/statistics'
import type { CfCallWorkBreakdown } from '../../src/xlsx/cf/formula'
import type { CfRule, CfValuePage, SourceRect } from '../../src/xlsx/cf/model'
import type { CfWorkbookSource } from '../../src/xlsx/cf/dependencies'

const MAX_ROW = 1_048_575
const PAGE = 256

/**
 * A lazily generated column-A source: rows 0..populated-1 hold `valueAt(row)`, nothing else is stored, so a declared
 * whole-column rule meets the populated bound only. Pages are generated on demand (a 1M-row column is never materialized).
 */
export function syntheticColumnSource(populated: number, valueAt: (row: number) => CellInput): FakeSource {
  const pageCalls: FakeSource['pageCalls'] = []
  const source: FakeSource = {
    pageCalls, boundsCalls: 0, valueGen: 1, dependencyGen: 0, unavailable: new Set(),
    stamp: () => ({ owner: 'audit-s5', sheetId: 's1', workbookValueGen: source.valueGen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1 }),
    async bounds() {
      source.boundsCalls++
      return { firstRow: 0, lastRow: Math.max(0, populated - 1), firstCol: 0, lastCol: 0 }
    },
    async *pages(_id, rects, _budget, signal) {
      const call = { rects, yielded: 0 }
      pageCalls.push(call)
      for (const r of [...rects].sort((a, b) => a.firstRow - b.firstRow)) {
        if (r.firstCol > 0) continue
        const last = Math.min(r.lastRow, populated - 1)
        for (let row = r.firstRow; row <= last; row += PAGE) {
          if (signal.aborted) throw new Error('CF cancelled')
          const end = Math.min(last, row + PAGE - 1)
          const n = end - row + 1
          const rows = new Uint32Array(n), cols = new Uint16Array(n), values = []
          for (let i = 0; i < n; i++) { rows[i] = row + i; values.push(literalOf(valueAt(row + i))) }
          call.yielded += n
          const page: CfValuePage = { rows, cols, values }
          yield { value: page, bytes: 64 * n, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
        }
      }
    },
  }
  return source
}

export interface Stage5PressureSample {
  name: string
  netBytes: number
  peakBytes: number
  wallMs: number
}

const col = (firstRow: number, lastRow: number): SourceRect => ({ firstRow, lastRow, firstCol: 0, lastCol: 0 })
const outcomeKinds = (o: { kind: string; reason?: string }): string => (o.kind === 'ready' ? 'ready' : `${o.kind}:${o.reason ?? ''}`)

/** 100 formula rules over a declared 1M-row A:A, 100 populated cells: one shared snapshot for all 100 rules. */
export async function pressureFormulaSharedSnapshot(): Promise<Stage5PressureSample & { rules: number; attempts: number; snapshotEntries: number; targetCells: number; ready: number }> {
  const started = Date.now()
  const memory = createCfMemory()
  const source = syntheticColumnSource(100, row => row % 10)
  const rules = Array.from({ length: 100 }, (_, i) => formulaRule(`A1>${i % 10}`, `A1:A${MAX_ROW + 1}`))
  const targets = Array.from({ length: 5 }, (_, i) => col(i * 8, i * 8 + 7))
  const batch = await batchOf(rules, source, { memory, targets, statistics: createCfStatistics(memory) })
  let ready = 0
  for (let t = 0; t < rules.length; t++) for (const r of targets) for (let row = r.firstRow; row <= r.lastRow; row++) if (batch.evaluate(t, row, 0).kind === 'ready') ready++
  const sample = { attempts: batch.attempts, snapshotEntries: batch.snapshotEntries }
  batch.dispose()
  return { name: 'formula-100-rules-1M-declared-shared-snapshot', rules: rules.length, ...sample, targetCells: 40, ready, netBytes: memory.currentBytes, peakBytes: memory.peakBytes, wallMs: Date.now() - started }
}

/**
 * SUM($A:$A)>0 and COUNTIF($A:$A,A2)>1 on 5,000 visible cells over a declared 1M-row A:A with 10,000 populated cells:
 * hoisted, the scan clamped to the populated bound, one chunked scan per rule (the session memoises it across batches),
 * every per-cell call within the 128-unit checkpoint. The same rules over ~1M populated rows refuse at the scan budget;
 * a relative-range SUM(A1:A200) refuses before any call is evaluated.
 */
export async function pressureFormulaHoists(): Promise<Stage5PressureSample & {
  visibleCells: number; columnScans: number; clampedScans: number; maxCallUnits: number; calls: number; refused: number
  millionOutcomes: string[]; millionCalls: number; relativeOutcomes: string[]; relativeCalls: number
}> {
  const started = Date.now()
  const memory = createCfMemory()
  const session = createCfStatistics(memory)
  const source = syntheticColumnSource(10_000, row => (row * 37) % 997)
  const rules: CfRule[] = [formulaRule('SUM($A:$A)>0', `A2:A${MAX_ROW + 1}`), formulaRule('COUNTIF($A:$A,A2)>1', `A2:A${MAX_ROW + 1}`)]
  const calls: CfCallWorkBreakdown[] = []
  const scanBudget = (): CfWorkBudget => createCfWorkBudget({ scan: true })
  let refused = 0, visibleCells = 0
  // 5,000 visible cells A2:A5001 in paint batches of 250 cells.
  for (let first = 1; first <= 5_000; first += 250) {
    const target = col(first, first + 249)
    const batch = await batchOf(rules, source, { memory, statistics: session, targets: [target], ceiling: createCfUnitCeiling(), scanBudget, onCall: b => { calls.push(b) } })
    for (let row = target.firstRow; row <= target.lastRow; row++) {
      visibleCells++
      for (let t = 0; t < rules.length; t++) if (batch.evaluate(t, row, 0).kind !== 'ready') refused++
    }
    batch.dispose()
  }
  const wholeColumn = source.pageCalls.filter(c => c.rects.length === 1 && c.rects[0]!.firstRow === 0 && c.rects[0]!.lastRow >= 9_999)
  const clampedScans = wholeColumn.filter(c => c.rects[0]!.lastRow === 9_999).length
  session.dispose()
  const net1 = memory.currentBytes, peak1 = memory.peakBytes

  // About a million populated rows: the scan budget refuses each rule (one diagnostic each), nothing stays charged.
  const bigMemory = createCfMemory()
  const big = syntheticColumnSource(MAX_ROW + 1, row => row % 997)
  const bigCalls: CfCallWorkBreakdown[] = []
  const bigBatch = await batchOf(rules, big, { memory: bigMemory, statistics: createCfStatistics(bigMemory), targets: [col(1, 256)], scanBudget, onCall: b => { bigCalls.push(b) } })
  const millionOutcomes = rules.map((_, t) => outcomeKinds(bigBatch.evaluate(t, 1, 0)))
  bigBatch.dispose()

  // A relative-range aggregate: 200 cells per call is over the 128-unit checkpoint, refused before any call.
  const relMemory = createCfMemory()
  const relCalls: CfCallWorkBreakdown[] = []
  const rel = await batchOf([formulaRule('SUM(A1:A200)>0', 'B1:B5000')], source, { memory: relMemory, targets: [{ firstRow: 0, lastRow: 249, firstCol: 1, lastCol: 1 }], onCall: b => { relCalls.push(b) } })
  const relativeOutcomes = [...new Set(Array.from({ length: 250 }, (_, row) => outcomeKinds(rel.evaluate(0, row, 1))))]
  rel.dispose()
  return {
    name: 'formula-hoists-5000-visible-over-1M-declared',
    visibleCells, columnScans: wholeColumn.length, clampedScans, maxCallUnits: Math.max(0, ...calls.map(c => c.total)), calls: calls.length, refused,
    millionOutcomes, millionCalls: bigCalls.length, relativeOutcomes, relativeCalls: relCalls.length,
    netBytes: net1 + bigMemory.currentBytes + relMemory.currentBytes, peakBytes: Math.max(peak1, bigMemory.peakBytes, relMemory.peakBytes), wallMs: Date.now() - started,
  }
}

/** The dependency snapshot at its cap (CF_LIMITS.dependencyCells) admits; one cell more refuses, with nothing left charged. */
export async function pressureFormulaSnapshotCap(): Promise<Stage5PressureSample & { cap: number; atCap: string; overCap: string; atCapEntries: number }> {
  const started = Date.now()
  const cap = CF_LIMITS.dependencyCells
  const source = syntheticColumnSource(cap + 1, row => row % 7)
  const run = async (cells: number, memory: CfMemory): Promise<{ outcome: string; entries: number }> => {
    const rule = formulaRule('A1>3', `A1:A${cap + 1}`)
    const batch = await batchOf([rule], source, { memory, targets: [col(0, cells - 1)] })
    const out = { outcome: outcomeKinds(batch.outcome(0)), entries: batch.snapshotEntries }
    batch.dispose()
    return out
  }
  const m1 = createCfMemory({ limitBytes: 256 * 1024 ** 2 }), m2 = createCfMemory({ limitBytes: 256 * 1024 ** 2 })
  const at = await run(cap, m1)
  const over = await run(cap + 1, m2)
  return { name: 'formula-snapshot-cap', cap, atCap: at.outcome, overCap: over.outcome, atCapEntries: at.entries, netBytes: m1.currentBytes + m2.currentBytes, peakBytes: Math.max(m1.peakBytes, m2.peakBytes), wallMs: Date.now() - started }
}

/** A refused footprint (a relative 200-cell range) beside a ready rule in one batch: the ready rule still evaluates. */
export async function pressureFormulaRefusedBesideReady(): Promise<Stage5PressureSample & { readyCells: number; refusedCells: number; refusedReasons: string[] }> {
  const started = Date.now()
  const memory = createCfMemory()
  const source = syntheticColumnSource(300, row => row % 5)
  const ready = formulaRule('A1>2', 'B1:B10')
  const refusedRule = formulaRule('SUM(A1:A200)>0', 'B1:B10')
  const batch = await batchOf([ready, refusedRule], source, { memory, targets: [{ firstRow: 0, lastRow: 9, firstCol: 1, lastCol: 1 }] })
  let readyCells = 0, refusedCells = 0
  const reasons = new Set<string>()
  for (let row = 0; row < 10; row++) {
    if (batch.evaluate(0, row, 1).kind === 'ready') readyCells++
    const o = batch.evaluate(1, row, 1)
    if (o.kind !== 'ready') { refusedCells++; reasons.add(outcomeKinds(o)) }
  }
  batch.dispose()
  return { name: 'formula-refused-footprint-beside-ready', readyCells, refusedCells, refusedReasons: [...reasons], netBytes: memory.currentBytes, peakBytes: memory.peakBytes, wallMs: Date.now() - started }
}

type WorkbookSheets = Record<string, { source: FakeSource; state?: 'ready' | 'pending' }>
/** A CfWorkbookSource over synthetic sheets (Task 36 seam): names are case-insensitive, ids are the names. */
function workbookOf(sheets: WorkbookSheets): CfWorkbookSource {
  const byLower = new Map(Object.keys(sheets).map(n => [n.toLowerCase(), n] as const))
  return {
    sheetIdOf: name => byLower.get(name.toLowerCase()),
    sheet: id => {
      const s = sheets[id]
      return s ? { sheetId: id, name: id, values: s.source, state: () => s.state ?? 'ready', valueGen: () => 0 } : undefined
    },
  }
}

/**
 * 16 referenced sheets evaluate; a formula naming 17 refuses at compile, and two rules whose union reaches 18 refuse the
 * one that crosses 16 in the batch (the other still evaluates).
 */
export async function pressureFormulaReferencedSheets(): Promise<Stage5PressureSample & { at16: string; compile17: string; union18: string[] }> {
  const started = Date.now()
  const memory = createCfMemory()
  const names = Array.from({ length: 18 }, (_, i) => `S${i + 2}`)
  const sheets: WorkbookSheets = Object.fromEntries(names.map(n => [n, { source: syntheticColumnSource(1, () => 1) }]))
  const sum = (list: readonly string[]): string => list.map(n => `${n}!$A$1`).join('+')
  const host = syntheticColumnSource(4, () => 1)
  const workbook = workbookOf(sheets)
  const sixteen = await batchOf([formulaRule(`${sum(names.slice(0, 16))}>0`, 'B1:B2')], host, { memory, workbook, targets: [{ firstRow: 0, lastRow: 1, firstCol: 1, lastCol: 1 }] })
  const at16 = outcomeKinds(sixteen.evaluate(0, 0, 1))
  sixteen.dispose()
  let compile17 = 'compiled'
  try { formulaRule(`${sum(names.slice(0, 17))}>0`, 'B1:B2') } catch (e) { compile17 = (e as Error).message.replace(/^compile refused .*: /, 'refused:') }
  const union = await batchOf([formulaRule(`${sum(names.slice(0, 9))}>0`, 'B1:B2'), formulaRule(`${sum(names.slice(9, 18))}>0`, 'B1:B2')], host, { memory, workbook, targets: [{ firstRow: 0, lastRow: 1, firstCol: 1, lastCol: 1 }] })
  const union18 = [outcomeKinds(union.outcome(0)), outcomeKinds(union.outcome(1))]
  union.dispose()
  return { name: 'formula-referenced-sheets-16-17', at16, compile17, union18, netBytes: memory.currentBytes, peakBytes: memory.peakBytes, wallMs: Date.now() - started }
}

/**
 * 20 open/read/cancel/close cycles over a cross-sheet rule: each cycle prepares a ready read, a read of a pending
 * referenced sheet, and a read cancelled (closed) while the referenced sheet's page read is in flight. The ledger is
 * checked after every transition; `maxNet` is the largest charge seen between a close and the next open (must be 0).
 */
export async function pressureFormulaCrossSheetCycles(rounds = 20): Promise<Stage5PressureSample & { cycles: number; ready: number; pending: number; cancelled: number; maxNetBetween: number }> {
  const started = Date.now()
  const memory = createCfMemory()
  let ready = 0, pending = 0, cancelled = 0, maxNetBetween = 0
  const host = syntheticColumnSource(10, () => 1)
  const targets = [{ firstRow: 0, lastRow: 9, firstCol: 1, lastCol: 1 }]
  for (let i = 0; i < rounds; i++) {
    const rule = formulaRule('Sheet2!A1+$A1>1', 'B1:B10')
    const ok = await batchOf([rule], host, { memory, workbook: workbookOf({ Sheet2: { source: syntheticColumnSource(10, () => 1) } }), targets })
    if (ok.outcome(0).kind === 'ready' && ok.evaluate(0, 0, 1).kind === 'ready') ready++
    ok.dispose()
    maxNetBetween = Math.max(maxNetBetween, memory.currentBytes)
    const wait = await batchOf([rule], host, { memory, workbook: workbookOf({ Sheet2: { source: syntheticColumnSource(10, () => 1), state: 'pending' } }), targets })
    if (wait.evaluate(0, 0, 1).kind !== 'ready') pending++
    wait.dispose()
    maxNetBetween = Math.max(maxNetBetween, memory.currentBytes)
    // Close during the referenced sheet's page read: its pages() blocks until the preparation is aborted.
    const ac = new AbortController()
    let reading!: () => void
    const started = new Promise<void>(res => { reading = res })
    const blocking = syntheticColumnSource(10, () => 1)
    const inner = blocking.pages.bind(blocking)
    blocking.pages = async function* (id, rects, budget, signal) {
      reading()
      await new Promise<void>(res => signal.aborted ? res() : signal.addEventListener('abort', () => res(), { once: true }))
      yield* inner(id, rects, budget, signal)
    }
    const run = batchOf([rule], host, { memory, signal: ac.signal, workbook: workbookOf({ Sheet2: { source: blocking } }), targets })
    await started
    ac.abort()
    try { (await run).dispose() } catch { /* cancelled: the preparation disposed what it held */ }
    cancelled++
    maxNetBetween = Math.max(maxNetBetween, memory.currentBytes)
  }
  return { name: 'formula-cross-sheet-20-cycles', cycles: rounds, ready, pending, cancelled, maxNetBetween, netBytes: memory.currentBytes, peakBytes: memory.peakBytes, wallMs: Date.now() - started }
}

/** The per-call checkpoint every formula call must stay within (CF_LIMITS.checkpointUnits). */
export const CHECKPOINT_UNITS: number = CF_LIMITS.checkpointUnits

/**
 * Task 39 fix round 1 (F5): the no-formula baseline keeps its conditional formatting (literal `cellIs` rules only) and
 * runs every production path that prepares CF for it: parse (resident), static paint with the document's evaluation
 * policy and sheets, and the text scope. It counts formula compiles, dependency-snapshot preparations and clock reads
 * (`cfFormulaActivity`), which must all stay 0. The positive control (one expression rule added) must move all three.
 */
type NoFormulaCounts = { compiles: number; snapshots: number; clockReads: number; painted: number }
export async function pressureNoFormulaBaseline(): Promise<{ literal: NoFormulaCounts; control: NoFormulaCounts; workerLiteral: NoFormulaCounts; workerControl: NoFormulaCounts }> {
  const { cfFormulaActivity } = await import('../../src/xlsx/cf/formula')
  const { workbookBytes, cfBlock, cellIsRule, exprRule, load } = await import('../../tests/helpers/cf-views')
  const { prepareCfPaint } = await import('../../src/xlsx/cf/render')
  const { prepareCfTextScope } = await import('../../src/xlsx/cf/text')
  const { cfDocumentEvaluationOf } = await import('../../src/xlsx/cf/formula')
  const measure = async (bytes: Uint8Array): Promise<{ compiles: number; snapshots: number; clockReads: number; painted: number }> => {
    const before = cfFormulaActivity()
    const { doc, sheet } = await load(bytes, 'cached')
    const evaluation = cfDocumentEvaluationOf(doc)
    let painted = 0
    for (let i = 0; i < 3; i++) {
      const batch = await prepareCfPaint(sheet, { evaluation, valuePolicy: 'cached', documentSheets: doc.sheets })
      for (let row = 1; row <= 3; row++) if (batch?.layer.styleAt(row, 1, undefined)?.fillColor) painted++
      batch?.dispose()
    }
    const budget = { reserve: () => {}, release: () => {}, retain: (_t: unknown, _b: number, commit: () => void) => commit(), settle: () => {} }
    const scope = await prepareCfTextScope(sheet, { rects: [{ r0: 1, r1: 3, c0: 1, c1: 1 }], signal: new AbortController().signal, locale: 'en-US', budget: budget as never, evaluation, documentSheets: doc.sheets })
    scope?.dispose()
    const after = cfFormulaActivity()
    return { compiles: after.compiles - before.compiles, snapshots: after.snapshots - before.snapshots, clockReads: after.clockReads - before.clockReads, painted }
  }
  // Final review F3: the worker (paged) route, three preparations through the real paged session. A clock read is any call
  // of the injected `now` (the worker's only clock) plus the evaluation context's own counter.
  const { openImportSession } = await import('../../src/xlsx/storage/pagedImport')
  const { createPagedWorkbookSource } = await import('../../src/xlsx/view/pagedSource')
  const { createCfMemory } = await import('../../src/xlsx/cf/limits')
  const measureWorker = async (bytes: Uint8Array): Promise<NoFormulaCounts> => {
    const before = cfFormulaActivity()
    let nowCalls = 0, painted = 0
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: `no-formula-${Math.random()}`, profile: 'D' })
    const source = createPagedWorkbookSource(session)
    try {
      await session.done
      const sheetId = source.sheets[0]!.sheetId
      const store = source.store(sheetId)
      for (let i = 0; i < 3; i++) {
        const result = await store.readConditionalFormatting!({ rects: [{ r0: 1, r1: 3, c0: 1, c1: 1 }], stamp: { owner: 'no-formula', sheetId, workbookValueGen: store.stamp.valueGen, cfGen: 0, viewGen: 0, structureGen: 0, clockGen: 0 }, grantBytes: 8 * 1024 * 1024 },
          createCfMemory(), { signal: new AbortController().signal, now: () => { nowCalls++; return Date.UTC(2026, 10, 16) } } as never)
        if (result?.kind === 'ready') painted += result.states.filter(s => s.properties['fill.bg'] !== undefined).length
      }
    } finally { await source.release() }
    const after = cfFormulaActivity()
    return { compiles: after.compiles - before.compiles, snapshots: after.snapshots - before.snapshots, clockReads: after.clockReads - before.clockReads + nowCalls, painted }
  }
  const cells = { A2: 1, A3: 2, A4: 3, B2: 1, B3: 5, B4: 9 }
  const literalBytes = workbookBytes(cells, cfBlock('B2:B4', cellIsRule('greaterThan', '4', 1, 0)))
  const controlBytes = workbookBytes(cells, cfBlock('B2:B4', cellIsRule('greaterThan', '4', 1, 0)) + cfBlock('C2:C4', exprRule('$A2&gt;1', 2, 1)))
  const literal = await measure(literalBytes)
  const control = await measure(controlBytes)
  const workerLiteral = await measureWorker(literalBytes)
  const workerControl = await measureWorker(controlBytes)
  return { literal, control, workerLiteral, workerControl }
}
