/**
 * Stage 5 Task 34 test helpers: cell maps in A1 notation, a counting fake `CfValueSource`, compiled formulas, and an
 * independent engine oracle that evaluates a formula with `evaluateNode` over the FULL store (every cell, reference
 * services, no CF snapshot, no CF budget). The oracle is the expectation the CF evaluator must reproduce; it shares no
 * code with `src/xlsx/cf/dependencies.ts`.
 */
import { evaluateNode } from '../../src/xlsx/formula/evaluator'
import { parseFormula } from '../../src/xlsx/formula/parser'
import { createReferenceServices, type SparseCellStore, type StoredCell } from '../../src/xlsx/formula/refs'
import type { CellRef, ElementValue, EvaluationContext, EvaluationValue, FormulaError, ResolvedSemantics } from '../../src/xlsx/formula/types'
import { compileCfFormula, isCfFormulaRefusal } from '../../src/xlsx/cf/formula'
import { createCfMemory } from '../../src/xlsx/cf/limits'
import type { CfFormula, CfLiteral, CfStamp, CfValuePage, CfValueSource, SourceRect } from '../../src/xlsx/cf/model'

export type CellInput = number | string | boolean | null | { error: FormulaError } | { date: number }
export type Cells = Readonly<Record<string, CellInput>>

export function a1(ref: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!m) throw new Error(`bad ref ${ref}`)
  const col = [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
  return { row: Number(m[2]) - 1, col }
}
export function refOf(row: number, col: number): string {
  let c = col + 1, s = ''
  while (c > 0) { const r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = Math.floor((c - 1) / 26) }
  return `${s}${row + 1}`
}
export const rectOf = (range: string): SourceRect => {
  const [from, to = from] = range.split(':')
  const a = a1(from!), b = a1(to!)
  return { firstRow: Math.min(a.row, b.row), lastRow: Math.max(a.row, b.row), firstCol: Math.min(a.col, b.col), lastCol: Math.max(a.col, b.col) }
}

export function literalOf(v: CellInput): CfLiteral {
  if (v === null) return { kind: 'blank' }
  if (typeof v === 'number') return { kind: 'number', value: v }
  if (typeof v === 'string') return { kind: 'string', value: v }
  if (typeof v === 'boolean') return { kind: 'boolean', value: v }
  if ('error' in v) return { kind: 'error', value: v.error }
  return { kind: 'number', value: v.date, dateSerial: true }
}
function elementOf(v: CellInput): ElementValue {
  if (v === null || typeof v !== 'object') return v
  if ('error' in v) return { kind: 'formula-error', code: v.error }
  return v.date
}

/** Stored cells in row-major order, nulls dropped (a null input is an unstored blank). */
export function storedCells(cells: Cells): Array<{ row: number; col: number; value: CellInput }> {
  return Object.entries(cells).filter(([, v]) => v !== null).map(([ref, value]) => ({ ...a1(ref), value }))
    .sort((a, b) => a.row - b.row || a.col - b.col)
}

export interface FakeSource extends CfValueSource {
  /** pages() calls, and the cells each call yielded. */
  readonly pageCalls: Array<{ rects: readonly SourceRect[]; yielded: number }>
  boundsCalls: number
  valueGen: number
  /** Referenced-sheet generation sum (Ruling S5-K); 0 is the same as absent. */
  dependencyGen: number
  /** Cells whose read throws the resident source's unavailable-value error (`render.ts residentSource`). */
  unavailable: Set<string>
}
export function fakeSource(cells: Cells, options: { sheetId?: string; bounds?: SourceRect; pending?: boolean } = {}): FakeSource {
  const sheetId = options.sheetId ?? 's1'
  const list = storedCells(cells)
  const pageCalls: FakeSource['pageCalls'] = []
  const source: FakeSource = {
    pageCalls, boundsCalls: 0, valueGen: 1, dependencyGen: 0, unavailable: new Set(),
    stamp: () => ({ owner: 'fake', sheetId, workbookValueGen: source.valueGen, cfGen: 1, viewGen: 1, structureGen: 1, clockGen: 1, ...(source.dependencyGen ? { dependencyGen: source.dependencyGen } : {}) }),
    async bounds() {
      source.boundsCalls++
      if (options.pending) return undefined
      if (options.bounds) return options.bounds
      if (list.length === 0) return { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }
      return { firstRow: Math.min(...list.map(c => c.row)), lastRow: Math.max(...list.map(c => c.row)), firstCol: Math.min(...list.map(c => c.col)), lastCol: Math.max(...list.map(c => c.col)) }
    },
    async *pages(_id, rects, _budget, signal) {
      const call = { rects, yielded: 0 }
      pageCalls.push(call)
      for (const cell of list) {
        if (signal.aborted) throw new Error('CF cancelled')
        if (!rects.some(r => cell.row >= r.firstRow && cell.row <= r.lastRow && cell.col >= r.firstCol && cell.col <= r.lastCol)) continue
        if (source.unavailable.has(refOf(cell.row, cell.col))) throw new Error('CF unavailable formula value')
        call.yielded++
        const page: CfValuePage = { rows: Uint32Array.of(cell.row), cols: Uint16Array.of(cell.col), values: [literalOf(cell.value)] }
        yield { value: page, bytes: 64, release() {}, transfer() { return { kind: 'unresolved', reason: 'borrowed-page' } } }
      }
    },
  }
  return source
}
export const stampOf = (source: CfValueSource, sheetId = 's1'): CfStamp => source.stamp(sheetId)

export function compile(text: string, anchor: string | { row: number; col: number } = 'A1'): CfFormula {
  const at = typeof anchor === 'string' ? a1(anchor) : anchor
  const out = compileCfFormula(text, at, createCfMemory())
  if (isCfFormulaRefusal(out)) throw new Error(`compile refused ${text}: ${out.reason}`)
  return out.value
}

export const SEMANTICS: ResolvedSemantics = Object.freeze({
  dateSystem: '1900', unicode: { version: 2, source: 'standalone-default' }, locale: 'en-US', timeZone: 'UTC', epochNowMs: Date.UTC(2026, 10, 16),
}) as ResolvedSemantics

/**
 * The engine's own answer on the FULL store: `evaluateNode` with reference services over every stored cell and a
 * typed `getCellValue`, the way the workbook evaluator wires them (workbook.ts). Returns the raw value (never the public
 * blank-to-0 boundary) and the unsupported features the engine recorded.
 */
export function engineEval(text: string, cells: Cells, at: { row: number; col: number }, sheet = 'Sheet1'): { value: EvaluationValue; features: string[] } {
  return engineOf(cells, sheet)(text, at)
}
/** The same oracle over one fixed store, reused across many evaluations. */
export function engineOf(cells: Cells, sheet = 'Sheet1'): (text: string, at: { row: number; col: number }) => { value: EvaluationValue; features: string[] } {
  const map = new Map<string, ElementValue>()
  for (const c of storedCells(cells)) map.set(`${c.row}:${c.col}`, elementOf(c.value))
  const store: SparseCellStore = {
    read: address => { const v = map.get(`${address.row}:${address.col}`); return v === undefined ? undefined : { value: v, origin: 'input' } },
    *stored(): Iterable<StoredCell> {
      for (const [key, value] of map) { const [row, col] = key.split(':').map(Number); yield { address: { sheetId: 'sheet', row: row!, col: col! }, value, origin: 'input' } }
    },
  }
  const references = createReferenceServices({
    generation: () => 0, sheets: [{ sheetId: 'sheet', name: sheet, workbookIndex: 0 }],
    sheetIdOfName: n => (n.toLowerCase() === sheet.toLowerCase() ? 'sheet' : undefined), sheetNameOfId: id => (id === 'sheet' ? sheet : undefined),
    store, definedNames: [], tables: undefined,
  })
  return (text, at) => {
    const currentCell: CellRef = { row: at.row, col: at.col, absRow: false, absCol: false }
    const features = new Set<string>()
    const ctx: EvaluationContext = {
      typedValues: true, currentSheet: sheet, currentCell, semantics: SEMANTICS, references, unsupportedFeatures: features,
      nodeValues: new WeakMap(), flatArgs: new WeakMap(), functionWork: new WeakMap(),
      getCellValue: (_s, col, row) => map.get(`${row}:${col}`) ?? null,
    }
    const value = evaluateNode(parseFormula(text.replace(/^=/, '')), ctx)
    return { value, features: [...features].sort() }
  }
}

/* ------------------------------------------------------------------ */
/* CF batch helpers (Task 34): one or more formula rules over a source. */
/* ------------------------------------------------------------------ */
import { cfDocumentEvaluationOf, cfEvaluationContextFor, type CfEdgeMode, type CfEvaluationContext, type CfResult } from '../../src/xlsx/cf/formula'
import { prepareCfFormulaBatch, type CfFormulaBatch, type CfFormulaBatchOptions } from '../../src/xlsx/cf/dependencies'
import { createCfWorkBudget, type CfMemory } from '../../src/xlsx/cf/limits'
import { createCfStatistics } from '../../src/xlsx/cf/statistics'
import type { CfOutcome, CfRule } from '../../src/xlsx/cf/model'

export const EVALUATION: CfEvaluationContext = cfEvaluationContextFor(
  cfDocumentEvaluationOf({ calc: undefined, valuePolicy: 'cached', semantics: { dateSystem: '1900', timeZone: 'UTC' } } as never), () => Date.UTC(2026, 10, 16))

let nextRuleId = 1
/** An expression rule over `range` (anchored at its top-left), with an optional DXF. */
export function formulaRule(text: string, range: string, extra: Partial<Pick<CfRule, 'priority' | 'stopIfTrue' | 'dxf' | 'writes'>> = {}): CfRule {
  const ranges = range.split(' ').map(rectOf)
  const formula = compile(text, { row: ranges[0]!.firstRow, col: ranges[0]!.firstCol })
  const id = nextRuleId++
  return {
    id, priority: extra.priority ?? id, documentOrder: id, ranges, stopIfTrue: extra.stopIfTrue ?? false,
    writes: extra.writes ?? new Set(Object.keys(extra.dxf ?? {}).map(k => k.split('.')[0]!)), ...(extra.dxf ? { dxf: extra.dxf } : {}),
    definition: { kind: 'expression', formula, sourceType: 'expression', dialect: 'base', formulas: [text] },
  }
}
export const formulaOf = (rule: CfRule) => (rule.definition as { formula: CfFormula }).formula

export type BatchOptions = Partial<Omit<CfFormulaBatchOptions, 'source' | 'stamp' | 'sheetId'>> & { edge?: CfEdgeMode }
export async function batchOf(rules: readonly CfRule[], source: FakeSource, options: BatchOptions = {}): Promise<CfFormulaBatch> {
  const memory: CfMemory = options.memory ?? createCfMemory()
  const signal = options.signal ?? new AbortController().signal
  return prepareCfFormulaBatch(rules.map(rule => ({ rule, formula: formulaOf(rule) })), {
    sheetName: 'Sheet1', evaluation: EVALUATION, edge: 'ref-error', memory, signal,
    budget: createCfWorkBudget({ signal, scan: true, maxUnits: 10_000_000 }), statistics: createCfStatistics(memory), hoist: true,
    ...options, source, sheetId: 's1', stamp: source.stamp('s1'),
  })
}
/** One rule's outcome at one cell. */
export async function cfEval(text: string, range: string, cells: Cells, target: string, options: BatchOptions = {}): Promise<CfOutcome<CfResult>> {
  const rule = formulaRule(text, range)
  const batch = await batchOf([rule], fakeSource(cells), options)
  try { const t = a1(target); return batch.evaluate(0, t.row, t.col) } finally { batch.dispose() }
}
