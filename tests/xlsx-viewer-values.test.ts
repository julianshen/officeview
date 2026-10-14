/**
 * S1.3 value authority (spec §4.4, plan S1.3): one declared value policy per
 * document, `engine` unchanged for small documents, `cached` never evaluates,
 * typed `formula-value-unavailable` refusals limited to key/value domains and
 * the saved-values status only for cached documents. Adapter parity under one
 * policy lives in xlsx-viewer-adapter-parity.test.ts.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import * as workbook from '../src/xlsx/formula/workbook'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { DocumentSession, type FilterEvaluator } from '../src/worker/viewerSession'
import { initialIdentity } from '../src/worker/viewerProtocol'
import { sortRowOrder } from '../src/xlsx/view/sort'
import {
  LANE_BLANK, LANE_BOOLEAN, LANE_ERROR, LANE_NUMBER, LANE_STRING, LANE_UNAVAILABLE,
  type FilterSpec, type SortSpec, type SourceStore,
} from '../src/xlsx/view/contracts'
import {
  FORMULA_VALUE_UNAVAILABLE, SAVED_VALUES_STATUS, countLoadEvaluatedFormulas, documentValuePolicy, filterValueDomain, isValueUnavailable,
  laneValue, savedValuesStatus, scanValueDomain, valueAuthority,
} from '../src/xlsx/view/values'
import { xlsxBytes } from './helpers/xlsx-package'

vi.mock('../src/xlsx/formula/workbook', async (importOriginal) => {
  const original = await importOriginal<typeof import('../src/xlsx/formula/workbook')>()
  return { ...original, evaluateWorkbookFormulas: vi.fn(original.evaluateWorkbookFormulas) }
})
const evaluator = vi.mocked(workbook.evaluateWorkbookFormulas)
beforeEach(() => { evaluator.mockClear() })

const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const load = async (sheetData: string, policy?: 'engine' | 'cached', extra: { fullCalcOnLoad?: boolean } = {}) =>
  parseXlsx(await OfficePackage.load(xlsxBytes({ sheetData, ...extra })), policy ? { valuePolicy: policy } : {})

// Header row, then A = 3,1,2 and B = A*2 with a cache only on B2.
const TABLE = [
  '<row r="1"><c r="A1" t="inlineStr"><is><t>k</t></is></c><c r="B1" t="inlineStr"><is><t>f</t></is></c></row>',
  '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f>A2*2</f><v>6</v></c></row>',
  '<row r="3"><c r="A3"><v>1</v></c><c r="B3"><f>A3*2</f></c></row>',
  '<row r="4"><c r="A4"><v>2</v></c><c r="B4"><f>A4*2</f></c></row>',
].join('')
const range = { r0: 0, r1: 3, c0: 0, c1: 1, headerRows: 1 as const, totalsRows: 0 as const }
const sortOn = (col: number): SortSpec => ({ range, keys: [{ col, order: 'asc' }], caseSensitive: false })
const filterOn = (col: number): FilterSpec => ({ range, columns: new Map([[col, { kind: 'number', op: 'gt', first: 0 }]]) })
const quiet = { yield: async () => {} }

describe('value authority classification (one rule for both adapters)', () => {
  it('a formula without a saved value is unavailable only under cached', () => {
    const uncached = { value: null, formula: 'A1*2', hasCachedValue: false }
    expect(isValueUnavailable(uncached, 'cached')).toBe(true)
    expect(isValueUnavailable(uncached, 'engine')).toBe(false)
    expect(isValueUnavailable({ value: null, sharedFormula: { si: 0 }, hasCachedValue: false }, 'cached')).toBe(true)
    // A literal never becomes unavailable, cached or not.
    expect(isValueUnavailable({ value: null, hasCachedValue: false }, 'cached')).toBe(false)
    expect(isValueUnavailable({ value: 7, formula: 'A1', hasCachedValue: true }, 'cached')).toBe(false)
  })

  it('authority carries value, provenance, error identity and the unavailable state', () => {
    expect(valueAuthority({ value: '#VALUE!', valueIsError: true, hasCachedValue: true }, 'cached'))
      .toEqual({ value: '#VALUE!', provenance: 'cached', valueIsError: true, unavailable: false })
    expect(valueAuthority({ value: '#VALUE!', valueIsError: false, hasCachedValue: true }, 'engine'))
      .toEqual({ value: '#VALUE!', provenance: 'engine', valueIsError: false, unavailable: false })
    // Unavailable is never blank, zero or an error value: the value is withheld.
    expect(valueAuthority({ value: null, formula: 'X', hasCachedValue: false, valueIsError: true }, 'cached'))
      .toEqual({ value: null, provenance: 'cached', valueIsError: false, unavailable: true })
  })

  it('lane classification separates error tokens from error-looking text and keeps lone surrogates', () => {
    expect(laneValue(1.5, false, false)).toEqual({ tag: LANE_NUMBER, num: 1.5, text: '' })
    expect(laneValue(true, false, false)).toEqual({ tag: LANE_BOOLEAN, num: 1, text: '' })
    expect(laneValue(null, false, false)).toEqual({ tag: LANE_BLANK, num: 0, text: '' })
    expect(laneValue('#VALUE!', true, false)).toEqual({ tag: LANE_ERROR, num: 0, text: '#VALUE!' })
    expect(laneValue('#VALUE!', false, false)).toEqual({ tag: LANE_STRING, num: 0, text: '#VALUE!' })
    expect(laneValue(Number.POSITIVE_INFINITY, false, false)).toEqual({ tag: LANE_ERROR, num: 0, text: '#NUM!' })
    expect(laneValue(5, false, true)).toEqual({ tag: LANE_UNAVAILABLE, num: 0, text: '' })
    expect(laneValue('a\ud800', false, false).text.charCodeAt(1)).toBe(0xd800)
  })
})

describe('per-document value policy', () => {
  it('a cached parse records its policy on the document; an engine parse adds nothing', async () => {
    const cached = await load(TABLE, 'cached')
    const engine = await load(TABLE)
    expect(documentValuePolicy(cached)).toBe('cached')
    expect(documentValuePolicy(engine)).toBe('engine')
    expect(Object.keys(engine)).not.toContain('valuePolicy')
  })

  it('a cached document wrapped without an explicit policy stays cached (no silent engine wrap)', async () => {
    const doc = await load(TABLE, 'cached')
    const source = createObjectWorkbookSource(doc, 'doc')
    expect(source.valuePolicy).toBe('cached')
    const store = source.store('1')
    const read = await store.readCell({ row: 2, col: 1 }, opts(store))
    expect(read).toMatchObject({ kind: 'unavailable', reason: FORMULA_VALUE_UNAVAILABLE, cell: { valueProvenance: 'cached', formula: 'A3*2' } })
  })

  it('a document is never wrapped under the other policy', async () => {
    const cached = await load(TABLE, 'cached'), engine = await load(TABLE)
    expect(() => createObjectWorkbookSource(cached, 'doc', { valuePolicy: 'engine' })).toThrow(RangeError)
    expect(() => createObjectWorkbookSource(engine, 'doc', { valuePolicy: 'cached' })).toThrow(RangeError)
  })

  it('the evaluator refuses a cached document', async () => {
    const doc = await load(TABLE, 'cached')
    expect(() => workbook.evaluateWorkbookFormulas(doc)).toThrow(/saved values/)
    expect(doc.sheets[0].rows[2].cells[1].value).toBeNull()
  })
})

describe('cached documents never invoke the evaluator', () => {
  it('parse, source queries, sort and filter run no evaluation; engine parse evaluates exactly once', async () => {
    const cached = await load(TABLE, 'cached')
    expect(evaluator).not.toHaveBeenCalled()
    const source = createObjectWorkbookSource(cached, 'doc')
    const store = source.store('1')
    await store.readRows({ r0: 0, r1: 10, c0: 0, c1: 5 }, undefined, opts(store))
    await store.queryColumn({ col: 1, rows: { kind: 'intervals', intervals: [{ r0: 0, r1: 10 }] }, limit: 50, encoding: 'lanes' }, opts(store))
    expect(await sortRowOrder(store, sortOn(0), quiet)).toMatchObject({ kind: 'complete' })
    expect(evaluator).not.toHaveBeenCalled()

    await load(TABLE)
    expect(evaluator).toHaveBeenCalledTimes(1)
  })

  it('view operations never recalculate an engine document either', async () => {
    const doc = await load(TABLE)
    evaluator.mockClear()
    const before = JSON.stringify(doc)
    const filterEvaluator: FilterEvaluator = async () => ({ kind: 'complete', matches: () => true })
    const session = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'), { ...quiet, filterEvaluator })
    const sorted = await session.sort(initialIdentity('doc', '1'), sortOn(1), new AbortController().signal)
    expect(sorted.kind).toBe('committed')
    expect(evaluator).not.toHaveBeenCalled()
    expect(JSON.stringify(doc)).toBe(before)
  })
})

describe('formula-value-unavailable refusals are limited to the key/value domain', () => {
  it('sort: an uncached formula in the key column refuses with count and first address', async () => {
    const store = createObjectWorkbookSource(await load(TABLE, 'cached'), 'doc').store('1')
    expect(await sortRowOrder(store, sortOn(1), quiet)).toEqual({
      kind: 'refused', refusal: { reason: FORMULA_VALUE_UNAVAILABLE, count: 2, firstAddress: { row: 2, col: 1 } },
    })
    // The same cells outside the key domain do not block.
    expect(await sortRowOrder(store, sortOn(0), quiet)).toMatchObject({ kind: 'complete' })
  })

  it('filter: a rule column holding uncached formulas refuses before any evaluation; another rule column does not', async () => {
    const doc = await load(TABLE, 'cached')
    let evaluated = 0
    const filterEvaluator: FilterEvaluator = async () => { evaluated++; return { kind: 'complete', matches: () => true } }
    const session = new DocumentSession('doc', createObjectWorkbookSource(doc, 'doc'), { ...quiet, filterEvaluator })
    const owner = initialIdentity('doc', '1')
    expect(await session.filter(owner, filterOn(1), 0, new AbortController().signal)).toEqual({
      kind: 'refused', refusal: { reason: FORMULA_VALUE_UNAVAILABLE, count: 2, firstAddress: { row: 2, col: 1 } },
    })
    expect(evaluated).toBe(0)
    const ok = await session.filter(owner, filterOn(0), 0, new AbortController().signal)
    expect(ok.kind).toBe('committed')
    expect(evaluated).toBe(1)
  })

  it('the filter precheck runs only for cached documents (engine values are never unavailable)', async () => {
    const reads: Record<string, number> = {}
    for (const policy of ['engine', 'cached'] as const) {
      const inner = createObjectWorkbookSource(await load(TABLE, policy), 'doc')
      const store = inner.store('1')
      reads[policy] = 0
      const counted: SourceStore = Object.assign(Object.create(Object.getPrototypeOf(store)), store, {
        queryColumn: (...args: Parameters<SourceStore['queryColumn']>) => { reads[policy]++; return store.queryColumn(...args) },
      })
      const source = { ...inner, store: () => counted }
      const filterEvaluator: FilterEvaluator = async () => ({ kind: 'complete', matches: () => true })
      const session = new DocumentSession('doc', source, { ...quiet, filterEvaluator })
      await session.filter(initialIdentity('doc', '1'), filterOn(0), 0, new AbortController().signal)
    }
    expect(reads.engine).toBe(0)
    expect(reads.cached).toBeGreaterThan(0)
  })

  it('the domain is the body band of the rule/key columns: header and totals rows are outside it', async () => {
    expect(filterValueDomain(filterOn(1))).toEqual({ cols: [1], r0: 1, r1: 3 })
    expect(filterValueDomain({ range: { ...range, totalsRows: 1 }, columns: new Map() })).toBeNull()
    const doc = await load('<row r="1"><c r="A1"><f>Z1</f></c></row><row r="2"><c r="A2"><v>1</v></c></row>', 'cached')
    const store = createObjectWorkbookSource(doc, 'doc').store('1')
    let pauses = 0
    expect(await scanValueDomain(store, { cols: [0], r0: 1, r1: 1 }, opts(store), async () => { pauses++ })).toEqual({ kind: 'clear' })
    expect(pauses).toBeGreaterThan(0)
    expect(await scanValueDomain(store, { cols: [0], r0: 0, r1: 1 }, opts(store))).toEqual({
      kind: 'refused', refusal: { reason: FORMULA_VALUE_UNAVAILABLE, count: 1, firstAddress: { row: 0, col: 0 } },
    })
  })

  it('under engine the same workbook publishes computed values and nothing refuses (an enumerated policy difference)', async () => {
    const store = createObjectWorkbookSource(await load(TABLE), 'doc').store('1')
    const read = await store.readCell({ row: 2, col: 1 }, opts(store))
    expect(read).toMatchObject({ kind: 'cell', cell: { value: 2, valueProvenance: 'engine' } })
    expect(await sortRowOrder(store, sortOn(1), quiet)).toMatchObject({ kind: 'complete' })
  })
})

describe('engine documents are unchanged', () => {
  it('explicit engine policy is byte-identical to the default parse and the adapter reads the published cell.value', async () => {
    const sheet = TABLE + '<row r="5"><c r="A5" t="e"><v>#DIV/0!</v></c><c r="B5" t="str"><v>#VALUE!</v></c><c r="C5" t="b"><v>0</v></c><c r="D5" s="0"/></row>'
    // Engine semantics sample the clock lazily; keep it pinned through serialization.
    const clock = vi.spyOn(Date, 'now').mockReturnValue(1_700_000_000_000)
    let plain, explicit
    try {
      plain = await load(sheet)
      explicit = await load(sheet, 'engine')
      expect(JSON.stringify(explicit)).toBe(JSON.stringify(plain))
    } finally { clock.mockRestore() }
    const store = createObjectWorkbookSource(plain, 'doc').store('1')
    for (const row of plain.sheets[0].rows) for (const cell of row.cells) {
      const read = await store.readCell({ row: cell.row, col: cell.col }, opts(store))
      if (cell.value === null && cell.formula === undefined) continue
      expect(read.kind).toBe('cell')
      if (read.kind !== 'cell') continue
      expect(read.cell.value).toBe(cell.value)
      expect(read.cell.valueIsError).toBe(cell.valueIsError)
      expect(read.cell.valueProvenance).toBe('engine')
    }
    expect(await store.readCell({ row: 4, col: 0 }, opts(store))).toMatchObject({ kind: 'cell', cell: { value: '#DIV/0!', valueIsError: true } })
    expect(await store.readCell({ row: 4, col: 1 }, opts(store))).toMatchObject({ kind: 'cell', cell: { value: '#VALUE!', valueIsError: false } })
  })
})

describe('saved-values status', () => {
  it('appears only for cached documents, with the approved count wording', async () => {
    expect(savedValuesStatus('engine', 4)).toBeUndefined()
    expect(savedValuesStatus('engine')).toBeUndefined()
    expect(savedValuesStatus('cached', 3)).toBe('Showing saved values — 3 formulas not recalculated (file too large)')
    expect(savedValuesStatus('cached')).toBe(SAVED_VALUES_STATUS)
    expect(SAVED_VALUES_STATUS).toBe('Showing saved values; formulas are not recalculated')
  })

  it('object sources report it on open only under cached, counting the formulas the engine would have evaluated', async () => {
    const cached = await load(TABLE, 'cached')
    // B2 has a saved value; B3/B4 do not: only those two are lost.
    expect(countLoadEvaluatedFormulas(cached)).toBe(2)
    const source = createObjectWorkbookSource(cached, 'doc')
    expect(source.openDiagnostics).toContainEqual({ code: 'cached-value-policy', message: 'Showing saved values — 2 formulas not recalculated (file too large)' })
    expect(source.valueStatus?.()).toEqual({ policy: 'cached', formulas: 2, message: 'Showing saved values — 2 formulas not recalculated (file too large)' })
    const engine = createObjectWorkbookSource(await load(TABLE), 'doc')
    expect(engine.openDiagnostics ?? []).toEqual([])
    expect(engine.valueStatus?.()).toMatchObject({ policy: 'engine', message: undefined })
  })
})
