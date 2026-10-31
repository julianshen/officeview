// @vitest-environment node
/**
 * Task 22: pivot calculated fields (spec §5.6). A calculated field's formula combines the SUMS of other fields,
 * evaluated per cell after aggregation, subtotal and grand-total cells included. Expected values here are computed
 * by hand: sum the source records, then apply the formula. Behaviour with no Excel-written oracle fails closed
 * (ruling S2-M): division by zero, error or text results, show-as, references to other calculated fields, formulas
 * over no field, text-valued source fields and empty intersections.
 */
import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { tokenize } from '../src/xlsx/formula/lexer'
import { compileCalculatedField, calculatedValue } from '../src/xlsx/pivots/engine/calculated'
import { computeAxes, lineCarriesValues } from '../src/xlsx/pivots/engine/axes'
import { savedPivotLayout, type PivotLine } from '../src/xlsx/pivots/savedLayout'
import { computeValues } from '../src/xlsx/pivots/engine/values'
import { checkSavedLayout, sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { initialViewState, type DataFieldState, type FieldState, type PivotViewState } from '../src/xlsx/pivots/engine/view'
import { VALUE_TAG, type RecordStore } from '../src/xlsx/pivots/engine/records'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import { loadAll } from './helpers/xlsx-pivots-load'

// ---------------------------------------------------------------- lexer mode (AGY-8)

describe('lexer pivot-formula mode', () => {
  it('quoted names with spaces and quotes lex as names only in the opt-in mode', () => {
    const names = tokenize("'Unit Price'*'Q''ty'", { quotedNames: true })
    expect(names.map(token => [token.type, token.raw ?? token.value])).toEqual([
      ['ident', 'Unit Price'], ['op', '*'], ['ident', "Q'ty"], ['eof', ''],
    ])
    // A quoted run before `!` is still a sheet qualifier in the mode.
    const sheet = tokenize("'Sheet 1'!A1+'Field A'", { quotedNames: true })
    expect(sheet.map(token => token.type)).toEqual(['cell', 'op', 'ident', 'eof'])
    expect(sheet[0]!.sheet).toBe('Sheet 1')
    // An unterminated quote stays an error token.
    expect(tokenize("'Field A", { quotedNames: true })[0]!.type).toBe('error')
  })

  it('lexer default mode unchanged: a quoted run without `!` is an error token', () => {
    for (const input of ["'Field A'*20", "SUM('Field A', 'Field B')", " 'Field4'/1"]) {
      expect(tokenize(input)).toEqual(tokenize(input, {}))
      expect(tokenize(input).some(token => token.type === 'error' && token.value.startsWith("'"))).toBe(true)
    }
    expect(tokenize("'Sheet 1'!A1")[0]).toMatchObject({ type: 'cell', sheet: 'Sheet 1' })
  })
})

// ---------------------------------------------------------------- hand-built pivot

const NAMES = ['Region', 'Unit Price', "Q'ty", 'Ratio', 'Twice', 'Note'] as const
const REGIONS = ['East', 'West', 'North']
// Records: [region, unit price, qty]. Sums: East 30/5, West 5/5, North 10/2, total 45/12.
const RECORDS: Array<[number, number, number]> = [[0, 10, 2], [0, 20, 3], [1, 5, 5], [2, 7, 1], [2, 3, 1]]

const fieldState = (count: number): FieldState => ({
  hidden: Array(count).fill(false), collapsed: Array(count).fill(false), sortType: 'manual',
  defaultSubtotal: true, subtotals: [], subtotalTop: true, compact: true, outline: true,
  insertBlankRow: false, fillDownLabels: false, showAll: false,
})
const sumOf = (field: number): DataFieldState => ({ field, fn: 'sum', showAs: 'normal', baseField: -1, baseItem: 1048832, name: '' })

interface Fixture { store: RecordStore; cache: PivotCacheDefinition; def: PivotTableDefinition; view: PivotViewState }
function fixture(formulas: Partial<Record<3 | 4, string>> = {}, records = RECORDS, note: Array<string | number> = records.map(() => 1)): Fixture {
  const inline = (values: Array<number | string>) => {
    const text = [...new Set(values.filter((value): value is string => typeof value === 'string'))]
    return {
      kind: 'inline' as const,
      tag: Uint8Array.from(values, value => typeof value === 'number' ? VALUE_TAG.number : VALUE_TAG.string),
      num: Float64Array.from(values, value => typeof value === 'number' ? value : 0),
      text, textIndex: Uint32Array.from(values, value => typeof value === 'string' ? text.indexOf(value) : 0),
    }
  }
  const cache = {
    source: { type: 'worksheet' }, olap: false,
    fields: [
      { name: NAMES[0], databaseField: true, sharedItems: { inline: false, items: REGIONS.map(value => ({ type: 'string', value })) } },
      { name: NAMES[1], databaseField: true, sharedItems: { inline: true, items: [] } },
      { name: NAMES[2], databaseField: true, sharedItems: { inline: true, items: [] } },
      { name: NAMES[3], databaseField: false, formula: formulas[3] ?? "'Unit Price'/'Q''ty'" },
      { name: NAMES[4], databaseField: false, formula: formulas[4] ?? "'Unit Price'*2" },
      { name: NAMES[5], databaseField: true, sharedItems: { inline: true, items: [] } },
    ],
  } as unknown as PivotCacheDefinition
  const def = {
    fields: [{ items: REGIONS.map((_, x) => ({ type: 'data', x })) }, {}, {}, {}, {}, {}],
    rowGrandTotals: true, colGrandTotals: true, mergeItem: false, gridDropZones: false, outlineData: false,
    unsupported: [],
  } as unknown as PivotTableDefinition
  const view: PivotViewState = {
    rows: [0], cols: [-2], pages: [], filters: [],
    fields: [fieldState(3), fieldState(0), fieldState(0), fieldState(0), fieldState(0), fieldState(0)],
    data: [sumOf(1), sumOf(3)],
  }
  const store: RecordStore = {
    count: records.length, bytes: 0,
    columns: [
      { kind: 'shared', index: Uint32Array.from(records, r => r[0]) },
      inline(records.map(r => r[1])), inline(records.map(r => r[2])), undefined, undefined, inline(note),
    ],
  }
  return { store, cache, def, view }
}

/** Every value cell as `rowLine/colLine: value`, with the data slot read from the Values level. */
function run(f: Fixture, view: PivotViewState = f.view) {
  const axes = computeAxes(f.store, f.cache, f.def, view)
  if (!axes.ok) return axes
  const values = computeValues(f.store, f.cache, f.def, view, axes.value)
  if (!values.ok) return values
  const cells: Record<string, unknown> = {}
  const onRows = view.rows.includes(-2), onCols = view.cols.includes(-2)
  // A grand line on the Values axis is one per data field: its label carries the slot.
  const label = (line: { type: string; items: readonly number[]; data: number }, values: boolean) =>
    `${line.type}:${line.items.join('.')}${line.type === 'grand' && values ? `#${line.data}` : ''}`
  axes.value.rows.forEach((row, r) => axes.value.cols.forEach((col, c) => {
    const d = onRows ? row.data : onCols ? col.data : 0
    const value = values.value.get(r, c, d)
    if (value.kind !== 'empty') cells[`${label(row, onRows)}/${label(col, onCols)}`] = value.kind === 'number' ? value.value : value
  }))
  return { ok: true as const, cells }
}
const diagnostic = (result: ReturnType<typeof run>) => result.ok ? undefined : { kind: result.diagnostic.kind, feature: result.diagnostic.feature }
const UNMEASURED = { kind: 'unsupported', feature: 'pivot-engine-unmeasured' }
const CALCULATED = { kind: 'unsupported', feature: 'pivot-calculated-field' }

/** Rows Region, then Note (one shared item, 1, on every record): each region gets a default subtotal line. */
function nestedFixture(): Fixture {
  const f = fixture()
  return {
    cache: { ...f.cache, fields: f.cache.fields.map((field, k) => k === 5 ? { ...field, sharedItems: { inline: false, items: [{ type: 'number', value: 1 }] } } : field) } as PivotCacheDefinition,
    store: { ...f.store, columns: f.store.columns.map((column, k) => k === 5 ? { kind: 'shared' as const, index: new Uint32Array(f.store.count) } : column) },
    def: { ...f.def, fields: f.def.fields.map((field, k) => k === 5 ? { items: [{ type: 'data', x: 0 }] } : field) } as PivotTableDefinition,
    view: { ...f.view, rows: [0, 5], cols: [-2], fields: f.view.fields.map((field, k) => k === 5 ? fieldState(1) : field) },
  }
}

describe('calculatedValue', () => {
  const { cache } = fixture()
  const compile = (formula: string) => {
    const compiled = compileCalculatedField(formula, cache, 3)
    if (!compiled.ok) throw new Error(compiled.diagnostic.message)
    return compiled.value
  }

  it('binds quoted names with spaces and quotes to their field sums', () => {
    const sums = new Map([[1, 45], [2, 12]])
    expect(calculatedValue(compile("'Unit Price'/'Q''ty'"), field => sums.get(field)!)).toEqual({ kind: 'number', value: 45 / 12 })
    expect(calculatedValue(compile(" SUM('Unit Price', 'Q''ty')*2"), field => sums.get(field)!)).toEqual({ kind: 'number', value: 114 })
    // Unquoted, case-insensitive field name.
    expect(calculatedValue(compile('region*0+REGION*0+1'), () => 0)).toEqual({ kind: 'number', value: 1 })
  })

  it('evaluates each cell from its own sums', () => {
    const ratio = compile("'Unit Price'/'Q''ty'")
    expect(calculatedValue(ratio, field => field === 1 ? 30 : 5)).toEqual({ kind: 'number', value: 6 })
    expect(calculatedValue(ratio, field => field === 1 ? 10 : 2)).toEqual({ kind: 'number', value: 5 })
  })

  it('reports error and text results as values (the pivot gate refuses them)', () => {
    expect(calculatedValue(compile("'Unit Price'/'Q''ty'"), field => field === 1 ? 1 : 0)).toEqual({ kind: 'error', value: '#DIV/0!' })
    expect(calculatedValue(compile('IF(\'Unit Price\'>1,"big","small")'), () => 5)).toEqual({ kind: 'text', value: 'big' })
  })

  it('refuses at compile time what the engine cannot evaluate', () => {
    const feature = (formula: string, self = 3) => {
      const compiled = compileCalculatedField(formula, cache, self)
      return compiled.ok ? 'ok' : `${compiled.diagnostic.kind}:${compiled.diagnostic.feature}`
    }
    expect(feature("SQRT('Unit Price')")).toBe('unsupported:pivot-calculated-field')
    expect(feature("'No Such Field'*2")).toBe('unsupported:pivot-calculated-field')
    expect(feature('A1*2')).toBe('unsupported:pivot-calculated-field')
    expect(feature("'Unit Price'*{1,2}")).toBe('unsupported:pivot-calculated-field')
    expect(feature("'Unit Price'*")).toBe('unsupported:pivot-calculated-field')
    expect(feature("RAND()*'Unit Price'")).toBe('unsupported:pivot-calculated-field')
    expect(feature("ROW()*'Unit Price'")).toBe('unsupported:pivot-calculated-field')
    expect(feature("YEAR('Unit Price')")).toBe('unsupported:pivot-calculated-field')
    // A reference to another calculated field re-applies its formula over the sums (oracle round B2,
    // calc-references-calculated); a cycle of references cannot be evaluated.
    expect(feature("'Ratio'/1", 4)).toBe('ok')
    expect(feature("'Twice'*1", 3)).toBe('ok')
    const cyclic = { ...cache, fields: cache.fields.map((field, k) => k === 4 ? { ...field, formula: "'Ratio'*2" } : field) } as PivotCacheDefinition
    const cycle = compileCalculatedField("'Twice'*1", cyclic, 3)
    expect(cycle.ok ? undefined : cycle.diagnostic).toEqual({ kind: 'unsupported', feature: 'pivot-calculated-field', message: 'calculated field 3: the formula refers to its own field Ratio' })
    // A cycle that does not pass through the field itself.
    const loop = { ...cyclic, fields: cyclic.fields.map((field, k) => k === 3 ? { ...field, formula: "'Twice'*1" } : field) } as PivotCacheDefinition
    const other = compileCalculatedField("'Ratio'+1", { ...loop, fields: [...loop.fields, { name: 'Third', databaseField: false, formula: "'Ratio'+1" }] } as PivotCacheDefinition, 6)
    expect(other.ok ? undefined : other.diagnostic.message).toBe('calculated field 6: calculated field Ratio refers back to itself')
    // A formula over no field (`2/0`): whether every cell shows it is not measured.
    expect(feature('2/0')).toBe('unsupported:pivot-engine-unmeasured')
  })
})

describe('calculated fields in the pivot', () => {
  it('is the formula over each cell\'s sums, subtotal and grand-total cells included', () => {
    expect(run(fixture())).toEqual({
      ok: true, cells: {
        // Sum of Unit Price, then 'Unit Price'/'Q''ty' over the summed values per cell.
        'data:0/data:0': 30, 'data:0/data:1': 30 / 5,
        'data:1/data:0': 5, 'data:1/data:1': 5 / 5,
        'data:2/data:0': 10, 'data:2/data:1': 10 / 2,
        // Grand total: 45 / 12, not the sum (12) or mean of the per-region ratios.
        'grand:/data:0': 45, 'grand:/data:1': 45 / 12,
      },
    })
  })

  it('applies at item subtotals of a nested row field and with the Values field on rows', () => {
    const f = fixture()
    const result = run(nestedFixture())
    expect(result.ok && result.cells).toEqual({
      'data:0/data:0': 30, 'data:0/data:1': 6, 'data:0.0/data:0': 30, 'data:0.0/data:1': 6,
      'data:1/data:0': 5, 'data:1/data:1': 1, 'data:1.0/data:0': 5, 'data:1.0/data:1': 1,
      'data:2/data:0': 10, 'data:2/data:1': 5, 'data:2.0/data:0': 10, 'data:2.0/data:1': 5,
      'grand:/data:0': 45, 'grand:/data:1': 3.75,
    })
    const onRows = run(f, { ...f.view, rows: [0, -2], cols: [] })
    expect(onRows.ok && onRows.cells).toEqual({
      'data:0.0/data:': 30, 'data:0.1/data:': 6, 'data:1.0/data:': 5, 'data:1.1/data:': 1,
      'data:2.0/data:': 10, 'data:2.1/data:': 5, 'grand:#0/data:': 45, 'grand:#1/data:': 3.75,
    })
  })

  it('division by zero gives #DIV/0! per measurement: unmeasured, so the pivot fails closed', () => {
    // West's Q'ty sums to 0. Excel's display of an error in a calculated cell (#DIV/0! or errorCaption) has no
    // Excel-written oracle yet (S2-M).
    const records: Array<[number, number, number]> = [[0, 10, 2], [1, 5, 0], [2, 3, 1]]
    expect(diagnostic(run(fixture({}, records)))).toEqual(UNMEASURED)
  })

  it('an unsupported function makes the pivot unsupported', () => {
    expect(diagnostic(run(fixture({ 3: "SQRT('Unit Price')" })))).toEqual(CALCULATED)
    expect(diagnostic(run(fixture({ 3: "'Unknown'*2" })))).toEqual(CALCULATED)
  })

  it('refuses a non-finite value read on the calculated-field aggregation path', () => {
    const f = fixture({}, RECORDS, RECORDS.map((_, index) => index === 0 ? Number.NaN : 1))
    f.view.data = [
      { ...sumOf(5), fn: 'average' },
      sumOf(4),
    ]
    expect(diagnostic(run(f))).toEqual(UNMEASURED)
  })

  it('fails closed on every unmeasured calculated-field behaviour', () => {
    const f = fixture()
    // Show-as and a function other than sum on the calculated field.
    expect(diagnostic(run(f, { ...f.view, data: [sumOf(1), { ...sumOf(3), showAs: 'percentOfTotal' }] }))).toEqual(UNMEASURED)
    expect(diagnostic(run(f, { ...f.view, data: [sumOf(1), { ...sumOf(3), fn: 'count' }] }))).toEqual(UNMEASURED)
    // A formula over no field (its every-cell rule is measured, calc-constant, but values.ts needs a field to sum).
    expect(diagnostic(run(fixture({ 3: '2/0' })))).toEqual(UNMEASURED)
    // A text result.
    expect(diagnostic(run(fixture({ 3: 'IF(\'Unit Price\'>1,"big","small")' })))).toEqual(UNMEASURED)
    // A text-valued source field ('Name'*2): Excel's sum-of-text rule for calculated fields is not measured.
    expect(diagnostic(run(fixture({ 3: "'Note'*2" }, RECORDS, ['a', 1, 1, 1, 1])))).toEqual(UNMEASURED)
    // A named subtotal line: which function a calculated cell takes there is not measured.
    const n = nestedFixture()
    const named = { ...n.view, fields: n.view.fields.map((field, k) => k === 0 ? { ...field, subtotals: ['avg' as const] } : field) }
    expect(diagnostic(run(n, named))).toEqual(UNMEASURED)
    // An empty intersection: with Note on columns and East's records moved to item 1, North x item 0 has no
    // records. Excel fills calculated cells without records; not measured.
    const sparse = nestedFixture()
    const twoNotes = {
      ...sparse,
      cache: { ...sparse.cache, fields: sparse.cache.fields.map((field, k) => k === 5 ? { ...field, sharedItems: { inline: false, items: [{ type: 'number', value: 1 }, { type: 'number', value: 2 }] } } : field) } as PivotCacheDefinition,
      store: { ...sparse.store, columns: sparse.store.columns.map((column, k) => k === 5 ? { kind: 'shared' as const, index: Uint32Array.from(RECORDS, r => r[0] === 2 ? 1 : 0) } : column) },
      def: { ...sparse.def, fields: sparse.def.fields.map((field, k) => k === 5 ? { items: [{ type: 'data', x: 0 }, { type: 'data', x: 1 }] } : field) } as PivotTableDefinition,
    }
    const columns = { ...twoNotes.view, rows: [0], cols: [5, -2], fields: twoNotes.view.fields.map((field, k) => k === 5 ? fieldState(2) : field) }
    expect(run(twoNotes, { ...columns, data: [sumOf(1)] }).ok).toBe(true)
    expect(diagnostic(run(twoNotes, columns))).toEqual(UNMEASURED)
  })

  it('a calculated field on an axis or page is unsupported; calculated items keep their feature', () => {
    const f = fixture()
    expect(diagnostic(run(f, { ...f.view, rows: [0, 3] }))).toEqual(CALCULATED)
    expect(diagnostic(run(f, { ...f.view, pages: [{ field: 3, selected: 'all' }] }))).toEqual(CALCULATED)
    const items = { ...f.def, fields: f.def.fields.map((field, k) => k === 0 ? { items: [...REGIONS.map((_, x) => ({ type: 'data', x })), { type: 'data', x: 0, calculated: true }] } : field) } as unknown as PivotTableDefinition
    expect(diagnostic(run({ ...f, def: items }))).toEqual({ kind: 'unsupported', feature: 'pivot-calculated-item' })
  })

  it('a value filter on a calculated data field fails closed', () => {
    const f = fixture()
    const filter = { fld: 0, type: 'valueGreaterThan', evalOrder: 0, id: 1, iMeasureFld: 1, doubleValue1: 2, autoFilter: { columns: [] } }
    expect(diagnostic(run(f, { ...f.view, filters: [filter as never] }))).toEqual(UNMEASURED)
  })
})

// ---------------------------------------------------------------- oracles

const CORPUS_FILE = resolve(process.cwd(), 'corpus/lo-calc-calcfields.xlsx')
describe.skipIf(!existsSync(CORPUS_FILE))('corpus calcfields.xlsx (Collabora-written: recorded, not an oracle)', () => {
  it('is written by Collabora, and its pivot fails closed on the unmeasured fields', async () => {
    const zip = await JSZip.loadAsync(readFileSync(CORPUS_FILE))
    expect(await zip.file('docProps/app.xml')!.async('string')).toMatch(/<Application>Collabora_Office/)
    const [p] = await loadAll(CORPUS_FILE)
    // Field3 'Name'*2 (a text field), Field4 2/0 (no field), Field5 'Field4'/1 (a calculated field).
    expect(checkSavedLayout(p!.store, p!.cache, p!.def, sheetCellReader(p!.sheet), { date1904: p!.date1904 }))
      .toMatchObject({ status: 'unverified', reason: 'unsupported', detail: 'pivot-engine-unmeasured' })
  })

  it('Field1 and Field2 are the formula over hand-summed Field A and Field B', async () => {
    const [p] = await loadAll(CORPUS_FILE)
    const view = initialViewState(p!.def, p!.cache)
    const data = view.data.slice(0, 4)   // Field A, Field B, Field1 = 'Field A'*20, Field2 = SUM('Field A', 'Field B')
    const result = run({ store: p!.store, cache: p!.cache, def: p!.def, view: { ...view, data } })
    // Hand sums from the records part, per Name.
    const names = p!.store.columns[0]!, a = p!.store.columns[1]!, b = p!.store.columns[2]!
    if (names.kind !== 'shared' || a.kind !== 'shared' || b.kind !== 'shared') throw new Error('shared columns expected')
    const value = (field: number, index: number) => (p!.cache.fields[field]!.sharedItems!.items[index] as { value: number }).value
    const sums = [0, 1, 2, 'grand'].map(name => {
      let sa = 0, sb = 0
      for (let r = 0; r < p!.store.count; r++) if (name === 'grand' || names.index[r] === name) { sa += value(1, a.index[r]!); sb += value(2, b.index[r]!) }
      return [sa, sb, sa * 20, sa + sb]
    })
    expect(sums.map(s => s.slice(0, 2))).toEqual([[24, 6], [32, 10], [58, 2], [114, 18]])
    const expected: Record<string, number> = {}
    sums.forEach((s, k) => s.forEach((v, d) => { expected[`${k === 3 ? 'grand:' : `data:${k}`}/data:${d}`] = v }))
    expect(result.ok && result.cells).toEqual(expected)
  })
})

/**
 * Ruling S2-AA: task-22 manifest cases the engine still refuses, each with its exact refusal and why. `lane A` marks
 * a rule this round measured (see the oracle round B2 tests above) whose gate lives in values.ts, lane A's file;
 * the report gives the exact change. `kept` marks a gate Excel's file does not settle.
 */
const TASK22_REFUSED: Record<string, { refusal: string; why: string }> = {
  'calc-cell-without-records': { refusal: 'pivot-engine-unmeasured: a calculated cell without records is not measured', why: 'lane A: measured, the formula over zero sums (R22C4 = 0); values.ts gate' },
  'calc-quoted-field-name': { refusal: 'pivot-engine-unmeasured: a calculated cell without records is not measured', why: 'lane A: measured, 0/0 shows #DIV/0! (R22C4); values.ts gate' },
  'calc-division-by-zero': { refusal: 'pivot-engine-unmeasured: calculated data field 2 gives error #DIV/0!; its display is not measured', why: 'lane A: measured, the error value shows (R20C4) with showError off; values.ts gate' },
  'calc-references-calculated': { refusal: 'pivot-engine-unmeasured: calculated data field 2 gives error #DIV/0!; its display is not measured', why: 'reference lifted in calculated.ts; lane A: Central shows #DIV/0! (R20C4); values.ts gate' },
  'calc-constant': { refusal: 'pivot-engine-unmeasured: calculated field 12 uses no field; its cells without records are not measured', why: 'measured (1 in every cell, R22C4 included), kept in calculated.ts until values.ts can sum no field (lane A)' },
  'calc-text-field': { refusal: 'pivot-engine-unmeasured: field 6 has text values; their sum in a calculated field is not measured', why: 'lane A: measured, text is left out of the sum (R13C13 = 84); values.ts gate' },
  'calc-function-count': { refusal: 'pivot-engine-unmeasured: calculated data field 2 with function count is not measured', why: 'lane A: measured, the function is ignored (R10C7 = Sales*2); values.ts gate' },
  'calc-show-as': { refusal: 'pivot-engine-unmeasured: show-as on calculated data field 2 is not measured', why: 'kept: percent of the calculated grand total, but 10 of 76 cells differ at the 15th digit (the percent-of-total rounding lane A also refuses)' },
  'calc-named-subtotals': { refusal: 'pivot-engine-unmeasured: a named subtotal of a calculated data field is not measured', why: 'lane A: measured, blank on named lines except at the column grand (R10C13 = R11C13 = 96237.92); values.ts gate' },
  'calc-value-sort': { refusal: 'pivot-engine-unmeasured: a filter or sort by calculated data field 2 is not measured', why: 'lane A: measured, the order is the formula over each item\'s sums (West, South, North, East, Central); values.ts aggregateByKey gate' },
  'calc-item': { refusal: "pivot-calculated-item: field 1 has a calculated item; Excel's saved layout lists it under every parent and includes it in totals, which this engine does not reproduce", why: 'kept: the item is listed under every parent and counted in its totals (calc-item R6, R28); needs axes and values work outside this lane' },
}

describe('Task 22 Excel manifest cases', () => {
  const fixtures = resolve(process.cwd(), 'tests/fixtures/xlsx-pivots')
  const manifestPath = resolve(fixtures, 'engine/manifest.json')
  if (!existsSync(manifestPath)) {
    it.skip('SKIPPED: tests/fixtures/xlsx-pivots/engine/manifest.json is absent; Task 22 Excel oracle cases not run', () => {})
    return
  }
  type Case = { task: number; status?: string; case: string; file: string; pivot: string }
  const raw = JSON.parse(readFileSync(manifestPath, 'utf8')) as { cases?: Case[] } | Case[]
  const cases = (Array.isArray(raw) ? raw : raw.cases ?? []).filter(entry => entry.task === 22 && entry.status !== 'pending')
  if (!cases.length) it.skip('SKIPPED: Task 22 has no runnable Excel manifest cases', () => {})
  else it.each(cases.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
    const local = resolve(fixtures, 'engine', entry.file)
    const loaded = await loadAll(existsSync(local) ? local : resolve(process.cwd(), entry.file))
    const pivot = loaded.find(candidate => candidate.def.name === entry.pivot)
    if (!pivot) throw new Error(`${entry.case}: pivot ${entry.pivot} not found`)
    const result = checkSavedLayout(pivot.store, pivot.cache, pivot.def, sheetCellReader(pivot.sheet), { date1904: pivot.date1904 })
    const refused = TASK22_REFUSED[entry.case]
    if (!refused) { expect(result, entry.case).toEqual({ status: 'ok' }); return }
    const view = initialViewState(pivot.def, pivot.cache)
    const axes = computeAxes(pivot.store, pivot.cache, pivot.def, view)
    const computed = axes.ok ? computeValues(pivot.store, pivot.cache, pivot.def, view, axes.value) : axes
    expect(computed.ok ? undefined : `${computed.diagnostic.feature}: ${computed.diagnostic.message}`, entry.case).toBe(refused.refusal)
    expect(result, entry.case).toEqual({ status: 'unverified', reason: 'unsupported', detail: refused.refusal.split(':')[0] })
  })
  it('the refusal map names task-22 cases only', () => {
    for (const name of Object.keys(TASK22_REFUSED)) expect(cases.map(entry => entry.case)).toContain(name)
  })
})

// ---------------------------------------------------------------- oracle round B2 (Task 15 Excel cases)

describe('oracle round B2: calculated fields and items against Excel', () => {
  const engine = async (name: string) => (await loadAll(resolve(process.cwd(), `tests/fixtures/xlsx-pivots/engine/${name}.xlsx`)))[0]!
  type Pivot = Awaited<ReturnType<typeof engine>>
  /** Each value cell of data slot `d`: Excel's saved value and the slot's formula over that cell's field sums (each
   * named field summed through computeValues, a cell without records summing to 0). */
  function formulaCells(p: Pivot, d: number) {
    const view = initialViewState(p.def, p.cache)
    const axes = computeAxes(p.store, p.cache, p.def, view)
    if (!axes.ok) throw new Error(axes.diagnostic.message)
    const compiled = compileCalculatedField(p.cache.fields[view.data[d]!.field]!.formula!, p.cache, view.data[d]!.field)
    if (!compiled.ok) throw new Error(compiled.diagnostic.message)
    const grids = new Map(compiled.value.fields.map(field => {
      const grid = computeValues(p.store, p.cache, p.def, { ...view, data: view.data.map(data => ({ ...data, field, fn: 'sum' as const, showAs: 'normal' as const })) }, axes.value)
      if (!grid.ok) throw new Error(grid.diagnostic.message)
      return [field, grid.value] as const
    }))
    const saved = savedPivotLayout(p.def)!, read = sheetCellReader(p.sheet)
    const onCols = view.cols.includes(-2)
    const cells: Array<{ row: number; col: number; rowLine: PivotLine; colLine: PivotLine; formula: unknown; saved: unknown }> = []
    axes.value.rows.forEach((rowLine, n) => axes.value.cols.forEach((colLine, m) => {
      if ((onCols ? colLine.data : 0) !== d) return
      if (!lineCarriesValues(view, 'rows', rowLine) || !lineCarriesValues(view, 'cols', colLine)) return
      const formula = calculatedValue(compiled.value, field => { const sum = grids.get(field)!.get(n, m, d); return sum.kind === 'number' ? sum.value : 0 })
      cells.push({ row: saved.dataRow0 + n, col: saved.dataCol0 + m, rowLine, colLine, formula, saved: read(saved.dataRow0 + n, saved.dataCol0 + m) })
    }))
    return cells
  }
  /** Saved and formula agree: numbers at 15 significant digits, errors by value. */
  const agrees = (cell: { formula: unknown; saved: unknown }) => {
    const f = cell.formula as { kind: string; value: number | string }, s = cell.saved as { kind: string; value: number | string } | undefined
    if (!s || f.kind !== s.kind) return false
    return f.kind === 'number' ? Number((f.value as number).toPrecision(15)) === Number((s.value as number).toPrecision(15)) : f.value === s.value
  }
  const disagreements = (cells: ReturnType<typeof formulaCells>) => cells.filter(cell => !agrees(cell)).map(cell => `R${cell.row + 1}C${cell.col + 1}`)

  it('every cell is the formula over its sums: cells without records sum to 0, errors show as the error', async () => {
    // calc-cell-without-records: Central x 2022 has no records; Sales and Units are blank, Double Sales shows 0 (R22C4).
    const empty = formulaCells(await engine('calc-cell-without-records'), 2)
    expect(disagreements(empty)).toEqual([])
    expect(empty.find(cell => cell.row === 21 && cell.col === 3)).toMatchObject({ formula: { kind: 'number', value: 0 }, saved: { kind: 'number', value: 0 } })
    // calc-division-by-zero: Central's Units sum to 0, so Sales/Units is #DIV/0! (R20C4, R21C4); totals are the
    // formula over the totals' sums (Grand Total 222603.36/2417).
    const div = formulaCells(await engine('calc-division-by-zero'), 2)
    expect(disagreements(div)).toEqual([])
    expect(div.filter(cell => (cell.saved as { kind: string }).kind === 'error').map(cell => `R${cell.row + 1}C${cell.col + 1}`)).toEqual(['R20C4', 'R21C4'])
    // calc-quoted-field-name: 'Net Amount'/Sales; a cell without records is 0/0, #DIV/0! (R22C4).
    expect(disagreements(formulaCells(await engine('calc-quoted-field-name'), 2))).toEqual([])
  })

  it('a reference to a calculated field re-applies its formula over the same sums (calc-references-calculated)', async () => {
    const p = await engine('calc-references-calculated')
    // Double Price = 'Price per Unit'*2, Price per Unit = Sales/Units: East is (48118.96/623)*2 (R4C4).
    const compiled = compileCalculatedField(p.cache.fields[13]!.formula!, p.cache, 13)
    expect(compiled.ok && compiled.value.fields).toEqual([3, 4])
    const cells = formulaCells(p, 2)
    expect(disagreements(cells)).toEqual([])
    expect(cells[0]).toMatchObject({ row: 3, col: 3, saved: { kind: 'number', value: 154.4749919743178 } })
  })

  it('text values are left out of the sum; the data field\'s function is ignored (calc-text-field, calc-function-count)', async () => {
    // Channel*2: Channel has text items and one number item (42), whose records give 84 (R13C13).
    expect(disagreements(formulaCells(await engine('calc-text-field'), 2))).toEqual([])
    // Count of Double Sales still shows Sales*2 over the sums (R10C7 North 2023: 47551.48).
    const count = await engine('calc-function-count')
    expect(count.def.dataFields[2]!.subtotal).toBe('count')
    expect(disagreements(formulaCells(count, 2))).toEqual([])
  })

  it('a formula over no field fills every value cell, cells without records included (calc-constant)', async () => {
    const p = await engine('calc-constant')
    const saved = savedPivotLayout(p.def)!, read = sheetCellReader(p.sheet)
    const values: unknown[] = []
    // Columns: Year × (Sales, Units, Flat One), then the three grand columns; Flat One is every third column.
    for (let row = saved.dataRow0; row <= saved.rect.maxRow; row++) for (let col = saved.dataCol0 + 2; col <= saved.rect.maxCol; col += 3) values.push(read(row, col))
    expect(new Set(values.map(value => JSON.stringify(value)))).toEqual(new Set([JSON.stringify({ kind: 'number', value: 1 })]))
    // R22C4 (Central x 2022) has no records and still shows 1. The engine keeps its gate until values.ts sums no
    // field (S2-AA map below).
    expect(read(21, 3)).toEqual({ kind: 'number', value: 1 })
  })

  it('named subtotals: calculated cells are blank on named-subtotal lines except at the column grand total (calc-named-subtotals)', async () => {
    const cells = formulaCells(await engine('calc-named-subtotals'), 2)
    const named = (line: PivotLine) => line.type !== 'data' && line.type !== 'grand' && line.type !== 'default'
    expect(disagreements(cells.filter(cell => !named(cell.rowLine)))).toEqual([])
    for (const cell of cells.filter(cell => named(cell.rowLine))) {
      if (cell.colLine.type !== 'grand') { expect(cell.saved, `R${cell.row + 1}C${cell.col + 1}`).toBeUndefined(); continue }
      // At the column grand total both the sum and the count line show the item's formula: East 96237.92 = 48118.96*2.
      const sumLine = cells.find(other => other.rowLine.type === 'sum' && other.rowLine.items[0] === cell.rowLine.items[0] && other.colLine === cell.colLine)!
      expect(cell.saved, `R${cell.row + 1}C${cell.col + 1}`).toEqual(sumLine.saved)
    }
  })

  it('calculated items: the cache item list parses, and the engine refuses with pivot-calculated-item (calc-item)', async () => {
    const p = await engine('calc-item')
    expect(p.cache.unsupported).toEqual([])
    expect(p.def.unsupported).toEqual([])
    expect(p.cache.calculatedItems).toEqual([{
      formula: 'Product[Apples]+Product[Bananas]',
      area: expect.objectContaining({ cacheIndex: true, outline: false, fieldPosition: 0, references: [expect.objectContaining({ field: 1, x: [4] })] }),
    }])
    expect(checkSavedLayout(p.store, p.cache, p.def, sheetCellReader(p.sheet), { date1904: p.date1904 }))
      .toEqual({ status: 'unverified', reason: 'unsupported', detail: 'pivot-calculated-item' })
  })
})
