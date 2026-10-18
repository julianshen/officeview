/**
 * S1.3 adapter parity under one declared value policy (spec §§4.2, 4.4): the
 * object adapter over a `cached` parse and the paged adapter over the same
 * bytes publish identical authoritative values. Engine-versus-cached
 * differences are enumerated case by case and never asserted equal.
 */
import { describe, expect, it, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import * as workbook from '../src/xlsx/formula/workbook'
import { createObjectWorkbookSource } from '../src/xlsx/view/source'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { SourcePageStore, memoryPageBackend } from '../src/xlsx/storage/sourcePages'
import type { WorkbookSource } from '../src/worker/viewerSession'
import type { CellRead, ColumnLanes, SourceStore } from '../src/xlsx/view/contracts'
import { FORMULA_VALUE_UNAVAILABLE, SAVED_VALUES_STATUS, countLoadEvaluatedFormulas } from '../src/xlsx/view/values'
import { loadOfficeFile } from '../src/components/OfficeFile'
import { xlsxEntries } from './helpers/xlsx-package'
import { buildZip } from './helpers/zip-builder'

vi.mock('../src/xlsx/formula/workbook', async (importOriginal) => {
  const original = await importOriginal<typeof import('../src/xlsx/formula/workbook')>()
  return { ...original, evaluateWorkbookFormulas: vi.fn(original.evaluateWorkbookFormulas) }
})
const evaluator = vi.mocked(workbook.evaluateWorkbookFormulas)

const DYNAMIC_METADATA = '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:da="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><da:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'

/** Cells: [A1 address, sheet XML, expected cached authority]. */
const CASES = [
  // numbers / booleans / blanks
  ['A1', '<c r="A1"><v>1.5</v></c>', { kind: 'cell', value: 1.5, valueIsError: false }],
  ['B1', '<c r="B1"><v>-0</v></c>', { kind: 'cell', value: -0, valueIsError: false }],
  ['C1', '<c r="C1" t="b"><v>1</v></c>', { kind: 'cell', value: true, valueIsError: false }],
  ['D1', '<c r="D1" t="b"><v>0</v></c>', { kind: 'cell', value: false, valueIsError: false }],
  ['E1', '<c r="E1" s="1"/>', { kind: 'cell', value: null, valueIsError: false }],
  // error tokens versus literal error-looking text
  ['A2', '<c r="A2" t="e"><v>#VALUE!</v></c>', { kind: 'cell', value: '#VALUE!', valueIsError: true }],
  ['B2', '<c r="B2" t="str"><v>#VALUE!</v></c>', { kind: 'cell', value: '#VALUE!', valueIsError: false }],
  ['C2', '<c r="C2" t="inlineStr"><is><t>#N/A</t></is></c>', { kind: 'cell', value: '#N/A', valueIsError: false }],
  ['D2', '<c r="D2" t="s"><v>1</v></c>', { kind: 'cell', value: '#DIV/0!', valueIsError: false }],
  // file caches (including a stale cache the engine would keep or recompute)
  ['A3', '<c r="A3"><f>A1*2</f><v>3</v></c>', { kind: 'cell', value: 3, valueIsError: false, formula: 'A1*2' }],
  ['B3', '<c r="B3" t="e"><f>1/0</f><v>#DIV/0!</v></c>', { kind: 'cell', value: '#DIV/0!', valueIsError: true, formula: '1/0' }],
  ['C3', '<c r="C3" t="str"><f>""</f><v></v></c>', { kind: 'cell', value: '', valueIsError: false, formula: '""' }],
  ['D3', '<c r="D3" t="b"><f>FALSE()</f><v>0</v></c>', { kind: 'cell', value: false, valueIsError: false, formula: 'FALSE()' }],
  // absent formula caches: unavailable, never blank/zero/error
  ['A4', '<c r="A4"><f>A1+1</f></c>', { kind: 'unavailable', formula: 'A1+1' }],
  ['B4', '<c r="B4"><f>A1+1</f><v></v></c>', { kind: 'unavailable', formula: 'A1+1' }],
  ['C4', '<c r="C4" t="e"><f>NA()</f></c>', { kind: 'unavailable', formula: 'NA()' }],
  // unsupported formula: the file cache is preserved verbatim
  ['A5', '<c r="A5"><f>_xlfn.NOSUCHFN(1)</f><v>42</v></c>', { kind: 'cell', value: 42, valueIsError: false, formula: '_xlfn.NOSUCHFN(1)' }],
  ['B5', '<c r="B5" t="str"><f>WEBSERVICE("x")</f><v>saved</v></c>', { kind: 'cell', value: 'saved', valueIsError: false, formula: 'WEBSERVICE("x")' }],
  // date styles keep the serial and style index
  ['A6', '<c r="A6" s="2"><v>45292</v></c>', { kind: 'cell', value: 45292, valueIsError: false, styleIndex: 2 }],
  ['B6', '<c r="B6" s="2"><f>A6+1</f><v>45293</v></c>', { kind: 'cell', value: 45293, valueIsError: false, styleIndex: 2, formula: 'A6+1' }],
  // lone UTF-16 surrogates in shared and inline strings
  ['A7', '<c r="A7" t="s"><v>0</v></c>', { kind: 'cell', value: 'a\ud800b', valueIsError: false }],
  ['B7', '<c r="B7" t="inlineStr"><is><t>x_xDC00_</t></is></c>', { kind: 'cell', value: 'x\udc00', valueIsError: false }],
] as const

// Saved spill children: a dynamic anchor (cm="1") and a legacy array whose
// children carry saved values but no formula; one child has no saved value.
const SPILL_ROWS = [
  '<row r="10"><c r="A10" cm="1"><f t="array" ref="A10:A12">SEQUENCE(3)</f><v>10</v></c><c r="C10"><f t="array" ref="C10:D10">{1,2}</f><v>1</v></c><c r="D10"><v>2</v></c></row>',
  '<row r="11"><c r="A11"><v>20</v></c></row>',
  '<row r="12"><c r="A12"/></row>',
].join('')

function packageBytes(): Uint8Array {
  const byRow = new Map<number, string[]>()
  for (const [ref, xml] of CASES) {
    const row = Number(ref.slice(1))
    byRow.set(row, [...(byRow.get(row) ?? []), xml])
  }
  const sheetData = [...byRow].sort((a, b) => a[0] - b[0]).map(([row, cells]) => `<row r="${row}">${cells.join('')}</row>`).join('') + SPILL_ROWS
  const entries = xlsxEntries({
    sheetData,
    sharedStrings: ['a_xD800_b', '#DIV/0!'],
    styles: '<numFmts count="0"/><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>',
  })
  const rels = entries.find(e => e.name === 'xl/_rels/workbook.xml.rels')!
  rels.data = String(rels.data).replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>')
  entries.push({ name: 'xl/metadata.xml', data: DYNAMIC_METADATA })
  return buildZip(entries).bytes
}

const opts = (store: SourceStore) => ({ signal: new AbortController().signal, stamp: { ...store.stamp } })
const a1 = (ref: string) => ({ row: Number(ref.slice(1)) - 1, col: ref.charCodeAt(0) - 65 })

async function bothAdapters(bytes: Uint8Array): Promise<{ object: WorkbookSource; paged: WorkbookSource }> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const object = createObjectWorkbookSource(doc, 'doc')
  const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
  const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store, pageRows: 4 })
  expect((await session.done).kind).toBe('complete')
  return { object, paged: createPagedWorkbookSource(session) }
}
function authority(read: CellRead) {
  if (read.kind !== 'cell' && read.kind !== 'unavailable') return { kind: read.kind }
  const { value, valueIsError, valueProvenance, formula, styleIndex } = read.cell
  return { kind: read.kind, value, valueIsError, valueProvenance, formula, styleIndex }
}

describe('adapter parity under the cached policy (S1.3 matrix)', () => {
  it('every case reads identically on both adapters and matches the declared cached authority', async () => {
    evaluator.mockClear()
    const { object, paged } = await bothAdapters(packageBytes())
    expect(evaluator).not.toHaveBeenCalled()
    expect([object.valuePolicy, paged.valuePolicy]).toEqual(['cached', 'cached'])
    const a = object.store('1'), b = paged.store('1')
    for (const [ref, , expected] of CASES) {
      const left = await a.readCell(a1(ref), opts(a)), right = await b.readCell(a1(ref), opts(b))
      expect(right, `${ref} parity`).toEqual(left)
      const got = authority(left)
      expect(got.valueProvenance, `${ref} provenance`).toBe('cached')
      if (expected.kind === 'unavailable') {
        expect(left, ref).toMatchObject({ kind: 'unavailable', reason: FORMULA_VALUE_UNAVAILABLE, cell: { formula: expected.formula } })
        // The withheld value is never published as blank, zero or an error.
        expect(left.kind === 'unavailable' && left.cell.valueIsError === true).toBe(false)
        continue
      }
      expect(got.kind, ref).toBe('cell')
      expect(Object.is(got.value, expected.value), `${ref} value ${String(got.value)}`).toBe(true)
      expect(got.valueIsError, ref).toBe(expected.valueIsError)
      if ('formula' in expected) expect(got.formula, ref).toBe(expected.formula)
      if ('styleIndex' in expected) expect(got.styleIndex, ref).toBe(expected.styleIndex)
    }
    await paged.release()
  })

  it('saved spill children keep their saved values (no formula, never unavailable) on both adapters', async () => {
    const { object, paged } = await bothAdapters(packageBytes())
    const a = object.store('1'), b = paged.store('1')
    for (const [at, value] of [[{ row: 9, col: 0 }, 10], [{ row: 10, col: 0 }, 20], [{ row: 9, col: 3 }, 2]] as const) {
      const left = await a.readCell(at, opts(a))
      expect(await b.readCell(at, opts(b))).toEqual(left)
      expect(left).toMatchObject({ kind: 'cell', cell: { value, valueProvenance: 'cached' } })
    }
    // A child record with no saved value is a blank record, not an unavailable formula.
    const empty = await a.readCell({ row: 11, col: 0 }, opts(a))
    expect(await b.readCell({ row: 11, col: 0 }, opts(b))).toEqual(empty)
    expect(empty.kind).not.toBe('unavailable')
    await paged.release()
  })

  it('typed lanes agree on both adapters: error token vs text, unavailable, lone surrogates', async () => {
    const { object, paged } = await bothAdapters(packageBytes())
    for (const col of [0, 1, 2, 3]) {
      const query = { col, rows: { kind: 'intervals' as const, intervals: [{ r0: 0, r1: 20 }] }, limit: 100, encoding: 'lanes' as const }
      const a = object.store('1'), b = paged.store('1')
      const left = await a.queryColumn(query, opts(a)), right = await b.queryColumn(query, opts(b))
      expect({ ...right, stamp: null }, `col ${col}`).toEqual({ ...left, stamp: null })
      expect((left.items[0] as ColumnLanes).kind).toBe('lanes')
    }
    await paged.release()
  })

  it('both cached sources state the saved-values status; the paged one gains the count once imported', async () => {
    const { object, paged } = await bothAdapters(packageBytes())
    const formulas = 3
    expect(object.valueStatus?.()).toEqual({ policy: 'cached', formulas, message: `Showing saved values — ${formulas} formulas not recalculated (file too large)` })
    expect(paged.valueStatus?.()).toEqual(object.valueStatus?.())
    expect(paged.openDiagnostics).toContainEqual({ code: 'cached-value-policy', message: expect.stringMatching(/^Showing saved values/) })
    await paged.release()
  })

  it('a paged source still importing states the generic status without a count', async () => {
    const bytes = packageBytes()
    const store = new SourcePageStore({ docId: 'doc', backend: memoryPageBackend() })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([bytes as BlobPart]) }, { docId: 'doc', profile: 'D', store, pageRows: 4 })
    const source = createPagedWorkbookSource(session)
    const early = source.valueStatus?.()
    expect(early?.policy).toBe('cached')
    if (early?.formulas === undefined) {
      expect(early?.message).toBe(SAVED_VALUES_STATUS)
      expect(source.openDiagnostics).toEqual([{ code: 'cached-value-policy', message: SAVED_VALUES_STATUS }])
    }
    await session.done
    expect(source.valueStatus?.().formulas).toBe(3)
    // The open notice is refreshed: read after completion it carries the counted wording.
    expect(source.openDiagnostics).toEqual([{ code: 'cached-value-policy', message: 'Showing saved values — 3 formulas not recalculated (file too large)' }])
    await source.release()
  })

  it('paged reads never change valueGen: committed pages are the single value publication', async () => {
    const { object, paged } = await bothAdapters(packageBytes())
    expect(paged.store('1').stamp.valueGen).toBe(0)
    expect(object.store('1').stamp.valueGen).toBe(0)
    await paged.release()
  })
})

describe('engine versus cached: enumerated differences, never asserted equal', () => {
  async function pair(sheetData: string, fullCalcOnLoad = false) {
    const bytes = buildZip(xlsxEntries({ sheetData, fullCalcOnLoad })).bytes
    const engine = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(bytes)), 'e').store('1')
    const cached = createObjectWorkbookSource(await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' }), 'c').store('1')
    return { engine, cached }
  }
  const at = { row: 0, col: 1 }

  it('1. an uncached supported formula: engine publishes the computed value, cached is unavailable', async () => {
    const { engine, cached } = await pair('<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>A1*2</f></c></row>')
    expect(authority(await engine.readCell(at, opts(engine)))).toMatchObject({ kind: 'cell', value: 8, valueProvenance: 'engine' })
    expect(authority(await cached.readCell(at, opts(cached)))).toMatchObject({ kind: 'unavailable', valueProvenance: 'cached' })
  })

  it('2. fullCalcOnLoad over a stale cache: engine recalculates, cached keeps the saved value', async () => {
    const { engine, cached } = await pair('<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>A1*2</f><v>99</v></c></row>', true)
    expect(authority(await engine.readCell(at, opts(engine)))).toMatchObject({ kind: 'cell', value: 8 })
    expect(authority(await cached.readCell(at, opts(cached)))).toMatchObject({ kind: 'cell', value: 99 })
  })

  it('3. a ca="1" (calculate-always) formula: engine re-evaluates over its cache, cached keeps the saved value', async () => {
    const { engine, cached } = await pair('<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f ca="1">A1*2</f><v>99</v></c></row>')
    expect(authority(await engine.readCell(at, opts(engine)))).toMatchObject({ kind: 'cell', value: 8 })
    expect(authority(await cached.readCell(at, opts(cached)))).toMatchObject({ kind: 'cell', value: 99 })
  })

  it('control: a file-load cache both policies keep is the same value with different provenance', async () => {
    const { engine, cached } = await pair('<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>A1*2</f><v>99</v></c></row>')
    expect(authority(await engine.readCell(at, opts(engine)))).toMatchObject({ value: 99, valueProvenance: 'engine' })
    expect(authority(await cached.readCell(at, opts(cached)))).toMatchObject({ value: 99, valueProvenance: 'cached' })
  })
})

describe('N in the saved-values status = formulas the engine would have evaluated at load (coordinator decision Q1)', () => {
  // Load-evaluated: B1 uncached, D1 cached but ca="1", E1 cached volatile NOW(). Kept: C1, F1 (cached, ordinary).
  const MIXED = '<row r="1"><c r="A1"><v>4</v></c><c r="B1"><f>A1+1</f></c><c r="C1"><f>A1*2</f><v>8</v></c><c r="D1"><f ca="1">A1*3</f><v>12</v></c><c r="E1"><f>NOW()</f><v>45000</v></c><c r="F1"><f>A1-1</f><v>3</v></c></row>'
  const message = (n: number) => `Showing saved values — ${n} formulas not recalculated (file too large)`

  it('object and paged sources agree on a mixed cached/uncached workbook; fullCalcOnLoad counts every formula', async () => {
    for (const [fullCalcOnLoad, n] of [[false, 3], [true, 5]] as const) {
      const { object, paged } = await bothAdapters(buildZip(xlsxEntries({ sheetData: MIXED, fullCalcOnLoad })).bytes)
      expect(object.valueStatus?.()).toEqual({ policy: 'cached', formulas: n, message: message(n) })
      expect(paged.valueStatus?.()).toEqual(object.valueStatus?.())
      await paged.release()
    }
  })

  it('the loader notice uses the same definition (cap-hit fallback over a mixed workbook)', async () => {
    // An uncached SEQUENCE(200000,1) trips the evaluator cap, so the file reopens with saved values.
    const sheetData = MIXED.replace('</row>', '<c r="G1"><f>SEQUENCE(200000,1)</f></c></row>')
    const doc = await loadOfficeFile(buildZip(xlsxEntries({ sheetData })).bytes) as { diagnostics?: Array<{ feature: string; message: string }> }
    expect(countLoadEvaluatedFormulas(doc as never)).toBe(4)
    // Same count definition; the cause is the evaluator limit, not the file's size.
    expect(doc.diagnostics?.find(d => d.feature === 'cached-value-policy')?.message).toBe(message(4).replace('(file too large)', '(formula evaluation limit)'))
  })
})
