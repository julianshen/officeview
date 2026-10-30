// @vitest-environment node
/**
 * The faithfulness check (spec §6.4, §7.6, §7.7; plan Task 19): the engine recomputes the SAVED state of a pivot and
 * proves it equals what Excel saved, or the pivot stays display-only. Oracles: the 103 Excel-saved Stage 1
 * references, the Task 15 manifest once it exists, and the gitignored corpus (status recorded, never tuned toward).
 */
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { savedPivotLayout, savedPivotRoles } from '../src/xlsx/pivots/savedLayout'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import { decodeRecords, RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { computePivot } from '../src/xlsx/pivots/engine/compute'
import { initialViewState } from '../src/xlsx/pivots/engine/view'
import {
  checkSavedLayout, engineEligibility, pagedCellReader, sheetCellReader, PAGED_CELL_BYTES_ESTIMATE, PAGED_CELL_LIMITS,
  type CellReader, type FaithfulResult, type SavedCell,
} from '../src/xlsx/pivots/engine/faithful'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import type { SourceStore } from '../src/xlsx/view/contracts'
import { MAIN_THREAD_BUDGETS } from '../src/xlsx/view/tileCache'
import { decodeParts, loadAll, loadPivotParts, once, type Loaded, type PivotParts } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const CORPUS = resolve(process.cwd(), 'corpus')
const MANIFEST = resolve(FIXTURES, 'engine/manifest.json')
const REFERENCES = readdirSync(FIXTURES).filter(name => /^pivot-.*\.xlsx$/.test(name)).sort()

const cached = new Map<string, Promise<Loaded[]>>()
const load = (path: string) => { let p = cached.get(path); if (!p) cached.set(path, p = loadAll(path)); return p }
const ref = async (name: string) => (await load(resolve(FIXTURES, `${name}.xlsx`)))[0]!
const check = (p: Loaded, saved: CellReader = sheetCellReader(p.sheet), over: Partial<Pick<Loaded, 'def' | 'cache'>> = {}) =>
  checkSavedLayout(p.store, over.cache ?? p.cache, over.def ?? p.def, saved, { date1904: p.date1904 })
/** A reader with some cells replaced (`null` clears one). */
const edited = (base: CellReader, edits: Record<string, SavedCell | null>): CellReader =>
  (row, col) => { const key = `${row},${col}`; return key in edits ? edits[key] ?? undefined : base(row, col) }

/** Engine refusals that name the later task which removes them (plan Stage 2). Any other refusal on an Excel-made
 * file is a finding, not a pass. */
const PENDING_TASK: Record<string, string> = {
  'pivot-engine-pending': 'Task 21 (sort, filters, grouping) or Task 22 (calculated fields)',
  'pivot-engine-unmeasured': 'Task 15 (Excel oracle measurement)',
}
const KNOWN_UNSUPPORTED = new Set(['pivot-filter-dynamic-date'])
const describeResult = (result: FaithfulResult) => result.status === 'mismatch'
  ? `mismatch R${result.firstDiff.row + 1}C${result.firstDiff.col + 1}: ${result.firstDiff.reason}`
  : result.status === 'unverified' ? `unverified:${result.reason}${result.detail ? `:${result.detail}` : ''}` : 'ok'
function expectOkOrPending(result: FaithfulResult, where: string) {
  if (result.status === 'ok') return
  expect(result.status === 'unverified' && result.reason === 'unsupported' && result.detail !== undefined &&
    (PENDING_TASK[result.detail] !== undefined || KNOWN_UNSUPPORTED.has(result.detail)),
    `${where}: ${describeResult(result)}`).toBe(true)
}

describe('every Excel-made fixture is ok or unverified with a named pending task', () => {
  it.each(REFERENCES)('%s', async name => {
    for (const p of await load(resolve(FIXTURES, name))) {
      expect(engineEligibility(p.cache, p.def, { sourceReadable: false }), name).toEqual({ ok: true, refresh: false })
      const result = check(p)
      expectOkOrPending(result, `${name} ${p.def.name}`)
      // Tasks 17 and 18 compute every Stage 1 reference: each is a full, verified pass today.
      expect(describeResult(result), name).toBe('ok')
    }
  })

  interface ManifestCase { case: string; file: string; features: string[]; task: number; pivot: string; status?: string }
  const manifest: ManifestCase[] | undefined = existsSync(MANIFEST)
    ? (raw => Array.isArray(raw) ? raw : raw.cases)(JSON.parse(readFileSync(MANIFEST, 'utf8')))
    : undefined
  if (!manifest) {
    it.skip('SKIPPED: tests/fixtures/xlsx-pivots/engine/manifest.json is absent (Task 15 oracle not built); manifest cases not run', () => {})
  } else {
    const runnable = manifest.filter(entry => entry.status !== 'pending')
    if (!runnable.length) it.skip('SKIPPED: the manifest has no runnable case (all pending)', () => {})
    else it.each(runnable.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
      const local = resolve(FIXTURES, 'engine', entry.file)
      const p = (await load(existsSync(local) ? local : resolve(process.cwd(), entry.file))).find(candidate => candidate.def.name === entry.pivot)
      if (!p) throw new Error(`${entry.case}: no pivot named ${entry.pivot} in ${entry.file}`)
      expectOkOrPending(check(p), entry.case)
    })
  }
})

describe('mismatch and unverified paths', () => {
  it('a one-cell edit is a mismatch naming that cell', async () => {
    const p = await ref('pivot-layout-compact')
    const saved = savedPivotLayout(p.def)!
    const base = sheetCellReader(p.sheet)
    // A value cell: the first data row's first value.
    const valueAt = { row: saved.dataRow0, col: saved.dataCol0 }
    const value = base(valueAt.row, valueAt.col)
    expect(value?.kind).toBe('number')
    const changed: SavedCell = { kind: 'number', value: (value as { value: number }).value + 1 }
    const result = check(p, edited(base, { [`${valueAt.row},${valueAt.col}`]: changed }))
    expect(result.status).toBe('mismatch')
    if (result.status !== 'mismatch') return
    expect(result.firstDiff).toMatchObject({ row: valueAt.row, col: valueAt.col, expected: changed, actual: value })
    // A label cell: the first row label.
    const labelAt = { row: saved.dataRow0, col: saved.rect.minCol }
    const label = check(p, edited(base, { [`${labelAt.row},${labelAt.col}`]: { kind: 'text', value: 'Elsewhere' } }))
    expect(label.status === 'mismatch' && { row: label.firstDiff.row, col: label.firstDiff.col }).toEqual(labelAt)
    // 15 significant digits: one unit of the 15th digit of this value is 10^(digits - 15).
    const v = (value as { value: number }).value
    const unit = 10 ** (Math.floor(Math.log10(Math.abs(v))) + 1 - 15)
    const within = v + 0.4 * unit, across = v + 0.6 * unit
    // Both perturbations are different Numbers; only the second changes the 15th significant digit.
    expect(within).not.toBe(v)
    expect(Number(within.toPrecision(15))).toBe(Number(v.toPrecision(15)))
    expect(Number(across.toPrecision(15))).not.toBe(Number(v.toPrecision(15)))
    const key = `${valueAt.row},${valueAt.col}`
    expect(check(p, edited(base, { [key]: { kind: 'number', value: within } })).status).toBe('ok')
    const crossed = check(p, edited(base, { [key]: { kind: 'number', value: across } }))
    expect(crossed.status === 'mismatch' && [crossed.firstDiff.row, crossed.firstDiff.col]).toEqual([valueAt.row, valueAt.col])
  })

  it('a cleared value cell is a mismatch (a blank saved value means no records), never cells-missing', async () => {
    const p = await ref('pivot-layout-compact')
    const saved = savedPivotLayout(p.def)!
    const result = check(p, edited(sheetCellReader(p.sheet), { [`${saved.dataRow0},${saved.dataCol0}`]: null }))
    expect(result.status === 'mismatch' && [result.firstDiff.row, result.firstDiff.col, result.firstDiff.expected]).toEqual([saved.dataRow0, saved.dataCol0, undefined])
  })

  it('a saved value the engine leaves blank is a mismatch', async () => {
    const p = await ref('pivot-layout-compact')
    const saved = savedPivotLayout(p.def)!
    // Find a blank value cell in the body (a combination with no records).
    const base = sheetCellReader(p.sheet)
    let blank: { row: number; col: number } | undefined
    for (let row = saved.dataRow0; row <= saved.rect.maxRow && !blank; row++) for (let col = saved.dataCol0; col <= saved.rect.maxCol && !blank; col++) if (!base(row, col)) blank = { row, col }
    expect(blank, 'pivot-layout-compact has an empty combination').toBeDefined()
    if (!blank) return
    const result = check(p, edited(base, { [`${blank.row},${blank.col}`]: { kind: 'number', value: 7 } }))
    expect(result.status === 'mismatch' && { row: result.firstDiff.row, col: result.firstDiff.col, actual: result.firstDiff.actual }).toEqual({ ...blank, actual: { kind: 'empty' } })
  })

  it('an implicit caption is compared for presence, not text', async () => {
    // These zh-TW references save localized default captions: the engine's English placeholders still pass...
    const p = await ref('pivot-layout-compact')
    expect(check(p).status).toBe('ok')
    // ...but a non-text saved cell under a default caption does not. Take the first implicit caption the engine places.
    const layout = computePivot(p.store, p.cache, p.def, initialViewState(p.def, p.cache), { date1904: p.date1904 })
    if (!layout.ok) throw new Error(layout.diagnostic.message)
    const caption = layout.value.cells.find(cell => cell.caption === 'implicit')!
    expect(caption).toBeDefined()
    const at = `${caption.row},${caption.col}`
    const base = sheetCellReader(p.sheet)
    const savedText = base(caption.row, caption.col)
    expect(savedText?.kind).toBe('text')
    expect(savedText).not.toEqual(caption.value)
    expect(check(p, edited(base, { [at]: { kind: 'text', value: 'any text' } })).status).toBe('ok')
    const number = check(p, edited(base, { [at]: { kind: 'number', value: 1 } }))
    expect(number.status === 'mismatch' && [number.firstDiff.row, number.firstDiff.col]).toEqual([caption.row, caption.col])
    expect(check(p, edited(base, { [at]: null }))).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
  })

  it('a line that differs from the saved rowItems is a line mismatch', async () => {
    const p = await ref('pivot-layout-compact')
    const rowItems = p.def.rowItems!.map((item, k) => k === 1 ? { ...item, x: [...item.x.slice(0, -1), item.x[item.x.length - 1]! + 1] } : item)
    const result = check(p, undefined, { def: { ...p.def, rowItems } })
    expect(result.status === 'mismatch' && result.firstDiff.actual).toBe('line')
    if (result.status === 'mismatch') expect(result.firstDiff.row).toBe(savedPivotLayout(p.def)!.dataRow0 + 1)
  })

  it('absent rowItems is unverified', async () => {
    const p = await ref('pivot-layout-compact')
    const { rowItems: _rowItems, ...def } = p.def
    expect(check(p, undefined, { def })).toEqual({ status: 'unverified', reason: 'no-row-items' })
  })

  it('absent colItems is unverified', async () => {
    // pivot-layout-compact has a column axis (Year, then Values).
    const p = await ref('pivot-layout-compact')
    expect(p.def.colFields.length).toBeGreaterThan(0)
    const { colItems: _colItems, ...def } = p.def
    expect(check(p, undefined, { def })).toEqual({ status: 'unverified', reason: 'no-col-items' })
  })

  it('missing or incomplete saved cells is unverified (cells-missing)', async () => {
    const p = await ref('pivot-layout-compact')
    const saved = savedPivotLayout(p.def)!
    const base = sheetCellReader(p.sheet)
    // Missing: no saved cell at all.
    expect(check(p, () => undefined)).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
    // Incomplete: the sheet stops before the grand-total row (labels gone, a truncated save).
    const truncated: CellReader = (row, col) => row >= saved.rect.maxRow ? undefined : base(row, col)
    expect(check(p, truncated)).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
    // Incomplete: one label cell absent.
    expect(check(p, edited(base, { [`${saved.dataRow0},${saved.rect.minCol}`]: null }))).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
    // Incomplete: a reader that says it could not read every cell (the paged reader on a failed page).
    const partial: CellReader = Object.assign((row: number, col: number) => base(row, col), { incomplete: 'read-failed' as const })
    expect(check(p, partial)).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
  })

  it('recordCount differing from the store is unverified (cells-predate-records)', async () => {
    const p = await ref('pivot-layout-compact')
    expect(p.cache.recordCount).toBe(p.store.count)
    const cache: PivotCacheDefinition = { ...p.cache, recordCount: p.store.count + 1 }
    expect(check(p, undefined, { cache })).toMatchObject({ status: 'unverified', reason: 'cells-predate-records' })
  })

  it('an engine refusal is unverified (unsupported) with the feature as detail', async () => {
    const p = await ref('pivot-layout-compact')
    const view = p.def.fields.map((field, k) => k === 0 ? { ...field, autoSortScope: { type: 'normal' as const, dataOnly: true, labelOnly: false, grandRow: false, grandCol: false, cacheIndex: false, outline: true, collapsedLevelsAreSubtotals: false, references: [] } } : field)
    expect(check(p, undefined, { def: { ...p.def, fields: view } })).toEqual({ status: 'unverified', reason: 'unsupported', detail: 'pivot-engine-unmeasured' })
  })
})

// ---------------------------------------------------------------- §7.6 refusals, built by editing Excel's XML

const CACHE_PART = 'xl/pivotCache/pivotCacheDefinition1.xml'
const TABLE_PART = 'xl/pivotTables/pivotTable1.xml'
async function editPackage(name: string, edit: (zip: JSZip) => Promise<void>): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(readFileSync(resolve(FIXTURES, `${name}.xlsx`)))
  await edit(zip)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
}
const rewrite = (zip: JSZip, path: string, change: (xml: string) => string) => zip.file(path, zip.file(path)!.async('string').then(change))
async function edit(name: string, change: (zip: JSZip) => Promise<void> | void): Promise<PivotParts> {
  const bytes = await editPackage(name, async zip => { await change(zip) })
  return (await loadPivotParts(`${name} (edited)`, bytes))[0]!
}
/** saveData="0" as Excel writes it: no records part, no relationship to it, no r:id. */
const dropRecords = (refreshOnLoad: boolean) => async (zip: JSZip) => {
  rewrite(zip, CACHE_PART, xml => xml.replace(' r:id="rId1"', ` saveData="0"${refreshOnLoad ? ' refreshOnLoad="1"' : ''}`))
  zip.remove('xl/pivotCache/pivotCacheRecords1.xml')
  zip.remove('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')
  rewrite(zip, '[Content_Types].xml', xml => xml.replace(/<Override PartName="\/xl\/pivotCache\/pivotCacheRecords1\.xml"[^>]*\/>/, ''))
}
const refusal = (result: ReturnType<typeof engineEligibility>) => result.ok ? result : { ok: false, kind: result.diagnostic.kind, feature: result.diagnostic.feature }

describe('§7.6 refusals are typed results, never throws', () => {
  it('saveData="0" without refreshOnLoad: pivot-records-absent', async () => {
    const p = await edit('pivot-layout-compact', dropRecords(false))
    expect(p.cache!.saveData).toBe(false)
    expect(p.cache!.recordsPartPath).toBeUndefined()
    for (const sourceReadable of [false, true]) expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-records-absent' })
  })

  it('saveData="0" with refreshOnLoad: pivot-records-absent unless the source is readable, then the §7.7 refresh route', async () => {
    const p = await edit('pivot-layout-compact', dropRecords(true))
    expect(p.cache!.refreshOnLoad).toBe(true)
    expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable: false }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-records-absent' })
    expect(engineEligibility(p.cache!, p.def, { sourceReadable: true })).toEqual({ ok: true, refresh: true })
  })

  it('refreshOnLoad with records: the §7.7 refresh route when readable, else not interactive (pivot-refresh-on-load)', async () => {
    const p = await edit('pivot-layout-compact', zip => { rewrite(zip, CACHE_PART, xml => xml.replace(' r:id="rId1"', ' r:id="rId1" refreshOnLoad="1"')) })
    expect(engineEligibility(p.cache!, p.def, { sourceReadable: true })).toEqual({ ok: true, refresh: true })
    expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable: false }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-refresh-on-load' })
  })

  it.each([
    ['external', '<cacheSource type="external" connectionId="1"/>'],
    ['consolidation', '<cacheSource type="consolidation"><consolidation><rangeSets count="1"><rangeSet ref="A1:E37" sheet="Data"/></rangeSets></consolidation></cacheSource>'],
    ['scenario', '<cacheSource type="scenario"/>'],
  ])('a %s source: pivot-source-unsupported', async (type, source) => {
    const p = await edit('pivot-layout-compact', zip => { rewrite(zip, CACHE_PART, xml => xml.replace(/<cacheSource\b[\s\S]*?<\/cacheSource>/, source)) })
    expect(p.cache!.source.type).toBe(type)
    for (const sourceReadable of [false, true]) expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-source-unsupported' })
  })

  it('a worksheet source in another workbook: pivot-source-unsupported', async () => {
    const p = await edit('pivot-layout-compact', zip => { rewrite(zip, CACHE_PART, xml => xml.replace('<worksheetSource ref="A1:E37" sheet="Data"/>', '<worksheetSource ref="A1:E37" sheet="Data" r:id="rId9"/>')) })
    expect(p.cache!.source.worksheet?.relId).toBe('rId9')
    expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable: true }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-source-unsupported' })
  })

  it('an OLAP cache: pivot-source-unsupported', async () => {
    const p = await edit('pivot-layout-compact', zip => { rewrite(zip, CACHE_PART, xml => xml.replace('</cacheFields>', '</cacheFields><cacheHierarchies count="0"/>')) })
    expect(p.cache!.olap).toBe(true)
    for (const sourceReadable of [false, true]) expect(refusal(engineEligibility(p.cache!, p.def, { sourceReadable }))).toEqual({ ok: false, kind: 'unsupported', feature: 'pivot-source-unsupported' })
  })

  it('an unknown definition construct: unverified (unsupported)', async () => {
    const parts = await edit('pivot-layout-compact', zip => { rewrite(zip, TABLE_PART, xml => xml.replace('<pivotTableDefinition ', '<pivotTableDefinition futureFlag="1" ')) })
    expect(parts.def.unsupported.length).toBeGreaterThan(0)
    expect(engineEligibility(parts.cache!, parts.def, { sourceReadable: false })).toEqual({ ok: true, refresh: false })
    const p = await decodeParts(parts)
    expect(check(p)).toEqual({ status: 'unverified', reason: 'unsupported', detail: 'pivot-definition-unsupported' })
  })
})

describe("a refused pivot's display is unchanged", () => {
  const deepFreeze = <T>(value: T): T => {
    if (value && typeof value === 'object' && !ArrayBuffer.isView(value) && !Object.isFrozen(value)) { Object.freeze(value); for (const child of Object.values(value)) deepFreeze(child) }
    return value
  }
  /** savedPivotRoles evaluated over every painted cell (it returns functions, so the snapshot is their output). */
  const rolesOf = (def: PivotTableDefinition) => {
    const roles = savedPivotRoles(def)
    if (!roles) return undefined
    const cells: unknown[] = []
    for (const rect of roles.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) cells.push([row, col, roles.rowClass(row), roles.roleAt(row, col)])
    return { rect: roles.rect, rects: roles.rects, cells }
  }
  it('savedPivotRoles and the sheet cells are identical before and after the check', async () => {
    const unknown = await decodeParts(await edit('pivot-layout-compact', zip => { rewrite(zip, TABLE_PART, xml => xml.replace('<pivotTableDefinition ', '<pivotTableDefinition futureFlag="1" ')) }))
    const compact = await decodeParts((await loadPivotParts(resolve(FIXTURES, 'pivot-layout-compact.xlsx')))[0]!)
    const saved = savedPivotLayout(compact.def)!
    const mismatchReader = edited(sheetCellReader(compact.sheet), { [`${saved.dataRow0},${saved.dataCol0}`]: { kind: 'number', value: -1 } })
    const cases: Array<[Loaded, CellReader, FaithfulResult['status']]> = [[unknown, sheetCellReader(unknown.sheet), 'unverified'], [compact, mismatchReader, 'mismatch']]
    for (const [p, reader, status] of cases) {
      const roles = rolesOf(p.def), rows = structuredClone(p.sheet.rows), def = structuredClone(p.def), cache = structuredClone(p.cache)
      expect(roles).toBeDefined()
      // Frozen inputs: any write by the check throws.
      deepFreeze(p.def); deepFreeze(p.cache); deepFreeze(p.sheet.rows)
      expect(engineEligibility(p.cache, p.def, { sourceReadable: false }).ok).toBe(true)
      expect(checkSavedLayout(p.store, p.cache, p.def, reader, { date1904: p.date1904 }).status).toBe(status)
      expect(rolesOf(p.def)).toEqual(roles)
      expect(p.sheet.rows).toEqual(rows)
      expect(p.def).toEqual(def)
      expect(p.cache).toEqual(cache)
    }
  })
})

// ---------------------------------------------------------------- the paged reader

describe('paged CellReader equals resident', () => {
  const PAGED = ['pivot-layout-compact', 'pivot-layout-outline', 'pivot-layout-tabular', 'pivot-page-field', 'pivot-values-on-rows', 'pivot-repeat-labels', 'pivot-subtotals-top', 'pivot-grand-totals-off']
  it.each(PAGED)('%s', async name => {
    const input = new Uint8Array(readFileSync(resolve(FIXTURES, `${name}.xlsx`)))
    const p = await ref(name)
    // The paged path exactly as tests/xlsx-pivots-paged.test.ts opens it: the worker import session and its paged source.
    const session = await openImportSession({ kind: 'blob', blob: new Blob([input]) }, { docId: `faithful-${name}`, profile: 'D' })
    await session.done
    const source = createPagedWorkbookSource(session)
    try {
      const sheet = source.sheets.find(item => item.name === p.def.sheet)!
      const store = source.store(sheet.sheetId)
      const paged = await pagedCellReader(store, p.def, { signal: new AbortController().signal, stamp: store.stamp }, 'D')
      expect(paged.incomplete).toBeUndefined()
      const resident = sheetCellReader(p.sheet)
      const saved = savedPivotLayout(p.def)!
      let compared = 0
      for (const rect of saved.rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
        expect(paged(row, col), `${name} R${row + 1}C${col + 1}`).toEqual(resident(row, col))
        if (resident(row, col)) compared++
      }
      expect(compared).toBeGreaterThan(0)
      expect(check(p, paged)).toEqual(check(p, resident))
      expect(check(p, paged).status).toBe('ok')
      // A read that cannot complete (here: cancelled) gives an incomplete reader, so the check is cells-missing.
      const aborted = new AbortController(); aborted.abort()
      const failed = await pagedCellReader(store, p.def, { signal: aborted.signal, stamp: store.stamp }, 'D')
      expect(failed.incomplete).toBe('read-failed')
      expect(check(p, failed)).toMatchObject({ status: 'unverified', reason: 'cells-missing' })
      // A cap below the area's cells: the reader stops, keeps nothing, and the check is cells-over-cap.
      const capped = await pagedCellReader(store, p.def, { signal: new AbortController().signal, stamp: store.stamp }, 'D', PAGED_CELL_BYTES_ESTIMATE.cell * 2)
      expect(capped.incomplete).toBe('over-cap')
      expect(check(p, capped)).toMatchObject({ status: 'unverified', reason: 'cells-over-cap' })
    } finally { await source.release(); await session.release() }
  })

  it('caps the retained cells per profile (a quarter of the main-thread budget)', () => {
    expect(PAGED_CELL_LIMITS).toEqual({ D: MAIN_THREAD_BUDGETS.D / 4, M: MAIN_THREAD_BUDGETS.M / 4 })
  })

  it.each(['D', 'M'] as const)('a pivot area over the %s cap is unverified (cells-over-cap) and reading stops at the cap', async profile => {
    const p = await ref('pivot-layout-compact')
    // A synthetic pivot: compact's definition with 400,000 data lines, so its area holds far more cells than any cap.
    const lines = 400_000
    const anchor = savedPivotLayout(p.def)!.rect
    const def: PivotTableDefinition = {
      ...p.def,
      location: { ...p.def.location, ref: p.def.location.ref.replace(/\d+$/, String(anchor.minRow + p.def.location.firstDataRow + lines)) },
      rowItems: Array.from({ length: lines }, () => ({ type: 'data' as const, r: 0, i: 0, x: [0] })),
    }
    const area = savedPivotLayout(def)!.rect
    expect((area.maxRow - area.minRow + 1) * (area.maxCol - area.minCol + 1) * PAGED_CELL_BYTES_ESTIMATE.cell).toBeGreaterThan(PAGED_CELL_LIMITS[profile])
    // A synthetic paged source: every cell of the requested range populated, 1,000 rows per page.
    const PAGE_ROWS = 1000, stamp = { docId: 'synthetic', sheetId: 's', valueGen: 0 }
    let served = 0, pages = 0
    const store: Pick<SourceStore, 'readRows'> = {
      async readRows(range, cursor) {
        pages++
        const r0 = cursor === undefined ? range.r0 : Number(cursor), r1 = Math.min(range.r1, r0 + PAGE_ROWS - 1)
        const items = []
        for (let row = r0; row <= r1; row++) {
          const cells = []
          for (let col = range.c0; col <= range.c1; col++) cells.push({ id: `${row}:${col}` as never, row, col, value: row * 10 + col, styleIndex: 0 })
          served += cells.length
          items.push({ row, cells })
        }
        return { items, completeness: 'complete', stamp, ...(r1 < range.r1 ? { nextCursor: String(r1 + 1) } : {}) }
      },
    }
    const reader = await pagedCellReader(store, def, { signal: new AbortController().signal, stamp }, profile)
    expect(reader.incomplete).toBe('over-cap')
    expect(reader(area.minRow + 5, area.minCol)).toBeUndefined()
    // Reading stopped at the cap: at most one page past the cells the cap admits.
    const capCells = Math.floor(PAGED_CELL_LIMITS[profile] / PAGED_CELL_BYTES_ESTIMATE.cell)
    const pageCells = PAGE_ROWS * (area.maxCol - area.minCol + 1)
    expect(served).toBeLessThanOrEqual(capCells + pageCells)
    expect(pages).toBeLessThan(lines / PAGE_ROWS)
    expect(check(p, reader)).toMatchObject({ status: 'unverified', reason: 'cells-over-cap' })
  })
})

// ---------------------------------------------------------------- corpus: writer files, status recorded

/** Each corpus pivot's status. Writer = workbook.xml `fileVersion@appName` ("xl" is Excel, "Calc" LibreOffice/Collabora).
 * LibreOffice and POI files are their writer's computation: their status is recorded here, never tuned toward. A
 * mismatch on an Excel-made file would be an engine bug (reported in the task report). */
const UNMEASURED = 'unverified:unsupported:pivot-engine-unmeasured'
const CORPUS_STATUS: Record<string, string[]> = {
  'lo-calc-Pivot2_Row_Compact.xlsx': ['xl', 'ok'],
  'lo-calc-Pivot2_Row_Grand_Subtotals.xlsx': ['xl', 'ok'],
  'lo-calc-Pivot3_Column_Grand_Subtotals.xlsx': ['xl', 'ok'],
  // Field 1 is sorted (Task 21).
  'lo-calc-Pivot4_Column_Grand_Subtotals_SortDescending.xlsx': ['xl', 'ok'],
  'lo-calc-PivotTableCellFormatsTest_7_TwoRowTwoColumnFields_DataColor.xlsx': ['xl', 'ok'],
  'lo-calc-PivotTableFormatsSubTotals.xlsx': ['xl', 'ok'],
  // worksheetSource r:id: the source is another workbook (§7.6).
  'lo-calc-PivotTableWithNoSourceData.xlsx': ['xl', 'refused:pivot-source-unsupported'],
  'lo-calc-PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx': ['xl', 'ok'],
  // Collabora-written, not an oracle. Calculated data fields (Task 22): Field4 `2/0` uses no field and Field5 refers to
  // a calculated field, both unmeasured (S2-M); Collabora shows #VALUE! for both.
  'lo-calc-calcfields.xlsx': ['Calc', UNMEASURED],
  // Count subtotals/grand totals from records are unmeasured (Task 15).
  'lo-calc-pivot_dark1.xlsx': ['xl', UNMEASURED],
  'lo-calc-pivotcompact.xlsx': ['xl', 'ok'],
  // Field 3 is sorted (Task 21).
  'lo-calc-pivottable_date_field_filter.xlsx': ['xl', UNMEASURED],
  'lo-calc-pivottable_outline_mode.xlsx': ['xl', 'ok'],
  // Field 6, on no axis, has a hidden item (Task 15 cross-field measurement).
  'lo-calc-pivottable_rowcolpage_field_filter.xlsx': ['xl', UNMEASURED],
  // Hidden items filter records: visible totals are unmeasured (Task 15).
  'lo-calc-pivottable_tabular_mode.xlsx': ['xl', UNMEASURED],
  // Date and number grouping (Task 21).
  'lo-calc-shared-dategroup.xlsx': ['xl', 'ok', 'ok'],
  'lo-calc-shared-numgroup.xlsx': ['xl', 'ok', 'ok'],
  // Count subtotals/grand totals from records are unmeasured (Task 15).
  'lo-calc-two-data-fields.xlsx': ['xl', UNMEASURED],
  // A page selection filters records: visible totals are unmeasured (Task 15).
  'poi-ExcelPivotTableSample.xlsx': ['xl', UNMEASURED, UNMEASURED],
}
const CORPUS_FILES = existsSync(CORPUS) ? readdirSync(CORPUS).filter(name => name.endsWith('.xlsx')).sort() : []
async function corpusStatus(file: string): Promise<{ writer: string; statuses: string[] }> {
  const bytes = new Uint8Array(readFileSync(resolve(CORPUS, file)))
  const zip = await JSZip.loadAsync(bytes)
  const workbook = await zip.file('xl/workbook.xml')?.async('string') ?? ''
  const writer = /appName="([^"]*)"/.exec(workbook)?.[1] ?? 'unknown'
  const statuses: string[] = []
  for (const parts of await loadPivotParts(file, bytes)) {
    if (!parts.cache) { statuses.push('no-cache'); continue }
    const eligible = engineEligibility(parts.cache, parts.def, { sourceReadable: false })
    if (!eligible.ok) { statuses.push(`refused:${eligible.diagnostic.feature}`); continue }
    const recordBytes = await parts.pkg.bytes(parts.cache.recordsPartPath!)
    const records = await decodeRecords(once(recordBytes!), parts.cache, RECORD_LIMITS.D, parts.date1904)
    if (!records.ok) { statuses.push(`records:${records.diagnostic.feature}`); continue }
    statuses.push(describeResult(checkSavedLayout(records.store, parts.cache, parts.def, sheetCellReader(parts.sheet), { date1904: parts.date1904 })))
  }
  return { writer, statuses }
}

describe.skipIf(!CORPUS_FILES.length)('corpus pivots', () => {
  it('LibreOffice files are ok, mismatch or unverified, never throw', async () => {
    const table: Record<string, string[]> = {}
    for (const file of CORPUS_FILES) {
      const zip = await JSZip.loadAsync(readFileSync(resolve(CORPUS, file)))
      if (!Object.keys(zip.files).some(path => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(path))) continue
      const { writer, statuses } = await corpusStatus(file)
      table[file] = [writer, ...statuses]
    }
    expect(table).toEqual(CORPUS_STATUS)
  })
})

