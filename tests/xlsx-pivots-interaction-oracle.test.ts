// @vitest-environment node
/**
 * Excel interaction oracle (plan Task 25 Step 1, spec §9, §11): a session opened on a seed, given the `PivotDelta`s
 * of the change Excel applied, must reproduce Excel's saved result: every saved cell of the case under
 * checkSavedLayout's rules (numbers to 15 significant digits, text exactly, blanks), and the live definition's
 * `rowItems`, `colItems` and `location`.
 *
 * Sources:
 * 1. Task 24's `interaction/manifest.json` (task-25 cases) and `interaction/delta-from-stage2.json`, when present.
 *    While absent, each shows a visible named skip, never a silent pass.
 * 2. Until then, deltas written by hand from scripts/xlsx-refs/engine-cases.ts for the Stage 2 cases that are Excel
 *    applying one Stage 3 UI change to the engine seed (they share the seed's cache and records exactly), run on
 *    engine/baseline.xlsx (the seed as Excel refreshed it). `ORACLE_VERIFIED_KINDS` may list only kinds with a
 *    passing case here.
 * 3. Cross-case evidence (two Excel saves of one pivot that differ by a UI change, run in reverse or between cases):
 *    pins semantics, never enables a kind.
 */
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import type { PivotFilter } from '../src/xlsx/pivots/model'
import { compareWithCase, describeOutcome, type Outcome } from './helpers/xlsx-pivots-oracle'
import { PIVOT_DELTA_KINDS, type PivotDelta, type PivotDeltaKind } from '../src/xlsx/pivots/session/delta'
import { openPivotSession, ORACLE_VERIFIED_KINDS } from '../src/xlsx/pivots/session/session'
import { F, thresholds } from '../scripts/xlsx-refs/engine-cases'
import { loadAll, type Loaded } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const INTERACTION = resolve(FIXTURES, 'interaction')
const MANIFEST = resolve(INTERACTION, 'manifest.json')
const FROM_STAGE2 = resolve(INTERACTION, 'delta-from-stage2.json')
/** Task 24's seeds (InteractionCase.seed). The engine seed's cache and records equal every same-cache Stage 2 case's. */
const SEED_FILES: Record<string, string> = {
  'engine-seed': resolve(FIXTURES, 'recipes/engine-seed.xlsx'),
  'chart-seed': resolve(FIXTURES, 'recipes/pivot-chart-seed.xlsx'),
  'getpivotdata-seed': resolve(INTERACTION, 'getpivotdata-seed.xlsx'),
  'refresh-seed': resolve(INTERACTION, 'refresh-seed.xlsx'),
}
const BASELINE = resolve(FIXTURES, 'engine/baseline.xlsx')
const engineCase = (name: string) => resolve(FIXTURES, `engine/${name}.xlsx`)

const cached = new Map<string, Promise<Loaded[]>>()
const pivotOf = async (file: string, name?: string) => {
  let p = cached.get(file)
  if (!p) cached.set(file, p = loadAll(file))
  const all = await p
  const found = name === undefined ? all[0] : all.find(entry => entry.def.name === name)
  if (!found) throw new Error(`${file}: no pivot ${name ?? ''}`)
  return found
}

const ALL_KINDS: ReadonlySet<PivotDeltaKind> = new Set(PIVOT_DELTA_KINDS)
/** Opens a session on `seed`, applies and commits `deltas` in order, and compares with `target`'s saved state. */
function replay(seed: Loaded, target: Loaded, deltas: readonly PivotDelta[]): Outcome {
  const savedSeed = sheetCellReader(seed.sheet)
  const opened = openPivotSession({
    def: seed.def, cache: seed.cache, store: seed.store, date1904: seed.date1904, saved: savedSeed,
    occupied: (row, col) => savedSeed(row, col) !== undefined, linkedCharts: [], verifiedKinds: ALL_KINDS,
  })
  if (!opened.ok) return { status: 'refused', feature: opened.diagnostic.feature, message: `open: ${opened.diagnostic.message}` }
  const session = opened.value
  for (const delta of deltas) {
    const applied = session.apply(delta)
    if (!applied.ok) return { status: 'refused', feature: applied.diagnostic.feature, message: applied.diagnostic.message }
    if (!session.commit(applied.value)) return { status: 'differs', reason: 'commit refused' }
  }
  return compareWithCase(session, target)
}

// ---- Hand-written deltas (the filters as the UI would build them, from engine-cases.ts `labelFilter` & co.) ----------
const T = thresholds()
let filterId = 100
const filterOf = (fld: number, type: PivotFilter['type'], column: PivotFilter['autoFilter']['columns'][number], extra: Partial<PivotFilter> = {}): PivotFilter =>
  ({ fld, type, evalOrder: -1, id: ++filterId, ...extra, autoFilter: { ref: 'A1', columns: [column] } })
const custom = (and: boolean, ...filters: Array<[NonNullable<PivotFilter['autoFilter']['columns'][number]['customFilters']>['filters'][number]['operator'], string]>) =>
  ({ colId: 0, hiddenButton: false, showButton: true, customFilters: { and, filters: filters.map(([operator, val]) => ({ operator, val })) } })
const label = (fld: number, type: 'captionBeginsWith' | 'captionEqual' | 'captionBetween', v1: string, v2?: string) =>
  filterOf(fld, type, type === 'captionBetween' ? custom(true, ['greaterThanOrEqual', v1], ['lessThanOrEqual', v2!]) : custom(false, ['equal', type === 'captionBeginsWith' ? `${v1}*` : v1]),
    v2 === undefined ? { stringValue1: v1 } : { stringValue1: v1, stringValue2: v2 })
const value = (fld: number, measure: number, type: 'valueGreaterThan' | 'valueBetween', v1: number, v2?: number) =>
  filterOf(fld, type, type === 'valueBetween' ? custom(true, ['greaterThanOrEqual', String(v1)], ['lessThanOrEqual', String(v2)]) : custom(false, ['greaterThan', String(v1)]), { iMeasureFld: measure })
const top = (fld: number, measure: number, type: 'count' | 'percent' | 'sum', val: number) =>
  filterOf(fld, type, { colId: 0, hiddenButton: false, showButton: true, top10: { top: true, percent: type === 'percent', val } }, { iMeasureFld: measure })
const filter = (f: PivotFilter): PivotDelta => ({ kind: 'filter', field: f.fld, filter: f })
const ELDERBERRIES = 3, WEST = 3, EAST = 0
const SALES = 0, UNITS = 1

/** Stage 2 cases that are one Stage 3 UI change to the engine seed (engine-cases.ts `lines`), as deltas. */
const FROM_BASELINE: ReadonlyArray<readonly [string, readonly PivotDelta[]]> = [
  ['layout-collapsed-item', [{ kind: 'collapse', field: F.Region, item: EAST, collapsed: true }]],
  ['quirk-hidden-row-item', [{ kind: 'items', field: F.Region, hidden: [WEST] }]],
  ['quirk-cross-hidden-item', [{ kind: 'items', field: F.Product, hidden: [ELDERBERRIES] }]],
  ['sort-text-ascending', [{ kind: 'sort', field: F.Product, order: 'ascending' }]],
  ['sort-text-descending', [{ kind: 'sort', field: F.Product, order: 'descending' }]],
  ['sort-number-descending', [{ kind: 'sort', field: F.Year, order: 'descending' }]],
  ['sort-value-top-level', [{ kind: 'sort', field: F.Region, order: 'descending', byData: SALES }]],
  ['sort-value-within-groups', [{ kind: 'sort', field: F.Product, order: 'descending', byData: SALES }]],
  ['sort-value-ties', [{ kind: 'sort', field: F.Region, order: 'descending', byData: UNITS }]],
  ['filter-label-begins-with', [filter(label(F.Product, 'captionBeginsWith', 'B'))]],
  ['filter-label-wildcard', [filter(label(F.Product, 'captionEqual', '*an?s'))]],
  ['filter-label-between', [filter(label(F.Product, 'captionBetween', 'B', 'D'))]],
  ['filter-label-cross-field', [filter(label(F.Product, 'captionBeginsWith', 'A'))]],
  ['filter-value-greater-than', [filter(value(F.Region, SALES, 'valueGreaterThan', T.regionValueGreater))]],
  ['filter-value-between', [filter(value(F.Region, SALES, 'valueBetween', T.regionValueLow, T.regionValueHigh))]],
  ['filter-value-nested', [filter(value(F.Product, SALES, 'valueGreaterThan', T.productValueGreater))]],
  ['filter-top-count', [filter(top(F.Region, SALES, 'count', 2))]],
  ['filter-top-count-ties', [filter(top(F.Region, UNITS, 'count', 2))]],
  ['filter-top-percent', [filter(top(F.Region, SALES, 'percent', 50))]],
  ['filter-top-percent-30', [filter(top(F.Region, SALES, 'percent', 30))]],
  ['filter-top-sum', [filter(top(F.Region, SALES, 'sum', T.regionTopSum))]],
  ['filter-top-nested', [filter(top(F.Product, SALES, 'count', 1))]],
]
/** Kinds with a passing seed-based case, collected as the cases run (the pin below reads it last). */
const passedKinds = new Set<PivotDeltaKind>()

describe('Excel interaction oracle', () => {
  describe('Task 24 manifests', () => {
    /** Ruling S3-D: `base` (relative to tests/fixtures/xlsx-pivots/, no extension) is the Excel-saved start file, else the
     * seed's; a `setup` case is a measurement or a base, not a delta oracle. Only built cases run. */
    interface InteractionCase { case: string; file: string; pivot: string; task: number; seed: string; delta: PivotDelta[]; base?: string; setup?: string; status?: string }
    if (!existsSync(MANIFEST)) {
      it.skip('SKIPPED: tests/fixtures/xlsx-pivots/interaction/manifest.json is absent (Task 24 interaction oracle not built); Task 24 cases not run', () => {})
    } else {
      const manifest = (raw => Array.isArray(raw) ? raw : raw.cases)(JSON.parse(readFileSync(MANIFEST, 'utf8'))) as InteractionCase[]
      const ours = manifest.filter(entry => entry.task === 25 || entry.task === 35 || entry.task === 36 || entry.task === 37)
      const unverifiedOf = (deltas: readonly PivotDelta[]) => [...new Set(deltas.map(d => d.kind).filter(kind => !ORACLE_VERIFIED_KINDS.has(kind)))]
      for (const entry of ours.filter(entry => entry.status !== 'built')) it.skip(`SKIPPED: manifest ${entry.case} is not built (status ${entry.status ?? 'none'})`, () => {})
      for (const entry of ours.filter(entry => entry.status === 'built' && entry.setup !== undefined)) it.skip(`SKIPPED: manifest ${entry.case} is a setup case, not a delta oracle (${entry.setup})`, () => {})
      for (const entry of ours.filter(entry => entry.status === 'built' && entry.setup === undefined && (unverifiedOf(entry.delta).length > 0 || (entry.task !== 25 && entry.delta.length === 0)))) {
        const waiting = unverifiedOf(entry.delta)
        it.skip(`SKIPPED: manifest ${entry.case} (task ${entry.task}) awaits Stage 4 implementation${waiting.length ? ` (${waiting.join(', ')})` : ''}`, () => {})
      }
      const cases = ours.filter(entry => entry.status === 'built' && entry.setup === undefined && unverifiedOf(entry.delta).length === 0 && (entry.task === 25 || entry.delta.length > 0))
      if (!cases.length) it.skip('SKIPPED: interaction/manifest.json has no built verified delta case', () => {})
      else it.each(cases.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
        const start = entry.base !== undefined ? resolve(FIXTURES, `${entry.base}.xlsx`) : SEED_FILES[entry.seed]
        expect(start && existsSync(start), `start file ${entry.base ?? entry.seed}`).toBe(true)
        const outcome = replay(await pivotOf(start!, entry.pivot), await pivotOf(resolve(process.cwd(), entry.file), entry.pivot), entry.delta)
        expect(describeOutcome(outcome), entry.case).toBe('ok')
        // Only a case that starts from the seed itself or an Excel save of it counts towards enabling its kinds.
        for (const delta of entry.delta) passedKinds.add(delta.kind)
      })
    }
    if (!existsSync(FROM_STAGE2)) {
      it.skip('SKIPPED: tests/fixtures/xlsx-pivots/interaction/delta-from-stage2.json is absent (Task 24 classification not built); hand-written deltas below run instead', () => {})
    } else {
      const allEntries = (JSON.parse(readFileSync(FROM_STAGE2, 'utf8')) as Array<{ case: string; delta: PivotDelta | PivotDelta[] | null }>).filter(entry => entry.delta !== null)
      const deltasOf = (entry: { delta: PivotDelta | PivotDelta[] | null }) => Array.isArray(entry.delta) ? entry.delta : [entry.delta!]
      const unverifiedOf = (deltas: readonly PivotDelta[]) => [...new Set(deltas.map(d => d.kind).filter(kind => !ORACLE_VERIFIED_KINDS.has(kind)))]
      for (const entry of allEntries.filter(e => unverifiedOf(deltasOf(e)).length > 0)) {
        it.skip(`SKIPPED: stage 2 ${entry.case} awaits Stage 4 implementation (${unverifiedOf(deltasOf(entry)).join(', ')})`, () => {})
      }
      const entries = allEntries.filter(e => unverifiedOf(deltasOf(e)).length === 0)
      if (!entries.length) it.skip('SKIPPED: delta-from-stage2.json has no verified entry with a delta', () => {})
      else it.each(entries.map(entry => [entry.case, entry] as const))('stage 2 %s', async (_name, entry) => {
        const deltas = deltasOf(entry)
        const outcome = replay(await pivotOf(SEED_FILES['engine-seed']!), await pivotOf(engineCase(entry.case)), deltas)
        expect(describeOutcome(outcome), entry.case).toBe('ok')
        for (const delta of deltas) passedKinds.add(delta.kind)
      })
    }
  })

  describe('hand-written deltas from the engine seed (baseline.xlsx)', () => {
    it.each(FROM_BASELINE.map(([name, deltas]) => [name, deltas] as const))('%s', async (name, deltas) => {
      const outcome = replay(await pivotOf(BASELINE), await pivotOf(engineCase(name)), deltas)
      expect(describeOutcome(outcome), name).toBe('ok')
      for (const delta of deltas) passedKinds.add(delta.kind)
    })
  })

  describe('cross-case evidence (semantics only; enables no kind)', () => {
    it('expanding the collapsed item gives the seed back', async () => {
      expect(describeOutcome(replay(await pivotOf(engineCase('layout-collapsed-item')), await pivotOf(BASELINE), [{ kind: 'collapse', field: F.Region, item: EAST, collapsed: false }]))).toBe('ok')
    })
    it('showing the hidden item gives the seed back', async () => {
      expect(describeOutcome(replay(await pivotOf(engineCase('quirk-hidden-row-item')), await pivotOf(BASELINE), [{ kind: 'items', field: F.Region, hidden: [] }]))).toBe('ok')
    })
    it('clearing a label or a value filter gives the seed back', async () => {
      expect(describeOutcome(replay(await pivotOf(engineCase('filter-label-begins-with')), await pivotOf(BASELINE), [{ kind: 'clearFilters', field: F.Product }]))).toBe('ok')
      expect(describeOutcome(replay(await pivotOf(engineCase('filter-value-greater-than')), await pivotOf(BASELINE), [{ kind: 'clearFilters', field: F.Region }]))).toBe('ok')
    })
    it('a page multi-selection narrowed to one item matches the single-selection save', async () => {
      // Both cases are the seed with Product moved to the page area; they differ only in the page selection.
      // Selecting Apples first: with every item shown under (All) the multi-selection save has no "(All)" caption.
      const deltas: PivotDelta[] = [{ kind: 'page', field: F.Product, selected: [0] }, { kind: 'items', field: F.Product, hidden: [] }]
      expect(describeOutcome(replay(await pivotOf(engineCase('quirk-page-multiple')), await pivotOf(engineCase('quirk-page-single')), deltas))).toBe('ok')
    })
    it('the reverse, as a multiple selection, matches the multi-selection save with the measured "(多重項目)"', async () => {
      const deltas: PivotDelta[] = [{ kind: 'page', field: F.Product, selected: [0, 2] }]
      expect(describeOutcome(replay(await pivotOf(engineCase('quirk-page-single')), await pivotOf(engineCase('quirk-page-multiple')), deltas))).toBe('ok')
    })
  })

  it('the comparison sees a missing or a wrong change', async () => {
    const seed = await pivotOf(BASELINE), target = await pivotOf(engineCase('quirk-hidden-row-item'))
    expect(replay(seed, target, []).status).toBe('differs')
    expect(replay(seed, target, [{ kind: 'items', field: F.Region, hidden: [EAST] }]).status).toBe('differs')
    // Same lines and location, one value off: the cells are compared too.
    const sorted = await pivotOf(engineCase('sort-value-top-level'))
    const edited = { ...sorted, sheet: { ...sorted.sheet, rows: sorted.sheet.rows.map(row => row.index !== 5 ? row : { ...row, cells: row.cells.map(cell => cell.col === 1 ? { ...cell, value: 1 } : cell) }) } }
    expect(replay(seed, edited, [{ kind: 'sort', field: F.Region, order: 'descending', byData: SALES }]).status).toBe('differs')
    // Same cells, lines and location, but the innermost field collapsed too, which Excel leaves alone: the item flags differ.
    const all = await pivotOf(resolve(INTERACTION, 'interaction-collapse-all-compact-top.xlsx'))
    expect(describeOutcome(replay(seed, all, [{ kind: 'collapseAll', collapsed: true }]))).toBe('ok')
    const extra = replay(seed, all, [{ kind: 'collapseAll', collapsed: true }, { kind: 'collapseField', field: F.Product, collapsed: true }])
    expect(extra.status === 'differs' && extra.reason.startsWith('item flags')).toBe(true)
  })

  it('ORACLE_VERIFIED_KINDS lists only kinds with a passing seed-based case', () => {
    for (const kind of ORACLE_VERIFIED_KINDS) expect(passedKinds.has(kind), `${kind} has no passing oracle case`).toBe(true)
  })
})
