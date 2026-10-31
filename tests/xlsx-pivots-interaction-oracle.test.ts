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
import { savedPivotLayout } from '../src/xlsx/pivots/savedLayout'
import { sheetCellReader, type SavedCell } from '../src/xlsx/pivots/engine/faithful'
import type { PivotLayout, PivotLayoutCell } from '../src/xlsx/pivots/engine/place'
import { PIVOT_VALUES_FIELD, type PivotFilter, type PivotTableDefinition } from '../src/xlsx/pivots/model'
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
const same15 = (a: number, b: number) => a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15))
/** checkSavedLayout's comparison (faithful.ts `differs`), with implicit captions compared for presence. */
function differs(cell: PivotLayoutCell | undefined, saved: SavedCell | undefined): string | undefined {
  const blank = !cell || cell.value.kind === 'empty' || (cell.value.kind === 'text' && cell.value.value === '')
  if (blank) return saved === undefined ? undefined : `saved ${JSON.stringify(saved)}, computed blank`
  const value = cell!.value
  if (saved === undefined) return `computed ${JSON.stringify(value)}, saved blank`
  if (cell!.caption === 'implicit') return saved.kind === 'text' ? undefined : 'implicit caption over a non-text cell'
  if (value.kind === 'number') return saved.kind === 'number' && same15(value.value, saved.value) ? undefined : `computed ${value.value}, saved ${JSON.stringify(saved)}`
  if (value.kind === 'text' || value.kind === 'error') return saved.kind === value.kind && saved.value === value.value ? undefined : `computed ${JSON.stringify(value)}, saved ${JSON.stringify(saved)}`
  return 'the value differs'
}
const at = (row: number, col: number) => `R${row + 1}C${col + 1}`

type Outcome = { status: 'ok' } | { status: 'refused'; feature: string; message: string } | { status: 'differs'; reason: string }
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
  const live = session.def
  for (const axis of ['rowItems', 'colItems'] as const) {
    const ours = lines(live, axis), theirs = lines(target.def, axis)
    if (ours !== theirs) return { status: 'differs', reason: `${axis}: live ${ours}, saved ${theirs}` }
  }
  if (JSON.stringify(live.location) !== JSON.stringify(target.def.location)) return { status: 'differs', reason: `location: live ${JSON.stringify(live.location)}, saved ${JSON.stringify(target.def.location)}` }
  return compareCells(session.layout, target)
}
/**
 * A definition's `rowItems`/`colItems` with each item index read through its field's `items` list (its cache item).
 * When Excel sorts a field it saves that field's `<items>` in the new order and indexes the lines into it; the live
 * definition keeps the saved list and indexes the lines into that. Both name the same items in the same order.
 */
function lines(def: PivotTableDefinition, axis: 'rowItems' | 'colItems'): string {
  const fields = axis === 'rowItems' ? def.rowFields : def.colFields
  return JSON.stringify((def[axis] ?? []).map(line => ({
    type: line.type, r: line.r, i: line.i,
    x: line.type === 'grand' ? line.x : line.x.map((x, k) => {
      const field = fields[line.r + k]
      return field === undefined || field === PIVOT_VALUES_FIELD ? `d${x}` : def.fields[field]?.items?.[x]?.x ?? x
    }),
  })))
}
function compareCells(layout: PivotLayout, target: Loaded): Outcome {
  const saved = sheetCellReader(target.sheet)
  const rects = savedPivotLayout(target.def)?.rects
  if (!rects) return { status: 'differs', reason: 'the case has no saved layout' }
  const cells = new Map(layout.cells.map(cell => [`${cell.row},${cell.col}`, cell]))
  const seen = new Set<string>()
  for (const rect of rects) for (let row = rect.minRow; row <= rect.maxRow; row++) for (let col = rect.minCol; col <= rect.maxCol; col++) {
    const key = `${row},${col}`
    if (seen.has(key)) continue
    seen.add(key)
    const reason = differs(cells.get(key), saved(row, col))
    if (reason) return { status: 'differs', reason: `${at(row, col)}: ${reason}` }
  }
  for (const cell of layout.cells) if (!seen.has(`${cell.row},${cell.col}`) && differs(cell, undefined)) return { status: 'differs', reason: `${at(cell.row, cell.col)} is outside the saved pivot area` }
  return { status: 'ok' }
}
const describeOutcome = (outcome: Outcome) => outcome.status === 'ok' ? 'ok' : outcome.status === 'refused' ? `refused ${outcome.feature}: ${outcome.message}` : `differs: ${outcome.reason}`

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
    interface InteractionCase { case: string; file: string; pivot: string; task: number; seed: string; delta: PivotDelta[] }
    if (!existsSync(MANIFEST)) {
      it.skip('SKIPPED: tests/fixtures/xlsx-pivots/interaction/manifest.json is absent (Task 24 interaction oracle not built); Task 24 cases not run', () => {})
    } else {
      const manifest = (raw => Array.isArray(raw) ? raw : raw.cases)(JSON.parse(readFileSync(MANIFEST, 'utf8'))) as InteractionCase[]
      const cases = manifest.filter(entry => entry.task === 25)
      if (!cases.length) it.skip('SKIPPED: interaction/manifest.json has no task-25 case', () => {})
      else it.each(cases.map(entry => [entry.case, entry] as const))('manifest %s', async (_name, entry) => {
        const seedFile = SEED_FILES[entry.seed]
        expect(seedFile && existsSync(seedFile), `seed ${entry.seed}`).toBe(true)
        const outcome = replay(await pivotOf(seedFile!), await pivotOf(resolve(process.cwd(), entry.file), entry.pivot), entry.delta)
        expect(describeOutcome(outcome), entry.case).toBe('ok')
        for (const delta of entry.delta) passedKinds.add(delta.kind)
      })
    }
    if (!existsSync(FROM_STAGE2)) {
      it.skip('SKIPPED: tests/fixtures/xlsx-pivots/interaction/delta-from-stage2.json is absent (Task 24 classification not built); hand-written deltas below run instead', () => {})
    } else {
      const entries = (JSON.parse(readFileSync(FROM_STAGE2, 'utf8')) as Array<{ case: string; delta: PivotDelta | PivotDelta[] | null }>).filter(entry => entry.delta !== null)
      if (!entries.length) it.skip('SKIPPED: delta-from-stage2.json has no entry with a delta', () => {})
      else it.each(entries.map(entry => [entry.case, entry] as const))('stage 2 %s', async (_name, entry) => {
        const deltas = Array.isArray(entry.delta) ? entry.delta : [entry.delta!]
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
    it('the reverse needs "(Multiple Items)", which the single-selection save never shows: pivot-caption-locale', async () => {
      const deltas: PivotDelta[] = [{ kind: 'items', field: F.Product, hidden: [1, 3] }, { kind: 'page', field: F.Product, selected: 'all' }]
      const outcome = replay(await pivotOf(engineCase('quirk-page-single')), await pivotOf(engineCase('quirk-page-multiple')), deltas)
      expect(outcome.status === 'refused' ? outcome.feature : describeOutcome(outcome)).toBe('pivot-caption-locale')
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
  })

  it('ORACLE_VERIFIED_KINDS lists only kinds with a passing seed-based case', () => {
    for (const kind of ORACLE_VERIFIED_KINDS) expect(passedKinds.has(kind), `${kind} has no passing oracle case`).toBe(true)
  })
})
