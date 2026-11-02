import { afterEach, describe, expect, it, vi } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { ENGINE_CASES, engineSeedStandIn, xmlDiffSummary, type PivotParts } from '../scripts/xlsx-refs/engine-cases'
import {
  INTERACTION_CASES, INTERACTION_PROBES, SEED_BUILDS, SEED_FILES, resolveDelta, startOf, startShapeOf, type InteractionCase, type PivotDelta,
} from '../scripts/xlsx-refs/interaction-cases'
import {
  CONTROL_INTERACTION_CASES, CONTROL_SEED_HEADERS, CONTROL_SEED_ROWS,
  type ControlInteractionCase, type ControlOracleOp,
} from '../scripts/xlsx-refs/control-interaction-cases'
import { classify } from '../scripts/xlsx-refs/classify-stage2'
import { interactionDryRun, verifyInteraction, verifyXmlRecomputation } from '../scripts/xlsx-refs/build-engine'

const PIVOTS = join(__dirname, 'fixtures/xlsx-pivots')
const fixture = (rel: string) => join(PIVOTS, `${rel}.xlsx`)
const CLASSIFICATION = join(PIVOTS, 'interaction/delta-from-stage2.json')
const ENGINE_MANIFEST = join(PIVOTS, 'engine/manifest.json')

async function zipOf(file: string) { return JSZip.loadAsync(readFileSync(file)) }
const firstPart = (zip: JSZip, re: RegExp) => Object.keys(zip.files).filter(n => re.test(n)).sort()[0]
async function parts(file: string): Promise<PivotParts> {
  const zip = await zipOf(file)
  return {
    pivot: await zip.file(firstPart(zip, /^xl\/pivotTables\/pivotTable\d+\.xml$/)!)!.async('string'),
    cache: await zip.file(firstPart(zip, /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/)!)!.async('string'),
  }
}
/** The engine/ cases whose cache fields and records equal baseline's: computed here independently of the script. */
async function sameCacheCases(): Promise<string[]> {
  const key = async (file: string) => {
    const zip = await zipOf(file)
    const def = await zip.file(firstPart(zip, /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/)!)!.async('string')
    const recordsPart = firstPart(zip, /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/)
    const records = recordsPart ? await zip.file(recordsPart)!.async('string') : ''
    return `${def.slice(def.indexOf('<cacheFields'), def.indexOf('</cacheFields>'))}|${records.slice(records.indexOf('<r>'))}`
  }
  const manifest = JSON.parse(readFileSync(ENGINE_MANIFEST, 'utf8')) as Array<{ case: string; status: string }>
  const base = await key(fixture('engine/baseline'))
  const out: string[] = []
  for (const e of manifest) if (e.status === 'built' && existsSync(fixture(`engine/${e.case}`)) && await key(fixture(`engine/${e.case}`)) === base) out.push(e.case)
  return out
}

const DATE_PERIODS = new Set(['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'])
const DATA_FNS = new Set(['average', 'count', 'countNums', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp'])
const SHOW_AS = new Set([
  'normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index',
  'percentOfParent', 'percentOfParentRow', 'percentOfParentCol', 'percentOfRunningTotal', 'rankAscending', 'rankDescending',
])

/** A structural check against PivotDelta; the Record over `kind` fails to compile when a kind is added. */
const KIND_CHECKS: Record<PivotDelta['kind'], (d: Record<string, unknown>) => boolean> = {
  collapse: d => Number.isInteger(d.field) && Number.isInteger(d.item) && typeof d.collapsed === 'boolean',
  collapseField: d => Number.isInteger(d.field) && typeof d.collapsed === 'boolean',
  collapseAll: d => typeof d.collapsed === 'boolean',
  items: d => Number.isInteger(d.field) && Array.isArray(d.hidden) && (d.hidden as unknown[]).every(Number.isInteger),
  page: d => Number.isInteger(d.field) && (d.selected === 'all' || (Array.isArray(d.selected) && (d.selected as unknown[]).every(Number.isInteger))),
  filter: d => {
    const f = d.filter as Record<string, unknown> | undefined
    return Number.isInteger(d.field) && !!f && f.fld === d.field && typeof f.type === 'string' && Number.isInteger(f.evalOrder) && Number.isInteger(f.id) &&
      Array.isArray((f.autoFilter as { columns?: unknown })?.columns)
  },
  clearFilters: d => Number.isInteger(d.field),
  sort: d => Number.isInteger(d.field) && (d.order === 'ascending' || d.order === 'descending') && (d.byData === undefined || Number.isInteger(d.byData)),
  moveField: d => Number.isInteger(d.field) && ['row', 'col', 'page', 'none'].includes(String(d.axis)) && (d.index === undefined || (Number.isInteger(d.index) && (d.index as number) >= 0)),
  addDataField: d => Number.isInteger(d.field) && (d.field as number) >= 0 && (d.index === undefined || (Number.isInteger(d.index) && (d.index as number) >= 0)),
  removeDataField: d => Number.isInteger(d.dataIndex) && (d.dataIndex as number) >= 0,
  moveDataField: d => Number.isInteger(d.dataIndex) && (d.dataIndex as number) >= 0 && Number.isInteger(d.toIndex) && (d.toIndex as number) >= 0,
  valueSettings: d => Number.isInteger(d.dataIndex) && (d.dataIndex as number) >= 0
    && (d.fn === undefined || DATA_FNS.has(String(d.fn)))
    && (d.showAs === undefined || SHOW_AS.has(String(d.showAs)))
    && (d.baseField === undefined || Number.isInteger(d.baseField))
    && (d.baseItem === undefined || Number.isInteger(d.baseItem))
    && (d.name === undefined || typeof d.name === 'string')
    && (d.numFmtId === undefined || d.numFmtId === null || Number.isInteger(d.numFmtId)),
  groupDate: d => Number.isInteger(d.field) && (d.field as number) >= 0
    && Array.isArray(d.periods) && d.periods.length > 0 && (d.periods as unknown[]).every(p => DATE_PERIODS.has(String(p)))
    && (d.startDate === undefined || typeof d.startDate === 'number')
    && (d.endDate === undefined || typeof d.endDate === 'number')
    && (d.daysInterval === undefined || Number.isInteger(d.daysInterval)),
  groupNumber: d => Number.isInteger(d.field) && (d.field as number) >= 0
    && typeof d.by === 'number' && d.by > 0
    && (d.start === undefined || typeof d.start === 'number')
    && (d.end === undefined || typeof d.end === 'number'),
  ungroup: d => Number.isInteger(d.field) && (d.field as number) >= 0,
}
function expectDelta(value: unknown): asserts value is PivotDelta {
  const d = value as Record<string, unknown>
  const check = KIND_CHECKS[d.kind as PivotDelta['kind']]
  expect(check, `unknown delta kind ${String(d.kind)}`).toBeTypeOf('function')
  expect(check!(d), JSON.stringify(d)).toBe(true)
}
const roundTrips = (v: unknown) => expect(JSON.parse(JSON.stringify(v))).toEqual(v)

describe('Stage 2 cases classified as interaction oracles (delta-from-stage2.json)', async () => {
  const classified = JSON.parse(readFileSync(CLASSIFICATION, 'utf8')) as Array<{ case: string; delta: unknown[] | null; reason?: string; note?: string }>
  const same = await sameCacheCases()
  it('lists every same-cache case exactly once, and nothing else', () => {
    expect(same.length).toBeGreaterThan(80)
    const ids = classified.map(c => c.case)
    expect(new Set(ids).size).toBe(ids.length)
    expect([...ids].sort()).toEqual([...same].sort())
  })
  it('gives each entry a delta list or a reason', () => {
    for (const c of classified) {
      if (c.delta === null) expect(c.reason, c.case).toBeTruthy()
      else for (const d of c.delta) expectDelta(d)
      roundTrips(c)
    }
  })
  it('indexes items by baseline\'s pivot-item lists', async () => {
    const shape = startShapeOf(await parts(fixture('engine/baseline')))
    for (const c of classified) for (const d of (c.delta ?? []) as PivotDelta[]) {
      const items = 'field' in d && d.field >= 0 ? shape.items[d.field] ?? shape.shared[d.field] ?? [] : []
      if (d.kind === 'collapse') expect(items[d.item], c.case).toBeTruthy()
      // An off-axis inline field has 0 items in baseline until an earlier moveField in c.delta materialises it.
      if (d.kind === 'items' && items.length > 0) for (const k of d.hidden) expect(items[k], c.case).toBeTruthy()
      if (d.kind === 'page' && d.selected !== 'all' && items.length > 0) for (const k of d.selected) expect(items[k], c.case).toBeTruthy()
    }
  })
  it('classifies Stage 3 and Stage 4 UI actions and keeps S2-AA variance-with-totals cases null', () => {
    const entry = (id: string) => classified.find(c => c.case === id)
    const delta = (id: string) => entry(id)?.delta
    expect(delta('baseline')).toEqual([])
    expect(delta('layout-collapsed-item')).toEqual([{ kind: 'collapse', field: 0, item: 0, collapsed: true }])
    expect(delta('quirk-hidden-row-item')).toEqual([{ kind: 'items', field: 0, hidden: [3] }])
    expect(delta('sort-value-ties')).toEqual([{ kind: 'sort', field: 0, order: 'descending', byData: 1 }])
    expect(delta('filter-top-count')).toMatchObject([{ kind: 'filter', field: 0, filter: { type: 'count', fld: 0 } }])
    // Stage 4 reclassified field moves, summary functions and show-values-as cases (all 36 pinned):
    expect(delta('layout-compact-three-rows')).toEqual([{ kind: 'moveField', field: 2, axis: 'row', index: 2 }])
    expect(delta('layout-values-on-rows-two-data')).toEqual([{ kind: 'moveField', field: -2, axis: 'row', index: 2 }])
    expect(delta('layout-collapsed-item-above-values')).toEqual([
      { kind: 'moveField', field: -2, axis: 'row', index: 2 },
      { kind: 'collapse', field: 0, item: 0, collapsed: true },
    ])
    expect(delta('quirk-page-single')).toEqual([
      { kind: 'moveField', field: 1, axis: 'page', index: 0 },
      { kind: 'page', field: 1, selected: [0] },
    ])
    expect(delta('quirk-page-multiple')).toEqual([
      { kind: 'moveField', field: 1, axis: 'page', index: 0 },
      { kind: 'page', field: 1, selected: [0, 2] },
    ])
    expect(delta('quirk-off-axis-hidden-item')).toEqual([
      { kind: 'moveField', field: 6, axis: 'row', index: 2 },
      { kind: 'items', field: 6, hidden: [1] },
      { kind: 'moveField', field: 6, axis: 'none' },
    ])
    for (const [id, fn] of [
      ['values-fn-sum', 'sum'], ['values-fn-count', 'count'], ['values-fn-average', 'average'],
      ['values-fn-max', 'max'], ['values-fn-min', 'min'], ['values-fn-product', 'product'], ['values-fn-countnums', 'countNums'],
    ] as const) {
      expect(delta(id), id).toEqual([{ kind: 'valueSettings', dataIndex: 0, fn }])
    }
    expect(delta('values-fn-countnums-mixed')).toEqual([
      { kind: 'addDataField', field: 6, index: 2 },
      { kind: 'valueSettings', dataIndex: 2, fn: 'countNums' },
    ])
    expect(delta('values-average-hidden-item')).toEqual([
      { kind: 'items', field: 0, hidden: [3] },
      { kind: 'valueSettings', dataIndex: 0, fn: 'average' },
    ])
    expect(delta('values-fn-max-of-date')).toEqual([
      { kind: 'addDataField', field: 5, index: 2 },
      { kind: 'valueSettings', dataIndex: 2, fn: 'max' },
    ])
    expect(delta('values-addition-order')).toEqual([
      { kind: 'addDataField', field: 7, index: 2 },
      { kind: 'valueSettings', dataIndex: 2, name: 'Sum of Net Amount' },
    ])
    expect(delta('values-showas-difference-previous')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 2, baseItem: 1048828 }])
    expect(delta('values-showas-difference-next')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 2, baseItem: 1048829 }])
    expect(delta('values-showas-difference-item')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 0, baseItem: 1 }])
    expect(delta('values-showas-percent-of-item')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percent', baseField: 0, baseItem: 1 }])
    expect(delta('values-showas-percent-diff-zero-base')).toEqual([{ kind: 'valueSettings', dataIndex: 1, showAs: 'percentDiff', baseField: 2, baseItem: 0 }])
    expect(delta('values-showas-percent-of-row')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfRow' }])
    expect(delta('values-showas-percent-of-column')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfCol' }])
    expect(delta('values-showas-percent-of-total')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfTotal' }])
    expect(delta('values-showas-percent-of-row-zero-total')).toEqual([{ kind: 'valueSettings', dataIndex: 1, showAs: 'percentOfRow' }])
    expect(delta('values-showas-index')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'index' }])
    expect(delta('values-showas-index-zero-total')).toEqual([{ kind: 'valueSettings', dataIndex: 1, showAs: 'index' }])
    expect(delta('values-showas-running-total-year')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'runTotal', baseField: 2 }])
    expect(delta('values-showas-running-total-region')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'runTotal', baseField: 0 }])
    expect(delta('values-showas-percent-of-parent-row-two-levels')).toEqual([
      { kind: 'moveField', field: 2, axis: 'row', index: 2 },
      { kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfParentRow' },
    ])
    expect(delta('values-showas-percent-of-parent-column')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfParentCol' }])
    expect(delta('values-showas-percent-of-parent-region')).toEqual([
      { kind: 'moveField', field: 2, axis: 'row', index: 2 },
      { kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfParent', baseField: 0 },
    ])
    expect(delta('values-showas-percent-running-total')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'percentOfRunningTotal', baseField: 2 }])
    expect(delta('values-showas-rank-descending-ties')).toEqual([{ kind: 'valueSettings', dataIndex: 1, showAs: 'rankDescending', baseField: 0 }])
    expect(delta('values-showas-rank-ascending')).toEqual([{ kind: 'valueSettings', dataIndex: 0, showAs: 'rankAscending', baseField: 1 }])
    // Explicit S2-AA variance-with-totals exclusions:
    for (const id of ['values-fn-stdev', 'values-fn-stdevp', 'values-fn-var', 'values-fn-varp']) {
      expect(delta(id), id).toBeNull()
      expect(entry(id)?.reason, id).toBe('refused by S2-AA variance-total 15th-digit rule when totals are shown; covered on total-less base in Stage 4 manifest')
    }
    // Non-UI options remain null:
    for (const id of ['layout-outline-three-rows', 'layout-grand-totals-off', 'values-showas-on-named-subtotals', 'filter-label-renamed-item', 'sort-value-showall-no-data'])
      expect(delta(id), id).toBeNull()
  })
  it('is what classify-stage2.ts writes', async () => {
    expect(await classify()).toEqual(classified)
  }, 120_000)
})

describe('interaction case manifest', () => {
  const ids = INTERACTION_CASES.map(c => c.id)
  const engineBuilt = new Set((JSON.parse(readFileSync(ENGINE_MANIFEST, 'utf8')) as Array<{ case: string; status: string }>).filter(e => e.status === 'built').map(e => e.case))
  it('has unique ids, a buildable route and a proof for every case', () => {
    expect(new Set(ids).size).toBe(ids.length)
    for (const c of INTERACTION_CASES) {
      expect([25, 29, 30, 31, 32, 35, 36, 37], c.id).toContain(c.task)
      expect(Object.keys(SEED_FILES), c.id).toContain(c.seed)
      if (c.route === 'applescript') expect(c.lines, c.id).toBeTypeOf('function')
      if (c.route === 'xml' || c.fallback === 'xml') expect(c.xmlEdit, c.id).toBeTypeOf('function')
      if (c.route === 'recipe') expect(c.recipe, c.id).toBeTruthy()
      const proofs = c.markers.length + (c.cacheMarkers?.length ?? 0) + (c.absent?.length ?? 0) + (c.cacheAbsent?.length ?? 0)
        + (c.recordMarkers?.length ?? 0) + (c.sheetMarkers?.length ?? 0)
      expect(proofs, `${c.id} has no marker`).toBeGreaterThan(0)
      expect(c.measures, c.id).toBeTruthy()
      if (c.probe) expect(INTERACTION_PROBES[c.probe], c.id).toBeTruthy()
    }
    for (const s of SEED_BUILDS) expect(INTERACTION_PROBES[s.probe], s.id).toBeTruthy()
  })
  it('starts each case from a file that exists or is built before it', () => {
    const seedFiles = new Set([...Object.values(SEED_FILES), ...SEED_BUILDS.map(s => s.file)])
    for (const [i, c] of INTERACTION_CASES.entries()) {
      if (!c.base) continue
      const [dir, name] = c.base.split('/') as [string, string]
      if (dir === 'engine') expect(engineBuilt.has(name), `${c.id}: base ${c.base}`).toBe(true)
      else if (dir === 'interaction') expect(ids.indexOf(name), `${c.id}: base ${c.base} is not an earlier case`).toBeGreaterThanOrEqual(0)
      else expect(seedFiles.has(c.base), `${c.id}: base ${c.base}`).toBe(true)
      if (dir === 'interaction') expect(ids.indexOf(name), `${c.id}: base ${c.base} must come first`).toBeLessThan(i)
    }
  })
  it('covers every interaction kind of the Stage 3 and Stage 4 briefs', () => {
    const kinds = new Set(INTERACTION_CASES.flatMap(c => c.delta.map(d => d.kind)))
    expect([...kinds].sort()).toEqual([
      'addDataField', 'clearFilters', 'collapse', 'collapseAll', 'collapseField', 'filter', 'groupDate', 'groupNumber',
      'items', 'moveDataField', 'moveField', 'page', 'removeDataField', 'sort', 'ungroup', 'valueSettings',
    ])
    for (const id of [
      'interaction-filter-all-hidden', 'interaction-page-multi-to-all', 'interaction-clear-label-filter', 'interaction-clear-value-filter',
      'caption-defaults', 'chart-baseline', 'chart-hide-item', 'chart-collapse-item', 'chart-value-filter', 'chart-sort', 'chart-values-on-rows', 'chart-page-select',
      'gpd-hide-item', 'gpd-collapse-item', 'gpd-filter-out', 'refresh-on-load', 'refresh-on-load-nodata',
      // Stage 4 Task 35:
      's4-move-inline-text-to-row', 's4-move-inline-date-to-col', 's4-move-inline-to-page', 's4-move-reorder-rows',
      's4-move-values-to-rows-first', 's4-move-row-to-col', 's4-move-row-to-col-and-back', 's4-move-filtered-field-to-page',
      's4-move-remove-filtered-field', 's4-data-remove-to-single', 's4-data-add-second-default-axis',
      's4-data-add-second-after-values-on-rows', 's4-data-add-numeric-vs-text', 's4-data-add-duplicate-field',
      's4-data-reorder', 's4-data-reorder-remap-measure', 's4-data-remove-all', 's4-empty-pivot-all-removed',
      's4-chart-values-to-rows', 's4-chart-remove-second-data',
      // Stage 4 Task 36:
      's4-values-fn-stdev-no-totals', 's4-values-fn-stdevp-no-totals', 's4-values-fn-var-no-totals', 's4-values-fn-varp-no-totals',
      's4-values-custom-name', 's4-values-numfmt-currency-percent', 's4-values-showas-previous-next', 's4-gpd-rename-and-move',
      // Stage 4 Task 37:
      's4-group-date-single-months', 's4-group-date-days-single', 's4-group-date-days-7', 's4-group-date-time-hms',
      's4-group-date-multi-yqm', 's4-group-date-manual-bounds', 's4-regroup-date-yqm-to-ym', 's4-ungroup-date-multi',
      's4-ungroup-date-single', 's4-group-number-units-auto', 's4-group-number-sales-blanks-manual',
      's4-regroup-number-units', 's4-ungroup-number-units', 's4-ungroup-manual-discrete',
    ])
      expect(ids, id).toContain(id)
    for (const f of ['compact-top', 'compact-bottom', 'outline-top', 'outline-bottom', 'tabular-top', 'tabular-bottom'])
      for (const what of ['item', 'field']) expect(ids).toContain(`interaction-collapse-${what}-${f}`)
    expect(new Set(INTERACTION_CASES.map(c => c.task))).toEqual(new Set([25, 29, 30, 31, 32, 35, 36, 37]))
  })
  it('starts every setup case from its seed (a setup on a base would blur which change made the start state)', () => {
    for (const c of INTERACTION_CASES.filter(x => x.setup)) expect(c.base, `${c.id}: a setup case starts from its seed`).toBeUndefined()
  })
})

/** A case's start parts: the start file when it exists, else (for an unbuilt interaction base or seed) undefined. */
async function startParts(c: InteractionCase): Promise<PivotParts | undefined> {
  const file = fixture(startOf(c))
  return existsSync(file) ? parts(file) : undefined
}
const standIn = async () => existsSync(fixture('recipes/engine-seed')) ? parts(fixture('recipes/engine-seed')) : engineSeedStandIn(await parts(fixture('recipes/pivot-seed')))
/** The start parts, or for a base not built yet: its own start with its xml edit applied (when it has one), down to
 * the engine seed. `exact` is false for anything synthesised. */
async function synthStart(c: InteractionCase): Promise<{ parts: PivotParts; exact: boolean }> {
  const built = await startParts(c)
  if (built) return { parts: built, exact: true }
  const parent = c.base?.startsWith('interaction/') ? INTERACTION_CASES.find(x => `interaction/${x.id}` === c.base) : undefined
  if (!parent) return { parts: await standIn(), exact: false }
  const above = await synthStart(parent)
  return { parts: parent.xmlEdit ? parent.xmlEdit({ ...above.parts }) : above.parts, exact: false }
}

describe('interaction deltas', async () => {
  const resolvable: Array<readonly [string, InteractionCase, PivotParts]> = []
  for (const c of INTERACTION_CASES) { const p = await startParts(c); if (p) resolvable.push([c.id, c, p] as const) }
  it('resolves against most cases\' start files already (the rest once their bases are built)', () => {
    expect(resolvable.length).toBeGreaterThan(INTERACTION_CASES.length / 2)
  })
  it.each(resolvable)('%s: resolves to PivotDelta and round-trips JSON', (_, c, p) => {
    const delta = c.delta.map(spec => resolveDelta(spec, startShapeOf(p)))
    for (const d of delta) expectDelta(d)
    roundTrips(delta)
  })
  it('resolves chart-seed items by its own item lists, not shared indices', async () => {
    const shape = startShapeOf(await parts(fixture(SEED_FILES['chart-seed'])))
    expect(resolveDelta({ kind: 'items', field: 'Region', hidden: ['West'] }, shape)).toEqual({ kind: 'items', field: 0, hidden: [3] })
    expect(resolveDelta({ kind: 'collapse', field: 'Region', item: 'East', collapsed: true }, shape)).toEqual({ kind: 'collapse', field: 0, item: 0, collapsed: true })
    expect(() => resolveDelta({ kind: 'items', field: 'Region', hidden: ['Central'] }, shape)).toThrow(/Central/)
  })
  it('resolves Stage 4 field moves, value settings sentinels and grouping specs', async () => {
    const shape = startShapeOf(await parts(fixture('engine/baseline')))
    expect(resolveDelta({ kind: 'moveField', field: 'Values', axis: 'row', index: 0 }, shape)).toEqual({ kind: 'moveField', field: -2, axis: 'row', index: 0 })
    expect(resolveDelta({ kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 'Year', baseItem: '(previous)' }, shape)).toEqual({
      kind: 'valueSettings', dataIndex: 0, showAs: 'difference', baseField: 2, baseItem: 1048828,
    })
    expect(resolveDelta({ kind: 'valueSettings', dataIndex: 1, showAs: 'percentDiff', baseField: 'Year', baseItem: '(next)' }, shape)).toEqual({
      kind: 'valueSettings', dataIndex: 1, showAs: 'percentDiff', baseField: 2, baseItem: 1048829,
    })
    expect(resolveDelta({ kind: 'valueSettings', dataIndex: 0, showAs: 'percent', baseField: 'Region', baseItem: 'North' }, shape)).toEqual({
      kind: 'valueSettings', dataIndex: 0, showAs: 'percent', baseField: 0, baseItem: 1,
    })
    expect(resolveDelta({ kind: 'groupDate', field: 'Date', periods: ['months', 'years'], startDate: 43831, endDate: 45291 }, shape)).toEqual({
      kind: 'groupDate', field: 5, periods: ['months', 'years'], startDate: 43831, endDate: 45291,
    })
    expect(resolveDelta({ kind: 'groupNumber', field: 'Sales', start: 0, end: 5000, by: 1000 }, shape)).toEqual({
      kind: 'groupNumber', field: 3, start: 0, end: 5000, by: 1000,
    })
    expect(resolveDelta({ kind: 'ungroup', field: 'Date' }, shape)).toEqual({ kind: 'ungroup', field: 5 })
  })
})

function parse(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  const error = doc.getElementsByTagName('parsererror')[0]
  if (error) throw new Error(`not well-formed: ${error.textContent}`)
  return doc
}
function expectCounts(p: PivotParts): void {
  const pivot = parse(p.pivot), cache = parse(p.cache)
  for (const name of ['rowFields', 'colFields', 'dataFields', 'filters', 'pageFields', 'pivotFields']) {
    const el = pivot.getElementsByTagName(name)[0]
    if (el) expect(Number(el.getAttribute('count')), name).toBe(el.children.length)
  }
  for (const items of Array.from(pivot.getElementsByTagName('items'))) expect(Number(items.getAttribute('count'))).toBe(items.children.length)
  const fields = cache.getElementsByTagName('cacheFields')[0]!
  expect(Number(fields.getAttribute('count'))).toBe(fields.children.length)
}

describe('interaction xml edits', async () => {
  const xmlCases = INTERACTION_CASES.filter(c => c.xmlEdit)
  const rows: Array<readonly [string, InteractionCase, PivotParts, boolean]> = []
  for (const c of xmlCases) { const { parts: p, exact } = await synthStart(c); rows.push([c.id, c, p, exact] as const) }
  it.each(rows)('%s: well-formed, consistent and carrying its marker', (_, c, start, exact) => {
    let next: PivotParts
    try { next = c.xmlEdit!({ ...start }) } catch (e) {
      // A base that is not built yet is replaced by the engine seed, which may lack what the edit needs (a page field).
      if (!exact) return
      throw e
    }
    expect(next.pivot !== start.pivot || next.cache !== start.cache).toBe(true)
    expectCounts(next)
    const fresh = [...c.markers.map(m => [m, start.pivot, next.pivot] as const), ...(c.cacheMarkers ?? []).map(m => [m, start.cache, next.cache] as const)]
      .filter(([m, before, after]) => !m.test(before) && m.test(after))
    const cleared = (c.absent ?? []).filter(m => m.test(start.pivot) && !m.test(next.pivot))
    expect(fresh.length + cleared.length, `${c.id}: the edit adds no marker and clears no absent pattern`).toBeGreaterThan(0)
  })
  it('adds the edit on an Excel-saved start file for the cases whose base exists', () => {
    expect(rows.filter(r => r[3]).map(r => r[0])).toEqual(expect.arrayContaining(['interaction-page-all', 'interaction-page-multi', 'caption-defaults', 'chart-value-filter']))
  })
  it('collapses and expands a field the way Excel saves it (default subtotal item included)', async () => {
    const saved = fixture('interaction/interaction-collapse-field-compact-top')
    if (!existsSync(saved)) return
    const regionItems = (pivot: string) => /<pivotField [^>]*axis="axisRow"[^>]*><items[^>]*>([\s\S]*?)<\/items>/.exec(pivot)![1]
    const excel = await parts(saved)
    const ours = INTERACTION_CASES.find(c => c.id === 'interaction-collapse-field-compact-top')!.xmlEdit!(await parts(fixture('recipes/engine-seed')))
    expect(regionItems(ours.pivot)).toBe(regionItems(excel.pivot))
    const expanded = INTERACTION_CASES.find(c => c.id === 'interaction-expand-field')!.xmlEdit!(excel)
    expect(expanded.pivot).not.toMatch(/\bsd="0"/)
  })
  it('summarises the page-all edit as the removed selection', () => {
    const row = rows.find(r => r[0] === 'interaction-page-all')!
    const { removed } = xmlDiffSummary(row[2].pivot, row[1].xmlEdit!({ ...row[2] }).pivot)
    expect(removed.some(t => /<pageField [^>]*item=/.test(t))).toBe(true)
  })
})

describe('build-engine --interaction --dry-run', () => {
  afterEach(() => { vi.restoreAllMocks() })
  it('prints every seed and case, with no failed edit or delta', async () => {
    const lines: string[] = []
    vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => { lines.push(a.join(' ')) })
    const failures = await interactionDryRun(INTERACTION_CASES, SEED_BUILDS)
    expect(failures).toBe(0)
    for (const c of INTERACTION_CASES) expect(lines.some(l => l.startsWith(`== ${c.id}  [`)), c.id).toBe(true)
    for (const s of SEED_BUILDS) expect(lines.some(l => l.startsWith(`== seed ${s.id}  [`)), s.id).toBe(true)
    // Stage 2 cases keep their own dry run (ids do not collide).
    expect(ENGINE_CASES.some(c => INTERACTION_CASES.some(i => i.id === c.id))).toBe(false)
  }, 60_000)
})

describe('verifyInteraction recomputation and strengthened markers', () => {
  it('rejects an unchanged / no-op pivot identical to its start file', async () => {
    const c = INTERACTION_CASES.find(x => x.id === 's4-values-showas-previous-next')!
    const startFile = fixture(startOf(c))
    await expect(verifyInteraction(c, startFile)).rejects.toThrow(/identical to start file|no-op interaction/)
  })
  it('rejects an XML refresh where the saved pivot cells in the location range are unchanged', async () => {
    const fixtureFile = fixture('interaction/s4-filtered-row-base')
    await expect(verifyXmlRecomputation(fixtureFile, fixtureFile)).rejects.toThrow(
      /saved pivot cells in location range .* are unchanged from edited input \(no recomputation\)/,
    )
  })
  it('proves XML recomputation on s4-filtered-row-base with cell diffs against the start file', async () => {
    const c = INTERACTION_CASES.find(x => x.id === 's4-filtered-row-base')!
    const startFile = fixture(startOf(c))
    const out = fixture(`interaction/${c.id}`)
    const recomputed = await verifyXmlRecomputation(startFile, out)
    expect(recomputed.cellDiffs).toBeGreaterThan(0)
    expect(recomputed.locChanged).toBe(true)
  })
  it('verifies built fixtures with their full measured deltas', async () => {
    for (const id of [
      's4-values-showas-previous-next',
      's4-group-date-manual-bounds',
      's4-group-number-sales-blanks-manual',
      's4-filtered-row-base',
      's4-group-number-units-auto',
      's4-regroup-number-units',
    ]) {
      const c = INTERACTION_CASES.find(x => x.id === id)!
      const out = fixture(`interaction/${id}`)
      if (existsSync(out)) {
        await expect(verifyInteraction(c, out)).resolves.toBeTruthy()
      }
    }
  })
})

const CONTROL_KINDS = new Set([
  'slicerSelect', 'slicerMultiSelect', 'slicerClear',
  'timelineSelect', 'timelineLevel', 'timelineClear',
  'tableSlicerSelect',
])

const CONTROL_KIND_CHECKS: Record<ControlOracleOp['kind'], (op: Record<string, unknown>) => boolean> = {
  slicerSelect: op => typeof op.name === 'string' && (op.itemIds === 'all' || (Array.isArray(op.itemIds) && (op.itemIds as unknown[]).every(Number.isInteger))),
  slicerMultiSelect: op => typeof op.name === 'string' && typeof op.enabled === 'boolean',
  slicerClear: op => typeof op.name === 'string',
  timelineSelect: op => typeof op.name === 'string' && typeof op.startDate === 'string' && typeof op.endDate === 'string',
  timelineLevel: op => typeof op.name === 'string' && ['years', 'quarters', 'months', 'days'].includes(String(op.level)),
  timelineClear: op => typeof op.name === 'string',
  tableSlicerSelect: op => typeof op.name === 'string' && (op.itemIds === 'all' || (Array.isArray(op.itemIds) && (op.itemIds as unknown[]).every(Number.isInteger))),
}

describe('Stage 5 control interaction cases (control-interaction-cases.ts)', () => {
  const CONTROLS_DIR = join(PIVOTS, 'controls')
  const CONTROLS_MANIFEST = join(CONTROLS_DIR, 'manifest.json')

  it('defines the exact Stage 5 data seed (5 rows, Date/Region/Product/Sales)', () => {
    expect(CONTROL_SEED_HEADERS).toEqual(['Date', 'Region', 'Product', 'Sales'])
    expect(CONTROL_SEED_ROWS).toEqual([
      ['2023-12-31', 'East', 'Apples', 10],
      ['2024-01-01', 'East', 'Bananas', 20],
      ['2024-02-15', 'West', 'Apples', 30],
      ['2024-03-31', 'West', 'Cherries', 40],
      ['2024-04-01', 'South', 'Apples', 50],
    ])
  })

  it('defines all 14 Stage 5 cases with exact names', () => {
    const EXPECTED_CASES = [
      's5-slicer-click-one',
      's5-slicer-click-selected',
      's5-slicer-modifier-add-remove',
      's5-slicer-multiselect-toggle',
      's5-slicer-clear',
      's5-slicer-crossfilter-order',
      's5-slicer-hidden-or-blank',
      's5-timeline-click-month',
      's5-timeline-drag-forward-reverse',
      's5-timeline-level-preserves-range',
      's5-timeline-clear-and-scroll',
      's5-linked-two-pivots',
      's5-linked-slicer-and-timeline',
      's5-table-slicer',
    ]
    expect(CONTROL_INTERACTION_CASES.length).toBe(14)
    expect(CONTROL_INTERACTION_CASES.map(c => c.case)).toEqual(EXPECTED_CASES)
  })

  it('validates each control interaction case against the schema', () => {
    for (const c of CONTROL_INTERACTION_CASES) {
      expect(typeof c.case).toBe('string')
      expect(typeof c.file).toBe('string')
      expect(typeof c.base).toBe('string')
      expect(c.target).toBeDefined()
      expect(typeof c.target.controlName).toBe('string')
      expect(['xml', 'applescript', 'recipe']).toContain(c.route)
      expect(['built', 'recipe', 'unverified']).toContain(c.status)
      expect(c.expected).toBeDefined()

      const actions = Array.isArray(c.action) ? c.action : [c.action]
      expect(actions.length).toBeGreaterThan(0)
      for (const op of actions) {
        const typedOp = op as ControlOracleOp
        expect(CONTROL_KINDS.has(typedOp.kind), `unknown op kind ${typedOp.kind}`).toBe(true)
        const check = CONTROL_KIND_CHECKS[typedOp.kind]
        expect(check(op as unknown as Record<string, unknown>), `invalid action shape: ${JSON.stringify(op)}`).toBe(true)
      }
    }
  })

  it('validates controls manifest matches cases when present', () => {
    if (!existsSync(CONTROLS_MANIFEST)) return
    const raw = JSON.parse(readFileSync(CONTROLS_MANIFEST, 'utf8')) as ControlInteractionCase[]
    expect(Array.isArray(raw)).toBe(true)
    expect(raw.length).toBe(14)
    expect(raw.map(c => c.case)).toEqual(CONTROL_INTERACTION_CASES.map(c => c.case))
  })
})

