import { afterEach, describe, expect, it, vi } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { ENGINE_CASES, engineSeedStandIn, xmlDiffSummary, type PivotParts } from '../scripts/xlsx-refs/engine-cases'
import {
  INTERACTION_CASES, INTERACTION_PROBES, SEED_BUILDS, SEED_FILES, resolveDelta, startOf, startShapeOf, type InteractionCase, type PivotDelta,
} from '../scripts/xlsx-refs/interaction-cases'
import { classify } from '../scripts/xlsx-refs/classify-stage2'
import { interactionDryRun } from '../scripts/xlsx-refs/build-engine'

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

/** A structural check against Task 25's PivotDelta; the Record over `kind` fails to compile when a kind is added. */
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
      const items = 'field' in d ? shape.items[d.field] ?? shape.shared[d.field]! : []
      if (d.kind === 'collapse') expect(items[d.item], c.case).toBeTruthy()
      if (d.kind === 'items') for (const k of d.hidden) expect(items[k], c.case).toBeTruthy()
    }
  })
  it('classifies the cases whose step is a Stage 3 UI action', () => {
    const delta = (id: string) => classified.find(c => c.case === id)?.delta
    expect(delta('baseline')).toEqual([])
    expect(delta('layout-collapsed-item')).toEqual([{ kind: 'collapse', field: 0, item: 0, collapsed: true }])
    expect(delta('quirk-hidden-row-item')).toEqual([{ kind: 'items', field: 0, hidden: [3] }])
    expect(delta('sort-value-ties')).toEqual([{ kind: 'sort', field: 0, order: 'descending', byData: 1 }])
    expect(delta('filter-top-count')).toMatchObject([{ kind: 'filter', field: 0, filter: { type: 'count', fld: 0 } }])
    for (const id of ['layout-collapsed-item-above-values', 'quirk-page-single', 'values-average-hidden-item', 'filter-label-renamed-item', 'sort-value-showall-no-data'])
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
      expect([25, 29, 30, 31, 32], c.id).toContain(c.task)
      expect(Object.keys(SEED_FILES), c.id).toContain(c.seed)
      if (c.route === 'applescript') expect(c.lines, c.id).toBeTypeOf('function')
      if (c.route === 'xml' || c.fallback === 'xml') expect(c.xmlEdit, c.id).toBeTypeOf('function')
      const proofs = c.markers.length + (c.cacheMarkers?.length ?? 0) + (c.absent?.length ?? 0) + (c.recordMarkers?.length ?? 0) + (c.sheetMarkers?.length ?? 0)
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
  it('covers every interaction kind of the brief', () => {
    const kinds = new Set(INTERACTION_CASES.flatMap(c => c.delta.map(d => d.kind)))
    expect([...kinds].sort()).toEqual(['clearFilters', 'collapse', 'collapseAll', 'collapseField', 'filter', 'items', 'page', 'sort'])
    for (const id of ['interaction-filter-all-hidden', 'interaction-page-multi-to-all', 'interaction-clear-label-filter', 'interaction-clear-value-filter',
      'caption-defaults', 'chart-baseline', 'chart-hide-item', 'chart-collapse-item', 'chart-value-filter', 'chart-sort', 'chart-values-on-rows', 'chart-page-select',
      'gpd-hide-item', 'gpd-collapse-item', 'gpd-filter-out', 'refresh-on-load', 'refresh-on-load-nodata'])
      expect(ids, id).toContain(id)
    for (const f of ['compact-top', 'compact-bottom', 'outline-top', 'outline-bottom', 'tabular-top', 'tabular-bottom'])
      for (const what of ['item', 'field']) expect(ids).toContain(`interaction-collapse-${what}-${f}`)
    expect(new Set(INTERACTION_CASES.map(c => c.task))).toEqual(new Set([25, 29, 30, 31, 32]))
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
