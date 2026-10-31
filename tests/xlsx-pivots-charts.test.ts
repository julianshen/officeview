// @vitest-environment node
/**
 * Pivot charts follow (plan Task 30, spec §6.6, §7.3, CODEX-3): a pivot chart's series, categories and caches are
 * rebuilt from its pivot's current layout, and must equal what Excel saved for the same change.
 *
 * Oracle: Task 24's chart cases (tests/fixtures/xlsx-pivots/interaction/manifest.json, task 30). Each opens on
 * `base ?? seed` and applies its `delta`; a case with `setup` (a field move Excel made before saving, which Stage 3
 * cannot apply) is checked on its own file (its saved layout against its own saved chart) and across files (the
 * seed's chart rebuilt over the case's layout). Measured in docs/validation/xlsx-pivots.md "Pivot charts follow" and
 * docs/validation/xlsx-pivots-stage3-measurements.md §(c).
 */
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { loadPivotWorkbook } from '../src/xlsx/pivots/load'
import { decodeRecords, RECORD_LIMITS } from '../src/xlsx/pivots/engine/records'
import { sheetCellReader } from '../src/xlsx/pivots/engine/faithful'
import { openPivotSession, type PivotSession } from '../src/xlsx/pivots/session/session'
import type { PivotDelta } from '../src/xlsx/pivots/session/delta'
import { pivotChartGate, pivotChartSeries } from '../src/xlsx/pivots/session/charts'
import type { PivotCacheDefinition, PivotTableDefinition } from '../src/xlsx/pivots/model'
import type { RecordStore } from '../src/xlsx/pivots/engine/records'
import type { ChartModel, ChartPoints, ChartSeries } from '../src/charts/model'
import type { XlsxDocument, XlsxSheet } from '../src/xlsx/types'
import { sheetCharts } from '../src/components/xlsx/sheetCharts'
import { installPivotCharts } from '../src/components/xlsx/sheetAssets'
import { chartDescription } from '../src/charts/describe'
import { refillChartModel } from '../src/charts/fill'
import { adoptChartModel, chartModelOwner, modelOwner, releaseChartModel } from '../src/charts/own'
import { createChartWorkspace } from '../src/charts/workspace'
import { once } from './helpers/xlsx-pivots-load'

const FIXTURES = fileURLToPath(new URL('./fixtures/xlsx-pivots/', import.meta.url))
const INTERACTION = resolve(FIXTURES, 'interaction')
const SEED = resolve(FIXTURES, 'recipes/pivot-chart-seed.xlsx')
interface ChartCase { case: string; file: string; pivot: string; task: number; seed: string; delta: PivotDelta[]; setup?: string; base?: string }
const CASES: ChartCase[] = (JSON.parse(readFileSync(resolve(INTERACTION, 'manifest.json'), 'utf8')) as ChartCase[]).filter(entry => entry.task === 30)
const caseFile = (entry: ChartCase) => resolve(FIXTURES, '..', '..', '..', entry.file)
const startFile = (entry: ChartCase) => entry.base ? resolve(FIXTURES, `${entry.base}.xlsx`) : SEED

interface Book { doc: XlsxDocument; sheet: XlsxSheet; def: PivotTableDefinition; cache: PivotCacheDefinition; store: RecordStore; charts: ChartModel[] }
const books = new Map<string, Promise<Book>>()
/** A workbook as the resident viewer has it (the parsed document, its charts) and its one pivot, decoded. */
function book(file: string): Promise<Book> {
  let found = books.get(file)
  if (!found) books.set(file, found = (async () => {
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(file)))
    const doc = await parseXlsx(pkg)
    const regions = new Map(doc.sheets.filter(sheet => sheet.pivotRegions?.length).map(sheet => [sheet.name, sheet.pivotRegions!]))
    const workbook = await loadPivotWorkbook(pkg, regions)
    const def = workbook.tables[0]!, cache = workbook.caches.get(def.cacheId)!
    const records = await decodeRecords(once((await pkg.bytes(cache.recordsPartPath!))!), cache, RECORD_LIMITS.D, false)
    if (!records.ok) throw new Error(records.diagnostic.message)
    const charts = doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings))
    return { doc, sheet: doc.sheets.find(sheet => sheet.name === def.sheet)!, def, cache, store: records.store, charts }
  })())
  return found
}
/** Opens a session on a fresh load of `file`'s pivot with the Task 30 gate over its linked charts. */
async function session(file: string, gate = pivotChartGate): Promise<{ book: Book; session: PivotSession }> {
  const b = await book(file)
  const opened = openPivotSession({
    def: b.def, cache: b.cache, store: b.store, date1904: false, saved: sheetCellReader(b.sheet), occupied: () => false,
    linkedCharts: b.charts, chartGate: gate,
  })
  if (!opened.ok) throw new Error(`${file}: open refused ${opened.diagnostic.feature}: ${opened.diagnostic.message}`)
  return { book: b, session: opened.value }
}

const n15 = (v: number) => Number(v.toPrecision(15))
const points = (p: ChartPoints) => Array.from({ length: p.length }, (_, i) => {
  const format = p.formatIndex && p.formats ? p.formats[(p.formatIndex[i] ?? 0) - 1] : undefined
  return p.states[i] === 1 ? [n15(p.values[i]!), format ?? null] : p.states[i]
})
/** What the chart shows of the pivot: categories, outer levels, references and each series' name, colour and points. */
function chartData(chart: ChartModel) {
  return {
    categories: chart.categories,
    levels: (chart.categoryLevels ?? []).map(level => ({ starts: [...level.starts], labels: level.labels })),
    categoryRef: chart.categoryRef,
    groups: chart.groups.length,
    series: chart.groups.flatMap(group => group.series.map((s: ChartSeries) => ({ index: s.index, name: s.name, color: s.color, ref: s.refs?.values, values: points(s.values), categories: s.categories }))),
  }
}

describe('pivot charts follow (Task 30)', () => {
  it('the manifest lists the measured chart cases', () => {
    expect(CASES.map(entry => entry.case).sort()).toEqual(['chart-baseline', 'chart-collapse-item', 'chart-hide-item', 'chart-page-base', 'chart-page-select', 'chart-sort', 'chart-value-filter', 'chart-values-on-rows'])
    for (const entry of CASES) expect(existsSync(caseFile(entry)), entry.file).toBe(true)
  })

  describe('series, categories and caches equal Excel\'s saved chart for every chart oracle case', () => {
    for (const entry of CASES) {
      it(entry.case, async () => {
        const target = await book(caseFile(entry))
        expect(target.charts).toHaveLength(1)
        // A setup case starts from its own saved state (the field move is Stage 4); the others replay their delta.
        const { session: s } = await session(entry.setup ? caseFile(entry) : startFile(entry))
        let charts = s.reset().charts
        for (const delta of entry.setup ? [] : entry.delta) {
          const applied = s.apply(delta)
          if (!applied.ok) throw new Error(`${entry.case}: ${applied.diagnostic.feature}: ${applied.diagnostic.message}`)
          expect(s.commit(applied.value)).toBe(true)
          charts = applied.value.charts
        }
        expect(charts).toHaveLength(1)
        expect(chartData(charts[0]!)).toEqual(chartData(target.charts[0]!))
        // The rebuilt chart keeps its link and its own chart-level settings.
        expect(charts[0]!.pivot).toEqual((await book(entry.setup ? caseFile(entry) : startFile(entry))).charts[0]!.pivot)
        expect(charts[0]!.preparedAxes).toBeDefined()
      })
    }
    it('the seed\'s chart rebuilt over each field-move case\'s saved layout equals that case\'s saved chart', async () => {
      const seed = await book(SEED)
      for (const entry of CASES.filter(c => c.setup)) {
        const { session: s, book: target } = await session(caseFile(entry))
        const rebuilt = pivotChartSeries(seed.charts[0]!, s.layout, s.def)
        if (!rebuilt.ok) throw new Error(`${entry.case}: ${rebuilt.diagnostic.message}`)
        expect(chartData(rebuilt.value), entry.case).toEqual(chartData(target.charts[0]!))
      }
    })
  })

  describe('pivotFmts follow the measured rule when a series appears or disappears', () => {
    const xml = async (file: string, part: string) => new TextDecoder().decode((await (await OfficePackage.load(new Uint8Array(readFileSync(file)))).bytes(part))!)
    it('Excel keeps every pivotFmt and the pivot\'s chartFormats when series disappear; series formats are automatic by position', async () => {
      const seedChart = await xml(SEED, 'xl/charts/chart1.xml')
      const seedFormats = /<chartFormats[\s\S]*?<\/chartFormats>/.exec(await xml(SEED, 'xl/pivotTables/pivotTable1.xml'))![0]
      const pivotFmts = (text: string) => text.match(/<c:pivotFmt>/g)?.length ?? 0
      expect(pivotFmts(seedChart)).toBe(6)
      for (const entry of CASES) {
        const chart = await xml(caseFile(entry), 'xl/charts/chart1.xml')
        // 6 pivotFmts in every case, values-on-rows' 3 series included; chartFormats byte-identical to the seed's.
        expect(pivotFmts(chart), entry.case).toBe(6)
        expect(/<chartFormats[\s\S]*?<\/chartFormats>/.exec(await xml(caseFile(entry), 'xl/pivotTables/pivotTable1.xml'))![0], entry.case).toBe(seedFormats)
        // No series carries its own format: c:idx and c:order are its position.
        const series = chart.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? []
        series.forEach((ser, k) => {
          expect(ser.includes('<c:spPr>'), `${entry.case} series ${k}`).toBe(false)
          expect(ser).toContain(`<c:idx val="${k}"/><c:order val="${k}"/>`)
        })
      }
    })
    it('a series that disappears takes nothing with it: the series left take the formats of their positions', async () => {
      const seed = await book(SEED), target = await book(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      const { session: s } = await session(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      const rebuilt = pivotChartSeries(seed.charts[0]!, s.layout, s.def)
      if (!rebuilt.ok) throw new Error(rebuilt.diagnostic.message)
      const series = rebuilt.value.groups[0]!.series
      expect(series.map(x => x.index)).toEqual([0, 1, 2])
      expect(series.map(x => x.color)).toEqual(seed.charts[0]!.groups[0]!.series.slice(0, 3).map(x => x.color))
      expect(series.map(x => x.color)).toEqual(target.charts[0]!.groups[0]!.series.map(x => x.color))
    })
    it('a series past the saved chart\'s series has no measured format and refuses pivot-chart-mapping', async () => {
      const fewer = await book(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      const { session: s } = await session(SEED)
      const rebuilt = pivotChartSeries(fewer.charts[0]!, s.layout, s.def)
      expect(rebuilt.ok ? 'ok' : rebuilt.diagnostic.feature).toBe('pivot-chart-mapping')
    })
    it('a single-colour chart style refuses a series change (its shades follow the series count)', async () => {
      const seed = await book(SEED)
      const { session: rows } = await session(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      const mono = pivotChartSeries({ ...seed.charts[0]!, styleId: 3 }, rows.layout, rows.def)
      expect(mono.ok ? 'ok' : mono.diagnostic.feature).toBe('pivot-chart-mapping')
      const { session: same } = await session(SEED)
      const kept = pivotChartSeries({ ...seed.charts[0]!, styleId: 3 }, same.layout, same.def)
      expect(kept.ok ? 'ok' : kept.diagnostic.message).toBe('ok')
    })
    it('a series with its own format refuses when the series change, and keeps it when they do not', async () => {
      const seed = await book(SEED)
      const chart = seed.charts[0]!
      const formatted: ChartModel = { ...chart, groups: chart.groups.map(group => ({ ...group, series: group.series.map((x, k) => k === 0 ? { ...x, fill: { kind: 'solid', color: 'FF0000' } as unknown as ChartSeries['fill'] } : x) })) }
      const { session: same } = await session(SEED)
      const kept = pivotChartSeries(formatted, same.layout, same.def)
      if (!kept.ok) throw new Error(kept.diagnostic.message)
      expect(kept.value.groups[0]!.series[0]!.fill).toEqual(formatted.groups[0]!.series[0]!.fill)
      const { session: rows } = await session(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      const moved = pivotChartSeries(formatted, rows.layout, rows.def)
      expect(moved.ok ? 'ok' : moved.diagnostic.feature).toBe('pivot-chart-mapping')
    })
  })

  describe('an uncovered mapping refuses and the pivot is not interactive', () => {
    const refusedAt = async (file: string, linked: (charts: ChartModel[]) => ChartModel[]) => {
      const b = await book(file)
      const opened = openPivotSession({ def: b.def, cache: b.cache, store: b.store, date1904: false, saved: sheetCellReader(b.sheet), occupied: () => false, linkedCharts: linked(b.charts), chartGate: pivotChartGate })
      return opened.ok ? 'ok' : opened.diagnostic.feature
    }
    it('a chart family no case measured refuses at open', async () => {
      expect(await refusedAt(SEED, charts => charts.map(chart => ({ ...chart, groups: chart.groups.map(group => ({ ...group, family: 'line' as const })) })))).toBe('pivot-chart-mapping')
    })
    it('a chart whose saved series are not its pivot\'s default mapping refuses at open', async () => {
      // The values-on-rows chart (3 series by Year) linked to the seed's pivot (6 series by Year × Values).
      // (A fresh object: the gate has not seen it. It refuses as soon as its 3 series cannot hold the seed's 6.)
      const other = await book(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      expect(await refusedAt(SEED, () => [{ ...other.charts[0]! }])).toBe('pivot-chart-mapping')
      // Reordered series are not the default mapping: the open-time check (rebuilt ≠ saved) refuses them.
      expect(await refusedAt(SEED, charts => charts.map(chart => ({ ...chart, groups: chart.groups.map(group => ({ ...group, series: [...group.series].reverse() })) })))).toBe('pivot-chart-mapping')
    })
    it('the gate opens a chart-linked pivot, and a later layout no case covers refuses with the session left at its revision', async () => {
      // Values on rows with an outer item collapsed: no case measured where the chart's points go.
      const { session: s } = await session(caseFile(CASES.find(c => c.case === 'chart-values-on-rows')!))
      expect(s.revision).toBe(0)
      const refused = s.apply({ kind: 'collapse', field: 0, item: 0, collapsed: true })
      expect(refused.ok ? 'ok' : refused.diagnostic.feature).toBe('pivot-chart-mapping')
      expect(s.revision).toBe(0)
      // A covered change on the same session still applies.
      const sorted = s.apply({ kind: 'sort', field: 0, order: 'descending' })
      expect(sorted.ok ? 'ok' : sorted.diagnostic.message).toBe('ok')
    })
    it('a pivot layout no case measured refuses the mapping (pivotChartSeries direct)', async () => {
      const { session: s, book: b } = await session(SEED)
      // Tabular form: two label columns.
      const tabular = { ...s.def, location: { ...s.def.location, firstDataCol: 2 } }
      const r1 = pivotChartSeries(b.charts[0]!, s.layout, tabular)
      expect(r1.ok ? 'ok' : r1.diagnostic.feature).toBe('pivot-chart-mapping')
      // Grand totals off.
      const r2 = pivotChartSeries(b.charts[0]!, s.layout, { ...s.def, colGrandTotals: false })
      expect(r2.ok ? 'ok' : r2.diagnostic.feature).toBe('pivot-chart-mapping')
      // A chart linked to another pivot.
      const r3 = pivotChartSeries({ ...b.charts[0]!, pivot: { ...b.charts[0]!.pivot!, name: '[x.xlsx]Pivot!Other' } }, s.layout, s.def)
      expect(r3.ok ? 'ok' : r3.diagnostic.feature).toBe('pivot-chart-mapping')
    })
  })

  describe('the chart repaints after a commit (resident)', () => {
    it('a commit\'s charts replace the linked charts in the drawing tree, and reset brings back the saved data', async () => {
      // A fresh document (the cached books are shared): the viewer's own doc.
      books.delete(SEED)
      const { session: s, book: b } = await session(SEED)
      const saved = chartData(b.charts[0]!)
      const hide = CASES.find(c => c.case === 'chart-hide-item')!
      const applied = s.apply(hide.delta[0]!)
      if (!applied.ok) throw new Error(applied.diagnostic.message)
      expect(s.commit(applied.value)).toBe(true)
      // The sheet holding the chart is returned for its tiles to be dropped.
      const touched = installPivotCharts(b.doc, b.def.name, applied.value.charts)
      expect(touched).toEqual([b.doc.sheets.findIndex(sheet => sheetCharts(sheet.drawings).length > 0)])
      const shown = b.doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings))
      expect(shown).toHaveLength(1)
      expect(shown[0]).toBe(applied.value.charts[0])
      expect(chartData(shown[0]!)).toEqual(chartData((await book(caseFile(hide))).charts[0]!))
      // What the sheet's chart list and painter read is the new model (9 points for 12).
      expect(chartDescription(shown[0]!)).not.toBe(chartDescription(b.charts[0]!))
      // Reset installs the saved state's charts again.
      const reset = s.reset()
      expect(s.commit(reset)).toBe(true)
      installPivotCharts(b.doc, b.def.name, reset.charts)
      expect(chartData(b.doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings))[0]!)).toEqual(saved)
    })
    it('charts of another pivot, and a count that does not match, are left alone', async () => {
      books.delete(SEED)
      const { book: b } = await session(SEED)
      const original = b.doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings))[0]
      expect(installPivotCharts(b.doc, 'Another pivot', [original!])).toEqual([])
      expect(installPivotCharts(b.doc, b.def.name, [])).toEqual([])
      expect(b.doc.sheets.flatMap(sheet => sheetCharts(sheet.drawings))[0]).toBe(original)
    })
  })

  describe('refillChartModel (src/charts/fill.ts)', () => {
    it('adopts the rebuilt model under its own tag in the source\'s workspace, leaving the source\'s charge alone', async () => {
      const seed = await book(SEED)
      const workspace = createChartWorkspace('desktop')
      const source: ChartModel = { ...seed.charts[0]! }
      const owner = modelOwner(workspace)
      expect(workspace.reserve(1000, owner.tag)).toBe(true)
      adoptChartModel(source, owner)
      const { session: s } = await session(SEED)
      const rebuilt = pivotChartSeries(source, s.layout, s.def)
      if (!rebuilt.ok) throw new Error(rebuilt.diagnostic.message)
      const next = rebuilt.value
      expect(chartModelOwner(source)).toBe(owner)
      expect(chartModelOwner(next)?.workspace).toBe(workspace)
      expect(chartModelOwner(next)?.tag).not.toBe(owner.tag)
      const charged = workspace.used
      expect(charged).toBeGreaterThan(1000)
      releaseChartModel(next)
      expect(workspace.used).toBeLessThan(charged)
      releaseChartModel(source)
    })
    it('a refused admission refunds the refill and the gate refuses the change', async () => {
      const seed = await book(SEED)
      const workspace = createChartWorkspace('mobile')
      const source: ChartModel = { ...seed.charts[0]! }
      const owner = modelOwner(workspace)
      // Leave less room than the rebuilt series need.
      expect(workspace.reserve(workspace.limit - 512, owner.tag)).toBe(true)
      adoptChartModel(source, owner)
      const used = workspace.used
      const refused = refillChartModel(source, charge => { charge(4096); return source })
      expect(refused.refused).toBe(true)
      expect(workspace.used).toBe(used)
      const { session: s } = await session(SEED)
      const rebuilt = pivotChartSeries(source, s.layout, s.def)
      expect(rebuilt.ok ? 'ok' : rebuilt.diagnostic.feature).toBe('pivot-chart-memory')
      expect(workspace.used).toBe(used)
      releaseChartModel(source)
    })
  })
})
