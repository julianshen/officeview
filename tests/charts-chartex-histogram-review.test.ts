/**
 * Task B3 review round 1 (codex): a fresh engine import (I1); a native chart refused once its data is read gives way
 * to its mc:Fallback with fallback coverage (I2); range-safe bins, sums and fractions (I3); hidden Pareto lines and
 * hidden columns are out of layout (I4); point overrides kept where a column is one source point (I5).
 *
 * HI2 variants are derived from the native Office part (tests/fixtures/chartex); inline parts are Office-shaped
 * synthetic.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContextFor } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { paintXlsxChart } from '../src/charts/paint'
import { pointList } from '../src/charts/points'
import type { ChartModel, ChartResult } from '../src/charts/model'
import { parseXlsx } from '../src/xlsx/parse'

const ROOT = join(__dirname, '..')
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const theme = parseThemeContext()
const HI2 = join(__dirname, 'fixtures', 'chartex', 'tdf163727_histogram_underflow_overflow', 'tdf163727_histogram_underflow_overflow.xlsx')
const space = (body: string) => `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</cx:chartSpace>`
const AXES = '<cx:axis id="0"><cx:catScaling gapWidth="0" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:tickLabels /></cx:axis><cx:axis id="2"><cx:valScaling max="1" min="0" /><cx:units unit="percentage" /><cx:tickLabels /></cx:axis>'
const lvl = (values: Array<number | string | undefined>) => `<cx:lvl ptCount="${values.length}">${values.map((v, i) => v === undefined ? '' : `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`
const data = (id: number, values: Array<number | string | undefined>, categories?: string[]) =>
  `<cx:data id="${id}">${categories ? `<cx:strDim type="cat">${lvl(categories)}</cx:strDim>` : ''}<cx:numDim type="val">${lvl(values)}</cx:numDim></cx:data>`
const column = (id: number, layoutPr: string, extra = '', attrs = '') => `<cx:series layoutId="clusteredColumn"${attrs}><cx:dataId val="${id}" />${extra}<cx:layoutPr>${layoutPr}</cx:layoutPr><cx:axisId val="1" /></cx:series>`
const chart = (sets: string, series: string) => space(`<cx:chartData>${sets}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${AXES}</cx:plotArea></cx:chart>`)
async function adapt(xml: string): Promise<ChartResult> {
  const zip = new JSZip(), path = 'ppt/charts/chartEx1.xml'
  zip.file(path, xml)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const read = await readChartPart(pkg, path, chartParseContextFor(pkg, path))
  if (read?.kind !== 'tree') throw new Error('read')
  try { return await parseChartPart(read.root, theme, chartParseContextFor(pkg, path)) } finally { read.release() }
}
async function native(xml: string): Promise<ChartModel> {
  const result = await adapt(xml)
  if (result.kind !== 'native') throw new Error(`${result.kind}: ${result.reason}`)
  return result.model
}
const nums = (model: ChartModel, group = 0, series = 0) => pointList(model.groups[group].series[series].values).map(p => p.state === 'value' ? p.v : p.state)
/** An HI2 variant: its chart part, sheet or drawing rewritten. */
async function hi2(edit: { chart?: (xml: string) => string; sheet?: (xml: string) => string; drawing?: (xml: string) => string }): Promise<OfficePackage> {
  const zip = await JSZip.loadAsync(readFileSync(HI2))
  for (const [path, fn] of [['xl/charts/chartEx1.xml', edit.chart], ['xl/worksheets/sheet1.xml', edit.sheet], ['xl/drawings/drawing1.xml', edit.drawing]] as const) {
    if (!fn) continue
    const before = await zip.file(path)!.async('string'), after = fn(before)
    if (after === before) throw new Error(`${path}: the variant edit did not apply`)
    zip.file(path, after)
  }
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}

describe('B3 review round 1', () => {
  it('I1: the engine entry and its modules import fresh, in any order (no initialization cycle)', () => {
    for (const entry of ['src/charts/index.ts', 'src/charts/chartex/histogram.ts', 'src/charts/prepare.ts', 'src/charts/fill.ts', 'src/charts/parse.ts', 'src/charts/chartex/parse.ts', 'src/charts/dataPrecedence.ts']) {
      const run = spawnSync('bun', ['-e', `await import(${JSON.stringify(join(ROOT, entry))}); console.log('imported')`], { cwd: ROOT, encoding: 'utf8', timeout: 60_000 })
      expect(`${run.stdout}${run.stderr}`, entry).toContain('imported')
      expect(run.status, entry).toBe(0)
    }
    // The family module depends only on leaf modules: never upward on the fill, parser, stage 1 or adapter.
    const source = readFileSync(join(ROOT, 'src/charts/chartex/histogram.ts'), 'utf8')
    const imports = [...source.matchAll(/^import (?!type)[^\n]* from '([^']+)'/gm)].map(m => m[1])
    for (const upward of ['../fill', '../parse', '../prepare', '../dataPrecedence', './parse']) expect(imports, upward).not.toContain(upward)
  }, 120_000)

  it('I2: a native chart refused once its data is read gives way to its mc:Fallback, with fallback coverage', async () => {
    // HI2 with binSize 0.000001: (14 - 11) / 1e-6 = 3,000,000 inner bins, refused at the fill (map §5.1 step 7);
    // HI2 with A1 a formula without a saved value (saved-values policy): an unavailable point, refused.
    const variants: Array<[string, Promise<OfficePackage>, string, object?]> = [
      ['binSize 0.000001', hi2({ chart: xml => xml.replace(/<cx:binSize val="2" ?\/>/, '<cx:binSize val="0.000001" />') }), 'invalid binning'],
      ['unavailable point', hi2({ sheet: xml => xml.replace(/<c r="A1"([^>]*)><v>12<\/v><\/c>/, '<c r="A1"$1><f>6*2</f></c>') }), "This chart's data is not available", { valuePolicy: 'cached' }],
    ]
    for (const [name, pkg, reason, options] of variants) {
      const xlsx = await parseXlsx(await pkg, options as never)
      const sheet = xlsx.sheets[0]
      const drawing = sheet.drawings?.[0]
      // The fallback text box, not the chart, not a card.
      expect(drawing?.content, name).toBeUndefined()
      expect(drawing?.source.representation, name).toBe('fallback')
      expect(drawing?.source.element, name).toBe('sp')
      expect(drawing?.textBody, name).toBeDefined()
      expect(drawing?.widthEmu, `${name}: placed on the anchor`).toBeGreaterThan(0)
      for (const coverage of [sheet.drawingCoverage ?? [], (xlsx.drawingCoverage ?? []).filter(e => e.unit === 0)]) {
        expect(coverage.some(e => e.selectedRepresentation === 'chart'), `${name}: no stale native entry`).toBe(false)
        expect(coverage.find(e => e.scope === 'original'), name).toMatchObject({ feature: 'ChartEx', status: 'fallback', representation: 'fallback', reason })
      }
    }
    // Control: the same refusal with no fallback keeps the chart's card, and the coverage says card (not native).
    const bare = await parseXlsx(await hi2({ chart: xml => xml.replace(/<cx:binSize val="2" ?\/>/, '<cx:binSize val="0.000001" />'), drawing: xml => xml.replace(/<mc:Fallback>[\s\S]*<\/mc:Fallback>/, '') }))
    const card = bare.sheets[0].drawings?.[0].content
    expect(card?.kind).toBe('xlsxChart')
    if (card?.kind === 'xlsxChart') expect(card.chart.unsupported).toBe('invalid binning')
    expect(bare.sheets[0].drawingCoverage?.find(e => e.element === 'graphicFrame')?.selectedRepresentation).not.toBe('chart')
    // The native HI2 itself is unchanged (still the chart).
    const ok = await parseXlsx(await hi2({}))
    expect(ok.sheets[0].drawings?.[0].content?.kind).toBe('xlsxChart')
    expect(ok.sheets[0].drawingCoverage?.find(e => e.element === 'graphicFrame')).toMatchObject({ selectedRepresentation: 'chart' })
  })

  it('I3: bins, sums and cumulative fractions stay finite on representable inputs; the unrepresentable is refused', async () => {
    // [-1e308, 1e308], binCount 2: edges -1e308, 0, 1e308; one value per bin.
    const wide = await native(chart(data(0, [-1e308, 1e308]), column(0, '<cx:binning><cx:binCount val="2" /></cx:binning>')))
    expect(wide.categories).toEqual(['[-1E+308, 0]', '(0, 1E+308]'])
    expect(nums(wide)).toEqual([1, 1])
    // By Category: A = 1e308 + 1e308 - 1e308 = 1e308 (a finite sum, whatever the order of partial sums).
    expect(nums(await native(chart(data(0, [1e308, 1e308, -1e308], ['A', 'A', 'A']), column(0, '<cx:aggregation />'))))).toEqual([1e308])
    // A Pareto over A = 1e308, B = 1e308: fractions 0.5, 1 (the total 2e308 is never formed).
    const pareto = await native(chart(data(0, [1e308, 1e308], ['A', 'B']), column(0, '<cx:aggregation />') + '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2" /></cx:series>'))
    expect(nums(pareto, 1)).toEqual([0.5, 1])
    // Unrepresentable: a column of 2e308 is refused, never drawn as Infinity.
    expect(await adapt(chart(data(0, [1e308, 1e308], ['A', 'A']), column(0, '<cx:aggregation />')))).toMatchObject({ kind: 'refused', reason: "This chart's values are out of range", feature: 'ChartEx' })
    // Degenerate geometry: edges lost to precision (1e16 + 1 = 1e16) are invalid binning.
    expect(await adapt(chart(data(0, [1e16, 1e16 + 4]), column(0, '<cx:binning><cx:binSize val="1" /></cx:binning>')))).toMatchObject({ kind: 'refused', reason: 'invalid binning' })
    // Every published point is finite.
    for (const model of [wide, pareto]) for (const g of model.groups) for (const s of g.series) expect(pointList(s.values).every(p => p.state === 'value' && Number.isFinite(p.v))).toBe(true)
  })

  it('I4: a hidden Pareto line and hidden columns are out of layout; ownerIdx keeps document positions', async () => {
    const line = (attrs = '', owner = 0) => `<cx:series layoutId="paretoLine" ownerIdx="${owner}"${attrs}><cx:axisId val="2" /></cx:series>`
    // A hidden line: the columns are still sorted by their owner, and no line is drawn.
    const hiddenLine = await native(chart(data(0, [1, 3], ['A', 'B']), column(0, '<cx:aggregation />') + line(' hidden="1"')))
    expect(hiddenLine.groups.map(g => g.family)).toEqual(['bar'])
    expect(hiddenLine.categories).toEqual(['B', 'A'])
    // A hidden sibling column (position 1) does not block the visible owner (position 0) and is not drawn.
    const sibling = await native(chart(data(0, [1, 3], ['A', 'B']) + data(1, [9], ['Z']), column(0, '<cx:aggregation />') + column(1, '<cx:aggregation />', '', ' hidden="1"') + line()))
    expect(sibling.groups[0].series).toHaveLength(1)
    expect(sibling.categories).toEqual(['B', 'A'])
    expect(nums(sibling, 1)).toEqual([0.75, 1])
    // A hidden column first: the owner at document position 1 is named as 1 (never renumbered).
    const second = await native(chart(data(0, [9], ['Z']) + data(1, [1, 3], ['A', 'B']), column(0, '<cx:aggregation />', '', ' hidden="1"') + column(1, '<cx:aggregation />') + line('', 1)))
    expect(second.categories).toEqual(['B', 'A'])
    expect(nums(second)).toEqual([3, 1])
    // A visible owner plus another visible column is still unsupported; so is a hidden owner beside a visible column.
    expect(await adapt(chart(data(0, [1]) + data(1, [2]), column(0, '<cx:binning />') + column(1, '<cx:binning />') + line()))).toMatchObject({ kind: 'unsupported' })
    expect(await adapt(chart(data(0, [1]) + data(1, [2]), column(0, '<cx:binning />', '', ' hidden="1"') + column(1, '<cx:binning />') + line()))).toMatchObject({ kind: 'unsupported' })
  })

  it('I5: point and label overrides follow their source point where a column is one point; binned columns drop them, diagnosed', async () => {
    const red = '<cx:dataPt idx="0"><cx:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></cx:spPr></cx:dataPt>'
    const labels = '<cx:dataLabels><cx:visibility value="1" /><cx:dataLabel idx="0"><cx:visibility value="0" /></cx:dataLabel></cx:dataLabels>'
    // Unique categories A = 5, B = 3: column 0 is point 0, so it is red and its label override applies.
    const unique = await native(chart(data(0, [5, 3], ['A', 'B']), column(0, '<cx:aggregation />', red + labels)))
    const s = unique.groups[0].series[0]
    expect(s.dPt?.get(0)?.color).toBe('FF0000')
    expect(s.dPt?.has(1)).toBe(false)
    expect(s.labels?.points?.[0]).toMatchObject({ showVal: false })
    const ctx = createCanvas(480, 320).getContext('2d')
    const fills: string[] = []
    const fillRect = ctx.fillRect.bind(ctx)
    ctx.fillRect = (x: number, y: number, w: number, h: number) => { if (w < 400 && h > 0) fills.push(String(ctx.fillStyle).toLowerCase()); fillRect(x, y, w, h) }
    paintXlsxChart(unique, ctx as unknown as CanvasRenderingContext2D, 480, 320)
    expect(fills).toContain('#ff0000')
    expect(fills).toContain(`#${s.color.toLowerCase()}`)
    // A Pareto moves the override with its column: A = 3, B = 5 sorts to B, A, so point 0 (A) is column 1.
    const sorted = await native(chart(data(0, [3, 5], ['A', 'B']), column(0, '<cx:aggregation />', red) + '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2" /></cx:series>'))
    expect(sorted.categories).toEqual(['B', 'A'])
    expect(sorted.groups[0].series[0].dPt?.get(1)?.color).toBe('FF0000')
    expect(sorted.groups[0].series[0].dPt?.has(0)).toBe(false)
    // Grouped categories (A twice) and bins: no column is one source point, so the overrides are dropped and diagnosed.
    for (const xml of [chart(data(0, [1, 2, 3], ['A', 'A', 'B']), column(0, '<cx:aggregation />', red)), chart(data(0, [1, 2, 3]), column(0, '<cx:binning><cx:binSize val="1" /></cx:binning>', red))]) {
      const model = await native(xml)
      expect(model.groups[0].series[0].dPt).toBeUndefined()
      expect(model.diagnostics?.map(d => d.code)).toContain('chart-histogram-point-overrides')
    }
  })
})
