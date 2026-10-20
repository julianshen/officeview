/**
 * Task B5: the treemap and sunburst ChartEx families (docs/charts/chartex-map.md §7).
 *
 * Native evidence: TM (treemap/treemap.xlsx) and SB (sunburst/sunburst.xlsx) in tests/fixtures/chartex, Office-authored
 * inputs (PROVENANCE.md). Their parts carry no derived tree and their fallback is a text box, so every hierarchy rule
 * asserted here is a project rule [P] (map §7), checked against hand-computed trees. Inline parts, level tables and
 * workbooks are Office-shaped synthetic, never Office evidence.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContext, documentChartContext } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { CHARTEX_PAINTED, isChartExGroup, parseChartEx, parseChartExPart } from '../src/charts/chartex/parse'
import { CHART_HIERARCHY_LEVEL_LIMIT, CHART_HIERARCHY_NODE_LIMIT, buildHierarchy, type ChartExHierarchy } from '../src/charts/chartex/hierarchy'
import { treemapLayout, treemapPaintGeometry } from '../src/charts/chartex/treemap'
import { sunburstLayout, sunburstPaintGeometry } from '../src/charts/chartex/sunburst'
import { fillWorkbookCharts } from '../src/charts/fill'
import { layoutChart } from '../src/charts/layout'
import { paintXlsxChart } from '../src/charts/paint'
import { releaseChartModel } from '../src/charts/own'
import { PointWriter } from '../src/charts/points'
import { createChartWorkspace } from '../src/charts/workspace'
import { prepareChartAxes } from '../src/charts/prepare'
import type { ChartExGroup, ChartModel, ChartPoints } from '../src/charts/model'
import type { ChartParseContext } from '../src/charts/context'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { buildChartPptx } from './helpers/chart-fixtures'
import { describeFamilyConformance } from './helpers/chart-family-conformance'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'chartex', name)))
const theme = parseThemeContext()
const TOO_MANY = 'This chart has too many data points to summarize'
const W = 480, H = 360

/** The derived tree of map §7 for TM/SB (pre-order, source order): [path, depth, weight]. */
const TM_TREE: Array<[string, number, number]> = [
  ['Best', 0, 27], ['Best/First', 1, 16], ['Best/First/A', 2, 3], ['Best/First/B', 2, 4], ['Best/First/C', 2, 2], ['Best/First/D', 2, 3], ['Best/First/E', 2, 4],
  ['Best/Second', 1, 11], ['Best/Second/C', 2, 5], ['Best/Second/D', 2, 6],
  ['Worst', 0, 16], ['Worst/Third', 1, 4], ['Worst/Third/E', 2, 4], ['Worst/Fourth', 1, 3], ['Worst/Fifth', 1, 4], ['Worst/Sixth', 1, 5], ['Worst/Sixth/F', 2, 5],
]
/** A tree as [full path, depth, weight] rows in pre-order. */
function flat(tree: ChartExHierarchy): Array<[string, number, number]> {
  const out: Array<[string, number, number]> = [], paths: string[] = []
  for (let i = 0; i < tree.count; i++) {
    paths[i] = tree.parent[i] < 0 ? tree.labels[i] : `${paths[tree.parent[i]]}/${tree.labels[i]}`
    out.push([paths[i], tree.depth[i], tree.weight[i]])
  }
  return out
}
const indexOf = (tree: ChartExHierarchy, path: string) => flat(tree).findIndex(row => row[0] === path)
const group = (model: ChartModel): ChartExGroup => {
  expect(model.unsupported, 'native').toBeUndefined()
  const g = model.groups[0]
  if (!g || !isChartExGroup(g)) throw new Error('not a ChartEx group')
  return g
}
const treeOf = (model: ChartModel): ChartExHierarchy => {
  group(model)
  const tree = model.preparedAxes?.hierarchy
  if (!tree) throw new Error('no derived hierarchy')
  return tree
}
/** Weights as compact points: a number, `undefined` (blank), an error code string, or 'unavailable'. */
function weights(items: Array<number | string | undefined>): ChartPoints {
  const out = new PointWriter(items.length)
  items.forEach((v, i) => { if (typeof v === 'number') out.value(i, v); else if (v === 'unavailable') out.unavailable(i); else if (typeof v === 'string') out.error(i, v) })
  return out.done()
}
/** Level columns (root first) from rows of cells; '' is a blank cell. */
const columns = (rows: string[][]): string[][] => rows[0].map((_, j) => rows.map(r => r[j] ?? ''))
const build = (rows: string[][], w: Array<number | string | undefined>) => buildHierarchy(columns(rows), weights(w))
const tree = (rows: string[][], w: Array<number | string | undefined>) => {
  const out = build(rows, w)
  if ('refused' in out) throw new Error(out.refused)
  return out
}

/** The resident XLSX path: the native workbook parsed, its chart filled from the workbook's cells. */
async function resident(file: string) {
  const doc = await parseXlsx(await OfficePackage.load(fixture(file)))
  const sheet = doc.sheets[0], drawing = sheet.drawings?.[0]
  const content = drawing?.content
  if (content?.kind !== 'xlsxChart') throw new Error(`${file}: no chart (${JSON.stringify(drawing?.source)})`)
  return { doc, sheet, chart: content.chart }
}
const opened: Array<{ release(): Promise<void> }> = []
afterEach(async () => { for (const source of opened.splice(0)) await source.release() })
/** The paged XLSX path (the worker's paged reader). */
async function paged(file: string): Promise<ChartModel> {
  const session = await openImportSession({ kind: 'blob', blob: new Blob([fixture(file)]) }, { docId: `b5-${file}`, profile: 'D' } as never)
  expect((await session.done).kind).toBe('complete')
  const source = createPagedWorkbookSource(session)
  opened.push(source)
  const store = source.store('1') as unknown as { readDrawings: (o: object, p: 'D') => Promise<{ completeness: string; drawings: Array<{ content?: { kind: string; chart?: ChartModel } }> }>; stamp: object }
  const result = await store.readDrawings({ stamp: store.stamp, signal: new AbortController().signal }, 'D')
  const chart = result.drawings[0]?.content?.chart
  if (!chart) throw new Error('paged: no chart')
  return chart
}
/** The DOCX/PPTX embedded-workbook path: TM's own part, given `cx:externalData`, embedded with TM's own workbook. */
async function embedded(transform: (xml: string) => string = xml => xml) {
  const part = new TextDecoder().decode(await (await JSZip.loadAsync(fixture('treemap/treemap.xlsx'))).file('xl/charts/chartEx1.xml')!.async('uint8array'))
  const chartXml = transform(part.replace('<cx:chartData>', '<cx:chartData><cx:externalData r:id="rId1" cx:autoUpdate="0" />'))
  const built = buildChartPptx([{ chartXml, workbook: fixture('treemap/treemap.xlsx') }])
  const pkg = await OfficePackage.load(built.bytes), doc = documentChartContext('desktop'), context = chartParseContext(doc, pkg, built.chartPaths[0])
  const read = await readChartPart(pkg, built.chartPaths[0], context)
  if (read?.kind !== 'tree') throw new Error('read')
  const result = await parseChartExPart(read.root, theme, context)
  return { result, release: read.release, doc }
}
/** An Office-shaped synthetic ChartEx part in a minimal package at `path`, adapted through the production reader. */
async function synthetic(body: string, path = 'ppt/charts/chartEx1.xml') {
  const zip = new JSZip()
  zip.file(path, `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</cx:chartSpace>`)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const doc = documentChartContext('desktop'), context = chartParseContext(doc, pkg, path)
  const read = await readChartPart(pkg, path, context)
  if (read?.kind !== 'tree') throw new Error('read')
  return { root: read.root, context, release: read.release, doc }
}
const hierarchyPart = (layout: 'treemap' | 'sunburst', data: string, layoutPr = '') =>
  `<cx:chartData><cx:data id="0">${data}</cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="1" value="0" /></cx:dataLabels><cx:dataId val="0" />${layoutPr}</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>`
const cachedLevel = (type: 'cat' | 'size', points: Array<string | number>) => {
  const dim = type === 'cat' ? 'strDim' : 'numDim'
  return `<cx:${dim} type="${type}"><cx:lvl ptCount="${points.length}">${points.map((p, i) => `<cx:pt idx="${i}">${p}</cx:pt>`).join('')}</cx:lvl></cx:${dim}>`
}

const measureCtx = createCanvas(W, H).getContext('2d')
const measure = (m: ChartModel) => ({ width: (text: string, css: string) => { measureCtx.font = css; return measureCtx.measureText(text).width }, lineHeight: () => m.fontSizePt * 96 / 72 })
/** The area the paint uses (same deterministic layout: size, dpr 1, font generation 0, default family). */
const areaOf = (m: ChartModel) => layoutChart(m, W, H, 1, 0, measure(m)).area
/** Paints and records the drawn text; returns the context for pixel reads. */
function record(m: ChartModel) {
  const ctx = createCanvas(W, H).getContext('2d') as any
  const texts: Array<{ text: string; x: number; y: number }> = []
  const fillText = ctx.fillText.bind(ctx)
  ctx.fillText = (text: string, x: number, y: number, ...rest: unknown[]) => { texts.push({ text, x, y }); fillText(text, x, y, ...rest) }
  paintXlsxChart(m, ctx, W, H)
  return { texts, ctx }
}
const pixel = (ctx: any, x: number, y: number) => { const d = ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data; return [d[0], d[1], d[2]].map((v: number) => v.toString(16).padStart(2, '0')).join('').toUpperCase() }
const withParentLayout = (m: ChartModel, value: 'none' | 'banner' | 'overlapping'): ChartModel => {
  const g = group(m)
  return { ...m, groups: [{ ...g, chartEx: { ...g.chartEx, series: g.chartEx.series.map(s => ({ ...s, layoutPr: { ...s.layoutPr, parentLabelLayout: value } })) } } as ChartExGroup] }
}

describe('treemap and sunburst (chartex-map §7)', () => {
  it('full-path identity; duplicates aggregate; missing level is a leaf', async () => {
    expect([...CHARTEX_PAINTED].sort()).toEqual(expect.arrayContaining(['sunburst', 'treemap']))
    // TM and SB natively (resident XLSX): every level of the 3-column `_xlchart.v1.0` range (B4:D14) is read, root
    // leftmost. "C" under Best/First and "C" under Best/Second are distinct nodes; Fourth and Fifth (D blank) are
    // depth-1 leaves; 17 nodes, total 43.
    for (const file of ['treemap/treemap.xlsx', 'sunburst/sunburst.xlsx']) {
      const { chart, sheet } = await resident(file)
      const t = treeOf(chart)
      expect(flat(t), file).toEqual(TM_TREE)
      expect(t.count).toBe(17)
      expect(t.total).toBe(43)
      expect(group(chart).family).toBe(file.startsWith('treemap') ? 'treemap' : 'sunburst')
      // Native representation: the cx Choice is selected and drawn as a chart, not its fallback text box.
      expect(sheet.drawingCoverage?.find(e => e.scope === 'original'), file).toMatchObject({ feature: 'chart', status: 'native', selectedRepresentation: 'chart' })
    }
    // The paged XLSX path reads every level too (its planned lines include each column of the range).
    expect(flat(treeOf(await paged('treemap/treemap.xlsx')))).toEqual(TM_TREE)
    // The DOCX/PPTX path: the embedded workbook's defined name resolves inside the bounded reader, every level read.
    const e = await embedded()
    expect(e.result.kind).toBe('native')
    if (e.result.kind === 'native') { expect(flat(treeOf(e.result.model))).toEqual(TM_TREE); releaseChartModel(e.result.model) }
    e.release()
    // A multi-level cache (order unconfirmed) is ignored: the referenced range is read instead.
    const twoLevels = `<cx:lvl ptCount="11"><cx:pt idx="0">A</cx:pt></cx:lvl><cx:lvl ptCount="11"><cx:pt idx="0">First</cx:pt></cx:lvl>`
    const m = await embedded(xml => xml.replace('<cx:f>_xlchart.v1.0</cx:f>', `<cx:f>_xlchart.v1.0</cx:f>${twoLevels}`))
    expect(m.result.kind).toBe('native')
    if (m.result.kind === 'native') { expect(flat(treeOf(m.result.model))).toEqual(TM_TREE); releaseChartModel(m.result.model) }
    m.release()
    // With no reference to fall back on, a multi-level cache is unsupported (its fallback is kept).
    const noRef = await synthetic(hierarchyPart('treemap', `<cx:strDim type="cat">${twoLevels}</cx:strDim>${cachedLevel('size', [1])}`))
    expect(await parseChartExPart(noRef.root, theme, noRef.context)).toMatchObject({ kind: 'unsupported', reason: 'multi-level cache order unconfirmed' })
    noRef.release()

    // Duplicate full paths are summed (P/Q 1 + 2); a parent is the sum of its descendants.
    expect(flat(tree([['P', 'Q'], ['P', 'Q'], ['P', 'R']], [1, 2, 4]).tree)).toEqual([['P', 0, 7], ['P/Q', 1, 3], ['P/R', 1, 4]])
    // Repeated identical ancestor: Best/First/A then Best/./B keeps the inherited child (Best/First/B).
    expect(flat(tree([['Best', 'First', 'A'], ['Best', '', 'B']], [1, 2]).tree)).toEqual([['Best', 0, 3], ['Best/First', 1, 3], ['Best/First/A', 2, 1], ['Best/First/B', 2, 2]])
    // Changed ancestor: X/Y/Z then W/./V clears Y; the blank level is an unnamed node "" (diagnosed).
    const changed = tree([['X', 'Y', 'Z'], ['W', '', 'V']], [1, 2])
    expect(flat(changed.tree)).toEqual([['X', 0, 1], ['X/Y', 1, 1], ['X/Y/Z', 2, 1], ['W', 0, 2], ['W/', 1, 2], ['W//V', 2, 2]])
    expect(changed.diagnostics.map(d => d.code)).toContain('chart-hierarchy-unnamed')
    // Same label at the same depth under different parents: distinct nodes (full-path identity).
    expect(flat(tree([['A', 'X'], ['B', 'X']], [1, 2]).tree)).toEqual([['A', 0, 1], ['A/X', 1, 1], ['B', 0, 2], ['B/X', 1, 2]])
    // A missing level is a leaf at the previous level; a path that is both a leaf and a prefix keeps its own weight
    // as an unnamed child "" (in source order).
    expect(flat(tree([['S', ''], ['S', 'T']], [2, 3]).tree)).toEqual([['S', 0, 5], ['S/', 1, 2], ['S/T', 1, 3]])
    // A row with no label at any level is excluded (diagnosed); later rows still carry forward.
    const blankRow = tree([['K', 'a'], ['', ''], ['', 'b']], [1, 5, 2])
    expect(flat(blankRow.tree)).toEqual([['K', 0, 3], ['K/a', 1, 1], ['K/b', 1, 2]])
    expect(blankRow.diagnostics.map(d => d.code)).toContain('chart-hierarchy-row-excluded')
  }, 60_000)

  it('non-positive weights excluded', async () => {
    // Zero, negative, blank and error weights exclude their row (diagnosed with the count); nodes exist only for
    // included rows, and parents sum what remains.
    const out = tree([['a'], ['b'], ['c'], ['d'], ['e'], ['f']], [3, 0, -2, undefined, '#N/A', 5])
    expect(flat(out.tree)).toEqual([['a', 0, 3], ['f', 0, 5]])
    expect(out.tree.total).toBe(8)
    expect(out.diagnostics.find(d => d.code === 'chart-hierarchy-weight-excluded')?.reason).toMatch(/^4 /)
    // An excluded row still moves the carry-forward path (its cells are part of the layout): G is cleared by H.
    expect(flat(tree([['G', 'a'], ['', 'b'], ['H', 'c']], [0, 2, 3]).tree)).toEqual([['G', 0, 2], ['G/b', 1, 2], ['H', 0, 3], ['H/c', 1, 3]])
    // An unavailable weight (a formula with no saved value) refuses the chart.
    expect(build([['a'], ['b']], [1, 'unavailable'])).toEqual({ refused: "This chart's data is not available" })
    // Nothing left: an empty (native) tree.
    const none = tree([['a'], ['b']], [0, -1])
    expect(none.tree.count).toBe(0)

    // End to end on the XLSX fill path (Office-shaped synthetic part and sheet): 0 and -1 drop out.
    const p = await synthetic(hierarchyPart('treemap', '<cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim><cx:numDim type="size"><cx:f>_xlchart.v1.1</cx:f></cx:numDim>'), 'xl/charts/chartEx1.xml')
    const result = await parseChartExPart(p.root, theme, p.context)
    expect(result.kind).toBe('native')
    if (result.kind !== 'native') return
    const cells = [['Top', 'x', 4], ['', 'y', 0], ['', 'z', -1], ['Low', 'w', 2]] as const
    const rows = cells.map((r, i) => ({ index: i, cells: r.flatMap((v, c) => v === '' ? [] : [{ row: i, col: c, value: v }]) }))
    const sheet = { name: 'Sheet1', workbookIndex: 0, rows, drawings: [{ content: { kind: 'xlsxChart', chart: result.model } }] }
    fillWorkbookCharts([sheet], { names: [{ name: '_xlchart.v1.0', source: 'Sheet1!$A$1:$B$4' }, { name: '_xlchart.v1.1', source: 'Sheet1!$C$1:$C$4' }], workspace: p.context.workspace })
    const filled = (sheet.drawings[0].content as { chart: ChartModel }).chart
    expect(flat(treeOf(filled))).toEqual([['Top', 0, 4], ['Top/x', 1, 4], ['Low', 0, 2], ['Low/w', 1, 2]])
    expect(filled.preparedAxes?.diagnostics.some(d => d.code === 'chart-hierarchy-weight-excluded')).toBe(true)
    releaseChartModel(filled)
    p.release()
  })

  it('squarified areas proportional to weight; parentLabelLayout none, banner and overlapping', async () => {
    // The squarified oracle of Bruls, Huizing and van Wijk: weights 6,6,4,3,2,2,1 in a 6 x 4 rectangle (hand-computed:
    // 6,6 stack in a 3-wide column; 4,3 share a 7/3-high strip; 2, 2 and 1 fill the rest column by column).
    const bruls = tree([['a'], ['b'], ['c'], ['d'], ['e'], ['f'], ['g']], [6, 6, 4, 3, 2, 2, 1]).tree
    const rects = treemapLayout(bruls, { x: 0, y: 0, w: 6, h: 4 }, 'none', 0)
    const expected = [[0, 0, 3, 2], [0, 2, 3, 2], [3, 0, 12 / 7, 7 / 3], [3 + 12 / 7, 0, 9 / 7, 7 / 3], [3, 7 / 3, 1.2, 5 / 3], [4.2, 7 / 3, 1.2, 5 / 3], [5.4, 7 / 3, 0.6, 5 / 3]]
    expected.forEach((r, i) => r.forEach((v, k) => expect(rects[i * 4 + k], `node ${i} [${k}]`).toBeCloseTo(v, 9)))

    // TM's tree: every node's area is its share of its parent's content area (the whole parent for none and
    // overlapping; the parent less its banner strip for banner); children stay inside it and never overlap.
    const { chart } = await resident('treemap/treemap.xlsx')
    const t = treeOf(chart)
    const box = { x: 10, y: 20, w: 400, h: 300 }, B = 18
    for (const layout of ['none', 'overlapping', 'banner'] as const) {
      const r = treemapLayout(t, box, layout, B)
      const area = (i: number) => r[i * 4 + 2] * r[i * 4 + 3]
      const content = (i: number) => {
        if (i < 0) return box
        const leaf = t.end[i] === i + 1, b = layout === 'banner' && !leaf ? Math.min(B, r[i * 4 + 3]) : 0
        return { x: r[i * 4], y: r[i * 4 + 1] + b, w: r[i * 4 + 2], h: r[i * 4 + 3] - b }
      }
      for (let i = 0; i < t.count; i++) {
        const p = t.parent[i], c = content(p), parentWeight = p < 0 ? t.total : t.weight[p]
        expect(area(i) / (c.w * c.h), `${layout} ${flat(t)[i][0]}`).toBeCloseTo(t.weight[i] / parentWeight, 9)
        expect(r[i * 4] >= c.x - 1e-9 && r[i * 4 + 1] >= c.y - 1e-9 && r[i * 4] + r[i * 4 + 2] <= c.x + c.w + 1e-9 && r[i * 4 + 1] + r[i * 4 + 3] <= c.y + c.h + 1e-9, `${layout} ${flat(t)[i][0]} inside`).toBe(true)
      }
      for (let i = 0; i < t.count; i++) for (let j = i + 1; j < t.count; j++) {
        if (t.parent[i] !== t.parent[j]) continue
        const ox = Math.min(r[i * 4] + r[i * 4 + 2], r[j * 4] + r[j * 4 + 2]) - Math.max(r[i * 4], r[j * 4])
        const oy = Math.min(r[i * 4 + 1] + r[i * 4 + 3], r[j * 4 + 1] + r[j * 4 + 3]) - Math.max(r[i * 4 + 1], r[j * 4 + 1])
        expect(ox <= 1e-9 || oy <= 1e-9, `${layout} siblings ${i} ${j} do not overlap`).toBe(true)
      }
    }

    // Painted natively: parent labels follow parentLabelLayout (TM writes an empty layoutPr; the default is
    // overlapping [U]); leaf labels are the category names (TM's dataLabels show categoryName).
    expect(group(chart).chartEx.series[0].layoutPr.parentLabelLayout).toBeUndefined()
    const parents = ['Best', 'First', 'Second', 'Worst', 'Third', 'Sixth']
    for (const layout of ['none', 'banner', 'overlapping', undefined] as const) {
      const m = layout ? withParentLayout(chart, layout) : chart
      const { texts, ctx } = record(m)
      const geometry = treemapPaintGeometry(m, areaOf(m))
      const drawn = new Set(texts.map(t => t.text))
      if (layout === 'none') for (const name of parents) expect(drawn.has(name), `none: ${name}`).toBe(false)
      else for (const name of parents) expect(drawn.has(name), `${layout ?? 'default'}: ${name}`).toBe(true)
      for (const name of ['A', 'B', 'D', 'F', 'Fourth', 'Fifth']) expect(drawn.has(name), `${layout ?? 'default'} leaf ${name}`).toBe(true)
      // Each parent label sits inside its own rectangle; with banner, inside the banner strip above the children.
      for (const name of layout === 'none' ? [] : parents) {
        const i = flat(t).findIndex((row, k) => row[0].split('/').at(-1) === name && t.end[k] > k + 1)
        const r = geometry.rects, label = texts.find(x => x.text === name)!
        expect(label.x >= r[i * 4] && label.x <= r[i * 4] + r[i * 4 + 2], `${layout} ${name} x`).toBe(true)
        const top = r[i * 4 + 1], bottom = layout === 'banner' ? top + geometry.banner : top + r[i * 4 + 3]
        expect(label.y >= top && label.y <= bottom, `${layout} ${name} y in [${top}, ${bottom}]`).toBe(true)
      }
      // Distinct native geometry: each root branch is filled with its own palette colour (sampled near the bottom-left
      // corner of a leaf, away from labels and borders).
      const palette = group(m).chartEx.palette!
      expect(palette[0]).not.toBe(palette[1])
      for (const [path, colour] of [['Best/First/A', palette[0]], ['Best/Second/D', palette[0]], ['Worst/Sixth/F', palette[1]], ['Worst/Fourth', palette[1]]] as const) {
        const i = indexOf(t, path), r = geometry.rects
        expect(pixel(ctx, r[i * 4] + 4, r[i * 4 + 1] + r[i * 4 + 3] - 4), `${layout ?? 'default'} ${path}`).toBe(colour)
      }
    }
  }, 60_000)

  it('sunburst rings by depth with angles proportional to subtree weight', async () => {
    const { chart, sheet } = await resident('sunburst/sunburst.xlsx')
    expect(sheet.drawingCoverage?.find(e => e.scope === 'original')).toMatchObject({ feature: 'chart', status: 'native', selectedRepresentation: 'chart' })
    const t = treeOf(chart), R = 90
    const s = sunburstLayout(t, R)
    const at = (path: string) => { const i = indexOf(t, path); return { start: s[i * 4], sweep: s[i * 4 + 1], r0: s[i * 4 + 2], r1: s[i * 4 + 3] } }
    const TAU = Math.PI * 2, top = -Math.PI / 2
    // The root level is the inner ring; each level one ring outward (3 levels: R/3 each). Angles are proportional
    // to subtree weight, children in source order from their parent's start, clockwise from 12 o'clock.
    const expected: Record<string, [number, number, number, number]> = {
      Best: [top, TAU * 27 / 43, 0, R / 3], Worst: [top + TAU * 27 / 43, TAU * 16 / 43, 0, R / 3],
      'Best/First': [top, TAU * 16 / 43, R / 3, 2 * R / 3], 'Best/Second': [top + TAU * 16 / 43, TAU * 11 / 43, R / 3, 2 * R / 3],
      'Best/First/A': [top, TAU * 3 / 43, 2 * R / 3, R], 'Best/First/B': [top + TAU * 3 / 43, TAU * 4 / 43, 2 * R / 3, R],
      'Best/Second/D': [top + TAU * 21 / 43, TAU * 6 / 43, 2 * R / 3, R],
      'Worst/Third': [top + TAU * 27 / 43, TAU * 4 / 43, R / 3, 2 * R / 3], 'Worst/Fourth': [top + TAU * 31 / 43, TAU * 3 / 43, R / 3, 2 * R / 3],
      'Worst/Sixth/F': [top + TAU * 38 / 43, TAU * 5 / 43, 2 * R / 3, R],
    }
    for (const [path, [start, sweep, r0, r1]] of Object.entries(expected)) {
      const got = at(path)
      expect(got.start, `${path} start`).toBeCloseTo(start, 9)
      expect(got.sweep, `${path} sweep`).toBeCloseTo(sweep, 9)
      expect(got.r0, `${path} r0`).toBeCloseTo(r0, 9)
      expect(got.r1, `${path} r1`).toBeCloseTo(r1, 9)
    }
    // Children exactly partition their parent's angle.
    for (let i = 0; i < t.count; i++) {
      if (t.end[i] === i + 1) continue
      let sum = 0
      for (let c = i + 1; c < t.end[i]; c = t.end[c]) sum += s[c * 4 + 1]
      expect(sum, flat(t)[i][0]).toBeCloseTo(s[i * 4 + 1], 9)
    }
    // Painted natively: ring segments of each root branch in its own palette colour (sampled a fifth into each
    // segment's angle, away from its centred label).
    const { ctx, texts } = record(chart)
    const g = sunburstPaintGeometry(chart, areaOf(chart))
    const palette = group(chart).chartEx.palette!
    expect(palette[0]).not.toBe(palette[1])
    for (const [path, ring, colour] of [['Best', 0.75, palette[0]], ['Worst', 0.75, palette[1]], ['Best/Second/D', 0.85, palette[0]], ['Worst/Sixth/F', 0.85, palette[1]]] as const) {
      const i = indexOf(t, path), mid = g.segments[i * 4] + g.segments[i * 4 + 1] / 5, r = g.segments[i * 4 + 2] + (g.segments[i * 4 + 3] - g.segments[i * 4 + 2]) * ring
      expect(pixel(ctx, g.cx + Math.cos(mid) * r, g.cy + Math.sin(mid) * r), path).toBe(colour)
    }
    expect(texts.map(x => x.text)).toEqual(expect.arrayContaining(['Best', 'Worst']))
  }, 60_000)

  it('over 4096 nodes refuses', async () => {
    expect(CHART_HIERARCHY_NODE_LIMIT).toBe(4096)
    const names = (n: number, prefix = 'n') => Array.from({ length: n }, (_, i) => [`${prefix}${i}`])
    // Single level: 4,096 distinct leaves are drawn; 4,097 refuse (never a prefix of the data).
    expect(tree(names(4096), Array(4096).fill(1)).tree.count).toBe(4096)
    expect(build(names(4097), Array(4097).fill(1))).toEqual({ refused: TOO_MANY })
    // Internal nodes count: 64 parents x 63 children = 4,096 nodes; one more child refuses.
    const twoLevel = (extra: number) => {
      const rows: string[][] = []
      for (let p = 0; p < 64; p++) for (let c = 0; c < 63 + (p === 0 ? extra : 0); c++) rows.push([`p${p}`, `c${c}`])
      return rows
    }
    expect(tree(twoLevel(0), Array(64 * 63).fill(1)).tree.count).toBe(4096)
    expect(build(twoLevel(1), Array(64 * 63 + 1).fill(1))).toEqual({ refused: TOO_MANY })
    // Unnamed nodes count: 4,095 leaves, then n0 becomes a prefix (child c plus its own "" child): 4,097 refuse.
    expect(build([...names(4095).map(r => [r[0], '']), ['n0', 'c']], Array(4096).fill(1))).toEqual({ refused: TOO_MANY })

    // On the cache path (an Office-shaped synthetic PPTX part with a single cached level): refused, not drawn.
    const labels = Array.from({ length: 4097 }, (_, i) => `n${i}`)
    const cached = await synthetic(hierarchyPart('treemap', `${cachedLevel('cat', labels)}${cachedLevel('size', labels.map(() => 1))}`))
    expect(await parseChartExPart(cached.root, theme, cached.context)).toMatchObject({ kind: 'refused', reason: TOO_MANY, feature: 'ChartEx' })
    cached.release()
    expect(cached.doc.workspace.used).toBe(0)
    const fits = await synthetic(hierarchyPart('sunburst', `${cachedLevel('cat', labels.slice(1))}${cachedLevel('size', labels.slice(1).map(() => 1))}`))
    const ok = await parseChartExPart(fits.root, theme, fits.context)
    expect(ok.kind).toBe('native')
    if (ok.kind === 'native') { expect(treeOf(ok.model).count).toBe(4096); releaseChartModel(ok.model) }
    fits.release()

    // On the XLSX fill path: 4,097 rows of distinct names refuse the chart.
    const p = await synthetic(hierarchyPart('treemap', '<cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$4097</cx:f></cx:strDim><cx:numDim type="size"><cx:f>Sheet1!$B$1:$B$4097</cx:f></cx:numDim>'), 'xl/charts/chartEx1.xml')
    const model = parseChartEx(p.root, theme, p.context)
    const sheet = { name: 'Sheet1', workbookIndex: 0, rows: labels.map((name, i) => ({ index: i, cells: [{ row: i, col: 0, value: name }, { row: i, col: 1, value: 1 }] })), drawings: [{ content: { kind: 'xlsxChart', chart: model } }] }
    fillWorkbookCharts([sheet], { workspace: p.context.workspace })
    expect((sheet.drawings[0].content as { chart: ChartModel }).chart).toMatchObject({ unsupported: TOO_MANY, refused: true })
    p.release()

    // Acquisition bounds before any read: a level over 65,536 points, or more than the level bound, refuses.
    const ctx = (): ChartParseContext => ({ partPath: 'xl/charts/chartEx1.xml', profile: 'desktop', workspace: createChartWorkspace('desktop'), pkg: {} as ChartParseContext['pkg'] })
    const column = (n: number) => { let out = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out }
    for (const range of ['Sheet1!$A$1:$B$65537', `Sheet1!$A$1:$${column(CHART_HIERARCHY_LEVEL_LIMIT + 1)}$3`]) {
      const c = ctx()
      const over = parseChartEx(parseXmlOrdered(`<chartSpace>${hierarchyPart('treemap', `<strDim type="cat"><f>${range}</f></strDim><numDim type="size"><f>Sheet1!$Z$1:$Z$3</f></numDim>`).replace(/cx:/g, '')}</chartSpace>`), undefined, c)
      const s = { name: 'Sheet1', workbookIndex: 0, rows: [], drawings: [{ content: { kind: 'xlsxChart', chart: over } }] }
      fillWorkbookCharts([s], { workspace: c.workspace })
      expect((s.drawings[0].content as { chart: ChartModel }).chart, range).toMatchObject({ unsupported: TOO_MANY, refused: true })
    }
  }, 120_000)
})

describe('treemap and sunburst after the feature merge (B5 review minors, A6b)', () => {
  it('value labels use the authored label number format; a label-less row is excluded before its weight is read; no error bars', async () => {
    // Office-shaped synthetic: values shown with numFmt "0.00" (formatted at stage 1, drawn as such).
    const part = hierarchyPart('treemap', `${cachedLevel('cat', ['a', 'b'])}${cachedLevel('size', [3, 4.5])}`)
      .replace('<cx:visibility seriesName="0" categoryName="1" value="0" />', '<cx:numFmt formatCode="0.00" sourceLinked="0" /><cx:visibility seriesName="0" categoryName="0" value="1" />')
    const p = await synthetic(part)
    const result = await parseChartExPart(p.root, theme, p.context)
    expect(result.kind).toBe('native')
    if (result.kind !== 'native') return
    expect(treeOf(result.model).values).toEqual(['3.00', '4.50'])
    expect(record(result.model).texts.map(t => t.text)).toEqual(expect.arrayContaining(['3.00', '4.50']))
    // A6b: treemap and sunburst never draw error bars (they are outside the drawable families): dropped, diagnosed.
    for (const family of ['treemap', 'sunburst'] as const) {
      const bars = { dir: 'y', type: 'both', valType: 'fixedVal', val: 1, noEndCap: false } as const
      const m: ChartModel = { ...result.model, preparedAxes: undefined, groups: result.model.groups.map(g => ({ ...g, family, series: g.series.map(s => ({ ...s, errorBars: [bars] })) })) }
      const prepared = prepareChartAxes(m)
      expect(prepared.preparedAxes!.errorBars ?? [], family).toHaveLength(0)
      expect(prepared.preparedAxes!.diagnostics.map(d => d.code), family).toContain('chart-error-bars-dropped')
      expect(prepared.preparedAxes!.hierarchy?.count, family).toBe(2)
    }
    releaseChartModel(result.model)
    p.release()
    // Map §7 order: a row with no label at any level is excluded (step 1) even when its weight is unavailable (step 6).
    const out = build([['a'], ['']], [1, 'unavailable'])
    if ('refused' in out) throw new Error(out.refused)
    expect(flat(out.tree)).toEqual([['a', 0, 1]])
    expect(out.diagnostics.map(d => d.code)).toContain('chart-hierarchy-row-excluded')
  })
})

/** The conformance harness (spec 7), with the numeric case as the size dimension and generated category names. */
const conformance = (layout: 'treemap' | 'sunburst') => ({
  name: layout,
  chart: (cache: string | null) => {
    if (cache === null) return `<cx:chartSpace xmlns:cx="${CX}"><cx:chartData /><cx:chart><cx:plotArea><cx:plotAreaRegion /></cx:plotArea></cx:chart></cx:chartSpace>`
    const count = Number(/<c:ptCount val="(\d+)"/.exec(cache)?.[1] ?? 0)
    const points = [...cache.matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v><\/c:pt>/g)].map(m => `<cx:pt idx="${m[1]}">${m[2]}</cx:pt>`).join('')
    const cats = Array.from({ length: count }, (_, i) => `<cx:pt idx="${i}">C${i}</cx:pt>`).join('')
    return `<cx:chartSpace xmlns:cx="${CX}"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="${count}">${cats}</cx:lvl></cx:strDim><cx:numDim type="size"><cx:lvl ptCount="${count}">${points}</cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="1" value="1" /></cx:dataLabels><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`
  },
  expect: {
    'no series': { kind: 'card' as const },
    'empty series': { kind: 'native' as const }, 'all blank': { kind: 'native' as const, diagnostics: ['chart-hierarchy-weight-excluded'] },
    'NaN and infinity': { kind: 'native' as const, diagnostics: ['chart-hierarchy-weight-excluded'] }, 'negative values': { kind: 'native' as const, diagnostics: ['chart-hierarchy-weight-excluded'] },
    'single point': { kind: 'native' as const },
    // An analytic family refuses an unavailable weight (map §2.4): a card, never a partial tree.
    'unavailable point': { kind: 'card' as const },
  },
})
describeFamilyConformance(conformance('treemap'))
describeFamilyConformance(conformance('sunburst'))
