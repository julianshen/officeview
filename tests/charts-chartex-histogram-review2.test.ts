/**
 * Task B3 review round 2 (codex re-review): the AlternateContent branch is re-selected as a unit and its fallback
 * descriptor lives until the data is final, paged pending sheets included (R1); ordinary arithmetic is kept where it
 * is representable (R2); override ambiguity is per category slot (R3); a selected picture fallback is indexed and
 * paints (R4).
 *
 * HI2 variants are derived from the native Office part (tests/fixtures/chartex); inline parts are Office-shaped
 * synthetic.
 */
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseThemeContext } from '../src/drawing/style'
import { chartParseContextFor } from '../src/charts/context'
import { readChartPart } from '../src/charts/read'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { pointList } from '../src/charts/points'
import type { ChartModel } from '../src/charts/model'
import { parseXlsx } from '../src/xlsx/parse'
import { parseWorksheetDrawings, settleChartFallbacks } from '../src/xlsx/drawing'
import type { XlsxSheet } from '../src/xlsx/types'
import { getPaintables } from '../src/render/paint'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'

const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const theme = parseThemeContext()
const HI2 = join(__dirname, 'fixtures', 'chartex', 'tdf163727_histogram_underflow_overflow', 'tdf163727_histogram_underflow_overflow.xlsx')
const TINY = (xml: string) => xml.replace(/<cx:binSize val="2" ?\/>/, '<cx:binSize val="0.000001" />')
/** An HI2 variant: each named part rewritten (the edit must apply) or added. */
async function hi2(edits: Record<string, (xml: string) => string>, added: Record<string, string | Uint8Array> = {}): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(readFileSync(HI2))
  for (const [path, fn] of Object.entries(edits)) {
    const before = await zip.file(path)!.async('string'), after = fn(before)
    if (after === before) throw new Error(`${path}: the variant edit did not apply`)
    zip.file(path, after)
  }
  for (const [path, data] of Object.entries(added)) zip.file(path, data)
  return zip.generateAsync({ type: 'uint8array' })
}
const choice = (fn: (frame: string) => string) => (xml: string) => xml.replace(/(<mc:Choice[^>]*>)(<xdr:graphicFrame[\s\S]*?<\/xdr:graphicFrame>)(<\/mc:Choice>)/, (_, open, frame, close) => `${open}${fn(frame)}${close}`)
const SHAPE = '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="9" name="Beside" /><xdr:cNvSpPr /></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="914400" cy="914400" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></xdr:spPr></xdr:sp>'
async function adapt(xml: string) {
  const zip = new JSZip(), path = 'ppt/charts/chartEx1.xml'
  zip.file(path, `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</cx:chartSpace>`)
  const pkg = await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
  const read = await readChartPart(pkg, path, chartParseContextFor(pkg, path))
  if (read?.kind !== 'tree') throw new Error('read')
  try {
    const result = await parseChartPart(read.root, theme, chartParseContextFor(pkg, path))
    if (result.kind !== 'native') throw new Error(`${result.kind}: ${result.reason}`)
    return result.model
  } finally { read.release() }
}
const lvl = (values: Array<number | string>) => `<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`
const AXES = '<cx:axis id="0"><cx:catScaling gapWidth="0" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:tickLabels /></cx:axis><cx:axis id="2"><cx:valScaling max="1" min="0" /><cx:units unit="percentage" /><cx:tickLabels /></cx:axis>'
const part = (values: number[], layoutPr: string, categories?: string[], extra = '', pareto = false) =>
  `<cx:chartData><cx:data id="0">${categories ? `<cx:strDim type="cat">${lvl(categories)}</cx:strDim>` : ''}<cx:numDim type="val">${lvl(values)}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0" />${extra}<cx:layoutPr>${layoutPr}</cx:layoutPr><cx:axisId val="1" /></cx:series>${pareto ? '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2" /></cx:series>' : ''}</cx:plotAreaRegion>${AXES}</cx:plotArea></cx:chart>`
const nums = (model: ChartModel, group = 0) => pointList(model.groups[group].series[0].values).map(p => p.state === 'value' ? p.v : p.state)
const red = (idx: number) => `<cx:dataPt idx="${idx}"><cx:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></cx:spPr></cx:dataPt>`

describe('B3 review round 2', () => {
  it('R1: the selected Choice is re-selected as a unit (a shape beside the chart, a grouped chart)', async () => {
    const group = (frame: string) => `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="7" name="Group" /><xdr:cNvGrpSpPr /></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="4572000" cy="2743200" /><a:chOff x="0" y="0" /><a:chExt cx="4572000" cy="2743200" /></a:xfrm></xdr:grpSpPr>${frame.replace('<a:ext cx="0" cy="0"/>', '<a:ext cx="4572000" cy="2743200"/>')}</xdr:grpSp>`
    for (const [name, fn] of [['shape beside the chart', (frame: string) => frame + SHAPE], ['grouped chart', group]] as const) {
      // Control: with the native binSize the Choice is drawn (the chart is native).
      const ok = await parseXlsx(await OfficePackage.load(await hi2({ 'xl/drawings/drawing1.xml': choice(fn) })))
      expect(ok.sheets[0].drawings?.map(d => d.source.representation), name).toEqual(name === 'grouped chart' ? ['choice'] : ['choice', 'choice'])
      // Refused once its data is read: every object of the Choice gives way to the authored Fallback.
      const xlsx = await parseXlsx(await OfficePackage.load(await hi2({ 'xl/drawings/drawing1.xml': choice(fn), 'xl/charts/chartEx1.xml': TINY })))
      const sheet = xlsx.sheets[0]
      expect(sheet.drawings?.map(d => [d.source.element, d.source.representation]), name).toEqual([['sp', 'fallback']])
      expect(sheet.drawingCoverage?.filter(e => e.representation === 'choice'), `${name}: no rejected Choice object is kept`).toEqual([])
      expect(sheet.drawingCoverage?.find(e => e.scope === 'original'), name).toMatchObject({ feature: 'ChartEx', status: 'fallback', representation: 'fallback', reason: 'invalid binning' })
      expect((xlsx.drawingCoverage ?? []).filter(e => e.representation === 'choice'), name).toEqual([])
    }
  })

  it('R1: a paged chart waiting on an importing sheet keeps its fallback route; the completed read selects the fallback', async () => {
    // HI2's data moved to a second sheet "Data" (its defined name re-pointed), with the refused binSize.
    const bytes = await hi2({
      'xl/charts/chartEx1.xml': TINY,
      'xl/workbook.xml': xml => xml.replace('</sheets>', '<sheet name="Data" sheetId="2" r:id="rData"/></sheets>').replace(/Sheet1!\$A\$1:\$A\$5/, 'Data!$A$1:$A$5'),
      'xl/_rels/workbook.xml.rels': xml => xml.replace('</Relationships>', '<Relationship Id="rData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'),
    }, { 'xl/worksheets/sheet2.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${[12, 10, 11, 13, 15].map((v, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${v}</v></c></row>`).join('')}</sheetData></worksheet>` })
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'b3-r2', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    try {
      const adapter = source.store('1') as unknown as { stamp: unknown; readDrawings(request: object, profile: string): Promise<{ pendingSheets?: string[]; drawings: Array<{ source: { element: string; representation: string }; content?: { kind: string; chart?: ChartModel } }> }> }
      const read = () => adapter.readDrawings({ stamp: adapter.stamp, signal: new AbortController().signal }, 'D')
      const data = session.workbook.sheets[1], through = data.importedThroughRow
      // The data sheet is still importing: the chart waits (not refused), so its Choice stays selected.
      data.importedThroughRow = -1
      const first = await read()
      expect(first.pendingSheets).toContain('2')
      expect(first.drawings.map(d => [d.source.element, d.source.representation])).toEqual([['graphicFrame', 'choice']])
      expect(first.drawings[0].content?.chart?.unsupported).toBeUndefined()
      // Imported: the read again refuses the chart and selects the authored fallback.
      data.importedThroughRow = through
      const second = await read()
      expect(second.pendingSheets ?? []).toEqual([])
      expect(second.drawings.map(d => [d.source.element, d.source.representation])).toEqual([['sp', 'fallback']])
    } finally { await source.release() }
  })

  it('R5: nested AlternateContent: the first usable inner fallback wins, whenever the chart is refused', async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    const CX1 = 'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"'
    const shape = (name: string, id: number) => SHAPE.replace('id="9" name="Beside"', `id="${id}" name="${name}"`)
    const badPicture = '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="31" name="BROKEN" /><xdr:cNvPicPr /></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rMissing" /></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="914400" cy="914400" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></xdr:spPr></xdr:pic>'
    /** The native chart AlternateContent, its Choice extended by `beside` and its Fallback replaced by `inner`, wrapped in an
     * outer AlternateContent whose Choice also holds `sibling` and whose Fallback is OUTER. */
    const nest = (inner: string, beside = '', sibling = '') => (xml: string) => xml.replace(/<mc:AlternateContent[\s\S]*<\/mc:AlternateContent>/, ac => {
      const innerAc = ac.replace(/(<\/xdr:graphicFrame>)(<\/mc:Choice>)/, `$1${beside}$2`).replace(/<mc:Fallback>[\s\S]*<\/mc:Fallback>/, `<mc:Fallback>${inner}</mc:Fallback>`)
      return `<mc:AlternateContent ${MC}><mc:Choice ${CX1}>${innerAc}${sibling}</mc:Choice><mc:Fallback>${shape('OUTER', 40)}</mc:Fallback></mc:AlternateContent>`
    })
    const late = TINY, early = (xml: string) => xml.replace(/<cx:binSize val="2" ?\/>/, '<cx:binSize val="-1" />')
    const result = async (drawing: (xml: string) => string, binSize: (xml: string) => string) => {
      const xlsx = await parseXlsx(await OfficePackage.load(await hi2({ 'xl/drawings/drawing1.xml': drawing, 'xl/charts/chartEx1.xml': binSize })))
      const sheet = xlsx.sheets[0]
      return { drawings: sheet.drawings?.map(d => [d.source.name, d.source.representation]), chart: sheet.drawings?.some(d => d.content), coverage: sheet.drawingCoverage?.filter(e => e.representation === 'choice' && e.selectedRepresentation === 'chart').length }
    }
    const cases: Array<[string, (xml: string) => string, Array<[string, string]>]> = [
      // codex's flattened nested wrapper: INNER (inner fallback) is selected, never OUTER.
      ['flattened nested wrapper', nest(shape('INNER', 30)), [['INNER', 'fallback']]],
      // The inner replacement shrinks the branch (chart + shape -> INNER); the outer branch keeps its sibling.
      ['inner replacement shrinks the branch', nest(shape('INNER', 30), shape('BESIDE', 32), shape('SIBLING', 33)), [['INNER', 'fallback'], ['SIBLING', 'choice']]],
      // The inner replacement grows the branch (chart -> INNER-A, INNER-B).
      ['inner replacement grows the branch', nest(shape('INNER-A', 30) + shape('INNER-B', 34), '', shape('SIBLING', 33)), [['INNER-A', 'fallback'], ['INNER-B', 'fallback'], ['SIBLING', 'choice']]],
      // An unusable inner fallback (a picture without its image part): the enclosing OUTER fallback answers.
      ['unusable inner fallback', nest(badPicture), [['OUTER', 'fallback']]],
    ]
    for (const [name, drawing, expected] of cases) {
      const lateResult = await result(drawing, late), earlyResult = await result(drawing, early)
      expect(lateResult.drawings, `${name} (refused after the fill)`).toEqual(expected)
      // The same authored representation as a chart refused at selection (binSize -1, invalid at parse).
      expect(lateResult.drawings, `${name}: late = early`).toEqual(earlyResult.drawings)
      expect(lateResult.chart, name).toBe(false)
      expect(lateResult.coverage, `${name}: no native chart coverage left`).toBe(0)
    }
  })

  it('merge: the late-refusal fallback seam serves every ChartEx family (treemap)', async () => {
    // TM (native treemap) with size cell E4 a formula without a saved value (saved-values policy): an analytic
    // family with an unavailable required point is refused once the fill reads it, so its mc:Fallback is selected.
    const zip = await JSZip.loadAsync(readFileSync(join(__dirname, 'fixtures', 'chartex', 'treemap', 'treemap.xlsx')))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const edited = sheet.replace(/<c r="E4"([^>]*)><v>3<\/v><\/c>/, '<c r="E4"$1><f>1+2</f></c>')
    expect(edited).not.toBe(sheet)
    zip.file('xl/worksheets/sheet1.xml', edited)
    const refused = await parseXlsx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })), { valuePolicy: 'cached' } as never)
    const drawing = refused.sheets[0].drawings?.[0]
    expect(drawing?.content).toBeUndefined()
    expect(drawing?.source.representation).toBe('fallback')
    expect(refused.sheets[0].drawingCoverage?.some(e => e.selectedRepresentation === 'chart')).toBe(false)
    expect(refused.sheets[0].drawingCoverage?.find(e => e.scope === 'original')).toMatchObject({ feature: 'ChartEx', status: 'fallback', representation: 'fallback', reason: "This chart's data is not available" })
    // Control: the unmodified treemap is native (its Choice selected).
    const native = await parseXlsx(await OfficePackage.load(new Uint8Array(readFileSync(join(__dirname, 'fixtures', 'chartex', 'treemap', 'treemap.xlsx')))), { valuePolicy: 'cached' } as never)
    expect(native.sheets[0].drawings?.[0].content?.kind).toBe('xlsxChart')
    expect(native.sheets[0].drawingCoverage?.find(e => e.element === 'graphicFrame')).toMatchObject({ selectedRepresentation: 'chart' })
  })

  it('R6: a three-level nest with a refused outer sibling selects the OUTER fallback, late as early', async () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    const CX1 = 'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"'
    // codex's case: INNER inside MIDDLE inside OUTER; OUTER's Choice also holds a second copy of the chart frame.
    const triple = (xml: string) => xml.replace(/<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/, ac => {
      const fallback = ac.match(/<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/)![1]
      const named = (name: string) => fallback.replace(/name="[^"]*"/, `name="${name}"`)
      const frame = ac.match(/<xdr:graphicFrame[\s\S]*?<\/xdr:graphicFrame>/)![0]
      const inner = ac.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/, `<mc:Fallback>${named('INNER')}</mc:Fallback>`)
      const middle = `<mc:AlternateContent ${MC}><mc:Choice ${CX1}>${inner}</mc:Choice><mc:Fallback>${named('MIDDLE')}</mc:Fallback></mc:AlternateContent>`
      return `<mc:AlternateContent ${MC}><mc:Choice ${CX1}>${middle}${frame}</mc:Choice><mc:Fallback>${named('OUTER')}</mc:Fallback></mc:AlternateContent>`
    })
    const run = async (binSize: (xml: string) => string) => {
      const xlsx = await parseXlsx(await OfficePackage.load(await hi2({ 'xl/drawings/drawing1.xml': triple, 'xl/charts/chartEx1.xml': binSize })))
      return { drawings: xlsx.sheets[0].drawings?.map(d => [d.source.name, d.source.representation]), choice: xlsx.sheets[0].drawingCoverage?.filter(e => e.representation === 'choice').length }
    }
    const late = await run(TINY), early = await run(xml => xml.replace(/<cx:binSize val="2" ?\/>/, '<cx:binSize val="-1" />'))
    // OUTER's Choice holds a refused chart (the sibling frame), so the whole outer branch gives way.
    expect(late.drawings).toEqual([['OUTER', 'fallback']])
    expect(late.drawings).toEqual(early.drawings)
    expect(late.choice).toBe(0)
  })

  it('R7: settlement returns promptly on a chart still waiting on data, and keeps its fallback route', async () => {
    // HI2's chart parsed alone (no fill): its defined-name reference is still pending.
    const pkg = await OfficePackage.load(new Uint8Array(readFileSync(HI2)))
    const sheet = { name: 'Sheet1', rows: [], cols: [], merges: [], mergeRanges: [] } as unknown as XlsxSheet
    await parseWorksheetDrawings(pkg, sheet, 'xl/worksheets/sheet1.xml', (await pkg.xml('xl/worksheets/sheet1.xml'))!, theme)
    const content = sheet.drawings![0].content
    if (content?.kind !== 'xlsxChart') throw new Error('chart')
    expect(content.chart.groups[0].series[0].pending?.values).toBeDefined()
    // Returns (bounded), nothing replaced: the chart is waiting, not refused.
    const timeout = new Promise<'timeout'>(resolve => setTimeout(() => resolve('timeout'), 2000))
    expect(await Promise.race([settleChartFallbacks(sheet), timeout])).toBe(false)
    expect(sheet.drawings!.map(d => d.source.representation)).toEqual(['choice'])
    // Its route stays: refused later (as the fill would), the next settlement selects the authored fallback.
    content.chart = { ...content.chart, unsupported: 'invalid binning', refused: true }
    expect(await Promise.race([settleChartFallbacks(sheet), timeout])).toBe(true)
    expect(sheet.drawings!.map(d => [d.source.element, d.source.representation])).toEqual([['sp', 'fallback']])
  })

  it('R2: ordinary arithmetic is kept where it is representable', async () => {
    // Distinct categories of very different magnitude: A = 1e308, B = 5e-324 (both sums representable).
    expect(nums(await adapt(part([1e308, 5e-324], '<cx:aggregation />', ['A', 'B'])))).toEqual([1e308, 5e-324])
    // A subnormal width: [0, 5e-324, 1e-323], binCount 2: h = 5e-324, edges 0, 5e-324, 1e-323; r: [0, 5e-324] {0, 5e-324},
    // (5e-324, 1e-323] {1e-323}.
    const tiny = await adapt(part([0, 5e-324, 1e-323], '<cx:binning><cx:binCount val="2" /></cx:binning>'))
    expect(nums(tiny)).toEqual([2, 1])
    expect(tiny.categories).toHaveLength(2)
    expect(tiny.categories[0]).toBe("[0, 5E-324]")
    // The large-number cases of round 1 still hold.
    expect(nums(await adapt(part([1e308, 1e308, -1e308], '<cx:aggregation />', ['A', 'A', 'A'])))).toEqual([1e308])
    expect(nums(await adapt(part([1e308, 1e308], '<cx:aggregation />', ['A', 'B'], '', true)), 1)).toEqual([0.5, 1])
  })

  it('R3: override ambiguity is per category slot; unique slots keep theirs, through the Pareto order', async () => {
    // A, A, B (1, 2, 3) with point 2 (B) red: B is one source point, so its column keeps the override; nothing dropped.
    const kept = await adapt(part([1, 2, 3], '<cx:aggregation />', ['A', 'A', 'B'], red(2)))
    expect(kept.categories).toEqual(['A', 'B'])
    expect(kept.groups[0].series[0].dPt?.get(1)?.color).toBe('FF0000')
    expect(kept.diagnostics?.some(d => d.code === 'chart-histogram-point-overrides') ?? false).toBe(false)
    // Point 0 (A, grouped from two points) is dropped and diagnosed; B's override is still kept.
    const mixed = await adapt(part([1, 2, 3], '<cx:aggregation />', ['A', 'A', 'B'], red(0) + red(2)))
    expect(mixed.groups[0].series[0].dPt?.has(0)).toBe(false)
    expect(mixed.groups[0].series[0].dPt?.get(1)?.color).toBe('FF0000')
    expect(mixed.diagnostics?.find(d => d.code === 'chart-histogram-point-overrides')?.reason).toMatch(/^1 point override is not drawn/)
    // Pareto: A, A, B (1, 1, 5) sorts to B 5, A 2; B's override moves to column 0.
    const pareto = await adapt(part([1, 1, 5], '<cx:aggregation />', ['A', 'A', 'B'], red(2), true))
    expect(pareto.categories).toEqual(['B', 'A'])
    expect(pareto.groups[0].series[0].dPt?.get(0)?.color).toBe('FF0000')
    expect(pareto.groups[0].series[0].dPt?.has(1)).toBe(false)
  })

  it('R4: a picture fallback selected after the fill is indexed and paints', async () => {
    const png = createCanvas(8, 8)
    const pctx = png.getContext('2d')
    pctx.fillStyle = '#00ff00'
    pctx.fillRect(0, 0, 8, 8)
    const picture = '<mc:Fallback><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1" /><xdr:cNvPicPr /></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rImg" /><a:stretch><a:fillRect /></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="1209675" y="47625" /><a:ext cx="4572000" cy="2743200" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></xdr:spPr></xdr:pic></mc:Fallback>'
    const variant = (binSize: boolean) => hi2({
      ...(binSize ? { 'xl/charts/chartEx1.xml': TINY } : {}),
      'xl/drawings/drawing1.xml': xml => xml.replace(/<mc:Fallback>[\s\S]*<\/mc:Fallback>/, picture),
      'xl/drawings/_rels/drawing1.xml.rels': xml => xml.replace('</Relationships>', '<Relationship Id="rImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>'),
      '[Content_Types].xml': xml => xml.includes('Extension="png"') ? xml : xml.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default '),
    }, { 'xl/media/image1.png': new Uint8Array(png.toBuffer('image/png')) })
    const green = async (bytes: Uint8Array) => {
      const doc = await parseXlsx(await OfficePackage.load(bytes))
      const units = await getPaintables(doc)
      try {
        const c = createCanvas(Math.ceil(units[0].spec.widthPx), Math.ceil(units[0].spec.heightPx))
        const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D
        units[0].paint(ctx)
        const d = ctx.getImageData(0, 0, c.width, c.height).data
        let n = 0
        for (let i = 0; i < d.length; i += 4) if (d[i] < 40 && d[i + 1] > 215 && d[i + 2] < 40) n++
        return { doc, n }
      } finally { units.dispose() }
    }
    const refused = await green(await variant(true))
    const drawing = refused.doc.sheets[0].drawings?.[0]
    expect(drawing?.source).toMatchObject({ element: 'pic', representation: 'fallback' })
    expect(drawing?.imageIndex).toBe(0)
    expect(refused.doc.images).toHaveLength(1)
    // Painted: the 480x288-px fallback picture's green fills most of its box.
    expect(refused.n).toBeGreaterThan(100_000)
    // Control: the native chart draws no green picture.
    expect((await green(await variant(false))).n).toBe(0)
  }, 60_000)
})
