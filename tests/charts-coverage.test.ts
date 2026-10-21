/**
 * Task A7 (spec 6.6): drawing coverage and chart diagnostics. Coverage reports a native chart as `chart`, a card
 * as `chart-placeholder`, and an authored fallback picture as the existing fallback representation that keeps the
 * original feature (`chart` / `ChartEx`), in DOCX, PPTX and XLSX. Every chart diagnostic code of spec 6.6 is
 * emitted with a human-readable reason. Fixtures are Office-shaped synthetic, never Office evidence.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { drawingPartContext } from '../src/drawing/parts'
import { parseDocx } from '../src/docx/parse'
import { parsePptx } from '../src/pptx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { parseXlsxChart } from '../src/charts/parse'
import { applyDataPrecedence } from '../src/charts/dataPrecedence'
import { readChartPart } from '../src/charts/read'
import { chartParseContextFor } from '../src/charts/context'
import { prepareChartAxes } from '../src/charts/prepare'
import { fillChartRanges } from '../src/charts/fill'
import { createChartWorkspace } from '../src/charts/workspace'
import type { DrawingCoverageEntry } from '../src/drawing/coverage'

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'
const EMU = 9525

const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
const ser = (i: number, values: number[], extra = '') => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>S${i + 1}</c:v></c:tx><c:cat><c:strLit><c:ptCount val="${values.length}"/>${values.map((_, k) => `<c:pt idx="${k}"><c:v>C${k}</c:v></c:pt>`).join('')}</c:strLit></c:cat><c:val>${lit(values)}</c:val>${extra}</c:ser>`
const axes = '<c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>'
const space = (chart: string) => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart>${chart}</c:chart></c:chartSpace>`
const COLUMN = space(`<c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [2, 3])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea>`)
const SURFACE = space(`<c:plotArea><c:surface3DChart>${ser(0, [1, 2])}</c:surface3DChart>${axes}</c:plotArea>`)
const png = () => { const c = createCanvas(2, 2); const g = c.getContext('2d'); g.fillStyle = '#00ff00'; g.fillRect(0, 0, 2, 2); return new Uint8Array(c.toBuffer('image/png')) }

type Format = 'docx' | 'pptx' | 'xlsx'
const FORMATS: Format[] = ['docx', 'pptx', 'xlsx']
const NAMESPACES = `xmlns:mc="${MC}" xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}" xmlns:cx="${CX}" xmlns:cx1="${CX1}" xmlns:pic="${PIC}"`
const rels = (...items: Array<[string, string, string]>) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`
function graphicObject(format: Format, id: number, payload: string, uri = C): string {
  const graphic = `<a:graphic><a:graphicData uri="${uri}">${payload}</a:graphicData></a:graphic>`
  if (format === 'docx') return `<w:drawing><wp:inline><wp:extent cx="${200 * EMU}" cy="${160 * EMU}"/><wp:docPr id="${id}" name="Object ${id}"/>${graphic}</wp:inline></w:drawing>`
  const xfrm = `<a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${160 * EMU}"/>`
  if (format === 'pptx') return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Object ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm>${xfrm}</p:xfrm>${graphic}</p:graphicFrame>`
  return `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Object ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm>${xfrm}</xdr:xfrm>${graphic}</xdr:graphicFrame>`
}
function picture(format: Format, id: number): string {
  if (format === 'docx') return graphicObject('docx', id, `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="im"/></pic:blipFill><pic:spPr/></pic:pic>`, PIC)
  const xfrm = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * EMU}" cy="${160 * EMU}"/></a:xfrm>`
  if (format === 'pptx') return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="im"/></p:blipFill><p:spPr>${xfrm}<a:prstGeom prst="rect"/></p:spPr></p:pic>`
  return `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="im"/></xdr:blipFill><xdr:spPr>${xfrm}<a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>`
}
const alternate = (choice: string, fallback: string, requires = 'c') => `<mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

async function packageOf(parts: Record<string, string | Uint8Array>): Promise<OfficePackage> {
  const zip = new JSZip()
  for (const [path, value] of Object.entries(parts)) zip.file(path, value)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
/** A one-object document whose body is `body`, with chart1.xml (`chart`), an empty chartEx1.xml and a picture. */
async function build(format: Format, body: string, chart: string, image: Uint8Array | null = png()): Promise<OfficePackage> {
  const chartEx = `<cx:chartSpace xmlns:cx="${CX}"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>`
  const relItems = (prefix: string): Array<[string, string, string]> => [['rC', 'chart', `${prefix}charts/chart1.xml`], ['rX', 'chartEx', `${prefix}charts/chartEx1.xml`], ['im', 'image', `${prefix}media/image1.png`]]
  if (format === 'docx') return packageOf({
    'word/document.xml': `<w:document xmlns:w="${W}" xmlns:wp="${WP}" ${NAMESPACES}><w:body><w:p><w:r>${body}</w:r></w:p></w:body></w:document>`,
    'word/_rels/document.xml.rels': rels(...relItems('')),
    'word/charts/chart1.xml': chart, 'word/charts/chartEx1.xml': chartEx, 'word/media/image1.png': png(),
  })
  if (format === 'pptx') return packageOf({
    'ppt/presentation.xml': `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldSz cx="${300 * EMU}" cy="${200 * EMU}"/><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(['s1', 'slide', 'slides/slide1.xml']),
    'ppt/slides/slide1.xml': `<p:sld xmlns:p="${P}" ${NAMESPACES}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': rels(...relItems('../')),
    'ppt/charts/chart1.xml': chart, 'ppt/charts/chartEx1.xml': chartEx, 'ppt/media/image1.png': png(),
  })
  return packageOf({
    'xl/workbook.xml': `<workbook xmlns:r="${R}"><sheets><sheet name="One" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels(['s1', 'worksheet', 'worksheets/sheet1.xml']),
    'xl/worksheets/sheet1.xml': `<worksheet xmlns:r="${R}"><sheetData/><drawing r:id="dr"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(['dr', 'drawing', '../drawings/drawing1.xml']),
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}" ${NAMESPACES}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="${200 * EMU}" cy="${160 * EMU}"/>${body}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': rels(...relItems('../')),
    'xl/charts/chart1.xml': chart, 'xl/charts/chartEx1.xml': chartEx, ...(image ? { 'xl/media/image1.png': image } : {}),
  })
}
async function coverage(format: Format, pkg: OfficePackage): Promise<DrawingCoverageEntry[]> {
  const doc = format === 'docx' ? await parseDocx(pkg) : format === 'pptx' ? await parsePptx(pkg) : await parseXlsx(pkg)
  return (doc.drawingCoverage ?? []).filter(entry => entry.scope === 'original')
}

describe('chart coverage (spec 6.6)', () => {
  it('coverage: native chart, card chart-placeholder, fallback keeps the original feature', async () => {
    for (const format of FORMATS) {
      const native = await coverage(format, await build(format, graphicObject(format, 10, '<c:chart r:id="rC"/>'), COLUMN))
      expect(native, `${format} native`).toEqual([expect.objectContaining({ id: '10', feature: 'chart', status: 'native', selectedRepresentation: 'chart' })])
      const card = await coverage(format, await build(format, graphicObject(format, 20, '<c:chart r:id="rC"/>'), SURFACE))
      expect(card, `${format} card`).toEqual([expect.objectContaining({ id: '20', feature: 'chart', selectedRepresentation: 'chart-placeholder' })])
      // An authored fallback picture: the existing fallback representation, with the original feature.
      const fallback = await coverage(format, await build(format, alternate(graphicObject(format, 30, '<c:chart r:id="rC"/>'), picture(format, 31)), SURFACE))
      expect(fallback, `${format} chart fallback`).toEqual([expect.objectContaining({ status: 'fallback', representation: 'fallback', selectedRepresentation: 'raster-fallback', feature: 'chart' })])
      const chartEx = await coverage(format, await build(format, alternate(graphicObject(format, 40, '<cx:chart r:id="rX"/>', CX), picture(format, 41), 'cx1'), COLUMN))
      expect(chartEx, `${format} ChartEx fallback`).toEqual([expect.objectContaining({ status: 'fallback', representation: 'fallback', selectedRepresentation: 'raster-fallback', feature: 'ChartEx' })])
    }
  })
})

/** A code is "emitted with a reason" when its reason is a non-empty sentence, not the code repeated. */
function expectReason(found: { code?: string; reason?: string; message?: string } | undefined, code: string): void {
  expect(found, `${code} emitted`).toBeDefined()
  const reason = found!.reason ?? found!.message ?? ''
  expect(reason.trim().length, `${code}: reason`).toBeGreaterThan(3)
  expect(reason.replace(/^[a-z-]+:\s*/, ''), `${code}: the reason is more than the code`).not.toBe(code)
  expect(reason.replace(/^[a-z-]+:\s*/, '').trim().length, `${code}: reason text`).toBeGreaterThan(3)
}

describe('chart diagnostics (spec 6.6)', () => {
  const theme = undefined
  const model = (xml: string) => parseXlsxChart(parseXmlOrdered(xml), theme)!
  const diagnostic = (list: Array<{ code: string; reason: string }> | undefined, code: string) => list?.find(d => d.code === code)

  it('every new diagnostic code is emitted with a reason', async () => {
    // chart-data-unavailable: a series with neither saved values nor a reference, observed on the produced model.
    const absentXml = space(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Empty</c:v></c:tx></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea>`)
    const absentPkg = await packageOf({ 'word/charts/chart1.xml': absentXml })
    const absentRoot = parseXmlOrdered(absentXml)
    const absent = await applyDataPrecedence(parseXlsxChart(absentRoot, theme)!, absentRoot, chartParseContextFor(absentPkg, 'word/charts/chart1.xml'))
    expect(absent.unsupported).toBeTruthy()
    expectReason(diagnostic(absent.diagnostics, 'chart-data-unavailable'), 'chart-data-unavailable')

    // chart-point-unavailable: a numeric point that does not exist (prepared at stage 1).
    const sparse = model(space(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea>`))
    sparse.groups[0].series[0].values.states[1] = 3 // POINT_UNAVAILABLE
    expectReason(diagnostic(prepareChartAxes(sparse).preparedAxes?.diagnostics, 'chart-point-unavailable'), 'chart-point-unavailable')

    // chart-date-unrepresentable: a 1900-system cell (the fictitious day 60) on a 1904-system date axis.
    const dated = model(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:date1904/><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>S!$A$1:$A$2</c:f></c:numRef></c:cat><c:val>${lit([1, 2])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:dateAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:dateAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`)
    const filled = fillChartRanges(dated, 'S', (_sheet: string, row: number) => { const v = row === 0 ? 60 : 45000; return { value: v, number: v, text: String(v) } }, { epoch: '1900' })
    expectReason(diagnostic(filled.diagnostics, 'chart-date-unrepresentable'), 'chart-date-unrepresentable')

    // chart-refused: a workspace that cannot hold the chart part's read; the read itself reports the code.
    const refusing = createChartWorkspace('mobile')
    refusing.reserve(32 * 1024 * 1024, 'full')
    const refused = await readChartPart(await packageOf({ 'word/charts/chart1.xml': COLUMN }), 'word/charts/chart1.xml', { profile: 'mobile', workspace: refusing })
    expect(refused).toMatchObject({ kind: 'refused', code: 'chart-refused' })
    expectReason(refused as { code: string; reason: string }, 'chart-refused')
    // The drawing-level record of a refused chart carries the code as its reason and the human reason as its message.
    const docRefusedPkg = await build('docx', graphicObject('docx', 60, '<c:chart r:id="rC"/>'), COLUMN)
    const { documentChartContext } = await import('../src/charts/context')
    const full = documentChartContext('mobile')
    full.workspace.reserve(32 * 1024 * 1024, 'full')
    await parseDocx(docRefusedPkg, { chartContext: full })
    const docRefused = drawingPartContext(docRefusedPkg).diagnostics.find(d => d.reason === 'chart-refused')
    expectReason(docRefused && { code: 'chart-refused', message: docRefused.message }, 'chart-refused')

    // chart-fallback-unusable: the Fallback picture has no image part; the card records why.
    const broken = await build('xlsx', alternate(graphicObject('xlsx', 50, '<c:chart r:id="rC"/>'), picture('xlsx', 51)), SURFACE, null)
    const brokenDoc = await parseXlsx(broken)
    const brokenSource = brokenDoc.sheets[0].drawings?.[0]?.source.reason
    expectReason({ code: 'chart-fallback-unusable', reason: brokenSource }, 'chart-fallback-unusable')
    // The drawing diagnostic keeps the code as its reason and carries the human reason in its message.
    const brokenDiagnostic = drawingPartContext(broken).diagnostics.find(d => d.reason === 'chart-fallback-unusable')
    expectReason(brokenDiagnostic && { code: 'chart-fallback-unusable', message: brokenDiagnostic.message }, 'chart-fallback-unusable')
    expect(brokenDiagnostic!.message).toContain('3-D surface charts are shown as a placeholder')

    // chart-invalid-log-value: a value <= 0 on a log axis.
    const log = model(space(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [-1, 10])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:logBase val="10"/></c:scaling><c:crossAx val="1"/></c:valAx></c:plotArea>`))
    expectReason(diagnostic(prepareChartAxes(log).preparedAxes?.diagnostics, 'chart-invalid-log-value'), 'chart-invalid-log-value')

    // chart-singular-trendline: an exponential trendline over non-positive values.
    const trend = model(space(`<c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:trendline><c:trendlineType val="exp"/></c:trendline><c:xVal>${lit([1, 2, 3])}</c:xVal><c:yVal>${lit([-1, -2, -3])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea>`))
    const trendCodes = [...(trend.diagnostics ?? []), ...(prepareChartAxes(trend).preparedAxes?.diagnostics ?? [])]
    expectReason(diagnostic(trendCodes, 'chart-singular-trendline'), 'chart-singular-trendline')

    // chart-label-feature-unsupported: a field-reference data label.
    const labelled = model(space(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [1, 2], '<c:dLbls><c:showVal val="1"/><c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"><c15:showDataLabelsRange xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" val="1"/></c:ext></c:extLst></c:dLbls>')}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea>`))
    expectReason(diagnostic(prepareChartAxes(labelled).preparedAxes?.diagnostics, 'chart-label-feature-unsupported'), 'chart-label-feature-unsupported')

    // chart-axis-unplaceable: a bar group and a scatter group on the same axes.
    const mixed = model(space(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:scatterChart><c:ser><c:idx val="1"/><c:order val="1"/><c:xVal>${lit([1, 2])}</c:xVal><c:yVal>${lit([1, 2])}</c:yVal></c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>${axes}</c:plotArea>`))
    expectReason(diagnostic(mixed.diagnostics, 'chart-axis-unplaceable'), 'chart-axis-unplaceable')
  })
})
