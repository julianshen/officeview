/**
 * S13.2 item 2 (spec RENDER-02 / AC-RENDER-01): a worker-backed (paged) sheet
 * paints the same drawables as the resident sheet of the same file. The paged
 * reader transports bounded drawing content (shape geometry, fill, line, text,
 * theme and pictures) from the worker; anything it cannot carry is an explicit
 * visible diagnostic, never a silent omission.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import type { XlsxDrawing, XlsxDocument } from '../src/xlsx/types'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { ViewerKernel, connectInProcess } from '../src/worker/viewerKernel'
import { createViewerClient } from '../src/worker/viewerClient'
import { RemoteWorkbook, openRemoteWorkbook, type RemoteOpen } from '../src/worker/viewerRemoteSource'
import { pagedSheetGeometry } from '../src/components/xlsx/pagedGeometry'
import { buildSheetGeometry, drawingExtent, type SheetGeometry } from '../src/xlsx/view/viewport'
import { createIdentityOrder } from '../src/xlsx/view/order'
import { createIntervalSet, visibilityFromHidden } from '../src/xlsx/view/visibility'
import { buildZip, type BuildEntry } from './helpers/zip-builder'
import { xlsxEntries } from './helpers/xlsx-package'
import { dataRows, EMU_PER_PX } from './helpers/xlsx-structure-fixtures'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
/** A 1x1 PNG (valid header, so image admission can probe it). */
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), c => c.charCodeAt(0))
const THEME = `<?xml version="1.0"?><a:theme xmlns:a="${A_NS}" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="12AB34"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`

/** A legacy VML WordArt shape (resident parse.ts paints these from <legacyDrawing>). */
const VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="WordArt_1" style="position:absolute;margin-left:30pt;margin-top:20pt;width:150pt;height:40pt;" fillcolor="#00FF00" strokecolor="#000000"><v:fill color="#00FF00"/><v:stroke color="#000000" weight="1pt"/><v:textpath on="True" style="font-family:'Calibri';font-size:20pt;" string="Legacy"/></v:shape></xml>`
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`
const DIAGRAM_DATA = `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:dsp="${DSP_NS}"><dgm:extLst><a:ext xmlns:a="${A_NS}" uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="rDsp"/></a:ext></dgm:extLst></dgm:dataModel>`
const DIAGRAM_DRAWING = `<?xml version="1.0"?><dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${40 * EMU_PER_PX}" cy="${30 * EMU_PER_PX}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`
const frame = (id: number, name: string, graphic: string, col: number) => `<xdr:oneCellAnchor>${anchorFrom(col, 9)}<xdr:ext cx="${120 * EMU_PER_PX}" cy="${90 * EMU_PER_PX}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${graphic.startsWith('<c:') ? C_NS : DGM_NS}">${graphic}</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`
const anchorFrom = (col: number, row: number, colOffPx = 0, rowOffPx = 0) =>
  `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOffPx * EMU_PER_PX}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOffPx * EMU_PER_PX}</xdr:rowOff></xdr:from>`
const xfrm = (wPx: number, hPx: number) => `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${wPx * EMU_PER_PX}" cy="${hPx * EMU_PER_PX}"/></a:xfrm>`
/** A filled, outlined rectangle with text. */
const filledShape = `<xdr:oneCellAnchor>${anchorFrom(0, 1)}<xdr:ext cx="${64 * EMU_PER_PX}" cy="${50 * EMU_PER_PX}"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Box"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrm(64, 50)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`
/** A theme-coloured ellipse below a hidden row and right of a custom-width column. */
const themedShape = `<xdr:twoCellAnchor>${anchorFrom(2, 4, 3, 2)}<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Oval"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>${xfrm(80, 60)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`
const picture = `<xdr:oneCellAnchor>${anchorFrom(5, 2)}<xdr:ext cx="${30 * EMU_PER_PX}" cy="${30 * EMU_PER_PX}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="4" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(30, 30)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`

interface DrawingWorkbook { anchors: string; theme?: boolean; media?: boolean; secondImage?: boolean; vml?: boolean; unusedMedia?: number; content?: boolean }
function drawingWorkbook(spec: DrawingWorkbook): Uint8Array {
  const extra: BuildEntry[] = [
    { name: 'xl/worksheets/_rels/sheet1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rD1" Type="${R_NS}/drawing" Target="../drawings/drawing1.xml"/>${spec.vml ? `<Relationship Id="rV1" Type="${R_NS}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` : ''}</Relationships>` },
    { name: 'xl/drawings/drawing1.xml', data: `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${spec.anchors}</xdr:wsDr>` },
  ]
  if (spec.media) {
    extra.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rImg" Type="${R_NS}/image" Target="../media/image1.png"/>${spec.secondImage ? `<Relationship Id="rImg2" Type="${R_NS}/image" Target="../media/image2.png"/>` : ''}${spec.unusedMedia ? `<Relationship Id="rUnused" Type="${R_NS}/image" Target="../media/unused.png"/>` : ''}</Relationships>` })
    if (spec.unusedMedia) extra.push({ name: 'xl/media/unused.png', data: new Uint8Array(spec.unusedMedia) })
    extra.push({ name: 'xl/media/image1.png', data: PNG })
    if (spec.secondImage) extra.push({ name: 'xl/media/image2.png', data: PNG.slice() })
  }
  if (spec.content) {
    extra.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rC1" Type="${R_NS}/chart" Target="../charts/chart1.xml"/><Relationship Id="rDm" Type="${R_NS}/diagramData" Target="../diagrams/data1.xml"/></Relationships>` })
    extra.push({ name: 'xl/charts/chart1.xml', data: CHART })
    extra.push({ name: 'xl/diagrams/data1.xml', data: DIAGRAM_DATA })
    extra.push({ name: 'xl/diagrams/_rels/data1.xml.rels', data: `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}"><Relationship Id="rDsp" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>` })
    extra.push({ name: 'xl/diagrams/drawing1.xml', data: DIAGRAM_DRAWING })
  }
  if (spec.theme) extra.push({ name: 'xl/theme/theme1.xml', data: THEME })
  if (spec.vml) extra.push({ name: 'xl/drawings/vmlDrawing1.vml', data: VML })
  // Column C is 20 chars wide (20*7+5 = 145 px); row 3 is hidden and row 4 is 30pt (40 px).
  const entries = xlsxEntries({
    beforeSheetData: '<cols><col min="3" max="3" width="20" customWidth="1"/></cols>',
    sheetData: dataRows(8, () => '', r => r === 2 ? ' hidden="1"' : r === 3 ? ' ht="30" customHeight="1"' : ''),
    afterSheetData: `<drawing xmlns:r="${R_NS}" r:id="rD1"/>${spec.vml ? `<legacyDrawing xmlns:r="${R_NS}" r:id="rV1"/>` : ''}`,
    extra,
  })
  if (spec.theme) {
    const wbRels = entries.find(e => e.name === 'xl/_rels/workbook.xml.rels')!
    wbRels.data = String(wbRels.data).replace('</Relationships>', `<Relationship Id="rId8" Type="${R_NS}/theme" Target="theme/theme1.xml"/></Relationships>`)
  }
  return buildZip(entries).bytes
}

const quiet = { yield: async () => {} }
const box = [0, 0, 4000, 4000, 20000] as const
const remotes: RemoteWorkbook[] = []
afterEach(async () => { for (const remote of remotes.splice(0)) await remote.close() })

async function openPaged(bytes: Uint8Array, sourceOptions?: Record<string, unknown>): Promise<RemoteWorkbook> {
  const blob = new Blob([new Uint8Array(bytes)])
  const session = await openImportSession({ kind: 'blob', blob }, { docId: 'P', profile: 'D' })
  expect((await session.done).kind).toBe('complete')
  const source = (createPagedWorkbookSource as (...args: unknown[]) => ReturnType<typeof createPagedWorkbookSource>)(session, sourceOptions)
  const kernel = new ViewerKernel({ parse: async () => { throw new Error('the paged source must serve this file') }, openSource: async () => source }, quiet)
  const client = createViewerClient(connectInProcess(kernel))
  const open = await openRemoteWorkbook(client, 'P', blob, 'D')
  expect('layouts' in open).toBe(true)
  const remote = new RemoteWorkbook(client, open as RemoteOpen)
  remotes.push(remote)
  remote.merges('1')
  for (let i = 0; i < 100 && remote.loadState('1').completeness !== 'complete'; i++) await new Promise(resolve => setTimeout(resolve, 2))
  expect(remote.loadState('1').completeness).toBe('complete')
  return remote
}
const pagedGeometry = (remote: RemoteWorkbook): SheetGeometry => pagedSheetGeometry(remote, '1', createIdentityOrder(), visibilityFromHidden(createIntervalSet([])))
/** Waits (bounded) until the paged geometry reports drawables or a diagnostic. */
async function settledGeometry(remote: RemoteWorkbook): Promise<SheetGeometry> {
  let geometry = pagedGeometry(remote)
  for (let i = 0; i < 200 && !geometry.drawingsIn(...box)?.length && !geometry.diagnostics.length; i++) {
    await new Promise(resolve => setTimeout(resolve, 2))
    geometry = pagedGeometry(remote)
  }
  return geometry
}
async function resident(bytes: Uint8Array): Promise<{ doc: XlsxDocument; geometry: SheetGeometry }> {
  const doc = await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })
  const geometry = buildSheetGeometry(doc.sheets[0])
  if (geometry.kind !== 'ready') throw new Error('resident geometry refused')
  return { doc, geometry }
}
/** Placement, painted extent and paint-relevant content of one drawable. */
const key = (node: XlsxDrawing) => ({
  element: node.source.element, name: node.source.name,
  xEmu: node.xEmu, yEmu: node.yEmu, widthEmu: node.widthEmu, heightEmu: node.heightEmu,
  extent: drawingExtent([node]),
  geometry: node.drawingGeometry, style: node.drawingStyle, fill: node.fill, line: node.line, text: node.textBody,
})

describe('paged drawings (RENDER-02 / AC-RENDER-01)', () => {
  it('paints the same drawables as the resident sheet: count, placement, geometry, fill, line, text and theme', async () => {
    const bytes = drawingWorkbook({ anchors: filledShape + themedShape, theme: true })
    const { doc, geometry: object } = await resident(bytes)
    const expected = object.drawingsIn(...box)!
    expect(expected).toHaveLength(2)
    const geometry = await settledGeometry(await openPaged(bytes))
    const actual = geometry.drawingsIn(...box)
    expect(actual?.length).toBe(2)
    expect(actual!.map(key)).toEqual(expected.map(key))
    expect(geometry.diagnostics).toEqual([])
    // A query window that misses every drawing returns none on both paths.
    expect(geometry.drawingsIn(2000, 2000, 2100, 2100, 100)).toEqual(object.drawingsIn(2000, 2000, 2100, 2100, 100))
    expect(doc.sheets[0].drawings).toHaveLength(2)
  })

  it('paints legacy VML drawings like the resident sheet', async () => {
    const bytes = drawingWorkbook({ anchors: filledShape, vml: true })
    const { geometry: object } = await resident(bytes)
    const expected = object.drawingsIn(...box)!
    expect(expected.map(node => node.source.partPath)).toContain('xl/drawings/vmlDrawing1.vml')
    const geometry = await settledGeometry(await openPaged(bytes))
    expect(geometry.drawingsIn(...box)!.map(key)).toEqual(expected.map(key))
    expect(geometry.diagnostics).toEqual([])
  })

  it('carries chart and SmartArt content across the worker boundary like the resident sheet', async () => {
    const chart = frame(6, 'Chart 1', `<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="rC1"/>`, 0)
    const smartArt = frame(7, 'Diagram 1', `<dgm:relIds xmlns:dgm="${DGM_NS}" xmlns:r="${R_NS}" r:dm="rDm"/>`, 4)
    const bytes = drawingWorkbook({ anchors: chart + smartArt, content: true })
    const { geometry: object } = await resident(bytes)
    const expected = object.drawingsIn(...box)!
    expect(expected.map(node => node.content?.kind)).toEqual(['chart', 'diagram'])
    const geometry = await settledGeometry(await openPaged(bytes))
    const actual = geometry.drawingsIn(...box)!
    expect(actual.map(key)).toEqual(expected.map(key))
    expect(actual.map(node => node.content)).toEqual(expected.map(node => node.content))
    expect(geometry.diagnostics).toEqual([])
  })

  it('transports pictures with their bytes within the image budgets', async () => {
    const bytes = drawingWorkbook({ anchors: filledShape + picture, media: true })
    const { doc, geometry: object } = await resident(bytes)
    const expected = object.drawingsIn(...box)!
    const geometry = await settledGeometry(await openPaged(bytes))
    const actual = geometry.drawingsIn(...box)
    expect(actual?.length).toBe(expected.length)
    expect(actual!.map(key)).toEqual(expected.map(key))
    const pic = actual!.find(node => node.source.element === 'pic')!
    expect([...pic.image!.data!]).toEqual([...PNG])
    expect([...pic.image!.data!]).toEqual([...doc.images![expected.find(node => node.source.element === 'pic')!.imageIndex!].data])
    expect(geometry.diagnostics).toEqual([])
  })

  it('names a picture part beyond the transfer budget in a visible diagnostic, keeping every other drawable', async () => {
    const bytes = drawingWorkbook({ anchors: filledShape + picture, media: true })
    const geometry = await settledGeometry(await openPaged(bytes, { drawingPayloadBytes: 16 }))
    expect(geometry.diagnostics.map(d => d.code)).toContain('drawing-part-refused')
    expect(geometry.diagnostics.find(d => d.code === 'drawing-part-refused')!.message).toContain('xl/media/image1.png')
    const actual = geometry.drawingsIn(...box)!
    expect(actual.map(node => node.source.element).sort()).toEqual(['pic', 'sp'])
    expect(actual.find(node => node.source.element === 'pic')!.image?.data).toBeUndefined()
  })

  it('reads only the parts the parser opens: an unused oversized relationship target costs nothing', async () => {
    const bytes = drawingWorkbook({ anchors: filledShape + picture, media: true, unusedMedia: 4096 })
    const geometry = await settledGeometry(await openPaged(bytes, { drawingPayloadBytes: PNG.byteLength + 10 }))
    expect(geometry.diagnostics).toEqual([])
    expect(geometry.drawingsIn(...box)!.find(node => node.source.element === 'pic')!.image?.data?.byteLength).toBe(PNG.byteLength)
  })

  it('withdraws pictures past the payload budget in paint order and says so', async () => {
    const second = picture.replace('id="4" name="Picture"', 'id="5" name="Picture 2"').replace('rImg', 'rImg2').replace(anchorFrom(5, 2), anchorFrom(6, 5))
    const bytes = drawingWorkbook({ anchors: filledShape + picture + second, media: true, secondImage: true })
    const geometry = await settledGeometry(await openPaged(bytes, { drawingPayloadBytes: PNG.byteLength + 10 }))
    expect(geometry.diagnostics.map(d => d.code)).toEqual(['drawing-asset-refused'])
    const pictures = geometry.drawingsIn(...box)!.filter(node => node.source.element === 'pic')
    expect(pictures.map(node => node.source.name)).toEqual(['Picture', 'Picture 2'])
    expect(pictures.map(node => node.image?.data?.byteLength)).toEqual([PNG.byteLength, undefined])
    expect(pictures.map(node => node.imageIndex)).toEqual([0, undefined])
  })
})
