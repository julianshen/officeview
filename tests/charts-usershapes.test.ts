// @vitest-environment node
/**
 * Task A4d (spec 5.11): chart user shapes (`c:userShapes`, cdr relSizeAnchor/absSizeAnchor) drawn over the chart
 * with the existing drawing preparation and scene painter. Packages are Office-shaped synthetic fixtures
 * (tests/helpers builders), never Office evidence.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { createCanvas } from 'canvas'
import type JSZip from 'jszip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseChartPart } from '../src/charts/dataPrecedence'
import { paintXlsxChart } from '../src/charts/paint'
import { releaseChartModel } from '../src/charts/own'
import { createChartWorkspace, type ChartWorkspace } from '../src/charts/workspace'
import { USER_SHAPE_NODE_BYTES, chartUserShapesReady } from '../src/charts/userShapes'
import { imageLimitsFor } from '../src/xlsx/storage/imageAdmission'
import type { ChartModel, ChartResult } from '../src/charts/model'
import { chartWorkbook, chartWorkbookBytes, emu } from './helpers/xlsx-chart-fixture'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { openImportSession } from '../src/xlsx/storage/pagedImport'
import { createPagedWorkbookSource } from '../src/xlsx/view/pagedSource'

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CDR = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const lit = (values: number[]) => `<c:numLit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`
/** One small bar series in the lower left; legend and title off, so the upper right of the chart is blank paper. */
const chartXml = (userShapes = '<c:userShapes r:id="rUS"/>') => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="808080"/></a:solidFill></c:spPr><c:val>${lit([1, 2])}</c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/><c:delete val="1"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/><c:delete val="1"/><c:min val="0"/><c:max val="10"/></c:valAx></c:plotArea></c:chart>${userShapes}</c:chartSpace>`
const rect = (hex: string, w: number, h: number) => `<cdr:sp><cdr:nvSpPr><cdr:cNvPr id="2" name="Box"/><cdr:cNvSpPr/></cdr:nvSpPr><cdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:ln><a:noFill/></a:ln></cdr:spPr></cdr:sp>`
const pic = (rId: string) => `<cdr:pic><cdr:nvPicPr><cdr:cNvPr id="3" name="Picture"/><cdr:cNvPicPr/></cdr:nvPicPr><cdr:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></cdr:blipFill><cdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(40)}" cy="${emu(40)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr></cdr:pic>`
const relAnchor = (from: [number, number], to: [number, number], body: string) => `<cdr:relSizeAnchor><cdr:from><cdr:x>${from[0]}</cdr:x><cdr:y>${from[1]}</cdr:y></cdr:from><cdr:to><cdr:x>${to[0]}</cdr:x><cdr:y>${to[1]}</cdr:y></cdr:to>${body}</cdr:relSizeAnchor>`
const absAnchor = (from: [number, number], cx: number, cy: number, body: string) => `<cdr:absSizeAnchor><cdr:from><cdr:x>${from[0]}</cdr:x><cdr:y>${from[1]}</cdr:y></cdr:from><cdr:ext cx="${cx}" cy="${cy}"/>${body}</cdr:absSizeAnchor>`
const userShapesXml = (anchors: string) => `<c:userShapes xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}" xmlns:cdr="${CDR}">${anchors}</c:userShapes>`
const rels = (...items: Array<[string, string, string]>) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`
const CHART = 'xl/charts/chart1.xml', SHAPES = 'xl/drawings/drawing2.xml'
const W = 400, H = 260

/** A 4x4 solid PNG (node-canvas encodes it). */
const png = (hex: string, size = 4): Uint8Array => {
  const canvas = createCanvas(size, size), ctx = canvas.getContext('2d')
  ctx.fillStyle = `#${hex}`; ctx.fillRect(0, 0, size, size)
  return new Uint8Array(canvas.toBuffer('image/png'))
}
/** A PNG header claiming `w` x `h` pixels (only the IHDR is read before decode). */
const forgedPng = (w: number, h: number): Uint8Array => {
  const bytes = png('00FF00')
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  view.setUint32(16, w); view.setUint32(20, h)
  return bytes
}

async function parsed(patch: (zip: JSZip) => void, options: { chart?: string; workspace?: ChartWorkspace } = {}): Promise<{ result: ChartResult; workspace: ChartWorkspace }> {
  const pkg = await chartWorkbook(options.chart ?? chartXml(), { patch })
  const workspace = options.workspace ?? createChartWorkspace('desktop')
  const root = parseXmlOrdered((await pkg.text(CHART))!)
  const result = await parseChartPart(root, undefined, { pkg, partPath: CHART, profile: 'desktop', workspace })
  return { result, workspace }
}
const native = (result: ChartResult): ChartModel => {
  expect(result.kind).toBe('native')
  return (result as { model: ChartModel }).model
}
const codes = (model: ChartModel) => (model.diagnostics ?? []).map(d => d.code)
function paint(model: ChartModel) {
  const canvas = createCanvas(W, H), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
  paintXlsxChart(model, ctx, W, H)
  const data = ctx.getImageData(0, 0, W, H).data
  const hex = (x: number, y: number) => { const i = (Math.round(y) * W + Math.round(x)) * 4; return [data[i], data[i + 1], data[i + 2]].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase() }
  return { data, hex }
}

describe('charts-usershapes (Task A4d)', () => {
  it('relSizeAnchor and absSizeAnchor place shapes over the chart', async () => {
    // relSizeAnchor: 50%..75% of the chart in both directions. absSizeAnchor: from 80%, 5%, a 40 x 30 px extent.
    const anchors = relAnchor([0.5, 0.1], [0.75, 0.4], rect('FF0000', 10, 10)) + absAnchor([0.8, 0.05], emu(40), emu(30), rect('0000FF', 10, 10))
    const { result, workspace } = await parsed(zip => {
      zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing2.xml']))
      zip.file(SHAPES, userShapesXml(anchors))
    })
    const model = native(result)
    expect(model.unsupported).toBeUndefined()
    expect(model.userShapes?.anchors.map(a => a.kind)).toEqual(['rel', 'abs'])
    expect(model.userShapes!.anchors[0]).toMatchObject({ x: 0.5, y: 0.1, x2: 0.75, y2: 0.4 })
    expect(model.userShapes!.anchors[1]).toMatchObject({ x: 0.8, y: 0.05, cx: emu(40), cy: emu(30) })
    const { hex } = paint(model)
    // Red covers 200..300 x 26..104 (the anchor box, not the shape's own 10 px extent).
    expect(hex(210, 30)).toBe('FF0000')
    expect(hex(290, 100)).toBe('FF0000')
    expect(hex(250, 65)).toBe('FF0000')
    expect(hex(195, 65)).not.toBe('FF0000')
    expect(hex(305, 65)).not.toBe('FF0000')
    expect(hex(250, 110)).not.toBe('FF0000')
    // Blue covers 320..360 x 13..43.
    expect(hex(322, 15)).toBe('0000FF')
    expect(hex(358, 41)).toBe('0000FF')
    expect(hex(364, 30)).not.toBe('0000FF')
    expect(hex(340, 47)).not.toBe('0000FF')
    // A relative anchor scales with the chart; the absolute extent does not.
    const canvas = createCanvas(800, 520), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintXlsxChart(model, ctx, 800, 520)
    const big = ctx.getImageData(0, 0, 800, 520).data, at = (x: number, y: number) => { const i = (y * 800 + x) * 4; return [big[i], big[i + 1], big[i + 2]] }
    expect(at(590, 200)).toEqual([255, 0, 0])
    expect(at(645, 30)).toEqual([0, 0, 255])
    expect(at(685, 30)).not.toEqual([0, 0, 255])
    // The chart itself is still drawn natively beneath (its grey bars), and nothing else changes: a chart without
    // the overlay paints exactly as before.
    const plain = native((await parsed(() => {}, { chart: chartXml('') })).result)
    const without = { ...model, userShapes: undefined }
    expect(Buffer.from(paint(without).data).equals(Buffer.from(paint(plain).data))).toBe(true)
    expect(paint(plain).data.some((_, i, d) => i % 4 === 0 && d[i] === 0x80 && d[i + 1] === 0x80 && d[i + 2] === 0x80)).toBe(true)
    // Charged before it was built, refunded with the model.
    expect(workspace.used).toBeGreaterThanOrEqual(2 * USER_SHAPE_NODE_BYTES)
    releaseChartModel(model)
    expect(workspace.used).toBe(0)
  })

  it('user-shape images use the image budget; relationship cycles are refused', async () => {
    const image = png('00FF00')
    const anchors = relAnchor([0.5, 0.1], [0.6, 0.3], pic('rImg')) + relAnchor([0.7, 0.1], [0.8, 0.3], pic('rBig')) + relAnchor([0.1, 0.1], [0.2, 0.3], pic('rSelf'))
    const { result, workspace } = await parsed(zip => {
      zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing2.xml']))
      zip.file(SHAPES, userShapesXml(anchors))
      zip.file('xl/drawings/_rels/drawing2.xml.rels', rels(['rImg', 'image', '../media/image1.png'], ['rBig', 'image', '../media/image2.png'], ['rSelf', 'image', 'drawing2.xml']))
      zip.file('xl/media/image1.png', image)
      zip.file('xl/media/image2.png', forgedPng(100_000, 100_000))
    })
    const model = native(result)
    const images = model.userShapes!.images
    // Only the admitted picture is kept: the oversized one (over the image budget's pixels) and the cyclic one are not.
    expect(images.length).toBe(1)
    expect(images[0].data.length).toBe(image.length)
    expect(images[0]).toMatchObject({ width: 4, height: 4 })
    expect(100_000 * 100_000).toBeGreaterThan(imageLimitsFor('D').maxPixels)
    expect(codes(model)).toContain('chart-user-shapes-image-limit')
    expect(codes(model)).toContain('chart-user-shapes-cycle')
    // Encoded and decoded bytes are both charged to the chart workspace.
    expect(workspace.used).toBeGreaterThanOrEqual(image.length + 4 * 4 * 4)
    await chartUserShapesReady(model)
    const { hex } = paint(model)
    expect(hex(220, 50)).toBe('00FF00')
    expect(hex(300, 50)).not.toBe('00FF00')
    releaseChartModel(model)
    expect(workspace.used).toBe(0)

    // The chart's user-shapes relationship pointing back at the chart part itself is never followed.
    const cyclic = await parsed(zip => { zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', 'chart1.xml'])) })
    const back = native(cyclic.result)
    expect(back.userShapes).toBeUndefined()
    expect(codes(back)).toContain('chart-user-shapes-cycle')
    releaseChartModel(back)
    expect(cyclic.workspace.used).toBe(0)

    // A workspace that cannot admit the overlay refuses the chart and refunds everything.
    let held = 0
    const refusing = createChartWorkspace('desktop', { reserve: (bytes, tag) => { if (tag.startsWith('chart-user-shapes')) return false; held += bytes; return true }, release: (_tag, bytes) => { held -= bytes } })
    const refused = await parsed(zip => {
      zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing2.xml']))
      zip.file(SHAPES, userShapesXml(relAnchor([0, 0], [1, 1], rect('FF0000', 10, 10))))
    }, { workspace: refusing })
    expect(refused.result.kind).toBe('refused')
    expect(refusing.used).toBe(0)
    expect(held).toBe(0)
  })

  it('missing user-shapes part leaves the chart native with a diagnostic', async () => {
    const plainPixels = paint(native((await parsed(() => {}, { chart: chartXml('') })).result)).data
    const cases: Array<[string, (zip: JSZip) => void, string]> = [
      ['no relationships part', () => {}, 'chart-user-shapes-missing'],
      ['relationship to an absent part', zip => { zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing9.xml'])) }, 'chart-user-shapes-missing'],
      ['malformed part', zip => {
        zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing2.xml']))
        zip.file(SHAPES, `<c:userShapes xmlns:c="${C}"><cdr:relSizeAnchor`)
      }, 'chart-user-shapes-unreadable'],
    ]
    for (const [name, patch, code] of cases) {
      const { result, workspace } = await parsed(patch)
      const model = native(result)
      expect(model.unsupported, name).toBeUndefined()
      expect(model.userShapes, name).toBeUndefined()
      expect(codes(model), name).toContain(code)
      // Drawn exactly as the same chart without the overlay.
      expect(Buffer.from(paint(model).data).equals(Buffer.from(plainPixels)), name).toBe(true)
      releaseChartModel(model)
      expect(workspace.used, name).toBe(0)
    }
  })

  const opened: Array<{ release(): Promise<void> }> = []
  afterEach(async () => { for (const source of opened.splice(0)) await source.release() })
  it('the resident and the paged XLSX reads carry the same user shapes', async () => {
    const bytes = await chartWorkbookBytes(chartXml(), { patch: zip => {
      zip.file('xl/charts/_rels/chart1.xml.rels', rels(['rUS', 'chartUserShapes', '../drawings/drawing2.xml']))
      zip.file(SHAPES, userShapesXml(relAnchor([0.5, 0.1], [0.75, 0.4], rect('FF0000', 10, 10)) + absAnchor([0.1, 0.1], emu(40), emu(40), pic('rImg'))))
      zip.file('xl/drawings/_rels/drawing2.xml.rels', rels(['rImg', 'image', '../media/image1.png']))
      zip.file('xl/media/image1.png', png('00FF00'))
    } })
    const resident = (await parseXlsx(await OfficePackage.load(bytes), { valuePolicy: 'cached' })).sheets[0].drawings?.[0].content
    const session = await openImportSession({ kind: 'blob', blob: new Blob([new Uint8Array(bytes)]) }, { docId: 'A4d', profile: 'D' })
    expect((await session.done).kind).toBe('complete')
    const source = createPagedWorkbookSource(session)
    opened.push(source)
    const store = source.store('1')
    const paged = (await store.readDrawings!({ stamp: store.stamp, signal: new AbortController().signal }, 'D')).drawings[0].content
    for (const content of [resident, paged]) {
      if (content?.kind !== 'xlsxChart') throw new Error('no chart')
      expect(content.chart.unsupported).toBeUndefined()
      expect(content.chart.userShapes?.anchors.map(a => a.kind)).toEqual(['rel', 'abs'])
      expect(content.chart.userShapes?.images.map(i => [i.width, i.height])).toEqual([[4, 4]])
    }
  })
})
