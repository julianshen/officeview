import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { getPaintables } from '../src/render/paint'
import { drawingPartContext } from '../src/drawing/parts'
import { buildXlsx } from '../src/testdata/ooxml-builders'

const emu = (px: number) => px * 9525
const shape = (id: number, color: string, x: number, y: number, w: number, h: number, rot = 0) =>
  `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="shape${id}"/></xdr:nvSpPr><xdr:spPr><a:xfrm${rot ? ` rot="${rot * 60000}"` : ''}><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></xdr:spPr></xdr:sp>`
const marker = (col: number, row: number, colOff = 0, rowOff = 0) => `<xdr:col>${col}</xdr:col><xdr:colOff>${emu(colOff)}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${emu(rowOff)}</xdr:rowOff>`
const one = (object: string, col: number, row: number, x = 0, y = 0, w = 100, h = 50) => `<xdr:oneCellAnchor><xdr:from>${marker(col, row, x, y)}</xdr:from><xdr:ext cx="${emu(w)}" cy="${emu(h)}"/>${object}<xdr:clientData/></xdr:oneCellAnchor>`
const absolute = (object: string, x: number, y: number, w: number, h: number) => `<xdr:absoluteAnchor><xdr:pos x="${emu(x)}" y="${emu(y)}"/><xdr:ext cx="${emu(w)}" cy="${emu(h)}"/>${object}<xdr:clientData/></xdr:absoluteAnchor>`

async function fixture(drawings: string, cols = '', rows = '', patch?: (zip: JSZip) => void | Promise<void>) {
  const base = await buildXlsx([{ name: 'Draw', rows: [{ r: 1, cells: [{ ref: 'A1', v: 9 }] }], cols }])
  const zip = await JSZip.loadAsync(base)
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  sheet = sheet.replace('</worksheet>', `${rows}<drawing r:id="rDraw"/></worksheet>`)
  zip.file('xl/worksheets/sheet1.xml', sheet)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rDraw" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing7.xml"/></Relationships>')
  zip.file('xl/drawings/drawing7.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${drawings}</xdr:wsDr>`)
  await patch?.(zip)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}

const group = (id: number, children: string) => `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${id}" name="group${id}"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(120)}" cy="${emu(60)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(120)}" cy="${emu(60)}"/></a:xfrm></xdr:grpSpPr>${children}</xdr:grpSp>`
const pic = (id: number, rid: string, x = 0, y = 0, w = 20, h = 20) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="picture${id}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"><a:alphaModFix amt="50000"/></a:blip><a:srcRect l="10000"/><a:stretch><a:fillRect r="10000"/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm></xdr:spPr></xdr:pic>`
const frame = (id: number, rid: string, tag = 'chart') => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="frame${id}"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(90)}" cy="${emu(50)}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/${tag}"><c:${tag} r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`

describe('worksheet drawing anchors', () => {
  test('uses physical custom markers and preserves source order on a drawing-only sheet', async () => {
    const pkg = await fixture(
      one(shape(3, 'FF0000', 0, 0, 100, 50), 1, 1, 1, 2) +
      absolute(shape(4, '0000FF', 0, 0, 80, 80), 2, 3, 80, 80),
      '<col min="1" max="1" width="5"/><col min="2" max="2" width="10.7142857"/>',
      '<row r="2" ht="22.5" customHeight="1"/>',
    )
    const doc = await parseXlsx(pkg)
    const nodes = doc.sheets[0].drawings ?? []
    expect(nodes.map(node => node.source.id)).toEqual(['3', '4'])
    expect(nodes.map(node => [node.xEmu / 9525, node.yEmu / 9525, node.widthEmu / 9525, node.heightEmu / 9525])).toEqual([[41, 22, 100, 50], [2, 3, 80, 80]])
    expect(computeMetrics(doc.sheets[0]).heightPx).toBeGreaterThanOrEqual(83)
  })

  test('expands for transformed bounds and clips huge finite drawings at bounded scale', async () => {
    const pkg = await fixture(absolute(shape(7, '00FF00', 0, 0, 100, 50, 90), 20, 30, 100, 50) + absolute(shape(8, 'FF0000', 0, 0, 10, 10), 200000, 200000, 10, 10))
    const sheet = (await parseXlsx(pkg)).sheets[0]
    const m = computeMetrics(sheet)
    expect(m.requestedDrawingBounds?.bottom).toBeGreaterThan(200000)
    expect(m.widthPx).toBeLessThanOrEqual(16384)
    expect(m.heightPx).toBeLessThanOrEqual(16384)
    expect(m.widthPx * m.heightPx).toBeLessThanOrEqual(16777216)
    const small = await fixture(absolute(shape(7, '00FF00', 0, 0, 100, 50, 90), 20, 30, 100, 50))
    expect(computeMetrics((await parseXlsx(small)).sheets[0]).heightPx).toBeGreaterThanOrEqual(105)
  })

  test('paints drawings after cells and restores caller state', async () => {
    const doc = await parseXlsx(await fixture(absolute(shape(9, 'FF0000', 0, 0, 30, 30), 2, 2, 30, 30)))
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d')
    ctx.globalAlpha = .5
    paintables[0].paint(ctx as unknown as CanvasRenderingContext2D)
    expect(ctx.globalAlpha).toBe(.5)
    const pixel = ctx.getImageData(15, 15, 1, 1).data
    expect(pixel[0]).toBeGreaterThan(240)
    expect(pixel[1]).toBeLessThan(128)
    paintables.dispose()
    // The synchronous public entry point remains usable.
    renderSheet(doc.sheets[0], ctx as unknown as CanvasRenderingContext2D)
  })

  test('two-cell anchors use hidden row and column prefixes and keep overflowing nested group children', async () => {
    const nested = group(10, group(11, shape(12, '00FF00', 130, 5, 20, 20)))
    const two = `<xdr:twoCellAnchor><xdr:from>${marker(1, 1)}</xdr:from><xdr:to>${marker(4, 4)}</xdr:to>${nested}<xdr:clientData/></xdr:twoCellAnchor>`
    const pkg = await fixture(two,
      '<col min="1" max="1" width="5"/><col min="2" max="2" hidden="1"/><col min="3" max="4" width="5"/>', '',
      async zip => {
        let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
        xml = xml.replace('</sheetData>', '<row r="2" ht="22.5" customHeight="1"/><row r="3" hidden="1"/></sheetData>')
        zip.file('xl/worksheets/sheet1.xml', xml)
      })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    const anchored = sheet.drawings![0]
    expect(sheet.drawings?.[0].source.id).toBe('10')
    expect(sheet.drawings?.[0].children?.[0].source.id).toBe('11')
    expect(sheet.drawings?.[0].children?.[0].children?.[0].source.id).toBe('12')
    expect(anchored.xEmu / 9525).toBe(40)
    expect(anchored.yEmu / 9525).toBe(20)
    expect(anchored.widthEmu / 9525).toBe(80)
    expect(anchored.heightEmu / 9525).toBe(50)
    const metrics = computeMetrics(sheet)
    expect(metrics.colWidthsPx[1]).toBe(0)
    expect(metrics.rowHeightsPx[2]).toBe(0)
    expect(metrics.widthPx).toBeGreaterThan(140)
  })

  test('deduplicates pictures across selected branches and preserves per-use crop and alpha', async () => {
    const imageCanvas = createCanvas(2, 2)
    const imageCtx = imageCanvas.getContext('2d')
    imageCtx.fillStyle = '#00ff00'; imageCtx.fillRect(0, 0, 2, 2)
    const alternate = `<mc:AlternateContent><mc:Choice Requires="unknown">${pic(99, 'rImg')}</mc:Choice><mc:Fallback>${pic(1, 'rImg')}</mc:Fallback></mc:AlternateContent>`
    const pkg = await fixture(absolute(group(2, alternate + pic(3, 'rImg', 40, 0)), 10, 10, 120, 60), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/picture.png"/></Relationships>')
      zip.file('xl/media/picture.png', imageCanvas.toBuffer('image/png'))
    })
    const doc = await parseXlsx(pkg)
    const children = doc.sheets[0].drawings?.[0].children ?? []
    expect(children.map(n => n.source.id)).toEqual(['1', '3'])
    expect(children[0].source.representation).toBe('fallback')
    expect(children.map(n => n.imageIndex)).toEqual([0, 0])
    expect(doc.images).toHaveLength(1)
    expect(children[0].image?.opacity).toBe(.5)
    expect(children[0].image?.srcRect?.l).toBe(.1)
    expect(children[0].image?.fillRect?.r).toBe(.1)
  })

  test('skips malformed optional drawings while keeping cells and valid neighboring drawing refs', async () => {
    const pkg = await fixture(absolute(shape(1, 'FF0000', 0, 0, 30, 30), 5, 5, 30, 30), '', '', async zip => {
      let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      sheet = sheet.replace('<drawing r:id="rDraw"/>', '<drawing r:id="bad"/><drawing r:id="rDraw"/>')
      zip.file('xl/worksheets/sheet1.xml', sheet)
      let rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')
      rels = rels.replace('</Relationships>', '<Relationship Id="bad" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/bad.xml"/></Relationships>')
      zip.file('xl/worksheets/_rels/sheet1.xml.rels', rels)
      zip.file('xl/drawings/bad.xml', '<broken')
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.rows[0].cells[0].value).toBe(9)
    expect(sheet.drawings?.map(n => n.source.id)).toEqual(['1'])
    expect(sheet.drawingDiagnostics?.some(issue => issue.kind === 'malformed-part')).toBe(true)
  })

  test('supported empty Choice intentionally suppresses Fallback picture', async () => {
    const alt = `<mc:AlternateContent><mc:Choice Requires="a"><a:extLst/></mc:Choice><mc:Fallback>${pic(1, 'missing')}</mc:Fallback></mc:AlternateContent>`
    const sheet = (await parseXlsx(await fixture(absolute(alt + shape(2, 'FF0000', 0, 0, 20, 20), 5, 5, 20, 20)))).sheets[0]
    expect(sheet.drawings?.map(n => n.source.id)).toEqual(['2'])
  })

  test('loads workbook-related theme and selected chart, diagram, and ink parts from drawing ownership', async () => {
    const diagram = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="31"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(60)}" cy="${emu(30)}"/></xdr:xfrm><a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const ink = `<xdr:contentPart r:id="rInk"/>`
    const drawing = absolute(frame(30, 'rChart'), 10, 10, 90, 50) + absolute(diagram, 20, 70, 60, 30) + absolute(ink, 40, 120, 40, 20)
    const pkg = await fixture(drawing, '', '', async zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', `<Relationships>
        <Relationship Id="rChart" Target="../charts/chart4.xml"/>
        <Relationship Id="rDiagram" Target="../diagrams/drawing2.xml"/>
        <Relationship Id="rInk" Target="../ink/ink3.xml"/>
      </Relationships>`)
      zip.file('xl/charts/chart4.xml', `<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`)
      zip.file('xl/diagrams/drawing2.xml', `<dsp:drawing><dsp:spTree>${shape(1, 'FF0000', 0, 0, 20, 20)}</dsp:spTree></dsp:drawing>`)
      zip.file('xl/ink/ink3.xml', '<ink><traceFormat><channel name="X" units="px"/><channel name="Y" units="px"/></traceFormat><trace>0 0, 10 10</trace></ink>')
      let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
      rels = rels.replace('</Relationships>', '<Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme5.xml"/></Relationships>')
      zip.file('xl/_rels/workbook.xml.rels', rels)
      zip.file('xl/theme/theme5.xml', '<a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme></a:themeElements></a:theme>')
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.map(n => n.source.id)).toEqual(['30', '31', 'rInk'])
    expect(sheet.drawings?.map(n => n.content?.kind)).toEqual(['chart', 'diagram', 'ink'])
    expect(sheet.drawings?.[0].content?.kind === 'chart' ? sheet.drawings[0].content.series[0].color : '').toBe('123456')
    expect(sheet.drawings?.[1].content?.kind === 'diagram' ? sheet.drawings[1].content.shapes.length : 0).toBe(1)
  })

  test('cached worksheet diagram text uses shared list defaults and direct false overrides', async () => {
    const diagram = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="34"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(80)}" cy="${emu(50)}"/></xdr:xfrm><a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const pkg = await fixture(absolute(diagram, 10, 10, 80, 50), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rDiagram" Target="../diagrams/cached.xml"/></Relationships>')
      zip.file('xl/diagrams/cached.xml', `<dsp:drawing><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(80)}" cy="${emu(50)}"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr><dsp:style><a:fontRef idx="minor"/></dsp:style><dsp:txBody><a:bodyPr vert="vert270"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400" b="1"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr b="0"/><a:t>off</a:t></a:r><a:r><a:t>on</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`)
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    const drawing = sheet.drawings?.[0].content
    const body = drawing?.kind === 'diagram' ? drawing.shapes[0].textBody : undefined
    expect(body).toMatchObject({ direction: 'vert270', paragraphs: [{ runs: [{ text: 'off', bold: false, fontSizePt: 24, fontFamily: '+mn-lt' }, { text: 'on', bold: true, fontSizePt: 24, fontFamily: '+mn-lt' }] }] })
  })

  test('ordinary worksheet shape retains fontRef defaults and direct run overrides', async () => {
    const ordinary = shape(35, 'FFFFFF', 0, 0, 100, 40)
      .replace('</xdr:spPr>', '</xdr:spPr><xdr:style><a:fontRef idx="major"><a:srgbClr val="112233"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Inherited</a:t></a:r><a:r><a:rPr><a:latin typeface="Direct"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Direct</a:t></a:r></a:p></xdr:txBody>')
    const sheet = (await parseXlsx(await fixture(absolute(ordinary, 10, 10, 100, 40)))).sheets[0]
    expect(sheet.drawings?.[0].textBody?.paragraphs[0].runs).toMatchObject([
      { text: 'Inherited', fontFamily: '+mj-lt', fontFamilyEastAsia: '+mj-ea', fontFamilyComplexScript: '+mj-cs', color: '#112233' },
      { text: 'Direct', fontFamily: 'Direct', color: '#FF0000' },
    ])
  })

  test('reports missing selected payload while preserving its next shape', async () => {
    const pkg = await fixture(absolute(frame(40, 'missing'), 1, 1, 60, 40) + absolute(shape(41, '00FF00', 0, 0, 20, 20), 70, 1, 20, 20))
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.map(n => n.source.id)).toEqual(['40', '41'])
    expect(sheet.drawingDiagnostics?.some(issue => issue.kind === 'missing-part' && issue.identity === 'missing')).toBe(true)
  })

  test('right and bottom viewport clipping retains scale under both dimension and area limits', async () => {
    const oneDimension = (await parseXlsx(await fixture(absolute(shape(50, 'FF0000', 0, 0, 10, 10), 30000, 10, 10, 10)))).sheets[0]
    const wide = computeMetrics(oneDimension)
    expect(wide.widthPx).toBe(16384)
    expect(wide.heightPx).toBeLessThan(16384)
    const bothDimensions = (await parseXlsx(await fixture(absolute(shape(51, 'FF0000', 0, 0, 10, 10), 30000, 30000, 10, 10)))).sheets[0]
    const clipped = computeMetrics(bothDimensions)
    expect(clipped.widthPx).toBe(16384)
    expect(clipped.heightPx).toBe(1024)
    expect(clipped.requestedDrawingBounds?.right).toBeGreaterThan(30000)
    expect(clipped.requestedDrawingBounds?.bottom).toBeGreaterThan(30000)
    expect(clipped.retainedDrawingBounds).toEqual({ right: 16384, bottom: 1024 })
  })

  test('expands custom cubic geometry to include off-rectangle control bounds', async () => {
    const custom = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="60"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(20)}" cy="${emu(20)}"/></a:xfrm><a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="200" y="0"/><a:pt x="200" y="100"/><a:pt x="100" y="100"/></a:cubicBezTo></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr></xdr:sp>`
    const m = computeMetrics((await parseXlsx(await fixture(absolute(custom, 60, 10, 20, 20)))).sheets[0])
    expect(m.widthPx).toBeGreaterThanOrEqual(100)
  })

  test('malformed geometry and optional relationships do not suppress valid neighboring shapes', async () => {
    const malformed = shape(70, 'FF0000', 0, 0, 20, 20).replace(`cx="${emu(20)}"`, 'cx="NaN"')
    const pkg = await fixture(absolute(malformed + shape(71, '00FF00', 0, 0, 20, 20), 2, 2, 20, 20), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><broken')
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.map(node => node.source.id)).toEqual(['70', '71'])
    expect(sheet.drawings?.[0].transformValid).toBe(false)
    expect(sheet.drawings?.[1].transformValid).toBe(true)
    expect(sheet.drawingDiagnostics?.some(issue => issue.kind === 'malformed-part' && issue.reason === 'invalid-relationship-xml')).toBe(true)
  })

  test('selects native picture fallback inside a graphic frame without using unselected media', async () => {
    const wrapped = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="80"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(30)}" cy="${emu(30)}"/></xdr:xfrm><a:graphic><a:graphicData><mc:AlternateContent><mc:Choice Requires="future">${pic(81, 'missing')}</mc:Choice><mc:Fallback>${pic(82, 'rImg')}</mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></xdr:graphicFrame>`
    const canvas = createCanvas(2, 2)
    const pkg = await fixture(absolute(wrapped, 10, 10, 30, 30), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rImg" Target="../media/pic.png"/></Relationships>')
      zip.file('xl/media/pic.png', canvas.toBuffer('image/png'))
    })
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].drawings?.[0].children?.map(n => n.source.id)).toEqual(['82'])
    expect(doc.images).toHaveLength(1)
    expect(doc.sheets[0].drawingDiagnostics?.some(issue => issue.identity === 'missing')).toBe(false)
  })

  test('drawing-only second worksheet shares first-use media index and keeps its drawing owner', async () => {
    const canvas = createCanvas(2, 2)
    const drawing = absolute(pic(91, 'rImg'), 2, 3, 20, 20)
    const pkg = await fixture(drawing, '', '', async zip => {
      const firstDrawing = await zip.file('xl/drawings/drawing7.xml')!.async('string')
      zip.file('xl/drawings/drawing8.xml', firstDrawing.replace('id="91"', 'id="92"'))
      const imageRels = '<Relationships><Relationship Id="rImg" Target="../media/shared.png"/></Relationships>'
      zip.file('xl/drawings/_rels/drawing7.xml.rels', imageRels)
      zip.file('xl/drawings/_rels/drawing8.xml.rels', imageRels)
      zip.file('xl/media/shared.png', canvas.toBuffer('image/png'))
      let workbook = await zip.file('xl/workbook.xml')!.async('string')
      workbook = workbook.replace('</sheets>', '<sheet name="EmptyDrawing" sheetId="2" r:id="rSheet2"/></sheets>')
      zip.file('xl/workbook.xml', workbook)
      let workbookRels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
      workbookRels = workbookRels.replace('</Relationships>', '<Relationship Id="rSheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
      zip.file('xl/_rels/workbook.xml.rels', workbookRels)
      zip.file('xl/worksheets/sheet2.xml', '<worksheet><sheetData/><drawing r:id="rDraw2"/></worksheet>')
      zip.file('xl/worksheets/_rels/sheet2.xml.rels', '<Relationships><Relationship Id="rDraw2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing8.xml"/></Relationships>')
    })
    const doc = await parseXlsx(pkg)
    expect(doc.sheets).toHaveLength(2)
    expect(doc.sheets[1].rows).toHaveLength(0)
    expect(doc.sheets.map(s => s.drawings?.[0].imageIndex)).toEqual([0, 0])
    expect(doc.sheets[1].drawings?.[0].source.partPath).toBe('xl/drawings/drawing8.xml')
    expect(doc.images).toHaveLength(1)
    const pages = await getPaintables(doc)
    expect(pages).toHaveLength(2)
    expect(pages[1].spec.widthPx).toBeGreaterThanOrEqual(22)
    pages.dispose()
  })

  test('unreadable optional picture bytes leave its later neighbor available', async () => {
    const canvas = createCanvas(2, 2)
    const pkg = await fixture(absolute(pic(101, 'rImg') + shape(102, '00FF00', 30, 0, 20, 20), 1, 1, 60, 30), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rImg" Target="../media/bad.png"/></Relationships>')
      zip.file('xl/media/bad.png', canvas.toBuffer('image/png'))
    })
    const original = pkg.bytes.bind(pkg)
    pkg.bytes = async (path, limit) => path === 'xl/media/bad.png' ? Promise.reject(new Error('corrupt optional media')) : original(path, limit)
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.map(node => node.source.id)).toEqual(['101', '102'])
    expect(sheet.drawings?.[0].image).toBeUndefined()
    expect(sheet.drawingDiagnostics?.some(issue => issue.kind === 'malformed-part' && issue.identity === 'xl/media/bad.png')).toBe(true)
  })

  test('allows 64 actual group levels plus a leaf and identifies the 65th group limit', async () => {
    const nesting = (count: number) => Array.from({ length: count }, (_, i) => i).reduceRight((child, i) => group(100 + i, child), shape(200, 'FF0000', 0, 0, 10, 10))
    const allowed = (await parseXlsx(await fixture(absolute(nesting(64), 1, 1, 120, 60)))).sheets[0]
    let leaf = allowed.drawings?.[0]
    for (let i = 0; i < 64; i++) leaf = leaf?.children?.[0]
    expect(leaf?.source.id).toBe('200')
    const limited = (await parseXlsx(await fixture(absolute(nesting(65), 1, 1, 120, 60)))).sheets[0]
    expect(limited.drawingDiagnostics?.some(issue => issue.kind === 'group-depth' && issue.identity === '164' && issue.limit === 64)).toBe(true)
  })

  test('gives repeated anchor kinds distinct source tree paths', async () => {
    const sheet = (await parseXlsx(await fixture(absolute(shape(301, 'FF0000', 0, 0, 10, 10), 1, 1, 10, 10) + absolute(shape(302, '00FF00', 0, 0, 10, 10), 20, 1, 10, 10)))).sheets[0]
    expect(sheet.drawings?.map(node => node.source.treePath)).toEqual(['absoluteAnchor[0]/sp[0]', 'absoluteAnchor[1]/sp[0]'])
  })

  test('paints watermark under an opaque drawing and that drawing over formatted cell borders', async () => {
    const pkg = await fixture(absolute(shape(401, 'FF0000', 0, 0, 64, 20), 0, 0, 64, 20), '', '', async zip => {
      let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      xml = xml.replace('<c r="A1">', '<c r="A1" s="1">')
      zip.file('xl/worksheets/sheet1.xml', xml)
    })
    const doc = await parseXlsx(pkg)
    const plain = await getPaintables(doc)
    const marked = await getPaintables(doc, { watermark: { text: 'TEST', placement: 'center', opacity: 1, color: '#000000', fontSizePt: 40 } })
    const render = (paint: typeof plain) => {
      const canvas = createCanvas(paint[0].spec.widthPx, paint[0].spec.heightPx)
      paint[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
      return canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data
    }
    const pixels = render(plain), stamped = render(marked)
    expect(pixels).toEqual(stamped)
    const edge = 4 * (10 * plain[0].spec.widthPx)
    expect(pixels[edge]).toBeGreaterThan(230)
    expect(pixels[edge + 1]).toBeLessThan(30)
    plain.dispose(); marked.dispose()
  })

  test('retains carrier identity and chosen vector fallback provenance', async () => {
    const graphic = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="501" name="chart carrier"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(80)}" cy="${emu(40)}"/></xdr:xfrm><a:graphic><a:graphicData><mc:AlternateContent><mc:Choice Requires="future"><c:chart r:id="missing"/></mc:Choice><mc:Fallback><c:chart r:id="rChart"/></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></xdr:graphicFrame>`
    const pkg = await fixture(absolute(graphic, 1, 1, 80, 40), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/chart9.xml"/></Relationships>')
      zip.file('xl/charts/chart9.xml', '<chartSpace><chart><plotArea><barChart><barDir val="col"/><ser><val><numLit><pt idx="0"><v>2</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>')
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.[0].content?.kind).toBe('chart')
    expect(sheet.drawings?.[0].source).toMatchObject({ id: '501', name: 'chart carrier', representation: 'fallback', partPath: 'xl/drawings/drawing7.xml' })
    expect(sheet.drawingDiagnostics?.some(issue => issue.identity === 'missing')).toBe(false)
  })

  test('finds InkML when compatibility wraps graphic, graphicData, or nested contentPart', async () => {
    const alt = (body: string) => `<mc:AlternateContent><mc:Choice Requires="a">${body}</mc:Choice><mc:Fallback/></mc:AlternateContent>`
    const ref = '<xdr:contentPart r:id="rInk"/>'
    const data = (body: string) => `<a:graphicData uri="ink">${body}</a:graphicData>`
    const graphic = (body: string) => `<a:graphic>${body}</a:graphic>`
    for (const body of [graphic(alt(data(ref))), alt(graphic(data(ref))), graphic(data(alt(ref))), graphic(data(alt(alt(ref))))]) {
      const wrapped = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="601"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(100)}"/></xdr:xfrm>${body}</xdr:graphicFrame>`
      const pkg = await fixture(absolute(wrapped, 0, 0, 100, 100), '', '', zip => {
        zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rInk" Target="../ink/ink.xml"/></Relationships>')
        zip.file('xl/ink/ink.xml', '<ink><traceFormat><channel name="X" units="px"/><channel name="Y" units="px"/></traceFormat><brush id="b"><brushProperty name="width" value="5" units="px"/><brushProperty name="color" value="#0000FF"/></brush><trace brushRef="#b">0 0, 100 100</trace></ink>')
      })
      const sheet = (await parseXlsx(pkg)).sheets[0]
      expect(sheet.drawings?.[0].content?.kind, body).toBe('ink')
      const m = computeMetrics(sheet), canvas = createCanvas(m.widthPx, m.heightPx)
      renderSheet(sheet, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, m)
      const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data
      let blue = 0
      for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 20 && pixels[i + 1] < 20 && pixels[i + 2] > 240) blue++
      expect(blue, body).toBeGreaterThan(100)
    }
  })

  test('recovers native picture fallback after an attempted horizontal-chart Choice at either MC level', async () => {
    const canvas = createCanvas(1, 1)
    canvas.getContext('2d').fillStyle = '#0000FF'; canvas.getContext('2d').fillRect(0, 0, 1, 1)
    const chart = '<c:chart r:id="rChart"/>'
    const frameBody = (body: string) => `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="701"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(100)}"/></xdr:xfrm><a:graphic><a:graphicData>${body}</a:graphicData></a:graphic></xdr:graphicFrame>`
    const alternate = (choice: string, fallback: string) => `<mc:AlternateContent><mc:Choice Requires="a">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    for (const markup of [alternate(frameBody(chart), pic(702, 'rImg')), frameBody(alternate(chart, pic(702, 'rImg')))]) {
      const pkg = await fixture(absolute(markup, 0, 0, 100, 100), '', '', zip => {
        zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/hbar.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
        zip.file('xl/charts/hbar.xml', '<chartSpace><chart><plotArea><barChart><barDir val="bar"/><ser><val><numLit><pt idx="0"><v>2</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>')
        zip.file('xl/media/blue.png', canvas.toBuffer('image/png'))
      })
      const sheet = (await parseXlsx(pkg)).sheets[0]
      const selected = sheet.drawings?.[0].children?.[0] ?? sheet.drawings?.[0]
      expect(selected?.source.id, markup).toBe('702')
      expect(selected?.source.representation, markup).toBe('fallback')
      expect(selected?.imageIndex, markup).toBe(0)
      expect(drawingPartContext(pkg).nodes, markup).toBeLessThanOrEqual(2)
    }
  })

  test('keeps a supported column-chart Choice selected over its picture fallback', async () => {
    const chartChoice = `<mc:AlternateContent><mc:Choice Requires="a">${frame(711, 'rChart')}</mc:Choice><mc:Fallback>${pic(712, 'missing')}</mc:Fallback></mc:AlternateContent>`
    const pkg = await fixture(absolute(chartChoice, 0, 0, 90, 50), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/column.xml"/></Relationships>')
      zip.file('xl/charts/column.xml', '<chartSpace><chart><plotArea><barChart><barDir val="col"/><ser><val><numLit><pt idx="0"><v>2</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>')
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.[0].source.id).toBe('711')
    expect(sheet.drawings?.[0].source.representation).toBe('choice')
    expect(sheet.drawings?.[0].content?.kind).toBe('chart')
    expect(sheet.drawingDiagnostics?.some(issue => issue.identity === 'missing')).toBe(false)
  })

  test('transformed full-circle custom arc stays inside requested viewport', async () => {
    const arc = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="801"/></xdr:nvSpPr><xdr:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="${emu(20)}" cy="${emu(20)}"/></a:xfrm><a:custGeom><a:pathLst><a:path w="20" h="20"><a:moveTo><a:pt x="100" y="50"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="0" swAng="21600000"/><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr></xdr:sp>`
    const sheet = (await parseXlsx(await fixture(absolute(arc, 100, 100, 20, 20)))).sheets[0]
    expect(computeMetrics(sheet).widthPx).toBeGreaterThanOrEqual(160)
  })

  test('cached diagram children beyond their frame expand the worksheet viewport', async () => {
    const frameXml = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="901"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(20)}" cy="${emu(20)}"/></xdr:xfrm><a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    const pkg = await fixture(absolute(frameXml, 10, 10, 20, 20), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rDiagram" Target="../diagrams/cache.xml"/></Relationships>')
      zip.file('xl/diagrams/cache.xml', `<drawing><spTree>${shape(1, 'FF0000', 100, 50, 40, 40)}</spTree></drawing>`)
    })
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.[0].content?.kind).toBe('diagram')
    const m = computeMetrics(sheet)
    expect(m.widthPx).toBeGreaterThanOrEqual(150)
    expect(m.heightPx).toBeGreaterThanOrEqual(100)
  })

  test('keeps drawing-free hidden-width metrics and sanitizes invalid drawing-bearing dimensions', async () => {
    const hidden = await fixture('', '<col min="1" max="1" width="5" hidden="1"/>', '', async zip => {
      let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      sheet = sheet.replace('<drawing r:id="rDraw"/>', '')
      zip.file('xl/worksheets/sheet1.xml', sheet)
    })
    expect(computeMetrics((await parseXlsx(hidden)).sheets[0]).colWidthsPx[0]).toBe(64)
    for (const [cols, row] of [['<col min="1" max="1" width="NaN"/>', ''], ['', '<row r="1" ht="NaN" customHeight="1"/>']] as const) {
      const pkg = await fixture(absolute(shape(1001, 'FF0000', 0, 0, 20, 20), 0, 0, 20, 20), cols, row, async zip => {
        if (row) {
          let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
          sheet = sheet.replace('<row r="1">', '<row r="1" ht="NaN" customHeight="1">')
          zip.file('xl/worksheets/sheet1.xml', sheet)
        }
      })
      const sheet = (await parseXlsx(pkg)).sheets[0], m = computeMetrics(sheet)
      expect(sheet.drawings).toHaveLength(1)
      expect(Number.isInteger(m.widthPx) && m.widthPx > 0).toBe(true)
      expect(Number.isInteger(m.heightPx) && m.heightPx > 0).toBe(true)
    }
  })

  test('nested supported empty Choices suppress all rejected fallback payloads but keep a neighbor', async () => {
    const alt = (choice: string, fallback: string) => `<mc:AlternateContent><mc:Choice Requires="a">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const failed = frame(1101, 'rChart')
    const nested = alt(alt('<a:extLst/>', failed), pic(1102, 'rImg'))
    const pkg = await fixture(absolute(nested + shape(1103, '00FF00', 0, 0, 10, 10), 0, 0, 100, 100), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/hbar.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
      zip.file('xl/charts/hbar.xml', '<chartSpace><chart><plotArea><barChart><barDir val="bar"/></barChart></plotArea></chart></chartSpace>')
      zip.file('xl/media/blue.png', createCanvas(1, 1).toBuffer('image/png'))
    })
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].drawings?.map(node => node.source.id)).toEqual(['1103'])
    expect(doc.images).toHaveLength(0)
    expect(drawingPartContext(pkg).nodes).toBe(1)
  })

  test('finds a native picture fallback below a selected graphicData wrapper', async () => {
    const inner = `<mc:AlternateContent><mc:Choice Requires="a"><c:chart r:id="rChart"/></mc:Choice><mc:Fallback>${pic(1202, 'rImg')}</mc:Fallback></mc:AlternateContent>`
    const outer = `<mc:AlternateContent><mc:Choice Requires="a"><a:graphicData>${inner}</a:graphicData></mc:Choice><mc:Fallback/></mc:AlternateContent>`
    const wrapped = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1201"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(100)}" cy="${emu(100)}"/></xdr:xfrm><a:graphic>${outer}</a:graphic></xdr:graphicFrame>`
    const pkg = await fixture(absolute(wrapped, 0, 0, 100, 100), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/hbar.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
      zip.file('xl/charts/hbar.xml', '<chartSpace><chart><plotArea><barChart><barDir val="bar"/></barChart></plotArea></chart></chartSpace>')
      zip.file('xl/media/blue.png', createCanvas(1, 1).toBuffer('image/png'))
    })
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].drawings?.[0].children?.map(node => node.source.id)).toEqual(['1202'])
    expect(doc.sheets[0].drawings?.[0].children?.[0].source.representation).toBe('fallback')
    expect(doc.images).toHaveLength(1)
    expect(drawingPartContext(pkg).nodes).toBe(2)
  })

  test('invalid row index cannot corrupt a valid drawing-bearing public viewport', async () => {
    const pkg = await fixture(absolute(shape(1301, 'FF0000', 0, 0, 20, 20), 0, 0, 20, 20), '', '', async zip => {
      let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      sheet = sheet.replace('<row r="1">', '<row r="NaN">')
      zip.file('xl/worksheets/sheet1.xml', sheet)
    })
    const doc = await parseXlsx(pkg), metrics = computeMetrics(doc.sheets[0])
    expect(Number.isSafeInteger(metrics.widthPx) && metrics.widthPx > 0).toBe(true)
    expect(Number.isSafeInteger(metrics.heightPx) && metrics.heightPx > 0).toBe(true)
    const pages = await getPaintables(doc)
    expect(pages[0].spec.widthPx).toBe(metrics.widthPx)
    pages.dispose()
  })

  test('a supported Choice can hold 3000 native rectangles without an inspection-budget rejection', async () => {
    const objects = Array.from({ length: 3000 }, (_, i) => shape(2000 + i, 'FF0000', 0, 0, 1, 1)).join('')
    const choice = `<mc:AlternateContent><mc:Choice Requires="a">${group(1401, objects)}</mc:Choice><mc:Fallback/></mc:AlternateContent>`
    const pkg = await fixture(absolute(choice, 0, 0, 120, 60))
    const sheet = (await parseXlsx(pkg)).sheets[0]
    expect(sheet.drawings?.[0].children).toHaveLength(3000)
    expect(drawingPartContext(pkg).nodes).toBe(3001)
    expect(sheet.drawingDiagnostics?.some(issue => issue.reason === 'source-probe-limit')).toBe(false)
  })

  test('outer fallback follows a selected graphic wrapper whose chart fails without an inner fallback', async () => {
    const alternate = (choice: string, fallback: string) => `<mc:AlternateContent><mc:Choice Requires="a">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    const bad = frame(1501, 'rChart')
    const wrapped = bad.replace('<a:graphic>', '<mc:AlternateContent><mc:Choice Requires="a"><a:graphic>').replace('</a:graphic>', '</a:graphic></mc:Choice></mc:AlternateContent>')
    const empty = bad.replace('<c:chart r:id="rChart"/>', '').replace('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">', '<a:graphicData>')
    for (const [choice, expected] of [[bad, '1502'], [wrapped, '1502'], [empty, '1501']] as const) {
      const pkg = await fixture(absolute(alternate(choice, pic(1502, 'rImg')), 0, 0, 100, 100), '', '', zip => {
        zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/hbar.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
        zip.file('xl/charts/hbar.xml', '<chartSpace><chart><plotArea><barChart><barDir val="bar"/></barChart></plotArea></chart></chartSpace>')
        zip.file('xl/media/blue.png', createCanvas(1, 1).toBuffer('image/png'))
      })
      const doc = await parseXlsx(pkg)
      expect(doc.sheets[0].drawings?.[0].source.id, choice).toBe(expected)
      expect(doc.images, choice).toHaveLength(expected === '1502' ? 1 : 0)
    }
  })

  test('MC preflight preserves the selected graphicData carrier URI when assessing a chart', async () => {
    const alt = (choice: string) => `<mc:AlternateContent><mc:Choice Requires="a">${choice}</mc:Choice><mc:Fallback>${pic(1602, 'rImg')}</mc:Fallback></mc:AlternateContent>`
    const supported = frame(1601, 'rChart')
    const unsupported = supported.replace('http://schemas.openxmlformats.org/drawingml/2006/chart', 'http://schemas.microsoft.com/office/drawing/2014/chartex')
    const wrapped = unsupported.replace('<a:graphicData ', '<mc:AlternateContent><mc:Choice Requires="a"><a:graphicData ').replace('</a:graphicData>', '</a:graphicData></mc:Choice></mc:AlternateContent>')
    const empty = unsupported.replace('<c:chart r:id="rChart"/>', '')
    for (const [choice, id, imageCount] of [[supported, '1601', 0], [unsupported, '1602', 1], [wrapped, '1602', 1], [empty, '1601', 0]] as const) {
      const pkg = await fixture(absolute(alt(choice), 0, 0, 90, 50), '', '', zip => {
        zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/column.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
        zip.file('xl/charts/column.xml', '<chartSpace><chart><plotArea><barChart><barDir val="col"/><ser><val><numLit><pt idx="0"><v>2</v></pt></numLit></val></ser></barChart></plotArea></chart></chartSpace>')
        zip.file('xl/media/blue.png', createCanvas(1, 1).toBuffer('image/png'))
      })
      const doc = await parseXlsx(pkg)
      expect(doc.sheets[0].drawings?.[0].source.id, choice).toBe(id)
      expect(doc.images, choice).toHaveLength(imageCount)
      if (id === '1601' && choice === supported) expect(doc.sheets[0].drawings?.[0].content?.kind).toBe('chart')
    }
    const nested = unsupported.replace('<c:chart r:id="rChart"/>', `<mc:AlternateContent><mc:Choice Requires="a"><c:chart r:id="rChart"/></mc:Choice><mc:Fallback>${pic(1603, 'rImg')}</mc:Fallback></mc:AlternateContent>`)
    const nestedPkg = await fixture(absolute(nested, 0, 0, 90, 50), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rChart" Target="../charts/column.xml"/><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
      zip.file('xl/charts/column.xml', '<chartSpace><chart><plotArea><barChart><barDir val="col"/></barChart></plotArea></chart></chartSpace>')
      zip.file('xl/media/blue.png', createCanvas(1, 1).toBuffer('image/png'))
    })
    const nestedDoc = await parseXlsx(nestedPkg)
    expect(nestedDoc.sheets[0].drawings?.[0].children?.[0].source.id).toBe('1603')
    expect(nestedDoc.images).toHaveLength(1)
  })

  test('one-cell anchors paint pictures and preset shapes with no optional local transform', async () => {
    const picture = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1701" name="anchor-only"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rImg"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>`
    const preset = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1702" name="anchor-only-shape"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr></xdr:sp>`
    const blue = createCanvas(1, 1)
    blue.getContext('2d').fillStyle = '#0000ff'
    blue.getContext('2d').fillRect(0, 0, 1, 1)
    const malformed = shape(1703, '00FF00', 0, 0, 20, 20).replace('<a:xfrm>', '<a:xfrm rot="NaN">')
    const pkg = await fixture(one(picture, 0, 0, 100, 40, 20, 20) + one(preset, 0, 0, 130, 40, 20, 20) + one(malformed, 0, 0, 400, 40, 20, 20), '', '', zip => {
      zip.file('xl/drawings/_rels/drawing7.xml.rels', '<Relationships><Relationship Id="rImg" Target="../media/blue.png"/></Relationships>')
      zip.file('xl/media/blue.png', blue.toBuffer('image/png'))
    })
    const doc = await parseXlsx(pkg)
    expect(doc.sheets[0].drawings?.map(node => node.transformValid)).toEqual([true, true, false])
    const paintables = await getPaintables(doc)
    try {
      expect(paintables[0].spec.widthPx).toBe(150)
      expect(paintables[0].spec.heightPx).toBe(60)
      const canvas = createCanvas(150, 60)
      paintables[0].paint(canvas.getContext('2d') as unknown as CanvasRenderingContext2D)
      const ctx = canvas.getContext('2d')
      expect(Array.from(ctx.getImageData(105, 45, 1, 1).data).slice(0, 3)).toEqual([0, 0, 255])
      expect(Array.from(ctx.getImageData(135, 45, 1, 1).data).slice(0, 3)).toEqual([255, 0, 0])
    } finally { paintables.dispose() }
  })

  test('drawing viewport expansion keeps grid edges at the physical sheet boundary', async () => {
    const pkg = await fixture(one(shape(1710, '0000FF', 0, 0, 20, 20), 0, 0, 100, 40, 20, 20))
    const sheet = (await parseXlsx(pkg)).sheets[0]
    const metrics = computeMetrics(sheet)
    expect([metrics.widthPx, metrics.heightPx]).toEqual([120, 60])
    expect([metrics.colWidthsPx[0], metrics.rowHeightsPx[0]]).toEqual([64, 20])
    const canvas = createCanvas(metrics.widthPx, metrics.heightPx)
    const ctx = canvas.getContext('2d')
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, metrics)
    expect(Array.from(ctx.getImageData(64, 10, 1, 1).data).slice(0, 3)).toEqual([208, 208, 208])
    expect(Array.from(ctx.getImageData(10, 20, 1, 1).data).slice(0, 3)).toEqual([208, 208, 208])
    expect(Array.from(ctx.getImageData(64, 30, 1, 1).data).slice(0, 3)).toEqual([255, 255, 255])
    expect(Array.from(ctx.getImageData(80, 20, 1, 1).data).slice(0, 3)).toEqual([255, 255, 255])
  })
})
