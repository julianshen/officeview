import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { collectDocImages, layoutDocx, type MeasureFn } from '../src/docx/layout'
import type { DocxDocument, DocxParagraph, DocxTable } from '../src/docx/types'
const fixed: MeasureFn = text => text.length * 8
const p = (text = ''): DocxParagraph => ({ runs: text ? [{ text }] : [], images: [], align: 'left' })
const table = (height = 1440): DocxTable => ({ gridColsTwips: [3000], cellMargins: { topTwips: 0, bottomTwips: 0, leftTwips: 0, rightTwips: 0 }, rows: [{ heightTwips: height, heightRule: 'exact', cells: [{ gridSpan: 1, paragraphs: [p('HEADER')], fill: '00AA00' }] }] })
function model(count = 1): DocxDocument {
 const paras = Array.from({ length: count }, () => p('BODY'))
 return { defaultFontFamily: 'Calibri', defaultFontSizePt: 11, styleDefaults: new Map(), sections: [{ margins: { topTwips: 1440, bottomTwips: 1440, leftTwips: 1440, rightTwips: 1440, headerTwips: 720, footerTwips: 720, gutterTwips: 0 }, pageSize: { widthTwips: 12240, heightTwips: 15840, orientation: 'portrait' }, paragraphs: paras, blocks: paras.map(paragraph => ({ kind: 'p', paragraph })) }] }
}
const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const para = (text: string, props = '') => `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
async function parse(body: string, styles = '', header?: string, sectionProps = '', footer?: string) {
 const zip = new JSZip()
 zip.file('word/document.xml', `<w:document ${ns}><w:body>${body}<w:sectPr>${header !== undefined ? '<w:headerReference w:type="default" r:id="h"/>' : ''}${footer !== undefined ? '<w:footerReference w:type="default" r:id="f"/>' : ''}${sectionProps}</w:sectPr></w:body></w:document>`)
 zip.file('word/styles.xml', `<w:styles ${ns}>${styles}</w:styles>`)
 if (header !== undefined) {
  zip.file('word/header1.xml', `<w:hdr ${ns}>${header}</w:hdr>`)
 }
 if (footer !== undefined) zip.file('word/footer1.xml', `<w:ftr ${ns}>${footer}</w:ftr>`)
 if (header !== undefined || footer !== undefined) zip.file('word/_rels/document.xml.rels', `<Relationships>${header !== undefined ? '<Relationship Id="h" Target="header1.xml"/>' : ''}${footer !== undefined ? '<Relationship Id="f" Target="footer1.xml"/>' : ''}</Relationships>`)
 return parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
}
describe('source driven DOCX reference flow', () => {
 test('document spacing defaults cascade field by field and direct zero survives', async () => {
  const doc = await parse(para('A') + para('B', '<w:spacing w:after="0"/>'), '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:before="30" w:after="200" w:line="288"/></w:pPr></w:pPrDefault></w:docDefaults>')
  expect(doc.sections[0].paragraphs[0]).toMatchObject({ spacingBeforeTwips: 30, spacingAfterTwips: 200, lineSpacing: { rule: 'auto', value: 288 } })
  expect(doc.sections[0].paragraphs[1].spacingAfterTwips).toBe(0)
  const pages = layoutDocx(doc, fixed)
  expect(pages[0].lines[1].yPx - pages[0].lines[0].yPx).toBeCloseTo(11 * 1.35 * (96 / 72) * 1.2 * 0.9 + 200 / 15 + 2)
 })
 test('controlled header cells and paragraphs retain ordered blocks', async () => {
  const hdr = '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:sdt><w:sdtContent><w:tr><w:sdt><w:sdtContent><w:tc><w:sdt><w:sdtContent>' + para('CELL') + '</w:sdtContent></w:sdt></w:tc></w:sdtContent></w:sdt></w:tr></w:sdtContent></w:sdt></w:tbl><w:sdt><w:sdtContent>' + para('TAIL') + '</w:sdtContent></w:sdt>'
  const doc = await parse(para('BODY'), '', hdr)
  expect(doc.sections[0].headerBlocks?.map(b => b.kind)).toEqual(['table', 'p'])
  expect(doc.sections[0].headerBlocks?.[0]).toMatchObject({ table: { rows: [{ cells: [{ paragraphs: [{ runs: [{ text: 'CELL' }] }] }] }] } })
 })
 test('tall header reserves table and trailing paragraph before body flow', () => {
  const doc = model()
  doc.sections[0].headerBlocks = [{ kind: 'table', table: table() }, { kind: 'p', paragraph: p() }]
  const page = layoutDocx(doc, fixed)[0]
  expect(page.lines[0].yPx).toBeCloseTo(48 + 96 + 11 * 1.35 * (96 / 72) * 0.9)
 })
 test('explicit empty block lists override stale legacy content and images', () => {
  const doc = model()
  doc.sections[0].header = [{ ...p('STALE'), images: [{ data: new Uint8Array(), widthEmu: 100, heightEmu: 100 }] }]
  doc.sections[0].headerBlocks = []
  expect(collectDocImages(doc)).toHaveLength(0)
  expect(layoutDocx(doc, fixed)[0].header?.paragraphs).toEqual([])
 })
 test('cell before and final after spacing includes empty paragraphs; tiny exact rows stay exact', () => {
  const doc = model()
  const t = table(15)
  t.rows.push({ cells: [{ gridSpan: 1, paragraphs: [{ ...p(), spacingBeforeTwips: 150, spacingAfterTwips: 300 }] }] })
  doc.sections[0].blocks = [{ kind: 'table', table: t }]
  const rows = layoutDocx(doc, fixed)[0].tables[0].rows
  expect(rows[0].heightPx).toBe(1)
  expect(rows[1].heightPx).toBeCloseTo(10 + 11 * 1.35 * (96 / 72) * 0.9 + 20)
 })
})

describe('conditional styles and bounded repeated content', () => {
 test('table defaults, bands, corners, paragraph style and direct properties cascade immutably', async () => {
  const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:b/><w:color w:val="111111"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="288"/></w:pPr></w:pPrDefault></w:docDefaults>' +
   '<w:style w:type="table" w:styleId="base"><w:rPr><w:color w:val="222222"/></w:rPr><w:pPr><w:spacing w:before="40"/></w:pPr><w:tblPr><w:tblCellMar><w:left w:w="30"/><w:top w:w="15"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:color="123456" w:sz="4"/></w:tblBorders></w:tblPr><w:tcPr><w:shd w:fill="EEEEEE"/></w:tcPr><w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:fill="555555"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr></w:tblStylePr></w:style>' +
   '<w:style w:type="table" w:styleId="derived"><w:basedOn w:val="base"/><w:tblPr><w:tblCellMar><w:right w:w="60"/></w:tblCellMar><w:tblBorders><w:top w:sz="8"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="nwCell"><w:tcPr><w:shd w:fill="AA0000"/><w:tcMar><w:bottom w:w="45"/></w:tcMar></w:tcPr></w:tblStylePr></w:style>' +
   '<w:style w:type="paragraph" w:styleId="p"><w:rPr><w:color w:val="00AA00"/><w:b w:val="0"/></w:rPr><w:pPr><w:spacing w:after="80"/></w:pPr></w:style>'
  const cell = (props = '', content = para('A')) => `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`
  const tbl = '<w:tbl><w:tblPr><w:tblStyle w:val="derived"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:lastColumn="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>' + cell() + cell('<w:gridSpan w:val="2"/>', para('B', '<w:pStyle w:val="p"/><w:spacing w:after="0"/>')) + '</w:tr><w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>' + cell('<w:tcMar><w:left w:w="0"/></w:tcMar><w:tcBorders><w:top w:val="nil"/></w:tcBorders>') + '</w:tr></w:tbl>'
  const doc = await parse(tbl, styles)
  const block = doc.sections[0].blocks[0]
  if (block.kind !== 'table') throw Error('table')
  const t = block.table
  expect(t.borders?.top).toMatchObject({ style: 'single', color: '123456', widthPt: 1 })
  expect(t.rows[0].cells[0]).toMatchObject({ fill: 'AA0000', margins: { leftTwips: 30, rightTwips: 60, topTwips: 15, bottomTwips: 45 } })
  expect(t.rows[0].cells[0].paragraphs[0].runs[0].color).toBe('FFFFFF')
  expect(t.rows[0].cells[1].paragraphs[0]).toMatchObject({ spacingBeforeTwips: 40, spacingAfterTwips: 0, runs: [{ color: '00AA00', bold: false }] })
  expect(t.rows[1].cells[0]).toMatchObject({ fill: 'DDDDDD', margins: { leftTwips: 0, rightTwips: 60 } })
  expect(t.rows[1].cells[0].borders).toHaveProperty('top', undefined)
  const page = layoutDocx(doc, fixed)[0]
  expect(page.tables[0].rows[1].cells[0].xPx).toBe(100)
  const again = await parse(tbl, styles)
  expect(again).toEqual(doc)
 })
 test('disabled regions do not leak; grid spans and ragged starts own columns', async () => {
  const styles = '<w:style w:type="table" w:styleId="s"><w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:fill="AA0000"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="lastCol"><w:tcPr><w:shd w:fill="00AA00"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="0000AA"/></w:tcPr></w:tblStylePr></w:style>'
  const doc = await parse('<w:tbl><w:tblPr><w:tblStyle w:val="s"/><w:tblLook w:firstColumn="1" w:lastColumn="1" w:firstRow="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="900"/><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc>' + para('MID') + '</w:tc><w:tc>' + para('LAST') + '</w:tc></w:tr></w:tbl>', styles)
  const block = doc.sections[0].blocks[0]
  if (block.kind !== 'table') throw Error('table')
  expect(block.table.rows[0].cells[0].fill).toBeUndefined()
  expect(block.table.rows[0].cells[1].fill).toBe('00AA00')
 })
 test('narrow field cells remeasure actual whole-document NUMPAGES before reserving flow', () => {
  const doc = model(60)
  doc.sections[0].pageSize.heightTwips = 2880
  doc.sections[0].margins.bottomTwips = 720
  doc.sections[0].margins.topTwips = 720
  const t = table()
  t.gridColsTwips = [450]
  t.rows[0].heightTwips = undefined
  t.rows[0].heightRule = 'auto'
  t.rows[0].cells[0].paragraphs = [{ ...p(), runs: [{ text: '0', field: 'NUMPAGES' }, { text: ' x' }] }]
  doc.sections[0].headerBlocks = [{ kind: 'table', table: t }]
  const pages = layoutDocx(doc, fixed)
  expect(pages.length).toBeGreaterThan(9)
  expect(pages[0].lines[0].yPx).toBeCloseTo(48 + 2 * 19.8)
  expect(pages.every(page => !page.diagnostics?.includes('field-layout-nonconvergence'))).toBe(true)
  expect(t.rows[0].cells[0].paragraphs[0].runs[0].text).toBe('0')
 })
 test('oversized combined repeated bounds terminate and every body row progresses', () => {
  const doc = model(5)
  doc.sections[0].headerBlocks = [{ kind: 'table', table: table(30000) }]
  doc.sections[0].footerBlocks = [{ kind: 'table', table: table(30000) }]
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(5)
  expect(pages.every(page => page.lines.length === 1 && page.diagnostics?.includes('repeated-content-overflow'))).toBe(true)
  const rows = table(1000)
  rows.rows[0].isHeader = true
  rows.rows.push(...Array.from({ length: 4 }, () => ({ cells: [{ gridSpan: 1, paragraphs: [p('ROW')] }] })))
  doc.sections[0].blocks = [{ kind: 'table', table: rows }]
  const rowPages = layoutDocx(doc, fixed)
  expect(rowPages.flatMap(page => page.lines.filter(line => line.segs.some(seg => seg.text === 'ROW')))).toHaveLength(4)
  expect(rowPages.slice(1).every(page => page.tables[0].rows.length >= 2)).toBe(true)
  expect(rowPages.length).toBeLessThanOrEqual(5)
 })
 test('first-page ordered variants and all references inherit across section boundaries', async () => {
  const doc = await parse(para('FIRST') + '<w:p><w:pPr><w:sectPr><w:titlePg/><w:headerReference w:type="default" r:id="h"/><w:headerReference w:type="first" r:id="h"/></w:sectPr></w:pPr></w:p>' + para('SECOND'), '', para('INHERITED'))
  expect(doc.sections).toHaveLength(2)
  expect(doc.sections[1].headerBlocks).toEqual(doc.sections[0].headerBlocks)
  expect(doc.sections[1].firstHeaderBlocks).toEqual(doc.sections[0].firstHeaderBlocks)
  doc.sections[0].firstHeaderBlocks = []
  expect(layoutDocx(doc, fixed)[0].header?.paragraphs).toEqual([])
  expect(layoutDocx(doc, fixed)[1].header?.paragraphs[0].runs[0].text).toBe('INHERITED')
 })
 test('paragraph mark style determines empty line height without changing document font defaults', async () => {
  const doc = await parse('<w:p/><w:p/>', '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="normal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>')
  expect(doc.defaultFontSizePt).toBe(11)
  const lines = layoutDocx(doc, fixed)[0].lines
  expect(lines[1].yPx - lines[0].yPx).toBeCloseTo(10 * (96 / 72) * 1.35 * 0.9)
 })
})

describe('repeated block paint compatibility', () => {
 test('table fields and inline images paint over fills and borders; stale legacy content stays excluded', async () => {
  const { createCanvas } = await import('canvas')
  const { renderPages } = await import('../src/docx/layout')
  const doc = model(90)
  const bitmap = createCanvas(8, 8)
  bitmap.getContext('2d').fillStyle = '#00FF00'
  bitmap.getContext('2d').fillRect(0, 0, 8, 8)
  const body = { data: new Uint8Array(), widthEmu: 76200, heightEmu: 76200 }
  const img = { ...body }
  doc.sections[0].paragraphs[0].images = [body]
  const t = table(480)
  t.rows[0].cells[0].borders = { bottom: { style: 'single', color: '0000FF', widthPt: 2 } }
  const field: DocxParagraph = { ...p(), runs: [{ text: '?', field: 'PAGE' }, { text: '/' }, { text: '?', field: 'NUMPAGES' }], images: [img] }
  field.inline = [{ kind: 'image', image: img }, ...field.runs.map(run => ({ kind: 'text' as const, run }))]
  t.rows[0].cells[0].paragraphs = [field]
  doc.sections[0].header = [p('STALE')]
  doc.sections[0].headerBlocks = [{ kind: 'table', table: t }]
  expect(collectDocImages(doc)).toEqual([body, img])
  const pages = layoutDocx(doc, fixed)
  expect(pages.length).toBeGreaterThan(1)
  for (const [index, page] of pages.entries()) {
   const ctx = createCanvas(page.widthPx, page.heightPx).getContext('2d')
   const texts: string[] = []
   const proxy = new Proxy(ctx, { get(target, key) { if (key === 'fillText') return (text: string) => { texts.push(text) }; const value = Reflect.get(target, key); return typeof value === 'function' ? value.bind(target) : value }, set(target, key, value) { Reflect.set(target, key, value); return true } })
   renderPages([page], proxy as never, [bitmap as unknown as CanvasImageSource, bitmap as unknown as CanvasImageSource], { pageNumberStart: index + 1, totalPages: pages.length })
   expect(texts.join('')).toContain(`${index + 1}/${pages.length}`)
   expect(texts.join('')).not.toContain('STALE')
   expect([...ctx.getImageData(98, 62, 1, 1).data].slice(0, 3)).toEqual([0, 255, 0])
   expect([...ctx.getImageData(150, 60, 1, 1).data].slice(0, 3)).toEqual([0, 170, 0])
   expect(ctx.getImageData(150, 79, 1, 1).data[2]).toBe(255)
  }
 })
 test('oversized repeating tables clip to fixed page viewports and leave the minimum body band', async () => {
  const { createCanvas } = await import('canvas')
  const { renderPages } = await import('../src/docx/layout')
  const doc = model()
  doc.sections[0].headerBlocks = [{ kind: 'table', table: table(30000) }]
  const page = layoutDocx(doc, fixed)[0]
  const ctx = createCanvas(page.widthPx, page.heightPx).getContext('2d')
  renderPages([page], ctx as never)
  expect(ctx.getImageData(150, 900, 1, 1).data[1]).toBe(170)
  expect(ctx.getImageData(150, page.heightPx - 2, 1, 1).data[3]).toBe(0)
  expect(page.header?.clipBottomPx).toBeCloseTo(page.heightPx - 19.8)
 })
 test('cell-specific spacing and margins measure image-only paragraphs in repeated tables', () => {
  const doc = model()
  const img = { data: new Uint8Array(), widthEmu: 76200, heightEmu: 152400 }
  const t = table()
  t.rows[0].heightTwips = undefined
  t.rows[0].heightRule = 'auto'
  t.rows[0].cells[0].margins = { leftTwips: 150, rightTwips: 0, topTwips: 75, bottomTwips: 90 }
  t.rows[0].cells[0].paragraphs = [{ ...p(), images: [img], spacingBeforeTwips: 45, spacingAfterTwips: 60 }]
  doc.sections[0].headerBlocks = [{ kind: 'table', table: t }]
  const page = layoutDocx(doc, fixed)[0]
  expect(page.images[0]).toMatchObject({ imageIndex: 0, xPx: 106, yPx: 56, widthPx: 8, heightPx: 16 })
  expect(page.lines[0].yPx).toBe(96) // occupied bottom82 stays above ordinary top margin96
 })
})

describe('spec review repeated-content corrections', () => {
 test('empty authoritative header/footer blocks reserve no space at intruding distances', () => {
  const doc = model(90)
  const baseline = layoutDocx(doc, fixed)
  doc.sections[0].margins.headerTwips = 3000
  doc.sections[0].margins.footerTwips = 3000
  doc.sections[0].header = [p('STALE HEADER')]
  doc.sections[0].footer = [p('STALE FOOTER')]
  doc.sections[0].headerBlocks = []
  doc.sections[0].footerBlocks = []
  const pages = layoutDocx(doc, fixed)
  expect(pages.map(page => page.lines.map(line => line.yPx))).toEqual(baseline.map(page => page.lines.map(line => line.yPx)))
  expect(pages[0].lines[0].yPx).toBe(96)
  expect(pages.every(page => !page.diagnostics?.length)).toBe(true)
 })
 test('floating images in header table cells retain cell anchors, reserve occupied bounds and paint', async () => {
  const { createCanvas } = await import('canvas')
  const { renderPages } = await import('../src/docx/layout')
  const doc = model()
  const img = { data: new Uint8Array(), widthEmu: 952500, heightEmu: 952500, floating: { behindDoc: false, relativeHeight: 1, wrap: 'none' as const, posH: { relativeFrom: 'column', offsetEmu: 95250 }, posV: { relativeFrom: 'paragraph', offsetEmu: 190500 } } }
  const t = table(480)
  t.gridColsTwips = [1500, 3000]
  t.rows[0].cells.unshift({ gridSpan: 1, paragraphs: [p()] })
  t.rows[0].cells[1].margins = { leftTwips: 150, rightTwips: 0, topTwips: 75, bottomTwips: 0 }
  t.rows[0].cells[1].paragraphs = [{ ...p(), images: [img], spacingBeforeTwips: 45 }]
  doc.sections[0].headerBlocks = [{ kind: 'table', table: t }]
  expect(collectDocImages(doc)).toEqual([img])
  const page = layoutDocx(doc, fixed)[0]
  expect(page.images).toHaveLength(1)
  expect(page.images[0]).toMatchObject({ xPx: 216, yPx: 76, widthPx: 100, heightPx: 100, imageIndex: 0, repeated: 'header' })
  expect(page.lines[0].yPx).toBe(176)
  const bitmap = createCanvas(100, 100)
  bitmap.getContext('2d').fillStyle = '#00FFFF'
  bitmap.getContext('2d').fillRect(0, 0, 100, 100)
  const ctx = createCanvas(page.widthPx, page.heightPx).getContext('2d')
  renderPages([page], ctx as never, [bitmap as unknown as CanvasImageSource])
  expect([...ctx.getImageData(250, 150, 1, 1).data].slice(0, 3)).toEqual([0, 255, 255])
 })
})

test('footer cell anchors keep page-relative coordinates and grow occupied bounds upward', () => {
 const doc = model(90)
 const paragraphImage = { data: new Uint8Array(), widthEmu: 952500, heightEmu: 952500, floating: { behindDoc: false, relativeHeight: 1, wrap: 'none' as const, posH: { relativeFrom: 'column', offsetEmu: 0 }, posV: { relativeFrom: 'paragraph', offsetEmu: -952500 } } }
 const pageImage = { ...paragraphImage, floating: { ...paragraphImage.floating, posH: { relativeFrom: 'page', offsetEmu: 476250 }, posV: { relativeFrom: 'page', offsetEmu: 1905000 } } }
 const t = table(480)
 t.rows[0].cells[0].paragraphs = [{ ...p(), images: [paragraphImage, pageImage] }]
 doc.sections[0].footerBlocks = [{ kind: 'table', table: t }]
 const pages = layoutDocx(doc, fixed)
 expect(pages.length).toBeGreaterThan(1)
 for (const page of pages) {
  expect(page.images).toHaveLength(2)
  expect(page.images[0]).toMatchObject({ xPx: 96, yPx: 876, repeated: 'footer' })
  expect(page.images[1]).toMatchObject({ xPx: 50, yPx: 200, repeated: 'footer' })
  expect(page.lines.every(line => line.yPx + line.heightPx <= 200)).toBe(true)
 }
})

describe('quality review section flow and physical margins', () => {
 test('current section ending properties govern its start: continuous follows a nextPage section on the same page', async () => {
  const doc = await parse(para('A', '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>') + para('B'), '', undefined, '<w:type w:val="continuous"/>')
  expect(doc.sections.map(section => section.type)).toEqual(['nextPage', 'continuous'])
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(1)
  expect(pages[0].lines.map(line => line.segs.map(seg => seg.text).join(''))).toEqual(['A', 'B'])
  expect(pages[0].lines.map(line => line.yPx)).toEqual([96, 115.8])
 })
 test('default nextPage on the current section starts a new page even after a continuous preceding section', async () => {
  const doc = await parse(para('A', '<w:sectPr><w:type w:val="continuous"/></w:sectPr>') + para('B'))
  expect(doc.sections.map(section => section.type)).toEqual(['continuous', 'nextPage'])
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(2)
  expect(pages.map(page => page.lines[0].yPx)).toEqual([96, 96])
 })
 test.each(['header', 'footer'] as const)('%s margin-relative anchors use physical section top margin', kind => {
  const doc = model()
  const img = { data: new Uint8Array(), widthEmu: 76200, heightEmu: 76200, floating: { behindDoc: false, relativeHeight: 1, wrap: 'none' as const, posH: { relativeFrom: 'margin', offsetEmu: 0 }, posV: { relativeFrom: 'margin', offsetEmu: 0 } } }
  doc.sections[0][`${kind}Blocks`] = [{ kind: 'p', paragraph: { ...p(), images: [img] } }]
  const page = layoutDocx(doc, fixed)[0]
  expect(page.images[0]).toMatchObject({ xPx: 96, yPx: 96, repeated: kind })
 })
})

test('continuous section overflow advances onto a fresh page with its own repeated bounds', () => {
 const doc = model(90)
 const first = doc.sections[0]
 first.paragraphs = first.paragraphs.slice(0, 20)
 first.blocks = first.paragraphs.map(paragraph => ({ kind: 'p' as const, paragraph }))
 first.headerBlocks = [{ kind: 'table', table: table(1440) }]
 const paras = Array.from({ length: 70 }, (_, i) => p(`CONT${i}`))
 doc.sections.push({ ...first, type: 'continuous', paragraphs: paras, blocks: paras.map(paragraph => ({ kind: 'p', paragraph })), headerBlocks: [{ kind: 'table', table: table(480) }] })
 const pages = layoutDocx(doc, fixed)
 expect(pages).toHaveLength(3)
 expect(pages[0].header?.blocks).toBe(first.headerBlocks)
 expect(pages[1].header?.blocks).toBe(doc.sections[1].headerBlocks)
 expect(pages[1].lines[0].yPx).toBe(96)
 expect(pages.flatMap(page => page.lines)).toHaveLength(90)
 expect(pages[0].lines.filter(line => line.segs.some(seg => seg.text.startsWith('CONT'))).length).toBeGreaterThan(0)
})

describe('physical margin alignment ranges for repeated table anchors', () => {
 test.each(['header', 'footer'] as const)('%s cell margin center/right alignment spans physical page margins', kind => {
  for (const [align, xPx] of [['center', 404], ['right', 712]] as const) {
   const doc = model()
   const image = { data: new Uint8Array(), widthEmu: 76200, heightEmu: 76200, floating: { behindDoc: false, relativeHeight: 1, wrap: 'none' as const, posH: { relativeFrom: 'margin', offsetEmu: 0, align }, posV: { relativeFrom: 'paragraph', offsetEmu: 0 } } }
   const t = table(480)
   t.rows[0].cells[0].paragraphs = [{ ...p(), images: [image] }]
   doc.sections[0][`${kind}Blocks`] = [{ kind: 'table', table: t }]
   const page = layoutDocx(doc, fixed)[0]
   expect(page.images[0]).toMatchObject({ xPx, yPx: kind === 'header' ? 48 : 976 })
  }
 })
 test.each(['header', 'footer'] as const)('%s cell margin vertical center/bottom alignment spans physical top/bottom margins', kind => {
  for (const [align, yPx] of [['center', 544], ['bottom', 988]] as const) {
   const doc = model()
   doc.sections[0].margins.topTwips = 1500
   doc.sections[0].margins.bottomTwips = 900
   const image = { data: new Uint8Array(), widthEmu: 76200, heightEmu: 76200, floating: { behindDoc: false, relativeHeight: 1, wrap: 'none' as const, posH: { relativeFrom: 'column', offsetEmu: 0 }, posV: { relativeFrom: 'margin', offsetEmu: 0, align } } }
   const t = table(480)
   t.rows[0].cells[0].paragraphs = [{ ...p(), images: [image] }]
   doc.sections[0][`${kind}Blocks`] = [{ kind: 'table', table: t }]
   const page = layoutDocx(doc, fixed)[0]
   expect(page.images[0]).toMatchObject({ xPx: 96, yPx })
  }
 })
})

describe('explicit empty final sections', () => {
 test('a finalized empty nextPage section retains its own header/footer and page', async () => {
  const doc = await parse(para('A', '<w:sectPr/>'), '', para('FINAL HEADER'), '', para('FINAL FOOTER'))
  expect(doc.sections).toHaveLength(2)
  expect(doc.sections[1]).toMatchObject({ type: 'nextPage', blocks: [], paragraphs: [], header: [{ runs: [{ text: 'FINAL HEADER' }] }], footer: [{ runs: [{ text: 'FINAL FOOTER' }] }] })
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(2)
  expect(pages[1].lines).toEqual([])
  expect(pages[1].header?.paragraphs[0].runs[0].text).toBe('FINAL HEADER')
  expect(pages[1].footer?.paragraphs[0].runs[0].text).toBe('FINAL FOOTER')
  const { createCanvas } = await import('canvas')
  const { renderPages } = await import('../src/docx/layout')
  const ctx = createCanvas(pages[1].widthPx, pages[1].heightPx).getContext('2d')
  const text: string[] = []
  const proxy = new Proxy(ctx, { get(target, key) { if (key === 'fillText') return (value: string) => text.push(value); const value = Reflect.get(target, key); return typeof value === 'function' ? value.bind(target) : value }, set(target, key, value) { Reflect.set(target, key, value); return true } })
  renderPages([pages[1]], proxy as never, undefined, { pageNumberStart: 2, totalPages: 2 })
  expect(text.join('')).toContain('FINAL HEADER')
  expect(text.join('')).toContain('FINAL FOOTER')
 })
 test('a handwritten explicitly empty section is retained after a content section', () => {
  const doc = model()
  doc.sections.push({ ...doc.sections[0], paragraphs: [], blocks: [], header: [p('EMPTY SECTION')] })
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(2)
  expect(pages[1].header?.paragraphs[0].runs[0].text).toBe('EMPTY SECTION')
 })
 test('an explicitly empty continuous section is retained without creating another page', async () => {
  const doc = await parse(para('A', '<w:sectPr/>'), '', para('FINAL HEADER'), '<w:type w:val="continuous"/>', para('FINAL FOOTER'))
  expect(doc.sections).toHaveLength(2)
  expect(doc.sections[1].blocks).toEqual([])
  expect(doc.sections[1].type).toBe('continuous')
  expect(layoutDocx(doc, fixed)).toHaveLength(1)
 })
 test('an ordinary finalized body does not retain the unused trailing parser section', async () => {
  const doc = await parse(para('A'), '', para('HEADER'), '', para('FOOTER'))
  expect(doc.sections).toHaveLength(1)
  expect(layoutDocx(doc, fixed)).toHaveLength(1)
 })
 test('a pure empty finalized document retains one page with its header/footer', async () => {
  const doc = await parse('', '', para('HEADER'), '', para('FOOTER'))
  expect(doc.sections).toHaveLength(1)
  const pages = layoutDocx(doc, fixed)
  expect(pages).toHaveLength(1)
  expect(pages[0].header?.paragraphs[0].runs[0].text).toBe('HEADER')
  expect(pages[0].footer?.paragraphs[0].runs[0].text).toBe('FOOTER')
 })
})
