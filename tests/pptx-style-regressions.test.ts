import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { renderSlide } from '../src/pptx/render'
import { buildPptx } from '../src/testdata/ooxml-builders'

async function backgroundSlide(flag: string, background = true) {
  const zip = new JSZip()
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="2857500" cy="1905000"/><p:sldIdLst><p:sldId r:id="s1"/></p:sldIdLst></p:presentation>')
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
  const transform = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm>'
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld>${background ? '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></p:bgPr></p:bg>' : ''}<p:spTree><p:grpSp><p:grpSpPr>${transform}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:grpSpPr><p:sp useBgFill="${flag}"><p:spPr>${transform}<a:prstGeom prst="rect"/></p:spPr></p:sp></p:grpSp></p:spTree></p:cSld></p:sld>`)
  return (await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))).slides[0]
}

describe('PPTX background and regional diagonal styles', () => {
  test.each(['1', 'true'])('shape useBgFill=%s uses slide background inside a colored group', async flag => {
    const slide = await backgroundSlide(flag)
    const ctx = createCanvas(300, 200).getContext('2d')
    renderSlide(slide, ctx as unknown as CanvasRenderingContext2D)
    expect(Array.from(ctx.getImageData(50, 50, 1, 1).data)).toEqual([0, 0, 255, 255])
    expect(slide.shapes[0].children![0].fill).toBe('#0000FF')
  })

  test('useBgFill uses white when the slide has no background; false keeps group fill', async () => {
    for (const [flag, expected] of [['1', [255, 255, 255, 255]], ['false', [255, 0, 0, 255]]] as const) {
      const slide = await backgroundSlide(flag, false)
      const ctx = createCanvas(300, 200).getContext('2d')
      renderSlide(slide, ctx as unknown as CanvasRenderingContext2D)
      expect(Array.from(ctx.getImageData(50, 50, 1, 1).data)).toEqual(expected)
    }
  })

  test.each(['tl2br', 'tr2bl'])('first-row %s diagonal overrides whole-table style only in header', async diagonal => {
    const border = (color: string) => `<a:tcStyle><a:tcBdr><a:${diagonal}><a:ln w="38100"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></a:${diagonal}></a:tcBdr></a:tcStyle>`
    const bytes = await buildPptx([{ off: ['0', '0'], ext: ['1905000', '1905000'], table: {
      firstRow: true, styleId: 'RegionalDiagonal', colWidths: ['1905000'],
      customStyleXml: `<a:tblStyle styleId="RegionalDiagonal"><a:wholeTbl>${border('0000FF')}</a:wholeTbl><a:firstRow>${border('FF0000')}</a:firstRow></a:tblStyle>`,
      rows: [{ h: '952500', cells: [{ paragraphs: [] }] }, { h: '952500', cells: [{ paragraphs: [] }] }],
    } }])
    const slide = (await parsePptx(await OfficePackage.load(bytes))).slides[0]
    const ctx = createCanvas(300, 300).getContext('2d')
    renderSlide(slide, ctx as unknown as CanvasRenderingContext2D)
    expect(Array.from(ctx.getImageData(100, 50, 1, 1).data)).toEqual([255, 0, 0, 255])
    expect(Array.from(ctx.getImageData(100, 150, 1, 1).data)).toEqual([0, 0, 255, 255])
  })

  test.each([false, true])('direct diagonal overrides first-row style, header enabled=%s', async firstRow => {
    const bytes = await buildPptx([{ off: ['0', '0'], ext: ['1905000', '952500'], table: {
      firstRow, styleId: 'DirectDiagonal', colWidths: ['952500', '952500'],
      customStyleXml: '<a:tblStyle styleId="DirectDiagonal"><a:firstRow><a:tcStyle><a:tcBdr><a:tl2br><a:ln w="38100"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:tl2br></a:tcBdr></a:tcStyle></a:firstRow></a:tblStyle>',
      rows: [{ h: '952500', cells: [
        { paragraphs: [], borders: '<a:lnTlToBr w="38100"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnTlToBr>' },
        { paragraphs: [], borders: '<a:lnTlToBr><a:noFill/></a:lnTlToBr>' },
      ] }],
    } }])
    const slide = (await parsePptx(await OfficePackage.load(bytes))).slides[0]
    const ctx = createCanvas(300, 200).getContext('2d')
    renderSlide(slide, ctx as unknown as CanvasRenderingContext2D)
    expect(Array.from(ctx.getImageData(50, 50, 1, 1).data)).toEqual([0, 255, 0, 255])
    expect(Array.from(ctx.getImageData(150, 50, 1, 1).data)).toEqual([255, 255, 255, 255])
  })
})
