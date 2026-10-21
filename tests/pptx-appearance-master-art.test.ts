/* Master and layout artwork (slideshow S0-6). Office-shaped synthetic decks from buildPptxDeck. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { OfficePackage } from '../src/core/zip'
import { slideTheme } from '../src/pptx/parse'
import { masterArtwork } from '../src/pptx/appearance/master-art'
import { buildPptxDeck, type PptxDeckSpec } from '../src/testdata/ooxml-builders'
import type { PptxShape } from '../src/pptx/types'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const pngBytes = () => Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0))

const SLIDE = 'ppt/slides/slide1.xml'
const MASTER = 'ppt/slideMasters/slideMaster1.xml'
const LAYOUT = 'ppt/slideLayouts/slideLayout1.xml'

function deck(options: { slideShowMasterSp?: boolean; layoutShowMasterSp?: boolean } = {}): PptxDeckSpec {
  return {
    masters: [{
      shapes: [
        { id: 10, name: 'Master Logo', image: { data: pngBytes(), name: 'logo.png' }, off: ['100000', '100000'], ext: ['500000', '500000'] },
        { id: 11, name: 'Master Title', placeholder: { type: 'title' }, off: ['0', '0'], ext: ['9144000', '1000000'] },
        { id: 12, name: 'Master Body', placeholder: { type: 'body', idx: 1 }, off: ['0', '1000000'], ext: ['9144000', '5000000'] },
      ],
      layouts: [{
        showMasterSp: options.layoutShowMasterSp,
        placeholders: [{ type: 'title', off: ['0', '0'], ext: ['9144000', '1000000'] }],
        shapes: [{ id: 20, name: 'Layout Band', fill: '336699', off: ['0', '6000000'], ext: ['9144000', '858000'] }],
      }],
    }],
    slides: [{
      showMasterSp: options.slideShowMasterSp,
      shapes: [
        { id: 2, name: 'Slide Title', placeholder: { type: 'title' }, paragraphs: [{ runs: [{ text: 'Hello' }] }] },
        { id: 3, name: 'Slide Box', fill: 'FF0000', off: ['1000000', '2000000'], ext: ['1000000', '1000000'] },
      ],
    }],
  }
}

async function open(bytes: Uint8Array) {
  const pkg = await OfficePackage.load(bytes)
  const theme = (await slideTheme(pkg, SLIDE)).context
  return masterArtwork(pkg, SLIDE, theme)
}

/** Insert raw XML at the end of a part's spTree. */
async function patchSpTree(bytes: Uint8Array, part: string, xml: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const text = await zip.file(part)!.async('text')
  zip.file(part, text.replace('</p:spTree>', `${xml}</p:spTree>`))
  return zip.generateAsync({ type: 'uint8array' })
}

function names(shapes: PptxShape[]): string[] {
  const out: string[] = []
  const walk = (list: PptxShape[]) => { for (const s of list) { out.push(s.source?.name ?? '?'); if (s.children) walk(s.children) } }
  walk(shapes)
  return out
}

describe('masterArtwork', () => {
  it('master logo and layout band are returned; title and body placeholders are not', async () => {
    const art = await open(await buildPptxDeck(deck()))
    expect(names(art.master)).toEqual(['Master Logo'])
    expect(names(art.layout)).toEqual(['Layout Band'])
    const logo = art.master[0]
    // the picture goes through the slide image loader, resolved against the master part
    expect(logo.image?.partPath).toBe('ppt/media/image1.png')
    expect(logo.image?.data.byteLength).toBe(pngBytes().byteLength)
    expect(art.layout[0].drawingStyle?.fill).toMatchObject({ kind: 'solid' })
    expect(art.layout[0].yEmu).toBe(6000000)
  })

  it('slide showMasterSp=0 drops both; layout showMasterSp=0 drops master only', async () => {
    const slideOff = await open(await buildPptxDeck(deck({ slideShowMasterSp: false })))
    expect(slideOff.master).toEqual([])
    expect(slideOff.layout).toEqual([])
    const layoutOff = await open(await buildPptxDeck(deck({ layoutShowMasterSp: false })))
    expect(layoutOff.master).toEqual([])
    expect(names(layoutOff.layout)).toEqual(['Layout Band'])
    // "false" is an xsd:boolean spelling too
    const bytes = await buildPptxDeck(deck())
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file(LAYOUT)!.async('text')
    zip.file(LAYOUT, xml.replace('<p:sldLayout ', '<p:sldLayout showMasterSp="false" '))
    const spelled = await open(await zip.generateAsync({ type: 'uint8array' }))
    expect(spelled.master).toEqual([])
    expect(names(spelled.layout)).toEqual(['Layout Band'])
  })

  it('a placeholder nested in a group is excluded', async () => {
    const group = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="Art Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="1000000"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="31" name="Nested Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="32" name="Nested Art"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="500000" y="500000"/><a:ext cx="500000" cy="500000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr></p:sp>' +
      // a picture placeholder (nvPicPr/nvPr/ph) in the group: pics are not marked by the parser, so the filter must catch it
      '<p:pic><p:nvPicPr><p:cNvPr id="33" name="Nested Picture Placeholder"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="3"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
      '</p:grpSp>' +
      // a top-level picture placeholder and a graphicFrame placeholder
      '<p:pic><p:nvPicPr><p:cNvPr id="34" name="Picture Placeholder"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="4"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="35" name="Table Placeholder"/><p:cNvGraphicFramePr/><p:nvPr><p:ph idx="5"/></p:nvPr></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="1000000"/></a:tblGrid>' +
      '<a:tr h="500000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
    const art = await open(await patchSpTree(await buildPptxDeck(deck()), MASTER, group))
    expect(names(art.master)).toEqual(['Master Logo', 'Art Group', 'Nested Art'])
    expect(art.master[1].children?.map(s => s.source?.id)).toEqual(['32'])
  })

  it('artwork source.partPath names the layout or master part', async () => {
    const art = await open(await buildPptxDeck(deck()))
    expect(art.master.map(s => s.source?.partPath)).toEqual([MASTER])
    expect(art.layout.map(s => s.source?.partPath)).toEqual([LAYOUT])
    expect(art.diagnostics).toEqual([])
  })

  it('a shape filled with the slide background (useBgFill) is diagnosed, not silently painted with its own fill', async () => {
    const sp = '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="40" name="Bg Filled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
    const art = await open(await patchSpTree(await buildPptxDeck(deck()), MASTER, sp))
    expect(names(art.master)).toEqual(['Master Logo', 'Bg Filled'])
    expect(art.diagnostics.filter(d => 'feature' in d && d.feature === 'useBgFill').map(d => d.message)).toEqual([`1 shape(s) in ${MASTER} fill with the slide background; their own fill paints instead`])
  })
  
  it('a shared cache parses each part once per colour map', async () => {
    const pkg = await OfficePackage.load(await buildPptxDeck(deck()))
    const theme = (await slideTheme(pkg, SLIDE)).context
    const cache = new Map()
    const first = await masterArtwork(pkg, SLIDE, theme, cache)
    const second = await masterArtwork(pkg, SLIDE, theme, cache)
    expect(cache.size).toBe(2)
    expect(second.master[0]).toBe(first.master[0])
    expect(second.layout[0]).toBe(first.layout[0])
    const remapped = await masterArtwork(pkg, SLIDE, { ...theme, colorMap: { ...theme.colorMap, bg1: 'dk1' } }, cache)
    expect(remapped.master[0]).not.toBe(first.master[0])
    expect(cache.size).toBe(4)
  })
})
