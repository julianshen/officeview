/** Office-shaped synthetic decks; these are format assertions, not Office visual evidence. */
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXmlOrdered } from '../src/core/xml'
import { parseThemeContext } from '../src/drawing/style'
import { buildPptxDeck, type PptxDeckSpec } from '../src/testdata/ooxml-builders'
import { backgroundNode, resolveSlideBackground } from '../src/pptx/appearance/background'

const slide = 'ppt/slides/slide1.xml'
const theme = () => parseThemeContext(undefined)
async function pkg(spec: PptxDeckSpec, edit?: (zip: JSZip) => Promise<void>) {
  const zip = await JSZip.loadAsync(await buildPptxDeck(spec))
  if (edit) await edit(zip)
  return OfficePackage.load(await zip.generateAsync({ type: 'uint8array' }))
}
const solid = (r: number, g: number, b: number) => ({ kind: 'fill', fill: { kind: 'solid', color: { r, g, b, a: 1 } } })
it('linear and radial gradients become a gradient DrawingFill with stops and angle', async () => {
  const spec: PptxDeckSpec = { slides: [{ shapes: [], bg: { gradient: { angle: 90, stops: [[0, 'FF0000'], [100000, '0000FF']] } } }] }
  for (const radial of [false, true]) {
    const source = await pkg(spec, async zip => {
      if (radial) zip.file(slide, (await zip.file(slide)!.async('string')).replace(/<a:lin[^>]*\/>/, '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'))
    })
    const { background, diagnostics } = await resolveSlideBackground(source, slide, theme())
    expect(diagnostics).toEqual([])
    expect(background).toMatchObject({ kind: 'fill', fill: { kind: 'gradient', gradient: radial ? 'circle' : 'linear', angle: radial ? 0 : Math.PI / 2, stops: [{ position: 0, color: { r: 255, g: 0, b: 0, a: 1 } }, { position: 1, color: { r: 0, g: 0, b: 255, a: 1 } }] } })
    expect(backgroundNode(background, 9144000, 6858000)).toMatchObject({ xEmu: 0, yEmu: 0, widthEmu: 9144000, heightEmu: 6858000, geometry: 'rect', drawingStyle: { fill: background.kind === 'fill' ? background.fill : undefined } })
  }
})
it('a stretched picture background becomes a full-slide picture node; tile records its scale and offset', async () => {
  const data = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
  for (const tiled of [false, true]) {
    const source = await pkg({ slides: [{ shapes: [], bg: { picture: { data, tile: tiled } } }] }, async zip => {
      zip.file(slide, (await zip.file(slide)!.async('string')).replace('<a:tile/>', '<a:tile sx="50000" sy="200000" tx="9144" ty="-9144" algn="ctr" flip="xy"/>'))
    })
    const { background } = await resolveSlideBackground(source, slide, theme())
    expect(background.kind).toBe('picture')
    if (background.kind !== 'picture') throw new Error('picture expected')
    expect(background.image.data).toEqual(data)
    if (tiled) expect(background.tile).toEqual({ sx: .5, sy: 2, tx: 9144, ty: -9144, algn: 'ctr', flip: 'xy' })
    else {
      expect(background.stretch).toEqual({ l: 0, t: 0, r: 0, b: 0 })
      expect(backgroundNode(background, 100, 200)).toMatchObject({ widthEmu: 100, heightEmu: 200, image: { data, fillRect: { l: 0, t: 0, r: 0, b: 0 } } })
    }
  }
})
it('pattern background keeps preset, fg and bg', async () => {
  const source = await pkg({ slides: [{ shapes: [], bg: { pattern: { prst: 'pct50', fg: 'FF0000', bg: '0000FF' } } }] })
  expect((await resolveSlideBackground(source, slide, theme())).background).toEqual({ kind: 'fill', fill: { kind: 'pattern', preset: 'pct50', fgColor: { r: 255, g: 0, b: 0, a: 1 }, bgColor: { r: 0, g: 0, b: 255, a: 1 } } })
})
it('bgRef 1001+ uses bgFillStyleLst with phClr replaced; 1-999 uses fillStyleLst', async () => {
  const t = parseThemeContext(parseXmlOrdered('<a:theme><a:themeElements><a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:bgFillStyleLst><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'))
  for (const idx of [1, 1002]) {
    const source = await pkg({ slides: [{ shapes: [], bg: { bgRef: { idx, color: '123456' } } }] })
    expect((await resolveSlideBackground(source, slide, t)).background).toEqual(solid(18, 52, 86))
  }
  for (const idx of [0, 1000]) {
    const source = await pkg({ slides: [{ shapes: [], bg: { bgRef: { idx, color: '123456' } } }] })
    expect((await resolveSlideBackground(source, slide, t)).background).toEqual({ kind: 'none' })
  }
})
it('slide bg beats layout beats master; absent everywhere is none', async () => {
  for (const level of [0, 1, 2, 3]) {
    const source = await pkg({ masters: [{ bg: level < 3 ? { solid: '0000FF' } : undefined, layouts: [{ bg: level < 2 ? { solid: '00FF00' } : undefined }] }], slides: [{ layout: [0, 0], shapes: [], bg: level === 0 ? { solid: 'FF0000' } : undefined }] })
    expect((await resolveSlideBackground(source, slide, theme())).background).toEqual([solid(255, 0, 0), solid(0, 255, 0), solid(0, 0, 255), { kind: 'none' }][level])
  }
  expect(backgroundNode({ kind: 'none' }, 100, 100)).toBeUndefined()
})
it('a missing picture part yields none plus a missing-image diagnostic', async () => {
  const source = await pkg({ slides: [{ shapes: [], bg: { picture: { data: new Uint8Array([1, 2, 3]) } } }] }, async zip => { for (const name of Object.keys(zip.files)) if (name.startsWith('ppt/media/')) zip.remove(name) })
  const result = await resolveSlideBackground(source, slide, theme())
  expect(result.background).toEqual({ kind: 'none' })
  expect(result.diagnostics).toEqual([expect.objectContaining({ kind: 'missing-image' })])
})
it('uncached reads refuse oversized XML and do not call materialising package APIs', async () => {
  const source = await pkg({ slides: [{ shapes: [], bg: { solid: '123456' } }] })
  source.text = source.bytes = source.xml = source.xmlOrdered = (() => { throw new Error('cached API used') }) as never
  expect((await resolveSlideBackground(source, slide, theme())).background).toEqual(solid(18, 52, 86))
  const oversized = await pkg({ slides: [{ shapes: [] }] }, async zip => { zip.file(slide, '<p:sld>' + ' '.repeat(4 * 1024 * 1024) + '</p:sld>') })
  expect(await resolveSlideBackground(oversized, slide, theme())).toMatchObject({ background: { kind: 'none' }, diagnostics: [expect.objectContaining({ feature: 'resource-limit' })] })
})
it('inherited pictures resolve relative to their owner and explicit noFill stops inheritance', async () => {
  const data = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
  const spec: PptxDeckSpec = { masters: [{ bg: { picture: { data } }, layouts: [{}] }], slides: [{ shapes: [], layout: [0, 0] }] }
  const source = await pkg(spec)
  const result = await resolveSlideBackground(source, slide, theme())
  expect(result.background).toMatchObject({ kind: 'picture', image: { data, partPath: 'ppt/media/image1.png' } })
  const noFill = await pkg(spec, async zip => { zip.file(slide, (await zip.file(slide)!.async('string')).replace('<p:cSld>', '<p:cSld><p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>')) })
  expect((await resolveSlideBackground(noFill, slide, theme())).background).toEqual({ kind: 'fill', fill: { kind: 'none' } })
})
it('theme blipFill resolves against the theme part and diagnoses duotone', async () => {
  const source = await pkg({ masters: [{ layouts: [{}], bg: { bgRef: { idx: 1001, color: 'FF0000' } } }], slides: [{ shapes: [], layout: [0, 0] }], theme: { bgFillStyles: ['<a:blipFill><a:blip r:embed="themeImage"><a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone></a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>'] } }, async zip => {
    zip.file('ppt/theme/_rels/theme1.xml.rels', '<Relationships><Relationship Id="themeImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/theme-only.png"/></Relationships>')
    zip.file('ppt/media/theme-only.png', new Uint8Array([137, 80, 78, 71]))
  })
  const t = parseThemeContext(await source.xmlOrdered('ppt/theme/theme1.xml'))
  const result = await resolveSlideBackground(source, slide, t)
  expect(result.background).toMatchObject({ kind: 'picture', image: { partPath: 'ppt/media/theme-only.png' } })
  expect(result.diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported-fill', feature: 'duotone' }))
})
it('lo-tdf123684 resolves its theme picture background', async ctx => {
  const { existsSync, readFileSync } = await import('node:fs')
  const path = 'corpus/lo-tdf123684.pptx'
  if (!existsSync(path)) { ctx.skip('run bun scripts/fetch-corpus.ts'); return }
  const source = await OfficePackage.load(readFileSync(path))
  const t = parseThemeContext(await source.xmlOrdered('ppt/theme/theme1.xml'))
  const result = await resolveSlideBackground(source, slide, t)
  expect(result.background).toMatchObject({ kind: 'picture', image: { data: expect.any(Uint8Array) } })
  if (result.background.kind === 'picture') expect(result.background.image.data.length).toBeGreaterThan(0)
})
it('picture reads share the HandlePackage aggregate and refuse with a diagnostic', async () => {
  const { HandlePackage, HANDLE_AGGREGATE_BYTES } = await import('../src/pptx/appearance/handle-package')
  const source = await pkg({ slides: [{ shapes: [], bg: { picture: { data: new Uint8Array(128 * 1024) } } }] }, async zip => {
    zip.file('pressure.bin', new Uint8Array(HANDLE_AGGREGATE_BYTES - 65536))
  })
  const over = HandlePackage.over({ pkg: source, slideParts: [slide] }, { partBytes: HANDLE_AGGREGATE_BYTES })
  await over.bytes('pressure.bin')
  const result = await resolveSlideBackground(over, slide, theme())
  expect(result.background).toEqual({ kind: 'none' })
  expect(result.diagnostics).toContainEqual(expect.objectContaining({ kind: 'missing-image', feature: 'resource-limit' }))
})
it('nonrotating background gradients do not report a rotation compensation warning', async () => {
  const source = await pkg({ slides: [{ shapes: [], bg: { gradient: { angle: 45, stops: [[0, 'FF0000'], [100000, '0000FF']] } } }] }, async zip => {
    zip.file(slide, (await zip.file(slide)!.async('string')).replace('<a:gradFill>', '<a:gradFill rotWithShape="0">'))
  })
  expect((await resolveSlideBackground(source, slide, theme())).diagnostics).toEqual([])
})
it('diagnoses unavailable bgRef, unsupported path gradients and external pictures', async () => {
  const missing = await pkg({ slides: [{ shapes: [], bg: { bgRef: { idx: 1007, color: '123456' } } }] })
  expect((await resolveSlideBackground(missing, slide, theme())).diagnostics).toContainEqual(expect.objectContaining({ kind: 'missing-theme-style' }))
  const path = await pkg({ slides: [{ shapes: [], bg: { gradient: { angle: 0, stops: [[0, 'FF0000'], [100000, '0000FF']] } } }] }, async zip => {
    zip.file(slide, (await zip.file(slide)!.async('string')).replace(/<a:lin[^>]*\/>/, '<a:path path="rect"/>'))
  })
  expect((await resolveSlideBackground(path, slide, theme())).diagnostics).toContainEqual(expect.objectContaining({ kind: 'unsupported-gradient' }))
  const external = await pkg({ slides: [{ shapes: [], bg: { picture: { data: new Uint8Array([1]) } } }] }, async zip => {
    const rels = 'ppt/slides/_rels/slide1.xml.rels'
    zip.file(rels, (await zip.file(rels)!.async('string')).replace('Target="../media/', 'TargetMode="External" Target="https://example.invalid/'))
  })
  expect(await resolveSlideBackground(external, slide, theme())).toMatchObject({ background: { kind: 'none' }, diagnostics: [expect.objectContaining({ kind: 'missing-image' })] })
})
it('uses the caller effective colour map for an inherited background', async () => {
  const source = await pkg({ masters: [{ layouts: [{}], bg: { bgRef: { idx: 1001, color: 'bg1' } } }], slides: [{ shapes: [], layout: [0, 0] }], theme: { bgFillStyles: ['<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'] } })
  const t = parseThemeContext(await source.xmlOrdered('ppt/theme/theme1.xml'), { bg1: 'accent1' })
  t.colors.accent1 = { kind: 'srgb', value: '123456', transforms: [] }
  expect((await resolveSlideBackground(source, slide, t)).background).toEqual(solid(18, 52, 86))
})
it('refuses a picture beyond the 32 MiB enrichment ceiling without losing diagnostics', async () => {
  const source = await pkg({ slides: [{ shapes: [], bg: { picture: { data: new Uint8Array(32 * 1024 * 1024 + 1) } } }] })
  expect(await resolveSlideBackground(source, slide, theme())).toMatchObject({ background: { kind: 'none' }, diagnostics: [expect.objectContaining({ kind: 'missing-image', feature: 'resource-limit' })] })
})
