/* Builder tests for the multi-slide deck builder (slideshow S0-3). Synthetic fixtures only. */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { buildPptx, buildPptxDeck, type PptxDeckSpec, type PptxShapeSpec } from '../src/testdata/ooxml-builders'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const pngBytes = () => Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0))

async function textEntries(buf: Uint8Array): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(buf)
  const out: Record<string, string> = {}
  for (const path of Object.keys(zip.files)) {
    if (!zip.files[path].dir) out[path] = await zip.files[path].async('text')
  }
  return out
}

async function shaByName(buf: Uint8Array): Promise<{ names: string[]; hashes: Record<string, string> }> {
  const zip = await JSZip.loadAsync(buf)
  const hashes: Record<string, string> = {}
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir) continue
    hashes[path] = createHash('sha256').update(await zip.files[path].async('uint8array')).digest('hex')
  }
  return { names: Object.keys(hashes).sort(), hashes }
}

interface RelRecord { id: string; type: string; target: string; external: boolean }

async function relsOf(entries: Record<string, string>, part: string): Promise<RelRecord[]> {
  const segment = part.indexOf('/') === -1 ? '_rels' : part.slice(0, part.lastIndexOf('/')) + '/_rels'
  const base = part.slice(part.lastIndexOf('/') + 1)
  const relPath = `${segment}/${base}.rels`
  const xml = entries[relPath]
  if (xml === undefined) return []
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => {
    const tag = m[0]
    const attr = (name: string) => tag.match(new RegExp(`${name}="([^"]*)"`))?.[1] ?? ''
    return {
      id: attr('Id'),
      type: attr('Type'),
      target: attr('Target'),
      external: /TargetMode="External"/.test(tag),
    }
  })
}

function resolveTarget(part: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const dir = part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : ''
  const out: string[] = dir ? dir.slice(0, -1).split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') out.pop()
    else if (seg !== '.') out.push(seg)
  }
  return out.join('/')
}

describe('buildPptx stays byte-identical', () => {
  const snapshots = JSON.parse(readFileSync('tests/__snapshots__/ooxml-builders-pptx.json', 'utf8')) as Record<string, { names: string[]; hashes: Record<string, string> }>
  const sameBytes = pngBytes()
  const plainShapes: PptxShapeSpec[] = [
    { prst: 'roundRect', off: ['914400', '914400'], ext: ['3657600', '2743200'], fill: 'FFCC00', paragraphs: [{ align: 'ctr', runs: [{ text: 'Slide Title', b: true, sz: '4400', color: '1F497D' }, { text: 'Body line' }] }] },
    { prst: 'ellipse', off: ['5486400', '914400'], ext: ['914400', '914400'] },
    { off: ['914400', '4114800'], ext: ['2286000', '914400'], fill: 'F2DCDB', rot: '5400000', paragraphs: [{ align: 'r', runs: [{ text: 'Rotated', i: true, sz: '1200', color: '3B3838' }] }] },
  ]
  const tableAndPicShapes: PptxShapeSpec[] = [
    { off: ['1828800', '914400'], ext: ['4572000', '1828800'], table: { colWidths: ['2286000', '2286000'], rows: [{ h: '500000', cells: [{ paragraphs: [{ runs: [{ text: 'A' }] }] }, { fill: 'DDEBF7', paragraphs: [{ runs: [{ text: 'B', b: true }] }] }] }], styleId: '{1A2B3C4D-0000-4000-8000-000000000001}', firstRow: true, bandRow: true } },
    { off: ['1828800', '3291840'], ext: ['1828800', '914400'], rot: '1200000', image: { data: sameBytes, srcRect: { l: 10000, t: 20000 } } },
    { off: ['4572000', '3291840'], ext: ['1828800', '914400'], image: { data: sameBytes } },
  ]

  it('buildPptx output is unchanged', async () => {
    expect(snapshots).toHaveProperty('plain')
    expect(snapshots).toHaveProperty('tableAndPic')
    const byCase: Record<string, PptxShapeSpec[]> = { plain: plainShapes, tableAndPic: tableAndPicShapes }
    for (const [caseName, shapes] of Object.entries(byCase)) {
      const actual = await shaByName(await buildPptx(shapes))
      const expected = snapshots[caseName]
      expect(actual.names).toEqual(expected.names)
      for (const [path, hash] of Object.entries(expected.hashes)) {
        expect(actual.hashes[path], path).toBe(hash)
      }
    }
  })
})

describe('buildPptxDeck structure', () => {
  const deck: PptxDeckSpec = {
    slideSize: [12192000, 6858000],
    masters: [
      {
        shapes: [{ prst: 'rect', off: ['0', '0'], ext: ['12192000', '274638'], fill: '1F4E79', id: 10, name: 'Band' }],
        layouts: [
          {
            placeholders: [{ type: 'title', off: ['838200', '1828800'], ext: ['7048500', '1325880'] }],
            shapes: [],
          },
          { shapes: [] },
        ],
      },
      { bg: { solid: '111111' }, layouts: [{ shapes: [] }] },
    ],
    slides: [
      { layout: [0, 0], shapes: [{ id: 2, name: 'Title', placeholder: { type: 'title' }, paragraphs: [{ runs: [{ text: 'Welcome' }] }] }, { id: 3, name: 'Mark', off: [ '1000125', '3000125' ], ext: ['1219200', '1219200'], fill: '3B6EF4' }] },
      { layout: [0, 1], shapes: [] },
      { layout: [1, 0], shapes: [] },
    ],
    theme: { bgFillStyles: ['<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>', '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>', '<a:solidFill><a:schemeClr val="accent3"/></a:solidFill>'] },
  }

  it('buildPptxDeck writes N slides, their layout and master rels, and a parseable theme', async () => {
    const buf = await buildPptxDeck(deck)
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slideWidthEmu).toBe(12192000)
    expect(doc.slides).toHaveLength(3)

    const entries = await textEntries(buf)
    expect(Object.keys(entries)).toContain('ppt/slideMasters/slideMaster1.xml')
    expect(Object.keys(entries)).toContain('ppt/slideMasters/slideMaster2.xml')
    expect(Object.keys(entries)).toContain('ppt/slideLayouts/slideLayout1.xml')
    expect(Object.keys(entries)).toContain('ppt/slideLayouts/slideLayout2.xml')
    expect(Object.keys(entries)).toContain('ppt/slideLayouts/slideLayout3.xml')
    expect(Object.keys(entries)).toContain('ppt/theme/theme1.xml')

    for (const part of Object.keys(entries)) {
      if (!part.endsWith('.xml') || part.endsWith('.rels')) continue
      for (const rel of await relsOf(entries, part)) {
        if (rel.external) continue
        expect(entries[resolveTarget(part, rel.target)], `${part} -> ${rel.target}`).toBeDefined()
      }
    }
    // presentation.xml links exactly the three slides and two masters with schema-valid ids
    const presentationRels = await relsOf(entries, 'ppt/presentation.xml')
    const slideRels = presentationRels.filter((r) => r.type.endsWith('/slide'))
    const masterRels = presentationRels.filter((r) => r.type.endsWith('/slideMaster'))
    expect(slideRels.map((r) => r.id)).toEqual(['rId1', 'rId2', 'rId3'])
    expect(masterRels.map((r) => r.id)).toEqual(['rId4', 'rId5'])
    const presentation = entries['ppt/presentation.xml']
    expect(presentation).toContain('<p:sldMasterIdLst>')
    expect(presentation).toMatch(/<p:sldMasterId id="2147483648" r:id="rId4"/)
    expect(presentation).toMatch(/<p:sldId id="256" r:id="rId1"/)
    expect(presentation).toContain('<p:sldSz cx="12192000" cy="6858000"/>')
    // every slide is wired to its own layout, and every layout to its master
    const slide1Layout = await relsOf(entries, 'ppt/slides/slide1.xml')
    expect(slide1Layout[0].type.endsWith('/slideLayout')).toBe(true)
    expect(slide1Layout[0].target).toContain('slideLayout1.xml')
    const layout3Rels = await relsOf(entries, 'ppt/slideLayouts/slideLayout3.xml')
    expect(layout3Rels[0].target).toContain('slideMaster2')
    // theme is parseable and surfaces the custom bgFillStyleLst
    for (const slide of doc.slides) {
      expect(slide.theme!.bgFillStyles.length).toBe(3)
      expect(slide.theme!.colors.accent1).toBeDefined()
    }
    // the theme part is a real DrawingML theme root, the custom bg fill list is in the XML (not the canned default), and the content type matches the Office theme type
    const themeXml = entries['ppt/theme/theme1.xml']
    expect(themeXml).toMatch(/<a:theme /)
    expect(themeXml).toContain('<a:themeElements>')
    expect(themeXml).toContain('<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>')
    expect(entries['[Content_Types].xml']).toContain('Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"')
  })

  it('presProps is written verbatim and linked from the presentation', async () => {
    const presProps = '<p:presentation xmlns:p="x"><p:props/>presentation-not-really</p:presentation>'
    const buf = await buildPptxDeck({ slides: [{ shapes: [] }], masters: [{ layouts: [{ shapes: [] }] }], presProps })
    const entries = await textEntries(buf)
    expect(entries['ppt/presProps.xml']).toBe(presProps)
    const pres = entries['ppt/presentation.xml']
    const presRels = await relsOf(entries, 'ppt/presentation.xml')
    expect(presRels.some((r) => r.type.endsWith('/presProps') && r.target === 'presProps.xml')).toBe(true)
    // every slide/master relationship id the presentation declares exists, and vice versa
    const declared = [...(pres.match(/<p:sldMasterIdLst>(.*?)<\/p:sldMasterIdLst>/s)?.[1] ?? '').matchAll(/r:id="([^"]+)"/g)].map((m) => m[1])
    for (const rel of presRels.filter((r) => r.type.endsWith('/slideMaster'))) {
      expect(declared, `${rel.id} referenced from sldMasterIdLst`).toContain(rel.id)
    }
  })

  it('slide sizes by spec and default slide size matches buildPptx', async () => {
    const buf = await buildPptxDeck({ slides: [{ shapes: [] }], masters: [{ layouts: [{ shapes: [] }] }] })
    const doc = await parsePptx(await OfficePackage.load(buf))
    expect(doc.slideWidthEmu).toBe(9144000)
    expect(doc.slideHeightEmu).toBe(6858000)
  })
})

describe('PptxBgSpec serialisation', () => {
  const deck: PptxDeckSpec = {
    masters: [
      { bg: { solid: '111111' }, layouts: [{ bg: { gradient: { angle: 45, stops: [[0, 'FFFFFF'], [100000, '1F4E79']] } }, shapes: [] }, { bg: { bgRef: { idx: 1002, color: 'accent1' } }, shapes: [] }] },
    ],
    slides: [
      { layout: [0, 0], bg: { picture: { data: pngBytes(), tile: true } }, shapes: [] },
      { layout: [0, 1], bg: { pattern: { prst: 'ltUpDiag', fg: '2E75B6', bg: 'FFFFFF' } }, shapes: [] },
    ],
  }

  it('each PptxBgSpec kind serialises to the matching p:bg child', async () => {
    const entries = await textEntries(await buildPptxDeck(deck))
    const master1 = entries['ppt/slideMasters/slideMaster1.xml']
    expect(master1).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>')
    const layout1 = entries['ppt/slideLayouts/slideLayout1.xml']
    expect(layout1).toContain('<a:gradFill')
    expect(layout1).toContain('<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>')
    expect(layout1).toContain('<a:gs pos="100000"><a:srgbClr val="1F4E79"/></a:gs>')
    expect(layout1).toMatch(/<a:lin ang="2700000"[^/]*\/>/)
    const layout2 = entries['ppt/slideLayouts/slideLayout2.xml']
    expect(layout2).toContain('<p:bg><p:bgRef idx="1002"><a:schemeClr val="accent1"/></p:bgRef></p:bg>')
    const slide1 = entries['ppt/slides/slide1.xml']
    expect(slide1).toContain('<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImg1"/><a:tile/></a:blipFill><a:effectLst/></p:bgPr></p:bg>')
    expect(entries['ppt/slides/_rels/slide1.xml.rels']).toContain('../media/')
    expect(Object.keys(entries).some((p) => p.startsWith('ppt/media/'))).toBe(true)
    const slide2 = entries['ppt/slides/slide2.xml']
    expect(slide2).toContain('<p:bg><p:bgPr><a:pattFill prst="ltUpDiag"><a:fgClr><a:srgbClr val="2E75B6"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill><a:effectLst/></p:bgPr></p:bg>')
    // content types cover slides, layouts, masters, theme and the image
    const ct = entries['[Content_Types].xml']
    expect(ct).toContain('Override PartName="/ppt/slides/slide1.xml"')
    expect(ct).toContain('Override PartName="/ppt/slideMasters/slideMaster1.xml"')
    expect(ct).toContain('Override PartName="/ppt/slideLayouts/slideLayout1.xml"')
    expect(ct).toContain('Override PartName="/ppt/theme/theme1.xml"')
    expect(ct).toContain('Default Extension="png"')
  })

  it('bgRef with a hex colour writes an srgbClr', async () => {
    const entries = await textEntries(await buildPptxDeck({ masters: [{ bg: { bgRef: { idx: 1003, color: '2E75B6' } }, layouts: [{ shapes: [] }] }], slides: [] }))
    expect(entries['ppt/slideMasters/slideMaster1.xml']).toContain('<p:bg><p:bgRef idx="1003"><a:srgbClr val="2E75B6"/></p:bgRef></p:bg>')
    // the default theme fill list is the canned phClr list, distinguishable from a custom one
    expect(entries['ppt/theme/theme1.xml']).toContain('<a:schemeClr val="phClr"')
  })
})

describe('showMasterSp', () => {
  it('slide and layout showMasterSp="0" are written only when false', async () => {
    const defaultEntries = await textEntries(await buildPptxDeck({
      masters: [{ layouts: [{ shapes: [] }] }],
      slides: [{ layout: [0, 0], shapes: [] }],
    }))
    for (const xml of ['ppt/slides/slide1.xml', 'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideMasters/slideMaster1.xml']) {
      expect(defaultEntries[xml]).not.toContain('showMasterSp')
    }
    const entries = await textEntries(await buildPptxDeck({
      masters: [{ layouts: [{ shapes: [], showMasterSp: false }] }],
      slides: [{ layout: [0, 0], shapes: [], showMasterSp: false }],
    }))
    expect(entries['ppt/slides/slide1.xml']).toMatch(/<p:sld [^>]*showMasterSp="0"/)
    expect(entries['ppt/slideLayouts/slideLayout1.xml']).toMatch(/<p:sldLayout [^>]*showMasterSp="0"/)
    expect(entries['ppt/slideMasters/slideMaster1.xml']).not.toContain('showMasterSp')
  })
})

describe('deck parse behaviour', () => {
  it('a slide placeholder without a transform inherits geometry from its layout', async () => {
    const doc = await parsePptx(await OfficePackage.load(await buildPptxDeck({
      masters: [{ layouts: [{ shapes: [], placeholders: [{ type: 'title', off: ['838200', '1828800'], ext: ['7048500', '1325880'] }, { type: 'body', idx: 1, off: ['838200', '3429000'], ext: ['6096000', '3657600'] }] }] }],
      slides: [{ layout: [0, 0], shapes: [{ placeholder: { type: 'title' }, paragraphs: [{ runs: [{ text: 'Inherited' }] }] }, { placeholder: { type: 'body', idx: 1 }, paragraphs: [{ runs: [{ text: 'Body' }] }] }] }],
    })))
    const [title, body] = doc.slides[0].shapes
    expect(title.placeholder).toEqual({ type: 'title', idx: 0 })
    expect(title.xEmu).toBe(838200)
    expect(title.widthEmu).toBe(7048500)
    expect(title.transformValid).toBe(true)
    expect(body.placeholder).toEqual({ type: 'body', idx: 1 })
    expect(body.yEmu).toBe(3429000)
    expect(body.heightEmu).toBe(3657600)
  })

  it('rejects a slide layout reference that no master defines', async () => {
    await expect(buildPptxDeck({ masters: [{ layouts: [] }], slides: [{ layout: [0, 0], shapes: [] }] })).rejects.toThrow()
    await expect(buildPptxDeck({ slides: [{ layout: [0, 0], shapes: [] }] })).rejects.toThrow()
  })

  it('shape id and name defaults match buildPptx and explicit ids win', async () => {
    const entries = await textEntries(await buildPptxDeck({
      masters: [{ layouts: [{ shapes: [] }] }],
      slides: [{ layout: [0, 0], shapes: [{ paragraphs: [{ runs: [{ text: 'Default' }] }] }, { id: 7, name: 'Second', paragraphs: [{ runs: [{ text: 'x' }] }] }] }],
    }))
    const slide = entries['ppt/slides/slide1.xml']
    expect(slide).toContain('<p:cNvPr id="2" name="Shape"/>')
    expect(slide).toContain('<p:cNvPr id="7" name="Second"/>')
  })
})

describe('deck XML hygiene', () => {
  const xmllintVersion = spawnSync('xmllint', ['--version'])
  const xmllintWorks = !xmllintVersion.error && xmllintVersion.status === 0
  const itIfXmllint = xmllintWorks ? it : it.skip
  const png = pngBytes()
  const lintDeck: PptxDeckSpec = {
    slideSize: [12192000, 6858000],
    masters: [
      {
        shapes: [{ prst: 'rect', off: ['0', '0'], ext: ['12192000', '274638'], fill: '1F4E79', id: 10, name: 'Band', paragraphs: [{ runs: [] }] }],
        layouts: [
          { placeholders: [{ type: 'title', off: ['838200', '1828800'], ext: ['7048500', '1325880'] }] },
          { showMasterSp: false, shapes: [{ prst: 'ellipse', off: ['0', '0'], ext: ['914400', '914400'], fill: '3B6EF4' }] },
        ],
      },
      { bg: { solid: '111111' }, shapes: [{ paragraphs: [{ runs: [{ text: 'default master shape' }] }] }], layouts: [{ shapes: [] }] },
    ],
    slides: [
      { layout: [0, 0], bg: { picture: { data: png, tile: true } }, shapes: [{ placeholder: { type: 'title' }, paragraphs: [{ runs: [{ text: 'Welcome' }] }] }, { paragraphs: [{ runs: [{ text: 'default id' }] }] }, { id: 9, name: 'Pic', off: ['0', '0'], ext: ['1', '1'], image: { data: png } }] },
      { layout: [0, 1], bg: { pattern: { prst: 'ltUpDiag', fg: '2E75B6', bg: 'FFFFFF' } }, shapes: [], extraXml: { beforeSpTree: '<p:custDataLst/>', afterCSld: '<p:custDataLst/>' } },
      { layout: [1, 0], shapes: [] },
    ],
    theme: { bgFillStyles: ['<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>', '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>', '<a:solidFill><a:schemeClr val="accent3"/></a:solidFill>'] },
  }

  itIfXmllint('every generated part is well-formed XML with bound namespace prefixes (xmllint --noout)', async () => {
    const buf = await buildPptxDeck(lintDeck)
    const zip = await JSZip.loadAsync(buf)
    for (const path of Object.keys(zip.files).filter((p) => p.endsWith('.xml'))) {
      const content = await zip.files[path].async('text')
      const run = spawnSync('xmllint', ['--noout', '-'], { input: content })
      const output = `${run.stdout?.toString() ?? ''}${run.stderr?.toString() ?? ''}`
      expect(output, path).toBe('')
      expect(run.status, `${path}: ${output}`).toBe(0)
    }
  })

  it('cNvPr ids are unique per part, defaulting to 2, 3, … skipping explicitly claimed ids', async () => {
    const entries = await textEntries(await buildPptxDeck(lintDeck))
    for (const [path, xml] of Object.entries(entries)) {
      if (!path.endsWith('.xml') || path.includes('/_rels/') || path.startsWith('[Content_Types]')) continue
      const ids = [...xml.matchAll(/<p:cNvPr id="([^"]+)"/g)].map((m) => m[1])
      const dupes = ids.filter((id, i) => ids.indexOf(id) !== i)
      expect(dupes, `duplicate cNvPr ids in ${path}: ${ids.join(',')}`).toEqual([])
    }
    // defaults are 2, 3, …: the unnamed slide shape after the placeholder takes 3; the master's unnamed shape (id 10 already claimed) takes 2
    const slide1 = entries['ppt/slides/slide1.xml']
    expect(slide1).toContain('<p:cNvPr id="2" name="title"/>')
    expect(slide1).toContain('<p:cNvPr id="3" name="Shape"/>')
  })

  it('sldMasterId and sldLayoutId values are unique across both lists and >= 2147483648', async () => {
    const entries = await textEntries(await buildPptxDeck(lintDeck))
    const masterIds = [...entries['ppt/presentation.xml'].matchAll(/<p:sldMasterId id="(\d+)"/g)].map((m) => m[1])
    const layoutIds = Object.keys(entries).filter((p) => /slideMasters\/slideMaster\d+\.xml$/.test(p)).flatMap((p) => [...entries[p].matchAll(/<p:sldLayoutId id="(\d+)"/g)].map((m) => m[1]))
    const all = [...masterIds, ...layoutIds]
    expect(all.length).toBeGreaterThan(0)
    for (const id of all) expect(Number(id)).toBeGreaterThanOrEqual(2147483648)
    expect(new Set(all).size).toBe(all.length)
  })

  it('extraXml.beforeSpTree lands between the background and the shape tree', async () => {
    const entries = await textEntries(await buildPptxDeck(lintDeck))
    const slide2 = entries['ppt/slides/slide2.xml']
    expect(slide2).toContain('</p:bg><p:custDataLst/><p:spTree>')
    expect(slide2).toContain('</p:cSld><p:custDataLst/>')
  })
})
