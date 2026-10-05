import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { createCanvas } from 'canvas'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { collectDocImages } from '../src/docx/layout'
import { paintDrawing } from '../src/docx/drawing'
import { getPaintables } from '../src/render/paint'
import { CT_TYPES, ROOT_RELS } from '../src/testdata/ooxml-builders'
import type { DocxDrawing, DocxImage } from '../src/docx/types'
import { RECORD_TEXT, type LogicalTextRange } from '../src/core/text-recording'
const WNS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
const emu = (px: number) => px * 9525
function fixedBoxPaint(paragraphs: Extract<DocxDrawing, { kind: 'textbox' }>['paragraphs'], direction: Extract<DocxDrawing, { kind: 'textbox' }>['direction'] = 'horz', width = 200, height = 200, normalHeight = 16, recordSource = false) {
  const ctx = createCanvas(width, height).getContext('2d')
  Object.defineProperty(ctx.canvas, 'ownerDocument', { value: { createElement: () => ({ getContext: () => ({ font: '', measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: normalHeight * .75, actualBoundingBoxDescent: normalHeight * .25, fontBoundingBoxAscent: normalHeight * .75, fontBoundingBoxDescent: normalHeight * .25 }) }) }) } })
  const texts: Array<{ text: string; x: number; y: number; font: string; matrix: number[] }> = []
  const images: Array<{ image: DocxImage; x: number; y: number; w: number; h: number }> = []
  ctx.fillText = (text, x, y) => { const m = ctx.getTransform(); texts.push({ text, x, y, font: ctx.font, matrix: [m.a, m.b, m.c, m.d, m.e, m.f] }) }
  const records: Array<{ text: string; x: number; y: number; width: number; logical: LogicalTextRange }> = []
  if (recordSource) Object.assign(ctx, { [RECORD_TEXT]: (text: string, x: number, y: number, width: number, logical: LogicalTextRange) => { records.push({ text, x, y, width, logical }); ctx.fillText(text, x, y) } })
  ctx.drawImage = ((image: DocxImage, x: number, y: number, w: number, h: number) => images.push({ image, x, y, w, h })) as never
  paintDrawing({ kind: 'textbox', direction, paragraphs, fontFamily: 'Fixed', fontSizePt: 10, insets: { left: 0, top: 0, right: 0, bottom: 0 } }, ctx as never, width, height, { imageFor: image => image as never })
  return { texts, images, records }
}
const picInline = (rid: string, id: string, px = 24) =>
  `<w:drawing><wp:inline><wp:extent cx="${emu(px)}" cy="${emu(px)}"/><wp:docPr id="${id}" name="pic${id}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
const textboxDrawing = (inner: string, vert: string | undefined, anchor: boolean, id: string, wPx: number, hPx: number) => {
  const extent = `<wp:extent cx="${emu(wPx)}" cy="${emu(hPx)}"/>`
  const carrier = anchor
    ? `<wp:anchor><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>${extent}<wp:wrapNone/>`
    : `<wp:inline>${extent}`
  const close = anchor ? '</wp:anchor>' : '</wp:inline>'
  return `<w:drawing>${carrier}<wp:docPr id="${id}" name="box${id}"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr${vert === undefined ? '' : ` vert="${vert}"`}/></wps:wsp></a:graphicData></a:graphic>${close}</w:drawing>`
}
const txRun = (text: string) => `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`
const txPara = (inner: string) => `<w:p>${inner}</w:p>`

async function png(color: string): Promise<Buffer> {
  const { createCanvas: make } = await import('canvas')
  const canvas = make(24, 24)
  const ctx = canvas.getContext('2d')
  ctx.fillStyle = color
  ctx.fillRect(0, 0, 24, 24)
  return canvas.toBuffer('image/png')
}

interface NestedFixture {
  doc: Awaited<ReturnType<typeof parseDocx>>
  keyOf: (img: DocxImage) => string
}

async function nestedFixture(): Promise<NestedFixture> {
  const [red, blue, green, shared, yellow, cyan, gray] = await Promise.all(
    ['#ff0000', '#0000ff', '#00ff00', '#ff00ff', '#ffff00', '#00ffff', '#808080'].map(png)
  )
  const files: Record<string, string> = {
    imgA: 'media/a.png', imgB: 'media/b.png', imgC: 'media/c.png', imgD: 'media/shared.png',
    imgE: 'media/shared2.png', imgF: 'media/f.png', imgG: 'media/g.png', imgH: 'media/h.png'
  }
  const buffers: Record<string, Buffer> = { imgA: red, imgB: blue, imgC: green, imgD: shared, imgE: shared, imgF: yellow, imgG: cyan, imgH: gray }
  const t1 = textboxDrawing(
    `${txPara(`${txRun('Nested ')}<w:r>${picInline('imgC', '11')}</w:r>${txRun(' text')}`)}${txPara(`<w:r>${picInline('imgE', '12')}</w:r>`)}`,
    undefined, true, '10', 200, 100
  )
  const t2 = textboxDrawing(txPara(`${txRun('Cell ')}<w:r>${picInline('imgG', '13')}</w:r>`), undefined, false, '14', 200, 60)
  const th = textboxDrawing(txPara(`<w:r>${picInline('imgH', '15')}</w:r>`), undefined, false, '16', 200, 40)
  const body =
    `<w:p><w:r><w:t>Before </w:t></w:r><w:r>${picInline('imgA', '1')}</w:r><w:r><w:t> After</w:t></w:r></w:p>` +
    `<w:p><w:r>${picInline('imgB', '2')}</w:r></w:p>` +
    `<w:p><w:r>${t1}</w:r></w:p>` +
    `<w:p><w:r>${picInline('imgD', '3')}</w:r></w:p>` +
    `<w:p><w:r>${picInline('imgF', '4')}</w:r></w:p>` +
    `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r>${t2}</w:r></w:p></w:tc></w:tr></w:tbl>` +
    `<w:sectPr><w:headerReference w:type="default" r:id="hdr"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file('word/document.xml', `<w:document ${WNS}><w:body>${body}</w:body></w:document>`)
  const rels = (entries: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`
  const imgRel = (id: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`
  zip.file(
    'word/_rels/document.xml.rels',
    rels(Object.entries(files).filter(([id]) => id !== 'imgH').map(([id, target]) => imgRel(id, target)).join('') + '<Relationship Id="hdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>')
  )
  for (const [id, target] of Object.entries(files)) {
    if (id === 'imgH') continue
    zip.file(`word/${target}`, buffers[id])
  }
  zip.file('word/header1.xml', `<w:hdr ${WNS}><w:p><w:r>${th}</w:r></w:p></w:hdr>`)
  zip.file('word/_rels/header1.xml.rels', rels(imgRel('imgH', 'media/h.png')))
  zip.file('word/media/h.png', gray)
  const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  const keyOf = (img: DocxImage): string => {
    for (const [id, buffer] of Object.entries(buffers)) if (img.data.length === buffer.length && img.data.every((b, i) => b === buffer[i])) return id === 'imgD' || id === 'imgE' ? 'shared' : id
    return img.data.length === 0 ? 'vector' : 'unknown'
  }
  return { doc, keyOf }
}

describe('word nested textbox images', () => {
  test('Word textbox auto spacing responds to measured normal height and atLeast is a minimum', () => {
    const paragraphs = ['文', '中'].map(text => ({ runs: [{ text }], images: [], align: 'left' as const, lineSpacing: { rule: 'auto' as const, value: 288 }, spacingAfterTwips: 200 }))
    const gaps = [16, 32].map(height => {
      const painted = fixedBoxPaint(paragraphs, 'eaVert', 300, 200, height)
      return painted.texts[0].matrix[4] - painted.texts[1].matrix[4]
    })
    expect(gaps[0]).toBeCloseTo(16 * 1.2 + 200 / 15)
    expect(gaps[1]).toBeCloseTo(32 * 1.2 + 200 / 15)
    const para = { runs: [{ text: 'A' }], images: [], align: 'left' as const }
    const exact = fixedBoxPaint([{ ...para, lineSpacing: { rule: 'exact', value: 120 } }, para], 'horz', 200, 200, 40)
    const minimum = fixedBoxPaint([{ ...para, lineSpacing: { rule: 'atLeast', value: 120 } }, para], 'horz', 200, 200, 40)
    expect(minimum.texts[1].y).toBeGreaterThan(exact.texts[1].y)
  })

  test('an image-only wrap line records its authored source boundary between neighboring words', () => {
    const image: DocxImage = { data: new Uint8Array(), widthEmu: emu(24), heightEmu: emu(30) }
    const painted = fixedBoxPaint([{ runs: [{ text: 'BeforeAfter' }], images: [image], inline: [
      { kind: 'text', run: { text: 'Before' } }, { kind: 'image', image }, { kind: 'text', run: { text: 'After' } }
    ], align: 'left' }], 'horz', 70, 200, 16, true)
    expect(painted.records.map(record => record.text)).toEqual(['Before', '', 'After'])
    expect(painted.records.map(record => [record.logical.start, record.logical.end])).toEqual([[0, 6], [6, 6], [6, 11]])
    expect(painted.records.every(record => record.logical.source.text === 'BeforeAfter')).toBe(true)
    expect(painted.images).toHaveLength(1)
    expect(painted.images[0].y).toBeGreaterThan(painted.records[0].y)
    expect(painted.images[0].y + painted.images[0].h).toBeLessThan(painted.records[2].y)
    expect(painted.records[1].x).toBe(painted.images[0].x)
  })

  test('source-free inline slots allocate separate image-only positions and tall paragraph height', () => {
    const image = (width: number, height: number): DocxImage => ({ data: new Uint8Array(), widthEmu: emu(width), heightEmu: emu(height) })
    const a = image(24, 24), b = image(24, 24), tall = image(24, 60)
    const after = { runs: [{ text: 'After' }], images: [], align: 'left' as const }
    const alone = fixedBoxPaint([{ runs: [], images: [a, b], inline: [{ kind: 'image', image: a }, { kind: 'image', image: b }], align: 'left' }, after])
    expect(alone.images.map(box => box.image)).toEqual([a, b])
    expect(alone.images[1].x).toBeGreaterThanOrEqual(alone.images[0].x + 24)
    for (const box of alone.images) { expect(box.y).toBeGreaterThanOrEqual(0); expect(box.y + box.h).toBeLessThanOrEqual(200) }
    const painted = fixedBoxPaint([{ runs: [{ text: 'AB' }], images: [tall], inline: [{ kind: 'text', run: { text: 'A' } }, { kind: 'image', image: tall }, { kind: 'text', run: { text: 'B' } }], align: 'left' }, after])
    expect(painted.images[0]).toMatchObject({ x: 10, y: 0, w: 24, h: 60 })
    expect(painted.texts.find(call => call.text === 'After')!.y).toBeGreaterThan(60)
    expect(painted.texts.map(call => call.text)).toEqual(['A', 'B', 'After'])
  })

  test.each(['horz', 'vert', 'vert270', 'eaVert'] as const)('%s keeps fitting non-square image slots bounded with complete logical text', direction => {
    const image: DocxImage = { data: new Uint8Array(), widthEmu: emu(30), heightEmu: emu(60) }
    const painted = fixedBoxPaint([{ runs: [{ text: 'AB中' }], images: [image], inline: [{ kind: 'text', run: { text: 'AB' } }, { kind: 'image', image }, { kind: 'text', run: { text: '中' } }], align: 'left' }], direction)
    expect(painted.texts.map(call => call.text).join('')).toBe('AB中')
    expect(painted.images).toHaveLength(1)
    const box = painted.images[0]
    expect(box.x).toBeGreaterThanOrEqual(0); expect(box.y).toBeGreaterThanOrEqual(0)
    expect(box.x + box.w).toBeLessThanOrEqual(200); expect(box.y + box.h).toBeLessThanOrEqual(200)
    expect([box.w, box.h]).toEqual([30, 60])
  })

  test.each(['horz', 'vert', 'vert270'] as const)('%s preserves contextual and kerned authored runs around physical image boundaries', direction => {
    for (const text of ['مرحبا', 'नमस्ते', 'AV']) {
      const image: DocxImage = { data: new Uint8Array(), widthEmu: emu(24), heightEmu: emu(24) }
      const run = { text, fontFamily: 'ScriptFace' }
      for (const inline of [[{ kind: 'text' as const, run }], [{ kind: 'image' as const, image }, { kind: 'text' as const, run }, { kind: 'image' as const, image }]]) {
        const painted = fixedBoxPaint([{ runs: [run], images: inline.length > 1 ? [image, image] : [], inline, align: 'left' }], direction)
        expect(painted.texts.map(call => call.text)).toEqual([text])
        expect(painted.texts[0].font).toContain('ScriptFace')
      }
    }
  })
  test('legacy vertical models keep fallback direction, explicit direction wins, and caller state survives', () => {
    const paint = (vertical: boolean, direction?: 'horz' | 'eaVert') => {
      const ctx = createCanvas(200, 200).getContext('2d')
      ctx.font = '17px serif'; ctx.fillStyle = '#123456'; ctx.translate(9, 11)
      const calls: Array<{ text: string; matrix: number[] }> = []
      ctx.fillText = text => { const m = ctx.getTransform(); calls.push({ text, matrix: [m.a, m.b, m.c, m.d, m.e, m.f] }) }
      const drawing: DocxDrawing = { kind: 'textbox', vertical, direction, paragraphs: [{ runs: [{ text: 'AB中' }], images: [], align: 'left' }], fontFamily: 'Calibri', fontSizePt: 12, insets: { left: 0, right: 0, top: 0, bottom: 0 } }
      paintDrawing(drawing, ctx as never, 180, 180)
      expect(ctx.font).toBe('17px serif'); expect(ctx.fillStyle).toBe('#123456')
      expect([ctx.getTransform().e, ctx.getTransform().f]).toEqual([9, 11])
      return calls
    }
    expect(paint(true)).toEqual(paint(false, 'eaVert'))
    expect(paint(true).some(call => call.matrix[1] === 1)).toBe(true)
    expect(paint(false).every(call => call.matrix[1] === 0)).toBe(true)
    expect(paint(true, 'horz')).toEqual(paint(false))
  })
  test('collects nested assets in source order with stable body-before-header indices', async () => {
    const { doc, keyOf } = await nestedFixture()
    const got = collectDocImages(doc).map(keyOf)
    // imgD/imgE share bytes so both key as 'shared'; vector owners precede
    // their nested images in document order (T2 before its picG).
    expect(got).toEqual(['imgA', 'imgB', 'vector', 'imgC', 'shared', 'shared', 'imgF', 'vector', 'imgG', 'vector', 'imgH'])
  })

  test('decodes each unique asset once per document', async () => {
    const { doc } = await nestedFixture()
    const calls: Uint8Array[] = []
    const paintables = await getPaintables(doc, {
      decodeImage: (async (bytes: Uint8Array) => {
        calls.push(bytes)
        const { decodeImage } = await import('../src/core/images')
        return decodeImage(bytes).catch(() => undefined)
      }) as never
    })
    expect(paintables.length).toBeGreaterThan(0)
    const unique = new Set(calls.map(b => `${b.length}:${Array.from(b).join(',')}`))
    expect(unique.size).toBe(7)
    expect(calls.length).toBe(7)
  })

  test('paints every nested asset with finite positive placements in order', async () => {
    const { doc } = await nestedFixture()
    const paintables = await getPaintables(doc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d')
    const t0 = ctx.getTransform()
    paintables[0].paint(ctx as never)
    expect(ctx.getTransform()).toEqual(t0)
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    const has = (r: number, g: number, b: number) => {
      for (let i = 0; i < data.length; i += 4) {
        if (Math.abs(data[i] - r) < 64 && Math.abs(data[i + 1] - g) < 64 && Math.abs(data[i + 2] - b) < 64) return true
      }
      return false
    }
    for (const [r, g, b] of [[255, 0, 0], [0, 0, 255], [0, 255, 0], [255, 0, 255], [255, 255, 0], [0, 255, 255], [128, 128, 128]]) {
      expect(has(r, g, b)).toBe(true)
    }
    paintables.dispose()
  })

  test('vertical textbox preserves paragraph spacing and line spacing', async () => {
    const box = async (body: string, vert: string | undefined) => {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', CT_TYPES)
      zip.file('_rels/.rels', ROOT_RELS)
      zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(200)}"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${body}</w:txbxContent></wps:txbx><wps:bodyPr${vert === undefined ? '' : ` vert="${vert}"`}/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
      zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
      const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
      const drawing = collectDocImages(doc)[0]?.drawing
      if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
      return drawing
    }
    const spaced = (extra: string) =>
      `<w:p><w:pPr>${extra}</w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>AB</w:t></w:r></w:p>` +
      `<w:p><w:pPr>${extra}</w:pPr><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>CD</w:t></w:r></w:p>`
    const gapsOf = async (body: string, vert: string | undefined) => {
      const drawing = await box(body, vert)
      const ctx = createCanvas(200, 300).getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, 200, 300)
      paintDrawing(drawing as never, ctx as never, 200, 300)
      const data = ctx.getImageData(0, 0, 200, 300).data
      const rows: number[] = []
      for (let y = 0; y < 300; y++) {
        for (let x = 0; x < 200; x++) {
          if (data[(y * 200 + x) * 4] < 128) { rows.push(y); break }
        }
      }
      let gap = 0, seen = false
      for (let y = 0; y < 300; y++) {
        const ink = rows.includes(y)
        if (ink) { seen = true; if (gap > 0) break }
        else if (seen) gap++
      }
      return { bands: rows.length, gap }
    }
    // Horizontal flow: spaceAfter separates paragraphs by a visible band.
    const plain = await gapsOf(spaced(''), undefined)
    const spacedResult = await gapsOf(spaced('<w:spacing w:after="200" w:line="288" w:lineRule="auto"/>'), undefined)
    expect(plain.bands).toBeGreaterThan(0)
    expect(spacedResult.bands).toBeGreaterThan(0)
    // 200tw after-spacing (10pt) dominates the natural line gap.
    expect(spacedResult.gap).toBeGreaterThan(plain.gap + 8)
    // Isolate spaceAfter exactly: identical line rules, only after differs.
    // 200tw = 10pt ≈ 13.33px at 96dpi, independent of font advances.
    const lined = await gapsOf(spaced('<w:spacing w:line="288" w:lineRule="auto"/>'), undefined)
    expect(spacedResult.gap - lined.gap).toBeGreaterThan(10)
    expect(spacedResult.gap - lined.gap).toBeLessThan(17)
    // Vertical flow keeps both paragraphs painted with wider gaps.
    const vplain = await gapsOf(spaced(''), 'eaVert')
    const vspaced = await gapsOf(spaced('<w:spacing w:after="200" w:line="288" w:lineRule="auto"/>'), 'eaVert')
    expect(vplain.bands).toBeGreaterThan(0)
    expect(vspaced.bands).toBeGreaterThan(0)
    expect(vspaced.gap).toBeGreaterThanOrEqual(vplain.gap)
  })

  test('copied textbox text contains exactly the authored source without slot spaces', async () => {
    const green = await png('#00ff00')
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    const inner =
      `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Before </w:t></w:r>` +
      `<w:r><w:drawing><wp:inline><wp:extent cx="${emu(24)}" cy="${emu(24)}"/><wp:docPr id="31" name="pic31"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="imgS"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>` +
      `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> After</w:t></w:r></w:p>`
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/s.png"/></Relationships>')
    zip.file('word/media/s.png', green)
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = collectDocImages(doc)[0]?.drawing
    if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const { textForRange } = await import('../src/core/selection')
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 120 }, paint: (ctx) => paintDrawing(drawing as never, ctx as never, 300, 120) }])
    expect(findMatches(index, 'Before').length).toBeGreaterThan(0)
    expect(findMatches(index, 'After').length).toBeGreaterThan(0)
    const end = { pageIndex: 0, lineIndex: 0, charIndex: 14 }
    expect(textForRange(index, { start: { ...end, charIndex: 0 }, end })).toBe('Before  After')
  })

  test('vertical textbox composes nested image flow', async () => {
    const { doc, keyOf } = await nestedFixture()
    expect(collectDocImages(doc).map(keyOf)).toContain('imgC')
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    const green = await png('#00ff00')
    const inner = txPara(`${txRun('AB')}<w:r>${picInline('imgV', '21')}</w:r>`)
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r>${textboxDrawing(inner, 'vert', false, '20', 120, 200)}</w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgV" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/v.png"/></Relationships>`)
    zip.file('word/media/v.png', green)
    const vdoc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const paintables = await getPaintables(vdoc)
    const canvas = createCanvas(paintables[0].spec.widthPx, paintables[0].spec.heightPx)
    const ctx = canvas.getContext('2d')
    paintables[0].paint(ctx as never)
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data
    let greenPixels = 0
    for (let i = 0; i < data.length; i += 4) if (data[i + 1] > 200 && data[i] < 128 && data[i + 2] < 128) greenPixels++
    expect(greenPixels).toBeGreaterThan(0)
    paintables.dispose()
  })

  test('authored phrase search matches without slot spaces', async () => {
    const green = await png('#00ff00')
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    const inner =
      `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Before </w:t></w:r>` +
      `<w:r><w:drawing><wp:inline><wp:extent cx="${emu(24)}" cy="${emu(24)}"/><wp:docPr id="31" name="pic31"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="imgS"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>` +
      `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> After</w:t></w:r></w:p>`
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/s.png"/></Relationships>')
    zip.file('word/media/s.png', green)
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = collectDocImages(doc)[0]?.drawing
    if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 120 }, paint: (ctx) => paintDrawing(drawing as never, ctx as never, 300, 120) }])
    expect(findMatches(index, 'Before  After').length).toBe(1)
    expect(findMatches(index, 'Before   After').length).toBe(0)
    const hit = findMatches(index, 'Before  After')[0]
    expect(hit.logical?.source.text).toBe('Before  After')
  })

  test('leading, lone, consecutive and trailing images keep exact copy with ordered placements', async () => {
    const colors = ['#ff0000', '#00ff00', '#0000ff', '#ffff00', '#00ffff']
    const buffers = await Promise.all(colors.map(png))
    const ids = ['i1', 'i2', 'i3', 'i4', 'i5']
    const paras =
      txPara(`<w:r>${picInline('i1', '41')}</w:r>${txRun('After')}`) +
      txPara(`<w:r>${picInline('i2', '42')}</w:r>`) +
      txPara(`${txRun('A')}<w:r>${picInline('i3', '43')}</w:r><w:r>${picInline('i4', '44')}</w:r>${txRun('B')}`) +
      txPara(`${txRun('Z')}<w:r>${picInline('i5', '45')}</w:r>`)
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r>${textboxDrawing(paras, undefined, false, '40', 300, 220)}</w:r></w:p></w:body></w:document>`)
    const relsXml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${ids.map((id) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${id}.png"/>`).join('')}</Relationships>`
    zip.file('word/_rels/document.xml.rels', relsXml)
    ids.forEach((id, k) => zip.file(`word/media/${id}.png`, buffers[k]))
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const { buildTextIndex, findMatches } = await import('../src/core/search')
    const { textForRange } = await import('../src/core/selection')
    const paintables = await getPaintables(doc)
    const boxes: Array<{ x: number; y: number; w: number; h: number }> = []
    const paint = (ctx: CanvasRenderingContext2D) => {
      const proxy = new Proxy(ctx, {
        get(t, p) {
          if (p === 'drawImage') {
            return (img: CanvasImageSource, x: number, y: number, w: number, h: number) => {
              boxes.push({ x, y, w, h })
              return (t as CanvasRenderingContext2D).drawImage(img, x, y, w, h)
            }
          }
          const v = Reflect.get(t, p, t)
          return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
        },
      })
      paintables[0].paint(proxy as never)
    }
    const index = await buildTextIndex([{ spec: paintables[0].spec, paint }])
    const lines = index.pages[0].lines
    // The image-only paragraph records an empty line (no slot blank).
    expect(lines.map(l => l.text)).toEqual(['After', '', 'AB', 'Z'])
    for (const [i, line] of lines.entries()) {
      const copied = textForRange(index, { start: { pageIndex: 0, lineIndex: i, charIndex: 0 }, end: { pageIndex: 0, lineIndex: i, charIndex: line.text.length } })
      expect(copied).toBe(line.text)
    }
    expect(findMatches(index, 'After').length).toBe(1)
    // Five inline images paint in source order with finite positive boxes.
    expect(boxes.length).toBe(5)
    for (const b of boxes) {
      expect(Number.isFinite(b.x) && Number.isFinite(b.y)).toBe(true)
      expect(b.w).toBeGreaterThan(0)
      expect(b.h).toBeGreaterThan(0)
    }
    // Consecutive images sit side by side in source order (boxes 2 and 3;
    // boxes 0/1/4 are the leading, lone and trailing images).
    expect(Math.abs(boxes[3].x - (boxes[2].x + boxes[2].w))).toBeLessThan(3)
    expect(Math.abs(boxes[3].y - boxes[2].y)).toBeLessThan(3)
    paintables.dispose()
  })

  test('hit on an image band lands on the neighboring text caret', async () => {
    const green = await png('#00ff00')
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    const inner =
      `<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Before </w:t></w:r>` +
      `<w:r><w:drawing><wp:inline><wp:extent cx="${emu(24)}" cy="${emu(24)}"/><wp:docPr id="31" name="pic31"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="imgS"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>` +
      `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> After</w:t></w:r></w:p>`
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="${emu(200)}" cy="${emu(100)}"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/s.png"/></Relationships>')
    zip.file('word/media/s.png', green)
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const drawing = collectDocImages(doc)[0]?.drawing
    if (drawing?.kind !== 'textbox') throw Error(`expected textbox, got ${drawing?.kind}`)
    const { buildTextIndex } = await import('../src/core/search')
    const { hitTest } = await import('../src/core/selection')
    const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 120 }, paint: (ctx) => paintDrawing(drawing as never, ctx as never, 300, 120) }])
    const line = index.pages[0].lines[0]
    expect(line.text).toBe('Before  After')
    const midY = (line.top + line.bottom) / 2
    // The pad lives on the space cluster (line offset 6..7). Clicks across
    // the image band resolve to an adjacent caret, never a phantom char.
    let end = 0
    const padSpan = line.spans.find(s => { const s0 = end; end += s.text.length; return s0 < 7 && end >= 7 })
    expect(padSpan).toBeDefined()
    const inPad = hitTest(index, 0, padSpan!.x + padSpan!.width - 2, midY)
    expect([6, 7]).toContain(inPad?.charIndex)
    const afterPad = hitTest(index, 0, padSpan!.x + padSpan!.width + 2, midY)
    expect(afterPad).toMatchObject({ pageIndex: 0, lineIndex: 0, charIndex: 7 })
  })

  test('vertical image-first textbox paints the image with exact copy', async () => {
    const green = await png('#00ff00')
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CT_TYPES)
    zip.file('_rels/.rels', ROOT_RELS)
    const inner = txPara(`<w:r>${picInline('imgW', '51')}</w:r>${txRun('AB')}`)
    zip.file('word/document.xml', `<w:document ${WNS}><w:body><w:p><w:r>${textboxDrawing(inner, 'vert', false, '50', 120, 200)}</w:r></w:p></w:body></w:document>`)
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="imgW" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/w.png"/></Relationships>')
    zip.file('word/media/w.png', green)
    const doc = await parseDocx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
    const { buildTextIndex } = await import('../src/core/search')
    const paintables = await getPaintables(doc)
    const boxes: Array<{ x: number; y: number; w: number; h: number }> = []
    const index = await buildTextIndex([{
      spec: paintables[0].spec,
      paint: (ctx) => {
        const proxy = new Proxy(ctx, {
          get(t, p) {
            if (p === 'drawImage') {
              return (img: CanvasImageSource, x: number, y: number, w: number, h: number) => {
                boxes.push({ x, y, w, h })
                return (t as CanvasRenderingContext2D).drawImage(img, x, y, w, h)
              }
            }
            const v = Reflect.get(t, p, t)
            return typeof v === 'function' ? (v as (...a: never[]) => unknown).bind(t) : v
          },
        })
        paintables[0].paint(proxy as never)
      },
    }])
    // No image glyph enters the recorded text even in vertical flow.
    expect(index.pages[0].lines.map(l => l.text).join('')).toBe('AB')
    expect(boxes.length).toBe(1)
    expect(boxes[0].w).toBeGreaterThan(0)
    expect(boxes[0].h).toBeGreaterThan(0)
    paintables.dispose()
  })
})
