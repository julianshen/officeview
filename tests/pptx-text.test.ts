import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { existsSync, readFileSync } from 'node:fs'
import { createCanvas } from 'canvas'
import { parseXml } from '../src/core/xml'
import { parseTextBody, textFontDefaults } from '../src/pptx/text-parse'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import * as rendering from '../src/pptx/render'
import { paintTextBody } from '../src/pptx/text-paint'
import { parseThemeContext } from '../src/drawing/style'
import { buildPptx } from '../src/testdata/ooxml-builders'
import type { PptxParagraph, PptxTextBody, PptxTextRun, PptxTextStyle } from '../src/pptx/types'

async function bodyXml(xml: string) {
  const zip = await JSZip.loadAsync(await buildPptx([{ off: ['0', '0'], ext: ['9525000', '9525000'], paragraphs: [{ runs: [{ text: 'replace' }] }] }]))
  const path = 'ppt/slides/slide1.xml'
  zip.file(path, (await zip.file(path)!.async('string')).replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, `<p:txBody>${xml}</p:txBody>`))
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  return doc.slides[0].shapes[0].textBody!
}

function body(paragraphs: PptxParagraph[], extra: Partial<PptxTextBody> = {}): PptxTextBody {
  return { paragraphs, anchor: 't', insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0, wrap: true, ...extra }
}
const p = (runs: PptxTextRun[], extra: Partial<PptxParagraph> = {}): PptxParagraph => ({ runs, align: 'left', level: 0, ...extra })
const run = (text: string, extra: Partial<PptxTextRun> = {}): PptxTextRun => ({ text, fontSizePt: 12, ...extra })
// Fixed local advances avoid platform/installed-font-dependent line counts.
function layout(b: PptxTextBody, width = 100, height = 200) {
  return rendering.layoutTextBody(b, width, height, (text: string) => ({ width: [...text].length * 10, ascent: 12, descent: 4 }))
}
const texts = (l: ReturnType<typeof layout>) => l.lines.map(line => line.segments.map(s => s.text).join(''))

describe('DrawingML independent source text styles', () => {
  test('each run inherits paragraph defaults, never its previous sibling; false and zero stay explicit', async () => {
    const b = await bodyXml('<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2000" b="1" spc="80"><a:latin typeface="List"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:defRPr sz="1800"><a:latin typeface="Paragraph"/></a:defRPr></a:pPr><a:r><a:rPr b="0" i="1" spc="0"><a:latin typeface="Direct"/><a:ea typeface="EA"/><a:cs typeface="CS"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>A</a:t></a:r><a:r><a:t>B</a:t></a:r><a:endParaRPr sz="4200"/></a:p>')
    expect(b.paragraphs[0].runs[0]).toMatchObject({ text: 'A', bold: false, italic: true, fontSizePt: 18, characterSpacingPt: 0, color: '#FF0000', fontFamily: 'Direct', fontFamilyEastAsia: 'EA', fontFamilyComplexScript: 'CS' })
    expect(b.paragraphs[0].runs[1]).toMatchObject({ text: 'B', bold: true, fontSizePt: 18, characterSpacingPt: .8, color: '#112233', fontFamily: 'Paragraph' })
    expect(b.paragraphs[0].runs[1].italic).toBeUndefined()
    expect(b.paragraphs[0].runs[0].directProperties).toMatchObject({ bold: false, characterSpacingPt: 0 })
  })
  test('list paragraph attributes, spacing, tab stops, explicit breaks and end properties are retained', async () => {
    const b = await bodyXml('<a:bodyPr wrap="none"/><a:lstStyle><a:lvl2pPr marL="95250" marR="190500" indent="-47625" defTabSz="381000" algn="r"><a:buChar char="•"/><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:spcBef><a:spcPts val="300"/></a:spcBef></a:lvl2pPr></a:lstStyle><a:p><a:pPr lvl="1"><a:spcAft><a:spcPts val="600"/></a:spcAft><a:tabLst><a:tab pos="190500" algn="l"/></a:tabLst></a:pPr><a:r><a:t>  code\tvalue</a:t></a:r><a:br><a:rPr sz="2400"/></a:br><a:r><a:t> next</a:t></a:r><a:endParaRPr sz="3000"/></a:p><a:p><a:endParaRPr sz="3200"/></a:p>')
    expect(b.wrap).toBe(false)
    expect(b.paragraphs[0]).toMatchObject({ align: 'right', marginLeftEmu: 95250, marginRightEmu: 190500, indentEmu: -47625, defaultTabSizeEmu: 381000, bulletCharacter: '•', lineSpacing: { kind: 'percent', value: 1.25 }, spaceBefore: { kind: 'points', value: 3 }, spaceAfter: { kind: 'points', value: 6 }, tabStops: [{ positionEmu: 190500, align: 'left' }], endProperties: { fontSizePt: 30 } })
    expect(b.paragraphs[0].runs.map(r => r.text)).toEqual(['  code\tvalue', '\n', ' next'])
    expect(b.paragraphs[0].runs[1].fontSizePt).toBe(24)
    expect(b.paragraphs[1]).toMatchObject({ endProperties: { fontSizePt: 32 } })
  })
  test('theme font aliases resolve child Latin, East Asian and supplemental script faces at use time', () => {
    const theme = parseThemeContext('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="test"><a:majorFont><a:latin typeface="Major"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hant" typeface="Traditional"/><a:font script="Arab" typeface="Arabic"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
    const fn = rendering.resolveTextFamily
    expect(typeof fn).toBe('function')
    expect(fn({ text: 'English', fontFamily: '+mn-lt' }, theme)).toBe('Minor')
    expect(fn({ text: '中文', fontFamilyEastAsia: '+mj-ea', language: 'zh-TW' }, theme)).toBe('Traditional')
    expect(fn({ text: 'مرحبا', fontFamilyComplexScript: '+mj-cs' }, theme)).toBe('Arabic')
  })
})

describe('fixed-measure styled paragraph layout', () => {
  test('adjacent colored/bold runs share one line and conserve whitespace/source offsets', () => {
    const l = layout(body([p([run('  One ', { bold: true }), run('two', { color: '#ff0000' })])]))
    expect(texts(l)).toEqual(['  One two'])
    expect(l.lines[0].segments[0].style.bold).toBe(true)
    expect(l.lines[0].segments.at(-1)?.sourceStart).toBe(6)
  })
  test('one explicit break advances once and preserves following indentation', () => {
    expect(texts(layout(body([p([run('a'), run('\n'), run('  b')])])))).toEqual(['a\n', '  b'])
  })
  test('wrap none retains long text on one line', () => {
    expect(texts(layout(body([p([run('one two three')])], { wrap: false }), 40))).toEqual(['one two three'])
  })
  test('long tokens always make progress and combining/emoji clusters are indivisible', () => {
    const source = 'e\u0301👨‍👩‍👧‍👦abcd'
    const l = layout(body([p([run(source)])]), 20)
    expect(texts(l).join('')).toBe(source)
    expect(texts(l).slice(0, 2)).toEqual(['e\u0301', '👨‍👩‍👧‍👦'])
    expect(l.lines.length).toBeLessThan(10)
  })
  test('paragraph margins and hanging indent govern first and continued lines; bullets paint once', () => {
    const l = layout(body([p([run('aa bb cc')], { marginLeftEmu: 190500, marginRightEmu: 95250, indentEmu: -95250, bullet: true, bulletCharacter: '•' })]), 80)
    expect(l.lines[0].x).toBe(10)
    expect(l.lines[1].x).toBe(20)
    expect(l.lines.filter(line => line.bullet).length).toBe(1)
    expect(texts(l).join('')).toBe('aa bb cc')
  })
  test('explicit and default tab stops preserve source tabs and use stops rather than glyph width', () => {
    const l = layout(body([p([run('a\tb\tc')], { defaultTabSizeEmu: 381000, tabStops: [{ positionEmu: 285750, align: 'left' }] })]), 200)
    expect(texts(l)).toEqual(['a\tb\tc'])
    expect(l.lines[0].segments.find(s => s.text === 'b')?.x).toBe(30)
    expect(l.lines[0].segments.find(s => s.text === 'c')?.x).toBe(80)
  })
  test('mixed sizes share a baseline and percentage/point/before/after spacing determines block anchor', () => {
    const measure = rendering.layoutTextBody
    expect(typeof measure).toBe('function')
    const b = body([p([run('big', { fontSizePt: 24 }), run('small')], { lineSpacing: { kind: 'percent', value: 1.5 }, spaceBefore: { kind: 'points', value: 3 }, spaceAfter: { kind: 'points', value: 6 } }), p([run('tail')], { lineSpacing: { kind: 'points', value: 30 } })], { anchor: 'ctr' })
    const l = measure(b, 300, 200, (text: string, style) => ({ width: text.length * 5, ascent: style.fontSizePt ?? 12, descent: (style.fontSizePt ?? 12) / 3 }))
    expect(l.height).toBeCloseTo(109.6)
    expect(l.lines[0].y).toBeCloseTo(49.2)
    expect(l.lines[0].height).toBeCloseTo(57.6)
    expect(l.lines[0].baseline).toBeCloseTo(86)
    expect(l.lines[1].baseline).toBeCloseTo(138.8)
    expect(l.lines[1].height).toBe(40)
  })
  test('center and right align move an entire styled line as a block', () => {
    const l = layout(body([p([run('ab'), run('cd')], { align: 'center' }), p([run('ab'), run('cd')], { align: 'right' })]), 100)
    expect(l.lines[0].segments[0].x).toBe(30)
    expect(l.lines[1].segments[0].x).toBe(60)
  })
  test('end properties size an empty paragraph without forcing nonempty paragraph size', () => {
    const l = layout(body([p([run('normal')], { endProperties: { fontSizePt: 48 } }), p([], { endProperties: { fontSizePt: 30 } })]), 300)
    expect(l.lines[0].height).toBeCloseTo(19.2)
    expect(l.lines[1].height).toBeCloseTo(48)
  })
})

describe('actual canvas styled painting', () => {
  test('styled runs share baseline and preserve caller font/spacing/transform', async () => {
    const b = await bodyXml('<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2000" b="1"/></a:pPr><a:r><a:t>Bold </a:t></a:r><a:r><a:rPr b="0"/><a:t>normal</a:t></a:r></a:p>')
    const ctx = createCanvas(1000, 1000).getContext('2d')
    ctx.font = '7px serif'; ctx.translate(13, 17); ctx.scale(2, 3)
    const before = ctx.getTransform(); const draws: any[] = []
    const proxy = new Proxy(ctx, { get(target, key) { if (key === 'fillText') return (text: string, x: number, y: number) => draws.push({ text, x, y, font: target.font }); const v = (target as any)[key]; return typeof v === 'function' ? v.bind(target) : v }, set(target, key, value) { (target as any)[key] = value; return true } })
    rendering.renderSlide({ index: 0, widthEmu: 9525000, heightEmu: 9525000, shapes: [{ geometry: 'rect', xEmu: 0, yEmu: 0, widthEmu: 9525000, heightEmu: 9525000, textBody: b }] }, proxy as any)
    expect(draws.map(d => d.text).join('')).toBe('Bold normal')
    expect(new Set(draws.map(d => d.y)).size).toBe(1)
    expect(draws[0].font).toContain('bold')
    expect(draws.at(-1).font).not.toContain('bold')
    expect(ctx.font).toBe('7px serif')
    expect(ctx.getTransform()).toEqual(before)
  })
})

describe('paragraph and table precedence edge cases', () => {
  test('table region properties override generic defaults while direct false/color stay authoritative', async () => {
    const b = await bodyXml('<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr><a:r><a:t>region</a:t></a:r><a:r><a:rPr b="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>direct</a:t></a:r></a:p>')
    const ctx = createCanvas(300, 100).getContext('2d'), draws: any[] = []
    ctx.fillText = (text, _x, y) => { draws.push({ text, y, color: ctx.fillStyle, font: ctx.font }) }
    rendering.renderSlide({ index: 0, widthEmu: 300 * 9525, heightEmu: 100 * 9525, shapes: [{ geometry: 'rect', xEmu: 0, yEmu: 0, widthEmu: 300 * 9525, heightEmu: 100 * 9525, table: { colWidthsEmu: [300 * 9525], firstRow: true, firstRowTextColor: '#FFFFFF', firstRowBold: true, rows: [{ cells: [{ paragraphs: b.paragraphs, gridSpan: 1, rowSpan: 1 }] }] } }] }, ctx as any)
    expect(draws[0]).toMatchObject({ text: 'region', color: '#ffffff' })
    expect(draws[0].font).toContain('bold')
    expect(draws[1]).toMatchObject({ text: 'direct', color: '#ff0000' })
    expect(draws[1].font).not.toContain('bold')
    expect(draws[1].y).toBe(draws[0].y)
  })
  test('right/center/decimal tabs measure following styles up to the next stop', () => {
    const l = layout(body([p([run('a\t'), run('bc'), run('d\te')], { tabStops: [{ positionEmu: 952500, align: 'right' }, { positionEmu: 1428750, align: 'center' }] })]), 300)
    expect(l.lines[0].segments.find(s => s.text === 'bc')?.x).toBe(70)
    expect(l.lines[0].segments.find(s => s.text === 'e')?.x).toBe(145)
  })
  test('a grapheme crossing styled runs cannot be emergency wrapped between its marks', () => {
    const l = layout(body([p([run('e'), run('\u0301', { bold: true }), run('x')])]), 10)
    expect(texts(l)).toEqual(['e\u0301', 'x'])
  })
  test('positive first-line indent is removed on wrap continuations and bottom anchoring uses the whole block', () => {
    const l = layout(body([p([run('aa bb cc')], { indentEmu: 95250 })], { anchor: 'b', insetTopEmu: 95250, insetBottomEmu: 190500 }), 60, 100)
    expect(l.lines[0].x).toBe(10)
    expect(l.lines[1].x).toBe(0)
    expect(l.lines.at(-1)!.y + l.lines.at(-1)!.height).toBe(80)
  })
  test('point line spacing and paragraph list explicit false/zero overrides remain independent', async () => {
    const b = await bodyXml('<a:bodyPr/><a:lstStyle><a:lvl1pPr marL="95250" indent="95250" algn="ctr"><a:buChar char="•"/><a:defRPr b="1" i="1" spc="100"/></a:lvl1pPr></a:lstStyle><a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/><a:lnSpc><a:spcPts val="2000"/></a:lnSpc><a:defRPr b="0" i="0" spc="0"/></a:pPr><a:r><a:t>plain</a:t></a:r></a:p>')
    expect(b.paragraphs[0]).toMatchObject({ marginLeftEmu: 0, indentEmu: 0, align: 'left', bullet: false, lineSpacing: { kind: 'points', value: 20 } })
    expect(b.paragraphs[0].runs[0]).toMatchObject({ bold: false, italic: false, characterSpacingPt: 0 })
  })
})

test('presentation generic defaults and shape fontRef resolve through actual slide theme with explicit source priority', async () => {
  const zip = await JSZip.loadAsync(await buildPptx([{ paragraphs: [{ runs: [{ text: 'Inherited' }, { text: 'Explicit', b: true, color: 'FF0000' }] }] }]))
  const presentation = await zip.file('ppt/presentation.xml')!.async('string')
  zip.file('ppt/presentation.xml', presentation.replace('</p:presentation>', '<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="2200" b="1"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>'))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file('ppt/slides/slide1.xml', slide.replace('b="1"', 'b="0"').replace('<p:txBody>', '<p:style><a:fontRef idx="major"><a:schemeClr val="accent1"/></a:fontRef></p:style><p:txBody>'))
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/custom.xml"/></Relationships>')
  zip.file('ppt/theme/custom.xml', '<a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="00AAFF"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Theme Major"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>')
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  const runs = doc.slides[0].shapes[0].textBody!.paragraphs[0].runs
  expect(runs[0]).toMatchObject({ fontSizePt: 22, bold: true, fontFamily: '+mj-lt', color: '#00AAFF', directProperties: {} })
  expect(rendering.resolveTextFamily(runs[0], doc.slides[0].theme)).toBe('Theme Major')
  expect(runs[1]).toMatchObject({ bold: false, color: '#FF0000' })
})

test('layout and master placeholder styles inherit field by field while slide false and zero win', async () => {
  const zip = await JSZip.loadAsync(await buildPptx([{ paragraphs: [{ runs: [{ text: 'Inherited' }] }] }]))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file('ppt/slides/slide1.xml', slide.replace('<p:cNvSpPr/><p:nvPr/>', '<p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>').replace('<a:bodyPr', '<a:bodyPr wrap="none" lIns="0"'))
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr vert="eaVert" tIns="0"/><a:lstStyle><a:lvl1pPr marL="0"><a:defRPr b="0" spc="0"/></a:lvl1pPr></a:lstStyle><a:p><a:pPr algn="r"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>')
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr rIns="0" anchor="b"/><a:lstStyle><a:lvl1pPr><a:defRPr i="1"><a:ea typeface="Master EA"/></a:defRPr></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400" b="1"><a:latin typeface="+mj-lt"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle></p:txStyles></p:sldMaster>')
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  const b = doc.slides[0].shapes[0].textBody!
  expect(b).toMatchObject({ direction: 'eaVert', wrap: false, insetLeftEmu: 0, insetTopEmu: 0, insetRightEmu: 0, anchor: 'b' })
  expect(b.paragraphs[0]).toMatchObject({ align: 'right', marginLeftEmu: 0 })
  expect(b.paragraphs[0].runs[0]).toMatchObject({ fontSizePt: 44, bold: false, italic: true, characterSpacingPt: 0, fontFamily: '+mj-lt', fontFamilyEastAsia: 'Master EA', fontFamilyComplexScript: '+mj-cs' })
})

const retainedA11 = '/tmp/officeview-a11-20261003/a11-open-snapshot.pptx'
test.skipIf(!existsSync(retainedA11))('retained open a11 title resolves to its 44 point master title style', async () => {
  const doc = await parsePptx(await OfficePackage.load(readFileSync(retainedA11)))
  const title = doc.slides.flatMap(slide => slide.shapes).find(shape => shape.placeholder?.type === 'title' && shape.textBody?.paragraphs.some(p => p.runs.some(r => r.text === 'asaasaasasasasasas')))
  expect(title?.textBody?.paragraphs[0].runs[0].fontSizePt).toBe(44)
  expect(doc.slides[2].shapes.filter(shape => shape.source?.representation === 'fallback').map(shape => shape.source?.id)).toEqual(['3', '4', '5'])
  const imageFallback = doc.slides[2].shapes.find(shape => shape.source?.representation === 'fallback' && shape.source.id === '3')!
  expect(imageFallback.textBody?.paragraphs[0].runs[0].noFill).toBe(true)
  const ctx = createCanvas(1200, 600).getContext('2d'), painted: string[] = []
  ctx.fillText = text => { painted.push(text) }
  paintTextBody(imageFallback.textBody!, ctx as unknown as CanvasRenderingContext2D, 0, 0, 1100, 450, f => f)
  expect(painted).toEqual([])
})

const retainedDirections = '/tmp/officeview-text-directions-20261003/drawingml-native.pptx'
test.skipIf(!existsSync(retainedDirections))('controlled native slide retains all 14 label and body objects and seven direction tokens', async () => {
  const doc = await parsePptx(await OfficePackage.load(readFileSync(retainedDirections)))
  const samples = doc.slides[1].shapes.filter(shape => /^(label|sample)-/.test(shape.source?.name ?? ''))
  expect(samples).toHaveLength(14)
  expect(samples.filter(shape => shape.source?.name?.startsWith('sample-')).map(shape => shape.textBody?.direction)).toEqual([
    'horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl',
  ])
  expect(samples.every(shape => shape.textBody?.paragraphs.some(p => p.runs.some(r => r.text)))).toBe(true)
})

test('parsed vertical paragraphs wrap inside rotated nonuniformly scaled groups', async () => {
  const zip = await JSZip.loadAsync(await buildPptx([{ ext: ['190500', '190500'], paragraphs: [{ runs: [{ text: 'ABCDE FGHIJ KLMNO' }] }] }]))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const grouped = slide.replace(/(<p:sp>[\s\S]*?<\/p:sp>)/, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="group"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000"><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1524000"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="381000"/></a:xfrm></p:grpSpPr>$1</p:grpSp>`)
    .replace('<a:bodyPr/>', '<a:bodyPr vert="vert" wrap="square" lIns="0" rIns="0" tIns="0" bIns="0"/>')
  zip.file('ppt/slides/slide1.xml', grouped)
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  const child = doc.slides[0].shapes[0].children?.[0]
  expect(child?.textBody?.direction).toBe('vert')
  expect(layout(child!.textBody!, 20, 20).lines.length).toBeGreaterThan(1)
  const ctx = createCanvas(500, 500).getContext('2d')
  const transforms: Array<{ a: number; b: number; c: number; d: number }> = []
  ctx.fillText = () => { transforms.push(ctx.getTransform()) }
  rendering.renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D)
  expect(transforms.length).toBeGreaterThan(1)
  const m = transforms[0]
  expect(m.a * m.d - m.b * m.c).toBeCloseTo(8)
  expect(Math.abs(m.b) + Math.abs(m.c)).toBeGreaterThan(0)
})

test('placeholder ancestry maps layout index to master type and retains explicit slide overrides', async () => {
  const zip = await JSZip.loadAsync(await buildPptx([{ paragraphs: [{ runs: [{ text: 'Body' }] }] }]))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file('ppt/slides/slide1.xml', slide.replace('<p:cNvSpPr/><p:nvPr/>', '<p:cNvSpPr/><p:nvPr><p:ph type="obj" idx="2"/></p:nvPr>')
    .replace('<a:bodyPr/>', '<a:bodyPr lIns="0"/>').replace('<a:p><a:r>', '<a:p><a:pPr><a:defRPr b="0"/></a:pPr><a:r>'))
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr i="1"/></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>')
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr anchor="ctr" rIns="0" vert="horz"/><a:lstStyle><a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>')
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  const body = doc.slides[0].shapes[0].textBody!
  expect(body).toMatchObject({ anchor: 'ctr', direction: 'horz', insetLeftEmu: 0, insetRightEmu: 0 })
  expect(body.paragraphs[0].runs[0]).toMatchObject({ text: 'Body', fontSizePt: 36, italic: true, bold: false })
})

test('ordinary PPTX shape inherits master otherStyle without a placeholder', async () => {
  const zip = await JSZip.loadAsync(await buildPptx([{ paragraphs: [{ runs: [{ text: 'Ordinary' }] }] }]))
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout/>')
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster><p:txStyles><p:otherStyle><a:lvl1pPr><a:defRPr sz="2700" i="1"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>')
  const doc = await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))
  expect(doc.slides[0].shapes[0].textBody?.paragraphs[0].runs[0]).toMatchObject({ fontSizePt: 27, italic: true })
})

test.each(['pic', 'chart', 'clipArt', 'dgm', 'media', 'tbl'])('%s content placeholder inherits master body metadata', async role => {
  const zip = await JSZip.loadAsync(await buildPptx([{ paragraphs: [{ runs: [{ text: 'Content' }] }] }]))
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file('ppt/slides/slide1.xml', slide.replace('<p:cNvSpPr/><p:nvPr/>', `<p:cNvSpPr/><p:nvPr><p:ph type="${role}" idx="2"/></p:nvPr>`))
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="${role}" idx="2"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr i="1"/></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="3600" b="1"/></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld><p:txStyles><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="2700"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>')
  const body = (await parsePptx(await OfficePackage.load(await zip.generateAsync({ type: 'uint8array' })))).slides[0].shapes[0].textBody!
  expect(body.anchor).toBe('ctr')
  expect(body.paragraphs[0].runs[0]).toMatchObject({ fontSizePt: 36, bold: true, italic: true })
})

test.skipIf(!existsSync('corpus/pypptx-test.pptx'))('corpus title placeholder inherits centered master anchor', async () => {
  const doc = await parsePptx(await OfficePackage.load(readFileSync('corpus/pypptx-test.pptx')))
  const title = doc.slides[0].shapes.find(shape => shape.placeholder?.type === 'ctrTitle')
  expect(title?.textBody).toMatchObject({ anchor: 'ctr', direction: 'horz' })
})

test('explicit noFill text keeps inherited bullet from painting over a fallback image', async () => {
  const b = await bodyXml('<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buChar char="•"/><a:defRPr sz="2800"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p>')
  expect(b.paragraphs[0].runs[0]).toMatchObject({ text: ' ', noFill: true, fontSizePt: 28 })
  const ctx = createCanvas(200, 100).getContext('2d')
  const painted: string[] = []
  ctx.fillText = text => { painted.push(text) }
  paintTextBody(b, ctx as unknown as CanvasRenderingContext2D, 0, 0, 200, 100, f => f)
  expect(painted).toEqual([])
})

test('empty bullet paragraph honors explicit noFill defaults and end properties', () => {
  const b = parseTextBody(parseXml('<a:txBody><a:bodyPr/><a:p><a:pPr><a:buChar char="•"/><a:defRPr><a:noFill/></a:defRPr></a:pPr><a:endParaRPr><a:noFill/></a:endParaRPr></a:p></a:txBody>'))
  const ctx = createCanvas(200, 100).getContext('2d'), painted: string[] = []
  ctx.fillText = text => { painted.push(text) }
  paintTextBody(b, ctx as unknown as CanvasRenderingContext2D, 0, 0, 200, 100, f => f)
  expect(painted).toEqual([])
})


test('endParaRPr cannot inflate nonempty percentage paragraph before/after spacing', () => {
  const l = layout(body([p([run('text')], { endProperties: { fontSizePt: 48 }, spaceBefore: { kind: 'percent', value: 1 }, spaceAfter: { kind: 'percent', value: 1 } })]), 300)
  expect(l.lines[0].y).toBe(16)
  expect(l.height).toBeCloseTo(51.2)
})

test('Office normal leading matches independent controlled-native default and percentage evidence', () => {
  // Controlled authored native fixture: 20pt horizontal body without lnSpc
  // advances 32px. Independent 10.5pt + 120% source advances 20.16px.
  const normal = layout(body([p([run('one two', { fontSizePt: 20 })])]), 30)
  expect(normal.lines[0].height).toBe(32)
  expect(normal.lines[0].baseline).toBe(20)
  expect(normal.lines[1].baseline).toBe(52)
  const percent = layout(body([p([run('one two', { fontSizePt: 10.5 })], { lineSpacing: { kind: 'percent', value: 1.2 } })]), 30)
  expect(percent.lines[0].height).toBeCloseTo(20.16)
  expect(percent.lines[0].baseline).toBeCloseTo(14.08)
  const points = layout(body([p([run('one two', { fontSizePt: 20 })], { lineSpacing: { kind: 'points', value: 20 } })]), 30)
  expect(points.lines[0].height).toBeCloseTo(26.6666666667)
})

describe('ordinary wrapped paragraph justification', () => {
  test('wrapped nonfinal styled lines stretch only interior word gaps to the right margin', () => {
    const l = layout(body([p([run('aa ', { bold: true }), run('bb cc')], { align: 'justify' })]), 70)
    expect(texts(l)).toEqual(['aa bb ', 'cc'])
    const first = l.lines[0]
    expect(first.segments.find(s => s.text === ' ')!.width).toBe(20)
    expect(first.segments.at(-1)!.x + first.segments.at(-1)!.width).toBe(70)
    expect(l.lines[1].segments[0].x).toBe(0)
    expect(l.lines[1].segments[0].width).toBe(20)
  })
  test('hard breaks and single long words never justify, and tabs retain their stop positions', () => {
    const hard = layout(body([p([run('aa bb\ncc')], { align: 'justify' })]), 100)
    expect(texts(hard)).toEqual(['aa bb\n', 'cc'])
    expect(hard.lines[0].segments.find(s => s.text === ' ')?.width).toBe(10)
    const token = layout(body([p([run('abcdef')], { align: 'justify' })]), 25)
    expect(token.lines[0].segments.reduce((n, s) => n + s.width, 0)).toBe(20)
    const tabs = layout(body([p([run('a\tbb cc dd')], { align: 'justify', tabStops: [{ positionEmu: 285750, align: 'left' }] })]), 100)
    expect(tabs.lines[0].segments.find(s => s.text === 'bb')?.x).toBe(30)
  })
})

test('generic theme aliases choose supplemental Tamil and East Asian faces when the child slot is omitted', () => {
  const theme = parseThemeContext('<a:theme><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="Latin"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Taml" typeface="Tamil"/><a:font script="Hant" typeface="Traditional"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>')
  expect(rendering.resolveTextFamily({ text: 'தமிழ்', fontFamily: '+mn-lt' }, theme)).toBe('Tamil')
  expect(rendering.resolveTextFamily({ text: '中文', fontFamily: '+mn-lt', language: 'zh-TW' }, theme)).toBe('Traditional')
})

test('deferred distributed alignment retains the source token and a focused diagnostic', async () => {
  const b = await bodyXml('<a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="thaiDist"/><a:r><a:t>source</a:t></a:r></a:p>')
  expect(b.paragraphs[0]).toMatchObject({ sourceAlign: 'thaiDist', align: 'left' })
  expect(b.diagnostics).toEqual([expect.objectContaining({ kind: 'unsupported-text-alignment', feature: 'thaiDist' })])
})


test('ordinary word wrapping ignores styled run boundaries and preserves source ranges', () => {
  for (const runs of [[run('a hel'), run('lo', { bold: true })], [run('a h'), run('el', { color: '#ff0000' }), run('lo', { bold: true })]]) {
    const placed = layout(body([p(runs)]), 50)
    expect(texts(placed)).toEqual(['a ', 'hello'])
    expect(placed.lines[1].segments.map(s => [s.sourceStart, s.sourceEnd])).toEqual(runs.length === 2 ? [[2, 5], [5, 7]] : [[2, 3], [3, 5], [5, 7]])
    expect(placed.lines[1].segments.at(-1)!.style.bold).toBe(true)
  }
  expect(texts(layout(body([p([run('a hello')])]), 50))).toEqual(['a ', 'hello'])
})

test('split words use complete styled advances for tracking and bounded emergency wrapping', () => {
  for (const spacing of [-1, 2]) {
    const measure = (text: string, style: PptxTextStyle) => ({ width: text.length * (10 + (style.characterSpacingPt ?? 0)), ascent: 12, descent: 4 })
    const split = body([p([run('a hel', { characterSpacingPt: spacing }), run('lo', { bold: true, color: '#ff0000', characterSpacingPt: spacing })])])
    const unsplit = body([p([run('a hello', { characterSpacingPt: spacing })])])
    const width = 5 * (10 + spacing)
    const placed = rendering.layoutTextBody(split, width, 200, measure)
    expect(texts(placed)).toEqual(texts(rendering.layoutTextBody(unsplit, width, 200, measure)))
    expect(texts(placed)).toEqual(['a ', 'hello'])
  }
  const oversized = layout(body([p([run('a abc'), run('defghij', { bold: true })])]), 30)
  expect(texts(oversized)).toEqual(['a ', 'abc', 'def', 'ghi', 'j'])
  expect(texts(oversized).join('')).toBe('a abcdefghij')
  expect(oversized.lines.length).toBeLessThan(12)
})


test('fontRef solid/theme alpha defaults match run colors and preserve direct overrides', () => {
  const theme = parseThemeContext('<a:theme><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="00AAFF"/></a:accent1></a:clrScheme></a:themeElements></a:theme>')
  for (const alpha of [0, 50000, 100000]) {
    for (const [tag, value] of [['srgbClr', '00AAFF'], ['schemeClr', 'accent1']]) {
      const color = `<a:${tag} val="${value}"><a:alpha val="${alpha}"/></a:${tag}>`
      const defaults = textFontDefaults(parseXml(`<a:style><a:fontRef idx="minor">${color}</a:fontRef></a:style>`), theme)
      const expected = alpha === 100000 ? '#00AAFF' : `rgba(0,170,255,${alpha / 100000})`
      expect(defaults.color).toBe(expected)
      const parsed = parseTextBody(parseXml(`<p:txBody><a:bodyPr/><a:p><a:r><a:t>default</a:t></a:r><a:r><a:rPr><a:solidFill>${color}</a:solidFill></a:rPr><a:t>same direct</a:t></a:r><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>override</a:t></a:r></a:p></p:txBody>`), theme, undefined, defaults)
      expect(parsed.paragraphs[0].runs.map(r => r.color)).toEqual([expected, expected, '#FF0000'])
      expect(parsed.paragraphs[0].runs[0].propertySources?.color).toBe('default')
      expect(parsed.paragraphs[0].runs[2].propertySources?.color).toBe('run')
    }
  }
})
