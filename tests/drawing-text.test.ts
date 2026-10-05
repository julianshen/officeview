import { expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { layoutTextBody, sortedBoundaryRange, type MeasureText } from '../src/drawing/text-layout'
import { paintTextBody } from '../src/drawing/text-paint'
import { clusterOrientation, verticalOrientation } from '../src/drawing/vertical-orientation'
import { buildTextIndex, findMatches } from '../src/core/search'
import { hitTest, textForRange } from '../src/core/selection'
import type { DrawingTextBody, TextDirection } from '../src/drawing/text'

const emu = (px: number) => px * 9525
const measure: MeasureText = (text, style) => ({ width: [...text].length * (style.fontSizePt ?? 12), ascent: 10, descent: 3 })

test('Word auto/exact/atLeast spacing resolves measured normal height without changing DrawingML spacing', () => {
  const make = (rule: 'auto' | 'exact' | 'atLeast', value: number): DrawingTextBody => ({ direction: 'eaVert', anchor: 't', wrap: true,
    insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0,
    paragraphs: ['文', '中'].map(text => ({ runs: [{ text, fontSizePt: 10 }], align: 'left', level: 0,
      wordLineSpacing: { rule, value }, spaceAfter: { kind: 'points', value: 10 }, spaceBefore: { kind: 'points', value: 3 } })) })
  const gap = (rule: 'auto' | 'exact' | 'atLeast', value: number, height: number) => {
    const lines = layoutTextBody(make(rule, value), 300, 200, text => ({ width: text.length * 10, ascent: height * .8, descent: height * .2, normalHeight: height })).lines
    return lines[0].x - lines[1].x
  }
  expect(gap('auto', 288, 16)).toBeCloseTo(16 * 1.2 + 13 * 96 / 72)
  expect(gap('auto', 288, 32)).toBeCloseTo(32 * 1.2 + 13 * 96 / 72)
  expect(gap('exact', 120, 40)).toBeCloseTo(8 + 13 * 96 / 72)
  expect(gap('atLeast', 120, 40)).toBeCloseTo(40 + 13 * 96 / 72)
  expect(gap('atLeast', 900, 40)).toBeCloseTo(60 + 13 * 96 / 72)
  const dml = make('auto', 288)
  dml.paragraphs.forEach(p => { delete p.wordLineSpacing; p.lineSpacing = { kind: 'percent', value: 1.2 } })
  const xs = [16, 32].map(height => layoutTextBody(dml, 300, 200, text => ({ width: text.length * 10, ascent: height * .8, descent: height * .2, normalHeight: height })).lines.map(line => line.x))
  expect(xs[0][0] - xs[0][1]).toBeCloseTo(xs[1][0] - xs[1][1])
})
function body(direction: TextDirection, source = 'AB12中文(甲)一乙\nCD34文句。蒙古'): DrawingTextBody {
  return { direction, paragraphs: [{ runs: [{ text: source, fontSizePt: 12, fontFamily: 'TestFace', color: '#123456' }], align: 'left', level: 0 }],
    anchor: 't', wrap: true, insetLeftEmu: emu(4), insetRightEmu: emu(6), insetTopEmu: emu(8), insetBottomEmu: emu(10) }
}

test('pinned Unicode orientation data and whole-cluster enclosing mark exception', () => {
  expect([0x41, 0x4e00, 0x3001, 0x3008, 0x1820, 0x1f600].map(verticalOrientation)).toEqual(['R', 'U', 'Tu', 'Tr', 'R', 'U'])
  expect(clusterOrientation('A\u20dd')).toBe('U')
  expect(clusterOrientation('A\u20d0')).toBe('R')
})

test.each(['horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'] as TextDirection[])(
  '%s retains source spans, explicit breaks, finite affine placements and wrapping columns', direction => {
    const result = layoutTextBody(body(direction), 90, 105, measure)
    expect(result.lines.length).toBeGreaterThan(1)
    const segments = result.lines.flatMap(line => line.segments)
    expect(segments.map(s => s.text).join('')).toBe('AB12中文(甲)一乙\nCD34文句。蒙古')
    expect(segments.every(s => s.style.fontFamily === 'TestFace' && s.style.color === '#123456')).toBe(true)
    expect(segments.every(s => [s.x, s.width, ...(s.transform ? Object.values(s.transform) : [])].every(Number.isFinite))).toBe(true)
    if (direction !== 'horz') expect(segments.filter(s => s.text !== '\n').every(s => s.transform && s.orientation)).toBe(true)
  },
)

test('mixed vertical orientation preserves shaped Latin and whole grapheme clusters', () => {
  const source = 'AB12中、é👩🏽‍💻ᠮᠣᠩᠭᠣᠯ'
  const result = layoutTextBody(body('eaVert', source), 200, 300, measure)
  const segments = result.lines.flatMap(line => line.segments)
  expect(segments.map(s => s.text).join('')).toBe(source)
  expect(segments.find(s => s.text === 'AB12')?.orientation).toBe('clockwise')
  expect(segments.find(s => s.text === '中')?.orientation).toBe('upright')
  expect(segments.find(s => s.text === '、')?.orientation).toBe('upright')
  expect(segments.find(s => s.text === 'e\u0301')?.orientation).toBe('clockwise')
  expect(segments.find(s => s.text === '👩🏽‍💻')?.orientation).toBe('upright')
  expect(segments.find(s => s.text.includes('ᠮ'))?.text).toBe('ᠮᠣᠩᠭᠣᠯ')
  expect(segments.every(s => s.runIndex === 0 && s.sourceEnd - s.sourceStart === s.text.length)).toBe(true)
  expect(segments.find(s => s.text === '👩🏽‍💻')?.graphemeBoundaries).toEqual([
    source.indexOf('👩'), source.indexOf('👩') + '👩🏽‍💻'.length,
  ])
})

test('vert rotates the whole line clockwise including CJK', () => {
  const result = layoutTextBody(body('vert', 'A中'), 200, 300, measure)
  const segments = result.lines.flatMap(line => line.segments)
  // Whole-line rotation shapes Latin and CJK together into one segment.
  expect(segments.map(s => s.text)).toEqual(['A中'])
  for (const segment of segments) expect(segment.orientation).toBe('clockwise')
})

test('vertical columns advance in the declared direction and alignment/anchor move placed runs', () => {
  for (const direction of ['vert', 'eaVert', 'wordArtVert', 'wordArtVertRtl', 'mongolianVert'] as TextDirection[]) {
    const b = body(direction, 'AAAA BBBB CCCC DDDD')
    const layout = layoutTextBody(b, 140, 50, measure)
    expect(layout.lines.length).toBeGreaterThan(1)
    const columns = layout.lines.map(l => l.segments.find(s => s.transform)?.transform?.e).filter((n): n is number => n !== undefined)
    const rtl = direction !== 'wordArtVert' && direction !== 'mongolianVert'
    expect(Math.sign(columns[1] - columns[0])).toBe(rtl ? -1 : 1)
    b.paragraphs[0].align = 'center'; b.anchor = 'ctr'
    const centered = layoutTextBody(b, 140, 80, measure)
    expect(centered.lines[0].segments[0].transform?.f).not.toBe(layout.lines[0].segments[0].transform?.f)
  }
})

test('native stacked columns: wordArtVert flows left-to-right, wordArtVertRtl right-to-left', () => {
  for (const [direction, sign] of [['wordArtVert', 1], ['wordArtVertRtl', -1]] as const) {
    const b = body(direction, 'AAAA BBBB CCCC DDDD')
    const layout = layoutTextBody(b, 140, 50, measure)
    expect(layout.lines.length).toBeGreaterThan(1)
    const columns = layout.lines.map(l => l.segments.find(s => s.transform)?.transform?.e).filter((n): n is number => n !== undefined)
    expect(columns.length).toBeGreaterThan(1)
    expect(Math.sign(columns[1] - columns[0])).toBe(sign)
  }
})

test('actual paint joins mixed orientation and stacked clusters for search/copy with grapheme-safe hits', async () => {
  const source = 'AB中e\u0301👩🏽‍💻CD'
  const b = body('eaVert', source)
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 300 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 200, 300, f => f) }])
  const match = findMatches(index, source)[0]
  expect(match?.logical?.source.text).toBe(source)
  expect(match.rects.length).toBeGreaterThan(1)
  const line = index.pages[0].lines[0]
  expect(line.text).toBe(source)
  const end = { pageIndex: 0, lineIndex: 0, charIndex: source.length }
  expect(textForRange(index, { start: { ...end, charIndex: 0 }, end })).toBe(source)
  const emoji = line.spans.find(s => s.text === '👩🏽‍💻')!
  const p = emoji.placement!
  const hit = hitTest(index, 0, p.transform.e + p.x + p.width / 2, p.transform.f + p.y, { strict: true })
  expect(hit).toBeDefined()
  expect([0, 2, 3, 5, 12, 14]).toContain(hit!.charIndex)
  expect(hitTest(index, 0, 199, 299, { strict: true })).toBeUndefined()
  const canvas = createCanvas(200, 300)
  paintTextBody(b, canvas.getContext('2d') as unknown as CanvasRenderingContext2D, 0, 0, 200, 300, f => f)
})

test('stacked word remains one logical searchable line after actual paint', async () => {
  const b = body('wordArtVert', 'STACK')
  const index = await buildTextIndex([{ spec: { widthPx: 80, heightPx: 200 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 80, 200, f => f) }])
  expect(index.pages[0].lines[0].text).toBe('STACK')
  expect(findMatches(index, 'STACK')).toHaveLength(1)
  expect(index.pages[0].lines[0].spans).toHaveLength(5)
})

test('copy follows source order across right-to-left vertical columns and preserves explicit breaks', async () => {
  const source = 'FIRST\nSECOND\nTHIRD'
  const b = body('eaVert', source)
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 200, 200, f => f) }])
  const lines = index.pages[0].lines
  expect(lines.map(line => line.text).join('')).toBe(source)
  expect(lines.length).toBe(3)
  expect(lines[0].spans[0].x).toBeGreaterThan(lines[1].spans[0].x)
  expect(lines[1].spans[0].x).toBeGreaterThan(lines[2].spans[0].x)
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 2, charIndex: lines[2].text.length } })).toBe(source)
  expect(findMatches(index, 'SECOND\nTHIRD')).toHaveLength(1)
})

test('outer nonuniform scale and rotation compose with each local direction transform', async () => {
  const b = body('eaVert', 'AB中')
  const index = await buildTextIndex([{ spec: { widthPx: 800, heightPx: 800 }, paint: ctx => {
    ctx.translate(300, 100); ctx.rotate(Math.PI / 6); ctx.scale(2, 3)
    paintTextBody(b, ctx, 0, 0, 150, 150, f => f)
  } }])
  const spans = index.pages[0].lines[0].spans
  expect(spans.map(s => s.text)).toEqual(['AB', '中'])
  const latin = spans[0].placement!.transform, han = spans[1].placement!.transform
  // The clockwise Latin baseline is the outer transform's vertical axis;
  // the upright Han baseline is its horizontal axis.
  expect(latin.a).toBeCloseTo(-1.5)
  expect(latin.b).toBeCloseTo(3 * Math.cos(Math.PI / 6))
  expect(han.a).toBeCloseTo(2 * Math.cos(Math.PI / 6))
  expect(han.b).toBeCloseTo(1)
  expect(findMatches(index, 'AB中')[0].rects).toHaveLength(2)
})

test('vertical painter preserves the caller canvas state and clips overflow', () => {
  const ctx = createCanvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.fillStyle = '#ff0000'; ctx.font = '9px serif'; ctx.globalAlpha = .5
  ctx.setTransform(1, .2, .1, 1, 3, 4)
  const transform = ctx.getTransform(), source = body('wordArtVert', 'LONG LONG LONG LONG LONG')
  paintTextBody(source, ctx, 0, 0, 25, 25, f => f)
  expect(ctx.getTransform()).toEqual(transform)
  expect(ctx.fillStyle).toBe('#ff0000')
  expect(ctx.font).toBe('9px serif')
  expect(ctx.globalAlpha).toBe(.5)
  // A fresh opaque mark outside the body would reveal a leaked clip.
  ctx.resetTransform(); ctx.fillRect(80, 80, 10, 10)
  expect(ctx.getImageData(85, 85, 1, 1).data[3]).toBe(128)
})

test('vertical run styles, script faces and UTF-16 offsets remain attached to their source runs', () => {
  const b = body('vert', '')
  b.paragraphs[0].runs = [
    { text: 'AB', fontFamily: 'LatinFace', bold: true, fontSizePt: 12 },
    { text: '中', fontFamily: 'LatinFace', fontFamilyEastAsia: 'HanFace', color: '#ff0000', fontSizePt: 12 },
    { text: 'e\u0301', fontFamily: 'LatinFace', italic: true, fontSizePt: 12 },
  ]
  const segments = layoutTextBody(b, 200, 200, measure).lines.flatMap(line => line.segments)
  expect(segments.map(s => [s.text, s.runIndex, s.sourceStart, s.sourceEnd, s.style.fontFamily])).toEqual([
    ['AB', 0, 0, 2, 'LatinFace'], ['中', 1, 2, 3, 'HanFace'], ['e\u0301', 2, 3, 5, 'LatinFace'],
  ])
  expect(segments.map(s => [s.style.bold, s.style.italic, s.style.color])).toEqual([
    [true, undefined, undefined], [undefined, undefined, '#ff0000'], [undefined, true, undefined],
  ])
})

test.each([
  { direction: 'horz', pieces: ['AB中'], rotations: [undefined], firstX: 4 },
  { direction: 'vert', pieces: ['AB中'], rotations: ['clockwise'], firstX: 130.9 },
  { direction: 'eaVert', pieces: ['AB', '中'], rotations: ['clockwise', 'upright'], firstX: 130.9 },
  { direction: 'eaVert', pieces: ['AB', '中'], rotations: ['clockwise', 'upright'], firstX: 130.9 },
  { direction: 'vert270', pieces: ['AB中'], rotations: ['counterclockwise'], firstX: 17.1 },
  { direction: 'wordArtVert', pieces: ['A', 'B', '中'], rotations: ['upright', 'upright', 'upright'], firstX: 17.1 - 6 },
  { direction: 'mongolianVert', pieces: ['AB', '中'], rotations: ['clockwise', 'upright'], firstX: 17.1 },
  { direction: 'wordArtVertRtl', pieces: ['A', 'B', '中'], rotations: ['upright', 'upright', 'upright'], firstX: 130.9 - 6 },
] as const)('$direction has independently placed fixed-measure runs', ({ direction, pieces, rotations, firstX }) => {
  const result = layoutTextBody(body(direction, 'AB中'), 150, 150, measure)
  const placed = result.lines[0].segments
  expect(placed.map(s => s.text)).toEqual(pieces)
  expect(placed.map(s => s.orientation)).toEqual(rotations)
  expect(placed[0].transform?.e ?? placed[0].x).toBeCloseTo(firstX)
  if (direction !== 'horz') {
    expect(placed[0].transform!.f).toBeGreaterThanOrEqual(8)
    const matrix = placed[0].transform!
    expect(matrix.a * matrix.d - matrix.b * matrix.c).toBe(1)
  }
})

test('Canvas paints visible upright and rotated glyph ink inside the local body clip', () => {
  for (const direction of ['vert', 'vert270', 'wordArtVert'] as TextDirection[]) {
    const canvas = createCanvas(180, 180), ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D
    paintTextBody(body(direction, 'AB中'), ctx, 20, 20, 120, 120, face => face)
    const alpha = ctx.getImageData(0, 0, 180, 180).data
    let inside = 0, outside = 0
    for (let y = 0; y < 180; y++) for (let x = 0; x < 180; x++) {
      if (alpha[(y * 180 + x) * 4 + 3] === 0) continue
      if (x >= 20 && x < 140 && y >= 20 && y < 140) inside++
      else outside++
    }
    expect(inside).toBeGreaterThan(10)
    expect(outside).toBe(0)
  }
})

test('vertical paragraphs and empty source lines copy in authored order despite right-to-left columns', async () => {
  const b = body('vert', 'FIRST')
  b.paragraphs.push({ runs: [], align: 'left', level: 0 }, { runs: [{ text: 'LAST', fontSizePt: 12 }], align: 'left', level: 0 })
  const index = await buildTextIndex([{ spec: { widthPx: 220, heightPx: 180 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 220, 180, f => f) }])
  const lines = index.pages[0].lines
  expect(lines.map(line => line.text)).toEqual(['FIRST', '', 'LAST'])
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 2, charIndex: 4 } })).toBe('FIRST\n\nLAST')
})

test.each([
  { name: 'combining', runs: ['e', '\u0301'], source: 'e\u0301', orientation: 'clockwise' },
  { name: 'enclosing', runs: ['A', '\u20dd'], source: 'A\u20dd', orientation: 'upright' },
  { name: 'emoji', runs: ['👩', '🏽‍💻'], source: '👩🏽‍💻', orientation: 'upright' },
])('complete $name cluster across styled runs keeps one cell and original provenance', ({ runs, source, orientation }) => {
  const b = body('eaVert', '')
  b.paragraphs[0].runs = runs.map((text, i) => ({ text, fontSizePt: 12, fontFamily: i ? 'SecondFace' : 'FirstFace', color: i ? '#ff0000' : '#000000' }))
  const segments = layoutTextBody(b, 200, 200, measure).lines.flatMap(line => line.segments)
  expect(segments).toHaveLength(1)
  expect(segments[0]).toMatchObject({ text: source, sourceStart: 0, sourceEnd: source.length, orientation })
  expect(segments[0].graphemeBoundaries).toEqual([0, source.length])
  expect(segments[0].sourceRuns.map(ref => [ref.runIndex, ref.start, ref.end, ref.style.fontFamily])).toEqual([
    [0, 0, runs[0].length, 'FirstFace'], [1, runs[0].length, source.length, 'SecondFace'],
  ])
})

test('vert270 starts at lower usable edge regardless of top inset, including breaks and tabs', () => {
  for (const source of ['AB', '\nAB', '\tAB']) {
    const first = body('vert270', source)
    first.insetTopEmu = 0; first.insetBottomEmu = emu(10)
    const second = { ...first, insetTopEmu: emu(8) }
    const a = layoutTextBody(first, 150, 150, measure).lines.flatMap(line => line.segments)
    const b = layoutTextBody(second, 150, 150, measure).lines.flatMap(line => line.segments)
    expect(b[0].transform?.f).toBeCloseTo(a[0].transform!.f)
    expect(b[0].transform?.f).toBeCloseTo(140)
  }
})

test.each([['A', '\u20dd'], ['e', '\u0301'], ['👩', '🏽‍💻']] as const)(
  'actual vertical paint records a single complete grapheme across runs %s + %s', async (first, second) => {
    const source = first + second, b = body('vert', '')
    b.paragraphs[0].runs = [
      { text: first, fontSizePt: 12, fontFamily: 'FirstFace' },
      { text: second, fontSizePt: 12, fontFamily: 'SecondFace', bold: true },
    ]
    const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 200, 200, f => f) }])
    const spans = index.pages[0].lines[0].spans
    expect(spans).toHaveLength(1)
    expect(spans[0].text).toBe(source)
    expect(spans[0].logical).toMatchObject({ start: 0, end: source.length, run: 0, graphemeBoundaries: [0, source.length] })
    expect(findMatches(index, source)).toHaveLength(1)
  },
)

test('horizontal run metadata advertises only complete paragraph grapheme boundaries without changing layout', () => {
  const b = body('horz', '')
  b.paragraphs[0].runs = [{ text: 'e', fontSizePt: 12 }, { text: '\u0301', fontSizePt: 12, bold: true }]
  const segments = layoutTextBody(b, 150, 150, measure).lines.flatMap(line => line.segments)
  expect(segments.map(segment => segment.text).join('')).toBe('e\u0301')
  expect(segments.flatMap(segment => segment.graphemeBoundaries)).toEqual([0, 2])
  expect(segments.flatMap(segment => segment.graphemeBoundaries)).not.toContain(1)
})

test.each(['vert', 'vert270', 'wordArtVert'] as TextDirection[])(
  '%s tabs record their true vertical advance and keep blank strict hits blank', async direction => {
    const b = body(direction, 'A\tB')
    b.insetLeftEmu = b.insetRightEmu = b.insetTopEmu = b.insetBottomEmu = 0
    b.wrap = false
    const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => paintTextBody(b, ctx, 0, 0, 100, 200, f => f) }])
    const tab = index.pages[0].lines[0].spans.find(span => span.text === '\t')!
    const rect = findMatches(index, '\t')[0].rects[0]
    expect(rect.height).toBeGreaterThan(rect.width * 2)
    expect(rect.width).toBeLessThan(30)
    expect(hitTest(index, 0, 140, 5, { strict: true })).toBeUndefined()
    const p = tab.placement!, t = p.transform
    const x = t.a * (p.x + p.width / 2) + t.c * p.y + t.e
    const y = t.b * (p.x + p.width / 2) + t.d * p.y + t.f
    expect(hitTest(index, 0, x, y, { strict: true })?.charIndex).toBe(2)
  },
)

test('wrapped source-boundary metadata does not iterate the whole paragraph per token', () => {
  const source = 'a '.repeat(1000), b = body('horz', source)
  b.insetLeftEmu = b.insetRightEmu = b.insetTopEmu = b.insetBottomEmu = 0
  const original = Set.prototype[Symbol.iterator]
  let wholeParagraphTraversals = 0
  Set.prototype[Symbol.iterator] = function () {
    if (this.size >= source.length / 2) wholeParagraphTraversals++
    return original.call(this)
  }
  try {
    const result = layoutTextBody(b, 100, 100, measure)
    expect(result.lines.length).toBeGreaterThan(100)
    expect(wholeParagraphTraversals).toBeLessThan(10)
  } finally { Set.prototype[Symbol.iterator] = original }
})

test('boundary ranges near a long paragraph end use logarithmic table reads', () => {
  const positions = Array.from({ length: 20001 }, (_, i) => i)
  let reads = 0
  const observed = new Proxy(positions, { get(target, key, receiver) {
    if (typeof key === 'string' && /^\d+$/.test(key)) reads++
    return Reflect.get(target, key, receiver)
  } })
  expect(sortedBoundaryRange(observed, 19995, 20000)).toEqual([19995, 20001])
  expect(reads).toBeLessThan(50)
})

test.each(['vert', 'vert270', 'wordArtVert'] as TextDirection[])(
  '%s tab hit geometry composes with rotated nonuniform outer scene transform', async direction => {
    const b = body(direction, 'A\tB')
    b.insetLeftEmu = b.insetRightEmu = b.insetTopEmu = b.insetBottomEmu = 0
    b.wrap = false
    const index = await buildTextIndex([{ spec: { widthPx: 600, heightPx: 600 }, paint: ctx => {
      ctx.translate(80, 50); ctx.rotate(Math.PI / 6); ctx.scale(2, 3)
      paintTextBody(b, ctx, 0, 0, 100, 200, f => f)
    } }])
    const tab = index.pages[0].lines[0].spans.find(span => span.text === '\t')!
    expect(tab.logical).toMatchObject({ start: 1, end: 2 })
    const p = tab.placement!, t = p.transform
    expect(Math.hypot(t.a, t.b)).toBeCloseTo(3)
    expect(Math.hypot(t.c, t.d)).toBeCloseTo(2)
    const x = t.a * (p.x + p.width * .75) + t.c * p.y + t.e
    const y = t.b * (p.x + p.width * .75) + t.d * p.y + t.f
    expect(hitTest(index, 0, x, y, { strict: true })?.charIndex).toBe(2)
    expect(hitTest(index, 0, 550, 550, { strict: true })).toBeUndefined()
    expect(findMatches(index, '\t')[0].rects).toHaveLength(1)
  },
)
