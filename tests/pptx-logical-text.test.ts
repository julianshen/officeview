import { describe, expect, test } from 'vitest'
import { createCanvas } from 'canvas'
import { buildTextIndex, createRecordingContext, findMatches, rectsForRange } from '../src/core/search'
import { hitTest, normalizeRange, rectsForSelectionOnPage, selectionSlices, textForRange } from '../src/core/selection'
import { renderSlide } from '../src/pptx/render'
import { paintTextBody } from '../src/pptx/text-paint'
import { graphemes } from '../src/core/text-recording'
import type { PptxShape, PptxSlide, PptxTextBody } from '../src/pptx/types'

const body = (text: string, wrap = true): PptxTextBody => ({ paragraphs: [{ runs: [{ text, fontSizePt: 12 }], align: 'left', level: 0 }], anchor: 't', wrap, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 })
function shape(b: PptxTextBody, width = 60, y = 0): PptxShape { return { geometry: 'rect', xEmu: 0, yEmu: y * 9525, widthEmu: width * 9525, heightEmu: 500 * 9525, textBody: b } }
async function indexFor(shapes: PptxShape[]) {
  const slide: PptxSlide = { index: 0, widthEmu: 1000 * 9525, heightEmu: 1000 * 9525, shapes }
  return buildTextIndex([{ spec: { widthPx: 1000, heightPx: 1000 }, paint: ctx => renderSlide(slide, ctx) }])
}

describe('source-scoped text from actual PPTX paint/index', () => {
  test('phrase search crosses adjacent styles and visual wraps without synthetic newlines', async () => {
    const b = body('')
    b.paragraphs[0].runs = [{ text: '  alpha ', bold: true, fontSizePt: 12 }, { text: 'beta gamma', fontSizePt: 12, color: '#ff0000' }]
    const index = await indexFor([shape(b)])
    expect(index.pages[0].lines.length).toBeGreaterThan(1)
    expect(findMatches(index, 'alpha beta gamma')).toHaveLength(1)
    expect(findMatches(index, 'alpha beta gamma')[0].rects.length).toBeGreaterThan(1)
    const lines = index.pages[0].lines
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines.at(-1)!.text.length } })).toBe('  alpha beta gamma')
  })
  test('source tabs, leading indent and explicit breaks survive painting and copy', async () => {
    const source = '  code\tvalue\n    next👩🏽‍💻e\u0301'
    const index = await indexFor([shape(body(source), 500)])
    const lines = index.pages[0].lines
    expect(findMatches(index, 'value\n    next')).toHaveLength(1)
    expect(lines.flatMap(line => line.spans).map(s => s.text).join('')).toBe(source)
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines.at(-1)!.text.length } })).toBe(source)
  })
  test('same baseline unrelated boxes/cells and reused paragraph objects never form phrases', async () => {
    const shared = body('alpha', false)
    const second = shape(body('beta', false), 100); second.xEmu = 100 * 9525
    const third = shape(shared, 100); third.xEmu = 200 * 9525
    const index = await indexFor([shape(shared, 100), second, third])
    expect(findMatches(index, 'alpha')).toHaveLength(2)
    expect(findMatches(index, 'alphabeta')).toHaveLength(0)
    expect(findMatches(index, 'betaalpha')).toHaveLength(0)
  })
  test('caret and copy offsets cannot split combining or emoji clusters', async () => {
    const source = 'Ae\u0301👩🏽‍💻Z'
    const index = await indexFor([shape(body(source, false), 1000)])
    const line = index.pages[0].lines[0]
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 2 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 5 } })).toBe('e\u0301👩🏽‍💻')
    const span = line.spans[0]
    const caret = hitTest(index, 0, span.x + span.width * 5 / source.length, span.y)!
    expect([0, 1, 3, 10, 11]).toContain(caret.charIndex)
  })
  test('tracking fallback records a single logical segment with tracked geometry', () => {
    const b = body('AB', false); b.paragraphs[0].runs[0].characterSpacingPt = 6
    const ctx = createCanvas(500, 500).getContext('2d') as unknown as CanvasRenderingContext2D
    const spans: Parameters<typeof createRecordingContext>[1] = []
    paintTextBody(b, createRecordingContext(ctx, spans), 0, 0, 500, 100, f => f)
    ctx.font = '12pt Calibri'
    expect(spans).toHaveLength(1)
    expect(spans[0].text).toBe('AB')
    expect(spans[0].width).toBeCloseTo(ctx.measureText('AB').width + 16)
  })
})

test('empty source paragraphs between selected paragraphs remain source breaks', async () => {
  const b = body('first')
  b.paragraphs.push({ runs: [], align: 'left', level: 0, endProperties: { fontSizePt: 20 } }, { runs: [{ text: 'last' }], align: 'left', level: 0 })
  const index = await indexFor([shape(b, 500)])
  const lines = index.pages[0].lines
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: 4 } })).toBe('first\n\nlast')
})

test('affine caller placement preserves local wraps and recorded tracked widths', () => {
  const b = body('AB CD EF'); b.paragraphs[0].runs[0].characterSpacingPt = 1
  const capture = (transformed: boolean) => {
    const ctx = createCanvas(200, 200).getContext('2d') as unknown as CanvasRenderingContext2D
    if (transformed) ctx.setTransform(-2, .3, .7, 3, 20, 30)
    const original = ctx.getTransform(), spans: Parameters<typeof createRecordingContext>[1] = []
    paintTextBody(b, createRecordingContext(ctx, spans), 0, 0, 40, 200, f => f)
    expect(ctx.getTransform()).toEqual(original)
    return spans
  }
  const a = capture(false), bPlaced = capture(true)
  expect(bPlaced.map(s => ({ text: s.text, x: s.placement!.x, y: s.placement!.y, width: s.placement!.width }))).toEqual(a.map(s => ({ text: s.text, x: s.placement!.x, y: s.placement!.y, width: s.placement!.width })))
})

test('native Canvas letterSpacing and deterministic headless fallback measure/place the same tracking', () => {
  const capture = (native: boolean, text = 'AB', spacingPt = 6, align: 'left' | 'center' | 'right' = 'left', width = 100, styled = false) => {
    const glyphs: Array<{ text: string; x: number }> = []
    const draws: Array<{ text: string; x: number; spacing?: string }> = []
    const measuredLanguages: string[] = [], paintedLanguages: string[] = []
    // A Canvas test double supplies deterministic local advances and the native
    // letterSpacing capability boundary; neither path depends on installed fonts.
    class FixedCanvas {
      width = 500; height = 100
      ownerDocument = { createElement: () => new FixedCanvas() }
      context: any
      constructor() {
        const stack: any[] = []
        this.context = { canvas: this, lang: 'fr', font: '7px serif', fillStyle: '#000', textAlign: 'left', textBaseline: 'alphabetic',
          ...(native ? { letterSpacing: '3px' } : {}),
          save() { stack.push({ lang: this.lang, font: this.font, fillStyle: this.fillStyle, textAlign: this.textAlign, textBaseline: this.textBaseline, ...(native ? { letterSpacing: this.letterSpacing } : {}) }) },
          restore() { Object.assign(this, stack.pop()) },
          measureText(text: string) { measuredLanguages.push(this.lang); const count = graphemes(text).filter(g => !/^[\p{Default_Ignorable_Code_Point}\ufffc]+$/u.test(g.text)).length; return { width: count * 10 + (native ? parseFloat(this.letterSpacing) * count : 0), actualBoundingBoxAscent: 12, actualBoundingBoxDescent: 4 } },
          getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } },
          fillText(text: string, x: number) { paintedLanguages.push(this.lang); draws.push({ text, x, spacing: this.letterSpacing }); graphemes(text).filter(g => !/^[\p{Default_Ignorable_Code_Point}\ufffc]+$/u.test(g.text)).forEach((g, i) => glyphs.push({ text: g.text, x: x + i * (10 + (native ? parseFloat(this.letterSpacing) : 0)) })) },
        }
      }
      getContext() { return this.context }
    }
    const ctx = new FixedCanvas().getContext() as CanvasRenderingContext2D
    const b = body(text); b.paragraphs[0].align = align
    b.paragraphs[0].runs = (styled ? graphemes(text).map(g => g.text) : [text]).map((t, i) => ({ text: t, fontSizePt: 12, characterSpacingPt: spacingPt, language: 'ja', bold: i % 2 === 1 }))
    paintTextBody(b, ctx, 0, 0, width, 100, f => f)
    const spans: Parameters<typeof createRecordingContext>[1] = []
    paintTextBody(b, createRecordingContext(ctx, spans), 0, 0, width, 100, f => f)
    expect(ctx.font).toBe('7px serif')
    if (native) expect(ctx.letterSpacing).toBe('3px')
    expect((ctx as any).lang).toBe('fr')
    return { draws, glyphs, spans, measuredLanguages, paintedLanguages }
  }
  const native = capture(true), fallback = capture(false)
  expect(native.draws).toEqual([{ text: 'AB', x: 0, spacing: '8px' }])
  expect(fallback.draws.map(d => ({ text: d.text, x: d.x }))).toEqual([{ text: 'A', x: 0 }, { text: 'B', x: 18 }])
  expect(native.spans[0].width).toBe(36)
  expect(fallback.spans[0].width).toBe(36)
  expect(native.spans[0].text).toBe(fallback.spans[0].text)
  expect(new Set([...native.measuredLanguages, ...native.paintedLanguages])).toEqual(new Set(['ja']))
  expect(new Set([...fallback.measuredLanguages, ...fallback.paintedLanguages])).toEqual(new Set(['ja']))
  // Primary WPT pins n × spacing advance, including the trailing gap. Blink
  // places LTR added spacing after each glyph, so the initial origin is fixed.
  for (const [text, spacing, align, width, styled] of [
    ['A', 6, 'left', 100, false], ['', 6, 'left', 100, false],
    ['AB', -3, 'center', 100, false], ['AB', 6, 'center', 100, false],
    ['AB', 6, 'right', 100, true], ['ABCD', 6, 'center', 35, false],
    ['e\u0301👩🏽‍💻', 3, 'left', 100, true],
    ['A\rB', 6, 'left', 100, false],
    ['A\u200bB', 6, 'left', 100, false], ['A\u200bB', -3, 'right', 100, true],
    ['\u200dAe\u0301👩🏽‍💻\ufeff', 3, 'center', 100, false],
  ] as const) {
    const native = capture(true, text, spacing, align, width, styled), fallback = capture(false, text, spacing, align, width, styled)
    expect(fallback.spans.map(s => ({ text: s.text, x: s.x, y: s.y, width: s.width }))).toEqual(native.spans.map(s => ({ text: s.text, x: s.x, y: s.y, width: s.width })))
    expect(fallback.glyphs).toEqual(native.glyphs)
    expect(native.spans.map(s => s.text).join('')).toBe(text)
  }
  expect(capture(false, 'A').spans[0].width).toBe(18)
  expect(capture(false, '').spans[0].width).toBe(0)
  expect(capture(false, 'AB', -3).spans[0].width).toBe(12)
  expect(capture(false, 'AB', 6, 'center').spans[0].x).toBe(32)
  expect(capture(false, 'AB', 6, 'right', 100, true).glyphs.map(g => g.x)).toEqual([64, 82])
  expect(capture(false, 'ABCD', 6, 'center', 35).spans).toHaveLength(4)

})

test('justified wrapped adjacent styles retain logical phrase search and exact source copy', async () => {
  const b = body(''); b.paragraphs[0].align = 'justify'
  b.paragraphs[0].runs = [{ text: '  alpha ', fontSizePt: 12, bold: true }, { text: 'beta gamma delta', fontSizePt: 12 }]
  const index = await indexFor([shape(b, 100)])
  const lines = index.pages[0].lines
  expect(findMatches(index, 'alpha beta gamma delta')).toHaveLength(1)
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines.at(-1)!.text.length } })).toBe('  alpha beta gamma delta')
})

test('empty paragraph metrics resolve the same theme font aliases as nonempty text', async () => {
  const { parseThemeContext } = await import('../src/drawing/style')
  const theme = parseThemeContext('<a:theme><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Theme Major"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>')
  const b = body(''); b.paragraphs[0].defaultProperties = { fontFamily: '+mj-lt', fontSizePt: 20 }
  const families: string[] = [], ctx = createCanvas(100, 100).getContext('2d')
  paintTextBody(b, ctx as any, 0, 0, 100, 100, f => { families.push(f); return f }, theme)
  expect(families).toContain('Theme Major')
  expect(families).not.toContain('+mj-lt')
})

test('grapheme fallback preserves modifier emoji, flags, variation selectors and ZWJ sequences', async () => {
  const { graphemes, snapGrapheme } = await import('../src/core/text-recording')
  const descriptor = Object.getOwnPropertyDescriptor(Intl, 'Segmenter')!
  Object.defineProperty(Intl, 'Segmenter', { configurable: true, value: undefined })
  try {
    expect(graphemes('👩🏽‍💻🇹🇼✈️e\u0301').map(g => g.text)).toEqual(['👩🏽‍💻', '🇹🇼', '✈️', 'e\u0301'])
    expect(snapGrapheme('🇹🇼', 2, 'ceil')).toBe(4)
  } finally { Object.defineProperty(Intl, 'Segmenter', descriptor) }
})

test('large end/break formatting cannot enlarge a nonempty line recorded text band', async () => {
  const b = body(''); b.paragraphs[0].runs = [{ text: 'A', fontSizePt: 12 }, { text: '\n', fontSizePt: 48 }, { text: 'B', fontSizePt: 12 }]
  const index = await indexFor([shape(b, 500)])
  expect(index.pages[0].lines[0].bottom - index.pages[0].lines[0].top).toBeCloseTo(16 * 1.1)
  expect(findMatches(index, 'A\nB')).toHaveLength(1)
})


test('headless tracking fallback retains a whole cursive/Indic shaped run and source properties', () => {
  for (const text of ['مرحبا', 'नमस्ते']) {
    const b = body(text, false); b.paragraphs[0].runs[0].characterSpacingPt = 6
    const ctx = createCanvas(500, 100).getContext('2d'), calls: string[] = []
    ctx.fillText = text => { calls.push(text) }
    paintTextBody(b, ctx as any, 0, 0, 500, 100, f => f)
    const spans: Parameters<typeof createRecordingContext>[1] = []
    paintTextBody(b, createRecordingContext(ctx as any, spans), 0, 0, 500, 100, f => f)
    expect(calls).toEqual([text])
    ctx.font = '12pt Calibri'
    expect(spans[0].width).toBeCloseTo(ctx.measureText(text).width)
    expect(b.paragraphs[0].runs[0].characterSpacingPt).toBe(6)
  }
})

test('headless tracking glyph origins retain measured Latin pair kerning', () => {
  const b = body('AV', false); b.paragraphs[0].runs[0].fontFamily = 'Arial'; b.paragraphs[0].runs[0].characterSpacingPt = 3
  const ctx = createCanvas(500, 100).getContext('2d'), calls: Array<{ text: string; x: number }> = []
  ctx.font = '12pt Arial'
  const expectedV = ctx.measureText('AV').width - ctx.measureText('V').width + 4
  ctx.fillText = (text, x) => { calls.push({ text, x }) }
  paintTextBody(b, ctx as any, 0, 0, 500, 100, f => f)
  expect(calls[0]).toEqual({ text: 'A', x: 0 })
  expect(calls[1].x).toBeCloseTo(expectedV)
})


test('rotated and reflected wrapped source copy, partial ranges and hits retain logical order', async () => {
  const source = 'alpha beta gamma delta'
  const transforms = [[1, 0, 0, 1, 0, 0], [-1, 0, 0, -1, 300, 300], [1, 0, 0, -1, 0, 300], [0, -1, 1, 0, 0, 300]]
  let ordinaryLines: string[] | undefined
  for (const [a, b, c, d, e, f] of transforms) {
    const index = await buildTextIndex([{ spec: { widthPx: 500, heightPx: 500 }, paint: ctx => {
      ctx.setTransform(a, b, c, d, e, f)
      paintTextBody(body(source), ctx, 0, 0, 50, 400, family => family)
    } }])
    const lines = index.pages[0].lines
    expect(lines.length).toBeGreaterThan(1)
    const texts = lines.map(line => line.text)
    ordinaryLines ??= texts
    expect(texts).toEqual(ordinaryLines)
    const start = { pageIndex: 0, lineIndex: 0, charIndex: 0 }
    const end = { pageIndex: 0, lineIndex: lines.length - 1, charIndex: lines.at(-1)!.text.length }
    expect(textForRange(index, normalizeRange(end, start))).toBe(source)
    const hitCaret = (lineIndex: number, charIndex: number) => {
      const line = lines[lineIndex]
      let offset = 0
      for (const span of line.spans) {
        if (charIndex >= offset && charIndex <= offset + span.text.length && span.text.length) {
          const p = span.placement!, m = p.transform
          const localX = p.x + p.width * (charIndex - offset) / span.text.length, localY = (p.top + p.bottom) / 2
          const hit = hitTest(index, 0, m.a * localX + m.c * localY + m.e, m.b * localX + m.d * localY + m.f, { strict: true })!
          expect(hit).toEqual({ pageIndex: 0, lineIndex, charIndex })
          return hit
        }
        offset += span.text.length
      }
      throw new Error('No span for caret')
    }
    const partialStart = hitCaret(0, 2), partialEnd = hitCaret(end.lineIndex, end.charIndex - 2)
    const partial = normalizeRange(partialEnd, partialStart)
    expect(partial).toEqual({ start: partialStart, end: partialEnd })
    expect(textForRange(index, partial)).toBe(source.slice(2, -2))
    const slices = selectionSlices(index, partial)
    expect(slices.map(s => lines[s.lineIndex].text.slice(s.from, s.to)).join('')).toBe(source.slice(2, -2))
    const expectedRects = slices.flatMap(s => rectsForRange(lines[s.lineIndex], s.from, s.to))
    expect(rectsForSelectionOnPage(index, 0, partial)).toEqual(expectedRects)
    expect(expectedRects.length).toBeGreaterThan(1)
    expect(expectedRects.every(r => r.width > 0 && r.height > 0)).toBe(true)
    for (const [lineIndex, line] of lines.entries()) {
      const span = line.spans.find(s => s.text.length)!, p = span.placement!, m = p.transform
      const localX = p.x, localY = (p.top + p.bottom) / 2
      expect(hitTest(index, 0, m.a * localX + m.c * localY + m.e, m.b * localX + m.d * localY + m.f, { strict: true })).toEqual({ pageIndex: 0, lineIndex, charIndex: 0 })
    }
    expect(findMatches(index, 'beta gamma')).toHaveLength(1)
  }
})


test('logical paragraph ordering leaves unrelated legacy spatial slots and copy intact', async () => {
  const index = await buildTextIndex([{ spec: { widthPx: 500, heightPx: 500 }, paint: ctx => {
    ctx.font = '12pt serif'
    ctx.fillText('bottom legacy', 0, 450)
    ctx.fillText('top legacy', 0, 20)
    ctx.save()
    ctx.setTransform(-1, 0, 0, -1, 300, 300)
    paintTextBody(body('alpha beta gamma delta'), ctx, 0, 0, 50, 400, family => family)
    ctx.restore()
  } }])
  const lines = index.pages[0].lines
  expect(lines[0].text).toBe('top legacy')
  expect(lines.at(-1)!.text).toBe('bottom legacy')
  expect(lines.filter(l => l.spans[0].logical).map(l => l.text).join('')).toBe('alpha beta gamma delta')
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 1 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 5 } })).toBe('op l')
  expect(findMatches(index, 'legacyalpha')).toHaveLength(0)
})


test('fallback zero-width format tracking preserves source but adds only eligible cluster gaps', () => {
  const source = 'A\u200bBe\u0301👩🏽‍💻\ufeff'
  const b = body(source, false); b.paragraphs[0].runs[0].characterSpacingPt = 6
  const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.font = '12pt "Calibri"'
  const ordinaryWidth = ctx.measureText(source).width
  const spans: Parameters<typeof createRecordingContext>[1] = []
  paintTextBody(b, createRecordingContext(ctx, spans), 0, 0, 500, 100, family => family)
  expect(spans.map(s => s.text).join('')).toBe(source)
  expect(spans[0].width).toBeCloseTo(ordinaryWidth + 4 * 8)
  const draws: string[] = []
  ctx.fillText = text => { draws.push(text) }
  paintTextBody(b, ctx, 0, 0, 500, 100, family => family)
  expect(draws).toEqual(['A', 'B', 'e\u0301', '👩🏽‍💻'])
})


test('direct source CR uses Canvas space preparation for tracking without changing copied source', async () => {
  const source = 'A\rB', b = body(source, false)
  b.paragraphs[0].runs[0].characterSpacingPt = 6
  const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
  ctx.font = '12pt "Calibri"'
  const ordinaryWidth = ctx.measureText('A B').width
  const spans: Parameters<typeof createRecordingContext>[1] = []
  paintTextBody(b, createRecordingContext(ctx, spans), 0, 0, 500, 100, family => family)
  expect(spans[0].width).toBeCloseTo(ordinaryWidth + 3 * 8)
  const draws: string[] = []
  ctx.fillText = text => { draws.push(text) }
  paintTextBody(b, ctx, 0, 0, 500, 100, family => family)
  expect(draws).toEqual(['A', ' ', 'B'])
  const index = await indexFor([shape(b, 500)])
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: source.length } })).toBe(source)
})


describe('drag selection between horizontally separated source scopes', () => {
  test.each([
    { name: 'two boxes with matching bands', cells: false, count: 2, overlap: false, mixed: false },
    { name: 'three boxes with overlapping bands', cells: false, count: 3, overlap: true, mixed: false },
    { name: 'overlapping boxes on a mixed rotated page', cells: false, count: 3, overlap: true, mixed: true },
    { name: 'three same-row table cells', cells: true, count: 3, overlap: false, mixed: false },
    { name: 'same-row cells on a mixed rotated page', cells: true, count: 3, overlap: false, mixed: true },
  ])('$name choose the pointer scope for strict starts and nonstrict drags', async ({ cells, count, overlap, mixed }) => {
    const words = ['lefttext', 'righttext', 'thirdtext'].slice(0, count)
    const shapes: PptxShape[] = cells ? [{ geometry: 'rect', xEmu: 0, yEmu: 0, widthEmu: count * 200 * 9525, heightEmu: 100 * 9525,
      table: { colWidthsEmu: words.map(() => 200 * 9525), rows: [{ cells: words.map(text => ({ paragraphs: body(text, false).paragraphs, gridSpan: 1, rowSpan: 1 })) }] } }]
      : words.map((text, i) => ({ ...shape(body(text, false), 150, overlap ? i * 3 : 0), xEmu: i * 200 * 9525 }))
    if (mixed) shapes.push({ ...shape(body('rotated', false), 100, 500), xEmu: 700 * 9525, rotationDeg: 90 })
    const index = await indexFor(shapes), lines = index.pages[0].lines
    for (const word of words) {
      const lineIndex = lines.findIndex(line => line.text === word), line = lines[lineIndex], span = line.spans[0]
      const y = (line.top + line.bottom) / 2
      const point = (char: number) => span.x + span.width * char / word.length
      const start = hitTest(index, 0, point(0), y, { strict: true })!
      const end = hitTest(index, 0, point(word.length) + 10, y)!
      expect(start).toEqual({ pageIndex: 0, lineIndex, charIndex: 0 })
      expect(end).toEqual({ pageIndex: 0, lineIndex, charIndex: word.length })
      const whole = normalizeRange(start, end)
      expect(textForRange(index, whole)).toBe(word)
      const wholeRects = rectsForSelectionOnPage(index, 0, whole)
      expect(wholeRects).toHaveLength(1)
      expect(wholeRects[0].x).toBeCloseTo(span.x)
      expect(wholeRects[0].y).toBeCloseTo(line.top)
      expect(wholeRects[0].width).toBeCloseTo(span.width)
      expect(wholeRects[0].height).toBeCloseTo(line.bottom - line.top)
      const inside = hitTest(index, 0, point(2), y)!, reverseStart = hitTest(index, 0, point(6), y, { strict: true })!
      expect(inside).toEqual({ pageIndex: 0, lineIndex, charIndex: 2 })
      const reversed = normalizeRange(reverseStart, inside)
      expect(textForRange(index, reversed)).toBe(word.slice(2, 6))
      expect(selectionSlices(index, reversed)).toEqual([{ pageIndex: 0, lineIndex, from: 2, to: 6 }])
      const rects = rectsForSelectionOnPage(index, 0, reversed)
      expect(rects).toHaveLength(1)
      expect(rects[0].x).toBeCloseTo(point(2))
      expect(rects[0].width).toBeCloseTo(span.width * 4 / word.length)
      expect(rects[0].y).toBeCloseTo(line.top)
      expect(rects[0].height).toBeCloseTo(line.bottom - line.top)
    }
  })
})

test('Canvas preparation preserves vertical tab while normalizing only ASCII whitespace', () => {
  const ctx = createCanvas(500, 100).getContext('2d') as unknown as CanvasRenderingContext2D
  const draws: string[] = []
  ctx.fillText = text => { draws.push(text) }
  paintTextBody(body('A\vB', false), ctx, 0, 0, 500, 100, family => family)
  expect(draws).toEqual(['A\vB'])
})


test.each([
  { name: 'reflected', matrix: [-1, 0, 0, -1, 300, 20] },
  { name: 'quarter rotated', matrix: [0, 1, -1, 0, 300, 0] },
])('$name scope wins actual geometry containment over an unrelated ordinary vertical band', async ({ matrix }) => {
  const [a, b, c, d, e, f] = matrix
  const index = await buildTextIndex([{ spec: { widthPx: 500, heightPx: 300 }, paint: ctx => {
    paintTextBody(body('left', false), ctx, 0, 0, 100, 100, family => family)
    ctx.setTransform(a, b, c, d, e, f)
    paintTextBody(body('right', false), ctx, 0, 0, 100, 100, family => family)
  } }])
  const lines = index.pages[0].lines, lineIndex = lines.findIndex(line => line.text === 'right'), line = lines[lineIndex], p = line.spans[0].placement!
  const point = (char: number) => {
    const x = p.x + p.width * char / 5, y = (p.top + p.bottom) / 2
    return [a * x + c * y + e, b * x + d * y + f] as const
  }
  const start = hitTest(index, 0, ...point(0), { strict: true })!
  expect(start).toEqual({ pageIndex: 0, lineIndex, charIndex: 0 })
  for (const char of [2, 5]) {
    const end = hitTest(index, 0, ...point(char))!
    expect(end).toEqual({ pageIndex: 0, lineIndex, charIndex: char })
    expect(textForRange(index, normalizeRange(start, end))).toBe('right'.slice(0, char))
    expect(rectsForSelectionOnPage(index, 0, normalizeRange(start, end))).toEqual(rectsForRange(line, 0, char))
  }
  const reverseStart = hitTest(index, 0, ...point(4), { strict: true })!, reverseEnd = hitTest(index, 0, ...point(1))!
  expect(textForRange(index, normalizeRange(reverseStart, reverseEnd))).toBe('igh')
  expect(selectionSlices(index, normalizeRange(reverseStart, reverseEnd))).toEqual([{ pageIndex: 0, lineIndex, from: 1, to: 4 }])
})
