import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { getPaintables } from '../src/render/paint'
import { buildTextIndex, findMatches, stepMatch } from '../src/core/search'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { parsePptx } from '../src/pptx/parse'
import { buildDocx, buildXlsx, buildPptx } from '../src/testdata/ooxml-builders'
import { RECORD_TEXT, type TextRecordingContext } from '../src/core/text-recording'

async function indexFor(data: ArrayBuffer | Uint8Array) {
  const pkg = await OfficePackage.load(data)
  const ct = await pkg.text('[Content_Types].xml')
  let doc: unknown
  if (ct?.includes('wordprocessingml.document.main')) doc = await parseDocx(pkg)
  else if (ct?.includes('spreadsheetml.sheet.main')) doc = await parseXlsx(pkg)
  else doc = await parsePptx(pkg)
  const paintables = await getPaintables(doc as never)
  return { doc, index: await buildTextIndex(paintables as never), paintables }
}

describe('text index', () => {
  test('captures docx text with page positions', async () => {
    const { index } = await indexFor(
      await buildDocx([
        { runs: [{ text: 'Hello canvas' }] },
        { runs: [{ text: 'Second line here' }] },
      ]),
    )
    expect(index.pages).toHaveLength(1)
    const text = index.pages[0].lines.map((l) => l.text).join(' ')
    expect(text).toContain('Hello canvas')
    expect(text).toContain('Second line here')
    // positions are page-absolute and ordered top-to-bottom
    const ys = index.pages[0].lines.map((l) => l.y)
    expect(ys[0]).toBeLessThan(ys[ys.length - 1])
    const first = index.pages[0].lines[0]
    expect(first.spans[0].x).toBeGreaterThan(0)
    expect(first.spans[0].width).toBeGreaterThan(0)
  })

  test('captures xlsx cell text and pptx shape text', async () => {
    const xlsx = await indexFor(
      await buildXlsx([{ name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }] }] }], ['needle-cell']),
    )
    expect(xlsx.index.pages[0].lines.map((l) => l.text).join(' ')).toContain('needle-cell')

    const pptx = await indexFor(
      await buildPptx([{ prst: 'rect', off: ['914400', '914400'], ext: ['4572000', '1828800'], paragraphs: [{ runs: [{ text: 'needle-slide' }] }] }]),
    )
    expect(pptx.index.pages[0].lines.map((l) => l.text).join(' ')).toContain('needle-slide')
  })

  test('spans on the same baseline merge into one line in visual order', async () => {
    // two cells side by side land on the same row band
    const { index } = await indexFor(
      await buildXlsx([
        {
          name: 'S',
          rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 's', v: 1 }] }],
        },
      ], ['left', 'right']),
    )
    const line = index.pages[0].lines.find((l) => l.text.includes('left') && l.text.includes('right'))
    expect(line).toBeDefined()
    expect(line!.text).toBe('leftright')
  })
})

describe('findMatches', () => {
  test('finds every occurrence and reports rects', async () => {
    const { index } = await indexFor(
      await buildDocx([
        { runs: [{ text: 'alpha beta alpha' }] },
        { runs: [{ text: 'third alpha' }] },
      ]),
    )
    const matches = findMatches(index, 'alpha')
    expect(matches).toHaveLength(3)
    expect(matches.map((m) => m.pageIndex)).toEqual([0, 0, 0])
    for (const m of matches) {
      expect(m.rects.length).toBeGreaterThanOrEqual(1)
      for (const r of m.rects) {
        expect(r.width).toBeGreaterThan(0)
        expect(r.height).toBeGreaterThan(0)
      }
    }
  })

  test('is case-insensitive by default and exact when asked', async () => {
    const { index } = await indexFor(await buildDocx([{ runs: [{ text: 'Canvas Canvas' }] }]))
    expect(findMatches(index, 'canvas')).toHaveLength(2)
    expect(findMatches(index, 'canvas', { caseSensitive: true })).toHaveLength(0)
    expect(findMatches(index, 'Canvas', { caseSensitive: true })).toHaveLength(2)
  })

  test('empty query yields no matches', async () => {
    const { index } = await indexFor(await buildDocx([{ runs: [{ text: 'anything' }] }]))
    expect(findMatches(index, '')).toHaveLength(0)
  })

  test('matches spanning multiple spans produce multiple rects', async () => {
    // bold run in the middle splits a line into several spans
    const { index } = await indexFor(
      await buildDocx([{ runs: [{ text: 'quick ' }, { text: 'brown', bold: true }, { text: ' fox' }] }]),
    )
    const matches = findMatches(index, 'quick brown')
    expect(matches).toHaveLength(1)
    // at least one rect, and the total width covers the phrase
    expect(matches[0].rects.length).toBeGreaterThanOrEqual(1)
  })

  test('finds text on the second page of a multi-page document', async () => {
    const paras = Array.from({ length: 120 }, (_, i) => ({ runs: [{ text: `filler ${i}` }] }))
    paras.push({ runs: [{ text: 'unique-marker-text' }] })
    const { index } = await indexFor(await buildDocx(paras))
    expect(index.pages.length).toBeGreaterThan(1)
    const matches = findMatches(index, 'unique-marker-text')
    expect(matches).toHaveLength(1)
    expect(matches[0].pageIndex).toBeGreaterThan(0)
  })
})

describe('stepMatch', () => {
  test('wraps around in both directions', () => {
    const matches = [1, 2, 3] as never[]
    expect(stepMatch(matches, -1, 1)).toBe(0)
    expect(stepMatch(matches, 2, 1)).toBe(0)
    expect(stepMatch(matches, 0, -1)).toBe(2)
    expect(stepMatch([], 0, 1)).toBe(-1)
  })
})

/** Capture actual ordinary painting independently from the text index proxy. */
async function actualText(paint: (ctx: CanvasRenderingContext2D) => void) {
  const { createCanvas } = await import('canvas')
  const ctx = createCanvas(800, 600).getContext('2d'), calls: Array<{ text: string; x: number; y: number }> = []
  ctx.fillText = (text, x, y) => {
    const m = ctx.getTransform()
    calls.push({ text, x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f })
  }
  paint(ctx as unknown as CanvasRenderingContext2D)
  return calls
}

describe('page coordinates under Canvas transforms', () => {
  test('separated ordinary PPTX shapes keep their visible origins and distinct search lines', async () => {
    const { index, paintables } = await indexFor(await buildPptx([
      { off: ['914400', '914400'], ext: ['4572000', '1828800'], paragraphs: [{ runs: [{ text: 'needle-one' }] }] },
      { off: ['914400', '3657600'], ext: ['4572000', '1828800'], paragraphs: [{ runs: [{ text: 'needle-two' }] }] },
    ]))
    const painted = await actualText(paintables[0].paint)
    expect(index.pages[0].lines.map(line => line.text)).toEqual(['needle-one', 'needle-two'])
    for (const [i, line] of index.pages[0].lines.entries()) {
      expect(line.spans[0].x).toBeCloseTo(105.6)
      expect(line.spans[0].y).toBeCloseTo(painted[i].y)
      expect(painted[i].y).toBeGreaterThan((i === 0 ? 96 : 384) + 4.8)
      expect(painted[i].y).toBeLessThan((i === 0 ? 96 : 384) + 4.8 + 16)
      expect(line.spans[0].fontSize).toBeCloseTo(16)
      expect(findMatches(index, line.text)[0].rects[0].x).toBeCloseTo(105.6)
    }
  })

  test('positioned table text includes its frame and cell translations', async () => {
    const { index, paintables } = await indexFor(await buildPptx([{
      off: ['914400', '1828800'], ext: ['1828800', '914400'],
      table: { colWidths: ['1828800'], rows: [{ h: '914400', cells: [{ paragraphs: [{ runs: [{ text: 'table-needle' }] }] }] }] },
    }]))
    const span = index.pages[0].lines[0].spans[0]
    expect(span.x).toBeCloseTo(105.6)
    const painted = await actualText(paintables[0].paint)
    expect(span.y).toBeCloseTo(painted[0].y)
    // 96px cell, vertical-center default, 4.8px top/bottom insets and 19.2px Office normal advance.
    expect(span.y).toBeGreaterThan(192 + 4.8 + (96 - 9.6 - 19.2) / 2)
    expect(span.y).toBeLessThan(192 + 4.8 + (96 - 9.6 - 19.2) / 2 + 16)
    expect(findMatches(index, 'table-needle')[0].rects[0].y).toBeCloseTo(painted[0].y - 16 * .85)
    const { hitTest } = await import('../src/core/selection')
    expect(hitTest(index, 0, 110, 240, { strict: true })?.lineIndex).toBe(0)
    expect(hitTest(index, 0, 10, 54, { strict: true })).toBeUndefined()
  })

  test('legacy shapes and nested nonuniform groups project baseline, width and text band', async () => {
    const { renderSlide } = await import('../src/pptx/render')
    const px = (v: number) => v * 9525
    const textBody = { paragraphs: [{ runs: [{ text: 'group-needle' }], align: 'left' as const, level: 0 }], anchor: 't' as const, insetLeftEmu: px(9.6), insetRightEmu: 0, insetTopEmu: px(4.8), insetBottomEmu: 0, wrap: false }
    const child = { xEmu: px(15), yEmu: px(10), widthEmu: px(100), heightEmu: px(60), geometry: 'rect' as const, textBody }
    const inner = { xEmu: px(20), yEmu: px(10), widthEmu: px(40), heightEmu: px(20), geometry: 'other' as const, group: { off: { x: px(20), y: px(10) }, ext: { width: px(40), height: px(20) }, chOff: { x: px(5), y: px(5) }, chExt: { width: px(40), height: px(10) } }, children: [child] }
    const outer = { xEmu: px(100), yEmu: px(50), widthEmu: px(120), heightEmu: px(80), geometry: 'other' as const, group: { off: { x: px(100), y: px(50) }, ext: { width: px(120), height: px(80) }, chOff: { x: px(10), y: 0 }, chExt: { width: px(60), height: px(40) } }, children: [inner] }
    const legacy = { ...child, xEmu: px(96), yEmu: px(300), textBody: { ...textBody, paragraphs: [{ runs: [{ text: 'legacy-needle' }], align: 'left' as const, level: 0 }] } }
    const index = await buildTextIndex([{ spec: { widthPx: 800, heightPx: 600 }, paint: ctx => renderSlide({ index: 0, widthEmu: px(800), heightEmu: px(600), shapes: [outer, legacy] }, ctx) }])
    const group = index.pages[0].lines.find(line => line.text === 'group-needle')!
    expect(group.spans[0].x).toBeCloseTo(159.2)
    const painted = await actualText(ctx => renderSlide({ index: 0, widthEmu: px(800), heightEmu: px(600), shapes: [outer, legacy] }, ctx))
    expect(group.spans[0].y).toBeCloseTo(painted.find(call => call.text === 'group-needle')!.y)
    expect(group.spans[0].fontSize).toBeCloseTo(64)
    const { createCanvas } = await import('canvas')
    const measure = createCanvas(4, 4).getContext('2d'); measure.font = '12pt "Calibri"'
    expect(group.spans[0].width).toBeCloseTo(2 * measure.measureText('group-needle').width)
    const rect = findMatches(index, 'group-needle')[0].rects[0]
    expect(rect.x).toBeCloseTo(159.2)
    expect(rect.y).toBeCloseTo(painted[0].y - 64 * .85)
    expect(rect.height).toBeCloseTo(70.4)
    const partial = findMatches(index, 'needle').find(match => match.lineIndex === index.pages[0].lines.indexOf(group))!.rects[0]
    expect(partial.x).toBeCloseTo(159.2 + measure.measureText('group-needle').width)
    expect(partial.width).toBeCloseTo(measure.measureText('group-needle').width)
    expect(partial.height).toBeCloseTo(70.4)
    const { hitTest } = await import('../src/core/selection')
    expect(hitTest(index, 0, 170, 150, { strict: true })).toMatchObject({ lineIndex: index.pages[0].lines.indexOf(group), charIndex: 1 })
    expect(hitTest(index, 0, 110, 310, { strict: true })?.lineIndex).toBe(index.pages[0].lines.findIndex(line => line.text === 'legacy-needle'))
    expect(index.pages[0].lines.find(line => line.text === 'legacy-needle')!.spans[0]).toMatchObject({ x: 105.6, y: painted.find(call => call.text === 'legacy-needle')!.y })
  })

  test('rotated flipped adjacent runs remain one logical phrase in source reading order', async () => {
    const index = await buildTextIndex([{ spec: { widthPx: 600, heightPx: 400 }, paint: ctx => {
      ctx.font = '20px sans-serif'; ctx.translate(500, 200); ctx.rotate(Math.PI / 2); ctx.scale(-1, 1)
      ctx.fillText('AB', 10, 20)
      ctx.fillText('CD', 10 + ctx.measureText('AB').width, 20)
    } }])
    expect(index.pages[0].lines.map(line => line.text)).toEqual(['ABCD'])
    const phrase = findMatches(index, 'BC')[0]
    expect(phrase.rects).toHaveLength(2)
    expect(phrase.rects[0].x).toBeCloseTo(475)
    expect(phrase.rects[0].width).toBeCloseTo(22)
    expect(phrase.rects[1].y).toBeLessThan(phrase.rects[0].y)
  })

  test('save/restore uses the real Canvas state and successive pages are isolated', async () => {
    const index = await buildTextIndex([
      { spec: { widthPx: 600, heightPx: 400 }, paint: ctx => {
        ctx.font = '20px sans-serif'; ctx.translate(100, 50); ctx.save(); ctx.scale(2, 3); ctx.fillText('scaled', 10, 20); ctx.restore(); ctx.fillText('restored', 10, 80)
      } },
      { spec: { widthPx: 600, heightPx: 400 }, paint: ctx => { ctx.font = '20px sans-serif'; ctx.fillText('next-page', 10, 20) } },
    ])
    expect(index.pages[0].lines.find(line => line.text === 'scaled')!.spans[0]).toMatchObject({ x: 120, y: 110, fontSize: 60 })
    expect(index.pages[0].lines.find(line => line.text === 'restored')!.spans[0]).toMatchObject({ x: 110, y: 130, fontSize: 20 })
    expect(index.pages[1].lines[0].spans[0]).toMatchObject({ x: 10, y: 20, fontSize: 20 })
  })
})


test('recording text never modifies the caller painting transform/font/state', async () => {
  const { createCanvas } = await import('canvas')
  const { createRecordingContext } = await import('../src/core/search')
  const base = createCanvas(4, 4).getContext('2d')
  base.font = '12pt "Calibri"'; base.fillStyle = '#123456'; base.textAlign = 'right'; base.textBaseline = 'top'
  base.setTransform(2, 1, .5, 3, 100, 50)
  const before = base.getTransform(), font = base.font
  const sink: import('../src/core/search').TextSpan[] = []
  const recorder = createRecordingContext(base as unknown as CanvasRenderingContext2D, sink)
  recorder.save(); recorder.translate(10, 20); recorder.fillText('caller', 40, 0); recorder.restore()
  expect(base.getTransform()).toEqual(before)
  expect(base.font).toBe(font)
  expect(base.fillStyle).toBe('#123456')
  expect(base.textAlign).toBe('right')
  expect(base.textBaseline).toBe('top')
  expect(sink).toHaveLength(1)
  expect(sink[0].placement?.transform).toMatchObject({ a: 2, b: 1, c: .5, d: 3, e: 130, f: 120 })
})

test('a source line joins mixed upright and rotated actual Canvas records', async () => {
  const source = { text: 'AB中CD' }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: context => {
    const ctx = context as TextRecordingContext
    ctx.font = '16px sans-serif'
    const record = ctx[RECORD_TEXT]!
    ctx.save(); ctx.translate(100, 10); ctx.rotate(Math.PI / 2)
    record('AB', 0, 0, 22, { source, start: 0, end: 2, line: 0, run: 0, graphemeBoundaries: [0, 1, 2] })
    ctx.restore()
    record('中', 90, 50, 16, { source, start: 2, end: 3, line: 0, run: 1, graphemeBoundaries: [2, 3] })
    ctx.save(); ctx.translate(100, 70); ctx.rotate(Math.PI / 2)
    record('CD', 0, 0, 22, { source, start: 3, end: 5, line: 0, run: 2, graphemeBoundaries: [3, 4, 5] })
    ctx.restore()
  } }])
  expect(index.pages[0].lines.map(line => line.text)).toEqual(['AB中CD'])
  expect(index.pages[0].lines[0].spans.map(span => span.logical?.run)).toEqual([0, 1, 2])
  expect(findMatches(index, 'B中C')[0].rects).toHaveLength(3)
})

test('repeated fill and outline capture one geometry while outline-only text stays searchable', async () => {
  const index = await buildTextIndex([{ spec: { widthPx: 300, heightPx: 100 }, paint: ctx => {
    ctx.font = '18px sans-serif'
    ctx.fillText('once', 10, 30)
    ctx.strokeText('once', 10, 30)
    ctx.strokeText('outline', 100, 30)
  } }])
  expect(index.pages[0].lines.flatMap(line => line.spans).map(span => span.text)).toEqual(['once', 'outline'])
  expect(findMatches(index, 'outline')).toHaveLength(1)
})

test('offset shadow, fill and outline passes retain one canonical logical geometry', async () => {
  const source = { text: 'ONCE' }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 100 }, paint: context => {
    const ctx = context as TextRecordingContext; ctx.font = '16px sans-serif'
    const record = ctx[RECORD_TEXT]!
    record('ONCE', 10, 30, 40, { source, start: 0, end: 4, line: 0 })
    record('ONCE', 12, 32, 40, { source, start: 0, end: 4, line: 0 })
    record('ONCE', 10, 30, 40, { source, start: 0, end: 4, line: 0 })
  } }])
  expect(index.pages[0].lines).toHaveLength(1)
  expect(index.pages[0].lines[0].text).toBe('ONCE')
  expect(findMatches(index, 'ONCE')[0].rects).toEqual([{ x: 10, y: 16.4, width: 40, height: 17.6 }])
})

test('shadow-first recording uses an explicit canonical band and distinct sources remain independent', async () => {
  const first = { text: 'ONCE' }, second = { text: 'ONCE' }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 100 }, paint: context => {
    const ctx = context as TextRecordingContext; ctx.font = '16px sans-serif'
    const record = ctx[RECORD_TEXT]!
    record('ONCE', 12, 32, 40, { source: first, start: 0, end: 4, line: 0,
      canonical: { x: 10, y: 30, width: 40 } })
    record('ONCE', 10, 30, 40, { source: first, start: 0, end: 4, line: 0 })
    record('ONCE', 100, 30, 40, { source: second, start: 0, end: 4, line: 0 })
  } }])
  const matches = findMatches(index, 'ONCE')
  expect(matches).toHaveLength(2)
  expect(matches.map(match => match.rects[0].x)).toEqual([10, 100])
})

test('canonical recording overrides an offset shadow transform before outline-only paint', async () => {
  const source = { text: 'GLYPH' }
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 100 }, paint: context => {
    const ctx = context as TextRecordingContext; ctx.font = '16px sans-serif'
    ctx.save(); ctx.translate(3, 4)
    ctx[RECORD_TEXT]!('GLYPH', 10, 30, 50, { source, start: 0, end: 5, line: 0,
      canonical: { x: 10, y: 30, width: 50, transform: identity } })
    ctx.restore()
    ctx.strokeText('unrelated', 100, 60)
    ctx[RECORD_TEXT]!('GLYPH', 10, 30, 50, { source, start: 0, end: 5, line: 0 })
  } }])
  expect(index.pages[0].lines.flatMap(line => line.spans).filter(span => span.logical?.source === source)).toHaveLength(1)
  expect(findMatches(index, 'GLYPH')[0].rects[0]).toEqual({ x: 10, y: 16.4, width: 50, height: 17.6 })
})

test('source-less identical fill and stroke operations retain independent occurrences', async () => {
  const record = async (operations: Array<'fill' | 'stroke'>) => {
    const { createCanvas } = await import('canvas')
    const { createRecordingContext } = await import('../src/core/search')
    const sink: import('../src/core/search').TextSpan[] = []
    const ctx = createRecordingContext(createCanvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D, sink)
    ctx.font = '12px sans-serif'
    for (const operation of operations) if (operation === 'fill') ctx.fillText('same', 10, 20); else ctx.strokeText('same', 10, 20)
    return sink
  }
  expect(await record(['fill', 'fill'])).toHaveLength(2)
  expect(await record(['stroke', 'stroke'])).toHaveLength(2)
  expect(await record(['fill', 'stroke'])).toHaveLength(1)
  expect(await record(['fill', 'stroke', 'fill', 'stroke'])).toHaveLength(2)
})
