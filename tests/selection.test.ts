import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { getPaintables } from '../src/render/paint'
import { buildTextIndex } from '../src/core/search'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import {
  hitTest,
  isEmptyRange,
  lineRangeAt,
  normalizeRange,
  rectsForSelectionOnPage,
  selectionSlices,
  textForRange,
  wordRangeAt,
  type CaretPos,
} from '../src/core/selection'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'
import { RECORD_TEXT, type TextRecordingContext } from '../src/core/text-recording'

test('optional cell clip retains full source while bounding rectangles and strict/non-strict carets', async () => {
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => {
    ctx.font = '20px sans-serif'; ctx.save(); ctx.transform(2, .3, .5, 1, 30, 40)
    const m = ctx.getTransform()
    const clip = { x: 0, y: 0, width: 30, height: 40, transform: { a: m.a, b: m.b, c: m.c, d: m.d, e: m.e, f: m.f } }
    const source = { text: 'ABCDE' }
    ;(ctx as TextRecordingContext)[RECORD_TEXT]!('ABCDE', 0, 20, 100, { source, start: 0, end: 5, clip })
    ctx.restore()
    expect([ctx.getTransform().e, ctx.getTransform().f]).toEqual([0, 0])
  } }])
  const line = index.pages[0].lines[0]
  const range = { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 5 } }
  expect(textForRange(index, range)).toBe('ABCDE')
  const rect = rectsForSelectionOnPage(index, 0, range)[0]
  expect(rect.x + rect.width).toBeLessThanOrEqual(110)
  expect(rect.y + rect.height).toBeLessThanOrEqual(89)
  // Local (15,20) is visible; local (80,20) is clipped past the cell edge.
  for (const strict of [true, false]) {
    expect(hitTest(index, 0, 70, 64.5, { strict })).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 1 })
    expect(hitTest(index, 0, 200, 84, { strict })).toBeUndefined()
  }
  expect(line.text).toBe('ABCDE')
})

async function indexDocx(paras: Parameters<typeof buildDocx>[0]) {
  const doc = await parseDocx(await OfficePackage.load(await buildDocx(paras)))
  const paintables = await getPaintables(doc as never)
  return buildTextIndex(paintables as never)
}

async function indexXlsx() {
  const doc = await parseXlsx(await OfficePackage.load(await buildXlsx([
    { name: 'S', rows: [{ r: 1, cells: [{ ref: 'A1', t: 's', v: 0 }, { ref: 'B1', t: 's', v: 1 }] }] },
  ], ['hello', 'world'])))
  const paintables = await getPaintables(doc as never)
  return buildTextIndex(paintables as never)
}

describe('hit testing', () => {
  test('maps a point to the nearest caret in docx text', async () => {
    const index = await indexDocx([{ runs: [{ text: 'hello world' }] }])
    const line = index.pages[0].lines[0]
    // a point in the middle of the line's vertical band
    const y = (line.top + line.bottom) / 2
    const caret = hitTest(index, 0, line.spans[0].x + line.spans[0].width * 0.5, y)
    expect(caret).toBeDefined()
    expect(caret!.lineIndex).toBe(0)
    expect(caret!.charIndex).toBeGreaterThan(0)
    expect(caret!.charIndex).toBeLessThan(line.text.length)
  })

  test('clamps past the end of a line to the line end', async () => {
    const index = await indexDocx([{ runs: [{ text: 'short' }] }])
    const line = index.pages[0].lines[0]
    const caret = hitTest(index, 0, line.spans[0].x + 5000, (line.top + line.bottom) / 2)
    expect(caret!.charIndex).toBe(line.text.length)
  })

  test('a point above the first line snaps to the first line', async () => {
    const index = await indexDocx([{ runs: [{ text: 'line one' }] }, { runs: [{ text: 'line two' }] }])
    expect(index.pages[0].lines.length).toBe(2)
    const caret = hitTest(index, 0, 10, -500)
    expect(caret!.lineIndex).toBe(0)
  })

  test('hit tests spreadsheet cells', async () => {
    const index = await indexXlsx()
    const line = index.pages[0].lines.find((l) => l.text.includes('hello'))
    expect(line).toBeDefined()
    const caret = hitTest(index, 0, line!.spans[0].x + 2, (line!.top + line!.bottom) / 2)
    // lands on the row-0 line, at or near the start of "hello"
    expect(caret!.lineIndex).toBeLessThanOrEqual(1)
    expect(index.pages[0].lines[caret!.lineIndex].text).toContain('hello')
  })

  test('returns undefined for an unknown page', async () => {
    const index = await indexDocx([{ runs: [{ text: 'x' }] }])
    expect(hitTest(index, 99, 10, 10)).toBeUndefined()
  })
})

describe('range normalization', () => {
  test('orders backwards selections', () => {
    const a: CaretPos = { pageIndex: 0, lineIndex: 2, charIndex: 5 }
    const b: CaretPos = { pageIndex: 0, lineIndex: 1, charIndex: 2 }
    const range = normalizeRange(a, b)
    expect(range.start).toEqual({ pageIndex: 0, lineIndex: 1, charIndex: 2 })
    expect(range.end).toEqual({ pageIndex: 0, lineIndex: 2, charIndex: 5 })
    expect(isEmptyRange(range)).toBe(false)
  })

  test('a caret on itself is empty', () => {
    const a: CaretPos = { pageIndex: 0, lineIndex: 0, charIndex: 3 }
    expect(isEmptyRange(normalizeRange(a, a))).toBe(true)
  })
})

describe('word and line expansion', () => {
  test('double-click expands to the word', async () => {
    const index = await indexDocx([{ runs: [{ text: 'alpha beta gamma' }] }])
    const line = index.pages[0].lines[0]
    expect(wordRangeAt(line, 7)).toEqual({ start: 6, end: 10 }) // "beta"
    expect(wordRangeAt(line, 0)).toEqual({ start: 0, end: 5 }) // "alpha"
    expect(wordRangeAt(line, line.text.length - 1)).toEqual({ start: 11, end: 16 }) // "gamma"
  })

  test('expands across run boundaries (bold word)', async () => {
    const index = await indexDocx([{ runs: [{ text: 'quick ' }, { text: 'brown', bold: true }] }])
    const line = index.pages[0].lines[0]
    const word = wordRangeAt(line, 8) // inside "brown"
    expect(line.text.slice(word.start, word.end)).toBe('brown')
  })

  test('triple-click selects the whole line', async () => {
    const index = await indexDocx([{ runs: [{ text: 'whole line here' }] }])
    const line = index.pages[0].lines[0]
    expect(lineRangeAt(line, 4)).toEqual({ start: 0, end: line.text.length })
  })
})

describe('selection text and rects', () => {
  test('copies the selected substring', async () => {
    const index = await indexDocx([{ runs: [{ text: 'hello world again' }] }])
    const text = textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 6 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 11 } })
    expect(text).toBe('world')
  })

  test('multi-line selection joins with newlines and trims trailing spaces', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'first line' }] },
      { runs: [{ text: 'second line' }] },
      { runs: [{ text: 'third line' }] },
    ])
    expect(index.pages[0].lines.length).toBe(3)
    const text = textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 2, charIndex: 10 } })
    expect(text.split('\n')).toHaveLength(3)
    expect(text.startsWith('first line')).toBe(true)
    expect(text.endsWith('third line')).toBe(true)
  })

  test('partial first/last lines are respected', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'aaaa bbbb cccc' }] },
      { runs: [{ text: 'dddd eeee ffff' }] },
    ])
    const text = textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 5 }, end: { pageIndex: 0, lineIndex: 1, charIndex: 4 } })
    expect(text).toBe('bbbb cccc\ndddd')
  })

  test('produces one rect per touched line', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'aaaa bbbb cccc' }] },
      { runs: [{ text: 'dddd eeee ffff' }] },
    ])
    const rects = rectsForSelectionOnPage(index, 0, { start: { pageIndex: 0, lineIndex: 0, charIndex: 5 }, end: { pageIndex: 0, lineIndex: 1, charIndex: 4 } })
    expect(rects.length).toBeGreaterThanOrEqual(2)
    // rects are top-to-bottom
    const ys = rects.map((r) => r.y).sort((a, b) => a - b)
    expect(ys[0]).toBeLessThanOrEqual(ys[ys.length - 1])
    for (const r of rects) {
      expect(r.width).toBeGreaterThan(0)
      expect(r.height).toBeGreaterThan(0)
    }
  })

  test('a whole-line selection covers the full line width', async () => {
    const index = await indexDocx([{ runs: [{ text: 'measure me' }] }])
    const line = index.pages[0].lines[0]
    const rects = rectsForSelectionOnPage(index, 0, {
      start: { pageIndex: 0, lineIndex: 0, charIndex: 0 },
      end: { pageIndex: 0, lineIndex: 0, charIndex: line.text.length },
    })
    const covered = rects.reduce((a, r) => a + r.width, 0)
    const total = line.spans.reduce((a, s) => a + s.width, 0)
    expect(covered).toBeCloseTo(total, 0)
  })
})

describe('selection across pages', () => {
  /**
   * A hand-built index so the geometry assertions are exact and do not depend
   * on how a real document happens to paginate.
   * Page 0 and page 1 each have two lines 20px apart.
   */
  function fakeIndex(): import('../src/core/search').TextIndex {
    const CHAR_W = 10
    const line = (text: string, y: number) => ({
      spans: [{ text, x: 50, y, width: text.length * CHAR_W, fontSize: 12 }],
      text,
      y,
      top: y - 12,
      bottom: y + 3,
    })
    return {
      pages: [
        { index: 0, lines: [line('alpha one', 100), line('bravo two', 120)] },
        { index: 1, lines: [line('delta four', 100), line('echo five', 120)] },
      ],
    }
  }

  const index = fakeIndex()

  test('normalizeRange orders a backwards cross-page drag', () => {
    const range = normalizeRange(
      { pageIndex: 1, lineIndex: 0, charIndex: 8 },
      { pageIndex: 0, lineIndex: 1, charIndex: 3 },
    )
    expect(range.start).toEqual({ pageIndex: 0, lineIndex: 1, charIndex: 3 })
    expect(range.end).toEqual({ pageIndex: 1, lineIndex: 0, charIndex: 8 })
    expect(isEmptyRange(range)).toBe(false)
  })

  test('slices cover only the touched lines, in page then line order', () => {
    const slices = selectionSlices(
      index,
      normalizeRange(
        { pageIndex: 0, lineIndex: 0, charIndex: 6 },
        { pageIndex: 1, lineIndex: 1, charIndex: 4 },
      ),
    )
    expect(slices).toEqual([
      { pageIndex: 0, lineIndex: 0, from: 6, to: 9 },
      { pageIndex: 0, lineIndex: 1, from: 0, to: 9 },
      { pageIndex: 1, lineIndex: 0, from: 0, to: 10 },
      { pageIndex: 1, lineIndex: 1, from: 0, to: 4 },
    ])
  })

  test('each page paints only its own rectangles, in its own coordinates', () => {
    const range = normalizeRange(
      { pageIndex: 0, lineIndex: 1, charIndex: 0 },
      { pageIndex: 1, lineIndex: 0, charIndex: 5 },
    )
    const onPage0 = rectsForSelectionOnPage(index, 0, range)
    const onPage1 = rectsForSelectionOnPage(index, 1, range)
    // page 0 contributes only its second line (y around 120)
    expect(onPage0).toHaveLength(1)
    expect(onPage0[0].y).toBeCloseTo(108, 0)
    // page 1 contributes only its first line (y around 100)
    expect(onPage1).toHaveLength(1)
    expect(onPage1[0].y).toBeCloseTo(88, 0)
    // a page outside the selection paints nothing
    expect(rectsForSelectionOnPage(index, 2, range)).toHaveLength(0)
  })

  test('copied text spans the page break, one line per visual line', () => {
    const text = textForRange(
      index,
      normalizeRange(
        { pageIndex: 0, lineIndex: 0, charIndex: 6 },
        { pageIndex: 1, lineIndex: 1, charIndex: 4 },
      ),
    )
    expect(text).toBe('one\nbravo two\ndelta four\necho')
  })

  test('a drag from a page end to the next page start paints no empty band', () => {
    const slices = selectionSlices(
      index,
      normalizeRange(
        { pageIndex: 0, lineIndex: 1, charIndex: 9 },
        { pageIndex: 1, lineIndex: 0, charIndex: 0 },
      ),
    )
    // both touch points are zero-length, so nothing is selected or painted
    expect(slices).toEqual([])
    expect(rectsForSelectionOnPage(index, 0, normalizeRange(
      { pageIndex: 0, lineIndex: 1, charIndex: 9 },
      { pageIndex: 1, lineIndex: 0, charIndex: 0 },
    ))).toHaveLength(0)
  })

  test('a selection inside one page is unaffected', () => {
    const range = normalizeRange(
      { pageIndex: 0, lineIndex: 0, charIndex: 0 },
      { pageIndex: 0, lineIndex: 0, charIndex: 4 },
    )
    expect(textForRange(index, range)).toBe('alph')
    expect(rectsForSelectionOnPage(index, 1, range)).toHaveLength(0)
  })

  test('a range spanning pages yields text from both pages', () => {
    const range = normalizeRange(
      { pageIndex: 0, lineIndex: 1, charIndex: 6 },
      { pageIndex: 1, lineIndex: 0, charIndex: 5 },
    )
    const text = textForRange(index, range)
    expect(text.startsWith('two')).toBe(true)
    expect(text.endsWith('delta')).toBe(true)
    expect(text.split('\n')).toHaveLength(2)
  })
})

describe('textForRange robustness', () => {
  test('two different documents with identical ranges do not interfere', () => {
    // guards the old module-level slice memo, which keyed only on range
    // coordinates and could return a stale slice list for a second document
    const mk = (lines: string[]) => ({
      pages: lines.map((text, i) => ({
        index: 0,
        lines: [{
          spans: [{ text, x: 0, y: 100 + i * 20, width: text.length * 10, fontSize: 12 }],
          text,
          y: 100 + i * 20,
          top: 88 + i * 20,
          bottom: 103 + i * 20,
        }],
      })),
    })
    const a = mk(['first doc line']) as never
    const b = mk(['second doc!!']) as never
    const range = normalizeRange({ pageIndex: 0, lineIndex: 0, charIndex: 0 }, { pageIndex: 0, lineIndex: 0, charIndex: 5 })
    expect(textForRange(a, range)).toBe('first')
    expect(textForRange(b, range)).toBe('secon')
  })

  test('trailing whitespace is trimmed on all but the last line', () => {
    const index = {
      pages: [
        { index: 0, lines: [
          { spans: [{ text: 'aaa   ', x: 0, y: 100, width: 60, fontSize: 12 }], text: 'aaa   ', y: 100, top: 88, bottom: 103 },
          { spans: [{ text: 'bbb   ', x: 0, y: 120, width: 60, fontSize: 12 }], text: 'bbb   ', y: 120, top: 108, bottom: 123 },
        ] },
      ],
    } as never
    const text = textForRange(index, normalizeRange(
      { pageIndex: 0, lineIndex: 0, charIndex: 0 },
      { pageIndex: 0, lineIndex: 1, charIndex: 6 },
    ))
    expect(text).toBe('aaa\nbbb   ')
  })

  test('a collapsed caret range selects no text', () => {
    const index = {
      pages: [{ index: 0, lines: [{ spans: [{ text: 'x', x: 0, y: 100, width: 10, fontSize: 12 }], text: 'x', y: 100, top: 88, bottom: 103 }] }],
    } as never
    // start == end is a caret, not a selection: nothing to copy
    const collapsed = normalizeRange({ pageIndex: 0, lineIndex: 0, charIndex: 1 }, { pageIndex: 0, lineIndex: 0, charIndex: 1 })
    expect(textForRange(index, collapsed)).toBe('')
    // a reversed drag still normalizes to a forward one-character selection
    expect(textForRange(index, normalizeRange(
      { pageIndex: 0, lineIndex: 0, charIndex: 1 },
      { pageIndex: 0, lineIndex: 0, charIndex: 0 },
    ))).toBe('x')
  })
})

describe('selection across three pages', () => {
  /** three pages, two lines each — the middle page is the one that matters */
  function threePageIndex() {
    const line = (text: string, y: number) => ({
      spans: [{ text, x: 50, y, width: text.length * 10, fontSize: 12 }],
      text,
      y,
      top: y - 12,
      bottom: y + 3,
    })
    return {
      pages: [
        { index: 0, lines: [line('p0a', 100), line('p0b', 120)] },
        { index: 1, lines: [line('p1a', 100), line('p1b', 120)] },
        { index: 2, lines: [line('p2a', 100), line('p2b', 120)] },
      ],
    } as never
  }

  const index = threePageIndex()

  test('a middle page contributes full lines, not the boundary char bounds', () => {
    // start on page 0 line 1 at char 2, end on page 2 line 0 at char 2
    const slices = selectionSlices(
      index,
      normalizeRange(
        { pageIndex: 0, lineIndex: 1, charIndex: 2 },
        { pageIndex: 2, lineIndex: 0, charIndex: 2 },
      ),
    )
    // the middle page must take BOTH lines in full — a page-blind
    // implementation would clip it to char 2 on every line
    expect(slices).toEqual([
      { pageIndex: 0, lineIndex: 1, from: 2, to: 3 },
      { pageIndex: 1, lineIndex: 0, from: 0, to: 3 },
      { pageIndex: 1, lineIndex: 1, from: 0, to: 3 },
      { pageIndex: 2, lineIndex: 0, from: 0, to: 2 },
    ])
  })

  test('the middle page paints its own coordinates only', () => {
    const range = normalizeRange(
      { pageIndex: 0, lineIndex: 1, charIndex: 0 },
      { pageIndex: 2, lineIndex: 0, charIndex: 3 },
    )
    const middle = rectsForSelectionOnPage(index, 1, range)
    expect(middle).toHaveLength(2)
    expect(middle[0].y).toBeCloseTo(88, 0)
    expect(middle[1].y).toBeCloseTo(108, 0)
    // the first page contributes only its last line, the last only its first
    expect(rectsForSelectionOnPage(index, 0, range)).toHaveLength(1)
    expect(rectsForSelectionOnPage(index, 2, range)).toHaveLength(1)
  })

  test('text spans all three pages in order', () => {
    const text = textForRange(
      index,
      normalizeRange(
        { pageIndex: 0, lineIndex: 0, charIndex: 0 },
        { pageIndex: 2, lineIndex: 1, charIndex: 3 },
      ),
    )
    expect(text).toBe('p0a\np0b\np1a\np1b\np2a\np2b')
  })

  test('a denormalized range still yields the forward selection', () => {
    const forward = selectionSlices(
      index,
      normalizeRange({ pageIndex: 1, lineIndex: 0, charIndex: 0 }, { pageIndex: 1, lineIndex: 1, charIndex: 2 }),
    )
    const reversed = selectionSlices(index, {
      start: { pageIndex: 1, lineIndex: 1, charIndex: 2 },
      end: { pageIndex: 1, lineIndex: 0, charIndex: 0 },
    })
    expect(reversed).toEqual(forward)
  })
})

describe('selection of transformed text', () => {
  async function affineIndex() {
    return buildTextIndex([{ spec: { widthPx: 500, heightPx: 300 }, paint: ctx => {
      ctx.font = '20px sans-serif'; ctx.setTransform(2, 1, .5, 3, 100, 50); ctx.fillText('ABCD', 10, 20)
    } }])
  }

  test('whole and partial highlights bound all four transformed text-band corners', async () => {
    const index = await affineIndex()
    const { createCanvas } = await import('canvas')
    const c = createCanvas(4, 4).getContext('2d'); c.font = '20px sans-serif'
    const width = c.measureText('ABCD').width
    const whole = rectsForSelectionOnPage(index, 0, { start: { pageIndex: 0, lineIndex: 0, charIndex: 0 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 4 } })[0]
    expect(whole.x).toBeCloseTo(121.5)
    expect(whole.y).toBeCloseTo(69)
    expect(whole.width).toBeCloseTo(2 * width + 11)
    expect(whole.height).toBeCloseTo(width + 66)
    const partial = rectsForSelectionOnPage(index, 0, { start: { pageIndex: 0, lineIndex: 0, charIndex: 1 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 3 } })[0]
    expect(partial.x).toBeCloseTo(121.5 + width / 2)
    expect(partial.y).toBeCloseTo(69 + width / 4)
    expect(partial.width).toBeCloseTo(width + 11)
    expect(partial.height).toBeCloseTo(width / 2 + 66)
    expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 1 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 3 } })).toBe('BC')
  })

  test('strict hit testing inverse-maps the visible band rather than accepting its bounding box', async () => {
    const index = await affineIndex()
    const { createCanvas } = await import('canvas')
    const c = createCanvas(4, 4).getContext('2d'); c.font = '20px sans-serif'
    const width = c.measureText('ABCD').width
    expect(hitTest(index, 0, 127 + 1.5 * width, 102 + .75 * width, { strict: true })).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 3 })
    // This point is inside the axis-aligned highlight but outside the sheared text band.
    expect(hitTest(index, 0, 122.5, 134 + width, { strict: true })).toBeUndefined()
  })

  test('rotation and horizontal/vertical flips preserve character progression across runs', async () => {
    const index = await buildTextIndex([{ spec: { widthPx: 600, heightPx: 400 }, paint: ctx => {
      ctx.font = '20px sans-serif'; ctx.translate(500, 200); ctx.rotate(Math.PI / 2); ctx.scale(-1, -1)
      ctx.fillText('AB', 10, 20); ctx.fillText('CD', 10 + ctx.measureText('AB').width, 20)
    } }])
    const { createCanvas } = await import('canvas')
    const c = createCanvas(4, 4).getContext('2d'); c.font = '20px sans-serif'
    const first = c.measureText('AB').width, second = c.measureText('CD').width
    expect(index.pages[0].lines[0].text).toBe('ABCD')
    expect(hitTest(index, 0, 514, 190 - first * .5, { strict: true })).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 1 })
    expect(hitTest(index, 0, 514, 190 - first - second * .5, { strict: true })).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 3 })
    expect(hitTest(index, 0, 550, 190 - first, { strict: true })).toBeUndefined()
  })

  test('visible text in separated PPTX shapes receives strict hits on the correct line', async () => {
    const { parsePptx } = await import('../src/pptx/parse')
    const { buildPptx } = await import('../src/testdata/ooxml-builders')
    const doc = await parsePptx(await OfficePackage.load(await buildPptx([
      { off: ['914400', '914400'], ext: ['4572000', '1828800'], paragraphs: [{ runs: [{ text: 'needle-one' }] }] },
      { off: ['914400', '3657600'], ext: ['4572000', '1828800'], paragraphs: [{ runs: [{ text: 'needle-two' }] }] },
    ])))
    const index = await buildTextIndex(await getPaintables(doc))
    expect(hitTest(index, 0, 110, 110, { strict: true })?.lineIndex).toBe(0)
    expect(hitTest(index, 0, 110, 398, { strict: true })?.lineIndex).toBe(1)
    expect(hitTest(index, 0, 110, 250, { strict: true })).toBeUndefined()
  })
})


test('a mixed index retains strict hits on handwritten spans without affine metadata', async () => {
  const index = await buildTextIndex([{ spec: { widthPx: 500, heightPx: 300 }, paint: ctx => { ctx.font = '20px sans-serif'; ctx.translate(100, 50); ctx.fillText('transformed', 0, 20) } }])
  index.pages[0].lines.push({ text: 'manual', y: 200, top: 188, bottom: 203, spans: [{ text: 'manual', x: 50, y: 200, width: 60, fontSize: 12 }] })
  expect(hitTest(index, 0, 80, 195, { strict: true })).toEqual({ pageIndex: 0, lineIndex: 1, charIndex: 3 })
})

describe('ordinary line drag compatibility on transformed indexes', () => {
  test.each([
    { name: 'identity', a: 1, d: 1, e: 0, f: 0, mixed: false },
    { name: 'translated and scaled', a: 2, d: 3, e: 100, f: 50, mixed: false },
    { name: 'mixed rotated and plain', a: 1, d: 1, e: 0, f: 0, mixed: true },
  ])('$name rightward and leftward drags stay on the containing short line', async ({ a, d, e, f, mixed }) => {
    const index = await buildTextIndex([{ spec: { widthPx: 1000, heightPx: 600 }, paint: ctx => {
      ctx.font = '20px sans-serif'; ctx.setTransform(a, 0, 0, d, e, f)
      ctx.fillText('short', 10, 30)
      ctx.fillText('a much longer line of text stretching to the right', 10, 60)
      if (mixed) { ctx.setTransform(0, 1, -1, 0, 450, 5); ctx.fillText('rotated', 10, 30) }
    } }])
    const shortLine = index.pages[0].lines.findIndex(line => line.text === 'short')
    const pointerY = f + d * 25
    expect(hitTest(index, 0, e + a * 400, pointerY)).toEqual({ pageIndex: 0, lineIndex: shortLine, charIndex: 5 })
    expect(hitTest(index, 0, e - a * 400, pointerY)).toEqual({ pageIndex: 0, lineIndex: shortLine, charIndex: 0 })
    // Drag snapping does not turn a blank click beyond the line into a strict text hit.
    expect(hitTest(index, 0, e + a * 400, pointerY, { strict: true })).toBeUndefined()
    expect(hitTest(index, 0, e - a * 400, pointerY, { strict: true })).toBeUndefined()
  })
})

describe('nearest ordinary line between text bands', () => {
  test.each([
    { name: 'identity', a: 1, d: 1, e: 0, f: 0, handwritten: false },
    { name: 'translated and scaled', a: 2, d: 3, e: 100, f: 50, handwritten: false },
    { name: 'mixed recorded and handwritten metadata', a: 1, d: 1, e: 0, f: 0, handwritten: true },
  ])('$name inter-line drags follow nearest y even far beyond the short line', async ({ a, d, e, f, handwritten }) => {
    const index = await buildTextIndex([{ spec: { widthPx: 1000, heightPx: 600 }, paint: ctx => {
      ctx.font = '20px sans-serif'; ctx.setTransform(a, 0, 0, d, e, f)
      ctx.fillText('short', 10, 30)
      ctx.fillText('a much longer line of text stretching to the right', 10, 60)
    } }])
    if (handwritten) delete index.pages[0].lines[1].spans[0].placement
    for (const localY of [36, 37, 38, 39]) {
      expect(hitTest(index, 0, e + a * 400, f + d * localY)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 5 })
      expect(hitTest(index, 0, e - a * 400, f + d * localY)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 0 })
      expect(hitTest(index, 0, e + a * 400, f + d * localY, { strict: true })).toBeUndefined()
    }
    expect(hitTest(index, 0, e + a * 400, f + d * 40)?.lineIndex).toBe(1)
  })
})


test('ordinary recorded pages retain y-first drag carets on legacy empty manual lines', async () => {
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 120 }, paint: ctx => { ctx.font = '20px sans-serif'; ctx.fillText('ordinary', 0, 80) } }])
  index.pages[0].lines.unshift({ text: '', spans: [], y: 20, top: 10, bottom: 30 })
  expect(hitTest(index, 0, 150, 20)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 0 })
  expect(hitTest(index, 0, 150, 35)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 0 })
})

test('handwritten same-band lines use horizontal distance while gap drags keep nearest y', () => {
  const line = (text: string, x: number, y: number) => ({ text, y, top: y - 10, bottom: y + 4, spans: [{ text, x, y, width: 40, fontSize: 12 }] })
  const index = { pages: [{ index: 0, lines: [line('left', 0, 20), line('right', 200, 20), line('lower', 400, 60)] }] }
  expect(hitTest(index, 0, 245, 20)).toEqual({ pageIndex: 0, lineIndex: 1, charIndex: 5 })
  expect(hitTest(index, 0, 190, 20)).toEqual({ pageIndex: 0, lineIndex: 1, charIndex: 0 })
  expect(hitTest(index, 0, -100, 20)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 0 })
  expect(hitTest(index, 0, 500, 30)).toEqual({ pageIndex: 0, lineIndex: 0, charIndex: 4 })
})

test('stacked actual-paint carets snap to complete combining and emoji clusters', async () => {
  const { paintTextBody } = await import('../src/drawing/text-paint')
  const source = 'Ae\u0301👩🏽‍💻B'
  const body = { direction: 'wordArtVert' as const, paragraphs: [{ runs: [{ text: source, fontSizePt: 16 }], align: 'left' as const, level: 0 }],
    anchor: 't' as const, wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 }
  const index = await buildTextIndex([{ spec: { widthPx: 150, heightPx: 250 }, paint: ctx => paintTextBody(body, ctx, 0, 0, 150, 250, face => face) }])
  const line = index.pages[0].lines[0]
  expect(line.text).toBe(source)
  const boundaries = [0, 1, 3, 10, 11]
  for (const span of line.spans) {
    const placement = span.placement!
    const x = placement.transform.e + placement.x + placement.width * .43
    const y = placement.transform.f + placement.y
    const hit = hitTest(index, 0, x, y, { strict: true })
    expect(hit).toBeDefined()
    expect(boundaries).toContain(hit!.charIndex)
  }
  expect(textForRange(index, { start: { pageIndex: 0, lineIndex: 0, charIndex: 2 }, end: { pageIndex: 0, lineIndex: 0, charIndex: 5 } })).toBe('e\u0301👩🏽‍💻')
  expect(hitTest(index, 0, 149, 249, { strict: true })).toBeUndefined()
})

test.each([
  ['wordArtVert', 'STACK'], ['wordArtVertRtl', 'STACK'], ['eaVert', '中文'],
] as const)('default %s hits follow each upright glyph while ordinary horizontal drags keep y-first behavior', async (direction, text) => {
  const { paintTextBody } = await import('../src/drawing/text-paint')
  const stacked = { direction, paragraphs: [{ runs: [{ text, fontSizePt: 12 }], align: 'left' as const, level: 0 }],
    anchor: 't' as const, wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => paintTextBody(stacked, ctx, 0, 0, 200, 200, face => face) }])
  expect(index.pages[0].lines[0].spans.map((span, i) => {
    const p = span.placement!, x = p.transform.e + p.x + p.width * .8, y = p.transform.f + p.y
    expect(hitTest(index, 0, x, y, { strict: true })?.charIndex).toBe(i + 1)
    return hitTest(index, 0, x, y)?.charIndex
  })).toEqual([...text].map((_, i) => i + 1))
  expect(hitTest(index, 0, 199, 199, { strict: true })).toBeUndefined()
})

// Whole-line clockwise 'vert' (DrawingML vert / Word tbRl) rotates CJK too. Probe
// each glyph centre via the affine and assert the caret resolves per-grapheme.
test.each([['vert', '中文'], ['vert270', '中文']] as const)('whole-line %s rotates CJK and resolves per-glyph carets', async (direction, text) => {
  const { paintTextBody } = await import('../src/drawing/text-paint')
  const stacked = { direction, paragraphs: [{ runs: [{ text, fontSizePt: 12 }], align: 'left' as const, level: 0 }],
    anchor: 't' as const, wrap: true, insetLeftEmu: 0, insetRightEmu: 0, insetTopEmu: 0, insetBottomEmu: 0 }
  const index = await buildTextIndex([{ spec: { widthPx: 200, heightPx: 200 }, paint: ctx => paintTextBody(stacked, ctx, 0, 0, 200, 200, face => face) }])
  const line = index.pages[0].lines[0]
  const span = line.spans[0]
  const p = span.placement!, m = p.transform
  // Whole-line rotation: CJK transform is a rotation, never identity.
  expect(m.a === 1 && m.b === 0 && m.c === 0 && m.d === 1).toBe(false)
  // Probe each glyph centre and assert the caret resolves per-grapheme.
  const n = [...text].length
  for (let i = 0; i < n; i++) {
    const localX = p.x + (i + .5) * (p.width / n)
    const localY = p.y
    const pageX = m.e + m.a * localX + m.c * localY
    const pageY = m.f + m.b * localX + m.d * localY
    expect(hitTest(index, 0, pageX, pageY, { strict: true })?.charIndex).toBe(i + 1)
  }
})

test('word expansion respects complete source graphemes', () => {
  const line = { text: 'e\u0301 👩🏽‍💻', y: 20, top: 10, bottom: 25, spans: [{ text: 'e\u0301 👩🏽‍💻', x: 0, y: 20, width: 100, fontSize: 16 }] }
  expect(wordRangeAt(line, 1)).toEqual({ start: 0, end: 2 })
  expect(wordRangeAt(line, 5)).toEqual({ start: 3, end: 10 })
})
