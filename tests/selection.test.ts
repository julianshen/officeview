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
  rectsForSelection,
  textForRange,
  wordRangeAt,
  type CaretPos,
} from '../src/core/selection'
import { buildDocx, buildXlsx } from '../src/testdata/ooxml-builders'

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
    expect(range.startLine).toBe(1)
    expect(range.startChar).toBe(2)
    expect(range.endLine).toBe(2)
    expect(range.endChar).toBe(5)
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
    const text = textForRange(index, {
      pageIndex: 0,
      startLine: 0,
      startChar: 6,
      endLine: 0,
      endChar: 11,
    })
    expect(text).toBe('world')
  })

  test('multi-line selection joins with newlines and trims trailing spaces', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'first line' }] },
      { runs: [{ text: 'second line' }] },
      { runs: [{ text: 'third line' }] },
    ])
    expect(index.pages[0].lines.length).toBe(3)
    const text = textForRange(index, {
      pageIndex: 0,
      startLine: 0,
      startChar: 0,
      endLine: 2,
      endChar: 10,
    })
    expect(text.split('\n')).toHaveLength(3)
    expect(text.startsWith('first line')).toBe(true)
    expect(text.endsWith('third line')).toBe(true)
  })

  test('partial first/last lines are respected', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'aaaa bbbb cccc' }] },
      { runs: [{ text: 'dddd eeee ffff' }] },
    ])
    const text = textForRange(index, {
      pageIndex: 0,
      startLine: 0,
      startChar: 5,
      endLine: 1,
      endChar: 4,
    })
    expect(text).toBe('bbbb cccc\ndddd')
  })

  test('produces one rect per touched line', async () => {
    const index = await indexDocx([
      { runs: [{ text: 'aaaa bbbb cccc' }] },
      { runs: [{ text: 'dddd eeee ffff' }] },
    ])
    const rects = rectsForSelection(index, {
      pageIndex: 0,
      startLine: 0,
      startChar: 5,
      endLine: 1,
      endChar: 4,
    })
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
    const rects = rectsForSelection(index, {
      pageIndex: 0,
      startLine: 0,
      startChar: 0,
      endLine: 0,
      endChar: line.text.length,
    })
    const covered = rects.reduce((a, r) => a + r.width, 0)
    const total = line.spans.reduce((a, s) => a + s.width, 0)
    expect(covered).toBeCloseTo(total, 0)
  })
})