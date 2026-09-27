import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { getPaintables } from '../src/render/paint'
import { buildTextIndex, findMatches, stepMatch } from '../src/core/search'
import { parseDocx } from '../src/docx/parse'
import { parseXlsx } from '../src/xlsx/parse'
import { parsePptx } from '../src/pptx/parse'
import { buildDocx, buildXlsx, buildPptx } from '../src/testdata/ooxml-builders'

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
