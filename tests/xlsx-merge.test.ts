import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx, parseMergeRange } from '../src/xlsx/parse'
import { computeMetrics, renderSheet } from '../src/xlsx/render'
import { buildXlsx } from '../src/testdata/ooxml-builders'

describe('merge range parsing', () => {
  test('A1:B3 -> 0-based inclusive', () => {
    expect(parseMergeRange('A1:B3')).toEqual({ minRow: 0, minCol: 0, maxRow: 2, maxCol: 1 })
    expect(parseMergeRange('C2')).toEqual({ minRow: 1, minCol: 2, maxRow: 1, maxCol: 2 })
  })
})

async function mergedFixture() {
  return buildXlsx([
    {
      name: 'M',
      rows: [
        { r: 1, cells: [{ ref: 'A1', t: 's', v: 0, style: 1 }, { ref: 'C1', t: 's', v: 1 }] },
        { r: 2, cells: [{ ref: 'A2', v: 99.5, style: 2 }] },
      ],
      merges: ['A1:B1'],
    },
  ], ['Merged', 'Plain'])
}

describe('merged cell rendering', () => {
  test('anchor fill and text span the merge; covered cell paints nothing', async () => {
    const doc = await parseXlsx(await OfficePackage.load(await mergedFixture()))
    const sheet = doc.sheets[0]
    expect(sheet.mergeRanges).toEqual([{ minRow: 0, minCol: 0, maxRow: 0, maxCol: 1 }])
    const m = computeMetrics(sheet)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)

    // yellow fill spans BOTH columns of row 1 (A1's fill)
    // Sample inside column B, safely clear of the anchor cell's text glyphs across font fallbacks
    const midB = Math.round(m.colWidthsPx[0] + m.colWidthsPx[1] - 15) // inside column B
    const px = ctx.getImageData(midB, 8, 1, 1).data
    expect([px[0], px[1], px[2]]).toEqual([255, 255, 0])

    // gridline between A and B is hidden in row 1 but visible in row 2
    // Sample in row 1's top padding (y=2) well above text ascenders
    const xb = m.colWidthsPx[0]
    const lineRow1 = ctx.getImageData(xb, 2, 1, 1).data
    expect(lineRow1[0]).toBe(255) // no gray line inside merge
    const lineRow2 = ctx.getImageData(xb, 24, 1, 1).data
    expect(lineRow2[0]).toBeLessThan(230) // gridline present below merge

    // covered B1 has no text: region right of the header text is pure yellow
    const tail = ctx.getImageData(m.colWidthsPx[0] + m.colWidthsPx[1] - 15, 5, 10, 12)
    let ink = 0
    for (let i = 0; i < tail.data.length; i += 4) if (tail.data[i] < 150 && tail.data[i + 2] > 100) ink++
    expect(ink).toBe(0)
  })

  test('right-aligned number in merged range aligns to merge right edge', async () => {
    const buf = await buildXlsx([
      {
        name: 'M2',
        rows: [
          { r: 1, cells: [{ ref: 'A1', v: 7.25, style: 2 }] },
          { r: 2, cells: [{ ref: 'A2', v: 1 }] },
        ],
        merges: ['A1:B1'],
      },
    ])
    const doc = await parseXlsx(await OfficePackage.load(buf))
    const sheet = doc.sheets[0]
    const m = computeMetrics(sheet)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSheet(sheet, ctx as unknown as CanvasRenderingContext2D, m)

    // text must end near merge right edge (col B), not col A's right edge
    const band = ctx.getImageData(0, 0, m.widthPx, m.rowHeightsPx[0])
    // find rightmost dark pixel in row 1
    let rightmost = -1
    for (let x = 0; x < m.widthPx; x++) {
      for (let y = 0; y < m.rowHeightsPx[0]; y++) {
        const px = ctx.getImageData(x, y, 1, 1).data
        if (px[0] < 100) { rightmost = x; y = m.rowHeightsPx[0]; break }
      }
    }
    expect(rightmost).toBeGreaterThan(m.colWidthsPx[0] - 10)
    void band
  })
})
