import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parsePptx } from '../src/pptx/parse'
import { slideMetrics, renderSlide } from '../src/pptx/render'
import { buildPptx, type PptxShapeSpec, type PptxTableSpec } from '../src/testdata/ooxml-builders'

const BORDERS = { border: 'C00000', accent: '1F497D' }

async function parseFixture(shapes: PptxShapeSpec[]) {
  return parsePptx(await OfficePackage.load(await buildPptx(shapes)))
}

/** Slide-sized frame wrapping a table spec. */
function tableFrame(table: PptxTableSpec): PptxShapeSpec {
  return {
    off: ['914400', '914400'],
    ext: ['4572000', '1828800'],
    table,
  }
}

const sampleTable = {
  colWidths: ['2286000', '2286000'],
  rows: [
    {
      cells: [
        { paragraphs: [{ align: 'ctr', runs: [{ text: 'Item', b: true, sz: '1600', color: 'FFFFFF' }] }], fill: BORDERS.border },
        { paragraphs: [{ align: 'ctr', runs: [{ text: 'Status', b: true, sz: '1600', color: 'FFFFFF' }] }], fill: BORDERS.accent },
      ],
    },
    {
      cells: [
        { paragraphs: [{ runs: [{ text: 'Row A', sz: '1400' }] }] },
        { paragraphs: [{ runs: [{ text: 'on track', sz: '1400', color: '0070C0' }] }] },
      ],
    },
  ],
}

describe('pptx table parse', () => {
  test('parses grid, cells, text and fills from p:graphicFrame/a:tbl', async () => {
    const doc = await parseFixture([tableFrame(sampleTable)])
    const frame = doc.slides[0].shapes[0]
    expect(frame.xEmu).toBe(914400)
    expect(frame.widthEmu).toBe(4572000)
    expect(frame.table).toBeDefined()
    const table = frame.table!
    expect(table.colWidthsEmu).toEqual([2286000, 2286000])
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0].cells[0].paragraphs[0].align).toBe('center')
    expect(table.rows[0].cells[0].paragraphs[0].runs[0]).toMatchObject({ text: 'Item', bold: true, fontSizePt: 16 })
    expect(table.rows[0].cells[0].fill).toBe('#C00000')
    expect(table.rows[1].cells[1].paragraphs[0].runs[0].text).toBe('on track')
  })

  test('parses gridSpan, rowSpan and merged-away cells', async () => {
    const doc = await parseFixture([
      tableFrame({
        colWidths: ['1143000', '1143000', '2286000'],
        rows: [
          {
            cells: [
              { gridSpan: 2, paragraphs: [{ runs: [{ text: 'Spans two' }] }] },
              { paragraphs: [{ runs: [{ text: 'third' }] }] },
            ],
          },
          {
            cells: [
              { merged: true },
              { merged: true },
              { paragraphs: [{ runs: [{ text: 'tail' }] }] },
            ],
          },
        ],
      }),
    ])
    const table = doc.slides[0].shapes[0].table!
    expect(table.rows[0].cells[0].gridSpan).toBe(2)
    expect(table.rows[1].cells[0].merged).toBe(true)
    expect(table.rows[1].cells[2].paragraphs[0].runs[0].text).toBe('tail')
  })

  test('reads explicit row heights', async () => {
    const doc = await parseFixture([
      tableFrame({
        colWidths: ['2286000'],
        rows: [
          { h: '457200', cells: [{ paragraphs: [{ runs: [{ text: 'fixed' }] }] }] },
          { cells: [{ paragraphs: [{ runs: [{ text: 'auto' }] }] }] },
        ],
      }),
    ])
    const table = doc.slides[0].shapes[0].table!
    expect(table.rows[0].heightEmu).toBe(457200)
    expect(table.rows[1].heightEmu).toBeUndefined()
  })
})

describe('pptx table render', () => {
  async function renderTable(table: PptxTableSpec) {
    const doc = await parseFixture([tableFrame(table)])
    const m = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m)
    return { canvas, ctx, m }
  }

  test('paints cell fills at the right position', async () => {
    const { ctx, m } = await renderTable(sampleTable)
    // frame at 1in,1in size 5in x 2in; first cell = left half of the top row
    // frame at (96,96) size 480x192 -> columns 96..336..576, rows 96..192..288
    // sample near each cell's bottom-left so we hit fill, not a glyph
    const px = ctx.getImageData(110, 182, 1, 1).data
    expect([px[0], px[1], px[2]]).toEqual([192, 0, 0]) // C00000
    // second header cell is the accent color
    const px2 = ctx.getImageData(350, 182, 1, 1).data
    expect([px2[0], px2[1], px2[2]]).toEqual([31, 73, 125]) // 1F497D
    expect(m.widthPx).toBeGreaterThan(0)
  })

  test('paints cell text inside the cell', async () => {
    const { ctx } = await renderTable(sampleTable)
    // header runs are white on a dark fill: count near-white glyph pixels
    const regionImg = ctx.getImageData(96, 120, 240, 48)
    let ink = 0
    for (let i = 0; i < regionImg.data.length; i += 4) {
      if (regionImg.data[i] > 230 && regionImg.data[i + 1] > 230 && regionImg.data[i + 2] > 230) ink++
    }
    expect(ink).toBeGreaterThan(20)
  })

  test('grid lines are drawn and the frame is clipped to its extent', async () => {
    const { ctx } = await renderTable(sampleTable)
    // sample just inside the left frame edge: should be grid/fill, not white page
    const onEdge = ctx.getImageData(97, 140, 1, 1).data
    expect(onEdge[3]).toBe(255)
    // well outside the frame stays page-white
    const outside = ctx.getImageData(800, 650, 1, 1).data
    expect([outside[0], outside[1], outside[2]]).toEqual([255, 255, 255])
  })
})

describe('pptx table style resolution', () => {
  const styled = {
    colWidths: ['2286000', '2286000'],
    firstRow: true,
    bandRow: true,
    styleId: 'FixtureTableStyle',
    rows: [
      { cells: [{ paragraphs: [{ align: 'ctr', runs: [{ text: 'Item' }] }] }, { paragraphs: [{ align: 'ctr', runs: [{ text: 'Qty' }] }] }] },
      { cells: [{ paragraphs: [{ runs: [{ text: 'a' }] }] }, { paragraphs: [{ runs: [{ text: '1' }] }] }] },
      { cells: [{ paragraphs: [{ runs: [{ text: 'b' }] }] }, { paragraphs: [{ runs: [{ text: '2' }] }] }] },
      { cells: [{ paragraphs: [{ runs: [{ text: 'c' }] }] }, { paragraphs: [{ runs: [{ text: '3' }] }] }] },
    ],
  }

  test('resolves styleId, flags and per-region fills from tableStyles.xml', async () => {
    const doc = await parseFixture([tableFrame(styled)])
    const table = doc.slides[0].shapes[0].table!
    expect(table.styleId).toBe('FixtureTableStyle')
    expect(table.firstRow).toBe(true)
    expect(table.bandRow).toBe(true)
    expect(table.styleFills?.firstRow).toBe('#1F4E79')
    expect(table.styleFills?.band1).toBe('#DDEBF7')
    expect(table.styleFills?.band2).toBe('#FFFFFF')
    // header text switches to the style's font color
    expect(table.firstRowTextColor).toBe('#FFFFFF')
  })

  test('paints the header fill from the style', async () => {
    const doc = await parseFixture([tableFrame(styled)])
    const m = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m)
    // frame at (96,96) 480x192; 4 rows -> 48px each; sample inside row 0
    const header = ctx.getImageData(110, 140, 1, 1).data
    expect([header[0], header[1], header[2]]).toEqual([31, 78, 121]) // 1F4E79
  })

  test('bands alternate below the header row', async () => {
    const doc = await parseFixture([tableFrame(styled)])
    const m = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m)
    const at = (rowIndex: number) => {
      const y = 96 + rowIndex * 48 + 44
      const px = ctx.getImageData(110, y, 1, 1).data
      return [px[0], px[1], px[2]]
    }
    expect(at(1)).toEqual([221, 235, 247]) // DDEBF7 (band1)
    expect(at(2)).toEqual([255, 255, 255]) // band2
    expect(at(3)).toEqual([221, 235, 247]) // band1 again
  })

  test('an explicit cell fill overrides the style banding', async () => {
    const doc = await parseFixture([
      tableFrame({
        ...styled,
        rows: [
          ...styled.rows.slice(0, 1),
          { cells: [{ paragraphs: [{ runs: [{ text: 'override' }] }], fill: 'FF0000' }, { paragraphs: [{ runs: [{ text: 'x' }] }] }] },
        ],
      }),
    ])
    const m = slideMetrics(doc)
    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(m.widthPx, m.heightPx)
    const ctx = canvas.getContext('2d')!
    renderSlide(doc.slides[0], ctx as unknown as CanvasRenderingContext2D, m)
    // this table has 2 rows -> 96px each; sample the overridden cell
    const px = ctx.getImageData(110, 96 + 96 + 80, 1, 1).data
    expect([px[0], px[1], px[2]]).toEqual([255, 0, 0])
  })
})
