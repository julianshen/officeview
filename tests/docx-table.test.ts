import { describe, test, expect } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages, type MeasureFn } from '../src/docx/layout'
import { buildDocx, type DocxParaSpec, type DocxTableSpec } from '../src/testdata/ooxml-builders'

const measureFixed: MeasureFn = (text, style) =>
  text.length * style.fontSizePt * 0.6 * (96 / 72)

const p = (text: string): DocxParaSpec => ({ runs: [{ text }] })

describe('docx table parse', () => {
  test('parses grid, spans, merges, shading, borders, margins', async () => {
    const table: DocxTableSpec = {
      gridCols: ['4320', '4320', '3600'],
      fill: 'D9D9D9',
      borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
      rows: [
        {
          heightTwips: '600',
          cells: [
            { paragraphs: [p('Head 1')], fill: 'FFCC00' },
            { paragraphs: [p('Head 2')], gridSpan: 2, borders: '<w:bottom w:val="thick"/>' },
          ],
        },
        { cells: [{ paragraphs: [p('A')], vMerge: 'restart' }, { paragraphs: [p('B')] }, { paragraphs: [p('C')] }] },
        { cells: [{ vMerge: 'continue' }, { paragraphs: [p('D')] }, { paragraphs: [p('E')] }] },
      ],
    }
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([p('before')], [table], )))
    const section = doc.sections[0]
    expect(section.blocks).toHaveLength(2)
    const [before, tblBlock] = section.blocks
    expect(before.kind).toBe('p')
    expect(tblBlock.kind).toBe('table')
    const table2 = (tblBlock as { kind: 'table'; table: import('../src/docx/types').DocxTable }).table
    expect(table2.gridColsTwips).toEqual([4320, 4320, 3600])
    expect(table2.fill).toBe('D9D9D9')
    expect(table2.borders?.top?.style).toBe('single')
    expect(table2.cellMargins.leftTwips).toBe(108)
    const [row1, row2, row3] = table2.rows
    expect(row1.heightTwips).toBe(600)
    expect(row1.cells[1].gridSpan).toBe(2)
    expect(row1.cells[0].fill).toBe('FFCC00')
    expect(row1.cells[1].borders?.bottom?.style).toBe('thick')
    expect(row2.cells[0].vMerge).toBe('restart')
    expect(row3.cells[0].vMerge).toBe('continue')
    // paragraphs accessor still contains top-level paragraphs
    expect(section.paragraphs.map((x) => x.runs[0].text)).toEqual(['before'])
  })
})

describe('docx table layout', () => {
  test('lays cells into boxes; text lines carry absolute positions', async () => {
    const table: DocxTableSpec = {
      gridCols: ['2880', '5760'],
      borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
      rows: [
        { cells: [{ paragraphs: [p('one')] }, { paragraphs: [p('two')] }] },
        { cells: [{ paragraphs: [p('three')] }, { paragraphs: [p('four')] }] },
      ],
    }
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([p('intro')], [table])))
    const measure = measureFixed
    const pages = layoutDocx(doc, measure)
    const page = pages[0]
    expect(page.tables).toHaveLength(1)
    const box = page.tables[0]
    // table starts at left margin: 1440 twips
    expect(box.xPx).toBeCloseTo(1440 / 15, 0)
    expect(box.rows).toHaveLength(2)
    expect(box.rows[0].cells[0].fill).toBeUndefined()
    expect(box.rows[0].cells[0].borders?.left).toBe('thin')
    expect(box.rows[1].cells[1].borders?.top).toBe('thin')
    // text lines inside cells exist on the page
    const cellTexts = page.lines.flatMap((l) => l.segs.map((s) => s.text)).join('')
    expect(cellTexts).toContain('intro')
    expect(cellTexts).toContain('three')
    // cell line x positions are inside the table
    const tableLine = page.lines.find((l) => l.segs.some((s) => s.text === 'three'))
    expect(tableLine!.xPx).toBeGreaterThanOrEqual(box.xPx)
    expect(tableLine!.xPx).toBeLessThan(box.xPx + box.widthPx)
  })

  test('explicit row height atLeast respected', async () => {
    const table: DocxTableSpec = {
      gridCols: ['8640'],
      rows: [{ heightTwips: '1200', cells: [{ paragraphs: [p('tall')] }] }],
    }
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [table])))
    const pages = layoutDocx(doc, measureFixed)
    const box = pages[0].tables[0]
    expect(box.rows[0].heightPx).toBeGreaterThanOrEqual(1200 / 15)
  })
})

describe('docx table render', () => {
  test('paints fills, borders, and cell text onto canvas', async () => {
    const table: DocxTableSpec = {
      gridCols: ['4320', '4320'],
      borders: '<w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/>',
      rows: [
        { cells: [{ paragraphs: [{ align: 'center', runs: [{ text: 'HEADER', bold: true }] }], fill: 'FFCC00' }, { paragraphs: [p('plain')] }] },
      ],
    }
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [table])))
    const pages = layoutDocx(doc, measureFixed)
    const page = pages[0]

    const { createCanvas } = await import('canvas')
    const canvas = createCanvas(Math.ceil(page.widthPx), Math.ceil(page.heightPx))
    const ctx = canvas.getContext('2d')!
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    renderPages(pages, ctx as unknown as CanvasRenderingContext2D)

    const box = page.tables[0]
    const c0 = box.rows[0].cells[0]
    // center of first cell must be the yellow fill (cell coords are table-relative)
    const cx = Math.round(box.xPx + c0.xPx + c0.widthPx / 2)
    const cy = Math.round(box.yPx + c0.yPx + c0.heightPx / 2)
    const px = ctx.getImageData(cx, cy, 1, 1).data
    expect([px[0], px[1], px[2]]).toEqual([255, 204, 0])
    // border line exists along the table's top edge (1px line straddles two
    // pixel rows antialiased at ~50%, so sample 3 rows, threshold < 200)
    let dark = 0
    const row = ctx.getImageData(Math.round(box.xPx + 5), Math.round(box.yPx) - 1, Math.round(box.widthPx) - 10, 3)
    for (let i = 0; i < row.data.length; i += 4) if (row.data[i] < 200) dark++
    expect(dark).toBeGreaterThan(30)
    // header text ink inside first cell
    const region = ctx.getImageData(Math.round(box.xPx + c0.xPx), Math.round(box.yPx + c0.yPx), Math.round(c0.widthPx), Math.round(c0.heightPx))
    let ink = 0
    for (let i = 0; i < region.data.length; i += 4) if (region.data[i] < 100) ink++
    expect(ink).toBeGreaterThan(50)
  })

  test('vMerge continue cells render no text', async () => {
    const table: DocxTableSpec = {
      gridCols: ['2880', '5760'],
      borders: '<w:insideH w:val="single"/>',
      rows: [
        { cells: [{ paragraphs: [p('merged')], vMerge: 'restart' }, { paragraphs: [p('r1c2')] }] },
        { cells: [{ vMerge: 'continue' }, { paragraphs: [p('r2c2')] }] },
      ],
    }
    const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [table])))
    const pages = layoutDocx(doc, measureFixed)
    const texts = pages[0].lines.flatMap((l) => l.segs.map((s) => s.text)).join('')
    expect(texts).toContain('merged')
    expect(texts).toContain('r1c2')
    expect(texts).toContain('r2c2')
    // no extra text for continue cell
    expect(texts).not.toContain('r2c1')
  })
})
