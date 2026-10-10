// @vitest-environment node
import { createCanvas } from 'canvas'
import { expect, test } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseDocx } from '../src/docx/parse'
import { layoutDocx, renderPages } from '../src/docx/layout'
import { buildDocx } from '../src/testdata/ooxml-builders'

test('a thin single border defeats a wide dotted border using eighth-point conflict weights', async () => {
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [{ gridCols: ['1440', '1440'], rows: [{ heightTwips: '1440', cells: [
    { borders: '<w:right w:val="dotted" w:sz="32" w:color="FF0000"/>', paragraphs: [] },
    { borders: '<w:left w:val="single" w:sz="2" w:color="000000"/>', paragraphs: [] },
  ] }] }])))
  const pages = layoutDocx(doc, () => 0), table = pages[0].tables[0]
  const ctx = createCanvas(816, 1056).getContext('2d')
  const strokes: Array<{ color: string; width: number }> = []
  const stroke = ctx.stroke.bind(ctx)
  ctx.stroke = () => { strokes.push({ color: ctx.strokeStyle as string, width: ctx.lineWidth }); stroke() }
  renderPages(pages, ctx as unknown as CanvasRenderingContext2D)
  expect(table.rows[0].cells).toHaveLength(2)
  expect(strokes).toEqual([{ color: '#000000', width: 1 / 3 }])
})

test.each([false, true])('explicit nil diagonal clears inherited table border (merged=%s)', async merged => {
  const cell = { paragraphs: [], borders: '<w:tl2br w:val="nil"/>', ...(merged ? { vMerge: 'restart' as const } : {}) }
  const doc = await parseDocx(await OfficePackage.load(await buildDocx([], [{
    gridCols: ['1440'], borders: '<w:tl2br w:val="single" w:sz="8" w:color="0000FF"/>',
    rows: [{ cells: [cell] }, ...(merged ? [{ cells: [{ paragraphs: [], vMerge: 'continue' as const }] }] : [])],
  }])))
  const pages = layoutDocx(doc, () => 0)
  expect(pages[0].tables[0].rows[0].cells[0].borderSpecs?.tl2br).toBeUndefined()
  const ctx = createCanvas(816, 1056).getContext('2d')
  let strokes = 0
  ctx.stroke = () => { strokes++ }
  renderPages(pages, ctx as unknown as CanvasRenderingContext2D)
  expect(strokes).toBe(0)
})
