// @vitest-environment node
import { createCanvas, loadImage } from 'canvas'
import { expect, it } from 'vitest'
import { getPaintables } from '../src/render/paint'
it('PNG encoder retains the prepared CF fill', async () => {
  const pages = await getPaintables({ sheets: [{ name: 'CF', merges: [], mergeRanges: [], rows: [], cols: [], conditionalFormatting: { diagnostics: [], rules: [{ id: 0, priority: 1, documentOrder: 0, stopIfTrue: false, ranges: [{ firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 }], writes: new Set(['fill']), dxf: { 'fill.bg': 'FF00FF00' }, definition: { kind: 'predicate', type: 'containsBlanks', dialect: 'base', formulas: [] } }] } }] })
  try {
    const c = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
    pages[0].paint(c.getContext('2d') as unknown as CanvasRenderingContext2D)
    const png = c.toBuffer('image/png'), decoded = await loadImage(png), copy = createCanvas(decoded.width, decoded.height)
    copy.getContext('2d').drawImage(decoded, 0, 0)
    expect([...copy.getContext('2d').getImageData(10, 5, 1, 1).data]).toEqual([0, 255, 0, 255])
  } finally { pages.dispose() }
})
