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
it('PNG encoder retains the prepared CF scale fill', async () => {
  const { withInjectedCfEvidence } = await import('../src/xlsx/cf/evidence')
  const scaleEvidence = [
    'scale:interpolation-color-space',
    'scale:interpolation-rounding-ladder',
    'scale:thresholds-min-max',
  ]
  await withInjectedCfEvidence(scaleEvidence, async () => {
    const pages = await getPaintables({
      sheets: [
        {
          name: 'ScaleCF',
          merges: [],
          mergeRanges: [],
          rows: [
            { index: 0, cells: [{ row: 0, col: 0, ref: 'A1', value: 10, styleIndex: 0 }] },
            { index: 1, cells: [{ row: 1, col: 0, ref: 'A2', value: 70, styleIndex: 0 }] },
          ],
          cols: [],
          conditionalFormatting: {
            diagnostics: [],
            rules: [
              {
                id: 1,
                priority: 1,
                documentOrder: 0,
                stopIfTrue: false,
                ranges: [{ firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 0 }],
                writes: new Set(['colorScale', 'fill']),
                definition: {
                  kind: 'colorScale',
                  points: [
                    { bound: { type: 'min' }, argb: 0xffff0000 },
                    { bound: { type: 'max' }, argb: 0xff00ff00 },
                  ],
                  dialect: 'base',
                  formulas: [],
                },
              },
            ],
          },
        },
      ],
    })
    try {
      const c = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
      pages[0].paint(c.getContext('2d') as unknown as CanvasRenderingContext2D)
      const png = c.toBuffer('image/png')
      const decoded = await loadImage(png)
      const copy = createCanvas(decoded.width, decoded.height)
      copy.getContext('2d').drawImage(decoded, 0, 0)
      // Cell A1 is red (min) -> [255, 0, 0, 255]
      expect([...copy.getContext('2d').getImageData(10, 5, 1, 1).data]).toEqual([255, 0, 0, 255])
    } finally {
      pages.dispose()
    }
  })
})
