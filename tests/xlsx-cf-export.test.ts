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
it('PNG encoder retains the prepared CF bar (injected evidence)', async () => {
  const { barSheet, withBars } = await import('./helpers/cf-bars')
  const { computeMetrics } = await import('../src/xlsx/render')
  await withBars(async () => {
    const sheet = barSheet([0, 50, 100])
    const m = computeMetrics(sheet)
    const rowMid = (r: number) => Math.round(m.rowHeightsPx.slice(0, r).reduce((a, b) => a + b, 0) + m.rowHeightsPx[r]! / 2)
    const pages = await getPaintables({ sheets: [sheet] })
    try {
      const c = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
      pages[0].paint(c.getContext('2d') as unknown as CanvasRenderingContext2D)
      const png = c.toBuffer('image/png'), decoded = await loadImage(png), copy = createCanvas(decoded.width, decoded.height)
      copy.getContext('2d').drawImage(decoded, 0, 0)
      const at = (x: number, y: number) => [...copy.getContext('2d').getImageData(x, y, 1, 1).data]
      expect(at(30, rowMid(2))).toEqual([0, 0, 255, 255])
      expect(at(30, rowMid(1))).toEqual([0, 0, 255, 255])
      expect(at(160, rowMid(1))).not.toEqual([0, 0, 255, 255])
      expect(at(30, rowMid(0))).not.toEqual([0, 0, 255, 255])
    } finally { pages.dispose() }
  })
})
it('PNG encoder retains the prepared CF icons (injected stubs); with no glyph table none paints', async () => {
  const { iconSheet, withIcons } = await import('./helpers/cf-icons')
  const { computeMetrics } = await import('../src/xlsx/render')
  const sample = async (sheet: import('../src/xlsx/types').XlsxSheet) => {
    const m = computeMetrics(sheet)
    const rowMid = (r: number) => Math.round(m.rowHeightsPx.slice(0, r).reduce((a, b) => a + b, 0) + m.rowHeightsPx[r]! / 2)
    const pages = await getPaintables({ sheets: [sheet] })
    try {
      const c = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
      pages[0].paint(c.getContext('2d') as unknown as CanvasRenderingContext2D)
      const png = c.toBuffer('image/png'), decoded = await loadImage(png), copy = createCanvas(decoded.width, decoded.height)
      copy.getContext('2d').drawImage(decoded, 0, 0)
      // The stub box: 2 pt in, 9 pt wide (2.7 px to 14.7 px), centred in the row.
      return [0, 1, 2].map(r => [...copy.getContext('2d').getImageData(8, rowMid(r), 1, 1).data])
    } finally { pages.dispose() }
  }
  const painted = await withIcons(() => sample(iconSheet([10, 50, 90])))
  expect(painted).toEqual([[0x11, 0xaa, 0x11, 255], [0xcc, 0x88, 0x00, 255], [0xcc, 0x00, 0x11, 255]])
  const withheld = await withIcons(() => sample(iconSheet([10, 50, 90])), { glyphs: () => undefined })
  for (const px of withheld) expect(px).toEqual([255, 255, 255, 255])
})

it('Stage 5: the PNG encoder retains a formula rule\'s prepared fill and carries nothing without evidence', async () => {
  const { load, withPaint, workbookBytes, cfBlock, exprRule } = await import('./helpers/cf-views')
  const bytes = workbookBytes({}, cfBlock('B2', exprRule('ISBLANK(B2)', 1, 1)))
  const render = async (): Promise<number[]> => {
    const pages = await getPaintables((await load(bytes, 'cached')).doc)
    try {
      const c = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
      pages[0].paint(c.getContext('2d') as unknown as CanvasRenderingContext2D)
      const decoded = await loadImage(c.toBuffer('image/png')), copy = createCanvas(decoded.width, decoded.height)
      copy.getContext('2d').drawImage(decoded, 0, 0)
      return [...copy.getContext('2d').getImageData(70, 25, 1, 1).data]
    } finally { pages.dispose() }
  }
  expect(await withPaint(render)).toEqual([0xc6, 0xef, 0xce, 255])
  // Task 39: "without evidence" is production before --publish-stage5 (Stage 1-4 evidence only).
  const { withoutStage5Evidence } = await import('./helpers/cf-stage5-off')
  expect(await withoutStage5Evidence(render)).not.toEqual([0xc6, 0xef, 0xce, 255])
})
