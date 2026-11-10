// @vitest-environment node
import { createCanvas } from 'canvas'
import { describe, expect, it, vi } from 'vitest'
import type { XlsxSheet, XlsxCell } from '../src/xlsx/types'
import { computeMetrics, computePrintMetrics, renderSheet } from '../src/xlsx/render'
import { buildSheetGeometry } from '../src/xlsx/view/viewport'
import { renderSheetViewport } from '../src/xlsx/view/renderViewport'
import { LayoutCache } from '../src/xlsx/view/textLayout'
import { prepareCfPaint } from '../src/xlsx/cf/render'
import * as regionGridModule from '../src/xlsx/regionGrid'
import { withInjectedCfEvidence } from '../src/xlsx/cf/evidence'
import { cfRuntimeEvidence } from '../src/xlsx/cf/predicates'
import { getPaintables } from '../src/render/paint'

const SCALE_EVIDENCE = [
  'scale:interpolation-color-space',
  'scale:interpolation-rounding-ladder',
  'scale:thresholds-min-max',
  'scale:scope-whole-column',
]

function makeScaleSheet(minRow = 1, maxRow = 5, minCol = 1, maxCol = 1): XlsxSheet {
  const cells: XlsxCell[] = []
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      cells.push({
        ref: `${String.fromCharCode(65 + c)}${r + 1}`,
        row: r,
        col: c,
        value: (r - minRow + 1) * 10,
        styleIndex: 0,
      })
    }
  }
  return {
    name: 'ScaleSheet',
    sheetId: '1',
    rows: Array.from({ length: 12 }, (_, r) => ({
      index: r,
      cells: cells.filter(cell => cell.row === r),
    })),
    cols: [],
    merges: [],
    mergeRanges: [],
    conditionalFormatting: {
      diagnostics: [],
      rules: [
        {
          id: 1,
          priority: 1,
          documentOrder: 0,
          stopIfTrue: false,
          ranges: [{ firstRow: minRow, lastRow: maxRow, firstCol: minCol, lastCol: maxCol }],
          writes: new Set(['colorScale', 'fill']),
          definition: {
            kind: 'colorScale',
            points: [
              { bound: { type: 'min' }, argb: 0xfff8696b },
              { bound: { type: 'max' }, argb: 0xff63be7b },
            ],
            dialect: 'base',
            formulas: [],
          },
        },
      ],
    },
  } as XlsxSheet
}

describe('Task 16 — All-view painting and gridlines', () => {
  it('scale fill suppresses gridlines on a CF-only sheet', async () => {
    // A sheet with no tables or pivots and a min/max scale on B2:B6 has no gridline
    // segments inside the filled cells and unchanged segments elsewhere.
    // A no-CF sheet keeps its existing metrics exactly. Assert that only the one
    // gridline path is used (regionGrid).
    await withInjectedCfEvidence(SCALE_EVIDENCE, async () => {
      const sheet = makeScaleSheet(1, 5, 1, 1) // B2:B6
      const noCfSheet: XlsxSheet = {
        name: 'PlainSheet',
        sheetId: 'plain',
        rows: sheet.rows.map(r => ({ ...r, cells: r.cells.map(c => ({ ...c })) })),
        cols: [],
        merges: [],
        mergeRanges: [],
      } as XlsxSheet

      // No-CF sheet keeps its existing metrics exactly
      const cfMetrics = computeMetrics(sheet)
      const plainMetrics = computeMetrics(noCfSheet)
      expect(cfMetrics.widthPx).toBe(plainMetrics.widthPx)
      expect(cfMetrics.heightPx).toBe(plainMetrics.heightPx)
      expect(cfMetrics.rowHeightsPx).toEqual(plainMetrics.rowHeightsPx)
      expect(cfMetrics.colWidthsPx).toEqual(plainMetrics.colWidthsPx)

      // Guard check: Plain sheet costs only the guard
      expect(regionGridModule.mayHaveRegionGrid(noCfSheet)).toBe(false)
      // CF-only sheet guard evaluates to true
      expect(regionGridModule.mayHaveRegionGrid(sheet)).toBe(true)

      const regionGridSpy = vi.spyOn(regionGridModule, 'regionGrid')

      // Paint plain sheet: regionGrid must NOT be called
      const plainCtx = createCanvas(cfMetrics.widthPx, cfMetrics.heightPx).getContext('2d')
      renderSheet(noCfSheet, plainCtx as unknown as CanvasRenderingContext2D, plainMetrics)
      expect(regionGridSpy).not.toHaveBeenCalled()

      // Prepare CF for scale sheet
      const batch = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      expect(batch).toBeDefined()

      // Track gridline stroke lines in scale sheet vs plain sheet
      type Segment = { x1: number; y1: number; x2: number; y2: number }
      function collectGridlines(s: XlsxSheet, m: ReturnType<typeof computeMetrics>, b?: typeof batch): Segment[] {
        const c = createCanvas(m.widthPx, m.heightPx)
        const ctx = c.getContext('2d')
        const segments: Segment[] = []
        let lastMove = { x: 0, y: 0 }
        const origMoveTo = ctx.moveTo.bind(ctx)
        const origLineTo = ctx.lineTo.bind(ctx)
        ctx.moveTo = (x, y) => { lastMove = { x, y }; origMoveTo(x, y) }
        ctx.lineTo = (x, y) => { segments.push({ x1: lastMove.x, y1: lastMove.y, x2: x, y2: y }); origLineTo(x, y) }
        renderSheet(s, ctx as unknown as CanvasRenderingContext2D, m, undefined, b ? { conditionalFormatting: b } : undefined)
        return segments
      }

      const cfSegments = collectGridlines(sheet, cfMetrics, batch!)
      expect(regionGridSpy).toHaveBeenCalled()

      // B2:B6 covers rows 1..5, col 1.
      // Row boundaries between row 1 and 2, 2 and 3, 3 and 4, 4 and 5 are internal horizontal gridlines.
      // For each row r in 1..4, the horizontal gridline at rowY[r+1] inside colX[1]..colX[2] should be suppressed.
      const col1Left = cfMetrics.colWidthsPx[0] ?? 0
      const col1Right = col1Left + (cfMetrics.colWidthsPx[1] ?? 0)
      let rowY = 0
      const rowTops: number[] = []
      for (const h of cfMetrics.rowHeightsPx) {
        rowTops.push(rowY)
        rowY += h
      }
      for (let r = 2; r <= 5; r++) {
        const y = rowTops[r]!
        const internalHLine = cfSegments.find(seg =>
          Math.abs(seg.y1 - (y + 0.5)) < 1 &&
          Math.abs(seg.y2 - (y + 0.5)) < 1 &&
          seg.x1 < col1Right - 1 && seg.x2 > col1Left + 1,
        )
        // Inside the filled cells B2:B6, internal horizontal gridline is suppressed
        expect(internalHLine, `horizontal gridline between B${r} and B${r + 1} should be suppressed`).toBeUndefined()
      }

      // Elsewhere (e.g. rows 7..10 across all cols): gridline segments exist unchanged
      const row8Y = rowTops[8]!
      const outsideLine = cfSegments.find(seg =>
        Math.abs(seg.y1 - (row8Y + 0.5)) < 1 &&
        Math.abs(seg.y2 - (row8Y + 0.5)) < 1,
      )
      expect(outsideLine, 'gridline outside filled cells should remain').toBeDefined()

      regionGridSpy.mockRestore()
      batch?.dispose()
    })
  })

  it('tile and full render carry the same scale fill at six scales', async () => {
    // At 50/100/175% zoom and DPR 1/2, a filled host that crosses a tile edge has
    // identical source-space fill primitives (rect and ARGB) in the full and tiled renders.
    await withInjectedCfEvidence(SCALE_EVIDENCE, async () => {
      // Create a sheet where B2 is at a coordinate crossing a 256px / 512px tile boundary
      const sheet = makeScaleSheet(1, 3, 1, 2)
      // Expand columns so cell crosses 256px tile edge
      sheet.cols = [
        { min: 0, max: 0, widthChars: 30 },
        { min: 1, max: 1, widthChars: 35 },
      ]
      const metrics = computeMetrics(sheet)
      const geometry = buildSheetGeometry(sheet)
      const batch = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      expect(batch).toBeDefined()

      const zooms = [0.5, 1.0, 1.75]
      const dprs = [1, 2]

      for (const zoom of zooms) {
        for (const dpr of dprs) {
          // 1. Full render via renderSheet
          const fullCanvas = createCanvas(Math.ceil(metrics.widthPx * zoom * dpr), Math.ceil(metrics.heightPx * zoom * dpr))
          const fullCtx = fullCanvas.getContext('2d')
          fullCtx.scale(zoom * dpr, zoom * dpr)

          type FillCall = { fillStyle: string; x: number; y: number; w: number; h: number; tileX?: number; tileY?: number }
          const fullFills: FillCall[] = []
          const origFullFillRect = fullCtx.fillRect.bind(fullCtx)
          fullCtx.fillRect = (x, y, w, h) => {
            fullFills.push({ fillStyle: String(fullCtx.fillStyle), x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h) })
            origFullFillRect(x, y, w, h)
          }
          renderSheet(sheet, fullCtx as unknown as CanvasRenderingContext2D, metrics, undefined, { conditionalFormatting: batch })

          // 2. Tiled render via renderSheetViewport
          const tileSize = 256
          const tiledFills: FillCall[] = []
          const cache = new LayoutCache()

          for (let tx = 0; tx < metrics.widthPx; tx += tileSize) {
            for (let ty = 0; ty < metrics.heightPx; ty += tileSize) {
              const tileCanvas = createCanvas(Math.ceil(tileSize * zoom * dpr), Math.ceil(tileSize * zoom * dpr))
              const tileCtx = tileCanvas.getContext('2d')
              tileCtx.scale(zoom * dpr, zoom * dpr)

              const origTileFillRect = tileCtx.fillRect.bind(tileCtx)
              tileCtx.fillRect = (x, y, w, h) => {
                const t = tileCtx.getTransform()
                const sourceX = tx + (t.e / (zoom * dpr)) + x
                const sourceY = ty + (t.f / (zoom * dpr)) + y
                tiledFills.push({
                  fillStyle: String(tileCtx.fillStyle),
                  x: Math.round(sourceX),
                  y: Math.round(sourceY),
                  w: Math.round(w),
                  h: Math.round(h),
                  tileX: tx,
                  tileY: ty,
                })
                origTileFillRect(x, y, w, h)
              }

              renderSheetViewport(
                sheet,
                tileCtx as never,
                geometry,
                { box: { x: tx, y: ty, width: tileSize, height: tileSize }, zoom, deviceScale: dpr },
                { conditionalFormatting: batch },
                { cache },
              )
            }
          }

          // Verify that for cell B2 (row 1, col 1), both full and tiled renders carry the exact same scale fill.
          // Cell B2 crosses the 256px tile seam: both the first tile (tx=0) and the seam tile (tx=256) must carry the fill.
          const b2FillColor = batch!.layer.styleAt(1, 1, undefined)?.fillColor
          expect(b2FillColor).toBeDefined()
          const expectedFillHex = `#${b2FillColor!.slice(2).toLowerCase()}`

          const fullB2 = fullFills.find(f => f.fillStyle.toLowerCase() === expectedFillHex)
          const firstTileB2 = tiledFills.find(f => f.tileX === 0 && f.fillStyle.toLowerCase() === expectedFillHex)
          const seamTileB2 = tiledFills.find(f => f.tileX === tileSize && f.fillStyle.toLowerCase() === expectedFillHex)

          expect(fullB2, `full render at zoom=${zoom} dpr=${dpr} must have B2 fill`).toBeDefined()
          expect(firstTileB2, `first tile (tx=0) render at zoom=${zoom} dpr=${dpr} must have B2 fill`).toBeDefined()
          expect(seamTileB2, `seam (second) tile (tx=${tileSize}) render at zoom=${zoom} dpr=${dpr} must have B2 fill`).toBeDefined()
          expect(firstTileB2!.fillStyle).toBe(fullB2!.fillStyle)
          expect(seamTileB2!.fillStyle).toBe(fullB2!.fillStyle)
          expect(firstTileB2!.x).toBe(fullB2!.x)
          expect(firstTileB2!.y).toBe(fullB2!.y)
          expect(firstTileB2!.w).toBe(fullB2!.w)
          expect(firstTileB2!.h).toBe(fullB2!.h)
        }
      }

      batch?.dispose()
    })
  })

  it('RED mutant proof: omitting CF fill on non-first tiles fails the seam tile check', async () => {
    // Proves that if CF fill is skipped on non-first tiles (tx > 0), the seam tile assertion fails.
    await withInjectedCfEvidence(SCALE_EVIDENCE, async () => {
      const sheet = makeScaleSheet(1, 3, 1, 2)
      sheet.cols = [
        { min: 0, max: 0, widthChars: 30 },
        { min: 1, max: 1, widthChars: 35 },
      ]
      const metrics = computeMetrics(sheet)
      const geometry = buildSheetGeometry(sheet)
      const batch = await prepareCfPaint(sheet, { evidence: cfRuntimeEvidence })
      const tileSize = 256
      const b2FillColor = batch!.layer.styleAt(1, 1, undefined)?.fillColor!
      const expectedFillHex = `#${b2FillColor.slice(2).toLowerCase()}`

      const mutantTiledFills: Array<{ tileX: number; fillStyle: string }> = []
      const cache = new LayoutCache()

      for (let tx = 0; tx < metrics.widthPx; tx += tileSize) {
        for (let ty = 0; ty < metrics.heightPx; ty += tileSize) {
          const tileCanvas = createCanvas(tileSize, tileSize)
          const tileCtx = tileCanvas.getContext('2d')
          const origTileFillRect = tileCtx.fillRect.bind(tileCtx)
          tileCtx.fillRect = (x, y, w, h) => {
            mutantTiledFills.push({ tileX: tx, fillStyle: String(tileCtx.fillStyle) })
            origTileFillRect(x, y, w, h)
          }

          // The mutant: skips CF fill on non-first tiles (tx > 0)
          const mutantPrepared = tx === 0 ? { conditionalFormatting: batch } : undefined
          renderSheetViewport(
            sheet,
            tileCtx as never,
            geometry,
            { box: { x: tx, y: ty, width: tileSize, height: tileSize }, zoom: 1, deviceScale: 1 },
            mutantPrepared,
            { cache },
          )
        }
      }

      const firstTile = mutantTiledFills.find(f => f.tileX === 0 && f.fillStyle.toLowerCase() === expectedFillHex)
      const seamTile = mutantTiledFills.find(f => f.tileX === tileSize && f.fillStyle.toLowerCase() === expectedFillHex)

      // First tile carries fill, but seam tile fails to carry it under the mutant
      expect(firstTile, 'mutant: first tile still has fill').toBeDefined()
      expect(seamTile, 'mutant: seam tile must NOT have fill under non-first tile omission').toBeUndefined()

      batch?.dispose()
    })
  })

  it('CF scales never expand static, print, or export extents (Lesson 4)', async () => {
    // A whole-column scale on A:A does not enlarge the print canvas or
    // exported PNG/PDF image dimensions beyond the data used extent.
    await withInjectedCfEvidence(SCALE_EVIDENCE, async () => {
      const dataRows = 10
      const sheet: XlsxSheet = {
        name: 'WholeColumnScale',
        sheetId: 'wc',
        rows: Array.from({ length: dataRows }, (_, r) => ({
          index: r,
          cells: [{ ref: `A${r + 1}`, row: r, col: 0, value: (r + 1) * 10, styleIndex: 0 }],
        })),
        cols: [],
        merges: [],
        mergeRanges: [],
        conditionalFormatting: {
          diagnostics: [],
          rules: [
            {
              id: 1,
              priority: 1,
              documentOrder: 0,
              stopIfTrue: false,
              ranges: [{ firstRow: 0, lastRow: 1048575, firstCol: 0, lastCol: 0 }], // A:A
              writes: new Set(['colorScale', 'fill']),
              definition: {
                kind: 'colorScale',
                points: [
                  { bound: { type: 'min' }, argb: 0xfff8696b },
                  { bound: { type: 'max' }, argb: 0xff63be7b },
                ],
                dialect: 'base',
                formulas: [],
              },
            },
          ],
        },
      } as XlsxSheet

      // 1. Static metrics: row count must equal data used extent, not 1048576!
      const m = computeMetrics(sheet)
      expect(m.rowHeightsPx.length).toBe(dataRows)
      expect(m.heightPx).toBeLessThan(1000) // ~200px, nowhere near 20 million px

      // 2. Print metrics: print page dimensions must remain bounded by data
      const print = computePrintMetrics(sheet, m)
      expect(print.pagesWide * print.pagesTall).toBe(1)
      expect(print.paperPx.height).toBeLessThan(2000)

      // 3. Export extents: getPaintables produces pages bounded by data used extent
      const pages = await getPaintables({ sheets: [sheet] })
      try {
        expect(pages).toHaveLength(1)
        expect(pages[0].spec.heightPx).toBeLessThan(1000)
        expect(pages[0].spec.widthPx).toBeLessThan(1000)
      } finally {
        pages.dispose()
      }
    })
  })
})
