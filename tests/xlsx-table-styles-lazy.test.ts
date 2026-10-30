// @vitest-environment node
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createCanvas } from 'canvas'
import { beforeEach, describe, expect, test, vi } from 'vitest'
import { OfficePackage } from '../src/core/zip'
import { parseXlsx } from '../src/xlsx/parse'

const fixture = () => readFileSync(fileURLToPath(new URL('./fixtures/xlsx-pivots/table-medium-2.xlsx', import.meta.url)))
const parsed = async () => parseXlsx(await OfficePackage.load(fixture()))

beforeEach(() => { vi.resetModules() })

describe('table-style lazy loading', () => {
  test('direct render falls back without preload; only styled tables load the lazy chunk', async () => {
    const direct = await parsed()
    const plain = await parsed()
    const styledSheet = direct.sheets[0]!
    const plainSheet = plain.sheets[0]!
    styledSheet.tables![0]!.styleInfo = { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false }
    plainSheet.tables![0]!.styleInfo = undefined

    const freshLoader = await import('../src/xlsx/tableStyleLoader')
    const freshRenderer = await import('../src/xlsx/render')
    expect(freshLoader.isTableStyleEngineLoaded()).toBe(false)
    const styledMetrics = freshRenderer.computeMetrics(styledSheet)
    const plainMetrics = freshRenderer.computeMetrics(plainSheet)
    const styledCanvas = createCanvas(styledMetrics.widthPx, styledMetrics.heightPx)
    const plainCanvas = createCanvas(plainMetrics.widthPx, plainMetrics.heightPx)
    freshRenderer.renderSheet(styledSheet, styledCanvas.getContext('2d') as unknown as CanvasRenderingContext2D, styledMetrics)
    freshRenderer.renderSheet(plainSheet, plainCanvas.getContext('2d') as unknown as CanvasRenderingContext2D, plainMetrics)
    expect(styledCanvas.toBuffer('raw')).toEqual(plainCanvas.toBuffer('raw'))
    expect(styledSheet.diagnostics).toContainEqual(expect.objectContaining({ feature: 'table-style-engine-unavailable' }))
    expect(freshLoader.isTableStyleEngineLoaded()).toBe(false)

    await vi.resetModules()
    let loaded = 0
    vi.doMock('../src/xlsx/tableStyles/regionLayer', async () => {
      loaded++
      const loader = await import('../src/xlsx/tableStyleLoader')
      loader.registerTableStyleEngine(() => undefined)
      return {}
    })
    try {
      const { getPaintables } = await import('../src/render/paint')
      const { computeMetrics, renderSheet } = await import('../src/xlsx/render')
      for (const unstyledTable of [true, false]) {
        const doc = await parsed()
        if (unstyledTable) doc.sheets[0]!.tables![0]!.styleInfo = undefined
        else doc.sheets[0]!.tables = undefined
        const pages = await getPaintables(doc)
        const metrics = computeMetrics(doc.sheets[0]!)
        const expected = createCanvas(metrics.widthPx, metrics.heightPx)
        const actual = createCanvas(pages[0].spec.widthPx, pages[0].spec.heightPx)
        renderSheet(doc.sheets[0]!, expected.getContext('2d') as unknown as CanvasRenderingContext2D, metrics)
        pages[0].paint(actual.getContext('2d') as unknown as CanvasRenderingContext2D)
        try { expect(actual.toBuffer('raw')).toEqual(expected.toBuffer('raw')) }
        finally { pages.dispose() }
      }
      expect(loaded).toBe(0)
      const styled = await parsed()
      styled.sheets[0]!.tables![0]!.styleInfo = { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false }
      const pages = await getPaintables(styled)
      pages.dispose()
      expect(loaded).toBe(1)
    } finally {
      vi.doUnmock('../src/xlsx/tableStyles/regionLayer')
      await vi.resetModules()
    }
  })
})
