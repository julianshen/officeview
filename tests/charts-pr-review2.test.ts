/**
 * PR #15 review (codex), round 2:
 * 1. A resident chart model paints in the workspace it is charged to (its owner), even when a receiving context
 *    is supplied; the receiving context serves ownerless (transferred) models only.
 * 2. An XLSX chart chunk that fails to load degrades to cards with a diagnostic (static getPaintables and the
 *    interactive sheet asset lease), never a rejection.
 * Fixtures are Office-shaped synthetic workbooks.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createCanvas } from 'canvas'
import { parseXlsx } from '../src/xlsx/parse'
import { paintXlsxChart } from '../src/charts/paint'
import { ownedChartWorkspace } from '../src/charts/layout'
import { documentChartContext } from '../src/charts/context'
import { releaseChartModel } from '../src/charts/own'
import type { ChartModel } from '../src/charts/model'
import type { XlsxDocument } from '../src/xlsx/types'
import { CACHED_COLUMN, chartWorkbook } from './helpers/xlsx-chart-fixture'

async function residentChart(): Promise<{ doc: XlsxDocument; chart: ChartModel }> {
  const doc = await parseXlsx(await chartWorkbook(CACHED_COLUMN))
  const content = doc.sheets[0].drawings![0].content
  if (content?.kind !== 'xlsxChart') throw new Error('no chart')
  return { doc, chart: content.chart }
}
const canvas = () => createCanvas(320, 200).getContext('2d') as unknown as CanvasRenderingContext2D

afterEach(() => { vi.doUnmock('../src/charts'); vi.resetModules() })

describe('PR #15 review round 2', () => {
  it('1: a resident model paints in its owner workspace, not the supplied receiving context', async () => {
    const { chart } = await residentChart()
    const owner = ownedChartWorkspace(chart)!
    expect(owner, 'a resident model has an owner').toBeDefined()
    const receiver = documentChartContext('desktop')
    const before = owner.used
    paintXlsxChart(chart, canvas(), 320, 200, undefined, receiver)
    // Layout (retained in the owner) and paint work were charged to the owner; nothing reached the receiver.
    expect(receiver.workspace.used).toBe(0)
    expect(owner.used).toBeGreaterThan(before)
    releaseChartModel(chart)
  })

  it('1: an ownerless (transferred) model still paints in the receiving context', async () => {
    const { chart } = await residentChart()
    const transferred = structuredClone(chart)
    expect(ownedChartWorkspace(transferred)).toBeUndefined()
    const receiver = documentChartContext('mobile')
    paintXlsxChart(transferred, canvas(), 320, 200, undefined, receiver)
    expect(receiver.workspace.used).toBeGreaterThan(0)
    releaseChartModel(chart)
  })

  it('2: an XLSX chart chunk that fails to load leaves cards and a diagnostic (getPaintables, sheet assets)', async () => {
    const { doc } = await residentChart()
    const transferred = structuredClone(doc) as XlsxDocument
    vi.resetModules()
    let attempts = 0
    vi.doMock('../src/charts', () => { attempts++; throw new Error('chunk failed to load') })
    const unhandled: unknown[] = []
    const onUnhandled = (reason: unknown) => { unhandled.push(reason) }
    process.on('unhandledRejection', onUnhandled)
    try {
      const { getPaintables } = await import('../src/render/paint')
      const pages = await getPaintables(transferred)
      expect(attempts).toBeGreaterThan(0)
      expect(pages.length).toBe(1)
      expect(pages.chartDiagnostics?.map(d => d.code)).toEqual(['chart-engine-unavailable'])
      // The page paints (the chart as its card).
      const { widthPx, heightPx } = pages[0].spec
      const ctx = createCanvas(Math.max(1, widthPx), Math.max(1, heightPx)).getContext('2d') as unknown as CanvasRenderingContext2D
      expect(() => pages[0].paint(ctx)).not.toThrow()
      pages.dispose()

      const { prepareSheetAssets } = await import('../src/components/xlsx/sheetAssets')
      const lease = await prepareSheetAssets(transferred)
      expect(lease.assets.chartContext).toBeUndefined()
      expect(lease.chartsUnavailable).toBe(true)
      lease.dispose()
      await new Promise(resolve => setTimeout(resolve, 0))
      expect(unhandled).toEqual([])
    } finally { process.off('unhandledRejection', onUnhandled) }
  })
})
